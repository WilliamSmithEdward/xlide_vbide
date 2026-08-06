// The split tree's arithmetic, tested by calling it.
//
// The live probes exercise this through real drags, which is the right test of the GESTURE
// and a slow, flaky way to test the ALGEBRA. A tree that mis-collapses shows up on screen as
// a pane that vanished, three drags later; here it shows up as one failing assertion naming
// the case.
//
// The source is TypeScript, so it is compiled to a temporary module first — esbuild is
// already a dependency and takes milliseconds.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-docktree-"));
const compiled = path.join(scratch, "docktree.mjs");

await build({
  entryPoints: [path.join(root, "src", "docktree.ts")],
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});

const tree = await import(pathToFileURL(compiled).href);

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

/** A group of the named tabs, the first one active. */
const group = (...tabs) => ({ kind: "group", tabs, active: tabs[0] });
const split = (direction, children, sizes) => ({ kind: "split", direction, children, sizes });

/** Every size in the tree sums to one at each split, which is what keeps layout honest. */
function assertSizesPartition(node, where = "root") {
  if (!node || node.kind !== "split") {
    return;
  }
  const total = node.sizes.reduce((sum, size) => sum + size, 0);
  assert.ok(
    Math.abs(total - 1) < 1e-9,
    `${where}: sizes sum to ${total}, not 1`,
  );
  assert.equal(node.sizes.length, node.children.length, `${where}: a size per child`);
  node.children.forEach((child, index) => assertSizesPartition(child, `${where}/${index}`));
}

check("an emptied group is dropped", () => {
  const kept = group("a");
  const emptied = { kind: "group", tabs: [], active: "" };
  const pruned = tree.prune(split("row", [kept, emptied], [0.5, 0.5]));

  assert.equal(pruned, kept, "a split with one survivor collapses into it");
});

check("a split emptied entirely becomes nothing", () => {
  const empty = { kind: "group", tabs: [], active: "" };
  assert.equal(tree.prune(split("row", [empty, empty], [0.5, 0.5])), null);
});

check("sizes are renormalised when a sibling leaves", () => {
  const pruned = tree.prune(split("row", [
    group("a"),
    group("b"),
    { kind: "group", tabs: [], active: "" },
  ], [0.25, 0.25, 0.5]));

  assert.equal(pruned.kind, "split");
  assertSizesPartition(pruned);
  // The survivors keep their RATIO to each other: equal before, equal after.
  assert.ok(Math.abs(pruned.sizes[0] - pruned.sizes[1]) < 1e-9);
});

check("nested emptiness collapses all the way up", () => {
  const survivor = group("a");
  const empty = { kind: "group", tabs: [], active: "" };
  const pruned = tree.prune(
    split("row", [
      survivor,
      split("column", [empty, split("row", [empty, empty], [0.5, 0.5])], [0.5, 0.5]),
    ], [0.5, 0.5]),
  );

  assert.equal(pruned, survivor, "a whole dead branch leaves no trace");
});

check("splitting beside a group makes a two-child split on the right axis", () => {
  const target = group("a");
  const newcomer = group("b");

  const right = tree.splitBeside(target, target, "right", newcomer);
  assert.equal(right.kind, "split");
  assert.equal(right.direction, "row");
  assert.deepEqual(right.children, [target, newcomer], "right puts the newcomer second");

  const top = tree.splitBeside(target, target, "top", newcomer);
  assert.equal(top.direction, "column");
  assert.deepEqual(top.children, [newcomer, target], "top puts the newcomer first");
});

check("a same-direction split absorbs rather than nesting", () => {
  // Three panes side by side must stay three panes side by side, not a pane beside a pair:
  // the nested form lays out with the wrong splitters and resizes wrongly.
  const first = group("a");
  const second = group("b");
  const third = group("c");

  const root = split("row", [first, second], [0.5, 0.5]);
  const grown = tree.splitBeside(root, second, "right", third);

  assert.equal(grown.kind, "split");
  assert.equal(grown.direction, "row");
  assert.equal(grown.children.length, 3, "absorbed as a sibling");
  assert.deepEqual(grown.children, [first, second, third]);
  assertSizesPartition(grown);
});

check("a cross-direction split nests, because it must", () => {
  const first = group("a");
  const second = group("b");
  const third = group("c");

  const root = split("row", [first, second], [0.5, 0.5]);
  const grown = tree.splitBeside(root, second, "bottom", third);

  assert.equal(grown.direction, "row");
  assert.equal(grown.children.length, 2);
  assert.equal(grown.children[1].kind, "split");
  assert.equal(grown.children[1].direction, "column");
  assertSizesPartition(grown);
});

check("resizing trades between one pair and leaves the rest alone", () => {
  const sizes = [0.25, 0.25, 0.5];
  const moved = tree.resizeAt(sizes, 1, 0.1, 0.05);

  assert.ok(Math.abs(moved[0] - 0.35) < 1e-9, "the divider's left grew");
  assert.ok(Math.abs(moved[1] - 0.15) < 1e-9, "its right gave way");
  assert.equal(moved[2], 0.5, "everything else stayed exactly where it was");
  assert.ok(Math.abs(moved.reduce((a, b) => a + b, 0) - 1) < 1e-9);
});

check("resizing cannot squeeze a pane out of existence", () => {
  const squeezed = tree.resizeAt([0.5, 0.5], 1, 10, 0.1);
  assert.ok(squeezed[1] >= 0.1 - 1e-9, "the minimum holds against any drag");
  assert.ok(Math.abs(squeezed[0] + squeezed[1] - 1) < 1e-9);

  const other = tree.resizeAt([0.5, 0.5], 1, -10, 0.1);
  assert.ok(other[0] >= 0.1 - 1e-9, "and holds in the other direction");
});

check("resizing at an edge index changes nothing", () => {
  assert.deepEqual(tree.resizeAt([0.5, 0.5], 0, 0.2, 0.1), [0.5, 0.5]);
  assert.deepEqual(tree.resizeAt([0.5, 0.5], 2, 0.2, 0.1), [0.5, 0.5]);
});

check("a strip position lands before the first midpoint not yet crossed", () => {
  const midpoints = [10, 30, 50];
  assert.equal(tree.indexAtMidpoints(midpoints, 5), 0, "before everything");
  assert.equal(tree.indexAtMidpoints(midpoints, 20), 1, "past the first");
  assert.equal(tree.indexAtMidpoints(midpoints, 40), 2);
  assert.equal(tree.indexAtMidpoints(midpoints, 99), 3, "past everything is the end");
  assert.equal(tree.indexAtMidpoints([], 42), 0, "an empty strip takes it at the start");
});

check("groups are found wherever they are nested", () => {
  const deep = group("needle");
  const root = split("row", [
    group("a"),
    split("column", [group("b"), deep], [0.5, 0.5]),
  ], [0.5, 0.5]);

  assert.equal(tree.groupHolding(root, "needle"), deep);
  assert.equal(tree.groupHolding(root, "absent"), null);
  assert.equal(tree.allGroups(root).length, 3);
  assert.equal(tree.firstGroup(root).tabs[0], "a", "first is depth-first leftmost");
  assert.equal(tree.firstGroup(null), null);
});

let failures = 0;
for (const { name, fn } of checks) {
  try {
    await fn();
    console.log(`ok   ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}`);
    console.error(`     ${error.message}`);
  }
}

await rm(scratch, { recursive: true, force: true });

console.log(`${checks.length - failures}/${checks.length} passed`);
if (failures > 0) {
  process.exitCode = 1;
}
