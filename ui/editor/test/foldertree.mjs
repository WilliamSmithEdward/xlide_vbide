// The folder view's arithmetic, tested by calling it.
//
// The live suite proves the annotations reach the tree; this proves what the tree DOES with
// them - merging, sorting, the unannotated root, the ancestors a follow has to unfold - one
// failing assertion at a time rather than one screenshot at a time.

import assert from "node:assert/strict";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const scratch = await mkdtemp(path.join(tmpdir(), "xlide-foldertree-"));
const compiled = path.join(scratch, "foldertree.mjs");

await build({
  entryPoints: [path.join(root, "src", "foldertree.ts")],
  outfile: compiled,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});

const { buildFolderTree, allFolders, ancestorPaths, folderSegments } = await import(pathToFileURL(compiled).href);

const checks = [];
function check(name, fn) {
  checks.push({ name, fn });
}

const byName = (a, b) => a.name.localeCompare(b.name);
const names = (modules) => modules.map((one) => one.name);

check("an unannotated module sits at the root", () => {
  const tree = buildFolderTree([{ name: "Loose" }, { name: "Also", folder: null }], byName);
  assert.deepEqual(names(tree.modules), ["Also", "Loose"]);
  assert.deepEqual(tree.folders, []);
});

check("a dotted path makes a folder inside a folder", () => {
  const tree = buildFolderTree([{ name: "Ledger", folder: "Accounts.Ledger" }], byName);
  assert.equal(tree.folders.length, 1);
  const accounts = tree.folders[0];
  assert.equal(accounts.name, "Accounts");
  assert.equal(accounts.path, "Accounts");
  assert.deepEqual(accounts.modules, []);
  const ledger = accounts.folders[0];
  assert.equal(ledger.path, "Accounts.Ledger");
  assert.deepEqual(names(ledger.modules), ["Ledger"]);
});

check("two spellings of one folder are one folder, called what the first module called it", () => {
  const tree = buildFolderTree([
    { name: "Helpers", folder: "Shared" },
    { name: "Tools", folder: "shared" },
    { name: "Inner", folder: "SHARED.Deep" },
  ], byName);
  assert.equal(tree.folders.length, 1);
  assert.equal(tree.folders[0].name, "Shared");
  assert.deepEqual(names(tree.folders[0].modules), ["Helpers", "Tools"]);
  assert.equal(tree.folders[0].folders[0].path, "Shared.Deep");
});

check("folders sort by name without regard to case, and modules by the order given", () => {
  const tree = buildFolderTree([
    { name: "b", folder: "zeta" },
    { name: "a", folder: "Alpha" },
    { name: "c", folder: "beta" },
    { name: "Z", folder: "Alpha" },
    { name: "Y", folder: "Alpha" },
  ], (x, y) => y.name.localeCompare(x.name));
  assert.deepEqual(tree.folders.map((folder) => folder.name), ["Alpha", "beta", "zeta"]);
  assert.deepEqual(names(tree.folders[0].modules), ["Z", "Y", "a"]);
});

check("a path with loose dots and spaces is read by its segments", () => {
  assert.deepEqual(folderSegments(" A . B .. C "), ["A", "B", "C"]);
  assert.deepEqual(folderSegments(""), []);
  assert.deepEqual(folderSegments(null), []);
  const tree = buildFolderTree([{ name: "M", folder: ".A..B." }], byName);
  assert.equal(tree.folders[0].folders[0].path, "A.B");
});

check("every folder is listed parents first", () => {
  const tree = buildFolderTree([
    { name: "M", folder: "A.B.C" },
    { name: "N", folder: "A.D" },
    { name: "O", folder: "Z" },
  ], byName);
  assert.deepEqual(allFolders(tree).map((folder) => folder.path), ["A", "A.B", "A.B.C", "A.D", "Z"]);
});

check("the ancestors of a path are the paths a follow unfolds", () => {
  assert.deepEqual(ancestorPaths("Accounts.Billing.Reminders"), ["Accounts", "Accounts.Billing", "Accounts.Billing.Reminders"]);
  assert.deepEqual(ancestorPaths(null), []);
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
