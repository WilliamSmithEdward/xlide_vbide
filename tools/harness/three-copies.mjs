/*
 * Three copies of every module, and which operations leave them disagreeing.
 *
 * A module open in this product exists three times over:
 *
 *   the WORKBOOK's, in the VBE's own code module, which is what gets saved and what Run executes
 *   the SURFACE's, which the page draws and the developer edits
 *   the ENGINE's, which every finding, completion, hover and rename is computed against
 *
 * They are meant to be the same text. Nothing could check the third one until 2026-08-08, and in
 * the week before that a disagreement between the first two was found by accident (a surface
 * holding an empty document for a module the host had 42 lines of) and a disagreement involving
 * the third was found by a developer noticing a red underline in the wrong place. Both were
 * invisible to every check in the repo.
 *
 * So this walks the operations that touch a module and asks all three after each one. It is a
 * HUNT rather than a regression suite: it is meant to be pointed at new operations as they are
 * added, and to fail loudly the first time one of them drifts.
 *
 * Ordered cheapest-to-recover-from first, because a failure leaves the fixture mid-edit and
 * everything after it inherits that state.
 *
 *   node tools\harness\three-copies.mjs
 */

import { open, wait, waitFor } from "./xlide-api.mjs";

const api = await open({});
const target = "HelpersExtra";

// The workbook holding the target, not whichever is active. The fifth suite to need this.
const home = await api.projectHolding(target);
const runnable = home !== null;

if (!runnable) {
  console.log(`no open workbook holds a module named ${target}.`);
  console.log("open the rename fixture and run this again:");
  console.log("  tools\\harness\\Start-Excel.ps1 -Workbook artifacts\\fixtures\\RenameFixture.xlsm");
  process.exitCode = 2;
}

const project = runnable ? await api.project(home.project) : { projectId: null };
const original = runnable ? (await api.readModule(target, project.projectId)).text ?? "" : "";

let passed = 0;
const failures = [];

function check(what, ok, detail) {
  if (ok) {
    passed++;
    console.log(`  ok   ${what}`);
  } else {
    failures.push(`${what}${detail ? `: ${detail}` : ""}`);
    console.log(`  FAIL ${what}${detail ? `\n       ${detail}` : ""}`);
  }
}

/**
 * THE OPERATION HAS LANDED SOMEWHERE, which is a different question from the three agreeing.
 *
 * Each step used to be `wait(2500)` then `agree(...)`, and the sleep was not the belt-and-braces
 * it looked like. `agree` polls until the copies match, so run before the operation has reached
 * any of them it finds them matching on the text from BEFORE and passes, having measured nothing
 * - the same vacuous pass a probe hit on 2026-08-10. The sleep was what made that unlikely rather
 * than impossible, at 27.6s a run.
 *
 * So: wait for the change to appear in ONE copy, then let `agree` decide about all three. The wait
 * names the surface; the assertion is host-against-surface and engine-against-surface, so nothing
 * is being waited into being true.
 */
async function landed(what, seen) {
  return waitFor(`${what} to reach the surface`, async () => {
    const text = (await api.native({ text: true })).surfaceText ?? "";
    return seen(text) ? text : null;
  });
}

/**
 * THE TWO STEPS WITH NO RELIABLE OBSERVABLE, kept as a bounded settle and named as such.
 *
 * Format Module is idempotent, and the line typed above arrives through the keyboard pipeline
 * already correctly indented, so there is frequently nothing for it to change. Undo does not
 * reliably take that line back out either. Both were measured on 2026-08-10 by writing the wait
 * and watching it time out on a healthy product.
 *
 * So these two keep a sleep, deliberately, rather than a wait that asserts something untrue. It
 * is a smaller sleep than the 2000 and 2500 it replaces because `agree` polls for twelve seconds
 * afterwards and is what actually decides; this only has to outrun the window in which all three
 * copies still agree on the text from BEFORE the operation.
 *
 * What the committed suite never established, and this does not either: what undo is expected to
 * DO here. It only ever asked whether the three copies agree, so an undo that did nothing at all
 * passed it. Worth pinning separately.
 */
const settle = async (what) => {
  void what;
  await wait(900);
};

/**
 * All three, compared. Settling is allowed for, because a write-back and a didChange are both
 * asynchronous and neither is instant; what is not allowed is settling never arriving.
 */
async function agree(after, { budgetMs = 12000 } = {}) {
  let last = null;

  const settled = await waitFor(`the three copies to agree after ${after}`, async () => {
    const [below, held] = await Promise.all([
      api.native({ text: true }),
      api.engineSource(target, { text: true }),
    ]);

    last = {
      host: below.nativeText,
      surface: below.surfaceText,
      engine: held.engineText ?? null,
      engineHolds: held.engineHolds,
      hostVsSurface: below.nativeContent === below.surfaceContent,
      engineVsSurface: held.engineContent === held.surfaceContent,
    };

    return last.hostVsSurface && (!last.engineHolds || last.engineVsSurface) ? last : null;
  }, { budgetMs }).catch(() => null);

  const state = settled ?? last;

  check(`after ${after}, the workbook and the surface hold the same text`,
    state?.hostVsSurface === true,
    `host ${JSON.stringify((state?.host ?? "").slice(0, 90))} against ` +
    `surface ${JSON.stringify((state?.surface ?? "").slice(0, 90))}`);

  // A engine holding nothing is not a disagreement: it has not been told about this module yet,
  // and will answer from the seeded copy, which is the workbook's.
  check(`after ${after}, the analyzer holds the same text as the surface`,
    state?.engineHolds !== true || state?.engineVsSurface === true,
    `engine ${JSON.stringify((state?.engine ?? "").slice(0, 90))} against ` +
    `surface ${JSON.stringify((state?.surface ?? "").slice(0, 90))}`);

  return state;
}

const SEED = [
  "Option Explicit",
  "",
  "Public Sub Probe()",
  "    Dim n As Long",
  "    n = 1",
  "    Debug.Print n",
  "End Sub",
  "",
];

const shown = async () => {
  await waitFor(`${target} shown`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith(`/${target.toLowerCase()}`) ? ui : null;
  });
};

if (runnable) try {
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await waitFor("the tab to go before the module is reseeded", async () =>
    !(await api.ui()).workspace.groups.some((group) =>
      group.tabs.some((tab) => tab.module?.toLowerCase() === target.toLowerCase())));

  await api.writeModule(target, SEED.join("\r\n"), project.projectId);
  await waitFor("the seed to be in the module", async () =>
    ((await api.readModule(target, project.projectId)).text ?? "").includes("Public Sub Probe()"));

  await api.pane("open", { module: target, project: project.projectId });
  await shown();

  console.log("\n1. opening a module\n");
  await agree("opening it");

  console.log("\n2. typing\n");
  const lines = ((await api.readModule(target, project.projectId, { live: true })).text ?? "")
    .split(/\r?\n/);
  const at = lines.findIndex((line) => line.includes("n = 1")) + 1;
  await api.caret(at, { module: target, project: project.projectId, column: (lines[at - 1] ?? "").length + 1 });
  // Load-bearing: `type` goes to wherever the caret IS, so typing before it lands puts the line in
  // another procedure and the suite reports a disagreement it caused itself.
  await waitFor("the caret to reach the line about to be typed on", async () =>
    (await api.ui()).focus?.line === at);

  await api.type("\nDim extra As String");
  await landed("the typed line", (text) => text.includes("Dim extra"));
  await agree("typing a line");

  console.log("\n3. Format Module\n");
  await api.act("format");
  await settle("formatting");
  await agree("formatting");

  console.log("\n4. undo\n");
  await api.act("undo");
  await settle("an undo");
  await agree("an undo");

  console.log("\n5. a write from OUTSIDE the surface, into the module on screen\n");
  await api.writeModule(target, [...SEED.slice(0, 5), "    Debug.Print n * 2", ...SEED.slice(6)].join("\r\n"),
    project.projectId);
  await landed("the outside write", (text) => text.includes("n * 2"));
  await agree("a write from outside");

  console.log("\n6. the same, while the module sits in a BACKGROUND tab\n");
  const other = "Helpers";
  await api.pane("open", { module: other, project: project.projectId });
  await waitFor(`${other} shown`, async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith(`/${other.toLowerCase()}`) ? ui : null;
  });

  await api.writeModule(target, [...SEED.slice(0, 5), "    Debug.Print n * 3", ...SEED.slice(6)].join("\r\n"),
    project.projectId);
  // The tab is in the BACKGROUND, so the surface is not where this shows up first. Waited for in
  // the workbook instead, which is the copy the write goes to.
  await waitFor("the background write to reach the workbook", async () =>
    ((await api.readModule(target, project.projectId)).text ?? "").includes("n * 3"));

  // Back to it, which is when a developer would notice.
  await api.act("activate", { module: target, project: project.projectId });
  await shown();
  await agree("coming back to a tab written behind its back");

  console.log("\n7. a rename, which rewrites every module that mentions the symbol\n");
  await api.act("rename", { word: "Probe", newName: "Probed" });
  await landed("the rename", (text) => text.includes("Probed"));
  await agree("a rename");

  await api.undoRename();
  await landed("the rename being undone", (text) => !text.includes("Probed"));
  await agree("undoing the rename");
} finally {
  await api.pane("close", { module: target, project: project.projectId, answer: "discard" });
  await api.writeModule(target, original, project.projectId);
  // Read back rather than slept on, and it must not throw: this is the `finally`, so a failure
  // here would replace whatever really went wrong with a timeout about the tidying up.
  await waitFor("the fixture to be back as it was", async () =>
    ((await api.readModule(target, project.projectId)).text ?? "").trim() === original.trim())
    .catch(() => null);

  const restored = ((await api.readModule(target, project.projectId)).text ?? "").trim() === original.trim();
  console.log(`\n${target} restored: ${restored}`);

  console.log(`\n${passed} passed, ${failures.length} failed`);
  for (const failure of failures) {
    console.log(`  ${failure}`);
  }

  process.exitCode = failures.length === 0 ? 0 : 1;
}
