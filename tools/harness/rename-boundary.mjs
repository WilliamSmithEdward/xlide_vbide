/*
 * A rename crosses MODULES and stops at the WORKBOOK.
 *
 * Both fixtures hold a module named Helpers, and both of those declare a Recalculate: a name is
 * not an identity across workbooks, and two open projects are exactly the state where a rename
 * that resolves by name instead of by project rewrites a stranger's code. rename-features.mjs
 * asks which modules a rename touches inside ONE workbook; this is the other half of what rename
 * has to get right, and it is only askable with two open.
 *
 * Run against both fixtures in one session, which one launch sets up:
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\RenameFixture.xlsm, artifacts\fixtures\TwinFixture.xlsm
 *   node tools\harness\rename-boundary.mjs
 *
 * Every assertion on the twin is BYTE-IDENTICAL text, not a token search: the defect this guards
 * does not announce itself, it just edits.
 */
import { open, waitFor } from "./xlide-api.mjs";

const api = await open();

let passed = 0;
const broken = [];
const check = (what, ok, detail) => {
  if (ok) { passed += 1; console.log(`ok   ${what}`); }
  else { broken.push(what); console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`); }
};

const projects = (await api.projects()).projects;
const rename = projects.find((p) => /RenameFixture/i.test(p.project));
const twin = projects.find((p) => /TwinFixture/i.test(p.project));

// A precondition that does not hold is a FAILURE here, not weather: this suite runs where the
// launcher opened both fixtures, so a missing twin means the wiring broke, and reporting it as
// a skip is how a boundary check stops existing without anyone deciding that.
if (!rename || !twin) {
  console.log(`FAIL both fixtures must be open; saw ${projects.map((p) => p.project).join(", ")}`);
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}

const read = async (module, project) =>
  (await api.readModule(module, project).catch(() => ({ text: null }))).text;

const before = {
  renameHelpers: await read("Helpers", rename.projectId),
  twinHelpers: await read("Helpers", twin.projectId),
};

check("RenameFixture's Helpers declares Recalculate", /Recalculate/.test(before.renameHelpers ?? ""));
check("TwinFixture's Helpers also names Recalculate, so the boundary is askable",
  /Recalculate/.test(before.twinHelpers ?? ""),
  "New-TwinFixture.ps1 builds the collision on purpose; without it every check below is vacuous");

// Drive the rename from RenameFixture's declaration, with the twin standing right there.
await api.pane("open", { module: "Helpers", project: rename.projectId });
await waitFor("RenameFixture's Helpers on screen", async () =>
  (await api.ui()).focus.model?.toLowerCase().includes("renamefixture"));

const said = await api.act("rename", { word: "Recalculate", newName: "Recompute" });
check("the rename was accepted", said.did === true, said.detail);

await waitFor("the rename to land in RenameFixture", async () =>
  /Recompute/.test((await read("Helpers", rename.projectId)) ?? ""));

const after = {
  renameHelpers: await read("Helpers", rename.projectId),
  twinHelpers: await read("Helpers", twin.projectId),
};

check("the declaration workbook took the new name", /Recompute/.test(after.renameHelpers ?? ""));
check("the twin's Helpers is BYTE-IDENTICAL to before the rename",
  after.twinHelpers === before.twinHelpers,
  "a rename that crossed the workbook boundary edited code the developer never asked about");

// Put it back, and hold the undo to the same boundary.
const undone = await api.undoRename();
check("the undo ran to the end", undone.undone === true, undone.stopped ?? JSON.stringify(undone));

await waitFor("RenameFixture back to Recalculate", async () =>
  /Recalculate/.test((await read("Helpers", rename.projectId)) ?? ""));

check("the undo restored the declaration workbook byte for byte",
  (await read("Helpers", rename.projectId)) === before.renameHelpers);
check("and the twin is still byte-identical, through rename AND undo",
  (await read("Helpers", twin.projectId)) === before.twinHelpers);

console.log(`\n${passed} passed, ${broken.length} failed`);
process.exit(broken.length === 0 ? 0 : 1);
