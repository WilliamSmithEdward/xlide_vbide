/*
 * Import and export act on ONE project, and the developer says which.
 *
 * THE DEFECT THIS EXISTS FOR. The dialog named no project, and the host filled the gap with two
 * different fallbacks: the plan's identity from the SHOWN project and its modules from the
 * editor's ACTIVE one. With two workbooks open those are routinely different - measured
 * 2026-08-21 with nothing contrived, seconds after opening DebugFixture and TwinFixture side by
 * side, a plan titled DebugFixture.xlsm whose rows were TwinFixture's six modules.
 *
 * That is not a cosmetic mislabel. Applying it would export one workbook's modules into the
 * other's remembered folder, or import that folder over the other workbook's code - and the
 * developer would have read a dialog that named the workbook they meant.
 *
 * So the dialog carries a project select, sends the choice with the plan AND with the apply, and
 * the host resolves one project for both the identity and the modules.
 *
 * TWO WORKBOOKS ARE THE WHOLE POINT. With one open there is nothing to get wrong, which is why
 * module-sync.mjs beside this - which runs against DebugFixture alone - passed throughout.
 *
 * Run with both fixtures open:
 *   tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\DebugFixture.xlsm,artifacts\fixtures\TwinFixture.xlsm
 *   node tools\harness\sync-scope.mjs
 */

import { mkdirSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { open, waitFor } from "./xlide-api.mjs";

const api = await open();
let passed = 0;
let failed = 0;

const check = (name, got, want = true) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (ok) {
    passed++;
    console.log(`ok   ${name}`);
  } else {
    failed++;
    console.log(`FAIL ${name}\n       got  ${JSON.stringify(got)}\n       want ${JSON.stringify(want)}`);
  }
  return ok;
};

const projects = (await api.projects()).projects;
if (projects.length < 2) {
  console.log(`FAIL this suite needs two workbooks open; found ${projects.length}`);
  console.log("\n0 passed, 1 failed");
  process.exit(1);
}

console.log(`open: ${projects.map((p) => p.project).join(", ")}\n`);

// What each project ACTUALLY holds, asked of each one by name - the yardstick every check below
// is measured against.
const held = new Map();
for (const one of projects) {
  const plan = await api.syncPlan("export", {
    project: one.project,
    folder: join(tmpdir(), `xlide-scope-${process.pid}-${held.size}`),
  });
  held.set(one.project, (plan.items ?? []).map((row) => row.module).sort());
}

for (const [name, modules] of held) {
  console.log(`  ${name}: ${modules.join(", ")}`);
}

check("the two workbooks hold different modules, so the questions below are not vacuous",
  JSON.stringify([...held.values()][0]) !== JSON.stringify([...held.values()][1]));

// ---- the dialog ---------------------------------------------------------------------------

const folder = join(tmpdir(), `xlide-scope-dialog-${process.pid}`);
rmSync(folder, { recursive: true, force: true });
mkdirSync(folder, { recursive: true });

await api.act("toolbar", { command: "openSync" });
const opened = await waitFor("the dialog to open", async () => (await api.ui()).sync);

check("it offers every open project",
  [...opened.projects].sort(), projects.map((p) => p.project).sort());

// The one being looked at, which is the answer a developer would give to "which am I working on".
const shownProject = (await api.state()).shownProject;
const shownName = projects.find((p) =>
  p.projectId.toLowerCase() === String(shownProject).toLowerCase()
  || p.project.toLowerCase() === String(shownProject).toLowerCase())?.project;

check("it starts on the project the surface is showing", opened.project, shownName);

await api.act("syncDialog", { folder });
const first = await waitFor("a plan", async () => {
  const now = (await api.ui()).sync;
  return now && !now.busy && now.rows.length > 0 ? now : false;
});

// THE DEFECT ITSELF: the rows must be the named project's, not whatever the editor calls active.
check("its rows are the chosen project's modules, not another project's",
  first.rows.map((row) => row.file.replace(/\.(bas|cls|frm)$/i, "")).sort(),
  held.get(opened.project));

// ---- and it can be pointed elsewhere ---------------------------------------------------------

const other = opened.projects.find((one) => one !== opened.project);
const moved = await api.act("syncDialog", { project: other });
check("the select points it at another open project", moved.did, true);

const second = await waitFor("a plan for the other project", async () => {
  const now = (await api.ui()).sync;
  return now && !now.busy && now.project === other ? now : false;
});

check("choosing a project forgets the folder, so it cannot inherit the other one's",
  second.folder !== folder);

await api.act("syncDialog", { folder: `${folder}-other` });
const replanned = await waitFor("the second plan", async () => {
  const now = (await api.ui()).sync;
  return now && !now.busy && now.rows.length > 0 ? now : false;
});

check("and its rows follow the choice",
  replanned.rows.map((row) => row.file.replace(/\.(bas|cls|frm)$/i, "")).sort(),
  held.get(other));

await api.act("syncDialog", { press: "close" });

rmSync(folder, { recursive: true, force: true });
rmSync(`${folder}-other`, { recursive: true, force: true });

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
