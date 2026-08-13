/*
 * An import must not write over edits the developer has not written yet.
 *
 * THE DEFECT. Import writes the module through the session's writer and then syncs the surface,
 * which replaces the document on screen. Measured 2026-08-09: typing into a module and importing
 * over it took the typing away with no question, no notice, and "1 changed, 0 failed" - while
 * CLOSING a tab in that same state raises a save/discard/cancel gate, because losing a developer's
 * work silently is not something this product does anywhere else. It is worse than the file cases
 * fixed the same day, because the developer is looking at the text when it disappears.
 *
 * Refused rather than asked. An import touching twelve modules would ask twelve questions, and the
 * row already has somewhere to say why. What matters as much as the refusal is that it is
 * RECOVERABLE, so the last checks discard and import again.
 *
 * WHY THIS IS NOT A SECTION OF module-sync.mjs. It needs the module on the surface with edits
 * pending, and by the end of that suite the surface has been opened, closed and re-opened enough
 * that the shim's document key for a freshly opened module no longer matches what a live read asks
 * with. The precondition became the thing being tested, which is how a check ends up reporting on
 * itself. Here it runs first, against a session that has done nothing else.
 *
 *   node tools\harness\import-guard.mjs
 */
import { mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { open, reporter, scratchModule, wait, waitFor } from "./xlide-api.mjs";

const api = await open();
const project = await api.project();
const name = `Guard${process.pid}`;
const folder = join(tmpdir(), `xlide-import-guard-${process.pid}`);
mkdirSync(folder, { recursive: true });

const { check, done } = reporter();
const scratch = scratchModule(api, project.projectId, name);

const ORIGINAL = ["Option Explicit", "", "Public Sub Go()", "    Debug.Print \"original\"", "End Sub", ""];
const FROM_FILE = ["Option Explicit", "", "Public Sub Go()", "    Debug.Print \"from the file\"", "End Sub", ""];

let made = false;
try {
  await api.component("add", { kind: "module", name, project: project.projectId });
  made = true;
  await api.writeModule(name, ORIGINAL.join("\r\n"), project.projectId);
  await waitFor("the seed to be in the module", async () =>
    ((await api.readModule(name, project.projectId)).text ?? "").includes(ORIGINAL[ORIGINAL.length - 2]));

  await api.pane("open", { module: name, project: project.projectId });
  await waitFor("the module to be the one on screen", async () =>
    (await api.ui()).focus.model?.toLowerCase().endsWith(`/${name.toLowerCase()}`));

  writeFileSync(
    join(folder, `${name}.bas`),
    [`Attribute VB_Name = "${name}"`, ...FROM_FILE].join("\r\n"),
    "utf8");
  await api.syncSettings({ folder });

  // Type, and do not pause long enough for the write-back.
  await api.caret(4, { module: name, project: project.projectId, column: 1 });
  await wait(400);
  await api.type("    Debug.Print \"UNSAVED\"\r\n", { waitMs: 2500 });

  const typed = (await api.readModule(name, project.projectId, { live: true })).text ?? "";
  const stored = (await api.readModule(name, project.projectId)).text ?? "";
  check("the surface is holding edits the workbook does not have",
    typed.includes("UNSAVED") && !stored.includes("UNSAVED"),
    `surface has it: ${typed.includes("UNSAVED")}, workbook has it: ${stored.includes("UNSAVED")}. `
    + "Without unwritten edits there is nothing here to protect and nothing below means anything.");

  const refused = await api.syncApply("import", { folder, select: "all" });
  await wait(1500);

  check("the import refuses that row rather than replacing them",
    (refused.failed ?? []).some((line) => /have not written yet/.test(line)),
    JSON.stringify(refused));
  check("and nothing is counted as changed", (refused.changed ?? []).length === 0,
    JSON.stringify(refused.changed));

  const after = (await api.readModule(name, project.projectId, { live: true })).text ?? "";
  check("the developer's typing is still on the surface", after.includes("UNSAVED"));
  check("and the file's text did not get in", !after.includes("from the file"));

  // RECOVERABLE, which is the half that makes refusing the right answer rather than a dead end.
  await api.pane("close", { module: name, project: project.projectId, answer: "discard" });
  await waitFor("the tab to go, which is what makes the second import recoverable", async () =>
    !(await api.ui()).workspace.groups.some((group) =>
      group.tabs.some((tab) => tab.module?.toLowerCase() === name.toLowerCase())));

  const second = await api.syncApply("import", { folder, select: "all" });
  await wait(1500);
  check("once the edits are discarded the same import goes through",
    (second.changed ?? []).includes(name), JSON.stringify(second));

  const imported = (await api.readModule(name, project.projectId)).text ?? "";
  check("and the module now holds the file's text", imported.includes("from the file"));
} finally {
  if (made) {
    await scratch.dispose();
  }

  rmSync(folder, { recursive: true, force: true });

  process.exitCode = done();
}
