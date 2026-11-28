/*
 * A code tab's unsaved dot, around a save.
 *
 * The dot is not Excel's Saved flag: that flag never goes back when an edit is undone to the
 * saved text. So the host keeps each module's text from the last moment the project read as
 * saved, and the dot is "the project is unsaved AND this module's text differs from that".
 * The whole thing rests on the kept text being the SAVED text, and it once was not: typing
 * reaches the editor at once and the module a pause later, the project still reads as saved in
 * between, and the kept text was taken from the editor in that gap. The first edit after a save
 * then showed no dot, and putting the text back to what was saved showed one (2026-11-28).
 *
 *   - after a save, does the FIRST edit show the dot?
 *   - does taking that edit back clear it, with the project still unsaved?
 *   - does a second edit show it again, and a save clear it?
 *
 * Typed as keys through page-keys.mjs, because the gap is between a key and its write-back.
 *
 * It saves the project it is run against, twice, with the module's text as it found it.
 *
 * Run against the debug fixture:
 *   node tools\harness\unsaved-dot.mjs
 */

import { open, reporter, wait, waitFor } from "./xlide-api.mjs";
import { connectToEditorPage, devtoolsPort, pageKeyboard, plainTyping } from "./page-keys.mjs";

const api = await open({});
const { check, done } = reporter();
const project = await api.project();
const MODULE = "Helper";

const original = (await api.readModule(MODULE, project.projectId)).text ?? "";
const page = await connectToEditorPage(devtoolsPort(api));
const keys = pageKeyboard(page);
const restoreTyping = await plainTyping(api);

const EDITOR = "window.xlideBridge.workspace.activeEditor()";
const dot = async () => (await api.ui()).workspace.groups.flatMap((group) => group.tabs).find((tab) => tab.module === MODULE)?.dirty;
const projectUnsaved = async () => (await api.scm({ action: "status", project: project.projectId })).dirty;
const firstLine = async () => {
  const answer = await api.ask(`JSON.stringify({ value: ${EDITOR}.getModel().getLineContent(1) })`);
  return (typeof answer === "string" ? JSON.parse(answer) : answer).value;
};
const moduleFirstLine = async () => ((await api.readModule(MODULE, project.projectId)).text ?? "").split(/\r?\n/)[0];

/** Waits for the dot to read `want`, and answers what it read last. */
const dotBecomes = (want) => waitFor(`the dot to be ${want}`, async () => (await dot()) === want, { budgetMs: 10000, pollMs: 200 })
  .then(() => want, async () => dot());

/** Saves, and waits for the project and the dot to say so. */
const save = async () => {
  await api.command("save");
  await api.waitUntilResponsive({ timeout: 20000 }).catch(() => {});
  await waitFor("the project to read as saved", async () => (await projectUnsaved()) === false, { budgetMs: 15000, pollMs: 250 }).catch(() => {});
};

try {
  console.log("the unsaved dot, around a save\n");

  await api.pane("open", { module: MODULE, project: project.projectId });
  await waitFor("the module to be on screen", async () => (await api.ui()).workspace.active?.module === MODULE, { budgetMs: 15000, pollMs: 250 });
  const wasFirst = await firstLine();

  await save();
  check("after a save the project reads as saved", (await projectUnsaved()) === false);
  check("and the tab carries no dot", (await dotBecomes(false)) === false, String(await dot()));

  // ---- the first edit after the save ----
  await api.ask(`(() => { const e = ${EDITOR}; e.focus(); e.setPosition({ lineNumber: 1, column: 1 }); return 1; })()`);
  await wait(200);
  await keys.press("'");
  check("a key typed goes into the editor", (await firstLine()) === `'${wasFirst}`, JSON.stringify(await firstLine()));
  const written = await waitFor("the key to reach the module", async () => (await moduleFirstLine()) === `'${wasFirst}`,
    { budgetMs: 10000, pollMs: 200 }).then(() => true, () => false);
  check("and into the project's module", written, JSON.stringify(await moduleFirstLine()));
  check("which leaves the project unsaved", (await projectUnsaved()) === true);
  check("the first edit after a save shows the dot", (await dotBecomes(true)) === true, String(await dot()));

  // ---- taken back ----
  await keys.press("Backspace");
  const back = await waitFor("the module to hold the saved text again", async () => (await moduleFirstLine()) === wasFirst,
    { budgetMs: 10000, pollMs: 200 }).then(() => true, () => false);
  check("taking the key back puts the module's text back", back && (await firstLine()) === wasFirst, JSON.stringify(await moduleFirstLine()));
  check("and clears the dot, though the project still reads as unsaved",
    (await dotBecomes(false)) === false && (await projectUnsaved()) === true, `dot ${await dot()}, project unsaved ${await projectUnsaved()}`);

  // ---- and again, through a save ----
  await keys.press("'");
  check("a second edit shows the dot again", (await dotBecomes(true)) === true, String(await dot()));
  await save();
  check("a save clears it", (await dotBecomes(false)) === false, String(await dot()));
  await keys.press("Backspace");
  check("and an edit to what was just saved shows it", (await dotBecomes(true)) === true, String(await dot()));
} finally {
  page.socket.close();
  await restoreTyping();
  await api.writeModule(MODULE, original, project.projectId).catch(() => {});
  await api.command("save").catch(() => {});
}

process.exit(done());
