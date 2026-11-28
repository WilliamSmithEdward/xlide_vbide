/*
 * Vim mode (#173): Vim's modes and keys in the code editors, behind the "Vim mode" setting.
 *
 * Which keys are Vim's is pinned without Excel, in ui\editor\test\vimkeys.mjs. What only a live
 * editor can say is here:
 *
 *   - is it off until asked for, and does a plain key type while it is off?
 *   - on, does a key in normal mode move the caret and type NOTHING, in the editor or the project?
 *   - does a Vim edit reach the project's module, and does `u` take it back there too?
 *   - do insert mode and Escape work, with the mode read off the status bar?
 *   - does `.` repeat what was typed and not what the editor wrote around it (the respelling)?
 *   - do the Ctrl keys divide as vimkeys.ts says, pressed as keys?
 *   - do the `/` and `:` lines work, and do `:w` and `:q` reach the host's Save and the tab?
 *   - does off put the caret and the keys back?
 *   - can the status line be read, in both themes?
 *
 * Keys are pressed through page-keys.mjs: real key events, to the page only. The mode is read
 * through `ui.vim`, which reports what the status bar shows.
 *
 * It writes a module of its own and takes it away again, and puts the developer's switch back.
 * `:w` saves the project it is run against, as the host's Save does.
 *
 * Run against the debug fixture:
 *   node tools\harness\vim.mjs
 */

import { open, reporter, scratchModule, wait, waitFor } from "./xlide-api.mjs";
import { contrastByTheme, contrastSaid } from "./contrast.mjs";
import { connectToEditorPage, devtoolsPort, pageKeyboard } from "./page-keys.mjs";

const api = await open({});
const { check, done } = reporter();
const project = await api.project();

const stamp = process.pid % 100000;
const probe = `VimProbe${stamp}`;

// Line numbers are part of what is asserted, so the text is laid out by line here.
const MODULE = [
  "Option Explicit",                    // 1
  "",                                   // 2
  `Public Sub First${stamp}()`,         // 3
  "    Dim alpha As Long",              // 4
  "    Dim beta As Long",               // 5
  "    alpha = 1",                      // 6
  "    beta = 2",                       // 7
  "End Sub",                            // 8
  "",                                   // 9
  `Public Sub Second${stamp}()`,        // 10
  "    Dim gamma As Long",              // 11
  "    gamma = 3",                      // 12
  "End Sub",                            // 13
].join("\r\n");

const switchWas = (await api.settings()).vimMode;
const probeModule = scratchModule(api, project.projectId, probe);
const page = await connectToEditorPage(devtoolsPort(api));
const keys = pageKeyboard(page);

const EDITOR = "window.xlideBridge.workspace.activeEditor()";
// In an object, as a string: the door unwraps a bare string answer and leaves an object's text
// alone, so this is the one shape that reads the same whatever the page handed back.
const ask = async (expression) => {
  const answer = await api.ask(`JSON.stringify({ value: ${expression} })`);
  return (typeof answer === "string" ? JSON.parse(answer) : answer).value;
};
const caret = () => ask(`(() => { const p = ${EDITOR}.getPosition(); return [p.lineNumber, p.column]; })()`);
const line = (number) => ask(`${EDITOR}.getModel().getLineContent(${number})`);
const lineCount = () => ask(`${EDITOR}.getModel().getLineCount()`);
const pageText = () => ask(`${EDITOR}.getModel().getValue()`);
const placeCaret = async (lineNumber, column = 1) => {
  await ask(`(() => { const e = ${EDITOR}; e.focus(); e.setPosition({ lineNumber: ${lineNumber}, column: ${column} }); return true; })()`);
  await wait(150);
};
const vim = async () => (await api.ui()).vim;
const mode = async () => (await vim()).mode;
const caretShape = () => ask(`${EDITOR}.getRawOptions().cursorStyle ?? "line"`);
const projectText = async () => ((await api.readModule(probe, project.projectId)).text ?? "").replace(/\r\n/g, "\n");
const said = (value) => JSON.stringify(value);

/**
 * Presses u until the page holds `text` again, and answers whether it came to.
 *
 * One u is one step of the editor's own undo, the step Ctrl+Z takes, and the editor counts the
 * respelling of a typed line as a step of its own. So how many presses a change takes is the
 * editor's business; that enough of them put everything back is the claim.
 */
const undoneTo = async (text) => {
  const wanted = text.replace(/\r\n/g, "\n");
  const holds = async () => String(await pageText()).replace(/\r\n/g, "\n") === wanted;
  for (let presses = 0; presses < 12 && !(await holds()); presses++) {
    await keys.press("u");
  }
  return holds();
};

/** Waits for the project's module to hold what the page holds, and answers whether it came to. */
const projectFollows = (what) => waitFor(what, async () =>
  (await projectText()).trimEnd() === String(await pageText()).replace(/\r\n/g, "\n").trimEnd(), { budgetMs: 8000, pollMs: 200 })
  .then(() => true, () => false);

try {
  console.log("vim mode\n");

  await api.settings({ vimMode: false });
  await api.component("add", { kind: "module", name: probe, project: project.projectId });
  await api.writeModule(probe, MODULE, project.projectId);
  await api.pane("open", { module: probe, project: project.projectId });
  await waitFor("the module to be on screen", async () => (await lineCount()) === 13, { budgetMs: 15000, pollMs: 250 });

  // ---- off, which is how it ships ----
  check("off, the status bar carries no mode", (await vim()).enabled === false && (await vim()).label === "", said(await vim()));
  await placeCaret(2);
  await keys.press("j");
  check("off, a plain key types", (await line(2)) === "j", said(await line(2)));
  await keys.press("Backspace");
  check("and Backspace takes it back", (await line(2)) === "", said(await line(2)));

  // ---- on ----
  await api.settings({ vimMode: true });
  const on = await waitFor("vim mode to come on", async () => {
    const now = await vim();
    return now.enabled && now.mode === "normal" ? now : null;
  }, { budgetMs: 10000, pollMs: 200 }).catch(async () => vim());
  check("the route's switch turns it on, in normal mode", on.enabled === true && on.mode === "normal", said(on));
  check("the status bar says so", on.label === "-- NORMAL --", said(on.label));
  check("the setting reads back on", (await api.settings()).vimMode === true);
  check("the caret is a block", (await caretShape()) === "block", said(await caretShape()));

  // ---- normal mode: keys move, and type nothing ----
  await placeCaret(4, 5);
  await keys.press("j");
  check("j moves the caret down a line", said(await caret()) === said([5, 5]), said(await caret()));
  await keys.press("w");
  check("w moves to the next word", said(await caret()) === said([5, 9]), said(await caret()));
  await keys.press("$");
  check("$ moves to the last character", said(await caret()) === said([5, 20]), said(await caret()));
  await keys.type("gg");
  check("gg moves to the first line", (await caret())[0] === 1, said(await caret()));
  await keys.press("G");
  check("G moves to the last line", (await caret())[0] === 13, said(await caret()));
  await keys.type("3G");
  check("a count is taken: 3G moves to line 3", (await caret())[0] === 3, said(await caret()));
  check("and none of those keys typed anything", String(await pageText()).replace(/\r\n/g, "\n") === MODULE.replace(/\r\n/g, "\n"));
  check("in the project either", (await projectText()).trimEnd() === MODULE.replace(/\r\n/g, "\n"));

  // ---- an edit reaches the project, and undo takes it back there ----
  await placeCaret(5, 5);
  await keys.type("dd");
  check("dd deletes the line", (await lineCount()) === 12 && (await line(5)) === "    alpha = 1", `${await lineCount()} lines; ${said(await line(5))}`);
  check("and the project's module follows", await projectFollows("the project to lose the line")
    && !(await projectText()).includes("Dim beta"));
  await keys.press("u");
  check("u puts it back", (await lineCount()) === 13 && (await line(5)) === "    Dim beta As Long", `${await lineCount()} lines; ${said(await line(5))}`);
  check("and the project's module follows", await projectFollows("the project to get the line back")
    && (await projectText()).includes("Dim beta"));
  await keys.type("yyjp");
  check("yy and p put a copy of the line below the next", (await line(7)) === "    Dim beta As Long", said(await line(7)));
  await keys.press("u");
  check("u takes the copy away", (await lineCount()) === 13, `${await lineCount()} lines`);

  // ---- a command still being typed shows in the status bar ----
  await keys.type("2d");
  check("a half-typed command is shown", (await vim()).keys === "2d", said((await vim()).keys));
  await keys.press("Escape");
  check("and Escape drops it", (await vim()).keys === "" && (await lineCount()) === 13, said(await vim()));

  // ---- insert mode, and the repeat ----
  await placeCaret(4, 9);
  await keys.type("ciw");
  check("ciw enters insert mode", (await mode()) === "insert" && (await vim()).label === "-- INSERT --", said(await vim()));
  check("with the caret a line again", (await caretShape()) === "line", said(await caretShape()));
  await keys.type("total");
  check("typing in insert mode types", (await line(4)) === "    Dim total As Long", said(await line(4)));
  await keys.press("Escape");
  check("Escape returns to normal mode", (await mode()) === "normal", said(await vim()));
  await placeCaret(5, 9);
  await keys.press(".");
  check(". repeats the change on another word", (await line(5)) === "    Dim total As Long", said(await line(5)));
  check("u, pressed until it is all undone, leaves the module as it was", await undoneTo(MODULE), said(await pageText()).slice(0, 160));

  // What the editor writes around the typing is not typing. A line typed in lower case is
  // respelled a moment after it rests; a repeat that had recorded the respelling as well would
  // type the line and its correction both.
  await placeCaret(7);
  await keys.press("o");
  check("o opens a line below, in insert mode", (await mode()) === "insert" && (await lineCount()) === 14, `${await mode()}, ${await lineCount()} lines`);
  await keys.type("dim delta as long");
  const respelled = await waitFor("the line to be respelled", async () => (await line(8)) === "    Dim delta As Long",
    { budgetMs: 8000, pollMs: 200 }).then(() => true, () => false);
  check("a line typed in lower case is respelled while still in insert mode", respelled, said(await line(8)));
  await keys.press("Escape");
  await placeCaret(12);
  await keys.press(".");
  const repeated = await waitFor("the repeat to be respelled", async () => (await line(13)) === "    Dim delta As Long",
    { budgetMs: 8000, pollMs: 200 }).then(() => true, () => false);
  check(". types the line once, and the editor respells it as it did the first", repeated && (await lineCount()) === 15,
    `${await lineCount()} lines; ${said(await line(13))}`);
  check("u, pressed until it is all undone, takes both lines away", await undoneTo(MODULE), `${await lineCount()} lines`);

  // ---- visual modes ----
  await placeCaret(6, 5);
  await keys.type("ve");
  check("v and a motion select, in visual mode",
    (await mode()) === "visual" && (await ask(`${EDITOR}.getModel().getValueInRange(${EDITOR}.getSelection())`)) === "alpha", said(await vim()));
  await keys.press("Escape");
  await keys.press("V");
  check("V is linewise", (await mode()) === "visual line" && (await vim()).label === "-- VISUAL LINE --", said(await vim()));
  await keys.press("Escape");
  await keys.type(">>");
  check(">> indents the line by the indent size", (await line(6)) === "        alpha = 1", said(await line(6)));
  await keys.type("<<");
  check("<< takes it back", (await line(6)) === "    alpha = 1", said(await line(6)));

  // ---- a macro, whose "(recording)" note is the status bar's too ----
  await placeCaret(6, 5);
  await keys.type("qa");
  check("qa says it is recording", (await vim()).prompt === "(recording)[a]", said(await vim()));
  await keys.press("x");
  await keys.press("q");
  const recorded = await vim();
  check("q ends the recording and leaves nothing in the bar", recorded.prompt === null && recorded.keys === "", said(recorded));
  await placeCaret(7, 5);
  await keys.type("@a");
  check("@a replays it on another line", (await line(6)) === "    lpha = 1" && (await line(7)) === "    eta = 2",
    `${said(await line(6))}, ${said(await line(7))}`);
  check("u, pressed until it is all undone, leaves the module as it was", await undoneTo(MODULE), said(await pageText()).slice(0, 160));

  // ---- the Ctrl keys ----
  await keys.press("q", { ctrl: true });
  check("Ctrl+Q starts a block selection", (await mode()) === "visual block", said(await vim()));
  await keys.press("[", { ctrl: true });
  check("Ctrl+[ is Escape", (await mode()) === "normal", said(await vim()));
  await placeCaret(1);
  await keys.press("d", { ctrl: true });
  check("Ctrl+D scrolls down, caret with it", (await caret())[0] > 1, said(await caret()));
  await keys.press("u", { ctrl: true });
  check("Ctrl+U scrolls back", (await caret())[0] === 1, said(await caret()));
  await placeCaret(6, 5);
  await keys.press("a", { ctrl: true });
  const all = await ask(`(() => { const s = ${EDITOR}.getSelection(); return [s.startLineNumber, s.endLineNumber]; })()`);
  check("Ctrl+A is still Select All, though Vim binds it", said(all) === said([1, 13]) && (await line(6)) === "    alpha = 1", said(all));
  await keys.press("Escape");
  await placeCaret(6, 5);
  await keys.press("i");
  await keys.press("d", { ctrl: true });
  check("in insert mode Ctrl+D is not Vim's: the line keeps its indent", (await mode()) === "insert" && (await line(6)).startsWith("    alpha"),
    `${await mode()}; ${said(await line(6))}`);
  await keys.press("[", { ctrl: true });
  check("and Ctrl+[ leaves insert mode", (await mode()) === "normal", said(await vim()));
  await ask(`(() => { ${EDITOR}.setSelection({ startLineNumber: 6, startColumn: 5, endLineNumber: 6, endColumn: 5 }); return true; })()`);

  // ---- the / and : lines ----
  await placeCaret(1);
  await keys.press("/");
  const searching = await vim();
  check("/ opens the search line in the status bar", searching.prompt === "", said(searching));
  await keys.type("gamma");
  check("what is typed goes into it, not the module", (await vim()).prompt === "gamma" && (await line(1)) === "Option Explicit", said(await vim()));
  await keys.press("Enter");
  check("Enter finds the first match", said(await caret()) === said([11, 9]) && (await vim()).prompt === null, said(await caret()));
  await keys.press("n");
  check("n finds the next", said(await caret()) === said([12, 5]), said(await caret()));
  await keys.press(":");
  await keys.type("4");
  await keys.press("Enter");
  check(":4 goes to line 4", (await caret())[0] === 4, said(await caret()));
  await keys.press(":");
  await keys.type("nosuchcommand");
  await keys.press("Enter");
  check("an unknown command says so in the status bar", /not an editor command/i.test((await vim()).notice), said((await vim()).notice));
  check("the editor has the focus again", (await ask(`${EDITOR}.hasTextFocus()`)) === true);

  // ---- can it be read ----
  const readable = await contrastByTheme(api, "#status");
  for (const theme of ["dark", "light"]) {
    check(`the status bar was read in ${theme}, the mode with it`, readable[theme].read >= 3, contrastSaid(readable[theme]));
    check(`and holds AA contrast in ${theme}`, readable[theme].fails.length === 0, contrastSaid(readable[theme]));
  }

  // ---- :wq, :w and :q ----
  const onScreen = () => waitFor("the module to be on screen", async () => (await lineCount()) === 13,
    { budgetMs: 15000, pollMs: 250 }).catch(() => {});
  const tabs = async () => (await api.ui()).workspace.groups.flatMap((group) => group.tabs).filter((tab) => tab.module === probe);
  const tabClosed = () => waitFor("the tab to close", async () => (await tabs()).length === 0, { budgetMs: 15000, pollMs: 200 })
    .then(() => true, () => false);

  // :wq first, while the module has never been saved. It is one request: sent as a save and
  // then a close, the close overtook the save and asked whether to save the changes.
  await placeCaret(6, 5);
  await keys.press("x");
  const unsaved = await waitFor("the tab to show as unsaved", async () => (await tabs()).some((tab) => tab.dirty),
    { budgetMs: 8000, pollMs: 200 }).then(() => true, () => false);
  check("the edited module's tab is unsaved", unsaved && (await line(6)) === "    lpha = 1", said(await line(6)));
  await keys.press(":");
  await keys.type("wq");
  await keys.press("Enter");
  const wroteAndQuit = await tabClosed();
  const asked = (await api.ui()).dialogs ?? [];
  check(":wq saves and closes without asking", wroteAndQuit && asked.length === 0, said(asked));
  await api.act("closeDialogs").catch(() => {});
  await api.waitUntilResponsive({ timeout: 20000 }).catch(() => {});
  await api.pane("open", { module: probe, project: project.projectId });
  await onScreen();
  check("and the edit it saved is still in the module", (await line(6)) === "    lpha = 1", said(await line(6)));

  await placeCaret(4);
  const mark = await api.mark(`vim.mjs ${stamp} :w`);
  await keys.press(":");
  await keys.type("w");
  await keys.press("Enter");
  // Command 3 is the host's Save (VbeCommands.Command.Save), logged as it is executed.
  const saved = await api.waitForLog("command: 3 executed", { since: mark.at, timeout: 10000 })
    .then((found) => found, () => null);
  check(":w runs the host's Save", (saved?.lines ?? []).length > 0, said(saved)?.slice(0, 200));
  await api.waitUntilResponsive({ timeout: 20000 }).catch(() => {});

  await keys.press(":");
  await keys.type("q");
  await keys.press("Enter");
  check(":q closes the tab", await tabClosed());

  // ---- off again, through the dialog's own path ----
  await api.pane("open", { module: probe, project: project.projectId });
  await onScreen();
  const viaPage = await api.act("settings", { key: "vimMode", value: "false" });
  const off = await waitFor("vim mode to go off", async () => {
    const now = await vim();
    return now.enabled ? null : now;
  }, { budgetMs: 10000, pollMs: 200 }).catch(async () => vim());
  check("the dialog's own switch turns it off", viaPage.did === true && off.enabled === false && off.label === "", `${viaPage.detail}; ${said(off)}`);
  check("the caret is a line again", (await caretShape()) === "line", said(await caretShape()));
  await placeCaret(2);
  await keys.press("j");
  check("and a plain key types again", (await line(2)) === "j", said(await line(2)));
  await keys.press("Backspace");
} finally {
  page.socket.close();
  await api.act("closeDialogs").catch(() => {});
  await api.settings({ vimMode: switchWas }).catch((error) => console.log(`  (the switch was left as it was: ${error.message})`));
  await probeModule.dispose();
}

process.exit(done());
