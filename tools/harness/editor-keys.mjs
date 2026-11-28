/*
 * The plain editing keys, pressed as keys (#103). Enter makes a new line in the code editor and
 * Up and Down move the caret; on the Immediate pane's entry line Enter runs the line and Up and
 * Down walk its history. The two editors share one keybinding service, and a binding meant for
 * the entry line once took all three keys from the code editor too.
 *
 * Pressed through page-keys.mjs, which delivers a real, trusted key event to the page and runs
 * it through Monaco's keybinding resolution - the step the api's `press` act skips, because it
 * types the newline directly. Nothing goes to the desktop: the event reaches only the page,
 * after the editor under test has been focused.
 *
 * Run against the debug fixture:
 *   node tools\harness\editor-keys.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";
import { connectToEditorPage, devtoolsPort, pageKeyboard, plainTyping } from "./page-keys.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

console.log("the editing keys, pressed as keys\n");

const page = await connectToEditorPage(devtoolsPort(api));
const evaluate = async (expression) => (await page.send("Runtime.evaluate",
  { expression, awaitPromise: true, returnByValue: true })).result?.value;
const pressKey = pageKeyboard(page).press;
const restoreTyping = await plainTyping(api);

const original = (await api.readModule("Helper", project.projectId)).text ?? "";
try {
  await api.guard(true);
  await api.pane("open", { module: "Helper", project: project.projectId });
  await wait(1500);

  // ---- the code editor ----------------------------------------------------------------
  const codeEditor = `window.xlideBridge.workspace.activeEditor()`;
  const caret = async () => evaluate(`(() => { const p = ${codeEditor}.getPosition(); return [p.lineNumber, p.column]; })()`);
  await evaluate(`(() => { const e = ${codeEditor}; e.focus(); e.setPosition({ lineNumber: 2, column: 1 }); })()`);
  await wait(300);

  const linesBefore = await evaluate(`${codeEditor}.getModel().getLineCount()`);
  await pressKey("Enter");
  const linesAfter = await evaluate(`${codeEditor}.getModel().getLineCount()`);
  check("Enter in the code editor makes a new line", linesAfter === linesBefore + 1, `${linesBefore} -> ${linesAfter}`);
  check("and leaves the caret on it", (await caret())[0] === 3, JSON.stringify(await caret()));

  // From a caret put there, so a key that does nothing cannot pass for one that moved it.
  const placeCaret = (line) => evaluate(`${codeEditor}.setPosition({ lineNumber: ${line}, column: 1 })`);
  await placeCaret(4);
  await pressKey("ArrowUp");
  check("Up moves the caret up a line", (await caret())[0] === 3, JSON.stringify(await caret()));
  await placeCaret(4);
  await pressKey("ArrowDown");
  check("Down moves the caret down a line", (await caret())[0] === 5, JSON.stringify(await caret()));

  // ---- the Immediate entry line, which keeps its own meaning for the same keys --------
  const entry = `window.xlideBridge.shell.immediateInput`;
  await evaluate(`window.xlideBridge.runEditorCommand("xlide.panel.immediate")`);
  await wait(500);
  await evaluate(`(() => { const e = ${entry}; e.focus(); e.setValue("?1+1"); })()`);
  await wait(200);
  await pressKey("Enter");
  await wait(1500);
  check("Enter on the entry line runs it and clears it", (await evaluate(`${entry}.value`)) === "",
    JSON.stringify(await evaluate(`${entry}.value`)));
  await pressKey("ArrowUp");
  check("Up on the entry line brings the last line back", (await evaluate(`${entry}.value`)) === "?1+1",
    JSON.stringify(await evaluate(`${entry}.value`)));
  await evaluate(`${entry}.setValue("")`);

  // ---- and the code editor again, after the entry line has been used ------------------
  await evaluate(`(() => { const e = ${codeEditor}; e.focus(); e.setPosition({ lineNumber: 2, column: 1 }); })()`);
  await wait(300);
  const before = await evaluate(`${codeEditor}.getModel().getLineCount()`);
  await pressKey("Enter");
  check("Enter in the code editor still makes a line after the entry line ran one",
    (await evaluate(`${codeEditor}.getModel().getLineCount()`)) === before + 1);
} finally {
  page.socket.close();
  await restoreTyping();
  await api.writeModule("Helper", original, project.projectId).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
