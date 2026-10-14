/*
 * The plain editing keys, pressed as keys (#103). Enter makes a new line in the code editor and
 * Up and Down move the caret; on the Immediate pane's entry line Enter runs the line and Up and
 * Down walk its history. The two editors share one keybinding service, and a binding meant for
 * the entry line once took all three keys from the code editor too.
 *
 * Pressed through the page's DevTools protocol (Input.dispatchKeyEvent), which delivers a real,
 * trusted key event to the page and runs it through Monaco's keybinding resolution - the step
 * the api's `press` act skips, because it types the newline directly. Nothing goes to the
 * desktop: the event reaches only the page, after the editor under test has been focused.
 *
 * Run against the debug fixture:
 *   node tools\harness\editor-keys.mjs
 */

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { open, reporter, wait } from "./xlide-api.mjs";
import { loopbackUrl } from "./loopback-url.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

/** The DevTools port of the session the api client is aimed at. */
function devtoolsPort() {
  const home = join(process.env.LOCALAPPDATA ?? "", "xlide_vbide");
  const pid = Number(process.env.XLIDE_PID) || api.pid;
  return JSON.parse(readFileSync(join(home, `xlide-api-${pid}.json`), "utf8")).devtoolsPort;
}

async function connectToEditorPage(port) {
  const version = await (await fetch(loopbackUrl(port, "/json/version"), { redirect: "error" })).json();
  const socket = new WebSocket(version.webSocketDebuggerUrl);
  let nextId = 0;
  const waiting = new Map();
  socket.onmessage = (event) => {
    const message = JSON.parse(event.data);
    const pending = message.id && waiting.get(message.id);
    if (pending) {
      waiting.delete(message.id);
      if (message.error) { pending.reject(new Error(message.error.message)); } else { pending.settle(message.result); }
    }
  };
  await new Promise((settle, reject) => {
    socket.onopen = settle;
    socket.onerror = () => reject(new Error("could not reach the DevTools door"));
  });
  const send = (method, params = {}, sessionId) => new Promise((settle, reject) => {
    const id = ++nextId;
    waiting.set(id, { settle, reject });
    socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
  });

  const targets = await (await fetch(loopbackUrl(port, "/json"), { redirect: "error" })).json();
  const page = targets.find((one) => one.type === "page" && one.title === "xlide editor");
  if (!page) { throw new Error(`no editor page; saw ${targets.map((one) => one.title).join("|")}`); }
  const { sessionId } = await send("Target.attachToTarget", { targetId: page.id, flatten: true });
  return { socket, send: (method, params) => send(method, params, sessionId) };
}

const KEYS = {
  Enter: { key: "Enter", code: "Enter", windowsVirtualKeyCode: 13, text: "\r" },
  ArrowUp: { key: "ArrowUp", code: "ArrowUp", windowsVirtualKeyCode: 38 },
  ArrowDown: { key: "ArrowDown", code: "ArrowDown", windowsVirtualKeyCode: 40 },
};

console.log("the editing keys, pressed as keys\n");

const page = await connectToEditorPage(devtoolsPort());
const evaluate = async (expression) => (await page.send("Runtime.evaluate",
  { expression, awaitPromise: true, returnByValue: true })).result?.value;
async function pressKey(name) {
  const key = KEYS[name];
  await page.send("Input.dispatchKeyEvent", { type: "keyDown", ...key });
  await page.send("Input.dispatchKeyEvent", { type: "keyUp", ...key, text: undefined });
  await wait(150);
}

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
  await api.writeModule("Helper", original, project.projectId).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
