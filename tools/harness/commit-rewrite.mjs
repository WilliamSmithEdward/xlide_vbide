/*
 * The VBE's rewrite of a committed line, shown as the native editor shows it. The VBE respaces
 * operators, closes a string left open at the end of a line, writes `1.0` as `1#` and `?` as
 * `Print` the moment a line is written; the page kept the text as typed until the module was
 * shown again, so it read differently from the code that compiles (2026-10-13). Each case is typed
 * into a scratch procedure, the line is left with Enter, and the page's line must then read what
 * the VBE holds. The line the caret is on is never rewritten under it.
 *
 * Run against the debug fixture:
 *   node tools\harness\commit-rewrite.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

const MODULE = "Helper";
const original = (await api.readModule(MODULE, project.projectId)).text ?? "";

// Typed without indentation: the procedure's own Enter supplies it.
const CASES = [
  ['Debug.Print  "What the heck", "Mannn', 'Debug.Print "What the heck", "Mannn"'],
  ["x=1+2", "x = 1 + 2"],
  ["x = 1.0", "x = 1#"],
  ["x = &hff", "x = &HFF"],
  ["x = \"a\"'trailing", "x = \"a\" 'trailing"],
  ["a=1:b=2", "a = 1: b = 2"],
  ["exit  sub", "Exit Sub"],
  ["x = .5", "x = 0.5"],
  ["? 1", "Print 1"],
];

const editor = (expr) => api.ask(`(() => { const e = window.xlideBridge.workspace.activeEditor(); const m = e.getModel(); ${expr} })()`);

console.log("the VBE's rewrite of a committed line, on the page\n");

try {
  await api.guard(true);
  await api.writeModule(MODULE, "Option Explicit\r\n\r\nSub CommitProbe()\r\n    Dim x, a, b\r\n    \r\nEnd Sub\r\n", project.projectId);
  await api.pane("open", { module: MODULE, project: project.projectId });
  await wait(1500);

  for (const [typed, expected] of CASES) {
    const line = await editor(`return m.getValue().split(/\\r?\\n/).lastIndexOf("End Sub");`);
    await api.ask(`(() => { const e = window.xlideBridge.workspace.activeEditor(); const m = e.getModel();
      e.setPosition({ lineNumber: ${line}, column: m.getLineMaxColumn(${line}) }); e.focus();
      // As plain edits, a character at a time: the editor's auto-closed quotes are a separate
      // behaviour, and here they would put a stray quote on the next line when Enter splits them.
      for (const ch of ${JSON.stringify(typed)}) { const c = m.getLineMaxColumn(${line});
        m.pushEditOperations([], [{ range: { startLineNumber: ${line}, startColumn: c, endLineNumber: ${line}, endColumn: c }, text: ch }], () => null);
        e.setPosition({ lineNumber: ${line}, column: m.getLineMaxColumn(${line}) }); }
      return true; })()`);

    // A pause long enough for the write, with the caret still on the line: not rewritten yet.
    await wait(1500);
    const held = await editor(`return m.getLineContent(${line}).trim();`);
    // Keywords are recased as they are typed, which is not the VBE's rewrite; the spacing is.
    check(`"${typed}" is left as typed while the caret is on it`, held.toLowerCase() === typed.trim().toLowerCase(), JSON.stringify(held));

    // Enter leaves the line, as a developer does.
    await api.ask(`(() => { const e = window.xlideBridge.workspace.activeEditor(); e.trigger("keyboard", "type", { text: "\\n" }); return true; })()`);
    let shown = "";
    for (let tries = 0; tries < 20 && shown !== expected; tries++) {
      await wait(250);
      shown = await editor(`return m.getLineContent(${line}).trim();`);
    }
    const vbe = ((await api.readModule(MODULE, project.projectId)).text ?? "").split(/\r?\n/)[line - 1]?.trim();
    check(`"${typed}" shows as the VBE wrote it once left`, shown === expected && vbe === expected,
      `page ${JSON.stringify(shown)}, VBE ${JSON.stringify(vbe)}`);
  }

  // The four copies still agree after all of it.
  await wait(2000);
  const page = await editor(`return m.getValue();`);
  const shadow = (await api.readModule(MODULE, project.projectId, { live: true })).text;
  const engine = await api.engineSource(MODULE, { text: true });
  const norm = (text) => String(text ?? "").replace(/\r\n?/g, "\n").replace(/\s+$/, "");
  check("the shadow matches the page", norm(shadow) === norm(page));
  check("the engine's mirror matches the page", !engine.engineHolds || norm(engine.engineText) === norm(page));
} finally {
  await api.writeModule(MODULE, original, project.projectId).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
