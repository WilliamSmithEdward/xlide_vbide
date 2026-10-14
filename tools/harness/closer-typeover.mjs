/*
 * Typing over the closer Smart Enter wrote. Enter after `If x Then` writes the `End If`; the
 * native editor writes none, so a developer used to it types their own - and had two, which does
 * not compile. A typed closer now replaces the inserted one when it would otherwise be a
 * duplicate, and is left alone when the code needs both (2026-10-13).
 *
 * Typed through the editor's own type command, Enter included, so Smart Enter runs as it does
 * for a person.
 *
 * Run against the debug fixture:
 *   node tools\harness\closer-typeover.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

const MODULE = "Helper";
const original = (await api.readModule(MODULE, project.projectId)).text ?? "";

const pageText = () => api.ask(`window.xlideBridge.workspace.activeEditor().getModel().getValue()`);
const lines = async () => (await pageText()).split(/\r?\n/).map((line) => line.trimEnd());
const count = (all, pattern) => all.filter((line) => pattern.test(line)).length;

/** Each string is a line typed and ended with Enter, in a fresh `Sub Probe()` body. */
async function typeInProcedure(typed, { editInsertedCloser = false } = {}) {
  await api.writeModule(MODULE, "Option Explicit\r\n\r\nSub Probe()\r\n    \r\nEnd Sub\r\n", project.projectId);
  await api.pane("open", { module: MODULE, project: project.projectId });
  await wait(1200);
  await api.ask(`(() => { const e = window.xlideBridge.workspace.activeEditor();
    e.setPosition({ lineNumber: 4, column: 5 }); e.focus(); return true; })()`);
  for (const [index, line] of typed.entries()) {
    await api.ask(`(async () => { const e = window.xlideBridge.workspace.activeEditor();
      for (const ch of ${JSON.stringify(line)}) e.trigger("keyboard", "type", { text: ch });
      e.trigger("keyboard", "type", { text: "\\n" });
      await new Promise((r) => setTimeout(r, 150)); return true; })()`);
    if (editInsertedCloser && index === 0) {
      // The developer edits the closer Smart Enter wrote: it is theirs from then on.
      await api.ask(`(() => { const m = window.xlideBridge.workspace.activeEditor().getModel();
        const at = m.getValue().split(/\\r?\\n/).findIndex((l) => /^\\s*End If\\s*$/.test(l)) + 1;
        m.pushEditOperations([], [{ range: { startLineNumber: at, startColumn: m.getLineMaxColumn(at), endLineNumber: at, endColumn: m.getLineMaxColumn(at) }, text: " ' mine" }], () => null);
        return at; })()`);
    }
  }
  await wait(1200);
  return lines();
}

console.log("typing over the closer Smart Enter wrote\n");

try {
  await api.guard(true);

  const ifBlock = await typeInProcedure(["If 1 Then", "x = 1", "endif"]);
  check("a typed endif over the inserted End If leaves one End If", count(ifBlock, /^\s*End If$/) === 1, ifBlock.join(" | "));
  const ifAt = ifBlock.findIndex((line) => /^\s*If 1 Then$/.test(line));
  const endAt = ifBlock.findIndex((line) => /^\s*End If$/.test(line));
  check("at the If's own indent", endAt > ifAt && ifBlock[endAt].match(/^\s*/)[0] === ifBlock[ifAt].match(/^\s*/)[0],
    JSON.stringify([ifBlock[ifAt], ifBlock[endAt]]));
  const caret = await api.ask(`window.xlideBridge.workspace.activeEditor().getPosition().lineNumber`);
  check("with the caret on the line below it", caret === endAt + 2, `caret ${caret}, End If at ${endAt + 1}`);

  const whileBlock = await typeInProcedure(["While 1", "x = 1", "Wend"]);
  check("a typed Wend over the inserted Wend leaves one", count(whileBlock, /^\s*Wend$/) === 1, whileBlock.join(" | "));

  // An inner block closed by hand inside an outer one: both closers are the code's.
  const nested = await typeInProcedure(["If 1 Then", "If 2 Then", "x = 1", "End If"]);
  check("nested blocks keep a closer each", count(nested, /^\s*End If$/) === 2, nested.join(" | "));

  // A closer the developer has edited is theirs: a second typed closer is not absorbed into it.
  const edited = await typeInProcedure(["If 1 Then", "x = 1", "End If"], { editInsertedCloser: true });
  check("an inserted closer edited by hand is never typed over", edited.some((line) => /End If ' mine/.test(line))
    && count(edited, /^\s*End If/) === 2, edited.join(" | "));

  // The page, the host's copy and the engine's agree after all of it.
  await wait(1500);
  const norm = (text) => String(text ?? "").replace(/\r\n?/g, "\n").replace(/\s+$/, "");
  const shadow = (await api.readModule(MODULE, project.projectId, { live: true })).text;
  check("the host's copy matches the page", norm(shadow) === norm(await pageText()));
} finally {
  await api.writeModule(MODULE, original, project.projectId).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
