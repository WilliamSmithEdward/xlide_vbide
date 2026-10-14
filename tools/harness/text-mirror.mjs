/*
 * The four copies of a module's text agree (#105): the page's model, the host's shadow, the
 * engine's mirror and the workbook. The engine answers every squiggle, completion and rename
 * from its mirror, so a mirror that drifts from the page draws phantom errors - and a rename
 * computed from it wrote the drifted text into the module.
 *
 * Each case changes the text from the page and from the host together, then compares the copies:
 *   - a line the VBE respaces when written, the module shown again, then typing (the #105 path);
 *   - one edit with several ranges, as a block comment makes;
 *   - typing while the module is changed from outside the page;
 *   - an edit stamped against a text the host has since replaced, which must be dropped;
 *   - an edit whose length disagrees with the page, which must heal from the page's text.
 *
 * Run against the debug fixture:
 *   node tools\harness\text-mirror.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

const norm = (text) => String(text ?? "").replace(/\r\n?/g, "\n").replace(/\s+$/, "");
const original = (await api.readModule("Helper", project.projectId)).text ?? "";
const codeModule = `Application.VBE.ActiveVBProject.VBComponents("Helper").CodeModule`;

async function copies() {
  const page = await api.ask(`window.xlideBridge.workspace.activeEditor().getModel()?.getValue() ?? null`);
  const shadow = (await api.readModule("Helper", project.projectId, { live: true })).text;
  const engine = await api.engineSource("Helper", { text: true });
  const host = (await api.readModule("Helper", project.projectId)).text;
  return {
    page: norm(page),
    shadow: norm(shadow),
    engine: engine.engineHolds ? norm(engine.engineText) : null,
    host: norm(host),
  };
}

function firstDifference(a, b) {
  const x = a.split("\n"), y = String(b).split("\n");
  for (let i = 0; i < Math.max(x.length, y.length); i++) {
    if (x[i] !== y[i]) { return `line ${i + 1}: ${JSON.stringify(x[i])} vs ${JSON.stringify(y[i])}`; }
  }
  return "same";
}

/** The page against the others, after the writes and passes have had time to land. */
async function agree(what, { host = true } = {}) {
  await wait(2500);
  const c = await copies();
  check(`${what}: the shadow matches the page`, c.shadow === c.page, firstDifference(c.page, c.shadow));
  check(`${what}: the engine's mirror matches the page, or it holds none`, c.engine === null || c.engine === c.page,
    c.engine === null ? "no mirror" : firstDifference(c.page, c.engine));
  if (host) {
    check(`${what}: the workbook matches the page`, c.host === c.page, firstDifference(c.page, c.host));
  }
  return c;
}

const typeAtEnd = (line, text, everyMs) => api.ask(`(async () => {
  const m = window.xlideBridge.workspace.activeEditor().getModel();
  for (const ch of ${JSON.stringify(text)}) {
    const col = m.getLineMaxColumn(${line});
    m.pushEditOperations([], [{ range: { startLineNumber: ${line}, startColumn: col, endLineNumber: ${line}, endColumn: col }, text: ch }], () => null);
    await new Promise((r) => setTimeout(r, ${everyMs}));
  }
  return true; })()`);

async function fresh() {
  await api.writeModule("Helper", original, project.projectId);
  await api.pane("open", { module: "Helper", project: project.projectId });
  await wait(1500);
}

console.log("the module's four copies agree\n");

try {
  await api.guard(true);

  // ---- the #105 path: respaced by the VBE, shown again, typed on --------------------------
  await fresh();
  await api.ask(`(() => {
    const m = window.xlideBridge.workspace.activeEditor().getModel();
    m.pushEditOperations([], [{ range: { startLineNumber: 2, startColumn: 1, endLineNumber: 2, endColumn: 1 },
      text: "Private Sub Spaced()\\n    Debug.Print Array(1,1)(0)+1\\nEnd Sub\\n" }], () => null);
    window.xlideBridge.workspace.activeEditor().setPosition({ lineNumber: 6, column: 1 });
    return true; })()`);
  await wait(2500);
  const respaced = (await api.readModule("Helper", project.projectId)).text.split(/\r?\n/)[2];
  check("the VBE respaced the written line, which is the case under test", respaced === "    Debug.Print Array(1, 1)(0) + 1",
    JSON.stringify(respaced));
  await api.pane("open", { module: "Runner", project: project.projectId });
  await wait(1200);
  await api.pane("open", { module: "Helper", project: project.projectId });
  await wait(1500);
  await agree("respaced and shown again");
  await typeAtEnd(3, " ' typed after", 10);
  await agree("typing on the respaced line", { host: false });

  // ---- one edit, several ranges ----------------------------------------------------------
  await fresh();
  await api.ask(`(() => {
    const m = window.xlideBridge.workspace.activeEditor().getModel();
    const at = (line) => ({ range: { startLineNumber: line, startColumn: 1, endLineNumber: line, endColumn: 1 }, text: "' " });
    m.pushEditOperations([], [at(3), at(4), at(5)], () => null);
    return true; })()`);
  await agree("one edit with three ranges", { host: false });

  // ---- typing while the module changes from outside --------------------------------------
  await fresh();
  const typing = typeAtEnd(1, " ' typed while the module changes underneath", 25);
  await wait(300);
  await api.immediate(`${codeModule}.InsertLines 1, "' inserted from outside"`);
  await typing;
  const outside = await agree("typing while lines are inserted from outside");
  check("and the workbook kept every line once, none duplicated", new Set(outside.host.split("\n").filter((l) => l.trim())).size
    === outside.host.split("\n").filter((l) => l.trim()).length, outside.host.split("\n").slice(0, 3).join(" | "));

  // ---- an edit made against a text the host has since replaced ----------------------------
  await fresh();
  const logMark = (await api.log({ max: 1 })).next;
  const before = await copies();
  await api.ask(`(() => {
    window.xlideBridge.transport.post({ type: "contentChanged", moduleName: "Helper", project: window.xlideBridge.hostActive?.project ?? undefined,
      revision: 999999, seq: -1, fullLength: 1,
      changes: [{ startLine: 1, startColumn: 1, endLine: 1, endColumn: 1, text: "STALE" }] });
    return true; })()`);
  const stale = await agree("an edit stamped with an older text", { host: false });
  const droppedLog = await api.log({ since: logMark, match: "arrived after text", max: 5 });
  check("the host said it dropped the stale edit", (droppedLog.lines ?? []).length > 0, JSON.stringify(droppedLog.lines ?? []).slice(0, 200));
  check("the stale edit was dropped, not applied", stale.shadow === before.shadow && !stale.shadow.includes("STALE"));

  // ---- an edit whose length disagrees: the shadow heals from the page ----------------------
  await fresh();
  const healMark = (await api.log({ max: 1 })).next;
  const seq = await api.ask(`window.xlideBridge.hostSeqs.get([...window.xlideBridge.hostSeqs.keys()].find((k) => /helper$/i.test(k)))`);
  await api.ask(`(() => {
    window.xlideBridge.transport.post({ type: "contentChanged", moduleName: "Helper", project: window.xlideBridge.hostActive?.project ?? undefined,
      revision: 999998, ${seq === undefined || seq === null ? "" : `seq: ${seq},`} fullLength: 3,
      changes: [{ startLine: 1, startColumn: 1, endLine: 1, endColumn: 1, text: "DRIFT" }] });
    return true; })()`);
  const healed = await agree("an edit whose length disagrees with the page");
  const healLog = await api.log({ since: healMark, match: "took the page", max: 5 });
  check("the host asked the page for its text and took it", (healLog.lines ?? []).length > 0, JSON.stringify(healLog.lines ?? []).slice(0, 200));
  check("the shadow took the page's text back, without the drifted edit", !healed.shadow.includes("DRIFT"));
} finally {
  await api.writeModule("Helper", original, project.projectId).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
