/*
 * The Immediate pane's entry line answers the code window's questions (#97): the member list,
 * quick info and parameter info, in the scope the line will run in. Outside break mode that is
 * the module on screen; stopped, it is the stopped procedure, whose locals the native window
 * offers too. Asked through the page's own requests, and once through the editor's suggest
 * widget, so the list a developer sees is the list checked.
 *
 * Run against the debug fixture, which is the one that COMPILES:
 *   node tools\harness\immediate-completions.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();
const { check, done } = reporter();

const stopped = async () => ((await api.breakpoints()).mode ?? (await api.state()).debugMode) === "break";
async function until(what, predicate, budgetMs = 20000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(250);
  }
  throw new Error(`timed out waiting for ${what}`);
}

/** The labels the host offers at the end of an Immediate line. */
const labels = async (line) => (await api.ask(`window.xlideBridge
  .requestImmediateCompletions(${JSON.stringify(line)}, ${line.length})
  .then((items) => items.map((item) => item.label))`)) ?? [];

console.log("the Immediate entry line's completions, hover and call tips\n");

const text = (await api.readModule("Runner", project.projectId)).text ?? "";
const lines = text.split(/\r?\n/);
const stopLine = lines.findIndex((l) => l.includes("counter = counter + 1")) + 1;

try {
  await api.guard(true);
  await api.pane("open", { module: "Runner", project: project.projectId });
  await until("Runner shown", async () => (await api.state()).shownModule === "Runner");

  // ---- design mode: the module on screen ------------------------------------------------
  const app = await labels("? Application.");
  check("a member list after a dot, in design mode", app.some((l) => /^ActiveWorkbook$/i.test(l)),
    `${app.length} item(s): ${app.slice(0, 6).join(", ")}`);

  const bare = await labels("? Wal");
  check("the module's own procedures are offered", bare.some((l) => /^Walk$/i.test(l)),
    bare.slice(0, 12).join(", "));
  check("the probe the line is set into is never offered",
    !bare.some((l) => /^XlideImmediateLine$/i.test(l)) && !(await labels("? Xlide")).some((l) => /^XlideImmediateLine$/i.test(l)));

  // The witness the break-mode check below needs: outside a stop, Walk's locals are not in
  // scope, and the native window does not offer them either.
  const designLocals = await labels("? cou");
  check("a procedure's locals are not offered outside break mode",
    !designLocals.some((l) => /^counter$/i.test(l)), designLocals.slice(0, 12).join(", "));

  const tip = await api.ask(`window.xlideBridge.requestImmediateSignatureHelp("? Helper.Describe(", 18)
    .then((info) => info && info.label)`);
  check("parameter info on an open paren", typeof tip === "string" && /Describe/i.test(tip), String(tip));

  const quick = await api.ask(`window.xlideBridge.requestImmediateHover("? Helper.Describe(1)", 10)
    .then((hover) => hover && { signature: hover.signature, start: hover.start, end: hover.end })`);
  check("quick info names the member, its span mapped back onto the line",
    quick && /Describe/i.test(quick.signature) && quick.start === 9 && quick.end === 17, JSON.stringify(quick));

  // ---- the widget a developer sees ------------------------------------------------------
  const widget = await api.ask(`(async () => {
    const entry = window.xlideBridge.shell.immediateInput;
    window.xlideBridge.runEditorCommand("xlide.panel.immediate");
    entry.focus();
    entry.setValue("? Application.");
    entry.editor.trigger("probe", "editor.action.triggerSuggest", {});
    for (let tries = 0; tries < 30; tries += 1) {
      await new Promise((r) => setTimeout(r, 100));
      const rows = [...document.querySelectorAll(".suggest-widget .monaco-list-row")];
      if (rows.length > 0) {
        const seen = rows.map((row) => row.getAttribute("aria-label") || row.textContent).slice(0, 5);
        entry.editor.trigger("probe", "hideSuggestWidget", {});
        entry.setValue("");
        return seen;
      }
    }
    entry.setValue("");
    return [];
  })()`);
  check("the suggest widget opens over the entry line with the list", (widget ?? []).length > 0,
    JSON.stringify(widget));

  // ---- break mode: the stopped procedure's locals ---------------------------------------
  await api.breakpoint("Runner", stopLine, { project: project.projectId, state: "on" });
  await wait(800);
  await api.caret(lines.findIndex((l) => l.includes("Public Sub Walk")) + 1,
    { module: "Runner", project: project.projectId });
  await wait(500);
  await api.command("run");
  const reached = await until("break mode", stopped, 25000).catch(() => false);
  check("the run stopped at the breakpoint", Boolean(reached));

  if (reached) {
    await wait(800);
    const locals = await labels("? cou");
    check("stopped, the procedure's locals are offered", locals.some((l) => /^counter$/i.test(l)),
      locals.slice(0, 12).join(", "));
    const label = await labels("? lab");
    check("and its other locals", label.some((l) => /^label$/i.test(l)), label.slice(0, 12).join(", "));
    const member = await labels("? label.");
    check("a String local's members are not invented", Array.isArray(member), JSON.stringify(member.slice(0, 5)));
  }
} finally {
  await api.command("reset").catch(() => {});
  await wait(1000);
  await api.breakpoint("Runner", stopLine, { project: project.projectId, state: "off" }).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
}

process.exit(done());
