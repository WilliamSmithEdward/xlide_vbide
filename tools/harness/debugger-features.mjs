/*
 * The run-and-stop cycle, with parity against the native editor at every stage.
 *
 * This is where parity matters most. Run, Step and ToggleBreakpoint act on the host's own ACTIVE
 * CODE PANE and the caret inside it - not on the page - so a surface that has drifted means a
 * breakpoint on the wrong line and a step into somewhere the developer is not looking. Checking
 * only the page here would pass while the debugger walked another module entirely.
 *
 * Run against the debug fixture, which is the only one that COMPILES:
 *   tools\New-DebugFixture.ps1
 *   node tools\harness\debugger-features.mjs
 *
 * Every run leaves break mode with a Reset, whatever happened, because a session left stopped
 * blocks everything a later test does and the next reader has no idea why.
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();

const { check, done } = reporter();

async function until(what, predicate, budgetMs = 20000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(250);
  }
  throw new Error(`timed out waiting for ${what}`);
}

/** Native, surface and page naming the same module: the definition of tested here. */
async function parity(after) {
  const sync = await api.inSync();
  check(`native parity ${after.padEnd(26)}`, sync.agreed,
    `native=${sync.nativeModule} surface=${sync.surfaceModule} page=${sync.pageModule}`);
  return sync;
}

/**
 * Stopped, as the product spells it.
 *
 * The mode is a STRING - "design", "run" or "break" - not a number. Comparing it to 2 is a
 * predicate that is false in every state including the one it is looking for, so the run
 * reported as never stopping while it was stopped the whole time (2026-08-08).
 */
const stopped = async () => {
  const mode = (await api.breakpoints()).mode ?? (await api.state()).debugMode;
  return mode === "break";
};

console.log("the debugger, with the native panes checked at every stage\n");

try {
  await api.guard(true);

  await api.pane("open", { module: "Runner", project: project.projectId });
  await until("Runner to be shown", async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith("/runner") ? ui : null;
  });
  await parity("with Runner shown");

  // It has to compile, or every command below is really a test of the dialog guard.
  //
  // `started` is checked first and separately. "No dialog appeared" was the whole of this
  // precondition until 2026-08-11, and a Compile item that was greyed raised no dialog either -
  // so the one check protecting the suite passed hardest in the case it exists to catch.
  const compiled = await api.compile();
  check("the compile command actually ran", compiled.started !== false,
    JSON.stringify(compiled).slice(0, 200));
  check("the fixture compiles", compiled.compiled !== false,
    JSON.stringify(compiled).slice(0, 200));

  const text = (await api.readModule("Runner", project.projectId)).text ?? "";
  const lines = text.split(/\r?\n/);
  const stopLine = lines.findIndex((l) => l.includes("counter = counter + 1")) + 1;
  console.log(`\n  stopping on line ${stopLine}: ${JSON.stringify(lines[stopLine - 1])}`);

  const set = await api.breakpoint("Runner", stopLine, { project: project.projectId, state: "on" });
  console.log(`  breakpoint: ${JSON.stringify(set)}`);
  await wait(1200);

  const recorded = (await api.breakpoints()).breakpoints ?? [];
  check("the breakpoint was recorded against Runner",
    recorded.some((r) => r.module.toLowerCase() === "runner" && r.lines.includes(stopLine)),
    JSON.stringify(recorded));

  await parity("after setting a breakpoint");

  console.log("\n  running Walk:");
  await api.caret(lines.findIndex((l) => l.includes("Public Sub Walk")) + 1,
    { module: "Runner", project: project.projectId });
  await wait(800);
  await api.command("run");

  const reached = await until("break mode", stopped, 25000).catch(() => false);
  check("the run stopped at the breakpoint", Boolean(reached));

  if (reached) {
    const below = await api.native();
    console.log(`  native pane: ${below.activeModule} at ${below.caretLine}:${below.caretColumn}`);
    check("the NATIVE pane is on the stopped module",
      (below.activeModule ?? "").toLowerCase() === "runner", below.activeModule ?? "none");
    check("the NATIVE caret is on the stopped line",
      below.caretLine === stopLine, `${below.caretLine}, wanted ${stopLine}`);

    // The STATUS BAR's claim against the native truth. The bar's caret readout is the one
    // place the developer sees where a Run would land, so a bar that drifts from the caret
    // the host acts on is misdirection - and until the snapshot carried these two fields the
    // bar could be driven all day and never asked (2026-08-12).
    const bar = await api.ui();
    check("the status bar reads the stopped line the native caret is on",
      (bar.statusPosition ?? "").startsWith(`Ln ${stopLine},`),
      `the bar says "${bar.statusPosition}", the native caret is on line ${below.caretLine}`);
    check("and names the module the stop is in",
      (bar.statusModule ?? "").toLowerCase().includes("runner"), bar.statusModule);

    await parity("while stopped");

    const locals = await api.locals();
    const rows = locals.rows ?? [];
    console.log(`  locals: ${JSON.stringify(rows.map((r) => `${r.expression}=${r.value}`)).slice(0, 220)}`);
    check("the Locals panel holds the procedure's variables",
      ["counter", "label", "ratio"].every((name) =>
        rows.some((r) => (r.expression ?? "").toLowerCase() === name)),
      rows.map((r) => r.expression).join(","));

    check("counter holds the value it was assigned before the stop",
      rows.some((r) => (r.expression ?? "").toLowerCase() === "counter" && String(r.value).trim() === "1"),
      rows.find((r) => (r.expression ?? "").toLowerCase() === "counter")?.value);

    console.log("\n  stepping:");
    await api.command("stepOver");
    await wait(1500);

    const afterStep = await api.native();
    check("Step Over advanced the native caret",
      afterStep.caretLine > stopLine, `${stopLine} -> ${afterStep.caretLine}`);
    await parity("after a step");
  }

  /*
   * THE TITLE BAR SURVIVES A RESET IT DID NOT NEED.
   *
   * Pressing Reset a second time changes no execution state and rewrites the caption: the editor
   * puts "Microsoft Visual Basic for Applications" back on its own window. The refresh used to
   * skip on a mode that had not changed, so the product's name stayed off its own window until
   * something else happened to move (2026-08-09, reported from the toolbar button).
   *
   * Read from the frame rather than from what the session thinks it wrote. Our record of the
   * caption is what was wrong.
   */
  console.log("\n  the title bar:");
  const captionNow = async () => (await api.state()).frameCaption ?? "";

  check("the caption is ours while stopped or just after",
    (await captionNow()).startsWith("XLIDE"), await captionNow());

  for (let press = 1; press <= 3; press += 1) {
    await api.command("reset");
    await wait(900);
    const caption = await captionNow();
    check(`reset #${press} leaves our name on the window`, caption.startsWith("XLIDE"), caption);
  }

  check("and it says design once the run is over",
    (await captionNow()).includes("[design]"), await captionNow());
} finally {
  // Whatever happened, leave break mode and take the breakpoint out: a session left stopped
  // blocks everything afterwards, and the next reader cannot tell why.
  await api.command("reset").catch(() => {});
  await wait(1200);
  const text = (await api.readModule("Runner", project.projectId)).text ?? "";
  const stopLine = text.split(/\r?\n/).findIndex((l) => l.includes("counter = counter + 1")) + 1;
  await api.breakpoint("Runner", stopLine, { project: project.projectId, state: "off" }).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
  await wait(800);
  console.log(`\n  left break mode: ${!(await stopped())}`);
}

process.exit(done());
