/*
 * The other direction of parity: the HOST moving on its own, and the surface following.
 *
 * Everything checked so far drove through the api and asked whether the native panes kept up.
 * This is the reverse, and it is the case a developer actually meets: Step Into crosses from
 * Runner into Helper, so the DEBUGGER activates a module the page was not showing and never
 * asked for. The surface has to catch up, or the developer steps into a procedure and watches a
 * different module's code.
 *
 * The native panes are covered by the surface, so a user cannot click one - which makes the
 * debugger the realistic driver of this direction, and the only one worth testing.
 */

import { open, reporter, wait } from "./xlide-api.mjs";

const api = await open({});
const project = await api.project();

const { check, done } = reporter();

async function until(what, predicate, budgetMs = 25000) {
  const deadline = Date.now() + budgetMs;
  while (Date.now() < deadline) {
    const answer = await predicate();
    if (answer) { return answer; }
    await wait(250);
  }
  throw new Error(`timed out waiting for ${what}`);
}

const stopped = async () => {
  const mode = (await api.breakpoints()).mode ?? (await api.state()).debugMode;
  return mode === "break";
};

console.log("Step Into, where the HOST changes module and the page must follow\n");

const text = (await api.readModule("Runner", project.projectId)).text ?? "";
const lines = text.split(/\r?\n/);
const callLine = lines.findIndex((l) => l.includes("answer = Helper.Describe(7)")) + 1;

try {
  await api.guard(true);

  await api.pane("open", { module: "Runner", project: project.projectId });
  await until("Runner shown", async () => {
    const ui = await api.ui();
    return ui.focus.model?.toLowerCase().endsWith("/runner") ? ui : null;
  });

  // Helper is CLOSED first rather than assumed closed. The page having no tab for it is the
  // harder half of this case - following the host means opening one - and a precondition that is
  // hoped for instead of established is a check that passes for the wrong reason. An earlier run
  // in the same session had left Helper open, and the assertion duly failed on the setup rather
  // than on the product (2026-08-08).
  const held = await api.ui();
  if (held.workspace.groups.flatMap((g) => g.tabs).some((t) => t.module === "Helper")) {
    await api.pane("close", { module: "Helper", project: project.projectId, answer: "discard" });
    await wait(2000);
  }

  const before = await api.ui();
  check("Helper has no tab before the step",
    !before.workspace.groups.flatMap((g) => g.tabs).some((t) => t.module === "Helper"),
    before.workspace.groups.flatMap((g) => g.tabs).map((t) => t.module).join(","));

  console.log(`  stopping on the call, line ${callLine}: ${JSON.stringify(lines[callLine - 1])}`);
  await api.breakpoint("Runner", callLine, { project: project.projectId, state: "on" });
  await wait(1200);

  await api.caret(lines.findIndex((l) => l.includes("Public Sub Calls")) + 1,
    { module: "Runner", project: project.projectId });
  await wait(800);
  await api.command("run");

  const reached = await until("break mode", stopped).catch(() => false);
  check("the run stopped on the call", Boolean(reached));

  if (reached) {
    const atCall = await api.inSync();
    check("everything agrees at the call", atCall.agreed,
      `native=${atCall.nativeModule} surface=${atCall.surfaceModule} page=${atCall.pageModule}`);

    console.log("\n  stepping INTO Helper:");
    await api.command("stepInto");

    // The host crosses into Helper. Waited for, not slept at.
    const crossed = await until("the native pane to be Helper", async () => {
      const below = await api.native();
      return (below.activeModule ?? "").toLowerCase() === "helper" ? below : null;
    }, 15000).catch(() => null);

    check("Step Into moved the NATIVE pane to Helper", Boolean(crossed),
      crossed ? `${crossed.activeModule} at ${crossed.caretLine}` : "it stayed in Runner");

    if (crossed) {
      // The whole point: the page never asked for Helper and must now be showing it.
      const followed = await until("the page to follow into Helper", async () => {
        const ui = await api.ui();
        return ui.focus.model?.toLowerCase().endsWith("/helper") ? ui : null;
      }, 12000).catch(() => null);

      check("the PAGE followed the host into Helper", Boolean(followed),
        followed ? "showing Helper" : `still showing ${(await api.ui()).focus.model}`);

      check("a tab was opened for Helper",
        (await api.ui()).workspace.groups.flatMap((g) => g.tabs).some((t) => t.module === "Helper"));

      const sync = await api.inSync();
      check("native, surface and page agree inside Helper", sync.agreed,
        `native=${sync.nativeModule} surface=${sync.surfaceModule} page=${sync.pageModule}`);

      const all = await api.parityAll();
      check("every open module's content still matches", all.agreed,
        all.stale.map((one) => one.module).join(","));

      // The Locals ghost republishes on its own tick, so the frame the panel is showing lags
      // the step by a beat. Wait for the panel's own CONTEXT line to be Helper's - a different
      // observable from the rows asserted below, so the check can still fail rather than only
      // time out - and then ask what it holds. Without this the read took the CALLER's frame
      // on a loaded gate session and failed there while every standalone run passed
      // (2026-08-15: "the Locals panel shows Helper's frame -- Runner").
      const locals = await until("the Locals panel to be on Helper's frame", async () => {
        const reading = await api.locals();
        return /helper/i.test(reading.context ?? "") ? reading : null;
      }, 12000).catch(() => api.locals());

      check("the Locals panel shows Helper's frame",
        (locals.rows ?? []).some((r) => (r.expression ?? "").toLowerCase() === "value"
          || (r.expression ?? "").toLowerCase() === "prefix"),
        (locals.rows ?? []).map((r) => r.expression).join(","));
    }
  }
} finally {
  await api.command("reset").catch(() => {});
  await wait(1200);
  await api.breakpoint("Runner", callLine, { project: project.projectId, state: "off" }).catch(() => {});
  await api.guard(false, { forget: true }).catch(() => {});
  await wait(800);
  console.log(`\n  left break mode: ${!(await stopped())}`);
}

process.exit(done());
