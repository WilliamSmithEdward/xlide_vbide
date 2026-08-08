/*
 * Findings stay true while the work behind them is skipped.
 *
 * A pass used to re-analyse every module of every project every time, however little had moved.
 * Adding one comment line to a 109-line module cost six analyses and 476ms, of which 446ms was
 * re-deriving findings for four modules that were byte-identical to the ones already on screen
 * (perf fixture, 2026-08-08). The pipe serves one request at a time, so that was also 446ms
 * during which a completion in another module waited: measured at 332ms, against 35ms idle.
 *
 * Two layers now skip that work. The shim leaves a project alone when no module's text has moved,
 * and the engine answers a module from its last analysis when the module's text AND the project
 * facts it depends on are both unchanged.
 *
 * The second one is where a wrong answer would hide, and this suite exists for it. A module's
 * findings depend on more than its own text: change a procedure's SIGNATURE in one module and
 * every call to it in every other module is right or wrong for a new reason, with none of those
 * modules' text having changed. A memo keyed on source alone reports the caller clean forever.
 *
 * That is a silent failure. Nothing goes red, nothing is slow, a squiggle that should be there
 * simply is not - and the developer's next hour is spent on a call the editor said was fine.
 *
 * Proven by failing: with the facts comparison removed from the engine's memo, step 3 reports
 * zero findings where it should report one (verified 2026-08-08).
 *
 * Runs against WHATEVER workbook is open. It brings its own two modules and takes them away
 * again, so it never edits a fixture's own code and never depends on one module being present.
 * The timing step needs a large module to mean anything and says so when it skips.
 *
 *   node tools\harness\analysis-freshness.mjs
 */

import { open } from "file:///F:/GitHub/xlide/xlide_vbide/tools/harness/xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const CALLEE = "XlideFreshCallee";
const CALLER = "XlideFreshCaller";

// The caller's text, written once and never touched again. Every finding that appears or clears
// below does so because the OTHER module moved, which is the only thing this suite is about.
const CALLER_SOURCE = [
  "Option Explicit",
  "",
  "Public Sub CallsAcross()",
  "    Dim r As Long",
  "    r = XlideFreshAdd(1)",
  "End Sub",
  "",
].join("\r\n");

const oneArgument = [
  "Option Explicit",
  "",
  "Public Function XlideFreshAdd(ByVal seed As Long) As Long",
  "    XlideFreshAdd = seed",
  "End Function",
  "",
].join("\r\n");

const twoArguments = [
  "Option Explicit",
  "",
  "Public Function XlideFreshAdd(ByVal seed As Long, ByVal extra As Long) As Long",
  "    XlideFreshAdd = seed + extra",
  "End Function",
  "",
].join("\r\n");

let passed = 0;
const failures = [];

function check(what, ok, detail) {
  if (ok) {
    passed += 1;
    console.log(`  ok    ${what}`);
  } else {
    failures.push(`${what}${detail ? ` (${detail})` : ""}`);
    console.log(`  FAIL  ${what}${detail ? ` (${detail})` : ""}`);
  }
}

async function findingsFor(moduleName) {
  const answer = await api.problems(moduleName);
  return answer.findings ?? answer.problems ?? [];
}

/** What the engine was asked to do since the last reset, by method. */
async function engineCalls() {
  const byMethod = {};
  for (const row of await api.engineCosts()) {
    byMethod[row.method] = { calls: row.calls, totalMs: row.totalMs };
  }
  return byMethod;
}

const project = await api.project();
if (!project?.projectId) {
  console.log("no workbook is open; start one with tools\\harness\\Start-Excel.ps1 -Fresh");
  process.exitCode = 2;
} else {
  // The biggest module already open, for the timing step. Its text is never changed.
  const biggest = (project.components ?? [])
    .filter((component) => component.kind === "module" && component.lines > 0)
    .sort((left, right) => right.lines - left.lines)[0];

  let made = [];

  try {
    console.log(`\nin ${project.projectId}\n`);

    console.log("1. a caller and a callee, agreeing\n");
    for (const name of [CALLEE, CALLER]) {
      const answer = await api.component("add", { kind: "module", name, project: project.projectId });
      if (answer?.ok) { made.push(answer.name ?? name); }
    }
    check("both modules were added", made.length === 2, made.join(", ") || "none");

    if (made.length === 2) {
      await api.writeModule(CALLEE, oneArgument, project.projectId);
      await wait(1500);
      await api.writeModule(CALLER, CALLER_SOURCE, project.projectId);
      await wait(3500);
      const clean = await findingsFor(CALLER);
      check("the caller is clean", clean.length === 0, `${clean.length} finding(s)`);

      console.log("\n2. a write-back that changes nothing at all\n");
      await api.perf({ reset: true });
      await api.writeModule(CALLER, CALLER_SOURCE, project.projectId);
      await wait(3500);
      const idle = await engineCalls();
      check(
        "the engine is not asked to analyse anything",
        idle["textDocument/diagnostics"] === undefined,
        `${idle["textDocument/diagnostics"]?.calls ?? 0} diagnostics call(s)`);
      check(
        "and the project is not re-seeded",
        idle["project/open"] === undefined,
        `${idle["project/open"]?.calls ?? 0} seed(s)`);
      check("the caller's findings survive the skip", (await findingsFor(CALLER)).length === 0);

      console.log("\n3. the callee grows a parameter; the caller is NOT touched\n");
      await api.writeModule(CALLEE, twoArguments, project.projectId);
      await wait(4500);
      const broken = await findingsFor(CALLER);
      check(
        "the caller is re-analysed and reports the call",
        broken.length === 1 && broken[0].code === "argument-count",
        broken.length === 0
          ? "no finding: the stale answer was served"
          : JSON.stringify(broken.map((finding) => finding.code)));

      console.log("\n4. and back; the caller is still NOT touched\n");
      await api.writeModule(CALLEE, oneArgument, project.projectId);
      await wait(4500);
      const healed = await findingsFor(CALLER);
      check("the caller's finding clears", healed.length === 0, `${healed.length} finding(s) left standing`);
    }

    console.log("\n5. a keystroke is not stuck behind a pass\n");
    if (!biggest || biggest.lines < 1500) {
      // Said rather than skipped in silence. On a small project a pass is short enough that the
      // stall this guards against cannot be produced, so a green result here would mean nothing.
      console.log(`     skipped: the biggest module open is ${biggest?.name ?? "none"} at `
        + `${biggest?.lines ?? 0} lines, and a pass over it is too short to stall anything.`);
      console.log("     run this against artifacts\\fixtures\\PerfFixture.xlsm for the real figure.");
    } else {
      await api.pane("open", { module: biggest.name, project: project.projectId });
      await wait(2500);
      const idlePipe = await api.timeFeature("completions", { line: 8, column: 13 }, { n: 10 });
      const writing = api.writeModule(CALLER, `${CALLER_SOURCE}\r\n' provokes a pass\r\n`, project.projectId);
      const duringPass = await api.timeFeature("completions", { line: 8, column: 13 }, { n: 10 });
      await writing;
      console.log(`     ${biggest.name}, ${biggest.lines} lines`);
      console.log(`     idle          ${idlePipe.medianMs}ms median, ${idlePipe.maxMs}ms worst`);
      console.log(`     during a pass ${duringPass.medianMs}ms median, ${duringPass.maxMs}ms worst`);

      // Generous on purpose. This guards against the head-of-line stall coming back, not a
      // particular timing: it was 332ms worst against 45ms idle before the skip, and the machine
      // this runs on is not quiet. A regression puts it back in the hundreds.
      check(
        "a completion during a pass is not stalled behind it",
        duringPass.maxMs < Math.max(150, idlePipe.maxMs * 3),
        `${duringPass.maxMs}ms worst against ${idlePipe.maxMs}ms idle`);
    }
  } finally {
    for (const name of made) {
      await api.component("remove", { name, project: project.projectId });
      await wait(800);
    }
    await wait(2000);

    const left = (await api.project(project.projectId)).components
      .filter((component) => made.includes(component.name));
    console.log(`\nthe modules it brought were taken away: ${left.length === 0}`);
    if (left.length > 0) {
      failures.push(`left ${left.map((component) => component.name).join(", ")} behind`);
    }

    console.log(`\n${passed} passed, ${failures.length} failed`);
    for (const failure of failures) {
      console.log(`  ${failure}`);
    }

    process.exitCode = failures.length === 0 ? 0 : 1;
  }
}
