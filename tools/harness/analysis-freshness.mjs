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
 *
 * NOT A GATE STEP YET, and the reason matters more than the check does.
 *
 * On ordinary fixtures it passes every time. On the perf fixture, which holds a module at VBA's
 * 65,534-line ceiling, step 3 has failed roughly two runs in five: the caller is never flagged,
 * for two minutes, across some twenty passes. Three separate sampling errors in THIS FILE were
 * found and fixed while chasing it - a fixed sleep shorter than a pass, a wait for "any finding"
 * that caught a transient `undeclared-variable`, and a "byte-identical" write of this file's own
 * constant rather than of what the editor had stored - and the failure outlived all three.
 *
 * So one of two things is true and it is not yet known which: the check is still sampling
 * something it should not, or a pass that skips a project can make a stale read PERMANENT, since
 * a skipped pass does not update what it holds and nothing re-triggers it. The second would be a
 * defect in the product, introduced by the skip, and it is the reason this is written down rather
 * than left as a flaky test somebody re-runs.
 *
 * Run it by hand against the fixture you care about. Do not put it back in the gate until the
 * intermittent is attributed.
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

/*
 * WAITED FOR, NOT SLEPT THROUGH.
 *
 * This suite slept a fixed few seconds after each write and read the answer. That is a guess
 * about how long a pass takes, and the guess was calibrated on a fixture where a pass takes
 * under a second. Pointed at one holding a module at VBA's line ceiling, three checks failed and
 * every one of them was this: the pass was still running, the caller had not been re-analysed
 * yet, and the suite called the interim state the final one. The product was right the whole
 * time - it re-analysed after 8 seconds (2026-08-08).
 *
 * A check whose verdict depends on the size of the fixture is not checking the product.
 */

/** Waits until the engine stops being asked things, so a measurement starts from quiet. */
async function settle({ quietFor = 2500, budgetMs = 120_000 } = {}) {
  let lastSeen = -1;
  let quietSince = 0;

  for (let waited = 0; waited < budgetMs; waited += 1000) {
    const calls = (await api.engineCosts()).reduce((sum, row) => sum + row.calls, 0);
    if (calls === lastSeen) {
      quietSince += 1000;
      if (quietSince >= quietFor) { return true; }
    } else {
      lastSeen = calls;
      quietSince = 0;
    }

    await wait(1000);
  }

  return false;
}

/** Waits for a module's findings to satisfy a predicate, and answers what it last saw. */
async function awaitFindings(moduleName, predicate, budgetMs = 120_000) {
  let last = [];
  for (let waited = 0; waited < budgetMs; waited += 1500) {
    last = await findingsFor(moduleName);
    if (predicate(last)) { return last; }
    await wait(1500);
  }
  return last;
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

      // SETTLED, then read. Waiting for "no findings" answers instantly and means nothing: a
      // module that has never been analysed has no findings either, so the wait would end before
      // the thing it is waiting for has happened, and every step after it would be racing the
      // caller's first analysis.
      await settle();
      const clean = await findingsFor(CALLER);
      check("the caller is clean", clean.length === 0, `${clean.length} finding(s)`);

      console.log("\n2. a write-back that changes nothing at all\n");
      // From quiet, or the pass still running from step 1 is what gets counted.
      check("the engine went quiet before the measurement", await settle());

      // WHAT THE MODULE HOLDS, not what was sent to it. The editor stores a module its own way -
      // a trailing line comes and goes - so writing this file's constant back is a write of
      // DIFFERENT text and provokes exactly the pass this step is checking does not happen.
      // Byte-identical has to mean identical to what is there.
      const stored = (await api.readModule(CALLER, project.projectId)).text ?? "";
      await api.perf({ reset: true });
      const mark = (await api.log({ max: 1 })).next;
      await api.writeModule(CALLER, stored, project.projectId);
      await settle();
      const idle = await engineCalls();

      // THE PASS IS WHAT MUST BE SKIPPED, not every request. A write legitimately provokes a LIVE
      // analysis of the module just written - that is the squiggle following the developer's own
      // edit, and it is one module's text, not the project's. Asserting "the engine was asked
      // nothing" failed on a fixture holding a large module for that reason, and the thing it was
      // meant to be checking was working the whole time.
      //
      // What must not happen is the seed and the sweep over every module behind it.
      check(
        "the project is not re-seeded",
        idle["project/open"] === undefined,
        `${idle["project/open"]?.calls ?? 0} seed(s)`);

      const said = await api.log({ since: mark, max: 400 });
      check(
        "and the pass says it left the project alone",
        said.lines.some((line) => /is unchanged, so its .* finding\(s\) stand/.test(line)),
        "no pass reported skipping the project");

      check("the caller's findings survive the skip", (await findingsFor(CALLER)).length === 0);

      console.log("\n3. the callee grows a parameter; the caller is NOT touched\n");
      await api.writeModule(CALLEE, twoArguments, project.projectId);
      // The SPECIFIC finding, not any finding. Waiting for "something appeared" catches the pass
      // mid-flight: for a moment the caller is analysed against a project that does not hold the
      // callee yet and reports `undeclared-variable`, which is true of that instant and not of
      // the state being checked. Sampling an interim state and calling it the outcome is the
      // same mistake as sleeping a fixed time, wearing a different hat.
      const broken = await awaitFindings(
        CALLER,
        (found) => found.some((finding) => finding.code === "argument-count"));

      check(
        "the caller is re-analysed and reports the call",
        broken.some((finding) => finding.code === "argument-count"),
        broken.length === 0
          ? "no finding after two minutes: the stale answer was served"
          : JSON.stringify(broken.map((finding) => finding.code)));

      console.log("\n4. and back; the caller is still NOT touched\n");
      await api.writeModule(CALLEE, oneArgument, project.projectId);
      const healed = await awaitFindings(CALLER, (found) => found.length === 0);
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
      await settle();
      const idlePipe = await api.timeFeature("completions", { line: 8, column: 13 }, { n: 10 });
      await api.perf({ reset: true });
      const writing = api.writeModule(CALLER, `${CALLER_SOURCE}\r\n' provokes a pass\r\n`, project.projectId);
      const duringPass = await api.timeFeature("completions", { line: 8, column: 13 }, { n: 10 });
      await writing;
      const spent = await engineCalls();
      const longestCall = Math.max(0, ...Object.values(spent).map((row) => row.worstMs ?? 0));

      console.log(`     ${biggest.name}, ${biggest.lines} lines`);
      console.log(`     idle          ${idlePipe.medianMs}ms median, ${idlePipe.maxMs}ms worst`);
      console.log(`     during a pass ${duringPass.medianMs}ms median, ${duringPass.maxMs}ms worst`);
      console.log(`     longest single analyzer call in that pass: ${longestCall}ms`);

      /*
       * MEASURED AGAINST THE BOUND, not against a number somebody liked.
       *
       * The analyzer is one thread and a request in flight cannot be preempted, so the most a
       * completion can wait is the longest single call it might land behind, plus what it costs
       * when the pipe is free. That is the theoretical bound, and a build that stays under it is
       * doing the best this architecture allows.
       *
       * A fixed threshold cannot express that. `max(150, idle * 3)` passed on an 11,000-line
       * fixture and failed on one holding a module at VBA's ceiling, where a single unavoidable
       * analysis is most of a second - and it failed for describing the fixture rather than the
       * product. What a REGRESSION looks like is a completion waiting for the whole pass rather
       * than for one call of it, and that is what this catches at any size.
       */
      const bound = idlePipe.maxMs + longestCall + 150;
      check(
        "a completion waits for at most one analysis, not for the whole pass",
        duringPass.maxMs <= bound,
        `${duringPass.maxMs}ms worst against a bound of ${Math.round(bound)}ms `
        + `(${idlePipe.maxMs}ms idle + ${longestCall}ms longest call)`);
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
