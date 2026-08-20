/*
 * Does the shim give back every wrapper it takes over the editor's objects?
 *
 * WHY THIS IS WORTH A SUITE OF ITS OWN. Every automation object this product touches is wrapped,
 * and the wrapper takes its own reference on top of the one the shim holds. A wrapper that is
 * never disposed is given back by the FINALIZER THREAD instead, and the editor's objects are
 * apartment-threaded: releasing one from the finalizer thread is not untidy, it is an access
 * violation, which ahead-of-time compilation cannot throw and so turns into a FailFast that ends
 * Excel.
 *
 * It also surfaces late and in the wrong place. On 2026-08-07 one missing dispose was reported
 * four times over as three different faults: an access violation blamed on this library, another
 * blamed on VBE7.DLL, and two heap corruptions blamed on ntdll, because a release on the wrong
 * thread corrupts COM's own bookkeeping and the damage is noticed by whoever touches it next.
 * Nothing connected them until the stack from the fourth one named ComObject.Finalize.
 *
 * So the leak is measured while it is still only a leak. Each operation runs many times over and
 * the live count has to come back to where it started, give or take the objects an operation is
 * entitled to keep. A count that grows with the iteration count is the crash, seen early.
 *
 * WHAT THIS CANNOT SEE, and it cost a second crash to learn: the count is of wrappers taken
 * through ComRuntime, so a wrapper the RUNTIME builds on its own is invisible here. One was, at
 * the bottom of a variant-to-text conversion, where `As<object>()` on a variant holding an
 * interface asked for a wrapper nobody owned. This suite read a perfectly balanced 13 across it,
 * every time, while Excel died on the finalizer thread. See lessons.md entry 40; the thing that
 * found it was the crash reporter, not this.
 *
 *   node tools\harness\com-leak.mjs
 *   node tools\harness\com-leak.mjs 40      # more rounds, for a slower leak
 */

import { open, wait } from "./xlide-api.mjs";
import { buildForm } from "./form-plan.mjs";

const api = await open({});
const rounds = Math.max(4, Number(process.argv[2] ?? 12) || 12);
const project = await api.project();

/*
 * A module that is actually in the ACTIVE workbook, rather than one this file was written beside.
 *
 * It named `Helpers` outright, which exists in the rename fixture and not in the language one, so
 * opening a second workbook made the sweep die on "no module named Helpers" -- the same mistake
 * the language suite made on the same afternoon, from the same cause: assuming the workbook.
 * A leak sweep should run against whatever is loaded, since a leak is not a property of a fixture.
 */
const sample = (project.components ?? [])
  .filter((one) => (one.kind ?? one.type) !== "document")
  .map((one) => one.name)
  .find(Boolean)
  ?? (project.components ?? [])[0]?.name;

if (!sample) {
  throw new Error(`the active workbook ${project.project} has no components to read`);
}

console.log(`sweeping against ${project.project}, sampling module ${sample}
`);

let passed = 0;
let totalRounds = 0;

/** Each row's handle delta, kept so one jump can be told apart from a leak spread over the sweep. */
const handleRows = [];
const failures = [];

function check(what, ok, detail) {
  if (ok) {
    passed++;
    console.log(`ok   ${what}`);
  } else {
    failures.push(`${what}${detail ? `: ${detail}` : ""}`);
    console.log(`FAIL ${what}${detail ? `\n     ${detail}` : ""}`);
  }
}

/**
 * What an operation is holding afterwards, in all three currencies that can run out.
 *
 * COM wrappers are the one that killed Excel, but a leak is not obliged to be made of them.
 * Windows handles are finite per process and their exhaustion is as fatal and as badly
 * attributed: the failure lands on whoever next asks for one, which is rarely the leaker.
 * Managed bytes are the mildest of the three and by far the noisiest, so they are printed for
 * the reader rather than asserted on.
 */
const held = async () => {
  // THE HANDLE COUNT IS NOISY, and the noise is bigger than the signal at these round counts.
  // Measured across one sweep: swings of -19, -7, -4 and +2 over eight rounds, from Excel's own
  // transient handles rather than from anything here. A single reading either side of an
  // operation would fail on a positive swing and pass on a negative one, which is a coin toss
  // wearing an assertion's clothes.
  //
  // The floor is the stable part. A transient handle is open at some sample and closed at
  // another, so the minimum of several readings is close to what is actually being held, and a
  // real leak raises the floor because a leaked handle is open at every sample.
  //
  // THREE SAMPLES, NOT FOUR, AND CLOSER TOGETHER. The original spacing cost 600ms per reading and
  // this is called twice a row across 35 rows: 42 seconds of the sweep's 123 were spent here, for
  // a floor that three readings establish as well as four. What the floor needs is more than one
  // look, not a long one.
  const samples = [];
  for (let i = 0; i < 3; i += 1) {
    samples.push(await api.stats());
    if (i < 2) { await wait(60); }
  }

  const last = samples[samples.length - 1];
  return {
    wrappers: last.comWrappersLive,
    handles: Math.min(...samples.map((s) => s.handleCount)),
    bytes: last.managedMemoryBytes,
  };
};

const live = async () => (await held()).wrappers;

/**
 * An operation run many times over, with the live count before and after.
 *
 * `allowance` is what the operation may legitimately still be holding once it is done: opening a
 * pane keeps the pane. Anything above that is per-iteration, which is the shape of a leak.
 */
async function repeat(what, allowance, body, howMany = rounds) {
  // A round first, so anything the operation sets up once is set up before the baseline.
  await body(0);
  await wait(250);

  const before = await held();
  for (let round = 1; round <= howMany; round += 1) {
    await body(round);
  }
  // Long enough for a write-back or a didChange to land, short enough that 35 rows do not cost a
  // minute in settling alone. Every operation above is awaited, so this covers what continues
  // AFTER the answer rather than the answer itself.
  await wait(500);
  let after = await held();
  let settled = "";

  // GROWTH IS RECOUNTED BEFORE IT IS CALLED A LEAK. The count is sampled at an instant, and the
  // session has background work of its own - an analysis pass, the placement poll - that takes a
  // wrapper and gives it back within its tick. A sample landing inside such a tick reads +1, and
  // on a session seconds old that is exactly when this sweep runs: the documents row failed a
  // gate on it while the sweep's own end-of-run reconciliation showed every wrapper released
  // (2026-08-12). One settle and one recount separates the two shapes, because a real leak is
  // still there a second later and a mid-tick hold is not. The allowance itself does not move.
  if (after.wrappers - before.wrappers > allowance) {
    await wait(1000);
    const recount = await held();
    if (recount.wrappers < after.wrappers) {
      after = recount;
      settled = ", settled after a recount";
    }
  }

  // Against howMany, not the file's default. The state-changing rows run fewer rounds, and
  // dividing by the default understated their per-round figure fourfold while the line above
  // announced a round count that had not happened. An instrument that misreports its own
  // denominator is the same failure as the two this suite was built to replace.
  const grew = after.wrappers - before.wrappers;
  const perRound = (grew / howMany).toFixed(2);

  const handles = after.handles - before.handles;
  const kb = Math.round((after.bytes - before.bytes) / 1024);

  console.log(`\n  ${what}: ${howMany} rounds, wrappers ${before.wrappers} -> ${after.wrappers} `
    + `(${perRound} per round), handles ${handles >= 0 ? "+" : ""}${handles}, `
    + `managed ${kb >= 0 ? "+" : ""}${kb}KB${settled}`);

  check(`${what} gives back the wrappers it takes`,
    grew <= allowance,
    `live grew by ${grew} over ${howMany} rounds, ${perRound} per round, allowance ${allowance}. ` +
    "A count that scales with the rounds is a wrapper reaching the finalizer thread.");

  // Handles are REPORTED per row and judged once at the end, over the whole sweep.
  //
  // They are worth watching: they are finite per process, and their exhaustion is as fatal as the
  // COM crash and as badly attributed, since the failure lands on whoever next asks for one. But
  // per row they are unjudgeable here. Excel opens and closes handles constantly and this product
  // is a guest in its process, so a single row's delta is mostly Excel. Taking the floor of four
  // samples cut the swing from ±19 to mostly zero, and what was left still tripped a per-round
  // threshold on rows that cannot leak, picking DIFFERENT rows on consecutive runs.
  //
  // A threshold that fires on noise is a coin toss wearing an assertion's clothes, so the
  // judgement moved to where the signal beats the noise: across the whole sweep, some 280
  // operations, where a leak of even one handle per operation is hundreds and Excel's churn is
  // still tens. The per-row numbers stay on screen, because that is where a reader looks to see
  // WHICH operation did it once the total says somebody did.
  totalRounds += howMany;
  handleRows.push({ what, handles, rounds: howMany });
}

const atStart = await held();
console.log(`resting live count: ${atStart.wrappers}, handles ${atStart.handles}\n`);

// Reading the project tree: the heaviest walk there is, one wrapper per component and several
// per project. If anything leaks, this is where it shows first.
await repeat("reading the project", 0, async () => {
  await api.project();
});

// The panes underneath, every one of them, with their code modules and their text.
await repeat("reading the native panes", 0, async () => {
  await api.native({ text: true });
});

// A module's text through the session's own reader.
await repeat("reading a module", 0, async () => {
  await api.readModule(sample, project.projectId);
});

// The windows collection, which is a COM collection walked by index.
await repeat("walking the editor's windows", 0, async () => {
  await api.windows();
});

// Moving the caret in the native pane, which resolves a pane and its code module each time.
await repeat("moving the caret", 0, async (round) => {
  await api.caret(1 + (round % 3), { module: sample, project: project.projectId, column: 1 });
});

// Breakpoint state, which reaches into the debugger's own objects.
await repeat("reading the breakpoints", 0, async () => {
  await api.breakpoints();
});

// And the whole doctor pass, which touches nearly everything at once.
await repeat("a doctor pass", 0, async () => {
  await api.doctor();
});

/*
 * EVERY OTHER READ ROUTE, because the seven above were chosen by hand and a leak does not care
 * which routes somebody thought to check.
 *
 * The defect that started this had been present since the type was written and reached the door
 * through whatever happened to call it. Picking likely-looking routes is how it survived: the
 * ones a developer reaches for are the ones already believed. So the sweep is exhaustive over
 * everything safe to call repeatedly, and anything left out is left out on purpose and named.
 *
 * Excluded, with reasons: `capture` returns a bitmap and is slow; `eval`, `await`, `assert`,
 * `bench`, `trip` and `pagecall` run scripts in the page rather than touching COM; `reload`
 * restarts the surface.
 *
 * The STATE-CHANGING routes are not excluded, and were only briefly: they do the most COM work
 * of anything here, so a guarantee that skips them is not a guarantee. They are swept below,
 * each paired with its own undo so the fixture comes back.
 */
const READ_ROUTES = [
  ["state", () => api.state()],
  ["documents", () => api.documents()],
  ["ui", () => api.ui()],
  ["layout", () => api.layout()],
  ["problems", () => api.problems()],
  ["outline", () => api.outline(sample)],
  ["locals", () => api.locals()],
  ["watches", () => api.watches()],
  ["immediate", () => api.immediate()],
  ["dialogs", () => api.dialogs()],
  ["journal", () => api.journal()],
  ["history", () => api.history()],
  ["messages", () => api.messages(5)],
  ["console", () => api.console(5)],
  ["perf", () => api.perf()],
  ["placement", () => api.placement()],
  ["engine", () => api.engineSource(sample)],
  ["breakpoints", () => api.breakpoints()],
  ["windows", () => api.windows()],
  ["inspect", () => api.inspect(".xlide-tab")],
  // `assert` reaches the debugger's own state through its claim predicates, which is COM the
  // read routes above never touch.
  ["assert stopped", () => api.assert("stopped", { timeoutMs: 200 })],
  ["assert shownModule", () => api.assert("shownModule", { timeoutMs: 200 })],
  ["assert localsHas", () => api.assert("localsHas", { value: "n", timeoutMs: 200 })],
];

for (const [name, call] of READ_ROUTES) {
  // eslint-disable-next-line no-await-in-loop
  await repeat(name, 0, async () => {
    await call().catch(() => null);
  });
}

/*
 * THE ROUTES THAT CHANGE STATE, which do the most COM work of anything here.
 *
 * Each is paired with its own undo, so a round leaves the fixture where it found it and the suite
 * can be run twice. That pairing is also what makes the count meaningful: a round that opens a
 * pane and does not close it is entitled to keep a wrapper, and an allowance big enough to cover
 * that is an allowance big enough to hide a leak.
 *
 * Fewer rounds than the read routes, because each round is a real editor operation with a real
 * settling time, and a leak that needs more than a few rounds to show is a leak the read sweep
 * would have found already.
 */
const changing = Math.max(2, Math.min(4, Math.floor(rounds / 3)));

await repeat("opening and closing a pane", 0, async () => {
  await api.pane("open", { module: sample, project: project.projectId });
  await wait(400);
  await api.pane("close", { module: sample, project: project.projectId, answer: "discard" });
  await wait(400);
}, changing);

// The host-originated direction: closeNative walks the editor's pane list per call, which is
// COM per pane, and the wrappers must all come home (route new 2026-08-12).
await repeat("closing a pane's native window", 0, async () => {
  await api.pane("open", { module: sample, project: project.projectId });
  await wait(400);
  await api.pane("closeNative", { module: sample, project: project.projectId });
  await wait(400);
}, changing);

// frame show reads MainWindow off the editor per call; frame close is a posted window message
// and takes nothing, but the close path it triggers runs the placement retreat and the palette
// follow, so the whole cycle is what has to come back to its resting count (route new
// 2026-08-12).
await repeat("closing the editor and reopening it", 0, async () => {
  await api.frame("close");
  await wait(600);
  await api.frame("show");
  await wait(600);
}, changing);

await repeat("setting and clearing a breakpoint", 0, async () => {
  await api.breakpoint(sample, 1, { project: project.projectId, state: "on" });
  await wait(250);
  await api.breakpoint(sample, 1, { project: project.projectId, state: "off" });
  await wait(250);
}, changing);

await repeat("reading and writing a module", 0, async () => {
  const held = (await api.readModule(sample, project.projectId)).text ?? "";
  await api.writeModule(sample, held, project.projectId);
  await wait(500);
}, changing);

await repeat("renaming a component and back", 0, async () => {
  await api.component("rename", { name: sample, newName: `${sample}Tmp`, project: project.projectId });
  await wait(600);
  await api.component("rename", { name: `${sample}Tmp`, newName: sample, project: project.projectId });
  await wait(600);
}, changing);

// The designer routes walk a form's whole control graph - every control, its font, its
// parent, its container's pages - which is more wrappers per round than anything else on this
// door, and every one must come home. The round builds the plan's form, reads it whole, and
// removes it, so the fixture comes back formless (routes new 2026-08-13).
//
// A FRESH NAME EVERY ROUND, not the plan's. A form name can be refused for the rest of the
// session - observed the day this row was written, cause unestablished, recorded in
// form-plan.mjs - and this sweep runs late in a session other suites have already worked
// over. The name is not what this row measures; the wrappers are.
let leakFormRound = 0;
await repeat("building, reading and removing a form", 0, async () => {
  const name = `LeakForm${++leakFormRound}`;
  await buildForm(api, project.projectId, name);
  await api.designer(name, project.projectId);
  await api.component("remove", { name, project: project.projectId });
  await wait(400);
}, changing);

// The test runner's whole COM road in one round: discovery walks every module's CodeModule,
// install writes the support module, a run injects the runner and dispatcher, calls
// Application.Run per test, and removes what it injected - each of them a wrapper this door
// takes and must give back.
let leakTestRound = 0;
await repeat("discovering and running a test", 0, async () => {
  const name = `LeakTest${++leakTestRound}`;
  await api.tests({ action: "install" });
  await api.component("add", { kind: "module", name, project: project.projectId });
  await api.writeModule(name, [
    "' @xlide-test",
    "Public Sub Green()",
    "    XlideAssert.IsTrue True",
    "End Sub",
  ].join("\r\n"), project.projectId);
  await api.tests({ action: "run", module: name, timeoutMs: 60000 });
  await api.component("remove", { name, project: project.projectId });
  await api.component("remove", { name: "XlideAssert", project: project.projectId });
  await wait(400);
}, changing);

// The markup language's vocabulary creates a BARE INSTANCE of every control coclass and walks
// each one's type library - fourteen objects this door MAKES rather than borrows, which is a
// shape nothing else here covers. Named without a module on purpose: the Form entry's designer
// walk is the row above's business, and this one is about the instances. They are cached per
// kind, so the rounds after the first create nothing at all, and the row asks both halves of the
// question - did the first round give its instances back, and does asking again cost anything
// (route new 2026-08-16).
await repeat("reading the markup vocabulary", 0, async () => {
  await api.markupVocabulary();
}, changing);

await repeat("changing a setting", 0, async () => {
  const was = (await api.settings()).formatIndentSize;
  await api.settings({ formatIndentSize: was === 4 ? 2 : 4 });
  await api.settings({ formatIndentSize: was });
}, changing);

/*
 * WHAT THIS SUITE DELIBERATELY DOES NOT DO: force a collection and see whether Excel survives.
 *
 * That was tried, on 2026-08-07, and it is a false-negative machine. A route was added to collect
 * and drain the finalizers on demand, on the theory that it would turn the crash from something
 * that arrives hours later into something that arrives now. Measured against a build carrying the
 * real defect and 8,734 leaked wrappers pending, the drain reported completely clean and the host
 * lived. Releasing an apartment-threaded object from the finalizer thread is only SOMETIMES
 * fatal, which is exactly why the crash took a day to attribute in the first place.
 *
 * So no suite here forces one. An instrument that agrees with the code and disagrees with the
 * product is worse than no instrument, because it is believed; and forcing a collection is a smell
 * in any case. The counter above is deterministic, costs two interlocked increments, and reported
 * 441 leaked wrappers from a single call to `project()` on the same broken build. That is the
 * instrument.
 *
 * The `drainfinalizers` ROUTE does still exist, and this said it had been removed until 2026-08-09.
 * It came back under a narrower contract: not a check but a BISECTOR. Run one operation, call it,
 * and if the host dies then that operation made the wrapper - which collapses the delay that made
 * five crashes across two days look unrelated. It answers about the outcome, the host being alive
 * or not, and about nothing else. Nothing in the gate calls it, on purpose, and audit-routes.mjs
 * holds that on the record rather than in a comment.
 */
const atEnd = await held();
console.log(`\nresting live count afterwards: ${atEnd.wrappers}, handles ${atEnd.handles}`);

/*
 * THE HANDLE VERDICT, over the whole sweep rather than per row.
 *
 * Several hundred operations by this point. A leak of one handle per operation is hundreds;
 * Excel's own churn across the same stretch is tens. That is the separation a per-row check never
 * had, and it is the whole reason this assertion is here and not up there.
 */
const handlesGrew = atEnd.handles - atStart.handles;
const perOperation = handlesGrew / Math.max(1, totalRounds);

console.log(`handles across ${totalRounds} operations: ${handlesGrew >= 0 ? "+" : ""}${handlesGrew}`
  + ` (${perOperation.toFixed(3)} per operation)`);

/*
 * PROVEN BOTH WAYS, on 2026-08-07, because a check that has only ever passed is not a check.
 *
 *   clean build:   -4, -26 and +1 across three runs
 *   leaking build: +538 over 190 operations, 2.832 each, and it failed
 *
 * The sabotage was a kernel handle taken per request and never given back, which is the shape of
 * a real leak: something the process owns and does not return.
 *
 * AND THE FIRST ATTEMPT AT SABOTAGE IS THE MORE USEFUL RESULT. It leaked on ONE route, and the
 * suite passed: eight handles across 190 operations is 0.04 each, far under this threshold. So
 * the honest limit of this check is that it catches a leak on a path the sweep exercises OFTEN,
 * and is blind to one confined to a rarely-called route. That is a fair trade for an assertion
 * that never fires on Excel's own churn, but it should be known rather than discovered later by
 * somebody trusting it further than it goes.
 */
/*
 * ONE JUMP IS NOT A LEAK PER OPERATION.
 *
 * The aggregate above went on to fail twice on a clean build, and both times for the same reason:
 * a SINGLE row jumped by about 350 while every other row sat at zero. It was `renaming a component`
 * once and `dialogs` the next time - both rows that legitimately make windows, and a window is
 * handles. Divided over the whole sweep that one jump reads as 0.9 per operation, which is most of
 * the way to the threshold on its own (2026-08-09).
 *
 * A real leak of one handle per request cannot hide in one row: every row calls the door, so every
 * row grows in proportion to its rounds. That is the difference this measures. The largest single
 * row is set aside and the REST is judged - the sabotage that proved this check (+538 over 190
 * operations, 2.832 each) fails just as loudly with its biggest row removed, because the growth
 * was everywhere.
 *
 * The outlier is still named, because a row that opens 350 handles and does not give them back
 * within the run is worth a developer's attention even when it is not this defect.
 */
const worstRow = handleRows.reduce(
  (worst, row) => (row.handles > worst.handles ? row : worst),
  { what: "none", handles: 0, rounds: 0 });
const spreadGrowth = handlesGrew - Math.max(0, worstRow.handles);
const spreadRounds = Math.max(1, totalRounds - worstRow.rounds);
const spreadPerOperation = spreadGrowth / spreadRounds;

if (worstRow.handles > 50) {
  console.log(`     the largest single row is ${worstRow.what} at +${worstRow.handles};`
    + ` the rest of the sweep grew ${spreadGrowth >= 0 ? "+" : ""}${spreadGrowth}`
    + ` over ${spreadRounds} operations (${spreadPerOperation.toFixed(3)} each)`);
}

check(`the sweep gives back the handles it takes, over all ${totalRounds} operations`,
  spreadPerOperation < 0.5,
  `handles grew by ${spreadGrowth} across ${spreadRounds} operations, ${spreadPerOperation.toFixed(3)} `
  + `each, with the largest single row (${worstRow.what}, +${worstRow.handles}) already set aside. `
  + "Excel's own churn over this stretch is tens, so a figure that scales with the operation count "
  + "is a handle this product opened and never closed. The per-row numbers above name which one.");

/*
 * GIVEN BACK AND ACTUALLY RELEASED MUST BE THE SAME NUMBER.
 *
 * This is the check that would have caught the defect every other check in this file missed for
 * the whole life of the product. `GiveBackWrapper` released with `(wrapper as IDisposable)?.
 * Dispose()`, and the wrapper StrategyBasedComWrappers hands back is a `ComObject`, which DOES
 * NOT IMPLEMENT IDisposable. The cast failed, `?.` swallowed it, and the counter incremented
 * anyway - so every wrapper this product ever took read as returned while its reference went to
 * the finalizer thread, where releasing an editor object FailFasts the host.
 *
 * Every row above passed throughout. Taken and given back were equal to the digit, the live count
 * sat at its resting 13, handles were flat. All true, and all measuring the INTENTION to release
 * rather than the release. Read `disposed: 0` against `givenBack: 1882` and there is nothing left
 * to argue about (2026-08-08, docs/com-wrapper-release.md).
 *
 * So the two are counted separately at the source and compared here. They can only diverge if a
 * wrapper is counted home without being released, which is the bug, exactly.
 */
const counters = await api.stats();
const notReleased = counters.comWrappersGivenBack - counters.comWrappersDisposed;
check("every wrapper given back was actually released",
  notReleased === 0,
  `${counters.comWrappersGivenBack} given back but only ${counters.comWrappersDisposed} released: `
  + `${notReleased} wrappers are counted home with their reference still held, and the finalizer `
  + "thread has them. This is the shape of the defect that killed Excel five times.");

console.log(`\n${passed} passed, ${failures.length} failed`);
for (const failure of failures) {
  console.log(`  ${failure}`);
}

process.exitCode = failures.length === 0 ? 0 : 1;
