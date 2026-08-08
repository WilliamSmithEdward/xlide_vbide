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
 *   node tools\harness\com-leak.mjs
 *   node tools\harness\com-leak.mjs 40      # more rounds, for a slower leak
 */

import { open } from "file:///F:/GitHub/xlide/xlide_vbide/tools/harness/xlide-api.mjs";

const api = await open({});
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const rounds = Math.max(4, Number(process.argv[2] ?? 12) || 12);
const project = await api.project();

let passed = 0;
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

const live = async () => (await api.stats()).comWrappersLive;

/**
 * An operation run many times over, with the live count before and after.
 *
 * `allowance` is what the operation may legitimately still be holding once it is done: opening a
 * pane keeps the pane. Anything above that is per-iteration, which is the shape of a leak.
 */
async function repeat(what, allowance, body) {
  // A round first, so anything the operation sets up once is set up before the baseline.
  await body(0);
  await wait(600);

  const before = await live();
  for (let round = 1; round <= rounds; round += 1) {
    await body(round);
  }
  await wait(1200);
  const after = await live();

  const grew = after - before;
  const perRound = (grew / rounds).toFixed(2);

  console.log(`\n  ${what}: ${rounds} rounds, live ${before} -> ${after} (${perRound} per round)`);

  check(`${what} gives back what it takes`,
    grew <= allowance,
    `live grew by ${grew} over ${rounds} rounds, ${perRound} per round, allowance ${allowance}. ` +
    "A count that scales with the rounds is a wrapper reaching the finalizer thread.");
}

console.log(`resting live count: ${await live()}\n`);

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
  await api.readModule("Helpers", project.projectId);
});

// The windows collection, which is a COM collection walked by index.
await repeat("walking the editor's windows", 0, async () => {
  await api.windows();
});

// Moving the caret in the native pane, which resolves a pane and its code module each time.
await repeat("moving the caret", 0, async (round) => {
  await api.caret(1 + (round % 3), { module: "Helpers", project: project.projectId, column: 1 });
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
 * WHAT THIS SUITE DELIBERATELY DOES NOT DO: force a collection and see whether Excel survives.
 *
 * That was tried, on 2026-08-07, and it is a false-negative machine. A route was added to collect
 * and drain the finalizers on demand, on the theory that it would turn the crash from something
 * that arrives hours later into something that arrives now. Measured against a build carrying the
 * real defect and 8,734 leaked wrappers pending, the drain reported completely clean and the host
 * lived. Releasing an apartment-threaded object from the finalizer thread is only SOMETIMES
 * fatal, which is exactly why the crash took a day to attribute in the first place.
 *
 * So the route was removed rather than kept as a weaker second opinion. An instrument that agrees
 * with the code and disagrees with the product is worse than no instrument, because it is
 * believed; and forcing a collection is a smell in any case. The counter above is deterministic,
 * costs two interlocked increments, and reported 441 leaked wrappers from a single call to
 * `project()` on the same broken build. That is the instrument.
 */
const resting = await live();
console.log(`\nresting live count afterwards: ${resting}`);

console.log(`\n${passed} passed, ${failures.length} failed`);
for (const failure of failures) {
  console.log(`  ${failure}`);
}

process.exitCode = failures.length === 0 ? 0 : 1;
