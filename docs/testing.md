# Testing this product

How the checks here are built, which kind to reach for, and the things that have wasted a day.

The short version: most of what breaks in this codebase breaks *between* parts that each work, so
most of the value is in probes that drive the real thing rather than tests of the pieces.

## The four kinds, and when each is right

**Unit tests** (`tests/`, 259 of them, none need Excel). Pure logic with a real answer: the split
tree's arithmetic, the lexer against its corpus, registration plans, pixel maths. Fast enough to
run on every gate. If a thing can be a unit test it should be, and if it cannot, that is usually a
sign the logic is tangled with the host and worth extracting.

**Seam checks** (`Test-Seam` in the harnesses). Assert that a load-bearing line still exists in the
source, the built bundle, and the published bundle. Crude, and they only prove a call is present,
not that it works. They earn their place for the host half that needs Excel, and for the
stale-deploy tripwire: a rebuilt page that never reached the publish tree has bitten more than once.

**Headless page probes** (`*-page-probe.mjs`). Drive the built bundle in a plain browser context.
Good for page-only behaviour: the close-confirm modal's queueing, the object browser's search.
Needs no Excel, so it runs in the ordinary gate.

**Live probes** (`tools\harness\Test-*.ps1`, gated behind `verify.ps1 -Live`). Drive a real editor
through the xlide api. Reach for these when the defect lives in the interaction: window events,
focus, the host rewriting something underneath the page, a module going away mid-flight. The gate
launches its own Excel per fixture; run one by hand against a `tools\dev.ps1 -KeepOpen
-Configuration Debug` session. Debug builds only, and they find their session through
`Get-XlideApi` (tools\harness\XlideApi.psm1), never by taking the first Excel.

**Live suites** (`tools\harness\*.mjs`, the same gate). These are most of the live coverage and
this document did not mention them at all until 2026-08-11, which is worth saying plainly: a
reader following it would have concluded the PowerShell probes were the whole of `-Live`. They
drive the same door from node through `xlide-api.mjs`, and they are where the behaviours that
cost the most are pinned - where a squiggle lands after Format Module, whether the workbook, the
surface and the analyzer hold the same text after every operation, what the Immediate window
answers, whether an import refuses to overwrite unwritten edits, and whether the two sync
planners still decide identically. `verify.ps1` groups them by the fixture each needs.

## The rule that matters most

**A test that passes before the fix is not a regression test.**

It is worth being blunt about this because it is the failure this document exists to prevent. The
sequence is: reproduce the bug, watch the new check go red, apply the fix, watch it go green. A
check written after the fix and never run against the broken build proves only that the current
build does something.

This was learned the hard way on 2026-08-06. A live probe was written for stale problems surviving a
discarded close, five checks, all passing. Disabling the fix and republishing, it still passed all
five. It exercised the whole flow faithfully and could not tell the two builds apart, so as a
regression test it was worthless, and only running it against the broken build revealed that.

When a probe cannot distinguish the builds, one of three things is true, and it is worth knowing
which: the fix is unnecessary, the probe misses the trigger, or something else repairs the state
before the probe looks. Waiting twenty seconds for a condition hides all three.

It happened twice more on 2026-08-07, on the same defect, in one afternoon. A route was added to
force a collection and drain the finalizers, on the theory that it would make a random crash
deterministic: run against the broken build with 8,734 leaked wrappers pending, it reported
completely clean. A counter was added beside it, incrementing next to the disposal rather than by
it: run against a build with the disposal deliberately deleted, it read perfectly balanced. Both
were believed for a while, because both passed. **An instrument is not proven by passing on a good
build.** The third attempt reported 441 leaked wrappers per call on the broken build and zero on
the fixed one, which is the only reason it is in the gate.

## No leaks, and it is a gate step

**All development guarantees no memory leaks** (the developer, 2026-08-07). This is
release-blocking, not an aspiration, because a leaked COM wrapper in this product does not waste
memory: it kills the host. The editor's objects are apartment-threaded, and one released by the
finalizer thread is an access violation the runtime cannot throw, so it FailFasts Excel. One
missing `Dispose` leaked 441 wrappers per `project()` call and killed Excel four times in a day,
reported as three different faults against three different libraries with nothing connecting them.
[lessons.md](lessons.md) entry 36 has the whole of it.

```bash
node tools\harness\com-leak.mjs        # every read route, many rounds, live count before and after
```

The gate runs it as the `no leaks` step. `ComRuntime.TakeWrapper` and `GiveBackWrapper` are the
only two doors for wrapping an automation object and each does its own counting, so a caller
cannot dispose without counting or count without disposing. New COM code goes through them.

Anything that changes COM, native handles, subscriptions, timers or caches ships with its own leak
check, and that check is shown failing on a deliberately broken build before it is trusted.

## Traps met in this codebase

**PowerShell unwraps a single-element array on return from a function.** A helper ending in
`@($things | Where-Object {...})` hands back the bare element when exactly one matches, and
`.Count` on a `PSCustomObject` is `$null` rather than `1`. So "exactly one finding" reads as "no
findings" and every wait times out while the api answers correctly the whole time. Return
`, @(...)` with the leading comma, and prefer `@(Helper $x).Count` at the call site.

**The tab close is armed at pointerdown and fired at pointerup, and never listens for a click.**
This is deliberate: a press has to survive the element being rebuilt underneath it by a host echo.
A synthetic `.click()` does nothing at all, silently. Dispatch `pointerdown` then `pointerup` with
`bubbles: true`, since the handler is delegated to the strip.

**Writing a module through the api is not the same as typing in it.** A host write carries a new
baseline with it, so the module is not unsaved afterwards and a close will not ask. Any probe about
unsaved state has to type, through `globalThis.xlideBridge.workspace.activeEditor()`.

**Choose the fixture module, do not assume it.** `state.module` is empty whenever the surface has
tabs but no active one, which is how a fresh `-KeepOpen` session starts. Read the tab strip. And if
the probe asserts that findings *clear*, pick a module with no findings of its own, or the assertion
can never be satisfied.

**A probe that mutates the fixture must put it back in a `finally`**, or a failed run leaves broken
VBA behind and the next run starts from somewhere new.

**Assert what the check is testing, not the weather.** A wait for `modules: publish` in the log only
appears while modules are open, so the check passed or failed on whether the fixture happened to
have a module open. Watch for the specific thing.

**A check with NO readiness wait is the mirror-image failure**, and it hides better: it passes on a
quiet machine and fails in the gate. The Locals panel is fed by a ghost reader's own tick, so the
frame it shows lags a Step Into by a beat; `step-into-features` read it with no wait at all, took
the CALLER's frame once - in the gate, after eight other suites had loaded that session - and
passed every standalone re-run afterwards (2026-08-15). The fix is a wait on the panel's own
context line, which is not what the rows below it assert.

**A readiness wait must name a different observable from the check that follows it.** Otherwise the
check cannot fail, only time out, and the assertion has moved into the setup where nobody counts it.
Replacing a sleep with `waitFor` is the moment this goes wrong: the obvious condition to wait for is
usually the thing under test. [driving-excel.md](driving-excel.md) has both cases that caught this
out, and the third one worth knowing - **a sleep can be the only thing preventing a vacuous pass**,
so deleting it makes a suite faster and blind at the same time.

**When a faithful reproduction disagrees with you, believe it.** Chasing a reported drag fault on
2026-08-06, a synthetic `pointerdown`/`pointermove` sequence produced the dim, the compass and its
five petals, and no drop preview however precisely it aimed. That was written off twice: first as
the synthetic events not being a real drag, then as the session running a stale bundle after an
intermediate "it works now". Both were wrong. The reproduction was accurate the whole time, and the
fault was real: with no module open the editor area is `display:none` and measures nothing, so the
guard that skips a zero-sized region skipped the entire editor-edge branch.

Two habits would have found it in one pass instead of three. Measure the thing being aimed at
rather than assuming it is where it looks - `getBoundingClientRect()` on the region under test says
`0x0` immediately. And take an intermittent report as a clue about state rather than noise: "works,
then doesn't" was the difference between a module being open and not, which is exactly what the
reporter eventually said.

**Fixture drift is self-sustaining, and it fails in a place unrelated to the change under
test.** `module-sync` writes a spare `ByTheDialog` sub into Helper, drives the apply - which
reaches the workbook - and restores the module in the SESSION only, so the file kept the dirty
text. With that sub present, the builtIn planner's round trip leaves the Locals ghost on the
CALLER's frame, `step-into-features` fails, and `Invoke-SuiteGroup` aborts the group there,
before the suites that would have re-saved the file. So the next run inherits the same drift and
fails the same way: two gates in a row, in a debugger suite, while the change under test was a
designer canvas (2026-08-15). Three habits fall out. A restore puts the FILE back, not just the
module. When a gate fails somewhere unrelated twice, suspect the fixture before the change -
`New-DebugFixture.ps1` takes forty seconds and settles it. And an interrupted gate run is a
fixture-corrupting event, because the cleanup half never ran.

**A suite that SAVES the workbook has to clean the file, not just the session.** Removing a
temporary component at the end leaves the session as found and the FILE holding it, because the
save in the middle already wrote it there - and the next run's other suites open the file. The
designer suite saves on purpose (its Ctrl+S rows exist to prove the designer's save reaches the
workbook), so every run since those rows landed had been writing a spare UserForm into
DebugFixture.xlsm; the gate's discard probe takes the project's FIRST module as its subject and
found that form there on 2026-08-15. The cleanup now saves after the removal, and the suite
asserts the write happened rather than assuming it. Worth separating from what it looked like:
the probe's timeouts did NOT pin to the change under suspicion - the SAME bundle failed twice
and passed twice on fresh sessions of the same fixture - so the pollution is the defect that
was real here, and the probe's waits are a second, still-open question. Two runs are not a
verdict either way: the discriminator has to be run both ways more than once before a build is
blamed or cleared.

**A rect is reported for an element the viewport is not showing, and a hit test aimed there
answers whatever paints at those coordinates.** The designer suite passed 214 checks on a session
the fixture generator had started and failed two of them on the same build in a 704px-wide window:
the form is wider than the canvas box, `RegionPick` sits clipped to the right, and
`document.elementFromPoint` at its perfectly-good-looking centre returned the dock tab strip. It
reads as a regression and is a window size. Anything that aims a synthetic pointer must
`scrollIntoView({ block: "nearest", inline: "nearest" })` first, which is what a hand does; the
resize act had learned this a day earlier and the drag act had not, so the lesson only half
landed. Suspect it whenever a hit-test row fails and the arithmetic looks right.

**A closed form is not yet an unloaded one, and its designer goes with it.** For roughly 400ms
after `userform?action=close` is posted, `designer?module=` answers "no designer to read" - the
object is absent, not empty, so the route throws and the whole group aborts on it. Neither
obvious signal helps: `debugMode` reads `design` while the form is still standing, and the
running-forms list empties before the designer comes back. Wait for the designer itself
(`api.designer(...).then(() => true, () => false)`), then assert on what it says - a wait that
throws is not a check that fails, and the difference is a run that reports nothing.

**Counts beat megabytes for leaks.** `Test-Churn.ps1` asserts that editors, models, dock groups and
DOM nodes return to their starting numbers after two dozen cycles. An exact number that must return
to its starting value is a far better detector than a heap figure nobody can interpret.

## Running them

```powershell
tools\verify.ps1                  # the headless gate, under a minute
tools\verify.ps1 -Live            # plus the live suites; launches its own Excel, twice
tools\verify.ps1 -Deep            # the pre-release tier: two more Excel sessions on top of -Live
tools\dev.ps1 -KeepOpen -Configuration Debug   # a session to run a single probe against by hand
```

The gate waits for the session to be healthy before probing rather than sleeping a guess: a session
that has only just launched is still seeding, and a probe asserting a healthy session then fails on
the truth that it is not healthy YET, which is a real answer to the wrong question.

Debug builds only for anything using the xlide api. Release has none of it, and the gate verifies
that by inspecting the built binary rather than trusting the compiler flag.

## What the xlide api gives a probe

`state`, `problems`, `module` read and write, `layout`, `eval` and `await` against the live page,
`command` by name, `caret`, `breakpoint`, `capture` for a cropped screenshot, `log` with `waitMs`
to block until a line appears, and `assert` to state an expectation and wait for it. Full reference
in [xlide-api.md](xlide-api.md).

`await` and `log?waitMs` exist so a probe can wait for the thing it means instead of sleeping a
guess. A fixed sleep races the fixture's own boot traffic, and the race is usually won on the
machine that wrote it and lost everywhere else.
