# The debug api

A local HTTP door into a running xlide session: ask what the editor is doing, drive it by
name, and read what its panels hold. It exists so diagnostics, tests, and performance work
can act semantically instead of posting mouse messages at measured pixels, and its route
shapes are deliberately product-grade because the same door is the intended path for
xlide_vscode and xlide_vbide to talk to each other one day.

**Debug builds only.** A Release shim carries no server, no port, and no route: the whole
feature is inside `#if DEBUG`, verified by publishing Release and finding none of its
strings in the binary.

## Finding an instance

The api belongs to an ADD-IN SESSION, which exists once per Excel process that has opened
the VBE. Each session writes a discovery file:

```text
%LOCALAPPDATA%\xlide_vbide\debug-api-{pid}.json
{"api":1,"port":61503,"token":"92a9...","devtoolsPort":61501,"pid":21916,"startedAt":"..."}
```

Every request goes to `http://127.0.0.1:{port}/{token}/{route}`. The port and token are
random per session; the token is the only gate, and a wrong one gets a 404.

**Several Excels.** Two situations look alike and are not:

- **Several workbooks, one Excel process.** The usual case. Excel reuses its process, so
  there is ONE session and one api, and the workbooks are told apart per request with
  `project=` (see the module route). `windows` lists every workbook's panes.
- **Several Excel processes** (`excel /x`, or separate logons). Each process that opens the
  VBE gets its own session, its own discovery file, its own port and token, and its own
  DevTools port. Nothing is shared and nothing collides.

A discovery file outlives a killed Excel, so answering is the only proof of life. Each new
session sweeps files whose process is gone, and clients should probe rather than trust.
`tools\harness\xlide-api.mjs` does all of this: `discover()` lists live instances,
`open({ workbook })` picks one, and it refuses to guess when the choice is ambiguous.

## Routes

Read routes are GET, acting routes are POST. Everything answers JSON except `capture`.
Routes that touch the session cross to the host thread with a three second deadline and
answer `{"error":"the host thread did not answer in time"}` rather than hanging.

| Route | Method | Arguments | Answers |
| --- | --- | --- | --- |
| `state` | GET | | shown module and project, debug mode, unwritten edits, engine up, frame and document-area handles and rects, palette open and visible, surface ready, DevTools port |
| `windows` | GET | | every editor window: type, caption, visible |
| `stats` | GET | | uptime, managed and working memory, handle count, GC counts, placement pass counts and timings, host-thread marshal count and timings, log lines, poll interval, message totals |
| `log` | GET | `since`, `match`, `max`, `waitMs` | a slice of the shim log and the next byte offset. With `waitMs` it BLOCKS until a matching line appears, so an event can be awaited rather than slept for |
| `doctor` | GET | | the start-of-session checklist: shim and bundle build times, the page's own build stamp, engine and ghost readers attached, dialogs standing. `findings` is empty when nothing is wrong |
| `perf` | GET | | recent raw placement and marshal durations, for medians and percentiles |
| `journal` | GET | `lines` | one capture of a whole moment: state, dialogs, counters, recent log, recent page traffic |
| `history` | GET | | every request this door has served, and a script that replays them |
| `assert` | POST | `that`, `value`, `timeoutMs` | states an expectation and waits for it, answering with what was actually seen |
| `messages` | GET | `last` | recent page traffic both directions, per surface |
| `problems` | GET | `module` | the analyzer's current findings |
| `locals` | GET | | the Locals panel's context and rows |
| `watches` | GET | | whether stopped, and the Watch panel's rows |
| `module` | GET | `name`, `project` | a module's text, read through the session's reader |
| `capture` | GET | `window=frame\|palette` | a BMP of the window, through PrintWindow |
| `module` | POST | `name`, `project`, body = text | writes the module through the session's writer, with the baseline and engine corrections a host rewrite carries |
| `dialogs` | GET | | native dialogs standing now, with their buttons, and how long the host thread has been quiet. Needs no host thread, so it answers while the editor is stuck |
| `eval` | POST | `surface`, `waitMs`, body = script | runs script in the live page and returns its result as JSON. A PROMISE is awaited, so `(async () => ...)()` works |
| `await` | POST | `surface`, `waitMs`, body = predicate | polls a predicate IN the page until it is true, answering `met` and how long it took. One request instead of a caller's poll loop |
| `layout` | GET | | the whole visible arrangement: dock sections and their groups, editor groups and their tabs, sizes, open documents, whether a drag is live |
| `reload` | POST | `waitMs` | reloads the page and waits for it to come back, answering with the bundle stamp it is now running and whether that is behind the one on disk |
| `dismiss` | POST | `button`, `caption` | clicks a dialog button by name. Explicit, so it will press OK if asked |
| `command` | POST | `name`, `keep` | runs an editor command by name (`VbeCommands.ForName`). `keep=1` exempts any dialog it opens from the guard below |
| `caret` | POST | `line`, `column`, `module`, `project` | puts the caret there, navigating first when a module is named |
| `breakpoint` | POST | `module`, `line`, `project`, `state=on\|off` | goes to the line and sets, clears, or toggles a breakpoint |
| `immediate` | POST | `text` | schedules an Immediate-window evaluation, fire and forget |
| `placement` | POST | | forces a placement pass |

### Awaiting, rather than sleeping

`eval` awaits a promise. The browser's own `ExecuteScript` hands back `{}` for one, which
looks exactly like a page fault: every async probe written against this door returned an
empty object and was read as broken until the shape was recognised. The script is evaluated
inside a wrapper that stashes a pending promise on the page and returns a ticket, and the
ticket is collected on a poll until it settles. So this works:

```powershell
Invoke-RestMethod "$api/eval" -Method Post -Body '(async () => { await something(); return answer; })()'
```

`await` takes a predicate instead of a script and polls it IN the page until it is true,
answering `met` and `elapsedMs`. It replaces the caller-side poll loop, which costs a round
trip per tick and invites a `Start-Sleep` chosen by guess — the loops written during the
workspace work raced the thing they were watching more than once. The elapsed time is the
interesting half of a PASS: a condition met in 4ms was already true, one met in 4 seconds
arrived.

One trap is written into the implementation. The page's content policy exempts the browser's
own synchronous evaluation, but NOT a later callback: a waiter that evaluated its predicate
string inside `setTimeout` was refused ("unsafe-eval is not an allowed source") and reported
every condition as unmet. The predicate is compiled to a function once, up front, and the
timer calls the function.

### Seeing the arrangement

`layout` answers with the whole visible shape in one request: each dock section with its
group tree and each group's tabs, the editor groups with their module tabs, the sizes, the
open documents, and whether a drag is live. It exists because building the docking layout
needed a dozen ad-hoc `eval` measurements per question — and it earned itself the day it
landed, answering "why is the Problems pane on the left?" in one call. (It was a probe that
had dragged it there and not put it back.)

`reload` reloads the page and WAITS for it to come back, answering with how long that took,
the build stamp the page is now running, and whether that stamp is behind the bundle on
disk. The manual version — reload, sleep a guess, hope — was run a dozen times in one
afternoon, and a guess that is too short reports on the page that is going away.

Two arguments deserve their reasons.

`breakpoint` takes `state=on|off` because a bare toggle is not safe for a script: a retry
clears what the first call set, which cost a live run its breakpoint during this api's own
verification. Without the argument it still toggles, the way the key does.

`immediate` only SCHEDULES. A statement that hits a breakpoint does not return until the
developer continues, so an api that waited for it would jam its own connection.

## Reading a break

The round trip the debugger milestone needs, and what `Test-DebugApi.ps1` walks:

1. `POST module?name=X` with the code, so the module has something to run.
2. `POST breakpoint?module=X&line=N&state=on`.
3. `POST caret?module=X&line=M` with M inside the procedure to run.
4. `POST command?name=run`. Expect this request to TIME OUT: a run that reaches a
   breakpoint does not let the host thread answer until it gets there. The timeout is not a
   failure, and the next step is the real assertion.
5. Poll `state` until `debugMode` reads `break`.
6. `GET locals` and `GET watches`.
7. `POST command?name=reset`.

Step 3 is not optional and not obvious. Every editor command acts on the CARET, and the host
copies the surface's caret into the native pane before running one, so a Run with the caret
on line 1 does what the editor always does there: it opens the Macros dialog and waits for a
person. `revealLine` scrolls without moving the caret and cannot substitute.

Two ways NOT to start the run, both learned the hard way on 2026-08-06:

- `Application.Run` from a process you might kill. The call BLOCKS inside the break, and
  killing its caller while Excel is suspended in it takes Excel down. It was a background
  job with a `Remove-Job -Force`; Excel crashed and restarted itself.
- `Application.OnTime`, which is silently unreliable here: it needs Excel idle and the macro
  name resolvable at fire time, and a module written seconds earlier did not always qualify.
  It works often enough to look correct and fail a probe at random.

Locals and watches come from the ghost reader thread's published snapshots, so these routes
never touch the accessibility layer themselves and cannot disturb a break. That matters: an
out-of-process accessibility client that dumps a ghost palette during a break can reset the
project (lesson 33), which is exactly why these routes exist instead.

## Never blocked by a modal

A modal dialog owns the editor until somebody answers it, and the two worst evenings this
project has had both ended that way: an Add Watch dialog whose filler mis-parsed its
arguments, and a Macros dialog raised by a Run with the caret on line one. Both times the
editor simply stopped, and nothing could say why.

Three things now stand between a modal and a wedged session.

**It can be seen.** `dialogs` enumerates windows, which needs no host thread, so it answers
while every other route would time out. It returns each dialog's caption and buttons, plus
`heartbeatAgeMs` - how long since the host thread last completed a poll tick. A large
heartbeat age with a dialog standing is a stuck editor, stated as two numbers.

**It is cleared automatically, if the door raised it.** Every request that touches the
session first sweeps dialogs this door is answerable for, pressing one SAFE button: Cancel,
then Close, then No. Never OK, Yes, Save, Delete, or Run - a dialog nobody read must not be
agreed with, and every safe button means "as you were". A request that means to open a
dialog passes `keep=1` and what it opens is exempt.

**A dialog you opened is never touched.** Attribution is the whole safety property, and it
took three attempts. A snapshot taken as a request ends catches nothing, because the dialog
arrives microseconds after the command returns. Comparing against "whatever stood when the
door last looked" then cancelled a developer's own Add Watch between requests - the one
outcome worth avoiding entirely. What works: after each request a pool thread looks again at
250ms, 750ms, and 1750ms, and the door owns only what appears in that window. A dialog that
opens while the door is idle is nobody's business but yours, however long it stands.

Verified in `Test-DebugApi.ps1`: a Run with a bad caret raises the Macros dialog, the door
sees it with its six buttons, clears it on the next request, and the heartbeat returns - and
separately, an Add Watch opened outside the api survives repeated api traffic untouched.

The mechanism behind all of this, and the rules for opening a modal at all, are in
[working-with-modals.md](working-with-modals.md).

## Stating expectations, capturing moments, replaying sessions

`assert` takes a named claim and waits for it: `stopped`, `running`, `surfaceReady`,
`shownModule`, `noDialogs`, `localsHas`, `watchHas`, `problemFree`, `responsive`. It answers
with what it SAW as well as whether the claim held, which is the half a bare false leaves
out - `{"held":false,"claim":"shownModule","expected":"NoSuchModule","saw":"BrokenModule"}`
diagnoses itself.

`journal` captures state, standing dialogs, the counters, the recent log, and the recent page
traffic in ONE request. Evidence gathered request by request describes several different
moments, and the interesting one is usually the one that has already passed.

`history` hands back every request the door has served, plus a runnable script of them. After
an investigation by hand, the useful sequence is normally reconstructed from a scrollback and
gets a step wrong; this turns a bug found by hand into a probe by copying.

## A note on the heartbeat

`heartbeatAgeMs` is the age of the host thread's last periodic tick, and it means "blocked"
only while something should be ticking. An idle editor stops polling by design, so a quiet
session legitimately shows a large age - `doctor` therefore only reports it as a finding when
polling is expected, after the first version of that check cried wolf on a quiet editor.

## Awaiting instead of sleeping

`log?match=...&waitMs=15000` returns the moment a matching line is written, or when the wait
expires. That removes the whole sleep-and-hope class of harness bug: a probe that slept a
guessed interval is slow when the guess is generous and flaky when it is not, and it can
never say whether the thing it wanted actually happened. Measured: a wait for a publish line
returned in 3.4 seconds, when the event occurred, not at the 15 second limit.

```js
const api = await open({ workbook: "scratch.xlsm" });
await api.command("save");
await api.waitForLog("modules: publish", { timeout: 10000 });
```

## When the page itself throws

A page exception is invisible without a DevTools client attached, which is never the case
during a live test: the surface misbehaves, the shim log says nothing, and all anyone has is
a description of symptoms. The page now forwards `error` and `unhandledrejection` to the host,
where they appear in the log as `page: page error: ...` with the message, the source
location, and the stack. Errors only - ordinary console noise would drown the log it is
meant to help.

## The DevTools door

Debug builds also start the browser's own DevTools protocol on a per-process port,
announced as `devtoolsPort` in the discovery file. That drives the LIVE page semantically:
real events on real elements, including double clicks, which pixel messages cannot produce.
`tools\harness\objbrowser-live-probe.mjs` is the reference client.

The port is picked per process rather than fixed. A fixed one (9333, in the first landing)
either collides between two Excels or, worse, lets them share a browser cluster and mix
their targets on one socket.

## Using it

```powershell
# Every live instance
node tools\harness\xlide-api.mjs instances

# One instance, chosen by workbook
node tools\harness\xlide-api.mjs --workbook scratch.xlsm state
node tools\harness\xlide-api.mjs --workbook scratch.xlsm locals

# The standing regression probe: every route plus the break round trip
powershell -File tools\harness\Test-DebugApi.ps1
```

```js
import { open } from "./tools/harness/xlide-api.mjs";

const api = await open({ workbook: "scratch.xlsm" });
await api.writeModule("CleanModule", source);
await api.breakpoint("CleanModule", 8, { state: "on" });
// start the macro, then:
await api.waitFor((s) => s.debugMode === "break");
console.log(await api.locals());
```

## History

This is the second landing. The first (branch `post-v010-experiments`, 2026-08-05) was
rolled back with the rest of that branch after a crash storm during break-mode work. The
crashes were not the api: they were an undersized VARIANT in the accessibility interop,
root-caused and fixed on 2026-08-05
([locals-break-investigation.md](locals-break-investigation.md)). What changed in this
landing, beyond the fix underneath it:

- locals and watches read the ghost reader thread's snapshots instead of mirrored fields
  the session had to maintain.
- the DevTools port and the discovery file are per process, for several Excels at once.
- dead instances' discovery files are swept at session start.
- `breakpoint` gained `state=on|off`, and `caret` is new: without it nothing outside the
  page can aim a Run at a procedure.
- a client library (`xlide-api.mjs`) and a standing probe (`Test-DebugApi.ps1`, 30 checks)
  ship with it.
- the modal guard above, `dialogs`, `dismiss`, `eval`, `doctor`, `perf`, `journal`,
  `history`, `assert`, the log wait, the host heartbeat, and page-error forwarding are all
  new in this landing; none of them existed on the branch.
