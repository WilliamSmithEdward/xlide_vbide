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
| `perf` | GET | `reset=1` | recent raw placement and marshal durations, **and the analyzer's cost per method**: calls, time spent QUEUED behind another call, time spent ON the pipe, median, p95, and the individual calls over 120ms. The engine serves one request at a time, so a diagnostics pass over a large module delays every keystroke's completion behind it and a combined figure blames completions. `reset=1` forgets the engine figures first, so an experiment measures what it provokes |
| `journal` | GET | `lines` | one capture of a whole moment: state, dialogs, counters, recent log, recent page traffic |
| `history` | GET | | every request this door has served, a script that replays them, and **`routeCosts`**: each route's count, total and worst time, slowest total first. The door is the instrument every other measurement is taken through, so a route that has quietly become slow makes everything look slow |
| `assert` | POST | `that`, `value`, `timeoutMs` | states an expectation and waits for it, answering with what was actually seen |
| `messages` | GET | `last` | recent page traffic both directions, per surface |
| `problems` | GET | `module` | the analyzer's current findings |
| `locals` | GET | | the Locals panel's context and rows |
| `watches` | GET | | whether stopped, and the Watch panel's rows |
| `module` | GET | `name`, `project` | a module's text, read through the session's reader |
| `capture` | GET | `window=frame\|palette`, `selector`, `pad` | a BMP of the window, through PrintWindow. With `selector` it is cropped to that element |
| `module` | POST | `name`, `project`, body = text | writes the module through the session's writer, with the baseline and engine corrections a host rewrite carries |
| `dialogs` | GET | | native dialogs standing now, with their buttons, and how long the host thread has been quiet. Needs no host thread, so it answers while the editor is stuck |
| `ui` | GET | `waitMs` | the surface as the PAGE describes it: tabs with the labels the strip drew and the MRU order a close falls through, the tree's expansion and unfolded module, panes, page-side dialogs, what text has not arrived yet, the caret, **the main-thread stalls over 50ms** (which is what jank is, and which no host counter can see), and the model/document census a leak shows up in. `dialogs` is asked of the DOM by `aria-modal`, not read off a list, so a dialog added later is reported without anything remembering to add it. Reported from the fields the render reads, so it cannot mistake a stale render for the state |
| `act` | POST | `do=closeActive\|answerCloseConfirm\|activate\|cycleTab\|split\|expandWorkbook\|unfoldModule\|settings\|sponsors\|closeDialogs\|key\|focusEditor`, plus that action's arguments | drives the surface through the methods a click reaches. Answers `{did, detail}`; `did: false` means the page declined or the action did not land, which is an answer and not an error. **`closeActive` reports the OUTCOME, not the request**: closing a module with unsaved changes raises a Save / Don't Save / Cancel box and the tab stays until it is answered, so `did` is false and `detail` names what is standing. `answerCloseConfirm` with `answer=save\|discard\|cancel` answers it |
| `trip` | GET | `what=pagecall`, `n` | times a round trip ACROSS the boundary, wall clock from asking to observable. `bench` times the page and `perf` reports the host; this is the crossing neither measures. Read `pagecall` alongside the rest, since it is the floor every other figure contains. **A route body holds the host thread, and a message POSTED to the page is delivered by that thread's pump, so nothing whose effect arrives that way can be observed from inside a request at all** — those are measured across requests, in the client (`tripCaret()`) |
| `eval` | POST | `surface`, `waitMs`, body = script | runs script in the live page and returns its result as JSON. A PROMISE is awaited, so `(async () => ...)()` works. Prefer `ui` and `act` where they fit |
| `await` | POST | `surface`, `waitMs`, body = predicate | polls a predicate IN the page until it is true, answering `met` and how long it took. One request instead of a caller's poll loop |
| `layout` | GET | | the whole visible arrangement: dock sections and their groups, editor groups and their tabs, sizes, open documents, whether a drag is live |
| `layout` | POST | `reset=1`, `waitMs` | puts the arrangement back to the default and waits for the page, so a probe that dragged panes about can clean up in one line |
| `inspect` | GET | `selector`, `styles`, `rules=1`, `max` | elements matching a selector: box, classes, hidden, the computed values of the styles asked for, and which CSS rules claim them |
| `bench` | GET | `what=tabswitch\|layout\|type`, `n` | times a scenario in the page and answers min, median, p95, max, and the raw samples |
| `console` | GET | `last` | what the page said to itself: a ring of console lines, installed at page ready |
| `reload` | POST | `waitMs` | reloads the page and waits for it to come back, answering with the bundle stamp it is now running and whether that is behind the one on disk |
| `dismiss` | POST | `button`, `caption` | clicks a dialog button by name. Explicit, so it will press OK if asked |
| `guard` | POST | `on=true\|false`, `forget` | turns the dialog guard on or off and answers what it has cleared. While ON, a NOTICE this door did not raise is cleared too — see below. Off by default |
| `compile` | POST | `waitMs` | compiles the project and answers its errors as DATA, clearing the modal it raises. `compiled` is false when anything appeared |
| `documents` | GET | | the documents the surface holds TEXT for, with line counts, unwritten flags, and which is active. Not the same list as the tabs |
| `component` | POST | `action=add\|remove\|rename`, `kind=1\|2\|3`, `name`, `newName`, `project` | adds, renames or removes a component FROM INSIDE, so a fixture can be built without "Trust access to the VBA project object model". `name` is read back from the component; a refused name adds nothing |
| `project` | GET | `project` | what the VBA project CONTAINS: every component with its kind, line count, and whether a pane is open on it, plus the project's execution mode. The object model's own answer, read from inside |
| `outline` | GET | `module`, `project` | a module's procedures, from the analyzer: name, kind, line. For asserting on shape without parsing the text again in the caller |
| `pane` | POST | `action=open\|close`, `module`, `project`, `answer=save\|discard` | opens or closes a module's TAB. A close goes through the same gate the tab's own X uses, so unwritten edits raise the question unless `answer` settles it |
| `settings` | GET/POST | any setting name | the developer's settings, and a POST changes only what it names — the page's own update takes the whole object |
| `undoRename` | POST | | puts the last rename back: every module it touched and the component's old name. The editor's own undo cannot, because a rename spans modules and the undo stack is per model |
| `breakpoints` | GET | | every recorded breakpoint as `{module, project, lines}`, and the project mode. **`project` is part of the identity**: two open workbooks can each hold a module of the same name, and the record was keyed by the name alone until 2026-08-08, so a breakpoint set in one was reported against the other and a run that should have stopped did not |
| `type` | POST | body = text, `waitMs` | types into the editor through its own KEYBOARD pipeline, so smart Enter, comment continuation and auto-indent all run. A `
` is an Enter, and each segment is given a turn of the loop — typing has gaps, and a script without them tests nothing |
| `mark` | POST | `text` | writes a labelled line into the shim log and answers the byte offset it landed at, so `log?since=` can return exactly the slice a step produced |
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

`layout?reset=1` (POST) puts the arrangement back and waits for the page. A probe that drags
panes about is testing the right thing and leaving the wrong thing behind — the layout is
persistent state, and clearing its storage key does not undo what the page already holds in
memory.

### Seeing why the page looks like that

`inspect` answers with the elements a selector matches: their boxes, classes, whether they
are hidden, the computed value of any styles named, and — with `rules=1` — which CSS rules
claim those properties, spelled out.

```powershell
Invoke-RestMethod "$api/inspect?selector=.dock-split&styles=align-items,display&rules=1"
```

The rule list is the point. This page shares a document with a large bundled stylesheet, and
a structural class of ours once inherited `align-items: baseline` from an unrelated rule,
collapsing every pane to its tab strip's height. It read as a flex bug in our own code and
cost an hour; the loop that eventually found it — walk every stylesheet, keep the rules this
element matches — is now this route.

`console` answers with what the page said to itself. Only UNCAUGHT errors reach the shim log,
deliberately, because forwarding every line would drown it — so a handled `console.error` or
a warning is invisible without DevTools attached, which is exactly the situation during a
live test. The ring is installed at page ready (including a reload's), wraps rather than
replaces the console, and keeps the last 500 lines.

### Looking at one widget

`capture?selector=` crops the screenshot to an element. A whole frame is a large picture in
which a 54-pixel drop zone cannot be seen, and a surface built by reading numbers rather than
looking at it is built with one eye shut. `tools\harness\Get-Shot.ps1` wraps this and
converts to PNG, which is what most things that want to LOOK at the result can open:

```powershell
tools\harness\Get-Shot.ps1 -Selector '.drop-compass' -Out compass.png
```

One trap is recorded in the code. The page reports coordinates inside its own client area,
and the crop has to place that area inside the captured frame. The surface's overlay window
is NOT the document area it is a child of — the surface draws the menu bar and toolbar too,
so it is taller — and a first cut that used the parent landed tens of pixels high, returning
the toolbar when asked for a pane header.

### Numbers for what a developer feels

`bench` times a named scenario in the page and answers min, median, p95, max, and the raw
samples. The counters elsewhere say what the HOST spent; this says what the surface costs,
which is where the risk moved when the workspace learned to split and dock.

| `what` | measures |
| --- | --- |
| `tabswitch` | putting another open document's model on screen — the live-model claim, in milliseconds |
| `layout` | re-measuring every editor, which is what a splitter drag and every dock change costs |
| `type` | an edit applied to the model and the page's work to show it |

```powershell
Invoke-RestMethod "$api/bench?what=tabswitch&n=40"
```

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

Four things now stand between a modal and a wedged session.

**It can be seen.** `dialogs` enumerates windows, which needs no host thread, so it answers
while every other route would time out. It returns each dialog's caption, **the text it
says**, and its buttons. The text matters: every VBA compile error wears the caption
"Microsoft Visual Basic for Applications", so a harness reading captions alone learns that
something is wrong and never what.

`doctor` names a standing dialog as a finding, because a session with one is not healthy
however well every other route answers.

**Do not use `heartbeatAgeMs` to detect one.** A VBA modal PUMPS messages, so the host thread
keeps completing poll ticks the whole time it is blocked: measured 2026-08-07, a compile error
stood for fourteen seconds with the heartbeat never above 140ms. What is standing is the only
evidence that something is standing.

**It is cleared automatically, if the door raised it.** Every request sweeps dialogs this door
is answerable for — including the routes that need no host thread, which is where the sweep
belongs, because `dialogs`, `dismiss` and `guard` are exactly what a caller reaches for while
something is stuck. It presses one SAFE button, and what counts as safe depends on what is
being asked:

- A **notice** — every button an acknowledgement (OK, Help, Close, Continue) — is answered with
  OK. It is reporting, not asking, so pressing it decides nothing. This case used to be
  missed: the policy was Cancel/Close/No only, so a compile error offering OK and Help matched
  nothing and stood with the host thread behind it.
- A **question** is only ever declined: Cancel, then No, then Close. Never OK, Yes, Save,
  Delete, or Run — a dialog nobody read must not be agreed with.

A request that means to open a dialog passes `keep=1` and what it opens is exempt.

**A harness can ask for more: `guard`.** While the guard is on, a NOTICE the door did NOT raise
is cleared too, and recorded in `cleared` so nothing is swallowed silently. This is off by
default and is never turned on by itself, because the attribution rule below is right for a
person at a keyboard and wrong for an unattended run: a compile error raised by an experiment
stood for six minutes with everything behind it (2026-08-07), and no route could say so.

A question is still never answered, guard or no guard.

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
