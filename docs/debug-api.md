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
| `log` | GET | `since`, `match`, `max` | a slice of the shim log and the next byte offset |
| `messages` | GET | `last` | recent page traffic both directions, per surface |
| `problems` | GET | `module` | the analyzer's current findings |
| `locals` | GET | | the Locals panel's context and rows |
| `watches` | GET | | whether stopped, and the Watch panel's rows |
| `module` | GET | `name`, `project` | a module's text, read through the session's reader |
| `capture` | GET | `window=frame\|palette` | a BMP of the window, through PrintWindow |
| `module` | POST | `name`, `project`, body = text | writes the module through the session's writer, with the baseline and engine corrections a host rewrite carries |
| `command` | POST | `name` | runs an editor command by name (`VbeCommands.ForName`) |
| `breakpoint` | POST | `module`, `line`, `project`, `state=on\|off` | goes to the line and sets, clears, or toggles a breakpoint |
| `immediate` | POST | `text` | schedules an Immediate-window evaluation, fire and forget |
| `placement` | POST | | forces a placement pass |

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
3. Start the macro (from outside: `Application.OnTime` with a workbook-qualified name; not
   `Application.Run`, which blocks while the code is stopped).
4. Poll `state` until `debugMode` reads `break`.
5. `GET locals` and `GET watches`.
6. `POST command?name=reset`.

Locals and watches come from the ghost reader thread's published snapshots, so these routes
never touch the accessibility layer themselves and cannot disturb a break. That matters: an
out-of-process accessibility client that dumps a ghost palette during a break can reset the
project (lesson 33), which is exactly why these routes exist instead.

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
- `breakpoint` gained `state=on|off`.
- a client library (`xlide-api.mjs`) and a standing probe (`Test-DebugApi.ps1`, 18 checks)
  ship with it.
