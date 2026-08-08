# Driving Excel

How to get a running Excel into a state something can drive, what the debug api can do once it
is there, where the api stops, and what still has to go through COM or the harness.

Written 2026-08-07 and **kept current**: anything added to the door, the client, or the harness
belongs here the same day. A door nobody can find the handle for is a door nobody opens.

> ## 📖 The API reference is [**debug-api.md**](debug-api.md)
>
> **That document is the authority on the api itself**: every route with its arguments and what it
> answers, the reasoning behind each one, and the failures each was built after. Read it when you
> need to know what a route DOES.
>
> **This document is the operational guide**: how to get a host into a state that can be driven at
> all, which loop to use for which change, where the api stops, and what still needs COM or CDP.
> Read it when you need to know how to GET somewhere.
>
> The two are kept in step by the gate — `tools\harness\audit-routes.mjs` reads the routes out of
> the shim and fails when one is missing from either document or has no client method.

Also: [working-with-modals.md](working-with-modals.md) is the rules for opening a modal at all,
and [testing.md](testing.md) is the gate.

---

## 1. Getting to a drivable state

Four things have to be true, and each of them has cost a session.

### The build must be Debug

The whole door is inside `#if DEBUG`. A Release shim has no api, no `/eval`, no DevTools port,
and the gate has a step that proves it (`Release carries no debug api`). If `discover()` finds
nothing, check this first.

```bash
dotnet publish src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj -c Debug -r win-x64
```

with `DOTNET_ROOT` at `%LOCALAPPDATA%\Microsoft\dotnet` and
`%ProgramFiles(x86)%\Microsoft Visual Studio\Installer` on PATH.

### Excel must be started as an ordinary process

```bash
tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\RenameFixture.xlsm -Fresh
```

Not `New-Object -ComObject Excel.Application`. A host created through automation runs in
**embedding mode** and loads **no add-ins**, so the thing under test is never there. It has to be
`EXCEL.EXE <workbook>` with a document on the command line, so it initialises promptly.

### The editor must be opened

Excel loads VBE add-ins when the **VBE** starts, not when Excel does. Nothing is under test until
the editor is up — and it is opened through Excel's own ribbon command, not the object model:

```powershell
$excel.CommandBars.ExecuteMso('VisualBasic')      # not $excel.VBE.MainWindow.Visible
```

`$excel.VBE` is one of the two properties "Trust access to the VBA project object model" gates,
and with that off it is **null**, so the obvious line fails — or worse, is swallowed by a
try/catch and the script carries on believing the editor is open. `ExecuteMso` is not gated.
`tools\harness\Open-VbeEditor.ps1` is that one line with the reason attached, and every harness
script goes through it.

`Start-Excel.ps1` does this, along with the two things around it that are easy to get wrong:

- **Attaching through the window, not the running object table.** A host publishes itself in the
  ROT ten to forty seconds after it is visibly ready, and none of that wait says anything about
  the add-in. `AccessibleObjectFromWindow(EXCEL7 window, OBJID_NATIVEOM)` answers in under a
  second and names the instance by window, so there is no chance of adopting a different one.
- **Clearing Excel's resiliency keys.** A harness terminates Excel by design and Excel reads
  termination as a crash: on the next start it offers document recovery and disables what it
  blames, both of which stand in front of the thing being tested.

### The door must be found

Each session writes `%LOCALAPPDATA%\xlide_vbide\debug-api-<pid>.json` with its port, token and
DevTools port. The client reads them:

```js
import { open, discover } from "./xlide-api.mjs";
const api = await open({ workbook: "RenameFixture.xlsm" });   // or { pid }, or nothing when there is one
```

A discovery file **outlives a killed Excel**, so `discover()` proves liveness by asking each one
for `state` before returning it.

### Then ask whether it is healthy

```bash
node tools\harness\xlide-api.mjs doctor
```

`findings` empty is the only summary worth trusting. It answers the questions that cost the most
when nobody thinks to ask them: is this session running the code that was just built, is the
engine up, are the ghost readers attached, **is a dialog standing**.

---

## 2. The development loop

**The page and the shim have different loops, and using the shim's loop for a page change wastes
about a minute an iteration.**

| Changed | What is needed | Why |
| --- | --- | --- |
| `ui/editor/**` | `tools\Update-Page.ps1` | The bundle is served over loopback from a folder on disk, read per request. Copy and reload. **No restart.** |
| `src/Xlide.Vbe.Shim/**` | close Excel, `dotnet publish`, restart | A host holds an add-in library open for its lifetime. Nothing can replace a file Excel is holding. |
| `engine/**` | `npm run package`, then restart | The add-in runs `engine\dist\xlide-engine.exe`. `npm run build` writes only the bundle, which nothing at runtime reads. |

```bash
tools\Update-Page.ps1
```

builds, copies into the published shim's `ui\editor\dist`, and reloads every live editor —
reporting the build stamp each is now running. `-NoBuild` copies what the gate just built.

> Excel holds **both** the shim DLL and the engine executable. `npm run package` fails with a
> copyfile error while Excel is running, and the gate's `engine executable is current` step then
> fails on the next run.

---

## 3. What the api can do

### Every route, and how to call it

The reasoning behind each route is in [debug-api.md](debug-api.md); this is the mapping from route
to client method. **`tools\harness\audit-routes.mjs` proves this table complete** — it reads the
route cases out of the shim and fails when one is missing here, missing from the reference, or has
no client method. It runs in the gate, because a route table is exactly the kind of thing that is
complete on the day it is written and quietly is not, six routes later.

| Route | Client | |
| --- | --- | --- |
| `act` | `act(name, args)` | drives the surface through the methods a click reaches |
| `assert` | `assert(that, {value, timeoutMs})` | states a claim and waits for it |
| `await` | `until(predicate, {waitMs})` | waits for a condition IN the page |
| `bench` | `bench(what, {n})` | times a scenario: min, median, p95, max, raw samples |
| `breakpoint` | `breakpoint(module, line, {project, state})` | set, clear or toggle |
| `capture` | `capture(window)` | a BMP of the window, through PrintWindow |
| `caret` | `caret(line, {module, column, project})` | navigates first when a module is named |
| `command` | `command(name)` | any editor command by name |
| `compile` | `compile({waitMs})` | compiles; errors as DATA, modal cleared |
| `component` | `component(action, {kind, name, newName, project})` | add, rename, remove — what a fixture is made of, from inside |
| `pane` | `pane(action, {module, project, answer})` | open or close a module's tab |
| `settings` | `settings()` / `settings({...})` | read them, or change one without restating the rest |
| `undoRename` | `undoRename()` | puts the last rename back, across every module it touched |
| `breakpoints` | `breakpoints()` / `breakpointsIn(project)` | what is set, per module AND workbook, and the mode |
| `type` | `type(text)` | types through the keyboard pipeline: smart Enter, comment continuation, auto-indent |
| `mark` | `mark(text)` | a labelled line in the log, and the offset to read back from |
| `outline` | `outline(module, project)` | a module's procedures, from the analyzer |
| `project` | `project(project)` | what the VBA project CONTAINS: components, kinds, line counts, panes |
| `console` | `console(last)` | what the page said to itself |
| `dialogs` | `dialogs()` | what is standing, with its TEXT. Needs no host thread |
| `dismiss` | `dismiss(button, caption)` | presses a button by name. Will press OK if asked |
| `doctor` | `doctor()` | the start-of-session checklist |
| `documents` | `documents()` | what the surface holds TEXT for |
| `eval` | `ask(script)` / `eval(script)` | runs script in the page. **Prefer `ask`** |
| `guard` | `guard(on, {forget})` | the dialog guard, and what it has cleared |
| `history` | `history()` | every request served, as a replayable script |
| `immediate` | `immediate(text)` | schedules an Immediate evaluation |
| `inspect` | `inspect(selector, {styles, rules, max})` | boxes, classes, computed styles, winning rules |
| `journal` | `journal(lines)` | one capture of a whole moment |
| `layout` | `layout()` / `resetLayout()` | the arrangement; and putting it back |
| `locals` | `locals()` | the Locals panel |
| `log` | `log({since, match, max, waitMs})` / `waitForLog(match)` | the shim log; BLOCKS with `waitMs` |
| `messages` | `messages(last)` | page traffic, both directions |
| `module` | `readModule(name, project, {live})` / `writeModule(name, text, project)` | through the session's own reader and writer; `live` reads what the surface holds, unwritten |
| `perf` | `perf({reset})` / `engineCosts()` | placement and marshal durations, and the ANALYZER's cost per method |
| `placement` | `placement()` | forces a placement pass |
| `problems` | `problems(module)` | the analyzer's findings |
| `reload` | `reload({waitMs})` | reloads the page and waits for it |
| `state` | `state(timeout)` | shown module, mode, handles, rects, DevTools port |
| `stats` | `stats()` | uptime, memory, handles, GC, placement and marshal counters |
| `trip` | `trip("pagecall", {n})` / `tripCaret()` | times what a person waits for, ACROSS the boundary |
| `ui` | `ui()` | the surface as the page describes it: tabs, tree, panes, dialogs, caret |
| `watches` | `watches()` | the Watch panel |
| `windows` | `windows()` | every editor window |
| `native` | `native({text})` / `inSync()` / `parityAll()` | the HOST's own panes, caret and CONTENT, under the surface |

Also on the client, built from those: `waitUntilResponsive()` and `ask()`.

### Grouped by what you would drive it for

### Looking

| Want | Call |
| --- | --- |
| Is this session sane | `doctor()` |
| What is shown, and where the windows are | `state()`, `windows()` |
| What the surface holds TEXT for | `documents()` |
| **What the surface LOOKS like: tabs, tree, panes, dialogs, caret** | **`ui()`** |
| The whole arrangement: docks, groups, tabs, sizes | `layout()` |
| A module's text, through the session's own reader | `readModule(name, project)` |
| The analyzer's findings | `problems(module)` |
| Locals, Watches | `locals()`, `watches()` |
| What the page said to itself | `console()` |
| Page traffic both directions | `messages()` |
| The shim log, optionally BLOCKING until a line appears | `log({ match, waitMs })` |
| One capture of a whole moment | `journal()` |
| A picture of the window | `capture("frame")` |

### Doing

| Want | Call |
| --- | --- |
| Put the caret somewhere, opening the module first | `caret(line, { module, column, project })` |
| Run an editor command by name | `command("compile")`, `command("run")` … |
| Set, clear or toggle a breakpoint | `breakpoint(module, line, { state })` |
| Evaluate in the Immediate window | `immediate(text)` |
| Write a module through the session's writer | `writeModule(name, text, project)` |
| Does the project compile, errors as DATA | `compile()` |
| **Close a tab, click the tree, send a chord, open a dialog** | **`act(name, args)`** |
| Run script in the page | `ask(script)` — see the trap below |
| Reload the page and wait for it | `reload()` |
| Put the arrangement back | `resetLayout()` |

### The surface, asked and driven

Two routes, and they replace almost every DOM script a probe used to carry.

```js
const ui = await api.ui();
ui.workspace.groups[0].tabs;      // module, project, LABEL as drawn, active, dirty, problems
ui.workspace.groups[0].recent;    // the MRU stack a close falls back through
ui.explorer.workbooks;            // expanded, and each module's unfolded state and procedures
ui.explorer.unfolded;             // the accordion's one open module, and its workbook
ui.dialogs;                       // settings, help, sponsors, references, object browser
ui.waiting.documents;             // text asked for and not yet arrived
ui.focus;                         // model, line, column, and whether the editor has the keyboard
ui.emptyViewShown;                // a DIFFERENT question from having no tabs
ui.search;                        // open, query, scope, matches, current
ui.bookmarks;                     // the marked lines of the model on screen
ui.longTasks;                     // main-thread stalls over 50ms, worst first

await api.act("closeActive");
await api.act("answerCloseConfirm", { answer: "discard" });   // the unsaved-changes box
await api.act("search", { query: "Recalculate", scope: "project" });
await api.act("bookmark", { which: "toggle" });   // `do` is reserved for the action name
await api.act("format");                        // Format Module; {selection: true} for the selection

// IntelliSense and the squiggles, at a word or a position
await api.act("hover", { word: "Recalculate" });
await api.act("completions", { line: 7, column: 12 });
await api.act("quickFixes", { word: "Recalcualte" });
await api.act("definition", { word: "Recalculate" });   // where F12 would go, caret unmoved
await api.act("rename", { word: "Recalculate", newName: "Recompute" });  // CHANGES STATE
await api.undoRename();                                 // and puts it back
await api.at("Recalculate");                    // colour as painted, and the markers on it
// -> { word, tokenClass: "mtk4", colour: "rgb(156, 220, 254)", squiggles: [{severity, message, code}] }
//
// `word` matches case-insensitively, because VBA does: the host RECASES identifiers on write, so
// `total = 1` comes back `Total = 1` and a case-sensitive lookup misses a word that is on screen.
```

> **These call the provider objects monaco calls**, with the arguments monaco passes, so they
> answer what the developer's editor would answer — including the refusal every provider opens
> with, that it will not answer for anything but the host-active module.
>
> That is the prime heuristic at work: **an api action must leave, and must report, the state the
> same action through the UI would.** The first version of these four called the bridge request
> underneath the providers instead, skipping that refusal, and reported hover healthy through a
> whole session in which hover was dead on screen (2026-08-08). Coverage that agrees with the code
> and disagrees with the product is worse than none.
>
> When IntelliSense is reported dead, the gate is still worth checking directly — it says WHY in
> one comparison, `ui.focus.model` against `ui.focus.host.model`:
>
> ```bash
> node -e "import('./tools/harness/xlide-api.mjs').then(async m => { const u = await (await m.open({})).ui(); console.log(u.focus.model === u.focus.host?.model ? 'providers answer' : 'PROVIDERS SILENT: ' + u.focus.model + ' vs ' + u.focus.host?.model); })"
> ```

```js
await api.act("expandWorkbook", { workbook: "TwinFixture.xlsm", open: true });
await api.act("unfoldModule", { module: "Helpers" });
await api.act("key", { code: "KeyW", ctrl: true, target: "document" });
await api.act("closeDialogs");
```

`act` answers `{did, detail}`. **`did: false` is an answer, not a failure** — closing a tab when
nothing is open, expanding a workbook that is not there. Scripts that treat it as a throw stop
distinguishing it from a broken door.

> **Closing a module with unsaved changes does not close it.** The host asks the page to confirm,
> and a Save / Don't Save / Cancel box stands until something answers. `closeActive` reports
> `did: false` and names what is standing; `answerCloseConfirm` gets past it. This is worth
> knowing before writing any close loop: five closes in a row once reported success five times
> while the tab never moved (2026-08-07).

> **Why these exist.** Both replace a habit that cost more than any product defect here. Reading
> the surface with `querySelectorAll` measures the RENDER and calls it the state, and a stale
> render is the bug worth catching; `ui()` reports from the fields the render reads. Driving the
> surface with synthesised events guesses which ones the handlers listen for, and guesses wrong
> silently: the tab close box arms at `pointerdown` and fires at `pointerup`, so `.click()` on it
> does nothing whatsoever and the probe reports a working feature broken (2026-08-07).
>
> If a question needs a DOM script twice, it belongs in `ui/editor/src/devsurface.ts`.

### Parity with the native editor

```js
await api.native();    // the host's active pane, its caret, and every pane it holds open
await api.inSync();    // one boolean over native, surface and page
```

> **A feature that touches the editor is not fully tested until it validates this.** The surface
> COVERS the host's code panes; it does not replace them. Run, Step, Compile and
> ToggleBreakpoint all act on the **native active code pane and the caret inside it**, not on the
> page — so a page showing one module while the native pane holds another is a Run that executes
> where the developer is not looking and a breakpoint on the wrong line, with nothing on screen
> to say so.
>
> **Parity means the CONTENT matches, not the names.** A surface holding an empty document for a
> module the host has 42 lines of passes every name comparison there is, and shows a blank
> editor — which is exactly how the first one was found. Both texts are reduced the same way in
> the shim, so a changed character registers and the host's CRLF does not; `native({text: true})`
> carries both texts for the run that fails.
>
> Every check in this repo read the page and the workbook and never the panes below, until
> 2026-08-08. It is an invariant in the surface walk now, checked after every step, and the
> rename and debugger suites assert it after every state change.

```bash
node tools\harness\debugger-features.mjs   # the run-and-stop cycle, parity at every stage
```

Run against `DebugFixture.xlsm` (`tools\New-DebugFixture.ps1`), the only fixture that **compiles**
— the rename fixture deliberately does not and the language fixture carries a module of deliberate
defects, so pressing Run on either raises a modal and tests the dialog guard instead of the
debugger. The suite sets a breakpoint, runs, and checks that the native pane is on the stopped
module at the stopped LINE, that the Locals panel holds the procedure's variables with the values
assigned before the stop, and that Step Over advances the native caret. It leaves break mode with a
Reset whatever happens: a session left stopped blocks everything after it.

```bash
node tools\harness\step-into-features.mjs   # the HOST moving, and the page following
```

The other direction of parity, and the one a developer actually meets. Everything else drives
through the api and asks whether the native panes keep up; this steps INTO a procedure in another
module, so the **debugger** activates a module the page was not showing and never asked for. The
page has to open a tab and follow, or you step into a procedure and watch a different module's
code.

The native panes are covered by the surface, so a user cannot click one — which makes the
debugger the realistic driver of this direction, and the only one worth testing. Measured: the
native pane crosses to `Helper`, the page follows, a tab appears, all three agree, every open
module's content still matches, and the Locals panel shows the new frame's `value` and `prefix`.

```bash
node tools\harness
ename-features.mjs   # rename and definition, with parity at each step
```

Rename is the one feature here that rewrites the developer's code, and it had no api at all
until 2026-08-08. The suite covers what the fixture was built for: a qualified call follows, a
bare call that two modules could own is **left alone**, a module whose name merely begins the
same way is untouched, and `undoRename` puts every module back.

### The language features, against real receivers

```bash
tools\New-LanguageFixture.ps1              # builds LanguageFixture.xlsm
node tools\harness\language-features.mjs   # the dot menu, by receiver kind
```

The rename fixture is shaped for renaming: one method on one class, and a project that
deliberately does not compile. That makes it the wrong workbook for asking what IntelliSense
offers, because a receiver with one member proves almost nothing and a project full of errors
buries the one error a quick-fix test wants to see. `LanguageFixture.xlsm` is shaped for these
questions instead:

| module | what it is for |
| --- | --- |
| `Gadget` | a class with a Sub, a Function, two `Property Get`s, a `Let`, and a **Private** Sub that must never appear in another module's menu |
| `Shapes` | an `Enum` and a user-defined `Type`, which resolve by their own paths |
| `Uses` | one receiver per line: class, type, enum, `Application`, `ActiveSheet`, and a call with parameters for signature help |
| `Defects` | the ONLY module with findings, so a quick-fix test sees exactly one |

**Two of its cases fail on purpose.** A project `Type` receiver offers its own name instead of its
fields, and an `Enum` receiver offers nothing. Both are analyzer defects, filed as
[xlide_vscode#11](https://github.com/WilliamSmithEdward/xlide_vscode/issues/11) rather than
patched here. They are left failing rather than deleted: a suite that drops the cases it cannot
pass stops being able to say when they start passing. The script exits 0 for those two and
non-zero for anything else, and announces it if one of them starts working.

### Walking the surface at random

```bash
tools\harness\Start-Excel.ps1 -Workbook artifactsixtures\RenameFixture.xlsm,artifactsixtures\TwinFixture.xlsm
node tools\harness\surface-walk.mjs --steps 80 --seed 424242
```

Picks each action from a deterministic stream and re-checks every invariant after every step, so
a failure replays from the seed it prints. **Start Excel with two workbooks** — `-Workbook` takes
a list, and they land in one process, which is one session and one door. Half of what this checks
does not exist with a single workbook open.

Two details are load-bearing, and both were added after the walk lied about a clean run:

- **The actions are weighted.** Unweighted, there are two ways to close a tab and one to open
  one, so the workspace drains and stays empty: 54 of 70 steps with nothing open at all.
- **It reports the states it reached.** A run that never holds two modules of the same name
  passes every label check vacuously. That line is what found the `pane` route dropping its
  project argument: the walk opened both workbooks' `Helpers`, reported `collision=0`, and was
  right.

### Measuring what a person actually waits for

```js
await api.trip("pagecall");   // the floor: a script into the page and back
await api.tripCaret();        // host caret set, to the PAGE agreeing where it is
```

### Where the time actually goes

```js
await api.perf({ reset: true });        // forget, then provoke the slowness
// ... type, hover, open modules, whatever feels slow ...
console.table(await api.engineCosts()); // ranked by total time spent
(await api.ui()).longTasks;             // main-thread stalls over 50ms, worst first
(await api.history()).routeCosts;       // the door's own cost, per route
```

**`engineCosts()` is the one to reach for first.** Every language feature goes down one pipe —
completions, hover, signature help, diagnostics, navigation, rename, semantic tokens, outline —
and that pipe serves **one request at a time**. So each method's latency is two things added
together, and only one is the analyzer's doing:

- **`waitMs`** — queued behind another call. A diagnostics pass over a large module delays every
  keystroke's completion request behind it.
- **`callMs`** — the round trip once on the pipe. This one is the analyzer's.

A combined figure reports the first as slow completions and sends the hunt at the wrong file.
Nothing measured any of this before 2026-08-07: the features a developer feels most had no
instrument at all.

`longTasks` is the other half. A frame is 16ms; anything over 50ms is a stretch where the surface
answered no key and painted nothing, and **no host counter can see it** — the host thread was fine
throughout.

`bench()` times the page's own work and `perf()` reports the host's. Both have read healthy while
the surface felt slow, because the cost was in the crossing that neither measures. `trip` is wall
clock from asking to observable, so the door's own cost is inside the number, which is why
`pagecall` has to be read alongside the rest. Without it, a 40ms feature and a 40ms door are the
same reading.

> **Why `tripCaret` is a client method and not a route.** A route body runs ON THE HOST THREAD,
> and a message posted to the page is delivered by that same thread's pump. So a route that posts
> and then waits to see the effect waits forever: the post cannot be delivered until the body
> returns. Written as a route, the caret trip sat four seconds a sample and reported that the
> caret never moved, while the identical sequence from outside landed on the first poll
> (2026-08-07). `Thread.Sleep` does not help; it yields the CPU and pumps nothing.
>
> `RunPageScript` survives this because `ExecuteScript`'s answer returns by a path the blocked
> thread still completes. `PostWebMessageAsString` does not. **Anything that observes a POSTED
> effect has to be measured across requests, on the client side of the door.** That applies to
> any future route, not just this one.

### Waiting, rather than sleeping

```js
await api.until("window.xlideBridge.documents.all().length > 1");
await api.waitForLog("rename:");
await api.assert("surfaceReady");
await api.waitUntilResponsive();
```

**Every fixed sleep in a probe is a race that has not lost yet.** A `settle(2500)` was right until
a host round trip joined the path it was waiting on, and then it reported a working feature
broken — twice in one afternoon (2026-08-07).

### Not being stuck behind a modal

```js
await api.guard(true);           // for the length of an unattended run
…
const g = await api.guard();     // g.cleared says what it took off the screen
await api.guard(false, { forget: true });
```

- `dialogs()` needs **no host thread**, so it answers while everything else would time out. It
  returns each dialog's caption, **the text it says**, and its buttons.
- A **notice** (every button an acknowledgement) is safe to clear; a **question** is only ever
  declined, guard or no guard.
- With the guard **off**, a dialog the door did not raise is left alone however long it stands —
  right for a person at a keyboard, wrong for an unattended run.

> **`heartbeatAgeMs` cannot detect a modal.** A VBA modal PUMPS messages, so the host thread keeps
> completing poll ticks the whole time it is blocked: measured at under 140ms through fourteen
> seconds of being stuck. What is standing is the only evidence that something is standing.

---

## 3a. Language, and what the host will not store

Two tests, because there are two different exposures and only one of them is ours.

```bash
node engine	est\language.mjs          # headless, a gate step, 18 scripts
tools\harness\Test-Language.ps1        # live, needs a running host
```

**The engine matrix is about OFFSETS.** The companion product decodes the workbook's VBA streams
itself, so its language risk is code pages; this product never touches those bytes, and its risk
is arithmetic. Every language feature names a position as a number of units into the source, and
a byte count or a code-point count anywhere in the chain drifts by the width of the non-ASCII
text to its left. That is invisible in an English module, gets worse further down the file, and
corrupts an edit rather than failing one. The matrix drives 18 scripts plus astral emoji through
open, diagnose, outline, definition (checking line **and column**) and rename (comparing the
whole produced text). All pass.

**The live probe is about what survives the round trip**, and the headline is not ours:

> **VBA stores module text in the system ANSI code page, not in Unicode.** A character outside
> that page does not survive being written at all — it comes back as a question mark, from Excel,
> before xlide sees it. On a Western European system, measured 2026-08-07: accented Latin
> survives, and Cyrillic, Greek, Hebrew, Arabic, Thai, CJK and emoji do not.
>
> This is why the companion product patches `PROJECTCODEPAGE` in its own tests — the code page is
> a property of the VBA project. Nothing in xlide_vbide can widen it.

The probe reports which scripts survive on the machine it runs on rather than asserting a list,
and fails only when text reaches COM and is then lost by **our** code.

---

## 4. Where the api stops

These are not oversights; they are the boundary.

| The api cannot | Because | Instead |
| --- | --- | --- |
| Start or stop Excel | It lives inside the process it would have to start | `Start-Excel.ps1`; no trust setting needed |
| Open, close or save a workbook | The door drives the EDITOR, not the host application | Excel's own object model; not gated |
| Answer a QUESTION dialog automatically | Deliberate. It declines or acknowledges; it never agrees | `dismiss(button)` by name, when you know |
| Send real mouse or keyboard input to the page | Synthesised DOM events do not reach the editor's own mouse handling | CDP, section 6 |
| Reach a second Excel process | One door per add-in session | `discover()` lists them, each with its own port |
| Run in a Release build | The whole door is `#if DEBUG` | — |

---

## 4a. "Trust access to the VBA project object model"

**Neither the add-in nor the normal harness needs it. Drive with it OFF.**

That setting — Trust Center → Macro Settings → *Trust access to the VBA project object model*,
`HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM` — is a barrier on the way IN from
outside. It gates `Workbook.VBProject` and `Application.VBE` when they are reached from
**automation**: another process, or VBA code, asking Excel for its VBA project. Off, those raise
1004, *"Programmatic access to Visual Basic Project is not trusted."*

| | Needs it | Why |
| --- | --- | --- |
| The add-in | **no** | It is a VBE add-in. The host hands it the `VBE` object at `OnConnection`, so it is already inside; it never asks Excel for a project. Every project access in `src/` is `VBE.ActiveVBProject`, never `Workbook.VBProject` |
| The debug api | **no** | The door runs inside the add-in's own process and calls the same objects |
| `Start-Excel.ps1` | **no** | Opens the editor through `CommandBars.ExecuteMso('VisualBasic')` — Excel running its own ribbon button — not through `$excel.VBE` |
| Building a fixture | **no** | `api.component()` adds, renames and removes from inside |
| Reading a module back | **no** | `api.readModule()` reads the real object model, from inside |
| Running a procedure | **no** | `Application.Run` is not gated (measured) |
| `New-RenameFixture.ps1` | **yes** | Predates the `component` route and still uses `$book.VBProject` |
| Verifying WITHOUT the add-in | **yes** | See the warning in section 5 |

This is worth caring about twice over: needing a security setting changed before a product works
is a real adoption barrier, and xlide does not have one — and a harness that leaves the setting off
is a harness that cannot be blamed for what a macro does while it runs.

**Verified 2026-08-07, with the box unticked and `AccessVBOM = 0`:**

| | |
| --- | --- |
| `Workbook.VBProject` from automation | **null** — refused |
| `Application.VBE` from automation | **null** — refused |
| xlide loading, surface up | **yes** (seen on screen) |
| `doctor` | healthy, no findings |
| `state` | named the shown module and its workbook |
| `readModule("Helpers")` | 442 characters of real code |
| `problems`, `compile` | answered; `compile` drove the VBE and read back its error |

The add-in was reading and driving the VBA project throughout, from inside, while nothing outside
could reach it at all.

> **The refusal is a NULL, not an exception** — at least through PowerShell's COM binder.
> `$excel.VBE` simply evaluates to nothing, so a probe written as `try { … } catch { }` reports
> success and prints an empty string. Test it as `$null -eq $excel.VBE`, or the experiment says
> the gate is open when it is shut.

To reproduce: untick the box, close Excel completely, open a macro workbook and press Alt+F11
(the harness cannot open the editor for you — `$excel.VBE` is exactly what is refused; Excel's own
ribbon button, `CommandBars.ExecuteMso('VisualBasic')`, is not). Tick it back afterwards to get the
harness working again.

---

## 5. COM, and how little of it is left

**Everything a harness normally needs works with the trust setting OFF.** Only the ungated part of
Excel's object model is used, and the rest goes through the door.

| Job | How | Trust setting |
| --- | --- | --- |
| Start Excel | `Start-Process EXCEL.EXE <workbook>` | not needed |
| Attach to Excel | `AccessibleObjectFromWindow(EXCEL7, OBJID_NATIVEOM)` | not needed |
| Open the editor | `$excel.CommandBars.ExecuteMso('VisualBasic')` | not needed |
| Open, close, save a workbook | `$excel.Workbooks…` | not needed |
| Run a procedure | `$excel.Run('Module.Proc')`, or `api.immediate()` | not needed |
| **Build a fixture** | `api.component("add"/"rename"/"remove")` + `api.writeModule()` | **not needed** |
| **Read a module back** | `api.readModule(name)` — the real object model, read from inside | **not needed** |
| Compile and read the errors | `api.compile()` | not needed |

Building a fixture through the door:

```js
await api.component("add", { kind: 1, name: "Helpers" });   // 1 standard, 2 class, 3 form
await api.writeModule("Helpers", source);
…
await api.component("rename", { name: "Helpers", newName: "Aides" });
await api.component("remove", { name: "Aides" });
```

`name` comes back as the component actually ended up named, not as it was asked for. A name the
editor refuses outright — `Circle` belongs to the Excel object library — is reported as a refusal
**and nothing is added**, so a failed name never leaves a stray `Module1` behind.

> A `Run` against a project that does not compile fails for a reason that has nothing to do with
> what you asked. Check `compile()` first, or your experiment is measuring the wrong thing — which
> is exactly how a shadowing question got answered wrongly on 2026-08-07 before being redone in a
> workbook of its own.

### ⚠️ The part that still needs the trust setting

> ## ⚠️ Requires "Trust access to the VBA project object model"
>
> **Everything below needs that box ticked in Trust Center → Macro Settings. Nothing above does.**
>
> **Consider carefully before turning it on.** It is off by default for a reason: with it on, ANY
> macro or automation running as you can rewrite the VBA of any open workbook. That is the
> mechanism of the self-replicating macro virus, and it is why the setting exists at all. Turn it
> on for a session that needs it, and turn it back off.

**Reaching the project from outside the add-in.** `Workbook.VBProject` and `Application.VBE`, from
a script rather than from inside:

```powershell
$project = $book.VBProject                  # null without the setting
$code = $project.VBComponents.Item('Helpers').CodeModule
$text = $code.Lines(1, $code.CountOfLines)
```

There is exactly one reason left to want this, and it is a good one: **verifying without trusting
the thing under test.** `api.readModule` reads the real object model, but it reads it through the
add-in. When the question is "did the add-in actually do what it said", an answer that comes back
through the add-in is weaker evidence than one that does not — and that distinction found two real
defects on 2026-08-06.

`tools\New-RenameFixture.ps1` no longer takes this route: it builds its nine modules through the
door and needs no trust setting at all. Its three phases are worth copying for any fixture —
Excel makes an empty `.xlsm` (automation is fine for that, no add-in needed), `Start-Excel.ps1`
opens it properly so the add-in loads, and `tools\harness\build-fixture.mjs` writes the components
through the api.

> **The refusal is a NULL, not an exception**, at least through PowerShell's COM binder. `$excel.VBE`
> simply evaluates to nothing, so `try { … } catch { }` reports success and prints an empty string.
> Test it as `$null -eq $excel.VBE`, or your experiment will say the gate is open when it is shut.

---

## 6. What needs CDP

`state()` reports `devtoolsPort`. `http://127.0.0.1:<port>/json/list` gives the page target.

**Real input.** The editor's mouse handling does not see synthesised DOM events, so a context
menu cannot be opened with `dispatchEvent`. `Input.dispatchMouseEvent` is a trusted event and
does:

```js
await send("Input.dispatchMouseEvent", { type: "mousePressed",  x, y, button: "right", buttons: 2, clickCount: 1 });
await send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "right", buttons: 2, clickCount: 1 });
```

**Reading what that opened.** The editor's menus render inside a **shadow root**, so a
document-only query finds nothing and reads as an empty menu. Walk the roots:

```js
const roots = [document];
document.querySelectorAll("*").forEach((n) => { if (n.shadowRoot) { roots.push(n.shadowRoot); } });
```

---

## 7. Traps, each paid for once

**`/eval` answers are encoded twice.** The browser returns a result as JSON, so a script that
builds its answer with `JSON.stringify` comes back quoted twice — and one parse leaves a string
that reads as an object right up until every property of it is `undefined`. That reported `false`
for working code twice in one day. Use `api.ask()`, or the reply's `value`.

**Documents are not tabs.** The page holds a module's text once it has been **activated**, not
because its pane is open: a workspace opened onto eight modules holds one. `documents()` shows
the difference. An empty peek window and a blank editor pane were both this.

**Display name is not project id.** The page names a workbook the way it is shown
(`RenameFixture.xlsm`); the object model wants the project's identity. Routes taking a project
from the page convert with `ProjectIdFromDisplay`. Forgetting it makes `FindComponent` quietly
find nothing.

**The rename fixture deliberately does not compile.** `Helpers` and `Rival` each declare a public
`Recalculate` and `Consumer` calls it bare — that ambiguity is what the fixture is FOR. Never
compile it as part of a wider experiment, and never take a `Run` failure against it as evidence
of anything.

**A write can fail and the reply will not say so.** `writeModule` answers the same way whether the
module took the text or not; the refusal is in the LOG. Building a fixture on unchecked writes
produced one with an empty `Rival`, so the duplicate `Recalculate` was not there, so it no longer
exercised the collision it exists for — and it looked fine. Read the line count back from
`project()` after writing anything you are going to make claims about. (The cause of that
particular failure is fixed: an empty module's baseline splits into one empty line, so the write
asked to delete a line from a module that had none, and the editor refused the whole thing.)

**PowerShell 5.1's `-Encoding utf8` writes a BOM.** `JSON.parse` refuses one, and names a
character that does not appear to be in the file. Use
`[System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding $false))`.

**A stale engine executable answers "Unknown method".** The add-in runs
`engine\dist\xlide-engine.exe`; `npm run build` writes only the bundle. Three features were
reported verified and were not. The gate now fails when any engine source is newer than the
executable.

---

## 8. A session, end to end

```bash
tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\RenameFixture.xlsm -Fresh
```

```js
import { open } from "./xlide-api.mjs";

const api = await open({ workbook: "RenameFixture.xlsm" });

const health = await api.doctor();
if (!health.healthy) { throw new Error(health.findings.join("; ")); }

await api.guard(true, { forget: true });        // unattended: clear notices, report them

await api.caret(7, { module: "Watcher", column: 13, project: "renamefixture.xlsm" });
await api.until("window.xlideBridge.workspace.activeEditor().getPosition().lineNumber === 7");

await api.ask(`(function () {
  window.xlideBridge.workspace.activeEditor().trigger('probe', 'editor.action.peekDefinition', null);
  return 'fired';
})()`);

await api.until("!!document.querySelector('.monaco-editor .reference-zone-widget')");

const drew = await api.ask(`JSON.stringify({
  title: (document.querySelector('.monaco-editor .peekview-title .filename') || {}).textContent,
  lines: document.querySelectorAll('.monaco-editor .reference-zone-widget .preview .view-line').length
})`);

console.log(drew, "shown module still", (await api.state()).shownModule);

const guard = await api.guard(false, { forget: true });
if (guard.cleared.length) { console.log("the guard swallowed:", guard.cleared); }
```

Read the result **out of the VBA project through COM** when the claim is about the project, not
out of the editor's report of what it did.
