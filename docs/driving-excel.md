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
> The two are kept in step by the gate. `tools\harness\audit-routes.mjs` reads the routes out of
> the shim and fails when one is missing from either document, has no client method, or is driven
> by no probe at all.

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

**Always with `-Fresh`.** Excel reuses one process for several workbooks, so launching while an
older instance is still shutting down attaches to that one and the script dies on "Could not reach
Excel NNNN through its window". It is not a real failure and it costs a minute every time. The
switch closes what is running first, and does it more reliably than killing the process by hand,
because the window can outlive the kill by a second or two.

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

### Day to day, drive it directly. Do not write a script

**The suites are for regressions, not for questions** (the developer, 2026-08-07). A full pass over
com-leak, format-positions, three-copies and immediate-watch is about two and a half minutes plus
an Excel restart, and paying that to find out what the caret is doing is what makes a loop feel
slow. A new `.mjs` per question is worse: it leaves the repo full of files nobody runs twice.

The client is a CLI, with a verb per route. No file needed:

```bash
node tools\harness\xlide-api.mjs doctor
node tools\harness\xlide-api.mjs ui
node tools\harness\xlide-api.mjs problems
node tools\harness\xlide-api.mjs module Helpers
node tools\harness\xlide-api.mjs immediate "?1+1"
node tools\harness\xlide-api.mjs act format
```

For anything the CLI does not cover, a `node -e` against the module answers in one line.

**Write a suite when the check is a regression worth keeping**: something that was broken, is now
fixed, and must fail if it breaks again. Those earn a place in the gate. **Run the full suite**
before a commit touching the editor surface, COM or the engine, and whenever a crash or a leak is
suspected. A manual check you keep repeating is the signal it has earned a suite, which is how all
four of the current ones arrived.

### Every route, and how to call it

The reasoning behind each route is in [debug-api.md](debug-api.md); this is the mapping from route
to client method. **`tools\harness\audit-routes.mjs` proves this table complete**: it reads the
route cases out of the shim and fails when one is missing here, missing from the reference, or has
no client method. It runs in the gate, because a route table is exactly the kind of thing that is
complete on the day it is written and quietly is not, six routes later.

It also asks the question the first three do not: **does anything DRIVE this route?** Documented
and reachable is not covered, and two routes were neither driven nor known to be undriven until
somebody counted (2026-08-09). A route left out has to be named in the script's
`NOT_DRIVEN_ON_PURPOSE` with a reason, and a name there that turns out to be driven fails just as
loudly, so the list cannot become a pile of things somebody meant to get round to. One entry
stands: `drainfinalizers`, which is a bisecting tool rather than an assertion.

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
| `sync` | `syncPlan(direction, {folder, mode, project})`, `syncApply(direction, {folder, mode, ids, select})`, `syncSettings({folder, exportMode, importMode})` | import and export. `syncPlan` answers what would happen without doing any of it; `syncApply` does it and answers what it did. Modes: export `exportAll\|trueUp`, import `updateOnly\|trueUpStandardClass` |
| `component` | `component(action, {kind, name, newName, project})` | add, rename, remove: what a fixture is made of, from inside. `kind` takes 1/`module`/`standard`, 2/`class`, 3/`form` |
| `pane` | `pane(action, {module, project, answer})` | open or close a module's tab; an open that finds no such module throws rather than answering ok |
| `projects` | `projects()` / `projectHolding(module)` | EVERY open workbook, which `project()` cannot answer: it answers about one |
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
| `immediate` | `immediate(text)` / `immediate()` | evaluates a line and answers WHAT IT CAME TO; without text, reads the window |
| `inspect` | `inspect(selector, {styles, rules, max})` | boxes, classes, computed styles, winning rules |
| `journal` | `journal(lines)` | one capture of a whole moment |
| `drainfinalizers` | `drainFinalizers()` | forces a collection and WAITS for the finalizers, so a leaked COM wrapper fails at the operation that leaked it rather than minutes later. A bisecting tool, not a health check |
| `layout` | `layout()` / `resetLayout()` | the arrangement; and putting it back |
| `locals` | `locals()` | the Locals panel |
| `log` | `log({since, match, max, waitMs})` / `waitForLog(match)` | the shim log; BLOCKS with `waitMs` |
| `messages` | `messages(last)` | page traffic, both directions |
| `module` | `readModule(name, project, {live})` / `writeModule(name, text, project)` | through the session's own reader and writer; `live` reads what the surface holds, unwritten. A write the editor refuses throws rather than answering ok |
| `perf` | `perf({reset})` / `engineCosts()` | placement and marshal durations, and the ANALYZER's cost per method |
| `placement` | `placement()` | forces a placement pass |
| `problems` | `problems(module)` | the analyzer's findings |
| `reload` | `reload({waitMs})` | reloads the page and waits for it |
| `state` | `state(timeout)` | shown module, mode, handles, rects, DevTools port |
| `stats` | `stats()` | uptime, memory, handles, GC, placement and marshal counters, and the COM WRAPPER counts |
| `trip` | `trip("pagecall", {n})` / `tripCaret()` | times what a person waits for, ACROSS the boundary |
| `ui` | `ui()` | the surface as the page describes it: tabs, tree, panes, dialogs, caret |
| `watches` | `watches()` | the Watch panel |
| `windows` | `windows()` | every editor window |
| `native` | `native({text})` / `inSync()` / `parityAll()` | the HOST's own panes, caret and CONTENT, under the surface |
| `engine` | `engineSource(module, {text})` | what the ENGINE is holding for a module, against what the surface holds |

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
| Evaluate in the Immediate window, and read the answer | `immediate(text)` |
| Read the Immediate window as it stands | `immediate()` |
| The Immediate window and Watch panel, end to end | `node tools\harness\immediate-watch.mjs` |
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
await api.act("dock", { pane: "properties", side: "bottom" });  // panes, through the method a
                                                //   real drop calls; resetLayout() puts it back
await api.act("backspace", { times: 1 });       // the one key `type` cannot send; takes back a
                                                // whole indent level in leading whitespace

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
await api.native();          // the host's active pane, its caret, and every pane it holds open
await api.inSync();          // one boolean over native, surface and page
await api.engineSource(m);   // and what the ANALYZER is holding, which is a third copy again
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

**There is a THIRD copy, and it is the one findings are measured in.** The engine keeps its own
live text per module, fed by the page's edits, and prefers it over the copy the project was
seeded with whenever a request does not carry a source. Nothing could see that copy until
`engineSource()`, and while it could not be seen, a squiggle drawn in the wrong place had no way
of being attributed to the side that had drifted. Ask it whenever a finding lands somewhere the
text does not agree with.

```bash
node tools\harness\debugger-features.mjs   # the run-and-stop cycle, parity at every stage
node tools\harness\format-positions.mjs    # where a squiggle lands after Format Module
node tools\harness\three-copies.mjs        # all three, after every operation that touches a module
node tools\harness\analysis-freshness.mjs  # findings stay true while the work behind them is skipped
```

### Work that is skipped, and how to see that it was skipped

A pass no longer re-analyses what has not moved. Two layers decide: the shim leaves a PROJECT
alone when no module's text has changed since it was seeded, and the engine answers a MODULE from
its last analysis when both the module's text and the project facts it depends on are unchanged.
Editing one line of a small module in a 17,000-line project went from 476ms of analyzer time to
34ms, and a completion asked during that pass went from 332ms to 64ms (2026-08-08).

Both layers are visible from here:

```js
await api.perf({ reset: true });          // forget, so what follows is what this provoked
await api.writeModule("Small", text);     // provoke a pass
await api.engineCosts();                  // what the analyzer was actually asked to do
```

`engineCosts()` returning nothing at all after a write-back is the project skip: no seed, no
diagnostics. A diagnostics row whose `calls` matches the module count but whose `totalMs` is a
few dozen is the engine's memo: every module was asked, most answered from the last pass.

The log says which projects were left alone:

```bash
node tools\harness\xlide-api.mjs log "finding(s) stand"
```

The trap to know about: a module's findings depend on more than its own text. Change a
procedure's SIGNATURE in one module and every call to it elsewhere is right or wrong for a new
reason, with none of those modules having been touched. That is what the engine's facts
comparison is for, and `analysis-freshness.mjs` is the check that it works. With the comparison
removed the suite reports zero findings where it should report one, which is the only way to know
a check of this kind is worth having.

### The COM wrappers, which killed Excel four times before anyone counted them

```js
const s = await api.stats();
s.comWrappersLive;                                 // returns to its resting level, always
s.comWrappersGivenBack - s.comWrappersDisposed;    // must be ZERO, always
await api.drainFinalizers();                       // makes a leak fail NOW, at the operation
```

**The second line is the one that matters, and it is newer than the first.** Given back and
actually released are counted separately because for the whole life of this product they were not
the same: the release was `(wrapper as IDisposable)?.Dispose()`, the wrapper is a `ComObject`, and
`ComObject` does not implement `IDisposable`. The cast failed silently, the counter incremented
anyway, and every wrapper ever taken went to the finalizer thread. `comWrappersLive` read a
perfectly balanced 13 throughout. See [com-wrapper-release.md](com-wrapper-release.md).

`drainFinalizers()` forces a collection and waits, so a wrapper that was leaked fails at the
operation that leaked it rather than minutes later in a stack that names nothing. Run an
operation, call it, and if it does not answer, that operation is the one.

Every automation object this product touches is wrapped, and **the wrapper takes its own
reference** on top of the one the shim holds. A wrapper that is never disposed is given back by
the FINALIZER THREAD instead, and the editor's objects are apartment-threaded: releasing one from
that thread is an access violation the runtime cannot throw, so it FailFasts and ends Excel.

It also surfaces late and in the wrong place. On 2026-08-07 one missing `Dispose` was reported
four times as three different faults, blamed on `ntdll` twice, on `VBE7.DLL` once, and on this
library once, hours apart, and nothing connected them until a stack finally named
`ComObject.Finalize`. See [lessons.md](lessons.md) entry 36.

```bash
node tools\harness\com-leak.mjs        # every route, many rounds, wrappers AND handles
node tools\harness\com-leak.mjs 40     # more rounds, for a slower leak
```

`ComRuntime.TakeWrapper` and `GiveBackWrapper` are the only two doors and each does its own
counting, so a caller cannot dispose without counting or count without disposing. On a build with
the defect restored deliberately a single `project()` call leaks **441** wrappers and every row
fails; on a good build every row is flat.

35 rows: every read route, the `assert` predicates (which reach the debugger's own objects), and
the STATE-CHANGING routes, each paired with its own undo so the fixture comes back and the suite
can be run twice. The state-changers were excluded at first and should not have been: they do the
most COM work of anything here, so a guarantee that skips them is not a guarantee.

**That two-way check is the only reason the instrument is worth having**: the first two attempts
at measuring this both passed on the broken build, and both were deleted. One of them was a route
that forced a collection and drained the finalizers, which reported completely clean with 8,734
leaked wrappers pending.

Handles ride along, reported per row and judged once over the whole sweep. Per row they are
unjudgeable: Excel opens and closes handles constantly and this product is a guest in its process,
so a row's delta is mostly Excel. Taking the floor of four samples cut the swing from plus or
minus 19 to mostly zero, and what was left still tripped a per-round threshold on rows that cannot
leak, picking DIFFERENT rows on consecutive runs. Across 250 operations a leak of even one handle
each is hundreds while the churn is tens, and that is a judgement worth making.

### When the host dies, the harness says why

```js
await whyDidItDie();   // and every failed call appends this by itself
```

A probe whose host has gone reports `fetch failed` and `ECONNREFUSED`, which is the shape of the
failure and not its cause. Windows knows the cause. On 2026-08-07 three crashes were read as
unrelated instabilities across an afternoon because nobody thought to look, and one line from the
fourth explained all four. Now `open()` and every route call ask, and print the faulting module,
the exception code, and **the managed stack when there is one**, which is the frame naming this
product's own code rather than whichever library noticed the damage.

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

### Does it still work in a big module?

```bash
tools\New-PerfFixture.ps1                # the same code at 109, 1127, 4502 and 11252 lines
node tools\harness\perf-scaling.mjs
```

A single timing answers the wrong question. What matters is not "is hover fast" but "is hover
still fast in the module I actually work in", and that needs a curve.

Measured 2026-08-08, across **103x the lines**:

| | 109 lines | 11,252 lines |
| --- | --- | --- |
| hover | 0.8ms | 15.3ms |
| completions | 0.6ms | 1.4ms |
| definition | 0.8ms | 15.0ms |
| *of which the analyzer* | *0ms* | *13ms* |

The feature cost IS the analyzer's cost, and it grows sub-linearly. Nothing anywhere near
anything a developer would feel.

> **Two measurements were the instrument, not the product, and finding that out was the whole
> exercise.**
>
> The curve first sat flat at 77ms from 109 lines to 11,252 — a number that will not move is
> usually a number about the harness. The door collects a promise by POLLING and slept a flat
> 40ms before the first poll, so every async route cost the call, the sleep and the poll
> whatever the feature had done underneath. The poll backs off from 2ms now.
>
> Then the page-side figures sat at ~31ms at every size, which is suspiciously two Windows timer
> ticks. It was: work marshalled to the host thread rode `SetTimer(..., 0, 0)`, and a
> zero-elapse timer does not fire immediately — Windows clamps it to the system timer
> resolution, 15.6ms. The shim's own marshal counter read a median of 16ms with samples at 31
> and 47, one two and three ticks. **Every hop to the host thread waited most of a tick, on every
> keystroke, completion, hover and api call.**
>
> A posted message now goes out beside the timer. The timer stays as the guarantee — posted
> app-range messages had never been seen to arrive at this window, which is why it was written
> that way — but they do arrive, and the queue drains at once when they do. Measured after:
> `pagecall` 15.5ms → **0.37ms**, the promise floor 47ms → **15ms**, completions on the largest
> module 31ms → **1.4ms**.
>
> A hover in the developer's editor never crosses the door at all. What they wait for is the last
> column, not the first three.

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

### What the editor refuses, and what a refused write used to cost

```bash
node tools\harness\write-rollback.mjs
```

**Restart Excel afterwards.** That is not tidiness, it is the price of the only refusal that can
be provoked on demand: a module pushed past the editor's identifier budget leaves the whole VBE
answering "Insufficient memory to continue the execution of the program" to every attempt to add a
component, for the rest of the session. Removing the offending module does not give the memory
back. This is why the probe is not in the gate.

The line ceiling is the limit everyone knows and not the one that bites. Measured 2026-08-09:

| written | what happened |
| --- | --- |
| 65,000 lines of procedures | accepted, and takes the editor some 17s to parse |
| 65,000 module-level constants | **refused at 32,446**, the identifier budget, nowhere near 65,534 lines |
| one line of 200,000 characters | accepted, and silently broken into 197 lines |
| a 60,000-character constant | accepted, broken into 60 |
| a null character | accepted as text |

**The line break is the one that costs code.** The editor holds 1,022 characters in a line.
1,023 and above are taken and split, at 1,023, with no continuation character: a 2,018
character `Debug.Print "aaa..."` became a 1,023 character fragment with an unterminated
string followed by a 995 character one. The module then held something that could not
compile, while the surface still showed the line the developer wrote.

So a line the editor would break is not written at all now, and the complaint names the line
and its length. Refused rather than repaired: a continuation would have to be inserted inside
the developer's expression, and inside a string literal that means splitting the literal and
concatenating it, which is rewriting their code to make it fit. `ModuleText.LongestLine`
carries the number and `Xlide.Vbe.Core.Tests` pins both sides of the boundary.

So a write can be refused at any size, for a reason no line count predicts, and the refusal
arrives partway: **the delete has landed and the add has not.** A module of 2,002 working lines
asked to take a body the editor would not have came back holding 31,956 lines of it, neither
body, and the route replied ok.

The writer keeps the previous text in hand and puts it back now, and returns the editor's own
words instead of only logging them. The copy costs a read of text the write reads back anyway:
3ms of a 1,037ms write at 1,002 lines, 66ms of a 12,594ms write at 40,002. Half a percent, at the
size where losing it hurts most. The line-diff path pays nothing at all, because the lines it
removes are already in hand.

`writeModule` throws on a refusal rather than answering ok. A large write still outlives the
door's three-second budget for the host thread, so a caller writing tens of thousands of lines
reads the line count back rather than trusting the reply either way. `build-fixture.mjs` does.

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

## 3b. Import and export, and why the api and the button cannot disagree

A project's modules go out to a folder as `.bas` and `.cls` files and come back from one. The same
files the companion editor writes and reads, so a repository can be worked on from either end.

The rule this surface is built to is the one that matters most about it: **an api action must leave
the same state the equivalent UI action would.** It is not upheld by care here, it is upheld by
construction. The dialog and the `sync` route reach the same call in the host, `HandleSync`, and
neither of them knows how to write a file by itself. There is no second implementation to drift.

```js
import { open } from "./tools/harness/xlide-api.mjs";
const api = await open();

// What an export WOULD do. Nothing is written.
const plan = await api.syncPlan("export", { folder: "C:\\src\\modules" });
for (const item of plan.items) {
  console.log(item.status, item.file, item.detail);
}
// will-create  Helper.bas    Create
// unchanged    Runner.bas    Already the same

// Do it. Without `ids`, `select` decides: "checked" takes the rows the plan itself ticked.
const wrote = await api.syncApply("export", { folder: "C:\\src\\modules" });
console.log(wrote.summary);   // 5 changed, 0 skipped, 0 removed, 0 failed

// And back the other way, after editing a file on disk.
const back = await api.syncPlan("import", { folder: "C:\\src\\modules" });
await api.syncApply("import", {
  folder: "C:\\src\\modules",
  ids: back.items.filter((i) => i.status === "will-update").map((i) => i.id),
});

// The folder is remembered per project, so later calls can leave it out entirely.
await api.syncPlan("import");
```

Each row carries a side-by-side comparison in `diff`, and the same one with the VBA attribute
headers left in as `diffWithHeaders`.

**Both are condensed, and capped.** A run of identical lines longer than three either side of a
change becomes one line of kind `gap`, whose `left` says how many were left out; and a comparison
longer than 400 lines stops with a gap saying how many were not shown. Nothing about the DECISION
passes through that - statuses are settled over the whole comparison - it is only what is drawn.

It matters more than it sounds. A first export compares every module against a file that is not
there, so every line is a change and there is nothing to condense: for a project of 81,795 lines
the plan was 15.1MB of which 100% was comparison lines, for a dialog that shows one row at a time.
Capped, the same plan is 0.31MB. The dialog draws the first and offers the second behind a
tick box, because nobody edits a header but it is exactly what decides whether an exported file
comes back as the same kind of module.

**A row with nothing to show does not draw itself at all.** Both planners stopped comparing a
row whose two sides are known to agree, a pipe apart and for the same reason: the status comes
from a string equality, never from the comparison, so the screenful of agreement was only ever
going to condense back to one line. The built-in planner skips the comparison; the engine stops
sending it and the shim writes the surviving line from the row's own text, using the same
method, so an unchanged row is the same object whichever planner answered.

Measured on the same project of 81,795 lines, and the second plan is the one a developer
actually waits for, because the first thing anybody does is export:

| | built in | shared |
| --- | --- | --- |
| a first export, every row a new file | 241ms | 1,724ms |
| planning again, every row unchanged | 249ms | **689ms**, was about 1,700ms |

The 163,000 comparison entries were 1,417ms of a 1,710ms plan, in the pipe and the two JSON
passes either side of it. What is left of the shared planner's 440ms is the module sources
going out and `leftRawCode` coming back, and that stays: it is what an apply writes, and it has
to be their bytes for the two products to write the same file.

A first export is unchanged at 1.7s, on purpose. Those rows have real differences to draw and
their planner is the one that draws them.

### What the statuses mean

| status | export | import |
| --- | --- | --- |
| `will-create` | the folder has no such file | the project has no such module |
| `will-write` | the file is there and differs | not applicable |
| `will-update` | not applicable | the module is there and differs |
| `will-remove` | a file naming no live module, and only with `mode=trueUp` | a module with no file, and only with `mode=trueUpStandardClass` |
| `unchanged` | both sides already agree | both sides already agree |
| `skipping-import` | not applicable | a worksheet or UserForm the project does not already have |
| `read-error` | not applicable | the file would not read |

### Which planner decided it

Two planners answer this, and the plan says which one did:

```js
const plan = await api.syncPlan("export", { folder });
plan.planner;   // "xlide" or "builtIn"
```

**`xlide`, the default.** The companion editor's own `moduleSyncPlan.ts`, imported out of the
xlide_vscode checkout the same way the analyzer is and running inside the engine. The two products
then decide identically about file names, module kinds, staleness and what counts as a change,
because it is one implementation. Defects in it are fixed upstream and both products get the fix.

**`builtIn`.** The same decisions worked out in the add-in, in `Xlide.Vbe.Core.Sync`. It needs no
engine, so import and export keep working when the engine is down.

Switch it with the setting, from the api or from the Settings dialog:

```js
await api.settings({ syncEngine: "builtIn" });
(await api.settings()).syncEngine;
```

**The fallback is silent, and this is why the plan reports its planner.** If the shared planner
cannot be reached the built-in one answers and the request succeeds, because a developer pressing Export
while the engine is starting should get their export, not a lecture. It is written to the log every
time. Do not write a test that asks for one planner without asserting `plan.planner`: this suite
did, and passed 31 checks against the built-in planner while asking for the shared one, because the
modules could not be serialised to the engine and nothing said so.

### Some things worth knowing before driving it

**A document module and a UserForm cannot be created from a file.** A sheet belongs to its
workbook and a form's designer is not in its `.cls`. Those rows come back `skipping-import` with a
warning rather than failing at apply time. When the module already exists, its code is replaced
normally.

**Removing is always opt-in and never touches a document.** `mode` defaults to the safe value in
both directions, and import true-up only ever deletes standard and class modules.

**Importing a class can change modules you did not select.** VBA unifies identifier case across a
project: bring in a class named `Tally` and a local variable spelled `tally` in some other module
becomes `Tally`. The next export will honestly report that module as changed. This is the editor's
doing, not this product's, and it is a good reason to look at an export plan after an import.

**An exported file is never half written.** The content goes to a `.<name>.xlide-partial` file
beside its destination and is then moved over the top, which the file system does as one step,
so nothing reading the folder can catch a truncated module: not the companion editor watching
it, not a build, not another Excel importing from the same folder. That last one is the case
this side has no lock for, and it is the reason the write is shaped this way rather than a
lock being added. A `.xlide-partial` left behind is junk from a process that died mid-write; it
is never read as a module, because it is neither a `.bas` nor a `.cls`.

**`action=browse` blocks.** It raises the system's folder chooser, which is what the dialog's
Browse button uses, and it does not answer until somebody closes it. A harness sets the folder with
`syncSettings({ folder })` instead.

## 3a. Language, and what the host will not store

Two tests, because there are two different exposures and only one of them is ours.

```bash
node engine\test\language.mjs          # headless, a gate step, 18 scripts
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

### The other thing the host will not store: a tab

> **VBA's code store cannot hold a tab character.** The editor expands every one it is handed to
> the next four-column stop. Measured 2026-08-07 on both write paths this product uses
> (`AddFromString` for a whole module, `DeleteLines` + `InsertLines` for a small change) and for
> tabs that are not leading whitespace either: `"    Dim n\tAs Long"` comes back as
> `"    Dim n   As Long"`.

So there is **no "indent with tabs" setting**, and there cannot be a working one. While there
was, the page indented with tabs and the workbook held spaces, and the two disagreed for as long
as the module stayed open, which is the one thing a surface covering the host's own editor must
never do. It was removed on the developer's call the day it was measured. A `format.useTabs` in
an older settings file is ignored rather than read.

What remains is **Indent size**, in spaces, governing typing, smart Enter, Backspace and Format
Module alike. Backspace is what makes that bearable: with `useTabStops` on, pressing it in a
line's **leading whitespace** takes back a whole indent level rather than one space, and
everywhere else it deletes one character as it always has.

```bash
node tools\harness\format-positions.mjs   # positions, parity, and the Backspace behaviour
```

`act("backspace", {times})` drives it. The `type` route cannot: it goes through
`trigger("keyboard", "type")`, which only ever inserts.

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
await api.component("add", { kind: "class", name: "Account" });  // the words work too
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
