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

Excel loads VBE add-ins when the **VBE** starts, not when Excel does. Nothing is under test until:

```powershell
$excel.VBE.MainWindow.Visible = $true
```

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
| `assert` | `assert(that, {value, timeoutMs})` | states a claim and waits for it |
| `await` | `until(predicate, {waitMs})` | waits for a condition IN the page |
| `bench` | `bench(what, {n})` | times a scenario: min, median, p95, max, raw samples |
| `breakpoint` | `breakpoint(module, line, {project, state})` | set, clear or toggle |
| `capture` | `capture(window)` | a BMP of the window, through PrintWindow |
| `caret` | `caret(line, {module, column, project})` | navigates first when a module is named |
| `command` | `command(name)` | any editor command by name |
| `compile` | `compile({waitMs})` | compiles; errors as DATA, modal cleared |
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
| `module` | `readModule(name, project)` / `writeModule(name, text, project)` | through the session's own reader and writer |
| `perf` | `perf()` | raw placement and marshal durations |
| `placement` | `placement()` | forces a placement pass |
| `problems` | `problems(module)` | the analyzer's findings |
| `reload` | `reload({waitMs})` | reloads the page and waits for it |
| `state` | `state(timeout)` | shown module, mode, handles, rects, DevTools port |
| `stats` | `stats()` | uptime, memory, handles, GC, placement and marshal counters |
| `watches` | `watches()` | the Watch panel |
| `windows` | `windows()` | every editor window |

Also on the client, built from those: `waitUntilResponsive()` and `ask()`.

### Grouped by what you would drive it for

### Looking

| Want | Call |
| --- | --- |
| Is this session sane | `doctor()` |
| What is shown, and where the windows are | `state()`, `windows()` |
| What the surface holds TEXT for | `documents()` |
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
| Run script in the page | `ask(script)` — see the trap below |
| Reload the page and wait for it | `reload()` |
| Put the arrangement back | `resetLayout()` |

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

## 4. Where the api stops

These are not oversights; they are the boundary. Everything here needs COM or CDP.

| The api cannot | Because |
| --- | --- |
| Start or stop Excel | It lives inside the process it would have to start |
| Open, close or save a workbook | The door drives the EDITOR, not the host application |
| Add, remove or rename a VBComponent | Only `renameModule` through the page's own flow; a fixture's components are COM |
| Answer a QUESTION dialog automatically | Deliberate. It declines or acknowledges; it never agrees |
| Send real mouse or keyboard input to the page | Synthesised DOM events do not reach the editor's own mouse handling — use CDP |
| Reach a second Excel process | One door per add-in session; `discover()` lists them, each with its own port |
| Run in a Release build | The whole door is `#if DEBUG` |

---

## 4a. "Trust access to the VBA project object model"

**The add-in does not need it. The harness does.**

That setting — Trust Center → Macro Settings → *Trust access to the VBA project object model*,
`HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM` — is a barrier on the way IN from
outside. It gates `Workbook.VBProject` and `Application.VBE` when they are reached from
**automation**: another process, or VBA code, asking Excel for its VBA project. Off, those raise
1004, *"Programmatic access to Visual Basic Project is not trusted."*

| | Needs it | Why |
| --- | --- | --- |
| The add-in | **no** | It is a VBE add-in. The host hands it the `VBE` object at `OnConnection`, so it is already inside; it never asks Excel for a project. Every project access in `src/` is `VBE.ActiveVBProject`, never `Workbook.VBProject` |
| The debug api | **no** | The door runs inside the add-in's own process and calls the same objects |
| `Start-Excel.ps1` | **yes** | `$excel.VBE.MainWindow.Visible` is automation reaching in |
| `New-RenameFixture.ps1` | **yes** | `$book.VBProject.VBComponents.Add` |
| Any COM verification | **yes** | Reading a module back independently is the same reach |

This is worth caring about: needing a security setting changed before a product works at all is a
real adoption barrier, and xlide does not have one.

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

## 5. What still needs COM

Through `Start-Excel.ps1`'s attach, or the same three calls in your own script.

**Building a fixture.** Adding components, naming them, filling them:

```powershell
$component = $project.VBComponents.Add(1)      # 1 standard, 2 class, 3 form
$component.Name = 'Helpers'
$component.CodeModule.AddFromString($code)
```

`tools\New-RenameFixture.ps1` is the worked example. Note `Circle` is not available as a module
name — the Excel object library owns it, and VBA refuses with a bare HRESULT.

**Verifying independently.** The most valuable thing COM does here: read the VBA project back
*without going through the add-in*, so a feature is proved by the object model rather than by the
editor's own report of what it did. That distinction found two real defects on 2026-08-06.

```powershell
$code = $project.VBComponents.Item('Helpers').CodeModule
$text = $code.Lines(1, $code.CountOfLines)
```

**Running a procedure.** `$excel.Run('Module.Procedure')`.

> A `Run` against a project that does not compile fails for a reason that has nothing to do with
> what you asked. Check `compile()` first, or your experiment is measuring the wrong thing —
> which is exactly how a shadowing question got answered wrongly on 2026-08-07 before being redone
> in a workbook of its own.

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
