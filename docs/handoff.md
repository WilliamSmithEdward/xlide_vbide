# Handoff

Everything needed to pick this up cold. Rewritten at the end of the session that made the surface
the whole visible editor; the previous version described a foundation, this one describes a product
with a short list of holes.

Read this, then [lessons.md](lessons.md) for the long-form findings with evidence,
[architecture.md](architecture.md) for the design, and [decisions.md](decisions.md) for choices
that would be expensive to reverse.

## START NEW SESSION HERE — 2026-08-02 18:00

**State.** The editor is in daily live use against a real 26,000-line workbook, and today's
second stretch closed every reported defect and the start-up complaint:

- **Boot is 181ms, was 2164ms.** The two seconds were the WebView2 folder mapping streaming
  the 3.4MB bundle through the browser's host pipe (~2MB/s); `LoopbackPageServer` now serves
  dist over 127.0.0.1 (ephemeral port, per-session path token, GET/HEAD, one directory), with
  the mapping kept as fallback. The ready log line itemises html/request/fetch/compile
  permanently (lesson 23), so a start-up regression names its own stage.
- **Typing structure is page-local.** Smart Enter (block closers, comment and With-member
  continuation), Smart Tab, Smart Backspace, and loop-iterator sync run in the page on the
  extension's own bundled helpers (`ui/editor/src/typing.ts`; tsc sees `src/spec/xlide-spec`
  declarations, esbuild aliases to the real sources) — no round trip between Enter and its
  `End If`. Canonical casing stays engine-side.
- **A line being typed gets no verdicts** (`ActiveLineHold`, Core, tested): typing hides
  findings touching that line from squiggles, panel, and badges at one filter point; the caret
  leaving republishes from cache. The VBE model, adopted deliberately (lesson 21).
- **The tree cannot blank an unfolded list**: outline timeout resolves null (never empty), the
  host marks failures, one request per module with a trailing refresh, engine memoises the
  outline per source string, and requests no longer carry the module text (lesson 21).
- **Immediate window mirrors again** — identified by caption match, not visibility diffing
  (lesson 22); a failed attach retries on first evaluation. Developer-confirmed.
- **Cancelled-shutdown revive waits for the save prompt**: an app-modal dialog disables the
  frame; only an enabled frame across two ticks revives (lesson 22). Reproduced both ways.
- Hover answers for class receivers (`ROneCOne` in `ROneCOne.DataView(...)`), the minimap
  slider is always visible, the tab X survives pointercancel and drag suppression.
- `ProjectReader.ReadAll` now runs on the host thread via `AnalysisService.HostMarshal`
  (declined marshal retries; unanswered marshal abandons the pass). The old pool-thread debt is
  paid.
- **Decision 10**: the product never requires "Trust access to the VBA project object model" —
  OnConnection instance + `AccessibleObjectFromWindow`/`Application.Run` only. Harness scripts
  are the sole (dev-only) exception.

**The scale architecture (unchanged, still the spine):** page ships Monaco change ranges (full
text under 64KB) → shim reconstructs and streams `textDocument/didChange` → engine holds one
live string per module; requests carry offsets only; FIFO order with the notification's wait
registered synchronously. Write-back is a line diff at one anchor. Analysis yields to typing
(live pass deferred to line-leave on large modules; full pass 3s-gated with catch-up). Every
page sink is idempotent and tree rebuilds restore scroll (lesson 20).

**Open threads, in order of likely next ask:**
1. `xlide_vscode` carries TWO local commits, neither pushed: `6453c20` (class/UserForm
   receivers in the shared resolver) and `1f9d8b8` (hover for object-module surfaces as bare
   receivers). Suite green at 2478 after both. Push is the developer's call.
2. #24 project-qualified addressing: the multi-workbook build, STARTED. The root defect was
   worse than "the UI is unqualified": `ProjectId` WAS `VBProject.Name`, which is "VBAProject"
   for nearly every workbook, so two open workbooks were one project to the engine. DONE:
   `ProjectReader.Identity` — id is the lowercased full file path (unique among saved
   workbooks, stable for the session), Name only for never-saved workbooks (two unsaved
   workbooks remain a residual collision, accepted); `ProjectSnapshot` carries `DisplayName`
   (file tail) alongside. Remaining, in dependency order:
   a. Pane→project: `CodePaneTracker.ReadOpenComponents` walks pane→CodeModule→Parent; extend
      one hop (component `Collection` → `Parent` = VBProject) and derive the same Identity, so
      a `CodePane` carries (ProjectId, Component). Captions cannot do this — a maximised MDI
      child's caption drops the workbook — the object model must.
   b. Shown identity: `ShowModuleInSurface`/`FindComponent`/`WriteModule`/`GoTo` currently find
      components by bare name across all projects; they must take (projectId, name), defaulting
      to the active pane's project. Track `_shownProject` beside `surface.Module`.
   c. Findings: the `Finding` record gains Project (from the snapshot when converting);
      `PublishMarkersForShownModule` filters by (shown project, module); the problems panel row
      shows the workbook when two projects share a module name; tree badge counts key on
      (project, component) — the explorer's rows are already per-workbook.
   d. DONE, probe-verified: `_moduleHomes` maps each name to EVERY home it has (swapped
      wholesale per pass); one `ResolveHome` replaces nine lookup copies. Requests about the
      shown module (everything the editor asks) resolve to `PreferredProject` even before a
      pass seeds it — the probe caught live text poisoning a same-named module in the other
      workbook when resolution fell back to the seeded home. The outline alone asks with
      `aboutShownModule: false` so tree rows for other workbooks keep their own homes; a
      COLLIDED name's tree outline still mis-addresses until (e) sends the project.
   e. Page protocol: `activateModule`/`closeModule`/`outline`/`navigate` gain an optional
      `project`; `setModules` sends (project, module) pairs; tabs render "Module1 — Book2.xlsm"
      only when bare names collide; explorer clicks pass their workbook.
   f. DONE (not yet probe-exercised — needs a workbook to close mid-session): the pass diffs
      `_openProjects` against the snapshot ids, closes the vanished projects engine-side
      (`EngineClient.CloseProjectAsync`), and prunes their homes.
   Steps (b) and (c) are DONE and harness-verified: `_shownProject` tracked from the pane's
   project, `FindComponent` project-scoped, writes/resync/live-merge target the shown project,
   findings carry Project, markers filter by it (null-tolerant during the transition).
   The colliding-modules probe (scratchpad `Test-CollidingModules.ps1`) then EXPOSED two
   defects the remaining steps must fix, both in its log:
   g. DONE, probe-verified: when the caption match is ambiguous (two same-named panes), the
      show consults `ActivePaneOwner` — the object model's ActiveCodePane names component and
      project without a caption. The probe log now shows Book1's 80-character module where the
      wrong-project 113-character one used to appear.
   h. DONE, probe-verified: showing a module whose project the engine has never been seeded
      with triggers `Reanalyse` (2s-gated; `AnalysisService.KnowsProject`, `_openProjects` now
      concurrent). The probe log reads "analysing 2 project(s)", each workbook's BrokenModule
      carrying its own finding at its own line.
3. Latent: something host-side once published a CHANGING payload every second in the
   developer's environment (never reproduced against the scratch workbook). The page now logs
   `page: tree: ... push changed, <diff>` whenever a push gets past its identity guards, so the
   next occurrence names the oscillating field itself.
4. Then the standing backlog: #20 right-click curation (needs the developer), #22 split
   groups, #12 settings, #13 tests panel, #14 debugging/forms designer, #10 typelib backfill,
   #9 the C# analyzer port.

**Before touching anything, know these:**
1. Registry writes from the agent shell are a MIRAGE (sandbox COW; lesson 17). Registration is
   the developer running `tools\Register-DevShim.ps1` themselves; verify with their regedit.
2. Cross-thread work into the browser rides the overlay's action timer, never a posted message
   (lesson 18). Log lines carry `[host]`/`[tN]` — and read the log's DATA before theorizing:
   lessons 19 and 20 were both settled by one log line after theories failed.
3. `PixelRect` is four EDGES, not origin-and-size (lesson 19).
4. The deploy dance: close Excel (the developer authorised force-close for republish) →
   `tools\dev.ps1 -NoRun`; page changes also robocopy `ui\editor\dist` into the publish tree;
   engine changes also `node build.mjs --package` in `engine/` (the shim loads the exe from the
   repo path). The DLL is locked while Excel runs.

Repository: `F:\GitHub\xlide\xlide_vbide`, public at
<https://github.com/WilliamSmithEdward/xlide_vbide>. The tree is committed and pushed at the end
of every session; `git log` names what each one did.

## 1. What this is, in one paragraph

An add-in that replaces the visible VBA editor from inside it. A native COM in-process server
(NativeAOT, no runtime in Excel's process) loads into the VBE, covers the editor with a WebView2
surface running Monaco, and drives the real editor underneath through its object model and command
bars. The developer sees our toolbar, project explorer, module tabs, editor, problems panel,
Immediate panel, and status bar; the native components stay alive as invisible engines. An
out-of-process language engine supplies diagnostics. One EXE installs it per user; nothing else is
required on the machine.

## 2. The two rules that decide the design

**The module is the source of truth. The surface is a view of it.** The compiler, the debugger, the
saved workbook and the analyzer all read `CodeModule`; nothing reads the surface. An edit that has
not reached the module has not happened. When the two disagree, the module wins.

**With one amendment, from the user, after living with the first rule:** the editor rewrites what
it is given (respells keywords, completes a procedure's parentheses, inserts blank bodies), and
those rewrites are exactly what a developer is mid-way through typing themselves. So resync
comparisons run against a baseline of *what the module read back as after our last write*, never
against the surface text, and an unfinished edit is never overwritten. Typing ergonomics follow
xlide_vscode, not the VBE. `Sub Later` stays `Sub Later`.

Anything added here should answer: what happens when the surface and the module disagree, and how
does it get back in step without the developer losing work.

## 3. Machine setup

Nothing below is on the system PATH; forgetting any of it fails confusingly rather than clearly.

```powershell
# The SDK (user-local .NET 10.0.302).
$env:PATH = "$env:LOCALAPPDATA\Microsoft\dotnet;$env:PATH"

# NativeAOT publishing finds the linker through vswhere. Without this the error text reads as a
# linker fault, not a discovery fault.
$env:PATH = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer;$env:PATH"

# Required to RUN anything framework-dependent, or it reports .NET is not installed.
$env:DOTNET_ROOT = "$env:LOCALAPPDATA\Microsoft\dotnet"
```

VS Build Tools 2026 (C++ workload), Node 24, Excel 365 x64 16.0.20228, VBA 7.1, Windows 11.

## 4. Running things

```powershell
tools\dev.ps1                                  # build, test, register, verify in a real editor
tools\dev.ps1 -Unregister                      # leave the machine clean
tools\harness\Get-EditorScreenshot.ps1 -KeepOpen   # launch, capture at true DPI, leave running
tools\Compare-Lexers.ps1                       # differential gate for the analyzer port
installer\build.ps1                            # produce xlide-setup.exe
cd ui\editor; node build.mjs                   # rebuild the surface bundle (esbuild + tsc)
cd engine;    npm run build                    # the TS language engine sidecar
```

The dev loop is: rebuild the bundle if the page changed, `dotnet publish` the shim, kill Excel, run
`Get-EditorScreenshot.ps1 -KeepOpen`, read `%LOCALAPPDATA%\xlide_vbide\logs\shim-*.log`. The
harness waits for `analysis:` and `editor surface: ready` lines before capturing, and restores
`LoadBehavior` and clears Excel's resiliency keys before each run. The fixture is
`tools\harness\fixtures\scratch.xlsm`, regenerated by `New-ScratchWorkbook.ps1`.

Throwaway probes from this session live in the session scratchpad (not the repo): window-tree
dumps, focus probes, menu dumps, synthetic-input senders. Synthetic input (SendKeys, mouse_event)
is dev-harness only and unreliable — foreground restrictions make it fail silently — and the user
has said plainly it must never be a production mechanism.

## 5. What works, verified against real Excel

- Add-in loads (NativeAOT COM server, IDispatch declared explicitly), engine connects over a named
  pipe, diagnostics arrive with correct positions.
- The surface covers the frame's whole client area below the command bars: toolbar (21 commands),
  project explorer with per-module problem badges, module tabs with error counts, Monaco with VBA
  grammar and themes, problems panel, Immediate panel, status bar with transient notices,
  draggable and keyboard-operable splitters.
- Typing writes back to `CodeModule` 400 ms after the last keystroke and immediately before
  run, step, module switch, and shutdown; re-analysis follows each write. Verified by reading the
  module back out of the VBA project.
- F5, F8, Shift+F8, Ctrl+F8, F9, Ctrl+S work; F1 opens the command palette instead of VBA help.
  All routes (key, toolbar button, glyph margin) go through one `ExecuteEditorCommand` path.
- Breakpoints render as red dots, with a hover preview dot and a visibly delineated margin.
  Non-executable lines are refused with a status-line notice before the VBE's modal can appear.
- Break mode is detected from `VBProject.Mode`; the stopped line comes from the pane's
  `GetSelection` (four by-ref ints) and is marked and revealed in Monaco.
- Immediate panel: `?expr` evaluates (scratch module + `Application.Run`, module removed always),
  plain lines run as statements, history on arrow keys. `?2*21` -> `42` verified. Declined during
  break, with the reason shown: adding a module would reset the project.
- `Debug.Print` output is captured from the hidden native Immediate window via UI Automation
  TextPattern and appended to our panel. Verified end to end with the window closed.
- Native Project Explorer, Properties and Immediate windows are closed via the object model; the
  native Standard toolbar is hidden (its commands were moved to ours first); title bar, caption
  and window border are dark via DWM attributes. The frame's own pale inner line is covered by the
  surface extending to the frame edge.
- The menu bar is ours. The surface draws its own File..Help bar and covers the native one, but
  only once its page has reported ready: a covered bar with no page behind it would take the only
  route to References and Options away. Each menu is read live from CommandBars at the moment it
  opens (VbeMenus.Read), so captions, enablement, checkmarks, shortcuts and the dynamic Window
  list are always the editor's own, and localisation comes free. Items execute by position chain
  (VbeMenus.ControlAt), never by identifier. Routed exceptions: run/step/save/breakpoint go
  through ExecuteEditorCommand (bookkeeping), undo/redo/find/replace run in the surface,
  Clear All Breakpoints also clears the surface's drawn record, View -> Immediate opens our
  panel, and the project explorer item answers with a notice. Anything that opens a native
  window or docked toolbar makes the surface retreat to the document area at once
  (RefreshSurfacePlacement) and withdraw its own menu row (setChrome), so nothing native is ever
  reachable-but-covered. Verified live: staged covering in the log (below command bars, then
  0,0 after ready) and 11 menus read; dropdowns, submenus, separators, disabled/checked items,
  shortcut column, execution paths, Alt+letter, F10, arrows and Escape verified in the browser
  demo (demoTransport serves a canned tree).
- Tabs have modern ergonomics: hover/active close button, middle-click close, Ctrl+W and
  Ctrl+F4 close the active module (claimed at the accelerator hook; unclaimed Ctrl+W tells the
  browser to close its own window), Ctrl+PageUp/PageDown cycle, drag to reorder with the order
  surviving host updates. Right clicks answer with class-curated menus everywhere (see 9.2).
  Closing goes through the pane's window; inserting goes through VBComponents.Add and opens what
  it made; commands that live only inside menus (References 942, ProjectProperties 2578) are
  found by a recursive descent that is sound only for UNIQUE identifiers.
- The Properties window is replaced, and the pane mirrors the native one. It follows the
  explorer's SELECTION (single click selects and publishes properties; double click or Enter
  opens the code, exactly like the native tree); an object header names the selection and its
  class ("Sheet1 Worksheet"); the code name appears as "(Name)" and sorts first; booleans are
  True/False dropdowns; the rest are inputs or read-only rows. Edits write in the type the
  property holds; refusals are reported in the editor's own words; renaming through "(Name)"
  rekeys the write baseline and breakpoint record, and only reloads the editor if the renamed
  module is the one being shown. View -> Properties Window opens this pane. CRITICAL SAFETY RULE
  (lessons.md 15): for document components, only the allowlisted browsable properties are read,
  names are enumerated before any value, and each value is read once - reading an unlisted
  getter such as a workbook's mail properties STARTS THE MAIL SYSTEM. The allowlists stand in
  for the type library's browsable flags until those are read directly (IntelliSense track).
  Resizable by a gripped splitter, like the other two.
- COMPLETIONS are served by the engine (mission pillar 3, first slice). The engine's
  textDocument/completion reuses xlide_vscode's analyzer resolvers directly (memberAccess,
  identifierCompletion, keywordCompletion) with project facts assembled from the seeded modules
  and the live text of the module being typed. After a dot: the verified Excel object model's
  members for the resolved receiver (ThisWorkbook. -> 186 members in the smoke test). At an
  expression position: in-scope identifiers, host globals, runtime functions, other modules'
  procedures, plus keywords; keyword-exclusive grammar positions show only keywords. The pipe
  client serialises calls (the protocol pairs answers by position); the surface asks with a
  correlation id + UTF-16 offset; the host answers off-thread and marshals back on the overlay's
  action timer (OverlayWindow.RunOnHostThread — never a posted message, see lesson 18); the page
  maps analyzer kinds to Monaco icons and inserts '${'-bearing insertText as snippets. Engine
  smoke test covers it end to end; log line per request:
  "completion: <module>@<offset> -> N item(s)".
- HOVER and SIGNATURE HELP ride the same channel (textDocument/hover, textDocument/signatureHelp),
  reusing resolveHover and resolveSignatureHelp over the shared project assembly
  (engine/src/moduleContext.ts — one assembly so completion, hover, and call tips describe the
  same project). Hover renders the declaration line as VBA code plus origin/visibility details
  and markdown docs; call tips carry the full parameter list with the active parameter tracked
  through both parenthesized and parenless calls, triggered on '(', ',' and space exactly as the
  extension triggers them. Smoke-tested end to end (local, host global, cross-module procedure,
  MsgBox tip, parenless second argument, null cases).
- 95 unit tests green; the lexer port agrees with the reference on 175/175 corpus files.

## 6. How the pieces fit

```text
EXCEL.EXE
  Xlide.Vbe.Shim.dll          native, no runtime loaded into the host
    AddInSession              owns lifetime + the editing contract: write-back, resync baseline,
                              debug state, breakpoint bookkeeping, panel routing, window hiding
    CodePaneTracker           which panes exist, where; VisiblePanes(); pane/frame class names
    EditorSurface             overlay + the page protocol (hold-until-ready queue, newest per
                              kind, order preserved; loadDocument before setDiagnostics)
    OverlayWindow             the window itself + two WM_TIMER timers (write debounce id 1,
                              poll id 2)
    VbeCommands               executes VBE commands via CommandBars (bars walked; FindControl
                              does not search menus and returns nothing rather than failing)
    ImmediateEvaluator        compiles a line into a scratch module, runs it by name, removes it
    ImmediateReader           reads the hidden native Immediate window via UIA TextPattern
    HostApplication           reaches Excel's Application from in-proc (XLMAIN->XLDESK->EXCEL7 +
                              AccessibleObjectFromWindow OBJID_NATIVEOM); never the ROT
    AnalysisService           owns the engine; Reanalyse() after writes
    DispatchObject            late-bound calls; property puts carry DISPID_PROPERTYPUT; EXCEPINFO
                              captured so VBA's own error text survives
  WebView2 process            the surface page

xlide-engine.exe              TS analyzer sidecar, named pipe, JSON-RPC line-framed

ui/editor/src
  main.ts        boot, per-feature Monaco imports (features are opt-in!), boot timings reported
  shell.ts       tabs, panel tabs (Problems/Immediate), splitters, notices, status line
  explorer.ts    project tree grouped the way the VBE groups it
  toolbar.ts     command table; host vs editor dispatch; availability check draws no dead buttons
  bridge.ts      every message both directions; syncDocument preserves undo/caret; markers re-set
                 after any whole-document replace
  format.ts      VBA formatter: indentation + canonical keyword case, nothing else
  vba.ts         Monarch grammar + CANONICAL_KEYWORDS (shared with the formatter)
```

## 7. Measured VBE facts (re-measure only if a host build disagrees)

Registry: `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64`, subkey = ProgID. Frame class
`wndclass_desked_gsk`, document area `MDIClient`, and the code panes, Immediate, Locals, Watch and
Object Browser windows all share class `VbaWindow` — a window's class says nothing about what it
is; only the object model knows.

Add-in discovery is **HKCU-only** — vbe7.dll's own strings carry
`HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0` and `...\Addins64`, and Rubberduck's installer
writes the Addins keys under HKCU even when elevated (HKLM there is only .NET COM classes). There
is no HKLM enumeration to fall back on, and the per-user registration needs no elevation — that
is the entire documented mechanism.

**The dev-loop registration mirage (lesson 17, read it):** the agent tool shell virtualizes
registry writes, so registrations written from it exist only in the sandbox's private layer.
Harness-launched Excel inherits that layer and loads the add-in; Excel the developer launches
reads the real hive, which never got the key. A day went to Click-to-Run/App-V theories before a
regedit screenshot showed the truth. Registration on the dev machine is therefore done by the
developer running `tools\Register-DevShim.ps1` in their own terminal (no admin); the published
shim path is stable across rebuilds so one real registration survives the dev loop. Elevated
child processes escape the sandbox (UAC-relaunched writes land for real) — useful, but not a
substitute. Verify persistence with the developer's regedit, never with in-sandbox reads.

Click-to-Run keeps machine-level VBA values (`Vbe71DllPath`) only inside its overlay
(`HKLM\...\ClickToRun\REGISTRY\MACHINE`); real HKLM has no VBA branch on C2R machines. The
installer writes the per-user registration only and never prompts for elevation; when it already
runs elevated on a C2R machine it also plants the overlay copy, and `--overlay-only` runs that
step deliberately (`tools\Register-InOfficeOverlay.ps1` is the standalone form).
`Register-MachineWide.ps1` (real-HKLM mirror) was retired — the VBE never reads that hive.

`Window.Type`: 0 code, 2 object browser, 3 watch, 4 locals, 5 immediate, 6 project, 7 properties,
10 linked frame, 15 our tool window. `VBProject.Mode`: 1 = break, 2 = design.

Command IDs (measured by enumerating CommandBars): Run 186, Break 189, Reset 228,
ToggleBreakpoint 51, StepInto 188, StepOver 194, StepOut 2559, RunToCursor 1811, QuickWatch 229,
AddWatch 1820, EditWatch 940, CallStack 620, Compile 578, ClearAllBreakpoints 579,
SetNextStatement 1812, ShowNextStatement 1813, Comment 192, Uncomment 2552, Save 3, Undo 128,
Redo 129, Find 141, Replace 313, ObjectBrowser 473, ImmediateWindow 2554, LocalsWindow 2555,
WatchWindow 2556, ProjectExplorer 2557, PropertiesWindow 222, References 942, Options 522,
Macros 930, ProjectProperties 2578, InsertProcedure 559, InsertUserForm 512, InsertModule 3039,
InsertClassModule 2579, Import 524, Export 525.

**Menu item IDs are NOT unique.** 746 is shared by New Project, Close Project, Remove <module>,
Make, four Insert placeholders, Digital Signature, and MSDN on the Web; 830 by every window-list
entry; 761 by every toolbar toggle. The current command set is by-ID and every ID in it is unique
(checked against the full dump), but menu replication MUST execute by path
(`bar.Controls.Item(i).Controls.Item(j).Execute()`), which is verified working live.

`MenuBar.Visible = false` returns E_FAIL — a genuine refusal (the identical call hides the
Standard toolbar). The menu bar cannot be hidden, only covered.

DWM: attribute 20 (19 on older builds) = dark title bar; 34 border colour, 35 caption colour,
36 caption text; values are COLORREF `0x00BBGGRR`; refused harmlessly before Windows 11. The
frame's client area is inset ~11 px and a 1-2 px pale line at that inset is drawn by the frame
itself; only covering it hides it. The menu bar is drawn by Office and no attribute reaches it.

A full menu tree dump (11 menus, ~90 items, captions + IDs + nesting + enabled state) was taken
2026-08-01; regenerate with a CommandBars walk when needed.

Pane lifecycle facts, measured while making tabs behave:
- With maximised panes the editor keeps a WINDOW only for the ACTIVE pane and destroys the
  others. The window map says one pane however many are open; only `CodePanes` knows the open
  set. Never feed the tab list from windows.
- `CodePanes` keeps a corpse entry for a just-closed pane: `Count` still counts it and
  `Item(n).CodeModule` throws. Read the collection per item, tolerantly, always.
- `CodePane.Show` displays but does not activate a pane that is already open behind another; no
  activation means no window event. Set `ActiveCodePane` directly - and that is an OBJECT
  assignment, which must go out as `DISPATCH_PROPERTYPUTREF`; an ordinary put is refused.
- Activating a native pane MOVES KEYBOARD FOCUS to it. Return focus to the surface
  (`controller.MoveFocus`) or every keyboard shortcut works exactly once.
- Ctrl+W and Ctrl+PageDown/PageUp are BROWSER accelerators (close window, switch browser tabs);
  unclaimed they never reach the page. Claim them at the accelerator hook like F1.
- The page has a `trace` message that writes into the shim log (`page: ...`), which is how a
  page-side silence is told apart from a transport one. `globalThis.xlideBridge` is reachable
  from devtools for the same reason.
- `tools`-side probes may use synthetic input against the harness (never in the product), BUT
  the user has judged SendKeys-style probing unreliable and it misfired once (it executed
  File -> Import through the menu). The standing verification loop is: the user drives, and the
  log answers - both halves log every evaluation, its outcome, and each change in what the
  hidden window holds, with non-printables escaped.

Immediate window facts, measured while making capture behave:
- The UIA TextPattern reading of the (hidden) Immediate window ends with a trailing empty line
  and a NUL, U+0000. New output is inserted BEFORE that tail, so the raw text never simply
  grows. Diff on readings trimmed of ALL trailing whitespace-or-control characters, by class,
  never by a list; a raw comparison first replayed the whole buffer per evaluation and then
  swallowed every print.
- A reading of empty is a project-reset artefact, not a cleared window (it is hidden; nobody
  can clear it): ignore it, never adopt it as the baseline. A reading that does not continue
  the baseline means trimming; adopt it silently.
- An unhandled run-time error in code started by Application.Run drops the editor into break
  mode INSIDE the calling frame and shows its dialog. The evaluator's scratch procedure carries
  its own On Error handler and returns the language's message as a Chr$(1)-marked value; the
  scratch module (ImmediateEvaluator.ScratchModule) is filtered out of the pane picker, the tab
  list, the explorer and the debugger's module switch, because it surfacing was one flash per
  evaluation.
- Evaluation must check VBProject.Mode FRESH, not the cached break flag: evaluating during an
  unnoticed break adds a module to a stopped project and fails incomprehensibly.
- A successful statement prints nothing (native parity); ?expr prints its value; errors print
  the editor's own message.

## 8. Lessons that cost real time this session (long form in lessons.md)

- **A dispatch property put carries a named argument** (DISPID_PROPERTYPUT). Two setters passed
  the value positionally; every assignment failed with no message; and the failures were recorded
  as "the VBE refuses to size tool windows". That claim was WRONG and is retracted in lessons.md.
  A failed call with no message says nothing about whose fault it is; compare against a sibling
  call that works (the boolean setter did) before blaming the host.
- **Pass EXCEPINFO to IDispatch::Invoke** or the callee's error text ("Type mismatch") is thrown
  away in favour of "Arg_COMException". For the Immediate panel that text is the answer.
- **The VBE rewrites AddFromString input.** Read the module back and adopt that as the baseline;
  never compare the module against the surface, or its own tidying reads as an external change
  and stomps the developer mid-keystroke.
- **Whole-document replaces collapse Monaco markers onto the end of the text.** Keep the host's
  markers and re-set them after any sync; `forceMoveMarkers: false`; restore selections around
  the edit so the caret and undo stack survive.
- **Undo/redo are not Monaco actions.** `editor.trigger("xlide", "undo", null)`; the toolbar's
  availability check must special-case them or their buttons are (correctly) never drawn.
- **A key claimed at the WebView2 accelerator hook never reaches the page.** Claiming F1 means
  explicitly messaging the page to open the palette. Keys the VBE owns (F5, F8, F9) must be
  claimed there too, or the browser acts on them (F5 reloads and destroys the document).
- **The editor reports windows Visible=true before creating them.** Identify the Immediate window
  only once the editor is genuinely up (the surface has a pane): show it, snapshot visible
  `VbaWindow`s recursively (not direct children — depths vary), hide it, diff.
- **The native Immediate window is unreadable by every normal means** (HWnd = 0 via the object
  model, WM_GETTEXT answers the caption, MSAA children fail) **but UIA TextPattern reads it in
  full, even while hidden**: CoCreate CUIAutomation, ElementFromHandle, FindFirst descendant,
  GetCurrentPattern(10014), QI IUIAutomationTextPattern, GetDocumentRange, GetText(-1). Poll it
  (300 ms, only while the panel is watched); do not subscribe — the text-changed event fires on
  the VBE thread while it is inside the developer's running code.
- **PrintWindow captures lie about parts of this window.** The menu bar renders light on screen
  and dark in capture. Measure pixels for edge/colour questions; only the user's own screenshot
  is on-screen truth. The capturing process must also be per-monitor DPI aware or every capture
  silently crops.
- **The scratch workbook must be .xlsm.** Excel saves an .xlsx by silently dropping its VBA.
- **VBA identifiers cannot start with an underscore** (bit the scratch-module naming).
- **VBProject.Mode is the execution state.** Command enabled-ness is not (Reset is enabled in
  design mode too), and running does not block the command that starts it — poll for up to 20 s
  after any run/step command, keep polling while stopped, stop when idle.

## 9. Open and known, in priority order

**THE MISSION (user, 2026-08-01), which frames everything below:**
1. Replicate all base functionality of the VBE editor.
2. Replace all native windows with custom GUI implementations.
3. Bring the editor into parity with xlide_vscode as far as feasible: tab completion, member
   (dot) menus, IntelliSense, and the rest of its feature set. xlide_vscode is the user's own
   MIT project at `F:\GitHub\xlide\xlide_vscode` and porting from it is sanctioned.

1. **Menu bar: first live click-through.** Built and verified as far as automation reaches (see
   section 5); no person has clicked the real menus yet. Watch the log for `menu: [n] read` on
   the first dropdown and `menu: [...] executed`. Known edges, all accepted for now: the native
   (light) bar shows during the ~2 s page boot, then ours covers it; keys the host claims (F5,
   F8, F9) still act while one of our menus is open, because the accelerator hook runs before
   the page sees anything; Alt alone does not focus the bar (Alt+letter and F10 do); floating
   native toolbars are left alone and uncontested; disabled items are as fresh as the moment the
   menu opened, and an execute on a stale item answers with a notice.
2. **Right-click surfaces, per object class (user directive 2026-08-01).** First pass SHIPPED:
   curated menus on tabs (close / close others / close all), explorer components (open or open
   code, rename, close), the project header (insert module/class/form, References, Project
   Properties), the breakpoint margin (toggle, clear all), problems rows (go to, copy) and the
   Immediate log (clear); the editor's own menu gained the host's run/breakpoint commands.
   Still to add as the features land: Export File and Remove (both need a native file dialog or
   a confirm flow), form designer items, and re-validation of every menu as classes grow.
3. **~2.1 s surface start-up**, entirely bundle fetch/parse (editor construction is 50 ms). The
   page reports its own timings in the ready message on every run. Untried: bundle splitting,
   warming during host start-up, cached compilation. (Task #17.)
4. **Panels are fixed.** The user wants VS-style drag-to-any-edge / stack-as-tabs.
5. **Locals, Watch, Object Browser are still native** (deliberate: no replacement yet; hiding
   them would remove features). While any is open the surface retreats to document-area bounds,
   so the frame's pale line reappears — accepted trade, logged when it happens.
6. **Debug.Print lands only while the Immediate panel is open** (poll-gated). The text still
   exists in the hidden window; backfill-on-first-open is a small known gap.
7. Stepping's current-line marker is implemented and log-verified, not yet pixel-verified in a
   live break by the user.
8. Forms designer untouched. Analyzer parser port (~3,400 lines) next in that track; lexer done.
   IntelliSense backfill from type libraries not started. No settings surface. Nothing signed.

## 10. Conventions and user directives (binding)

- Never mention the other add-in product in anything public. Clean-room; cite only Microsoft
  specs and documented interfaces.
- ASCII prose, no em dashes, wrap at 100. Comments explain constraints, never narrate. Commit
  messages say what changed, why, and what the defect looked like.
- Report status literally; a check that passes by not looking hard enough is worse than none.
- The user rejects backwards-compatibility hacks — full refactors are fine.
- No synthetic input (SendKeys) in production, ever.
- The whole UI should end up ours: consistently dark, VS-style ergonomics, the VBE alive
  underneath as the engine. The module is the source of truth; typing follows xlide_vscode.
- Every native window should eventually be replaced by the surface (user, 2026-08-01): Locals,
  Watch, Object Browser, Properties, Call Stack, and dialogs wherever the object model allows a
  faithful rebuild (References and Macros are scriptable; parts of Options are not). Until a
  replacement exists, the native window stays reachable and the surface retreats for it; the menu
  routing table in RouteMenuCommand is where "open ours instead" gets decided per window.
