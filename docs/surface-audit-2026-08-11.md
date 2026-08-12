# Surface audit, 2026-08-11

A whole-repo walk for three things: gaps that stop the debug api from driving and observing Excel
first-class, code that can be removed or simplified, and performance opportunities. This is a
findings document, not a plan. It records what was established and how, so each item can be checked
against the code rather than taken on trust.

## Implementation status, 2026-08-11

Five of the eight batches are applied and in the working tree. The gate was run to PASS after each
one, and the Debug shim republished at page build 291.

| Batch | State | What it covered |
| --- | --- | --- |
| Leaks | done | `_hostApp` surviving `Stop()`, two `GetItem` loops that drop a wrapper on a throw, every `eval` permanently rooting its handler |
| Routes that answered a constant | done | `command`, `compile`, `pane close`, `undoRename`, `caret`, `breakpoint`, `component rename`, `assert` on an unknown claim |
| Settings | done | `syncEngine` dropped by the page's post, and the six keys spelled out five times |
| Dead code | done | 10 HResult constants, 15 Win32 declarations, `VbeCommands.Describe`, the `setTheme` and `applyEdit` paths, `module/didClose`, the stale README tables, a form feed that made a documented command uncopyable, the missing `TwinFixture` generator |
| Gate blind spots | done | the driver check counting unrun suites and `console.log`, four suites wired in, engine typecheck before packaging, engine currency on the `-SkipGate` release path, freshness and the two cheap checks in CI |
| API coverage | done | 18 items. Closed: scoped search and Replace All (A6-ish, prior), Properties pane (A5), status bar readouts (A15), the window gestures (A12: `frame` close/show, `pane closeNative`), the references dialog's open form (A16), the sync dialog's read/drive pair (A6), the immediate/history doc-truth claims (A17), and the session teardown/revival (A10, a Debug-only `session` route driving the real OnBeginShutdown, `session-lifecycle.mjs` proving the revival in -Deep). A18 (populating the Watch panel) is resolved as a designated exception rather than built: its only mechanism is synthetic keyboard input into the native dialog, which the standing no-synthetic-input rule forbids, so `watchHas` is documented as having no api driver by design with `Test-WatchPanel.ps1` its out-of-gate exception - the audit's own fallback |
| Performance | open | 20 items; `dev.ps1` repackaging the engine every run is done, the rest are not |
| Duplication | open | 13 items |

Two findings were NOT implemented as written, and both are recorded here rather than quietly
dropped:

- **"Delete rename-features.mjs, three-copies.mjs covers it"** is wrong, and the adversary should
  not have confirmed it. `three-copies.mjs` asks one question - do the workbook, the surface and
  the analyzer hold the same text - across eight operations, one of which happens to be a rename of
  its own scratch symbol. `rename-features.mjs` asks whether the rename touched the right things
  and left the wrong ones alone: the qualified call follows, the ambiguous bare call does not,
  `Rival`'s own `Recalculate` is untouched, `HelpersExtra` is untouched despite the shared prefix.
  A rename that wrongly rewrote `Rival` would leave all three copies in agreement and pass. The
  suite was wired into the gate instead.
- **`trip` has no driver in the gate.** Its only caller is `perf-scaling.mjs`, which needs a fourth
  fixture and exists to be read rather than to go green. It is excused by name in
  `audit-routes.mjs` with that reason, which is the pattern already there for `drainfinalizers`.

## Addendum, 2026-08-12: the orphaned suites, and a correction to the row above

The "Gate blind spots: done" row was overstated, and the overstatement is worth recording: four
suites were wired into the live gate on 2026-08-11 and the live gate was never then run to green.
It could not have been - four wired suites printed "N checks, M broken" where the runner parses
"N passed, M failed", so every -Live run died at the first of them. Running it for real on
2026-08-12 surfaced, one layer at a time: the verdict mismatch (menu-bar, debugger-features,
step-into-features, rename-features); Test-DebugApi's reader check passing only against its own
saved residue, and the save that put that residue into DebugFixture.xlsm on disk; module-sync
leaving a ByTheDialog sub in the fixture's Helpers, which one run gets away with and the gate's
two planner runs stack into "Ambiguous name detected"; write-rollback wired mid-group against its
own header's stated terms, wedging every suite behind it; and a vacuous "0 passed, 0 failed"
being read as green. All fixed; lessons 67 and 68. The -Live gate passes 17 steps.

The 14 suites nothing ran were then triaged by eight readers (cost grounded line by line) and
placed:

- **Headless gate**: Test-CloseConfirm.ps1, whose engine leg (engine-live-probe.mjs) is the only
  didChange coverage over the engine pipe anywhere; its page leg moved out of the probes list
  rather than launching Edge twice. Net headless cost under a second.
- **-Live, sharing already-open sessions**: objbrowser-live-probe.mjs (the palette's only live
  coverage; self-resolves both doors from the discovery file, targets Runner, 9 checks),
  Test-ResizeFollow.ps1 (the only thing that resizes the host window; pane closes moved from
  Application.VBE to the pane route, so no project trust), Test-CloseVbe.ps1 (lesson 27's crash
  class; reopen via ExecuteMso, runs as the live half's final step by construction).
- **-Deep, the new pre-release tier**: language-features.mjs on LanguageFixture (tolerating the
  two xlide_vscode#11 upstream defects by name), language-live-probe.mjs and surface-walk.mjs
  --steps 80 sharing one RenameFixture + TwinFixture session. The walk now fails on vacuity,
  exits nonzero, and its replay hint names this file rather than a file that does not exist.
- **Kept manual**: perf-scaling.mjs (asserts nothing; the instrument is for reading).
- **Deleted**: Test-Language.ps1 (a zero-check wrapper around language-live-probe.mjs).
- **Blocked on routes, hand-run until then**: Test-CloseHiddenPane.ps1, Test-GhostLocalsPanel.ps1,
  Test-WatchPanel.ps1 and Test-ObjectBrowser.ps1's lifecycle trio all need the VBA project object
  model, and the machine runs with that trust OFF. Their path into a tier is A12-shaped routes
  (frame visibility, native pane close) and something that can populate a watch.

  Later the same day, two of those routes arrived: `pane?action=closeNative` (the host-originated
  close of a hidden pane, through the editor's own pane list) and `frame?action=close|show` (the
  developer's X click posted through the pump, and the window brought back), with
  `palette?action=hide` beside them and `state.frameVisible` to observe by. `window-routes.mjs`
  drives all three in the -Live gate's DebugFixture group, holding the follow contracts too - the
  palette going down with the frame, staying away on its return, the hidden pane's tab leaving
  the strip with the shown module unstolen. Test-CloseHiddenPane.ps1 is deleted, its whole
  subject now api-driven and gate-run with trust off. Still waiting on routes: a way to populate
  a watch (Test-WatchPanel), the setLocals push assertion (Test-GhostLocalsPanel), and
  Test-ObjectBrowser's icon and full-restart residue.

The two-workbook hole is closed at both tiers: the -Live gate's second group now opens
RenameFixture AND TwinFixture in one session (zero extra launches; all six suites proven green in
the double session before the widening), with rename-boundary.mjs new and last in the group - a
rename must cross modules and stop at the workbook, byte for byte, through the rename and its
undo. The randomized cross-workbook walk runs in -Deep.

## How it was produced, and what that is worth

Twelve read-only agents surveyed the surface in parallel, one per area. Every finding was then
handed to a separate adversary agent that reopened the cited lines and tried to refute it, with
instructions to default to refuted when the claim could not be established from the code. Three
dimension synthesizers deduplicated and ranked what survived.

| | |
| --- | --- |
| Findings raised | 123 |
| Refuted by the adversary | 3 |
| Never judged | 2 |
| **Standing** | **118** |
| Of those, verified / supported | 108 / 10 |
| Severity high / medium / low | 3 / 44 / 71 |
| Effort small / medium / large | 82 / 34 / 2 |

Two caveats that belong at the top rather than the bottom. Confidence is the finder's own, carried
through the adversary pass: `verified` means every line the claim rests on was read, `supported`
means the definition was read but not all of its uses. And no finding here was checked against a
running Excel, because none was running; anything whose truth depends on live behaviour says so in
its own entry.

---

## A. API coverage

Gaps that stop a script from driving or observing Excel first-class through the debug api, and
defects in the api that already exists.

The debug api covers the keyboard and the shim well and covers clicks, dialogs and second surfaces
badly. Three shapes recur. First, actions that can be driven answer a constant: five routes wrap a
`private void` session method and synthesise true, so `command`, `compile`, `pane action=close`,
`undoRename`, `caret` and `breakpoint` all report the request rather than the outcome, and every
decline path in the callee reaches only the log. Second, observers read the shim's own record and
are presented as the user's view: `ui.search.matches` is structurally 0 for any non-module scope,
`native` compares the host against the shim's reconstruction and never asks the page, and `assert
that=shownModule` compares our record with our record. Third, whole click-driven surfaces have no
vocabulary at all: the toolbar strip, the Properties pane, the import/export dialog, the Panes menu,
the tab context menu and the entire Object Browser palette are reachable only by hand-rolled
querySelector through `eval`, or in six harness scripts by Application.VBE, which needs the
AccessVBOM trust setting the product is supposed to be free of. audit-routes.mjs cannot see any of
it: it asks whether every route is documented and driven, never whether the harness went around the
api, and its driven-check for the `log` route is satisfied by 203 console.log lines.

### The ranked list

| # | Finding | Effort | Risk | Confidence |
| --- | --- | --- | --- | --- |
| 1 | Scoped search has no driver and its observer reports the module-scope fields, so Replace All is untested and act search scope=project answers did:true having searched nothing | medium | medium | verified |
| 2 | ExecuteEditorCommand discards the bool it computes, so command answers ran:true for a greyed menu item and compile answers compiled:true when the compile never started | small | low | verified |
| 3 | No act presses a toolbar button or runs an editor action by id, so three harness sites click the strip by selector and the whole command surface behind it is undrivable | small | low | verified |
| 4 | component?action=rename skips AdoptRename, so a scripted rename strands the breakpoint record and the write baseline under the old module name | small | low | verified |
| 5 | The Properties pane writes real component state and has no route, no act, no snapshot field and no suite - the only user-visible surface with zero api presence in either direction | medium | low | verified |
| 6 | The import/export dialog has no act and no ui field, so the gate's module-sync suite drives it with eleven eval scripts over seven private CSS selectors, and has already gone red for that reason | large | medium | verified |
| 7 | breakpoint and caret answer the state they were asked for, and GoTo's silent failure lets a breakpoint land on the previously shown module | small | medium | verified |
| 8 | pane action=close and undoRename answer a constant true over callees whose failures reach only the page and the log | small | low | verified |
| 9 | ui and act are hard-wired to the editor surface, so the Object Browser palette - the second-largest page in the product - has no snapshot, no action vocabulary and a probe that hand-rolls DevTools | medium | low | verified |
| 10 | Session teardown and the cancelled-shutdown revival - a documented field failure that left the add-in dead inside a living Excel - cannot be driven or observed at all | medium | medium | verified |
| 11 | Nothing in the gate notices when the harness goes around the api, and the one driven-check that exists is satisfied for log by 203 console.log lines | small | low | verified |
| 12 | No route hides the editor frame or closes a native code pane, so six harness scripts still reach Application.VBE and need the AccessVBOM trust setting the product forbids | small | medium | verified |
| 13 | The parity instrument compares the host against the shim's reconstruction of the page, and the detector that watches that hop reaches only the log | medium | low | verified |
| 14 | Panes cannot be opened or closed through the api and ui.panes omits which side a pane is on, so a gate probe drives the Panes menu by selector | small | low | verified |
| 15 | The status bar - caret readout, module name and the notice line - is in no snapshot, so a whole class of declined actions cannot be shown to have been reported | small | low | verified |
| 16 | act references answers the lookup and never opens the dialog whose rendering is the feature, and the act appears in neither api doc | small | low | verified |
| 17 | The reference document makes five claims the code does not honour, two of them repeated in code comments where the next maintainer will read them | small | low | verified |
| 18 | The Watch panel can be read and asserted but never populated, so its only live coverage spawns a PowerShell helper that types into the native dialog with WM_CHAR | medium | medium | verified |
| 19 | build-fixture's write-retry can no longer run for the refusal it was written to absorb, because the write route learned to report it and the client turns that into a throw | small | low | supported |
| 20 | Two orphaned probes still hold VBProject call sites for work that routes now do | small | low | verified |
| 21 | The client cannot send keep on command or selector/pad on capture, so no .mjs suite can open a dialog deliberately or take a cropped screenshot | small | low | verified |
| 22 | history drops five routes from a transcript documented as every request, so the replay script it generates loses exactly the waits that made the session work | small | low | verified |
| 23 | assert polls a misspelt claim name for the full timeout and answers it in the same shape as a real failure | small | low | verified |

#### A1. Scoped search has no driver and its observer reports the module-scope fields, so Replace All is untested and act search scope=project answers did:true having searched nothing

**Where.** ui/editor/src/devsurface.ts:943-962 (the whole search act); ui/editor/src/devsurface.ts:120-125
(DevSurfaceParts.search); ui/editor/src/searchwidget.ts:502-503, 319, 557-561 (find dispatches
`input`, whose only handler searches when scope is module); ui/editor/src/searchwidget.ts:563-576
(scopeChanged empties matches/current for every other scope); ui/editor/src/searchwidget.ts:466-477
(state()); ui/editor/src/searchwidget.ts:322-333 (replace, replaceAll, findAll, prev/next are click
listeners only); ui/editor/src/searchwidget.ts:195-219 (showSearchResults never touches
this.matches); docs/debug-api.md:69; docs/driving-excel.md:300,306

**What.** The search act is open/find/close and nothing else. find() sets the input and fires `input`;
queryChanged() runs a search only when scope() === "module". open() calls scopeChanged(), which for
project or workbook scope does `this.matches = []; this.current = -1; this.counter.textContent =
""`. So act("search", {query, scope:"project"}) types a query, runs no search, and returns did:true,
while ui.search.matches is 0 and current is -1 by construction whatever the panel drew. The scoped
result the panel actually renders is built by showSearchResults from its own arguments and never
lands in a field the snapshot reads. Replace, Replace All, Find All, next and previous hang off
click listeners with no method exposed on DevSurfaceParts, so none can be driven; Replace All
rewrites text across every module of a project (EditorSurface.cs:1904 raises ReplaceAllRequested,
AddInSession.cs:1104 wires it to OnReplaceAllRequested at 2742) and has no driver, no route and no
suite.

**Why it matters.** This is the only place where the driver and the observer fail in the same direction, so a
project-scope search test written today passes green forever: the action does nothing and the field
it would be checked against is unconditionally 0. Behind that trap sits the most destructive
operation on the surface, a cross-module text rewrite, with zero automated coverage of any kind.

**Fix.** Widen the existing act rather than adding routes: accept
run=find|findAll|next|previous|replace|replaceAll plus a replacement argument, and expose runScoped,
showModuleResults, next, previous, replaceCurrent and replaceAllRun on DevSurfaceParts.search so
each goes through the method the button's own click handler calls. Add the scoped result to
UiSnapshot["search"] as NEW fields (the count handed to showSearchResults, truncated, replaced, and
the grouped rows), read from those arguments rather than from the .search-row DOM, so no existing
field changes meaning. Update the act row in docs/debug-api.md, the act list in
docs/driving-excel.md and the client helper.

**Size.** 6 widget methods to expose, roughly 30 lines across devsurface.ts and searchwidget.ts, plus one added record on the search snapshot. Effort medium, risk medium, confidence verified.

#### A2. ExecuteEditorCommand discards the bool it computes, so command answers ran:true for a greyed menu item and compile answers compiled:true when the compile never started

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1722 (private void), :1756 (`var ran =
VbeCommands.Execute(_editor, command);`), :1762-1780 (ran used only for three Notify blocks);
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3197-3199 (command route); :1642-1643 and
:1670-1671 (compile route); src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:709-711 (the field is
named `ran`); src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:97-131 (four false returns, including
`command: {id} is currently disabled`)

**What.** One cause, two lying routes. VbeCommands.Execute returns false when there are no CommandBars, when
the control is absent, when the control is disabled, and on exception; ExecuteEditorCommand captures
that bool and uses it only to raise three Notify messages, then returns void. The command route
calls it and returns DebugCommandReply(true, command) unconditionally, so POST command?name=stepOver
outside break mode logs "currently disabled" and answers {"ran":true}. The compile route calls the
same method through RunOnHostThread and derives its verdict solely from dialogs seen inside a fixed
watch loop with no early exit, so "the project compiled", "the Compile item was greyed and nothing
ran" and "the error box arrived after waitMs" all answer compiled:true, and every call burns the
full default 6000ms.

**Why it matters.** command is the harness's only way to drive Run, Step, Reset, Compile, Save and the watch commands.
compile is the precondition gate for the whole debugger suite, guarded by the comment "It has to
compile, or every command below is really a test of the dialog guard" (debugger-features.mjs:74) and
read as compiled.compiled !== false - a gate that passes when the compile never happened turns the
suite it protects into exactly what the comment forbids. For command the cost is a misattributed
diagnosis rather than a silent pass, since debugger-features.mjs never reads .ran.

**Fix.** Return the bool (or a small record carrying the decline reason already built for Notify) out of
ExecuteEditorCommand and put it in `ran`; the two early returns at AddInSession.cs:1732
(ObjectBrowser) and :1748 (ToggleBreakpoint) run before `ran` exists and must answer true
explicitly. Add a `detail` field distinguishing "not present" from "disabled". Add `started` to
DebugCompileReply so compiled:true, started:false is expressible, and either stop the watch loop
early once the project's Mode settles or state in the route table that compiled means "no dialog
appeared within waitMs". No route or field is renamed; ran starts meaning what its name says. 7 call
sites of ExecuteEditorCommand, 3 of them in the debug api.

**Size.** one bool plumbed out of a void method, 7 call sites, one new field on each of two reply records. Effort small, risk low, confidence verified.

#### A3. No act presses a toolbar button or runs an editor action by id, so three harness sites click the strip by selector and the whole command surface behind it is undrivable

**Where.** ui/editor/src/toolbar.ts:35-92 (COMMANDS, about thirty entries: objectBrowser :85, openSync :87,
openPanes :88, openHelp :90, plus the indent/fold/gotoLine/quickCommand cluster :68-83);
ui/editor/src/toolbar.ts:165 (`button.addEventListener("click", () => run(command))`);
ui/editor/src/main.ts:468-508 (the shell handler); tools/harness/module-sync.mjs:248;
tools/harness/objbrowser-live-probe.mjs:105; tools/harness/menu-bar.mjs:157;
ui/editor/src/main.ts:1120-1131 (the Undo Rename editor action, registered and never run by
anything)

**What.** The act vocabulary has 34 actions and none takes a command id. `press` is a keyboard key, and the
shim's `command` route is the NATIVE editor's command by name, not a page toolbar id. So the only
way to reach openSync, openPanes, openHelp or objectBrowser is to find the button by its
data-command attribute and click it, which is what module-sync.mjs and objbrowser-live-probe.mjs do
and what menu-bar.mjs scrapes to assert the strip's contents. The same hole covers editor actions
registered by id: xlide.undoRename's run() body, its context-menu placement and its notice are never
executed by any test, only the shim function underneath it is (the undoRename route calls UndoRename
directly and its bypass is undocumented at docs/debug-api.md:91 and docs/driving-excel.md:210).

**Why it matters.** This is the highest-leverage single act in the list. It is the driver half of the sync dialog, the
Panes menu, the Object Browser summons and the Undo Rename menu entry, all of which appear
separately below; one action retires the selector-clicking in three harness files and unblocks four
other gaps. A test that clicks a button it found itself keeps passing after the button is renamed,
removed, disabled by the needsBreak gate or scrolled out of the strip, because it never went through
buildToolbar's dispatch.

**Fix.** Add an act taking a command id, looking it up in COMMANDS and calling the same run(command) callback
buildToolbar hands the button, refusing a command that is absent or disabled and saying which.
Generalise it (or add a sibling) to run a registered editor action by id via
editor.getAction(id).run(), the way the format act already calls getAction, which gives
xlide.undoRename a driver. Document in both api docs with a client method, and designate the
undoRename route as a page bypass in its two doc rows.

**Size.** about 25 lines in devsurface.ts plus a shell accessor for the command callback; 3 harness selector sites retired. Effort small, risk low, confidence verified.

#### A4. component?action=rename skips AdoptRename, so a scripted rename strands the breakpoint record and the write baseline under the old module name

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2852-2867 (the route: FindComponent,
SetString("Name"), Log.Info, ComponentsChanged, reply);
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3344-3385 (AdoptRename re-keys _writtenModules and
_breakpoints, moves _propertiesTarget, reloads the shown module, then calls ComponentsChanged
itself); :3230-3240 (the Properties grid "(Name)" row, the UI equivalent, does call it); :6841-6845
(RemoveComponent already learned this for removal); docs/debug-api.md:85 (the designation `remove`
got and `rename` did not)

**What.** The route sets the COM Name and calls ComponentsChanged, which is only PublishModules +
PublishProjects + Reanalyse and cannot re-key anything. The UI path for the same operation goes
through AdoptRename. After a scripted rename the session's _breakpoints entry still carries the old
key and the old Module value, so GET breakpoints reports a breakpoint on a module that no longer
exists and none on the module that now carries it; the write baseline is filed under the old key
too, so ModuleDiffersFromSaved answers false for the renamed module and the tab close gate silently
stops offering Save/Don't Save.

**Why it matters.** This is design rule (a) failing in exactly the shape `remove` was already fixed for, still live in
its sibling in the same route: a fixture built with this action leaves a machine no developer can
produce, and the divergence reads as a product bug. It is latent today only because the sole caller
(com-leak.mjs:301-303) renames and immediately renames back.

**Fix.** Call AdoptRename after the SetString instead of ComponentsChanged (AdoptRename ends by calling it).
Not a bare one-line swap: AdoptRename re-keys against DisplayFromProjectId(_shownProject) at
AddInSession.cs:3349 and :3360 while the route resolves its target from componentOwner
(DebugApi.cs:2718), so the owner must be threaded in the way RemoveComponent computes it at
AddInSession.cs:6822, or a rename in a non-shown workbook migrates keys under the wrong one. If the
bare-COM primitive is wanted deliberately, keep it behind an explicit argument and designate the
deviation in the three places `remove` designates its: the docs/debug-api.md row, the xlide-api.mjs
method comment, and the code. Worth checking with the fix: whether EditorSurface's per-module
document table re-keys on rename (RemoveComponent additionally calls DiscardEdits at
AddInSession.cs:6825).

**Size.** two dictionaries left unmigrated per rename; one route action, one owner argument threaded into AdoptRename, one docs row. Effort small, risk low, confidence verified.

#### A5. The Properties pane writes real component state and has no route, no act, no snapshot field and no suite - the only user-visible surface with zero api presence in either direction

**Where.** ui/editor/src/shell.ts:492 and :539 (the select and the text input both call handlers.editProperty);
ui/editor/src/main.ts:522-523 and ui/editor/src/bridge.ts:647-653 (the wiring);
src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1943-1953 (the page message);
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1114 (PropertyEditRequested) and :3216-3259
(OnPropertyEdit, including the "(Name)" branch that calls AdoptRename at :3239 and WriteProperty's
complaint at :3251-3259); :3074-3130 (PublishProperties computes the rows, the kind and the
allowed-name list); ui/editor/src/devsurface.ts:38-93 (no properties block on UiSnapshot)

**What.** A repo-wide grep for editProperty, selectComponent and propertiesComponent across ui, tools, docs,
src, installer and .github returns nothing outside that wiring chain. There is no case in the route
switch, no entry in the act table, no client method, no doc row and no suite. The read side is
equally absent: PublishProperties computes the grid and sends it as SurfacePropertyEntry[], and no
route reads it. The only way to exercise any of it is to synthesise change events on shell.ts's
controls through eval.

**Why it matters.** This is a whole pane, one of six seats, whose edits change component state on the host, and the
"(Name)" row is the UI entry point for AdoptRename - the exact bookkeeping the rename route above is
missing, so today neither the route nor the UI path has a test. A refused write returns a complaint
from WriteProperty that nothing can read, and ClassifyDocument gates which rows a document component
may edit with nothing checking it.

**Fix.** Add a `properties` route. GET (optional component) answers the rows PublishProperties computed -
component, kind, and each entry's name, value and editability - so the observation reads the same
fields the panel renders. POST with component, name and value calls OnPropertyEdit, the exact entry
point the page's message reaches, and returns the complaint when the write is refused. Add a
matching act calling handlers.editProperty for the page-side half, a properties block on UiSnapshot,
both doc rows, and a properties()/setProperty() pair in xlide-api.mjs.

**Size.** one route with two verbs, one act, one snapshot block, roughly 40 lines of page code plus the shim route. Effort medium, risk low, confidence verified.

#### A6. The import/export dialog has no act and no ui field, so the gate's module-sync suite drives it with eleven eval scripts over seven private CSS selectors, and has already gone red for that reason

**Where.** tools/harness/module-sync.mjs:20-22 (SYNC_ROWS scrapes .sync-item/.sync-item-name/.sync-chip), :248,
250, 255, 270, 271, 273, 274, 276, 280, 289, 311 (eleven api.eval calls), :258-268 (the comment
recording the 2026-08-10 red-on-one-planner false failure); ui/editor/src/syncdialog.ts:90, 105,
121, 164, 365, 471, 510 (the selectors), :375-380 and :707-711 (Apply composes its own request),
:481-492 (per-row ticks); ui/editor/src/main.ts:478-483 (opened only from the toolbar branch);
ui/editor/src/devsurface.ts:182-192 (dialogsUp yields only {id, title}); tools/verify.ps1:422 (runs
the suite twice, once per planner)

**What.** The suite that exists to prove the dialog and the sync route leave the same state builds the UI half
of that comparison by scraping the render: it opens the dialog by clicking
button[data-command=openSync], sets #sync-folder.value and dispatches a synthetic change, clicks
.sync-direction[data-direction=import], finds Apply by textContent and closes with #sync-close. No
act names the dialog and the ui snapshot reports only that some aria-modal element is up, so the
plan the dialog drew is invisible to any typed answer.

**Why it matters.** The sync route is honestly the same HandleSync call the dialog reaches, so everything the route
proves is about the planner and nothing is about the dialog: which rows it ticked, whether it redrew
for the new direction, what its status line said. That half rests on a scrape that no compiler and
no gate step protects, that cannot tell a stale render from a correct one, and that has already
produced one false red - in a suite that runs twice per gate.

**Fix.** Give the dialog the treatment the tree got: acts for open/close, folder, direction, mode, ticking a
row by id and pressing Apply, each calling the same plain click and change handlers the dialog's own
listeners call (syncdialog.ts:140, 169, 375, 381, 753 - no pointer sequence is involved). Report the
drawn plan as a new ui.sync block - {open, folder, direction, mode, rows:[{name, status, selected}]}
- taken from the item array the row builder at syncdialog.ts:471 renders. The open half comes free
if the toolbar act (rank 3) lands first.

**Size.** a 768-line dialog with roughly 6 controls to expose; retires 11 eval calls and 7 private selectors from a suite that runs twice per gate. Effort large, risk medium, confidence verified.

#### A7. breakpoint and caret answer the state they were asked for, and GoTo's silent failure lets a breakpoint land on the previously shown module

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3217-3238 (GoTo, then alreadySet from
_editorSurface.Module, then ToggleBreakpoint, then DebugCommandReply(shouldSet, ...)); :3298-3315
(caret returns DebugCommandReply(true, 0) unconditionally);
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1354-1364 (GoTo is private void with a silent early return
when the pane is null) and :1404-1407 (swallowing catch); :2929-2935 (ToggleBreakpoint acts on
_editorSurface.Module) and :2943-2951, :2963-2967 (two more silent refusals);
tools/harness/debugger-features.mjs:84-91 (a 1200ms sleep plus a second breakpoints round trip to
find out what happened)

**What.** The route navigates, then reads the shown module, then toggles. If GoTo cannot resolve the component
- a misspelt module name, or a name that exists only in another workbook when `project` is wrong -
it logs and returns, _editorSurface.Module is still the previous module, and ToggleBreakpoint sets a
breakpoint there. The reply carries shouldSet in the field named `ran`, so state=off on a clean line
answers ran:false on success and the field cannot be used as a success flag in either direction.
Note the pane-not-open case is NOT the failure: FindCodePane creates a pane that was never opened
(AddInSession.cs:7023-7025).

**Why it matters.** Setting a breakpoint on an unresolvable module silently sets one somewhere else and answers ok,
which is the class of defect the route's own state=on|off argument was added to prevent. Because ran
carries the requested state, every breakpoint assertion in the harness costs a fixed sleep plus a
second round trip.

**Fix.** Make GoTo return a complaint string the way WriteModule does, and have both caret and breakpoint
answer it as `error`. Add `set` (the observed state, re-read from BreakpointsFor after the toggle)
and `module` (the module actually acted on) to the breakpoint reply so ran stops carrying two
meanings and one call answers what two calls answer now.

**Size.** one return type on GoTo, 2 routes, 2 new fields; removes a 1200ms sleep and a round trip per breakpoint assertion. Effort small, risk medium, confidence verified.

#### A8. pane action=close and undoRename answer a constant true over callees whose failures reach only the page and the log

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2919-2928 (close: OnModuleCloseRequested then
DebugCommandReply(true, 0)) against :2909-2918 (the `open` branch twenty lines above, whose comment
records the same lie being fixed and which now returns showed's error);
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:6472 (OnModuleCloseRequested is private void), :6484-6491
(failed save returns), :6544-6550 (unanswered confirm returns); DebugApi.cs:2938-2945 (undoRename:
UndoRename(0) then true); AddInSession.cs:5030 (private void), :5038-5042 ("There is no rename to
undo" goes to the page for the request id the route invented), :5058-5062 and :5066-5072 (a partial
revert sets `stopped` and breaks); tools/harness/rename-features.mjs:115-118 (prints the reply,
never checks it, sleeps 3000) against :82-84 (the forward half checks did/detail);
docs/debug-api.md:89

**What.** Two routes, one shape. Close has two paths that do not close - a workbook that would not save, and a
dirty module that raised ConfirmClose instead - and both answer ok, so a caller cannot tell "the tab
closed" from "a Save box is standing and the tab is still there". (The
discard-whose-revert-was-refused path at AddInSession.cs:6506-6512 does fall through to CloseModule,
so the tab closes; that is a separate defect, a silent failed revert.) UndoRename cannot express
"there was nothing to undo" or "it stopped half way", so the worst outcome, a partially reverted
multi-module rename, is indistinguishable from a clean one.

**Why it matters.** Close is the interesting half of the pane route, the one with the save/discard/cancel gate the tab's
X uses, and a probe testing tab-strip behaviour after a close is testing a state it has not
established. Rename reversibility is the property the undoRename route exists to let a probe assert,
and the two halves of that one feature answer in two different shapes.

**Fix.** Give OnModuleCloseRequested a string? complaint return like WriteModule and RemoveComponent already
have and answer it, reporting the ConfirmClose case as a distinct non-error outcome (closed:false,
awaiting:"confirm") so a caller can drive act("answerCloseConfirm") next. Give undoRename the shape
the page gets: modules restored, the restored name, and the `stopped` reason; act("rename")'s {did,
detail} is the precedent.

**Size.** two return types, five existing early-return sites; removes a 3000ms sleep and two read-backs from rename-features.mjs. Effort small, risk low, confidence verified.

#### A9. ui and act are hard-wired to the editor surface, so the Object Browser palette - the second-largest page in the product - has no snapshot, no action vocabulary and a probe that hand-rolls DevTools

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1054 (ui passes null) and :1080-1083 (act passes
null); :220 (RunPageScriptOnce already resolves surface == "palette" to _browserPalette); :1101 and
:1119 (eval and await do read `surface`; capture takes a window); ui/editor/src/main.ts:1242-1244
(?view=objbrowser boots before boot, so installDevSurface at devsurface.ts:1462 never runs on that
page); tools/harness/objbrowser-live-probe.mjs:24-76 (hand-rolled connect/attachToPage/evaluate over
a DevTools WebSocket), :106, 133, 139, 145 (all four interactions are synthesised DOM events);
tools/harness/Test-DebugApi.ps1:484 (eval?surface=palette already works live)

**What.** Two routes pass a hardcoded null surface even though the plumbing beneath them already branches on
it, and the palette page has no xlideUi installed to answer if they did. The palette is a 568-line
surface with its own search box, tree, member list, detail pane, Escape handling and arrow-key
resizing, and every assertion about it is a document.querySelector scrape - the practice
devsurface.ts was written to end. The live probe reaches it through a second door it opens itself,
sixty lines of DevTools plumbing, although the eval route already targets that surface and every
interaction the probe drives is a synthesised DOM event eval can deliver.

**Why it matters.** A scrape cannot distinguish a stale render from the state, and the probe's hand-rolled transport has
none of the client's error-field check, so a route failing quietly reads as the page misbehaving.
Note the CDP transport is not why the probe is outside the gate - Test-ObjectBrowser.ps1 throws when
Application.VBE is null, which is the harder blocker.

**Fix.** Honour the existing `surface` argument on ui and act (default unchanged, so no route shape changes:
this is a new argument, not a rename) and install a small xlideUi on the palette page with its own
state() and act() covering the search box, the libraries picker, the selected node and the detail
pane. Then replace the probe's CDP transport with api.eval(script, "palette"), which deletes
connect/attachToPage/evaluate and the --cdp argument and removes the need to export the client's
private clientFor; a straight port needs an explicit waitMs, since the eval route defaults to 5000ms
while the probe's inner polls run to 10000ms.

**Size.** two argument reads in the shim plus a palette-side devsurface; about 60 lines of DevTools plumbing retired from the only live palette probe. Effort medium, risk low, confidence verified.

#### A10. Session teardown and the cancelled-shutdown revival - a documented field failure that left the add-in dead inside a living Excel - cannot be driven or observed at all

**Where.** src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs:116-140 (OnBeginShutdown calls _session.Stop() then arms the
watchdog), :151-230 (OnWatchdogTick revives from retained pointers behind a two-consecutive-tick
guard), :250-264 (the 2026-08-02 note); src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:7605-7606 (Stop
disposes the DebugServer first), :7632 (SetNativeChromeBands visible), :7644-7645
(RestoreLocalsPalette/RestoreWatchPalette), :7654-7661;
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs (no session/stop/shutdown/disconnect/revive case
in the route switch), :1833 (ComWrappersLive in stats); tools/harness (no reference to shutdown,
revival or the watchdog anywhere)

**What.** Stop() is the only place the native menu and toolbar bands, the ghosted Locals and Watches palettes
and the host chrome are put back, and OnWatchdogTick is the guard against the observed failure where
a developer cancelled Excel's save prompt and the editor came back with the add-in dead. Neither is
reachable from any route and neither is exercised by anything in the harness; the nearest coverage,
Test-CloseVbe.ps1, closes the VBE frame with WM_SYSCOMMAND and never reaches OnBeginShutdown.

**Why it matters.** Two release-blocking properties with no test. A regression that leaves the editor menu-less after
unload is invisible until a user hits it, and the revival path can only be reached today by closing
Excel by hand and pressing Cancel. The COM wrapper counters that make the leak gate meaningful are
never checked across a teardown, which is where a leak would actually show.

**Fix.** Add a Debug-only lifecycle route (session?action=beginShutdown, session?action=disconnect) calling
the same XlideAddIn entry points the host calls through a static reference set in OnConnection, not
a private teardown of its own - it is OnBeginShutdown that arms the watchdog, and the revival is the
thing under test. One mechanical constraint: Stop() disposes the DebugServer that is serving the
in-flight request, so the teardown must be posted to the host thread after the response is flushed
and the route documented as answer-less, or every suite using it fails on a transport error rather
than an assertion. The client reconnects by re-reading debug-api-{pid}.json, which the revived
session rewrites.

**Size.** one route with two actions plus a static hook in XlideAddIn; unlocks three assertions that are impossible today. Effort medium, risk medium, confidence verified.

#### A11. Nothing in the gate notices when the harness goes around the api, and the one driven-check that exists is satisfied for log by 203 console.log lines

**Where.** tools/harness/audit-routes.mjs:143-146 (isDriven's bare `\.method\s*\(` match), :100-121
(methodsByRoute walks back to `log:` at xlide-api.mjs:421-422), :137-140 (the corpus is already
every harness .mjs and .ps1), :9-14 (the header stating that documented and reachable is not
covered), :28-37 and :162-164 (NOT_DRIVEN_ON_PURPOSE, one entry, with the symmetric check that makes
it honest)

**What.** Two defects in the gate's own coverage guarantee. First, the driven-check for route log is
/\.log\s*\(/ against a corpus containing 203 console.log lines and 5 real api.log calls, so the
check would stay green if every real call were deleted - and log?waitMs is the mechanism the
harness's whole no-sleep policy rests on. Second, the audit asks three questions of each route and
no question of the harness, so a suite reaching Excel through COM or DevTools, or scraping a value
the render already computed, passes twelve green steps in silence.

**Why it matters.** Every gap in this report is invisible to the gate. That is how Test-GhostLocalsPanel and
Get-EditorScreenshot stayed in the tree after their routes arrived, and how a gate suite came to
drive the sync dialog with eleven eval calls. The audit's own exemption design - a fallback must be
named with a reason, and an excuse that turns out to be unnecessary fails the build - is the fix
already proven in this file.

**Fix.** Anchor the method match to a receiver, /\b(?:api|client)\w*\.method\s*\(/, and print the file each
route was found driven in so a vacuous match is visible in the ok line. Add a second pass over the
corpus the audit already reads, flagging GetActiveObject, Application.VBE, VBProject,
AccessibleObjectFromWindow and raw DevTools use (new WebSocket, /json/list, Runtime.evaluate), each
allowed only by an OFF_THE_API list naming the file and the reason, in the same shape as
NOT_DRIVEN_ON_PURPOSE. Do NOT extend it to api.eval: there are roughly 60 eval calls in the harness
and the eval route's own contract invites one-off use, so that list would start with dozens of
entries and never shrink, which is the failure mode the drainfinalizers design avoids. The COM axis
seeds with exactly six files.

**Size.** one regex plus a second pass with a six-entry seed list; 203 false matches on one route today. Effort small, risk low, confidence verified.

#### A12. No route hides the editor frame or closes a native code pane, so six harness scripts still reach Application.VBE and need the AccessVBOM trust setting the product forbids

**Where.** tools/harness/Test-ObjectBrowser.ps1:170-177 (the comment stating the dependency) and :179, 208, 213
($excel.VBE.MainWindow.Visible); tools/harness/Test-CloseVbe.ps1:36, 53;
tools/harness/Test-CloseHiddenPane.ps1:14-15, :33-40 ($pane.Window.Close(), with the comment
"exactly as the host's close path does"), :47 (concludes from a setModules log grep);
tools/harness/Test-ResizeFollow.ps1:59-60, :103-105; tools/harness/Test-GhostLocalsPanel.ps1:8;
tools/harness/Test-WatchPanel.ps1:60; src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2610-2628
(the windows route reads Type, Caption and Visible with no setter) and :2920-2922 (pane close goes
"through the same gate the tab's own X uses"); docs/driving-excel.md:1155-1163 (the trust table
names only two of the six)

**What.** Two host-window capabilities the api never grew. Nothing can set VBE.MainWindow.Visible, which the
palette lifecycle test needs because Excel's ribbon command opens the editor but does not close it;
and nothing can close a code pane from the host side, which is the direction two probes are written
to test and which is a different event from the tab's X. The add-in is handed the VBE object at
OnConnection and needs no trust setting to do either, so these dependencies exist only because the
capability was never given a door. The trust table in the operational guide understates the problem
by four scripts.

**Why it matters.** A probe that cannot run with AccessVBOM off is a probe the developer skips, so the palette's whole
lifecycle story and the hidden-pane close - a real host event with a real fix behind it, where the
pane tracker kept a dead tab - have no coverage anyone runs, and what coverage exists concludes from
a shim-log grep rather than the tab strip.

**Fix.** Add frame?visible=0|1 answering the shape the windows route reports, driving VBE.MainWindow.Visible
on the host thread, and pane?action=closeNative&module=... calling Close on the code pane's Window
from inside. Both are deliberate deviations from api rule (a) - no UI action hides the VBE main
window without closing it, and closeNative is the host's close rather than the tab's, which is the
point - so both must be designated in the route table, the client method and the code, exactly as
`remove` designates its. The probes then assert through ui().workspace tabs. Correct the trust table
to name all six scripts.

**Size.** two routes; 6 harness scripts currently gated on the trust setting, 2 of them asserting through log greps. Effort small, risk medium, confidence verified.

#### A13. The parity instrument compares the host against the shim's reconstruction of the page, and the detector that watches that hop reaches only the log

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2560 and :2588 (both surface sides come from
_editorSurface.Text / TextOf; the host side IS read live at :2554 and :2573-2575 via
ProjectReader.ReadSource); src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:949-953 (Text and TextOf),
:1387-1406 (the text is rebuilt from incremental edits), :1394-1399 (the drift check `surface:
reconstructed N character(s) where the page holds M`, whose only outlet is Log.Error);
ui/editor/src/bridge.ts:1626, 1633 (the page does send fullLength);
tools/harness/xlide-api.mjs:384-403 (inSync's page contribution is a module NAME);
tools/harness/debugger-features.mjs:188

**What.** native compares the live host pane against the shim's own reconstruction of the page's text, which
the shim builds by applying incremental change messages to its own string. No route ever asks the
page what its model holds; the page contributes only a URI to inSync. A reconstruction bug - a
mis-parsed change range, a dropped message - is already detected on every edit by a length
comparison against the page's own fullLength, and that detector writes a log line and nothing else.
Grep for "reconstructed" finds it in no route, no reply and no counter.

**Why it matters.** Editor parity is a definition-of-done property and this is the instrument that certifies it. The
layer stops ARE designated in the docs (the reply's field is named surfaceContent and the route
table says "what the surface believes it is showing"), so this is not a false observation - but the
shim-to-page hop is the one link nothing asserts, and the check that would catch it cannot be read
by any suite.

**Fix.** The load-bearing half is small: expose the reconstruction-mismatch count from EditorSurface.cs:1398
as a counter in `stats`, so a suite can assert it is zero without reading the log. Optionally add
pageContent/pageContentKey to native, read through RunPageScript from the model the render uses, so
the comparison becomes three-way; that is a nice-to-have next to the counter.

**Size.** one counter in stats plus one page round trip if the third column is wanted. Effort medium, risk low, confidence verified.

#### A14. Panes cannot be opened or closed through the api and ui.panes omits which side a pane is on, so a gate probe drives the Panes menu by selector

**Where.** ui/editor/src/devsurface.ts:127-131 (DevSurfaceParts.panes declares only list and moveTo),
:1235-1251 (dock is the only pane act), :461 (the snapshot drops even the permanent flag);
ui/editor/src/shell.ts:349-362 (the object main.ts actually passes already carries setOpen over
docks.open/close), :268-278 (the six seats), :293 and ui/editor/src/main.ts:518 (visibilityChanged
reaches the host); ui/editor/src/paneldocks.ts:222-265 (open/close), :269-278 (findPane holds {side,
group}); src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3659-3666 (immediate visibility gates polling);
tools/harness/Test-SplitWorkspace.ps1:258, 270-271, 305 (six selectors), driven by
tools/verify.ps1:369

**What.** The concrete pane object already has setOpen; the interface the act table sees does not declare it,
so there is no way to show or hide a pane. A gate probe therefore opens one by clicking #toolbar
[data-command=openPanes] and then #panes-menu .panes-menu-item, closes one with .dock-close, and
answers "which side is this pane on" by walking dock-left/right/top/bottom for a .panel-tab. Note
the layout route already answers the side question (LayoutScript at
AddInSession.DebugApi.cs:443-455), also by scraping the DOM from C# rather than reading PaneDocks -
so the observation exists twice as a scrape and zero times as structure.

**Why it matters.** Opening a panel is an everyday action reachable only by knowing three private details of the toolbar
and the menu. Pane visibility is not cosmetic: the immediate pane's open state flips
_watchingImmediate and changes the host poll interval. The dock act can flip visibility indirectly,
so that transition is not wholly unexercised, but the pane's own open/close is not drivable at all -
and the locals and watches routes read shim state, so a suite can assert on rows the developer
cannot see because the pane is closed.

**Fix.** Add setOpen to the DevSurfaceParts.panes interface and an act that calls it, answering the pane's
open state afterwards from panes.list() so the reply is the outcome rather than the request. Carry
side and permanent on the existing snapshot, taken from PaneDocks.findPane rather than the DOM,
which also gives the layout route's scrape a typed answer to fall back on. The menu-driven half
comes free with the toolbar act at rank 3.

**Size.** one interface line, about 15 lines of act code, 2 fields added to ui.panes; 6 private selectors retired from a gate probe. Effort small, risk low, confidence verified.

#### A15. The status bar - caret readout, module name and the notice line - is in no snapshot, so a whole class of declined actions cannot be shown to have been reported

**Where.** ui/editor/src/shell.ts:155-157 and :200-202 (statusPosition, statusModule, statusNotice), :747-760
(notify sets the text, adds .visible, clears after 5000ms), :876-878 (setPosition);
ui/editor/src/devsurface.ts:38-93 and :458-475 (UiSnapshot has no status block); call sites
ui/editor/src/bridge.ts:1125 (host notice), ui/editor/src/main.ts:526-530 (tree rename summary),
:833 (F2 rename summary), :1127-1129 (Undo Rename); ui/editor/src/devsurface.ts:751-772 and :745-746
(the renameModule act calls the bridge directly and its comment claims otherwise);
docs/debug-api.md:70

**What.** notify is the documented reply to actions that were legitimately declined, the deliberate
alternative to a dialog, and nothing can read it. Host-originated notices do cross the transport as
a `notice` message and are visible in the messages route, so that class is observable at the
transport - but that the shell rendered it is not, and the three page-local notify calls never leave
the page at all. Coupled to this: the renameModule act awaits bridge.requestModuleRename directly
rather than the shell handler the tree's Rename item runs, so the notice never fires, while the
act's own comment and the route table both claim it is the path the menu item takes.

**Why it matters.** "The refusal was shown to the developer" is untestable, and a rename that silently reports nothing
looks identical to one that reported correctly. The false parity claim in the docs is what makes
that silent: a reader has no reason to add the missing check.

**Fix.** Add a status: {position, module, notice, visible} block to UiSnapshot read from the shell's own
last-set values rather than textContent, with a flag for whether the notice is still inside its
five-second window. Then either route the renameModule act through the shell's handler (passed into
DevSurfaceParts the way openSettings is) so the notice fires, or keep the direct call and designate
the omission in the act comment, the docs/debug-api.md act row and the client method. The doc half
is worth doing regardless of the snapshot.

**Size.** three fields, about 15 lines, plus one handler reference or three doc edits. Effort small, risk low, confidence verified.

#### A16. act references answers the lookup and never opens the dialog whose rendering is the feature, and the act appears in neither api doc

**Where.** ui/editor/src/devsurface.ts:1117-1135 (the act ends at parts.referencesAt and returns the found
list), :1114-1115 (the comment designating the difference); ui/editor/src/main.ts:944-948 (the
wiring), :1139-1147 (the Shift+F12 action and context-menu entry run showReferences), :203-213
(showReferences calls the same lookup then openReferencesDialog), :1135-1138 (why the dialog exists:
monaco's peek cannot render a module with no tab open); docs/debug-api.md:70 (references absent from
the do= list); docs/driving-excel.md (references absent from every act example)

**What.** Find All References exists precisely because the editor's own peek window cannot render a module
with no tab open, so the dialog's rendering is the feature. The act that carries its name performs
the lookup and returns data, leaving no dialog on screen, so ui.dialogs after a references act shows
the same surface as before. The deviation is designated in the code comment and in neither document,
and audit-routes audits routes rather than act actions, so nothing in the gate catches the omission.

**Why it matters.** A broken references dialog passes every check that exists, and a reader consulting either api
document cannot tell the act stops short - the act is not listed in them at all.

**Fix.** Add an open=1 argument that calls the same showReferences the menu entry runs, so the dialog stands
afterwards and ui.dialogs sees it; keep the data-only form as the default. List references in the
act row of docs/debug-api.md and in the act examples of docs/driving-excel.md, naming which layer
each form stops at.

**Size.** one argument plus one parts entry, about 20 lines, and two doc rows. Effort small, risk low, confidence verified.

#### A17. The reference document makes five claims the code does not honour, two of them repeated in code comments where the next maintainer will read them

**Where.** docs/debug-api.md:207-208 ("immediate only SCHEDULES") against
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:756-780, 792-794 (it waits, default 15000) and
:642-657 (the header recording the change); docs/debug-api.md:99 ("ten second wait" against a 15000
default); src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:497-499 (the dispatcher summary repeats
the retired claim); docs/debug-api.md:59 and :316 ("every request this door has served") against
src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:235-241; docs/debug-api.md:70 (renameModule "the one
the tree's Rename item runs"); docs/debug-api.md:91 and docs/driving-excel.md:210 (undoRename, no
mention that it bypasses the page); docs/debug-api.md:305-307 (the assert claim list says which
claims exist and never which layer any of them reads, e.g. shownModule reads _editorSurface.Module
at DebugApi.cs:2104-2110 while the same file reads ActiveCodePane at :2517-2542)

**What.** Five separate statements in the reference document contradict or overstate the code, and the
`immediate` one is contradicted twice within that document and once more in the summary of the
method that dispatches every route. A reader deciding whether immediate can run a statement that
might break is told the opposite of what the code does. The ten-versus-fifteen-second figure sets
the wrong client timeout; xlide-api.mjs:934 hardcodes 20000, which clears 15000 by accident.

**Why it matters.** This is the cheapest item here and it is pure grounding: the reference document is where every
future driver of this api starts, and a doc that claims parity the code does not have is exactly
what makes a false negative silent. No risk, no code behaviour changes.

**Fix.** Delete the "only SCHEDULES" paragraph, correct the ten-second figure and name waitMs as the
override, fix the stale sentence in the AnswerDebugRequest summary, correct the two "every request"
sentences, qualify the renameModule and undoRename rows, and say for assert which layer each claim
reads (or add a nativeModule claim beside shownModule, evaluated from ActiveCodePane the way native
does).

**Size.** about eight sentences across docs/debug-api.md, docs/driving-excel.md and two code comments. Effort small, risk low, confidence verified.

#### A18. The Watch panel can be read and asserted but never populated, so its only live coverage spawns a PowerShell helper that types into the native dialog with WM_CHAR

**Where.** tools/harness/immediate-watch.mjs:10 (the header declaring the gap), :162-168 (command("addWatch")
checked only for reaching the editor), :170-177 (dismisses whatever it opened);
tools/harness/Test-WatchPanel.ps1:93-95 (writes a helper to TEMP and spawns it), helper :26
(EnumWindows), :45 (GetDlgItem 4853), :48-51 (one WM_CHAR per character, with the comment that
WM_SETTEXT produces "Empty watch expression"), :55 (BM_CLICK), :135 (the verdict is a log grep);
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:816 (the watches route) and :2127 (watchHas);
src/Xlide.Vbe.Shim/Diagnostics/DialogWatch.cs:1 (#if DEBUG), :143-153 (walks children, can only post
BM_CLICK to a Button)

**What.** The observation half exists first-class and the action half stops one step short: the gate suite
reaches the honest limit of the api, opens the Add Watch dialog and cancels, so it asserts a list
exists and never puts a row in it. The only thing that populates the panel is an out-of-gate probe
that needs AccessVBOM, spawns a second process, and concludes from a log line rather than from the
panel.

**Why it matters.** The Watch panel is a shipped surface whose value-tracking behaviour is proven only by a probe nobody
runs, while the suite that is in the gate cancels. The watchHas assertion cannot be made true from
the api at all.

**Fix.** Give DialogWatch a fill capability and expose it as a route (fill?caption=Add Watch&field=<control
id or preceding label>&text=Counter), typing WM_CHAR the way the helper already proves works, then
dismiss("OK"). Ranked here rather than higher because it is a judgment call the owner has to make,
not a free extension: DialogWatch is compiled out of Release so this is not shipped product code,
but it does move keystroke synthesis into shim source, against the standing no-synthetic-input rule.
If that is refused, the honest alternative is to say in the route table that watchHas has no api
driver and keep Test-WatchPanel.ps1 as the designated exception.

**Size.** 1 gate suite that cancels rather than asserts; 1 out-of-gate probe with a spawned helper and about 45 lines of Win32. Effort medium, risk medium, confidence verified.

#### A19. build-fixture's write-retry can no longer run for the refusal it was written to absorb, because the write route learned to report it and the client turns that into a throw

**Where.** tools/harness/build-fixture.mjs:38-49 (the stated premise: "the write route reports that in the log
rather than in its reply"), :53-66 (the catch rethrows anything not matching /did not answer in
time/); src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3258-3263 (the route now answers
WriteModule's complaint as an error) and :3253-3257 (the comment recording the change and naming the
fixture builder's workaround); src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1946-1954 and :2061-2063
(WriteModule returns a string for the AddFromString refusal and from its outer catch);
tools/harness/xlide-api.mjs:295-298 (the client throws on any error field)

**What.** The first write to a module that is freshly added and currently shown is refused with "Invalid
procedure call or argument"; the retry exists for exactly that. The route now reports it, the client
throws, and the message does not match the timeout regex, so attempt 1 rethrows and the retry never
runs for its own case. Narrower than it first looks: the loop's other path, the line-count poll at
:78-105, still retries a write that fails silently.

**Why it matters.** Fixtures are the precondition for the live suites, so this surfaces as a whole -Live pass that
cannot start, reported as an error about a module rather than about a retry that stopped working.

**Fix.** Retry when the message matches the door timeout OR the writer's refusal wording, rethrow otherwise,
and rewrite the comment at lines 38-49, which describes behaviour the route no longer has.

**Size.** one regex in one catch, plus a stale comment. Effort small, risk low, confidence supported.

#### A20. Two orphaned probes still hold VBProject call sites for work that routes now do

**Where.** tools/harness/Test-GhostLocalsPanel.ps1:10-27 (VBProject fixture), :29 and :45-53
(CommandBars.FindControl for stepInto and reset), :36 ($app.OnTime to provoke the break), :63-68
(verdict from log greps); tools/harness/Get-EditorScreenshot.ps1:252 (VBProject fixture); the
replacements: src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2707 (component add),
src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:236-240 (reset, stepInto, stepOver, stepOut),
DebugApi.cs:797 (locals) and :2118 (localsHas), tools/harness/Get-Shot.ps1:60 (wraps the capture
route); referenced only by docs/locals-break-investigation.md:24,163 and Get-EditorScreenshot's own
help text

**What.** Every step of the ghost-locals probe has had a route for some time, and debugger-features.mjs and
step-into-features.mjs already drive break, step and parity entirely through the client. Nothing
references either script from code. Left in tools/harness they read as the way to test the Locals
panel and the way to take a screenshot, they need the trust setting to run at all, and the ghost
probe's log-grep assertion passes on a session whose panel renders nothing, because a pushed row and
a drawn row are different claims.

**Why it matters.** This is the api-coverage gap seen from the far end: the route arrived and the probe was never
retired, which is the rot the audit's missing reverse pass (rank 11) would prevent. It also keeps
two of the harness's remaining VBProject call sites alive.

**Fix.** Fold the ghost probe's one distinct claim - locals values changing across steps, not merely being
pushed - into debugger-features.mjs using component/writeModule for the fixture, the
breakpoint-plus-run path that suite already uses (its $app.OnTime provocation has no route
equivalent and must be replaced, not ported), command?name=stepInto for the steps and
locals()/assert localsHas for the assertion. Then delete both scripts; Get-EditorScreenshot is
superseded by Get-Shot.ps1 over the capture route.

**Size.** 2 orphaned scripts, 447 lines, carrying 2 of the harness's remaining VBProject call sites. Effort small, risk low, confidence verified.

#### A21. The client cannot send keep on command or selector/pad on capture, so no .mjs suite can open a dialog deliberately or take a cropped screenshot

**Where.** tools/harness/xlide-api.mjs:426 (command takes only name) and :424 (capture takes only window);
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1894 and :1903 (keep is honoured generically on
every host-thread route), :583-600 (the crop); docs/debug-api.md:96, :279 ("A request that means to
open a dialog passes keep=1 and what it opens is exempt"), :66; the only users hand-build the URL:
tools/harness/Test-DebugApi.ps1:466, :514, :416 and tools/harness/Get-Shot.ps1:54-55; grep of the
.mjs corpus for .capture( returns zero calls

**What.** Two pass-through query arguments the routes support and the client does not expose. Without keep, a
.mjs suite that opens a dialog on purpose is fought by the dialog guard -
immediate-watch.mjs:170-177 has to sweep up after it - so the deliberate-modal path is only ever
exercised from PowerShell. Without the crop, capture returns a whole frame, which is a picture in
which a 54-pixel drop zone cannot be seen, and no .mjs suite screenshots at all.

**Why it matters.** keep is a safety argument, and its absence from the client is why an entire interaction class is
confined to the PowerShell probes that are outside the gate.

**Fix.** Add optional arguments to the two existing client methods: command(name, {keep}) and capture(window,
{selector, pad}). No route changes. Update both client rows in docs/driving-excel.md (:201 for
capture, :203 for command).

**Size.** two client signatures, two doc rows. Effort small, risk low, confidence verified.

#### A22. history drops five routes from a transcript documented as every request, so the replay script it generates loses exactly the waits that made the session work

**Where.** src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:235-241 (`if (route is "history" or "log" or "journal"
or "state" or "dialogs") return;`), :1034-1039 (DebugHistoryReply carries no exclusion list and no
dropped count); src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:885-889 ("a bug found by hand
becomes a probe by copying") and :891-902 (the script builder emits one line per recorded request
with nothing between them); docs/debug-api.md:59 and :316

**What.** The exclusion is sound - polled routes would be the whole transcript - but it is invisible, and the
routes excluded are the log?match=&waitMs= waits and the state polls that sequence everything else.
Copy the generated script out and the steps run back to back with no waits, so a session that only
worked because of its waits replays as a race, with nothing in the reply hinting that anything was
removed.

**Why it matters.** The route exists so a hand-driven session becomes a probe by copying, and the copy is missing the
part that made it deterministic.

**Fix.** Keep the exclusion, make it visible: add a notRecorded field listing the five route names, and emit
a `# waited on log/state here` marker in the generated script where a polled route was skipped.
Correct the two "every request" sentences (folded into the doc sweep at rank 17 if that lands
first).

**Size.** one field plus a script marker and two doc sentences. Effort small, risk low, confidence verified.

#### A23. assert polls a misspelt claim name for the full timeout and answers it in the same shape as a real failure

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:916-933 (the poll loop, timeout default 10000, no
pre-check), :2147-2148 (`default: return (false, $"unknown claim {claim}");` - a value, not a
refusal), :935-937 (a normal DebugAssertReply, so the client's error check at xlide-api.mjs:295-298
does not fire); the precedent in the same file: :1368 (bench answers `unknown benchmark {what}; try
tabswitch, layout, or type`) and :1492-1498 (trip)

**What.** assert?that=shownModul sleeps ten seconds and answers
{"held":false,"claim":"shownModul","saw":"unknown claim shownModul"}. A check(..., answer.held, ...)
call site never reads the prose in `saw`, so the typo presents as a failed assertion against the
product. The claim vocabulary is a string list in three places that can drift: the switch at
:2093-2150, docs/debug-api.md:306-307, and the client docstring at xlide-api.mjs:592-594.

**Why it matters.** An assertion instrument that reports its own typo as a product failure sends the reader to the wrong
place, and each occurrence costs the caller its whole timeout.

**Fix.** Evaluate once before entering the poll and answer {"error":"unknown claim X; known claims are ..."}
immediately when the name is unrecognised, the treatment bench and trip already give an unknown
argument.

**Size.** one pre-check; saves 10s per typo. Effort small, risk low, confidence verified.

### Themes

- A route wraps a private void session method and synthesises a constant reply, so every decline path
  in the callee reaches only the log: command, compile, pane close, undoRename, caret and breakpoint
  all report the request rather than the outcome. The rule this suggests is that any session method a
  route calls must return its refusal, and no route may hard-code true.
- The act vocabulary covers what a keystroke reaches and stops where a click begins. Keys, carets,
  language providers and monaco actions are first-class; toolbar buttons, menu items, dialog controls,
  property grids and tab gestures have no action at all, so the harness clicks by selector - and every
  selector it knows is a private detail no compiler protects.
- An instrument reads the shim's own record and is presented as the user's view. ui.search.matches,
  assert shownModule, native's surface side and inSync's page contribution all stop at a layer the
  caller cannot see, and the ones that are honest about it are honest only in the reference document,
  never in the reply.
- Half a feature gets an api and the missing half is where the false pass lives: watches and locals
  have observers with no driver, sync and references have drivers with no observer, and scoped search
  has a driver that does nothing paired with an observer that is structurally zero.
- A capability the api never grew pushes the harness back onto COM or DevTools, and the gate cannot
  see it happen. Six scripts still need the AccessVBOM trust setting the product forbids, one probe
  hand-rolls a DevTools client, and audit-routes asks only whether routes are driven, never whether
  the harness went around them.
- One sibling learns a lesson and the others do not, and the docs keep the old claim. remove was
  taught to carry the bookkeeping and rename was not; open was taught to answer honestly and close was
  not; immediate was taught to wait and its own reference row, the dispatcher summary and a second doc
  line still say it only schedules.

### What this section leaves out

Merges: 34 surviving findings became 23 items. Merged pairs and triples - search driver plus search
observer (rank 1); command plus compile, one discarded bool in ExecuteEditorCommand (rank 2); the
two Properties-pane findings, filed independently by the ui and shim finders (rank 5); the two
sync-dialog findings, one from the harness end and one from the page end (rank 6); pane close plus
undoRename as one constant-true family (rank 8); palette-has-no-ui-act plus the CDP probe, since the
probe's transport exists because ui/act cannot address that surface (rank 9); frame visibility plus
native pane close as one AccessVBOM cause (rank 12); audit-routes' vacuous log match plus its
missing reverse pass (rank 11); pane open/close plus ui.panes-omits-side (rank 14);
statusbar-no-observer plus renameModule's unraised notice (rank 15); and five separate doc-truth
defects into one sweep (rank 17). undorename-menu-item-no-driver was folded into the toolbar act at
rank 3, since its own proposed fix is that act. Dropped 2 as trivial or subsumed:
native-route-thin-for-background-panes, because the adversary retired three of its four supporting
claims (the hash IS stable within a process and carries a portable length prefix, both texts are
already reachable via module and module?live=1, and Run/Step act on ActiveCodePane whose caret the
reply already carries), leaving a convenience improvement; and clientfor-not-exported, one missing
export whose sole consumer stops needing its own transport if rank 9 lands. What I verified myself
in this pass: the search act and DevSurfaceParts.search (devsurface.ts:116-131, 943-962), the
widget's state/find/queryChanged/scopeChanged/showSearchResults and its click wiring
(searchwidget.ts:190-229, 316-335, 460-520, 545-585), ExecuteEditorCommand end to end including its
two early returns (AddInSession.cs:1718-1787), the command route (DebugApi.cs:3188-3200), and the
component rename route against AdoptRename (DebugApi.cs:2836-2874, AddInSession.cs:3340-3385).
Everything else I carried from the adversary's CONFIRMED verdicts with their corrections applied, so
line numbers in files I did not open are the verifiers' and not mine - I did not independently
re-read the shim's lifecycle code, the PowerShell probes, the sync dialog, or audit-routes' corpus
counts. Not examined by anyone in this set as far as I can tell: the engine's own JSON-RPC surface
under engine/src (only engine-live-probe.mjs is cited, and only in passing), the journal, perf,
bench, trip, doctor, guard and outline routes beyond incidental mentions, and whether the roughly 60
api.eval sites outside the two named suites hide further scrapes of state a typed field could
answer.

---

## B. Complexity and dead code

Code that can be removed or simplified without changing behaviour.

The repo is disciplined about the things it has been burned by and undisciplined about everything
adjacent to them. There is no dead weight in the product's core behaviour: the menu machinery, the
placement code and the engine protocol are all live, and the biggest single file (AddInSession.cs,
7,741 lines) is repetitive rather than rotten. What the audit found instead is three separable
classes. First, ownership gaps at the edges of the COM discipline the codebase otherwise enforces
perfectly: one cached Application wrapper that Stop() never releases, two collection loops whose
manual Dispose is skipped on a throw, and a Debug-only handler list that only ever grows - the last
two of which the leak instrument (com-leak.mjs, stats.comWrappersLive) cannot see by construction.
Second, shapes written out three to thirteen times instead of once, in every layer: the engine
round-trip handler (7 copies), the debug-api reply shaping (5 small duplications), the page's
pending-request tables (10), the modal plumbing (6), the settings key list (6 on the page plus 3 on
the host), and the harness scaffolding (4 engine tests, 9 mjs suites, 4 fixture generators, 9
PowerShell discovery blocks). The settings copy is not merely ugly - it is the mechanism that
dropped syncEngine from the page's update, which now persists "xlide" over the user's choice on
every unrelated settings change. Third, a runner list the gate maintains by hand while
audit-routes.mjs counts the whole harness directory as coverage, so four routes, ten seam checks,
two live probes, three shipped-crash guards and one engine test report as covered while running
nowhere. Dead code proper is small and mostly declarations: 25 unreferenced constants and imports
across HResult.cs and Win32*.cs, four page-protocol message kinds with no sender, one unreachable
40-line COM enumerator, one engine method nothing sends.

### The ranked list

| # | Finding | Effort | Risk | Confidence |
| --- | --- | --- | --- | --- |
| 1 | Stop() releases every automation reference except _hostApp, so a teardown inside a living process abandons a wrapper on Excel's Application to the finalizer thread | small | low | verified |
| 2 | The page's updateSettings omits syncEngine, so the setting is unreachable from the dialog and every other settings change persists "xlide" over it | small | low | verified |
| 3 | Every page script call permanently roots its handler: ForgetScriptHandler exists, has no caller, and _pendingScripts is never cleared | small | low | verified |
| 4 | Two collection loops take a COM item into a plain local and dispose it by hand, so any throw between the take and the dispose drops the wrapper | small | low | verified |
| 5 | The six settings keys are spelled out structurally nine times and nothing makes the copies agree - this is the cause behind item 2, not another instance of it | small | low | verified |
| 6 | engine/test/freshness.mjs runs in no runner: CI runs three of the four engine tests and verify.ps1 runs one | small | low | verified |
| 7 | release.ps1 -SkipGate removes the only engine-freshness check and repackage, and nothing later in the release path re-checks or identifies the engine that shipped | small | low | verified |
| 8 | audit-routes.mjs counts every file in tools/harness as a driver, so routes pass the gate's coverage check while nothing the gate runs drives them | medium | low | verified |
| 9 | Seven harness scripts run nowhere, and among them are the only seam checks, the only engine and object-browser live probes, and the only guards for three shipped crashes | medium | medium | verified |
| 10 | Query flags are parsed by two incompatible rules in one switch, so perf?reset=false clears the counters | small | medium | verified |
| 11 | The page's host-message protocol carries three dead kinds and its only written reference documents a fourth that never existed | small | low | verified |
| 12 | AnswerBlockedRequest's only explanation sits above the wrong method and promises a retry the code does not perform, and the wire field that would report it is a constant false | small | low | verified |
| 13 | Five duplicated shapes in the debug-api file, each small, all fixable with a local helper the file simply never grew | small | low | verified |
| 14 | Seven engine round-trip handlers are the same capture-guard-deadline-map-marshal block, and the drift has already started | medium | low | verified |
| 15 | bridge.ts carries ten hand-copied pending-request tables where one helper would do | medium | medium | verified |
| 16 | WatchReader is a transcription of LocalsReader, including the hardening that came out of the 2026-08-05 crash | medium | medium | verified |
| 17 | Six dialogs hand-roll the same modal plumbing, five re-declare CSS a general .modal-backdrop rule already provides, and none of the six traps focus | medium | low | verified |
| 18 | The whole-module read and the component walk are each reimplemented three times, once in production and twice behind the debug api | medium | medium | verified |
| 19 | The pool-side route switch has no default, so a route whose argument guard fails is marshalled to the host thread it exists to avoid | small | medium | verified |
| 20 | The import path reads the whole module twice before writing it, and the second read is provably the same text as the first | small | low | verified |
| 21 | The engine answers a protocol version on initialize and the shim discards it | small | low | verified |
| 22 | Comments and doc blocks that describe something other than the code beneath them, in the three files whose comments are most load-bearing | small | low | verified |
| 23 | PublishModules walks Excel's Workbooks collection once per distinct workbook on every poll tick, and the change-key that would make it free is computed after that work | medium | low | supported |
| 24 | The test harness has no shared layer, so the same scaffolding is written four, nine and four times | medium | low | verified |
| 25 | Nine PowerShell scripts rebuild the debug-api discovery path by hand and six of them pick the first Excel they find | medium | low | verified |
| 26 | Twenty-five unreferenced constants and P/Invokes in the interop surface, four of them describing an object model the shim does not implement | small | low | verified |
| 27 | Native declarations and window helpers written outside Interop, duplicating what Interop already owns | small | low | verified |
| 28 | VbeCommands.Describe is unreachable, and the doc comment that sends a maintainer to it is the only reference to its name | small | low | verified |
| 29 | The split-tree prune and same-direction absorb are written three times, and only the copy in the module built to be tested is tested | large | medium | verified |
| 30 | docs/testing.md describes the live gate as the PowerShell probes and never mentions the ten node suites that are most of it | small | low | verified |
| 31 | The engine routes and documents module/didClose, which nothing has ever sent, and the dispatcher says so thirty lines below | small | low | verified |
| 32 | tools/page.ps1 and tools/Update-Page.ps1 are the same page loop, and the operational guide points at the thinner one | small | low | verified |
| 33 | surface-walk.mjs and four doc examples require TwinFixture.xlsm, and no generator makes it | small | low | verified |
| 34 | A literal form feed byte sits where a backslash-f belongs in the driving guide's two-workbook command, so the sample cannot be copied and run | small | low | verified |

#### B1. Stop() releases every automation reference except _hostApp, so a teardown inside a living process abandons a wrapper on Excel's Application to the finalizer thread

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5664 (field), :5685 and :6583 (taken), :5708 and :6609 (the
only Dispose calls, both in catch), :7593-7666 (Stop), :7733-7740 (Dispose);
teardown-in-a-living-process callers at src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs:208-210 and :250-252

**What.** `private DispatchObject? _hostApp;` is populated by `_hostApp ??= HostApplication.Find();` on the
WorkbookSaved and SaveWorkbookOf paths, and HostApplication.Find's own doc says the caller owns the
result. I read Stop() in full: it releases _debugServer, _analysis, _frameSubclass,
_immediateReader, _ghostReaders, _typeLibraries, _browserPalette, _codePanes, _hostChrome,
_editorSurface and both ghost palettes, and _hostApp appears nowhere; Dispose() is `Stop();
_addIn?.Dispose(); _editor.Dispose();`. The only two disposals are inside stale-answer catch blocks.
DispatchObject has no finalizer and Dispose is the only route to ComRuntime.GiveBackWrapper, so the
wrapper is orphaned and its inner UniqueInstance ComObject is released on the finalizer thread.

**Why it matters.** src/Xlide.Vbe.Shim/Com/DispatchObject.cs:724-740 records what that costs: a finalizer-thread release
read as an access violation inside Marshal.Release, which under ahead-of-time compilation becomes a
FailFast that takes Excel with it. WorkbookSaved runs on the publish path so _hostApp is populated
in essentially every session, and Stop()+Dispose() run in a living process on watchdog revival after
a cancelled shutdown and on a non-HostShutdown OnDisconnection. com-leak.mjs cannot catch it: it
reads stats while the session is alive, and this wrapper is steady during the session and only
orphaned at teardown.

**Fix.** Add `_hostApp?.Dispose(); _hostApp = null;` to Stop(), beside the other automation-reference
releases (after _codePanes, before _hostChrome).

**Size.** 2 lines added; one leaked COM wrapper per non-shutdown teardown. Effort small, risk low, confidence verified.

#### B2. The page's updateSettings omits syncEngine, so the setting is unreachable from the dialog and every other settings change persists "xlide" over it

**Where.** ui/editor/src/bridge.ts:672-681 (post body) and :300-308 (outbound union member);
ui/editor/src/settingsdialog.ts:56 and :346-348; ui/editor/src/main.ts:471 and :570;
src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1623-1636;
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:861-869; contrast
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2997-2999

**What.** I read the post body: it sends blockLayout, continueCommentOnNewline, mirrorCommentSpacing,
treeFollowsEditor and formatIndentSize, and no syncEngine; the outbound union member lacks it too
while the inbound setSettings type has it. settingsdialog.ts declares a syncEngine choice whose
handler spreads currentSettings() into bridge.updateSettings. On the host,
EditorSurface.cs:1623-1626 turns a MISSING syncEngine into the literal "xlide" rather than the
stored value, builds a fresh ProductSettings and fires SettingsChangeRequested;
AddInSession.OnSettingsChanged then assigns `_settings = updated.Normalized()` and immediately
writes the settings file. The debug-api route does the opposite, falling back to
`settings.SyncEngine`.

**Why it matters.** Choosing the built-in planner in the dialog does nothing, and changing indent size or block layout
silently overwrites a planner choice made through the api - persisted to disk, not just to the
session. It is the api-mirrors-the-UI rule inverted: the only driver in the repo is
tools/harness/module-sync.mjs:34-36 going through api.settings, so the whole harness passes over a
broken UI path, and docs/debug-api.md:90 asserts the page's update takes the whole object.

**Fix.** Add `syncEngine: string` to the updateSettings member and pass `settings.syncEngine`. Separately
make EditorSurface's absent-field branch fall back to the stored value the way the debug-api route
does, so no page can ever reset a field by omitting it. Drive the dialog's select through the page
in whichever live suite covers settings.

**Size.** one omitted field; one setting unreachable from the UI and overwritten on disk by five other UI actions. Effort small, risk low, confidence verified.

#### B3. Every page script call permanently roots its handler: ForgetScriptHandler exists, has no caller, and _pendingScripts is never cleared

**Where.** src/Xlide.Vbe.Shim/WebView/WebView2Surface.cs:132 (Add), :141 (the list), :143 (the un-root nobody
calls), :929-940 (CreateCallback), :1065-1077 (Invoke), :942-998 (Dispose); sole caller
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:221, reached from RunPageScript's poll loop at
DebugApi.cs:143-148

**What.** I grepped the repo: `internal void ForgetScriptHandler(object handler) =>
_pendingScripts.Remove(handler);` has no call site anywhere, and the list has exactly one Add and no
Clear. ExecuteScriptCompletedHandler.Invoke calls the completion delegate and returns HResult.Ok
without un-rooting; Dispose nulls the five other handler fields and never touches _pendingScripts.
CreateCallback registers each handler with ComRuntime.Wrappers.GetOrCreateComInterfaceForObject,
which is not one of the two doors (TakeWrapper / GiveBackWrapper) that WrappersTaken/GivenBack/Live
count.

**Why it matters.** Growth is per RunPageScriptOnce, not per route call - the backoff loop re-invokes on every tick and
ui, act, eval, run, wait, layout and a dozen more routes all go through it, so a live suite
accumulates thousands of rooted handlers plus their closures (which pin the already-disposed
ManualResetEventSlim). It is Debug-only so nothing ships, but Debug is exactly where the leak
instrument must be trustworthy, and this growth is invisible to WrappersLive by construction. The
un-root method being present makes the code read as if the lifetime were managed.

**Fix.** Wrap the caller's callback in ExecuteScript so completion calls `completed(code, json)` then
`ForgetScriptHandler(handler)`, and clear _pendingScripts in Dispose. If a handler is meant to
outlive its call, delete ForgetScriptHandler and say so at the list.

**Size.** one rooted handler plus closure per page-script call, unbounded over a session; or 1 dead method removed. Effort small, risk low, confidence verified.

#### B4. Two collection loops take a COM item into a plain local and dispose it by hand, so any throw between the take and the dispose drops the wrapper

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:6010-6031 (FindProjectByDisplayName, catch at :6027
swallows without disposing) and :6997-7007 (FindComponent, no try at all); the reads that can throw
at src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:92 and AddInSession.cs:5969

**What.** `grep -n "GetItem(" AddInSession.cs | grep -v "using var"` returns exactly these two sites. Both do
`var x = collection.GetItem(i);` then a property read then `x.Dispose()` on the non-match branch.
ProjectReader.Identity's `project.GetString("Name")` sits outside its own try and
WorkbookDisplayName does the same after its FileName guard; GetString goes through InvokeCore, which
throws on hr<0. In FindProjectByDisplayName the enclosing catch logs and returns without disposing
the item in hand; in FindComponent the throw simply escapes.

**Why it matters.** Same consequence as item 1 - a dropped wrapper is a finalizer-thread release, which the codebase has
already paid for three times. The manual-dispose pattern is correct only on the happy path, and
these are the lookup helpers behind write, rename, close, replace-all and publish, so the exposure
is any transient refusal from the editor mid-enumeration.

**Fix.** Wrap each loop body in `try { ... } catch { item.Dispose(); throw; }`, or take the item into a
`using var` and hand ownership out deliberately on the match branch. About three lines per site, no
happy-path change.

**Size.** ~6 lines across 2 sites. Effort small, risk low, confidence verified.

#### B5. The six settings keys are spelled out structurally nine times and nothing makes the copies agree - this is the cause behind item 2, not another instance of it

**Where.** ui/editor/src/settings.ts:10-27 and :49-58; ui/editor/src/bridge.ts:91-99, :300-308, :672-681,
:1332-1340; src/Xlide.Vbe.Shim/Editor/EditorMessages.cs:84;
src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:953;
src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1600-1636

**What.** Six independent structural spellings on the page (interface, inbound message member, outbound
message member, hand-written post body, hand-written adopt block, normaliser) plus three on the host
(the SetSettings record, the debug reply record, the field-by-field parse). Neither bridge.ts
message member references the EditorSettings type - both restate the members inline - so the
compiler cannot notice that the outbound shape is one member short of the inbound one.

**Why it matters.** Adding a setting means five coordinated edits in two files with no type relating them, and the two
shapes that must agree (what the page sends and what it receives) are the two that diverged. Fixing
item 2 without fixing this leaves the next setting to fail the same way.

**Fix.** Define the payload once in settings.ts and use it in both members: `{ type: "setSettings" } &
Partial<SettingsPayload>` (the inbound side deliberately tolerates a host that predates a field) and
`{ type: "updateSettings" } & SettingsPayload`. Then updateSettings can post `{ type:
"updateSettings", ...settings }` and the adopt case can call applySettings(message), leaving
coercion in the one place that already does it.

**Size.** 6 page-side spellings collapse to 1; roughly 35 lines out of bridge.ts. Effort small, risk low, confidence verified.

#### B6. engine/test/freshness.mjs runs in no runner: CI runs three of the four engine tests and verify.ps1 runs one

**Where.** engine/package.json:12; .github/workflows/build.yml:69, :85, :94 and :138-139;
tools/verify.ps1:229-244

**What.** package.json's test script names all four tests. The CI engine job runs smoke, language and
positions as individual `node test/*.mjs` steps and never freshness; the `npm test` at build.yml:139
is under `working-directory: ui/editor`, so it is the page's script. verify.ps1's engine step runs
language.mjs alone. A repo-wide grep for freshness.mjs finds only package.json and
docs/lessons.md:1439.

**Why it matters.** 11,989 bytes and five checks over the re-analysis-on-callee-change invariant never execute. That is
the same invariant tools/harness/analysis-freshness.mjs guards at the live layer, so when the live
suite goes red there is no cheap host-free test to bisect against, and the engine half can rot
silently in the meantime.

**Fix.** Add `node test/freshness.mjs` to the engine CI job beside Positions, or make the job run `npm test`
so package.json stays the single list. CI rather than verify.ps1: it needs no Excel and both runners
already build the bundle.

**Size.** 1 CI line; 5 checks that never run. Effort small, risk low, confidence verified.

#### B7. release.ps1 -SkipGate removes the only engine-freshness check and repackage, and nothing later in the release path re-checks or identifies the engine that shipped

**Where.** tools/release.ps1:53-57 (the skip), :60 (installer build), :74 (upload), :81-85 (the report);
tools/verify.ps1:94-162 (the only freshness comparison, which also repackages at :139-162);
installer/build.ps1:76-91

**What.** release.ps1 gates verify.ps1 behind `-not $SkipGate`, then builds the installer and uploads it. The
freshness comparison - engine\dist\xlide-engine.exe's LastWriteTimeUtc against both engine\src and
the neighbouring analyzer checkout - exists only inside verify.ps1, and that step does not merely
check: it repackages when a source is newer. installer/build.ps1 only asks whether any *.exe exists
in the engine publish folder. The closing report prints asset names and sizes from `gh release view`
and no hash or timestamp.

**Why it matters.** -SkipGate is documented for the ordinary case, and a session where the gate ran an hour ago and the
analyzer checkout was pulled since is precisely the 2026-08-06 shape the freshness step was written
for. This one ships to users, and nothing on the release page lets anyone determine afterwards which
engine went out. It also contradicts the standing verify-engine-on-release rule.

**Fix.** Move the comparison (and its repackage) into a function verify.ps1 and release.ps1 both call, run it
in release.ps1 unconditionally after -SkipGate, and add the uploaded exe's SHA-256 and
LastWriteTimeUtc to the closing report.

**Size.** 1 skippable step standing between a stale engine and an upload. Effort small, risk low, confidence verified.

#### B8. audit-routes.mjs counts every file in tools/harness as a driver, so routes pass the gate's coverage check while nothing the gate runs drives them

**Where.** tools/harness/audit-routes.mjs:137-140 (corpus), :144-148 (isDriven), :174 (the summary line);
tools/verify.ps1:246-252, :369, :420-427, :469 (what the gate actually runs)

**What.** The corpus is readdirSync over the whole harness directory, filtered only to .mjs/.ps1, and isDriven
regexes that blob. Intersecting each route's drivers against the files verify.ps1 invokes leaves
four routes whose only drivers never run: compile and guard (debugger-features.mjs,
step-into-features.mjs), mark (write-rollback.mjs), trip (perf-scaling.mjs). The audit's own failure
text says 'nothing in tools/harness drives it', so the check is honest about what it enforces; the
summary line reporting N routes 'driven by a probe', which verify.ps1 prints as the step verdict, is
not.

**Why it matters.** guard is the break-mode guard the two debugger suites arm and disarm, and compile is a real
behaviour; both can regress with every gate green. This is also the mechanism that hides item 9: the
audit is the only thing looking at harness coverage and it cannot distinguish a driver from a driver
that runs.

**Fix.** Give audit-routes.mjs the list of files the gate runs (read it out of verify.ps1) and split the
verdict into driven-by-something-that-runs versus driven-only-by-an-unrun-script, failing on the
second unless the file is excused by name the way NOT_DRIVEN_ON_PURPOSE already excuses routes.

**Size.** 2 routes with real behaviour and no automated coverage, reported as covered. Effort medium, risk low, confidence verified.

#### B9. Seven harness scripts run nowhere, and among them are the only seam checks, the only engine and object-browser live probes, and the only guards for three shipped crashes

**Where.** tools/harness/Test-CloseConfirm.ps1:31 (Test-Seam, 10 calls at :51-88) and :125
(engine-live-probe.mjs); tools/harness/Test-ObjectBrowser.ps1:29 (the second Test-Seam) and
objbrowser-live-probe.mjs; tools/harness/Test-CloseVbe.ps1, Test-CloseHiddenPane.ps1,
Test-ResizeFollow.ps1, rename-features.mjs, Get-EditorScreenshot.ps1; tools/verify.ps1:247 and :369;
docs/testing.md:15 and :24

**What.** verify.ps1 runs close-confirm-page-probe.mjs and objbrowser-page-probe.mjs directly and never the
wrappers, so legs 1 and 3 of each (the seam checks and the only invokers of engine-live-probe.mjs
and objbrowser-live-probe.mjs) execute nowhere. Separately, five scripts have no runner and no
mention anywhere outside themselves. Three of those five are unique cover: Test-CloseVbe.ps1 closes
the frame by WM_SYSCOMMAND three times and checks Excel survives (the 2026-08-04 crash),
Test-CloseHiddenPane.ps1 checks a hidden pane's close removes its tab, Test-ResizeFollow.ps1 checks
the overlay and browser child follow a frame resize - grep for resize or
SC_CLOSE/closeFrame/hidePane across tools/harness returns no .mjs at all. rename-features.mjs is
also unique cover, not superseded: it checks the ambiguous bare call left alone, Rival's own
Recalculate untouched, prefix-sharing HelpersExtra untouched and Go to Definition from a call site,
none of which three-copies.mjs does. Get-EditorScreenshot.ps1 largely duplicates Get-Shot.ps1 but
launches and closes its own host, which Get-Shot does not.

**Why it matters.** docs/testing.md keeps seam checks specifically because a rebuilt page that never reached the publish
tree has bitten more than once, and both Test-Seam implementations sit in files nothing invokes -
including Test-CloseConfirm's published-bundle stale-deploy check. The gate looks like it covers
close-confirm and the object browser and covers one third of each, and three named shipped crashes
have a guard that reads as present and is not.

**Fix.** Add Test-CloseConfirm.ps1 and Test-ObjectBrowser.ps1 to verify.ps1's live probe list (they emit the
RESULT: line that loop parses) or promote their missing legs - engine-live-probe.mjs needs no Excel
and can sit beside the page probes, and Test-Seam deserves a step of its own. For the five orphans,
decide per file: run it, or record in docs/testing.md that it is a hand-run reproduction, so the gap
is a decision rather than an accident.

**Size.** 10 seam checks, 2 live probes, 3 crash guards and 1 rename suite running nowhere. Effort medium, risk medium, confidence verified.

#### B10. Query flags are parsed by two incompatible rules in one switch, so perf?reset=false clears the counters

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1014, :1231, :1515, :2469, :2561, :3273
(anything-but-"0" is true) against :1764 and :2984 (affirmative word required, and :2982-2985
already has the local tri-state helper); docs/debug-api.md:81

**What.** Six sites read a flag as `value != "0"` and two require `is "1" or "true" or "yes" or "on"`.
`perf?reset=false` therefore clears the engine counters and `problems?rules=off` returns the rules,
while `guard?on=false` correctly leaves the guard off. The two conventions sit 750 lines apart in
the same switch. They are not arbitrary - the six are per-request toggles the docs only spell as
`=1`, the two are setters of persisted state where false must mean false - but docs/debug-api.md:81
teaches `on=true|false` on this same door, so a caller does meet the word convention before meeting
a toggle.

**Why it matters.** The failure is silent and destructive: the counters are gone by the time anyone notices the flag was
misread. It is also the kind of inconsistency a future product bridge inherits, because route shapes
are deliberately stable.

**Fix.** Lift the existing local Flag helper at :2982-2985 to a private static used by all eight sites,
keeping absent-means-unchanged for guard and settings and absent-means-false for the toggles. Accept
1/true/yes/on and 0/false/no/off and fall back for anything else.

**Size.** 8 sites to one existing helper, about 6 lines. Effort small, risk medium, confidence verified.

#### B11. The page's host-message protocol carries three dead kinds and its only written reference documents a fourth that never existed

**Where.** setTheme: ui/editor/src/bridge.ts:62, :477, :548-550, :1162-1165, :1350 and
src/Xlide.Vbe.Shim/Editor/EditorMessages.cs:43-45 and :522. applyEdit: bridge.ts:61, :1159-1161,
:1572-1597 (no host record exists at all). README: ui/editor/README.md:44-73, the loadDocument row
at :48 and the revision paragraph at :70-73

**What.** Two independent finders reached setTheme from opposite ends; it is one path. A repo-wide grep finds
setTheme only in the page's union member, its handler, the monaco call and the README - nothing in
C# ever constructs SetThemeMessage, and no host code carries a theme at all, so the README documents
pinning the theme against prefers-color-scheme as a capability the host does not have. Because
nothing sends it, `themePinned` is only ever set inside the dead case and the public `isThemePinned`
getter has zero callers. applyEdit is 26 lines of pushEditOperations, echo suppression and revision
adoption with no sender and no host record or serializer entry, so the host could not send it even
if someone wrote the call; the real write path is syncDocument at bridge.ts:1565. The README's
host-to-page table has seven rows, of which loadDocument does not exist anywhere in the page (the
live kind is openDocument at :1115), and covers four of about 39 dispatch cases; its page-to-host
table has four rows against roughly 50 client kinds.

**Why it matters.** applyEdit reads as the mechanism by which the host writes into a document, so the next person
implementing a host write wires to unreachable, untested code that touches undo-stack semantics and
revision authority. Unknown kinds only console.warn at bridge.ts:1341-1344, so a host message
written from the README's loadDocument row fails silently. A confidently specific wrong protocol
table is worse than none.

**Fix.** Delete the setTheme path (both ends, plus themePinned and isThemePinned), the applyEdit path, and
the loadDocument row and revision paragraph. Then either regenerate the two tables from the
ClientMessage and HostMessage unions or cut them to a pointer at bridge.ts:40-100 and :260-310 as
the authority.

**Size.** about 45 lines across bridge.ts, EditorMessages.cs and README.md, plus one AOT serializer registration. Effort small, risk low, confidence verified.

#### B12. AnswerBlockedRequest's only explanation sits above the wrong method and promises a retry the code does not perform, and the wire field that would report it is a constant false

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2070-2084 (the stranded summary), :2085-2091
(EvaluateClaim, which it now decorates), :2239 (AnswerBlockedRequest, undocumented), :2272 (the
wait), :2259 and :2285 (`Retried: false`); src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:744 (the
field)

**What.** A complete summary beginning 'What to say - and do - when the host thread did not answer' is
immediately followed by a second summary and then by EvaluateClaim, so one member carries two
summaries and AnswerBlockedRequest carries none. The first summary ends by saying the dialog is
dismissed and the request retried once; the code waits on work already queued
(`done.Wait(TimeSpan.FromSeconds(3))` under its own comment saying so) and both reply constructions
pass Retried: false. A grep across src, tools, docs, the page and the engine finds nothing that sets
it true or reads it, and it is absent from docs/debug-api.md.

**Why it matters.** The blocked-request policy is the hardest thing in that file to reason about and its explanation is
attached to the wrong method and is false where it is read. A client checking `retried` to see
whether its request was re-run gets a field that is always false and undocumented.

**Fix.** Move the summary above AnswerBlockedRequest, correct 'retried once' to what the code does, and
either set Retried honestly or drop it from DebugBlockedReply.

**Size.** 1 doc block moved and corrected, 1 dead wire field removed. Effort small, risk low, confidence verified.

#### B13. Five duplicated shapes in the debug-api file, each small, all fixable with a local helper the file simply never grew

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs - the on-host error and ok replies (21 sites from
:2499 to :3339, 6 more at :2916, :2928, :2945, :3261, :3315, :3321, against the pool-side DebugError
at :33); the marshal-and-wait scaffold (:965-987 in journal against :1858-1882 in the main
dispatch); the page-eval reply shaping (:1058, :1088, :1107, :1731); the wait-for-workspace loop
(:1532-1543 against :1579-1589); the quantile block (:1411-1422 against :1500-1512, with a third
copy at src/Xlide.Vbe.Shim/Diagnostics/EngineCounters.cs:122-123 and two more at
tools/harness/xlide-api.mjs:830 and :871)

**What.** AnswerDebugRequestOnHost returns string rather than DebugReply, so it cannot use DebugError and
spells the serializer call out 21 times; the success reply is spelled out 6 more. The journal route
re-implements the marshal-and-wait (same event, same callback shape, same hardcoded 3s) and is the
one host crossing missing `PerfCounters.Marshal`, despite the comment at :1880 saying every
marshaled request doubles as a probe of host responsiveness. Four page-script routes repeat the
identical error-or-DebugEvalReply five-liner over the same RunPageScript tuple. Two routes repeat
the same 150ms probe loop for `window.xlideBridge.workspace`. Two routes plus EngineCounters plus
the harness carry five copies of the same p50/p95 expression.

**Why it matters.** Individually trivial; together they are why the on-host switch is 1,058 lines, and each one is a
place where a copy can drift. The two that already cost something: the journal's host crossing is
invisible to the perf and stats routes, and one definition of p95 living in five places means the
door can report two different quantile conventions.

**Fix.** Add HostError(string) and HostOk(int) beside DebugError; one OnHost(request) helper holding the
event, callback, catch, deadline and the PerfCounters.Marshal sample, with the main dispatch keeping
the dialog attribution around it; PageReply(tuple) next to RunPageScript;
WaitForWorkspace(budgetMs); and one Quantiles(what, samples, detail). Roughly 90 lines out and five
conventions stated once.

**Size.** about 90 lines removed across 5 shapes; 33 call sites. Effort small, risk low, confidence verified.

#### B14. Seven engine round-trip handlers are the same capture-guard-deadline-map-marshal block, and the drift has already started

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:4363-4406, :4413-4454, :4457-4499, :4506-4552, :4558-4597,
:4603-4642 and :4650-4691; the byte-identical guard line at :4369, :4419, :4463, :4512, :4564, :4609

**What.** Each handler captures surface/module/source, guards with an empty reply, then Task.Run { using
deadline; await analysis.XAsync; map; log; catch; surface.RunOnHostThread(reply) }. Only the
analysis method, the projection, the empty value, the log verb and the deadline differ.
OnCodeActionsRequested already diverges: its guard omits `source` and its success log sits inside
`if (actions.Length > 0)` while the others log unconditionally.

**Why it matters.** Any change to the shared policy - the deadline, the failure-log shape, the empty-reply contract, or
cancelling a request superseded by a newer one - has to be made seven times, and a handler that
misses one drifts silently because they all answer empty on failure. Note what is NOT in scope: the
other six deadline sites (outline, semantic tokens, rename, module rename, navigation, live
analysis) diverge materially in shape and must not be forced into the same helper.

**Fix.** One generic AnswerFromEngine<T>(requestId, empty, deadline, verb, ask, reply) on the session doing
the capture, guard, Task.Run, deadline, catch, log and host-thread reply. Closed generic
instantiations only, so NativeAOT is unaffected.

**Size.** 60-80 boilerplate lines of about 260, across 7 handlers. Effort medium, risk low, confidence verified.

#### B15. bridge.ts carries ten hand-copied pending-request tables where one helper would do

**Where.** ui/editor/src/bridge.ts:405-462 (ten maps), :464-473 (ten counters), twelve request bodies at :710,
:728, :746, :764, :782, :824, :843, :862, :896, :910, :924, :1045, and ten resolve blocks at :1200,
:1209, :1218, :1227, :1236, :1245, :1258, :1267, :1277, :1290

**What.** Each request body is the same eight lines - allocate an id, setTimeout that deletes from the map and
resolves a fallback, map.set, transport.post - and each resolve block is the same four - get,
delete, clearTimeout, resolve. Twelve methods over ten tables (the three rename methods share one).
The timeouts are NOT uniform: 2000 at five sites, 8000 for outline, navigation and semantic tokens,
30000 for the three rename methods, 120000 for sync with a comment saying why, and the fallbacks
differ (empty array, null, an error object, a refused rename answer).

**Why it matters.** Roughly 400 lines of a 2,356-line file carrying one idea. A new host round trip costs four
coordinated edits in four separate regions, and a missed clearTimeout or a wrong map in a resolve
block is a leaked timer or a promise that resolves empty after the budget - both of which read as a
slow host rather than a bug.

**Fix.** One RequestTable with `ask<T>(kind, body, fallback, timeoutMs)` and `settle(kind, id, value)`. Keep
timeout and fallback as per-call arguments; do not collapse them to one constant, since four
distinct budgets are deliberate.

**Size.** about 400 lines carrying one pattern. Effort medium, risk medium, confidence verified.

#### B16. WatchReader is a transcription of LocalsReader, including the hardening that came out of the 2026-08-05 crash

**Where.** src/Xlide.Vbe.Shim/Editor/WatchReader.cs:25-33, :41-50, :52-92, :95-185, :236-249 against
src/Xlide.Vbe.Shim/Editor/LocalsReader.cs:37-49, :57-66, :68-110, :119-249, :303-316; both
constructed on one thread at src/Xlide.Vbe.Shim/Editor/GhostReaderThread.cs:88, :95, :103

**What.** The two Dispose bodies are byte-identical (verified by diff, no output). Connect differs only in a
log prefix and two comment lines: same CoCreateInstance of the automation class, same ComHandle
ownership, same ElementFromHandle, same CreateTrueCondition. Field blocks, Create factories and the
5000ms backoff with its first-of-streak log all match. The Read loops share the FindAll -> GetLength
-> GetElement -> ControlType -> Name walk with the same per-element try/catch and poisoned counter,
diverging only in accepted control types, LocalsReader's stage string and Edit/Pane branch, and the
row parser. WatchReader's own comments say 'the same manner as the Locals reader' three times. Both
readers create their own CUIAutomation client inside one CoInitializeEx(MTA).

**Why it matters.** This is the path that killed the host on 2026-08-05, and its fix - the sized variant out-parameter,
the separate thread, the per-element try/catch, the backoff - now exists twice. The next correction
to the accessibility walk has to be made in both, and a maintainer who fixes the file in front of
them leaves the other panel on the old behaviour.

**Fix.** One base or one parameterised class holding Create/Connect/Dispose plus a ForEachListItem callback
that owns the FindAll walk, the per-element catch, the poisoned count and the backoff; each reader
keeps only its control-type rule and its parser. Optionally hoist the automation client and true
condition into GhostReaderThread, which already owns the apartment - that saves one CoCreateInstance
at start, not anything per tick.

**Size.** 74 byte-identical lines plus a ~55-line loop differing in three places. Effort medium, risk medium, confidence verified.

#### B17. Six dialogs hand-roll the same modal plumbing, five re-declare CSS a general .modal-backdrop rule already provides, and none of the six traps focus

**Where.** ui/editor/src/helpdialog.ts:76-90 and :195-215; sponsordialog.ts:59-73 and :169-190;
settingsdialog.ts:408-429; referencesdialog.ts:142-163; syncdialog.ts:739-761; shell.ts:796,
:819-826, :950; ui/editor/src/styles.css:1877-1886 (.modal-backdrop, used twice) against :1190,
:1360, :1660, :2798, :2922

**What.** Each dialog creates its own backdrop and card, sets role and aria-modal, installs a captured Escape
handler and a `event.target === backdrop` mousedown, and defines the same three-line dismiss. The
general .modal-backdrop/.modal-card rules exist and only shell.ts uses them; the other five repeat
position, inset, rgba(0,0,0,0.35), flex, alignment and centring verbatim under their own ids,
differing only in padding-top and z-index. A grep for Tab handling across the five dialog modules
returns only an unrelated comment.

**Why it matters.** Six copies means the thing a modal should do gets written zero times: six cards carry
aria-modal="true" and Tab walks straight out of every one of them into the Monaco surface behind.
That is a WCAG-level defect a shared helper fixes once. (The z-index difference is correct as it
stands - the close-confirm at 95 should sit above the dialogs at 90 - and the padding-top values are
defensible per card height.)

**Fix.** Promote the shell.ts pattern to openModal({ id, label, build }) in one module: backdrop and card,
role and aria-modal, captured Escape, backdrop mousedown, a Tab trap inside the card, focus restore
on dismiss, returning dismiss. Point the five dialogs at it and delete the five #*-backdrop rules,
keeping only genuinely per-dialog width and padding.

**Size.** 6 copies of ~20 lines of TS plus ~50 lines of CSS; one focus trap instead of six missing ones. Effort medium, risk low, confidence verified.

#### B18. The whole-module read and the component walk are each reimplemented three times, once in production and twice behind the debug api

**Where.** ReadSource copies at src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:2534-2540, :4073-4083 and :5005-5012
against the helper at src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:199-216. Component walk at
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3057-3070 and :3110-3122 and in production at
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5919-5926

**What.** Three sites hand-roll FindComponent -> GetObject("CodeModule") -> GetInt32("CountOfLines") ->
GetStringIndexed("Lines", 1, count) with their own empty-module guard, which is exactly what
ProjectReader.ReadSource owns and documents ('asking an empty module for line one raises').
Separately, three sites walk VBComponents with their own count loop and their own copy of the
scratch-module filter, written positively in one and negatively in another. The two debug-api copies
also differ in failure handling in a way that matters: the projects route's per-component read sits
inside the try that wraps the whole project entry, so one unreadable component drops the ENTIRE
project from the reply, while the project route skips the component and returns the rest.

**Why it matters.** The scratch-module exclusion is a fixture-correctness property (a fixture that counts it counts
wrong) now stated three times in two layers, and the empty-module edge case is re-guarded by hand at
four sites. All the copies happen to be right today, so this is maintenance cost plus one genuine
inconsistency in the projects route.

**Fix.** Replace the three inline reads with ProjectReader.ReadSource, deciding once whether a failed module
read is Log.Error (as the helper does) or Log.Verbose (as all three sites do) and whether 'no
CodeModule' must stay distinguishable from 'the read threw' - two of the three sites branch on that
today. For the walk, add one component-enumeration helper holding the count loop, name read, scratch
filter and per-entry catch; it must live outside the #if DEBUG region so the production walk can
share it.

**Size.** ~9 lines from the read copies, ~12 from the walk copies, plus one behaviour inconsistency fixed. Effort medium, risk medium, confidence verified.

#### B19. The pool-side route switch has no default, so a route whose argument guard fails is marshalled to the host thread it exists to avoid

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:516 (switch opens), :1837 (closes, no default),
the 11 guarded cases at :909, :1062, :1092, :1111, :1218, :1297, :1425, :1515, :1675, :1735, :1783,
and the fallthrough at :1839-1878; the on-host default at :3323-3344; the design intent at :503-513

**What.** Eleven cases carry `when` guards. When a guard does not hold, control leaves the switch, takes
`_editorSurface`, and marshals to the host thread. In the ordinary case the on-host default already
answers accurately - it names both possibilities and lists the arguments given - so the harm is
narrower than it looks. What is unconditional: assert, dismiss and guard exist precisely so a caller
can get an answer while something is standing, and they lose that property the moment their guard
does not hold; and with no surface up, any argument-rejected pool route answers 'the surface is not
up yet' instead of naming the argument.

**Why it matters.** dismiss with an empty button is the request a caller makes while a modal stands, and it is the one
that stops being answerable without the host. The file already paid for this class of confusion
once: the on-host default carries a six-line apology written on 2026-08-07 after caret?line=-1 was
answered 'unknown route caret'.

**Fix.** Add a pool-side default returning the same message the on-host default already composes, or a
pre-switch guard limited to the host-free routes. Do NOT build a pre-switch
route-to-required-arguments table as first proposed: `layout` appears twice, guarded at :1515 and
plain at :1549, so at least one `when` pair is overload dispatch and a table would swallow the plain
case.

**Size.** about 10 lines; the change is behavioural, not size. Effort small, risk medium, confidence verified.

#### B20. The import path reads the whole module twice before writing it, and the second read is provably the same text as the first

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1903 (the first read), :1915 (TryWriteLineDiff), :1928 (the
second read), :1982 (where the first is already used); TryWriteLineDiff at :2110-2190; the only
keepEveryCharacter:true caller at :395

**What.** On the import path wasHoldingBefore is read at :1903 and a second ReadSource runs at :1928 inside
`if (!wroteDiff)`. TryWriteLineDiff has exactly one `return false`, at :2142, guarded on window size
and reached before any COM call - its first COM touch is CountOfLines at :2154 and its first
mutation is DeleteLines at :2161 - and its failure path rethrows. So when wroteDiff is false the
module is untouched between the two reads and they return identical text.

**Why it matters.** The code's own measurement two lines above says what a read costs: 3ms of a 1,037ms write at 1,002
lines, 66ms of a 12,594ms write at 40,002. A fifty-module import pays that per module in an
operation the user is waiting on. Half a percent of the accompanying write, so small, but free.

**Fix.** `var wasHolding = wasHoldingBefore ?? ProjectReader.ReadSource(found);`. Note wasHoldingBefore is
not spare - it already backs the character-loss restore at :1982 - so this is a read elision only
and changes no restore semantics.

**Size.** 1 line; 3ms per 1,002-line module, 66ms per 40,002-line module, per import. Effort small, risk low, confidence verified.

#### B21. The engine answers a protocol version on initialize and the shim discards it

**Where.** src/Xlide.Vbe.Shim/Engine/EngineClient.cs:114 (the last statement of the connect path);
engine/src/dispatcher.ts:272-274; the failure it would catch at tools/verify.ps1:94-102

**What.** `await CallAsync("initialize", ...)` awaits and does not assign, and nothing in src reads an engine
or protocol field off any response. The engine returns `{ engine: 'xlide', protocol: 1 }`. That
literal 1 has never been bumped, and a grep for any version or build stamp across src and engine/src
returns nothing, so there is no build identity on the wire at all.

**Why it matters.** verify.ps1's freshness step exists because an engine change can be built, tested, committed and
published while the running engine is hours old and refuses every new method as unknown - which
happened on 2026-08-06 and surfaced only through a live session's log. The gate's timestamp check
covers the developer's machine; an installed user with a mismatched engine gets MethodNotFound per
feature with nothing naming the cause. Pairs with item 7: same failure, opposite end.

**Fix.** Read the result and log at Warn (or refuse to start) on mismatch, naming both numbers - but note
that comparing protocol only helps once the engine starts bumping it. Answering a build stamp from
the packaged exe catches the 2026-08-06 case directly. Additive to an existing response, not a new
route.

**Size.** 1 discarded response; no build identity anywhere on the wire. Effort small, risk low, confidence verified.

#### B22. Comments and doc blocks that describe something other than the code beneath them, in the three files whose comments are most load-bearing

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs - ten members carrying two summaries, at :131, :1247,
:3720, :4007, :4974, :5720, :6420, :6875, :6932, :7224.
src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs:186-187 (the patience budget).
ui/editor/src/devsurface.ts:17-19 (a few hundred bytes)

**What.** Ten times in AddInSession.cs a summary sits immediately above a member it does not describe, while
the member it does describe has none: the block about publishing findings across all modules
decorates DropFindingsFor, the one about naming a workbook decorates ContentKey, the one about
finding a component across every open project decorates ActivePaneOwner, the one about recomputing
surface placement decorates PlaceSurfaceFast, and the poll-interval pair leaves
EmptyWorkspacePollMilliseconds bare. In XlideAddIn.cs one comment says the enabled-tick wait 'does
not spend the patience budget above' when the comment fourteen lines earlier correctly says there is
no budget (and _watchdogTicks, the counter that looks like one, only feeds a Log.Verbose line the
standing verbose-logging directive protects - keep it). In devsurface.ts the comment justifying
shipping the module in Release calls it 'a few hundred bytes of read-only reporting'; the file is
63,412 bytes and 1,483 lines, it carries the act() driving surface as well as reporting, and it was
already 11,724 bytes in the commit that wrote the comment.

**Why it matters.** These comments are the map of a 7,741-line file and an add-in whose watchdog the codebase calls its
single most important line, and they carry the measurements, dates and failure histories that
justify the code. Ten members currently lie about what they do and ten more have lost their
explanation; a maintainer reading the watchdog goes looking for a threshold that was deliberately
removed; and anyone revisiting the ship-devsurface-in-Release tradeoff reads a number two orders of
magnitude off and closes the question. The decision itself is sound - a door only present in Debug
is a door nobody trusts - but it should be defended on the real figure, including the always-on
PerformanceObserver at devsurface.ts:220.

**Fix.** Move the ten summaries onto the members they describe (turn the second one at :6875 into a
<returns>); fix the one stale sentence in XlideAddIn.cs; restate the devsurface comment with the
real source size, labelled as source bytes since the minified contribution is unmeasured. No code
lines change. Do not expect a compiler guard: there is no diagnostic for a duplicated <summary>, so
GenerateDocumentationFile would not catch a recurrence.

**Size.** 12 doc blocks, 0 code lines. Effort small, risk low, confidence verified.

#### B23. PublishModules walks Excel's Workbooks collection once per distinct workbook on every poll tick, and the change-key that would make it free is computed after that work

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3934 (called every tick), :5641 (the walk), :5646-5649 (the
key, built after), :5574 and :5587 (per-call memo), :5681-5713 (WorkbookSaved) duplicated at
:6589-6601 (SaveWorkbookOf), :5773-5841 (ReadOpenModules); intervals at :3703-3707, :3726, :3728 and
:2244

**What.** The dirty array is built before the change-key is computed, so the cheap exit never avoids the walk.
DirtyOf memoises in a dictionary local to the call, so WorkbookSaved runs once per distinct workbook
per tick, and each run is GetObject("Workbooks") + Count + a GetItem/Name pair per workbook until it
matches + Saved. ReadOpenModules is about eight invokes per open pane. Structurally that is 8P +
D*(2+2W) IDispatch invokes before discovering nothing changed, against roughly ten for everything
else on the tick. WorkbookSaved and SaveWorkbookOf are the same
Workbooks/Count/GetItem/match-on-Name loop written twice, differing only in reading Saved versus
invoking Save.

**Why it matters.** The comment at :3930-3933 states the unchanged cost as 'a read', which understates it by an order of
magnitude, so anyone tuning the tick reasons from a wrong number. Calibrate the urgency, though: the
interval is 0 unless a resync burst, a debug episode, immediate watching or an empty workspace is in
play; the 16ms close-resync is about six ticks over a tenth of a second, and the sustained worst
case is 150ms while stopped. So this is a wrong-comment-plus-structure finding, not a live stall.

**Fix.** Add a PerfCounters stamp around PublishModules first - PlacementFull and PlacementFast are stamped
and publish is not, which is why the cost is asserted rather than measured. Then read Workbooks once
per call into a name -> Saved map, which also collapses the duplicated walk into one helper both
WorkbookSaved and SaveWorkbookOf use, and move the change-key ahead of the dirty array where it can
actually short-circuit.

**Size.** structurally 8P + D*(2+2W) invokes per tick against ~10 for the rest; unmeasured. Effort medium, risk low, confidence supported.

#### B24. The test harness has no shared layer, so the same scaffolding is written four, nine and four times

**Where.** engine/test/positions.mjs:23-100, freshness.mjs:22-99, language.mjs:26-112, smoke.mjs:7-80 (no
shared module exists in engine/test). tools/harness: the identical `const check = (what, ok, detail)
=>` in colouring, debugger-features, import-guard, language-features, menu-bar, rename-features,
settings-bite, step-into-features, write-rollback; the verdict epilogue in analysis-freshness:344,
colouring:200, com-leak:432, format-positions:435, immediate-watch:197, import-guard:108,
settings-bite:152, three-copies:246; local `wait` consts in nine suites shadowing xlide-api.mjs:110;
the pane-close/component-remove teardown pair at colouring:196-197, import-guard:102-103,
settings-bite:148-149. tools/New-DebugFixture.ps1:116, New-LanguageFixture.ps1:193,
New-PerfFixture.ps1:107, New-RenameFixture.ps1:211

**What.** Each engine test carries its own spawn, stderr drain, listening promise with the same 30s reject,
named-pipe connect, newline framer, call() with timeout and check() reporter - measured at 0.85 to
0.997 similarity - and they have already drifted: smoke and language accept --exe so they can run
against the packaged executable, positions and freshness hardcode process.execPath with engine.cjs
and cannot be pointed at what the add-in launches. The mjs suites repeat the reporter in three
shapes (two suites use different variable names), the verdict line verify.ps1:439 greps for, and the
scratch-module teardown. The four fixture generators repeat the same three-phase driver verbatim,
including the BOM-free WriteAllText whose comment exists because PowerShell 5.1's -Encoding utf8
writes a BOM that JSON.parse refuses - that lesson is now recorded four times.

**Why it matters.** About 600 duplicated lines. The engine copies mean a protocol change is four edits and two of the
tests cannot exercise the packaged exe at all. The gate's ability to read a verdict rests on nine
hand-copied console.log formats staying byte-identical (all three shapes still parse today, so this
is latent, not live). A fifth fixture is a copy-paste and a change to how build-fixture.mjs is
invoked is four edits with nothing to catch a miss.

**Fix.** engine/test/harness.mjs exporting startEngine({ label, useExe }) -> { call, stop } plus the
reporter, which gives --exe to all four for free. In xlide-api.mjs, which already absorbed
wait/waitFor/waitUntilStable, add reporter() -> { check, done } printing the one line verify.ps1
parses, and scratchModule(api, project, name) returning a disposable that owns the teardown; delete
the nine local wait consts. A shared tools/FixtureDriver.ps1 taking
-Path/-Modules/-SheetCode/-OpenAtEnd, leaving each generator as its VBA bodies plus one call.

**Size.** about 600 duplicated lines across 17 files. Effort medium, risk low, confidence verified.

#### B25. Nine PowerShell scripts rebuild the debug-api discovery path by hand and six of them pick the first Excel they find

**Where.** tools/harness/Test-DebugApi.ps1:17-20, Test-SplitWorkspace.ps1:20-23,
Test-DiscardProblems.ps1:33-36, Test-Churn.ps1:31-34, Get-Shot.ps1:44-47,
Test-ObjectBrowser.ps1:183, Open-VbeIn.ps1:103, tools/page.ps1:70, tools/verify.ps1:352; contrast
tools/harness/xlide-api.mjs:179-217 and :224, and tools/page.ps1:66-90

**What.** Nine sites build the same `$env:LOCALAPPDATA\xlide_vbide\debug-api-<pid>.json` path and the same
base URL from port and token. Six take the pid from `Get-Process EXCEL | Select-Object -First 1`.
None of the six proves the discovery file belongs to a live listener. The node client and page.ps1
both already do it properly: enumerate the directory, prove each candidate with a /state call, and
throw with the list rather than guessing, because guessing which Excel to drive is how a test writes
into the wrong workbook.

**Why it matters.** Four of the six first-Excel-wins scripts are run by the gate (Test-DebugApi, Test-SplitWorkspace,
Test-DiscardProblems, Test-Churn), so with two Excels open the live gate can assert against the
wrong session. The lesson was learned once on the node side and not carried across. Any change to
the discovery file's name, location or fields is nine edits.

**Fix.** tools/harness/XlideApi.psm1 exporting Get-XlideApi [-ProcessId] [-Workbook] mirroring
discover()/open(): enumerate, prove liveness with /state, throw with the candidate list. Convert the
nine call sites. Route shapes untouched - this is only how the base URL is found.

**Size.** 9 hand-rolled copies; 6 of them pick the wrong Excel when two are open. Effort medium, risk low, confidence verified.

#### B26. Twenty-five unreferenced constants and P/Invokes in the interop surface, four of them describing an object model the shim does not implement

**Where.** src/Xlide.Vbe.Shim/Com/HResult.cs:12, :13, :16, :17, :20, :24, :28, :31, :33, :34;
src/Xlide.Vbe.Shim/Interop/Win32.cs:114, :129, :132, :136, :140, :141, :143, :154, :227, :230;
src/Xlide.Vbe.Shim/Interop/Win32.Events.cs:34, :70, :201, :203, :206

**What.** Ten of HResult's eighteen constants have no reference anywhere in src, tests or tools - confirmed by
qualified grep, bare-identifier grep and by checking no `using static` path exists - and no call
site uses the raw literals either. Four of the ten (OLE_S_USEREG, OLEOBJ_S_INVALIDVERB,
OLE_E_ADVISENOTSUPPORTED, OLE_E_NOCONNECTION) are the vocabulary of an in-place-activated embedded
object, and there is no IOleObject implementation anywhere. In Win32, SetParent and
GetForegroundWindow have no caller, and thirteen message and style constants appear only at their
own declarations (WsVisible is not folded into the composite WsOverlappedWindow either).

**Why it matters.** Readability only - consts are compile-time and an unreferenced LibraryImport partial is never rooted
by ILC, so nothing ships. But HResult.cs's summary claims these are the values this server returns
or inspects, and SetParent sitting in the interop surface of a product whose entire placement
strategy is a layered overlay that deliberately does not reparent the host's windows suggests
reparenting is an available move. GwlStyle with WsMaximize, and WmSysCommand with ScClose, read the
same way.

**Fix.** Delete them. Anything genuinely reserved for imminent work gets a one-line comment naming that work,
so the next dead-code pass does not re-derive the same answer.

**Size.** 25 declarations, about 37 lines across three files, 0 bytes shipped. Effort small, risk low, confidence verified.

#### B27. Native declarations and window helpers written outside Interop, duplicating what Interop already owns

**Where.** src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:488-495 (a private Rect), :463-465 (GetWindowRect),
:478-479 (SelectObject), :481-483 (DeleteObject) against src/Xlide.Vbe.Shim/Interop/Win32.cs:6-13,
:323, :326 and Win32.Events.cs:88. src/Xlide.Vbe.Shim/Diagnostics/DialogWatch.cs:215-229,
src/Xlide.Vbe.Shim/Editor/CodePaneTracker.cs:910-924 and
src/Xlide.Vbe.Shim/Editor/HostChrome.cs:200-206 (three private copies of the same window-text and
class reads)

**What.** DebugCapture re-declares a Rect identical in layout to Interop.Win32.Rect and three imports the same
assembly already binds; GetDC, ReleaseDC, PrintWindow, CreateCompatibleDC, DeleteDC and
CreateDIBSection are genuinely new and belong there. Separately, the same GetClassName /
GetWindowText wrapper is written three times with three buffer capacities (64, 256, 512) and two
empty conventions (string.Empty in two, null in HostChrome).

**Why it matters.** Modest. The Rect is private and nested so nothing can pick the wrong one, and DebugServer.cs is
entirely inside #if DEBUG so nothing ships - this is the Interop folder's
one-declaration-per-entry-point rule quietly not holding, which means the next person adding a GDI
or window-inspection call has two or three places to look and no way to tell which is canonical. The
three window-text copies invite a fourth.

**Fix.** Point DebugCapture at Interop.Win32.Rect and the three existing imports (GetWindowRect differs in
shape, `Rect*` against `out Rect`; the pointer form works from a local with &rect, as
CodePaneTracker.ReadBounds already does), keeping only the six new ones. Move ReadClassName and
ReadWindowText onto Interop.Win32 beside the imports they wrap, reconciling the three capacities and
the two empty conventions in the move.

**Size.** 1 duplicate struct, 3 duplicate imports, 3 copies of 2 helpers, about 35 lines. Effort small, risk low, confidence verified.

#### B28. VbeCommands.Describe is unreachable, and the doc comment that sends a maintainer to it is the only reference to its name

**Where.** src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:322-361 (the method), :29-30 (the doc comment pointing at
it), :23 (PreferredBars); the surviving instrument at src/Xlide.Vbe.Shim/Editor/VbeMenus.cs:401,
driven from src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2652 and documented at
docs/debug-api.md:86

**What.** A repo-wide grep finds no call site for VbeCommands.Describe; the only reference to the name is the
Command class's own doc comment telling the reader to re-run it if a host build ever disagrees about
command identifiers. The method walks every CommandBar and logs each control's Id and Caption, and
its own summary admits it is not called during normal operation. The menus route now answers the
same question through VbeMenus.Describe, and reports suppression as well as identity.

**Why it matters.** Forty lines of live COM enumeration nothing can reach, presented by a doc comment as the recovery
procedure. One caveat that changes the fix: VbeMenus.Describe reaches controls only from the menu
bar (FindMenuBar plus a position chain), while VbeCommands.Describe enumerates every CommandBar -
which is where VbeCommands.Find looks, at the Debug and Edit bars. Deleting it loses that dump.

**Fix.** Either delete :322-361 and repoint the :29 comment at the menus route, accepting that the
whole-CommandBars dump is gone, or add a bar selector as a new field on the menus door first and
then delete.

**Size.** 40 lines removable. Effort small, risk low, confidence verified.

#### B29. The split-tree prune and same-direction absorb are written three times, and only the copy in the module built to be tested is tested

**Where.** ui/editor/src/docktree.ts:78-105 and :128-162 (the tested original, covered by
ui/editor/test/docktree.mjs:60-186); ui/editor/src/workspace.ts:29 (imports only resizeAt), :893-929
and :946-970; ui/editor/src/paneldocks.ts:21 (imports prune), :803 (uses it), :308-338 (a local
const prune shadowing the import inside pruneUnknown)

**What.** docktree.ts's header says it lives there, separately and purely, so the arithmetic can be tested by
calling it. workspace.ts re-implements prune and the splitBeside replace, including the identical
same-direction absorb arithmetic. paneldocks.ts imports prune and then declares a local one whose
body is the imported function verbatim wrapped around a tab filter, shadowing the import inside that
method. The only divergence is at workspace.ts:928, where the absorb path returns sizes
unrenormalised - arithmetically a no-op while sizes partition one, so a latent robustness difference
rather than an observed layout bug.

**Why it matters.** The bug class docktree's own header names - removing a pane empties a group, which leaves a split
with one child, which must collapse, and sizes must stay a partition of one - is live in two
untested copies. The shadowed name is the sharper problem: a reader who sees prune in pruneUnknown
and follows the import at line 21 reads the wrong function.

**Fix.** Two separately sized jobs. Cheap and low risk: split the tab filter from the collapse in
paneldocks.pruneUnknown - map the tree's groups through the known/placed filter, then call the
imported prune - and remove the shadow. Expensive: having workspace.ts adopt docktree's TreeNode
requires moving the DOM element that LayoutNode carries on every split node into a side map, which
the pure tree has no place for. Do the first now and treat the second as its own change.

**Size.** 3 implementations of 2 functions, about 80 duplicated lines, 2 of 3 untested. Effort large, risk medium, confidence verified.

#### B30. docs/testing.md describes the live gate as the PowerShell probes and never mentions the ten node suites that are most of it

**Where.** docs/testing.md:15, :20, :24, :67; tools/verify.ps1:369 (four Test-*.ps1 of twelve), :395-443 (the
node suites), :464-478 (com-leak)

**What.** The document presents live probes as tools\harness\Test-*.ps1 gated behind -Live. The gate runs four
of the twelve such files; the bulk of -Live is the step whose own comment opens 'THE NODE SUITES,
which existed and passed and which nothing ran', running import-guard, immediate-watch,
analysis-freshness, menu-bar, module-sync twice, format-positions, three-copies, colouring and
settings-bite. The document names none of them; its only .mjs mentions are the page-probe family and
com-leak. It also presents seam checks as one of four kinds, and both Test-Seam implementations are
in files the gate never invokes (item 9).

**Why it matters.** This is the document that answers 'which kind of check should I write'. A new suite written from it
arrives in the wrong language, in a file the gate does not run, with a verdict line verify.ps1's
parser does not recognise.

**Fix.** Add a kind for the node suites - what they are, that they must print `N passed, M failed` because
verify.ps1 parses that, that they import open/waitFor from xlide-api.mjs and must restore the
fixture in a finally - and correct line 24 to name the four Test-*.ps1 the gate runs. Fix or remove
the seam-check paragraph depending on what happens to item 9.

**Size.** 10 of the 14 live-gate suites undocumented. Effort small, risk low, confidence verified.

#### B31. The engine routes and documents module/didClose, which nothing has ever sent, and the dispatcher says so thirty lines below

**Where.** engine/src/dispatcher.ts:286-288 (the case), :448-457 (the comment explaining it is never sent and
that closeProject releases instead); engine/README.md:36

**What.** A whole-repo grep for didClose outside node_modules and engine/dist returns four hits: the case, the
two comment lines, and a README protocol-table row advertising it as live protocol. No .cs, .ts,
.mjs or .ps1 sends it.

**Why it matters.** The next person wiring tab-close plumbing finds a documented hook that duplicates what closeProject
already does correctly and at the right moment. Two ways to forget a document, one of them
unreachable and untested.

**Fix.** Delete the case and the README row and move the closeProject comment's explanation up to where a
reader looks for it. If it is kept as a future hook, say so in the README row rather than describing
it as protocol.

**Size.** 3 lines plus one README row. Effort small, risk low, confidence verified.

#### B32. tools/page.ps1 and tools/Update-Page.ps1 are the same page loop, and the operational guide points at the thinner one

**Where.** tools/Update-Page.ps1:38-65; tools/page.ps1:16-28 and :53-56; docs/driving-excel.md:134 and :139
against docs/status.md:57 and README.md:100

**What.** Both compute the same artifacts\publish\Xlide.Vbe.Shim\<config>_win-x64\ui\editor\dist target, build
the page unless told not to, copy into it and reload. page.ps1 is the superset: -Watch, -Reset,
-NoTypecheck, -NoDeploy. Nothing in verify.ps1, dev.ps1 or CI runs either. docs/status.md and
README.md name page.ps1; the driving guide's what-to-run-after-changing-what table names
Update-Page.ps1.

**Why it matters.** docs/driving-excel.md is the operational guide under a standing rule to stay current with the
harness, and its change-to-command table sends the reader to the loop without the typecheck and
without -Watch. Two scripts computing the same publish path also means a layout change fixes one
loop and breaks the other. (The reload itself is fine in both: Update-Page.ps1 goes through
reload-page.mjs, which discovers and reloads every live editor.)

**Fix.** Delete tools/Update-Page.ps1 and repoint docs/driving-excel.md:134 and :139 at tools\page.ps1,
mapping -NoBuild to -NoTypecheck/-NoDeploy. If a thin scripted form is wanted, make it a flag
combination rather than a second file.

**Size.** 66-line script duplicating a 200-line one; two stale doc rows. Effort small, risk low, confidence verified.

#### B33. surface-walk.mjs and four doc examples require TwinFixture.xlsm, and no generator makes it

**Where.** tools/harness/surface-walk.mjs:15-21; docs/driving-excel.md:346, :371, :372, :618; tools/ holds
New-DebugFixture.ps1, New-LanguageFixture.ps1, New-PerfFixture.ps1 and New-RenameFixture.ps1 and no
twin; .gitignore:2 and :55

**What.** The suite's header says a run that never holds two modules of the same name passes every label check
vacuously, and that this is how the pane route's dropped project argument was found; it names a
two-workbook Start-Excel line with RenameFixture.xlsm and TwinFixture.xlsm. The file exists on this
machine under artifacts/fixtures, but artifacts is gitignored and git ls-files artifacts is empty,
so no clone can obtain or reproduce it. Run without a twin the suite still exits zero reporting
collision=0.

**Why it matters.** The one suite written to catch cross-workbook ambiguity - the class of defect that produced the
dropped project argument - is unreproducible for anyone who does not already have the file, and its
failure mode when the fixture is missing is the vacuous pass its own header warns about.

**Fix.** Add tools/New-TwinFixture.ps1 using the same three-phase driver, producing a workbook whose module
names deliberately collide with RenameFixture.xlsm's. If the twin is meant to be a second copy of
the rename fixture, say that in surface-walk.mjs's header and give New-RenameFixture.ps1 a -Path
default that makes the copy obvious.

**Size.** 1 suite and 4 doc examples resting on an untracked file. Effort small, risk low, confidence verified.

#### B34. A literal form feed byte sits where a backslash-f belongs in the driving guide's two-workbook command, so the sample cannot be copied and run

**Where.** docs/driving-excel.md:618 and tools/harness/surface-walk.mjs:21, two occurrences each

**What.** The raw bytes are `artifacts` 0x0C `ixtures`, twice per line - the `\f` of `artifacts\fixtures`
interpreted as a form feed. A scan of every .md, .mjs, .ps1, .cs, .ts, .yml and .json outside .git,
node_modules, dist, obj, bin and artifacts finds 0x0C in these two files and nowhere else.

**Why it matters.** docs/driving-excel.md is under a standing rule that its samples are meant to be run rather than
read, and this one resolves to a path that does not exist. It renders and greps as
`artifactsixtures`, so a reader searching for the fixtures directory will not find the line, and
whoever fixes the doc will not know the same bytes are in the suite header.

**Fix.** Replace the 0x0C with a literal backslash in both files, and single-quote the path wherever it is
being emitted from.

**Size.** 4 bytes across 2 files. Effort small, risk low, confidence verified.

### Themes

- Ownership at the edges of a discipline that is otherwise exact. Every COM wrapper in the product
  goes through TakeWrapper/GiveBackWrapper and the counting works - and then _hostApp is never
  released at teardown, two GetItem loops dispose by hand where a throw skips it, and a Debug-only
  handler list only grows. All three sit outside what com-leak.mjs can observe, because the instrument
  watches deltas during a live session through counted doors, and these three are
  steady-during-session, throw-only, or through an uncounted door. The rule is not tighter review; it
  is that any wrapper cached in a field must be released in Stop(), any wrapper taken in a loop must
  be a using, and any object rooted for a callback must be un-rooted where it fires.
- Duplication concentrated exactly where a helper would have to cross a boundary. The engine round
  trip is copied seven times inside one class, the reply shaping five times inside one switch, the
  modal plumbing six times across six modules, the pending-request table ten times inside one file. In
  every case the reason the helper does not exist is that nobody wanted to add a private member to a
  large file or a new module to a directory. The cheapest structural rule this repo could adopt is
  that the second copy of anything gets the helper, and the helper lives beside the first copy.
- Stated coverage that no runner enforces. audit-routes.mjs treats the harness directory as the driver
  set, verify.ps1 keeps its own hand-written list, CI runs three engine tests by name out of four,
  docs/testing.md describes the wrong kind of suite, and release.ps1 -SkipGate drops the only
  engine-freshness check. Five separate places where the list of what runs is maintained by hand and
  diverges from the list of what exists. One derived list - the gate reads what it runs, the audit
  reads the gate, the doc points at both - would collapse the whole class.
- Comments as the only specification, and no mechanism keeping them true. Ten doc blocks decorate the
  wrong member, the watchdog promises a removed budget, PublishModules is described as costing 'a
  read', devsurface is described as 'a few hundred bytes', AnswerBlockedRequest's explanation says it
  retries when it waits. These comments carry the measurements and failure dates that justify
  decisions, which makes them load-bearing, and nothing checks them - there is not even a compiler
  diagnostic for a duplicated <summary>. Where a comment states a number, the number should come from
  a counter the code emits (PerfCounters has no publish stamp) rather than from a memory of a
  measurement.
- One shape written twice diverges silently when both halves answer plausibly. The page's settings
  payload against the host's parse (syncEngine silently reset), the two flag conventions in one switch
  (perf?reset=false clears the counters), workspace.ts's absorb against docktree's (unrenormalised
  sizes), the two Test-Seam copies. Nothing errors in any of these; each just answers something
  reasonable. The type system could catch the first (share the payload type), a helper the second, a
  test the third - but the common cause is a wire or a convention with two independent implementations
  and no single definition.
- The api-mirrors-the-UI rule catches what the harness cannot. syncEngine is reachable through
  api.settings and unreachable through the dialog, which is why every suite passes over a broken user
  path; module-sync.mjs drives the setting the way no user can. Any route that reaches a state the UI
  cannot reach is a place to look for a UI path that was never wired, not just a documented deviation.

### What this section leaves out

"WHAT I DROPPED (3). (1) 'VbeMenus is fully live' was a verified NEGATIVE result proposing no change
- all four members reachable, the wrench reuses the whole position-chain machinery, nothing to
delete. Recorded here so nobody sweeps it: the leftovers from the v0.6.0 menu retirement are in
VbeCommands, not VbeMenus. Its one real observation is a product item, not cleanup - the Options-522
comment at VbeMenus.cs:150-158 flags Require Variable Declaration as a capability the suppression
dropped with no replacement. (2) 'The editor frame is located four ways' (CodePaneTracker.cs:135,
:625, :681, :884) survived verification but the rationale did not: FindWindowEx with a null parent
and GetTopWindow+GW_HWNDNEXT both walk Z order and both filter to this process and class, so the two
locators cannot disagree. That leaves a 24-line redundant implementation plus three copies of an
8-line prologue, at medium risk, with no behaviour to fix - not worth the change. (3) 'The log's
duplicate-collapse marker hardcodes [info] [host]' (Log.cs:144-145) is real and three lines, but the
marker is always preceded in the file by the same line carrying its true level and origin, nothing
machine-reads the origin, and the storm the collapsing exists for (resize drag, in the frame's
message chain) genuinely is host-thread work. Too small to spend a reviewer on.\n\nMERGES. 52
surviving findings became 34 items. Multi-finding merges: setTheme was found independently from the
shim end and the page end and is one path, now folded with applyEdit and the stale README table into
item 11 (one edit pass, three files). Five separate debug-api duplications became item 13.
ReadSource-inline-three-times and the component-walk duplication became item 18 (same cause: the
session reimplements COM helpers it already has, and the debug-api half cannot share a helper that
lives inside #if DEBUG). HResult and Win32 dead declarations became item 26; DebugCapture's GDI
copies and the three window-text readers became item 27. The engine-test, mjs-suite and
fixture-generator scaffolds became item 24. Misattached summaries, the watchdog patience budget and
the devsurface size comment became item 22.\n\nCAUSE VERSUS INSTANCE, named explicitly. Item 5
(settings shape spelled nine times) is the CAUSE of item 2 (syncEngine dropped); fixing item 2 alone
leaves the next setting to fail identically. Item 8 (audit-routes counts unrun files) is the CAUSE
that hides item 9 (seven unrun scripts) and would have caught item 6 (freshness.mjs) had it looked
at CI. Item 7 and item 21 are the same failure - a stale engine reaching a user - approached from
the release script and from the wire.\n\nWHAT I DID NOT DO. I did not re-derive the finders'
evidence from scratch. I spot-checked the three highest-ranked items against the code directly and
confirmed each: _hostApp is declared at AddInSession.cs:5664, taken at :5685 and :6583, disposed
only in the two catch blocks, and absent from both Stop() (read in full) and Dispose();
bridge.ts:672-681 posts five keys with no syncEngine while EditorSurface.cs:1623-1626 defaults a
missing one to the literal 'xlide'; ForgetScriptHandler at WebView2Surface.cs:143 has no caller and
_pendingScripts has one Add and no Clear. Items 4 through 34 rest on the finders' evidence as
re-checked by the adversary pass, whose corrections I carried into the text - notably the corrected
counts (7 fungible engine handlers not 13, twelve Test-*.ps1 not fifteen, three window-text copies
not two, five timeout budgets not one) and the corrected consequences (no shipped binary cost for
unreferenced consts or LibraryImports, no client-visible error-shape divergence on the debug api, no
z-index inversion between the modals). Item 23 is the only one carrying less than verified
confidence: the per-tick IDispatch arithmetic is counted from call sites, not measured, and a
PerfCounters stamp around PublishModules is the measurement that would settle whether the walk
matters at all.\n\nI did not build, run the gate, launch Excel or execute any harness script, so no
finding here is backed by a live observation - all of it is source reading. I did not audit
../xlide_vscode/src (out of scope), engine/dist, ui/editor/dist, node_modules or the installer's WiX
sources beyond installer/build.ps1's engine handling."

---

## C. Performance

Opportunities to make the product or the development loop faster.

Performance in this repo splits cleanly into three populations. The development loop has one large,
certain, cheap win: tools/dev.ps1 repackages the 90 MB engine executable on every run and then
computes, four lines later, the staleness comparison that would have let it skip - a comparison
tools/verify.ps1 already performs first and whose own comment prices the injection at "the better
part of a minute". The shim's host thread is where the product's latency lives, and it is
uninstrumented: every VBE object-model read pays two cross-COM calls because DISPIDs are never
cached, a keystroke in a module over 64,000 characters scans and reallocates the whole module on the
UI thread, and four separate paths (the empty-workspace poll, the pane tick, a tab switch, every
analysis pass) re-read state that has not changed, because in each case the skip or the change key
was placed after the expensive work rather than before it. The page is in better shape than the shim
- explorer.ts, workspace.ts and the host's own findings publisher all carry equality guards - but
the same guard is missing from the problems panel and the diagnostics republish, and three splitter
drag paths still call editor.layout() per pointermove even though main.ts documents that exact
defect being fixed for window resize. Almost nothing here is measured: the perf route times engine
methods and placement passes, so every shim-side item lands with a verified mechanism and an
unmeasured magnitude, which is why the two measurement items (bundle metafile, retained boot
timings) matter out of proportion to their size.

### The ranked list

| # | Finding | Effort | Risk | Confidence |
| --- | --- | --- | --- | --- |
| 1 | tools/dev.ps1 repackages the 90 MB engine on every run, then computes the staleness check that would have skipped it | small | low | verified |
| 2 | Every VBE object-model read costs two cross-COM calls: the member name is resolved through GetIDsOfNames before every Invoke and no DISPID is ever cached | small | low | verified |
| 3 | A keystroke in a module over 64,000 characters scans and reallocates the module's entire text on the VBE host thread | medium | medium | verified |
| 4 | PublishProjects enumerates every project and component twice and posts the whole tree unconditionally, including once a second forever while the workspace is empty | small | low | verified |
| 5 | project/close prunes two of the four per-module maps, so every closed workbook's live module text and outlines stay in the engine for the process lifetime | small | low | verified |
| 6 | The poll tick re-derives each open pane's project identity from scratch, including a thrown-and-caught exception per unsaved workbook, duplicating a walk CodePaneTracker already caches | small | low | verified |
| 7 | Switching tabs re-reads the entire source of every open document over COM, each preceded by an unindexed name scan of the project's components | medium | medium | verified |
| 8 | Every analysis pass reads every module's full source over COM on the host thread before the unchanged-sources skip can decide the pass was unnecessary, and one caller fires it per Immediate-window line | medium | medium | verified |
| 9 | Startup is unmeasurable end to end: the build emits no metafile, and the page's boot breakdown reaches a log line and nothing else | small | low | verified |
| 10 | Every splitter drag runs a synchronous Monaco layout per pointermove on top of automaticLayout, the exact doubling main.ts already diagnosed and fixed for window resize | small | low | verified |
| 11 | engine/test/freshness.mjs, the only headless guard on the memo that saves 446 ms of a 476 ms pass, is run by neither the gate nor CI | small | low | verified |
| 12 | One caret movement republishes diagnostics to every open document and rebuilds the whole problems panel, with no equality guard on either side of the wire | small | low | verified |
| 13 | The whole-project index is rebuilt once per module on every context-cache miss, and past eight modules each rebuild re-parses the entire project | medium | medium | supported |
| 14 | For any module under 64,000 characters the page flattens and ships the whole document on every keystroke, and on a For/Next line it flattens it twice | small | medium | supported |
| 15 | The gate runs the xunit v3 assembly through dotnet test's VSTest host when the assembly is already a runnable executable | small | low | supported |
| 16 | CI runs neither the route audit nor the variant-as-object shape guard, the two cheapest checks in the local gate | small | low | verified |
| 17 | Object Browser search re-sorts the library and rebuilds both panes synchronously on every keystroke, with two fresh listeners per row | small | low | supported |
| 18 | The page is one eager iife with no code splitting and no dynamic imports, so the dev surface, six dialogs and the Object Browser document are all parsed before first paint | medium | medium | verified |
| 19 | The gate packages and ships the engine executable without ever typechecking engine/src | small | low | verified |
| 20 | Nothing reports whether the engine's five memos are hitting, so a caching regression shows only as latency against a baseline nobody keeps | medium | low | verified |

#### C1. tools/dev.ps1 repackages the 90 MB engine on every run, then computes the staleness check that would have skipped it

**Where.** tools/dev.ps1:143-189 (unconditional `node build.mjs --package` at 159, comparison at 172-185);
contrast tools/verify.ps1:107-162

**What.** The step 'Build the engine (bundle, then executable)' runs `node build.mjs --package` with no
condition in front of it whenever -NoBuild is absent (dev.ps1:134, 159). The
newest-source-vs-executable comparison is assembled immediately afterwards at 172-185 and used only
to throw. verify.ps1 does the identical comparison first (113-122) and packages only when
$newer.Count -gt 0, with its own comment at 135-137: the injection 'writes a 90 MB executable and
takes the better part of a minute'. -NoBuild is all-or-nothing (it also skips the unit tests and the
shim publish, dev.ps1:20, 134), so there is no way to keep the loop and drop the packaging.

**Why it matters.** dev.ps1 is the inner loop - build, register, launch - and under the standing directive to republish
the dev shim after every change it runs many times a day, most of those for a C#-only change that
cannot have invalidated the engine. verify.ps1 already judged this cost to be most of a
twenty-second gate and removed it; the script run far more often still pays it, and the data needed
to skip it is gathered four lines later.

**Fix.** Hoist the comparison above the package call and package only when a watched source is newer than
engine\dist\xlide-engine.exe, keeping the post-package read-back assertion for the runs where it did
fire. Adopt verify.ps1's include list (*.ts, *.mjs, *.js - dev.ps1:176 globs only *.ts, so the two
scripts would otherwise still disagree, which is the defect being fixed) and carry across
verify.ps1:145-150, the guard that refuses to package while EXCEL or xlide-engine holds the
executable; dev.ps1 has no equivalent.

**Size.** About a minute per dev.ps1 run that touches no engine or analyzer source, which is most runs. Confirm with one timed `node build.mjs --package`. Effort small, risk low, confidence verified.

#### C2. Every VBE object-model read costs two cross-COM calls: the member name is resolved through GetIDsOfNames before every Invoke and no DISPID is ever cached

**Where.** src/Xlide.Vbe.Shim/Com/DispatchObject.cs:81-96 (GetDispId), :101 (GetProperty), :446 and :462
(GetItem re-resolving the constant "Item" per loop iteration); hot callers at
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:6976-7007, :5781-5814,
src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:162-183

**What.** GetDispId is an unconditional IDispatch::GetIDsOfNames round trip, and the class holds only _pointer
and _dispatch (DispatchObject.cs:23-24), so no memo exists anywhere on the type. All sixteen
internal call sites go straight through, and GetItem re-resolves the literal "Item" inside every
collection loop. Nothing caches on the caller side either. The class doc at DispatchObject.cs:14-17
defends this by claiming hot per-keystroke paths use early-bound interfaces; that is false for the
VBE object model - the only early-bound wrapper, ComHandle<T>, is used for UI Automation and
WebView2 only, never for a VBE member.

**Why it matters.** This is the multiplier under items 6, 7 and 8: every host-thread walk of a VBE collection pays a
name lookup per item, on the thread that freezes the editor. FindComponent over one 60-component
project is about 240 cross-COM calls where 122 would do. Fixing it does not remove those walks, but
it makes each one cheaper everywhere at once, with no behaviour change.

**Fix.** Give DispatchObject a per-instance Dictionary<string,int> populated on first resolve. Same object
means same type means same DISPIDs, so it is exactly safe, and it needs no reflection or codegen, so
NativeAOT is unaffected. That removes the repeated GetIDsOfNames("Item") from every collection loop
and every repeated read on a held object. Correct the class doc while there. A second tier keyed on
type info would also cover the fresh-object-per-iteration case (component.GetString("Name"));
measure the free tier first.

**Size.** Halves the calls for repeated reads on one object. For the 60-component FindComponent walk: 240 calls to 181 with the per-instance tier alone, to about 122 only if a type-keyed tier is added. Per-call microsecond cost unmeasured; the shim has no counter for object-model calls. Effort small, risk low, confidence verified.

#### C3. A keystroke in a module over 64,000 characters scans and reallocates the module's entire text on the VBE host thread

**Where.** src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1388-1389 (reconstruct branch), :1970 (ParseChanges opens
with a full LineStarts scan), :2048 (ApplyEdits allocates a whole new module string per edit);
src/Xlide.Vbe.Core/Engine/EngineProtocol.cs:255-270; trigger at ui/editor/src/bridge.ts:1626-1638

**What.** The page attaches fullText only below 64,000 characters, so above it the shim reconstructs:
ParseChanges runs TextPositions.LineStarts over every character into a List<int> that is then copied
to an array, and ApplyEdits does string.Concat over the whole module per edit. Both are static and
take the text fresh, so nothing survives between keystrokes, and the poster is wired to
model.onDidChangeContent (bridge.ts:491) with no debounce. It runs on the host UI thread:
WebView2Surface.OnWebMessageReceived calls the handler inline (WebView2Surface.cs:846) and its own
comment at :834 says so.

**Why it matters.** The repo's Massive fixture is 64,802 lines / about 1.42 MB and the 11,000-line perf fixture is also
over the threshold. One character costs a 1.4-million-character scan, a ~65,000-entry list grown and
copied, and a 2.84 MB string that lands on the Large Object Heap - per keystroke, on the thread that
draws the editor, forcing gen2 collections during typing. EditorSurface.cs:1460 shows the engine is
fed only the parsed edits for a large module, so this rebuild is the one remaining place in the
per-keystroke path that materialises the whole module.

**Fix.** Hold the reconstruction state on OpenDoc: cache the line-start table beside doc.Text and invalidate
it when the text is replaced, and apply the bottom-up edits into a reusable buffer or a per-document
StringBuilder so one keystroke costs one splice. Materialise the string only where something needs
it - the write-back debounce, the live-analysis pass, the HasUnwritten comparison - not on every
change message.

**Size.** Per keystroke in the 1.42 MB fixture: ~1.4M char reads and ~3.4 MB of allocation, 2.84 MB of it LOH. Wall clock unmeasured; no counter exists on this path. Effort medium, risk medium, confidence verified.

#### C4. PublishProjects enumerates every project and component twice and posts the whole tree unconditionally, including once a second forever while the workspace is empty

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5850, :5880-5884 (first enumeration), :5907-5939 (second),
:5941 (unconditional send); callers at :1604 and :3847; interval at :3728; contrast the change key
at :5647-5652

**What.** The method walks the project collection once for display names and again for the tree, reading Name
and Type per component and calling ProjectReader.Identity a second time per project, then sends
unconditionally. EditorSurface.ShowProjects (EditorSurface.cs:1019-1026) serialises and posts with
no comparison of its own. PollDebugState calls it unguarded whenever _watchingEmpty (:3847), and
UpdatePolling resolves the interval to EmptyWorkspacePollMilliseconds = 1000 in that state with no
decay. PublishModules on the same tick does keep a change key; PublishProjects does not.

**Why it matters.** With no module open - a fresh workbook, or after the last tab closes - the shim re-reads every
component of every open workbook and posts a full setProjects message once a second, indefinitely,
whether or not anything changed. That is the same push that a visibly cycling tree was once traced
to; the page was made idempotent, the source of the push was not. A tab switch also pays a double
enumeration of the object model on the host thread before the new module appears.

**Fix.** Give PublishProjects the treatment PublishModules already has: compose the tree, build a change key
from the project displays plus each project's component names and types, and return before the send
when it matches. Then fold the two enumerations into one pass - the display names can be collected
in the loop that builds the members, with the numbering applied afterwards.

**Size.** Removes one full setProjects serialisation and one full component enumeration per second in the empty-workspace state, and halves the COM enumeration on every tab switch. Absolute cost unmeasured. Effort small, risk low, confidence verified.

#### C5. project/close prunes two of the four per-module maps, so every closed workbook's live module text and outlines stay in the engine for the process lifetime

**Where.** engine/src/dispatcher.ts:444-480 (closeProject prunes lastAnalysis and semanticMemo only); contrast
:427-434 (openProject prunes all four); maps at :199 (liveSources) and :206 (outlineMemo)

**What.** closeProject loops the `${projectId}\0` prefix over lastAnalysis and semanticMemo, deletes
symbolsMemo and calls forgetProjectWords. outlineMemo and liveSources use the identical liveKey
(dispatcher.ts:75-77) and are never touched. openProject's prune does cover all four, but it is
scoped to the opening project and only removes modules absent from the reseed, so a workbook closed
and never reopened is never reclaimed at all. liveSources holds one whole module source per module;
outlineMemo holds a source plus its OutlineResult.

**Why it matters.** The engine is one process per Excel session and outlives every workbook. Closing a workbook releases
the analyzer's per-document state and the analysis memos - the fix the closeProject comment was
written for - and keeps a full copy of the text of every module the developer typed in, plus its
outline. A developer moving between workbooks all day accumulates all of it. This is that same fix,
two maps short.

**Fix.** Extend the existing prefix loop to cover outlineMemo and liveSources. Cheapest correct form is one
loop over a list of the four maps rather than four copies of the loop, so the next map added cannot
be missed the same way.

**Size.** Retained bytes are the sum of every closed project's module sources. Measurable today with no new code by driving project/open and project/close and reading process.memoryUsage(). Effort small, risk low, confidence verified.

#### C6. The poll tick re-derives each open pane's project identity from scratch, including a thrown-and-caught exception per unsaved workbook, duplicating a walk CodePaneTracker already caches

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5777-5814 (ReadOpenModules), :3934 (called every tick),
:2244 and :2247 (150 ms while stepping, 300 ms whenever the Immediate panel is open), :5647-5652
(change key compared only after every read); src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:90-104;
src/Xlide.Vbe.Shim/Editor/CodePaneTracker.cs:340, :502-543

**What.** Per pane per tick the shim walks CodePanes.Item -> CodeModule -> Parent -> Name, then Collection ->
Parent -> ProjectReader.Identity - roughly six to eight property reads. Identity reads Name then
FileName inside a try whose own comment says an unsaved project raises rather than answering empty,
so an unsaved workbook throws once per pane per tick. CodePaneTracker.ReadPaneComponents performs
the identical walk and memoises it in _openComponents with invalidation on rename and on pane
create/destroy, so the tick's walk is a literal duplicate of an already-cached one.

**Why it matters.** Steady state with the Immediate panel open - a panel developers leave open - and five tabs is
roughly 42 property reads, about 84 cross-COM calls at today's two-per-read, every 300 ms on the VBE
host thread, to compute an answer that almost never changes. The dirty-flag half of PublishModules
genuinely has to be polled; the pane-to-project mapping does not.

**Fix.** Expose the tracker's _openComponents and have ReadOpenModules take the owner from there, falling
back to the current derivation only when the tracker has no entry for the pane. This does not
violate ReadOpenModules' own prohibition at :5767-5771, which forbids deriving the list from the
tracker's WINDOW map (maximised panes have only one window) and not from _openComponents, which is
read from the CodePanes collection itself. Separately, cache Identity per project object so the
FileName raise happens once rather than once per pane per tick.

**Size.** Removes roughly six property reads and up to one throw/catch per open pane per tick, at 3.3 ticks per second with the Immediate panel open. Per-read cost unmeasured. Effort small, risk low, confidence verified.

#### C7. Switching tabs re-reads the entire source of every open document over COM, each preceded by an unindexed name scan of the project's components

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3762-3778 (ResyncFromModule), :6969-7011 (FindComponent),
:1610 (called from FollowActivePane past the substance gate at :1573), :2778 (search and replace);
src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:210

**What.** For each open document that has no unwritten edits, ResyncFromModule calls FindComponent - which
walks VBProjects, calls ProjectReader.Identity per project, then walks VBComponents comparing
GetString("Name") per item, with no index - and then ProjectReader.ReadSource, which pulls the whole
module text across. The comment at :3751-3754 acknowledges the resolve walks the project collection
on every pane follow. Only the project identities are cached, and only within one call.

**Why it matters.** Eight tabs in a 60-component workbook means eight unindexed name scans plus eight full module text
transfers, synchronously on the host thread, before the new module appears. The felt threshold for a
tab gesture is around 40 ms and this path has no measurement at all.

**Fix.** Two independent cuts. Resolve the component once per document and keep a (projectId, module) to
index memo, invalidated on the component add/remove/rename paths the session already owns. Then gate
the text read: CountOfLines is one property read, and a module whose line count matches its
_writtenModules baseline needs no full transfer, falling through to the full read when the counts
differ or there is no baseline. Instrument first - a duration around ResyncFromModule on the
existing perf route would say whether this is the tab-switch cost or a rounding error.

**Size.** One full module read per open tab that is not being typed in, plus O(tabs x components) property reads, per tab switch. Unmeasured; no counter exists. Effort medium, risk medium, confidence verified.

#### C8. Every analysis pass reads every module's full source over COM on the host thread before the unchanged-sources skip can decide the pass was unnecessary, and one caller fires it per Immediate-window line

**Where.** src/Xlide.Vbe.Shim/Engine/AnalysisService.cs:845 and :797 (ReadProjectsAsync marshals
ProjectReader.ReadAll onto the host thread), :893-904 (the skip, reached afterwards);
src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:39-67, :143-193, :199-217; trigger at
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:4192-4193

**What.** AnalyseEverythingAsync's first act is the host-thread read of every project and every component,
pulling each module's whole text with GetStringIndexed("Lines", 1, lines). Only then does
SameSources get to conclude nothing changed and skip the engine work. ProjectReader's own header
prices a large project's read at 'tens of milliseconds, which is fine when a project opens and not
fine per keystroke'. AddInSession.cs:4193 calls Reanalyse unconditionally after every
Immediate-window evaluation, and because the scratch module is already removed in
ImmediateEvaluator's finally, that pass almost always reaches the skip and makes no engine call -
the COM read is its entire cost. Two findings, one cause: the read is unconditional, so every caller
pays it whether or not anything changed.

**Why it matters.** The documented skip removed the engine's work for a byte-identical write-back and left the read that
decides it is identical. Reanalyse is not per keystroke (the typing callers sit behind the quiet
timer, and a burst coalesces to at most two passes), but the Immediate path is per gesture, and
Immediate use is rapid-fire during a debug session. The work arrives as a host-thread item of tens
of milliseconds shortly after each line, not as a stall inside Evaluate.

**Fix.** Give the read a cheap host-thread pre-check: component count plus CountOfLines per component -
CountOfLines is already the first read inside ReadSource - and skip the GetStringIndexed transfer
for a project whose shape and line counts are unchanged, falling back to the full read whenever
anything differs. That is one change that pays off for every Reanalyse caller. Do not instead gate
the Immediate call site on whether cleanup removed something unexpected: :4193 corrects for a pass
that READ while the scratch module existed, which is reachable because VBA execution pumps messages
and a marshalled read can land mid-Evaluate. Instrument ProjectReader.ReadAll on the perf route
first; today it reports engine timings and marshal samples and nothing about this read.

**Size.** One full source read of every open workbook per write-back burst, plus one per Immediate line. ProjectReader.cs:19 puts a large project at tens of milliseconds; unmeasured here. Effort medium, risk medium, confidence verified.

#### C9. Startup is unmeasurable end to end: the build emits no metafile, and the page's boot breakdown reaches a log line and nothing else

**Where.** ui/editor/build.mjs:253-285 (no metafile) and :301-317 (reports only total bytes);
tools/verify.ps1:207-217 (no size budget) and :309;
src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:2119-2150 (DescribeTimings formats and discards) and
:1307

**What.** The esbuild common options set bundle, minify, target, alias and loader with no `metafile: true`,
and the build reports only readdir/stat byte counts: editor.js 3,642,564 bytes, editor.css 184,668,
editor.worker.js 304,346. There is no per-input attribution and no ceiling anywhere; the only size
assertion in the repo is a lower bound on the worker. Separately, the page posts a full breakdown
(main.ts:984-997: scriptMs, createMs, totalMs, fetchMs, transferBytes, requestStartMs, htmlMs) and
DescribeTimings reads each field into a local, formats a string for one Log.Info, and retains
nothing. Grepping the route switch, the client and docs/debug-api.md for scriptMs or timings finds
nothing.

**Why it matters.** Startup is the page's largest single cost and every decision about what to cut from that 3.64 MB -
which of the ~40 eager Monaco feature registers is expensive, whether the object browser split is
worth it - currently rests on guesswork, while esbuild computes the answer during the build and
throws it away. And a change that adds 400 ms to bundle parse ships with nothing failing: a suite
can pull the formatted log line through the existing 'log' route, but nothing can hold scriptMs to a
number. Items 14 and 18 both end in 'measure this first' and this is the item that lets them.

**Fix.** Add `metafile: true` to the common options, print the top 20 inputs by bytesInOutput and write
dist/metafile.json for esbuild's analyzer, then consider a size assertion in verify.ps1 beside the
existing warnings check. Separately, retain the parsed timings on the surface as a small record and
answer them from a NEW debug-api noun (not a field on perf, and no existing route touched),
documented in both docs/debug-api.md and docs/driving-excel.md and driven by a harness check so
audit-routes.mjs is satisfied, with the route entry saying plainly that it reports the page's own
measurement.

**Size.** Zero runtime cost. Two lines in build.mjs plus one route; it is the instrument that sizes items 14, 18 and 19. Effort small, risk low, confidence verified.

#### C10. Every splitter drag runs a synchronous Monaco layout per pointermove on top of automaticLayout, the exact doubling main.ts already diagnosed and fixed for window resize

**Where.** ui/editor/src/paneldocks.ts:820-825 and :874-890 (applies), :849 and :914 (pointermove listeners);
ui/editor/src/workspace.ts:831-846, :869, :470-472; ui/editor/src/main.ts:327, :438, :467, :444-461

**What.** Each apply() writes a style then calls this.handlers.layoutChanged(), which is
editors().forEach(editor => editor.layout()) over every open group, and each apply() is called once
per pointermove with no rAF or settle. In the installed monaco 0.56, layout() calls observeContainer
(a forced clientWidth read) then a synchronous render(), while every editor is already created with
automaticLayout: true, whose ResizeObserver covers the same resize. main.ts:444-461 records the
identical defect on the window-resize path - 'running it per event doubled every layout of a drag,
which read as latency and churn' - and its 150 ms settle plus body.live-resize fix was applied only
there; grepping live-resize finds it only in main.ts and styles.css, so no splitter drag gets it.

**Why it matters.** Dragging a dock or group splitter is the most common layout gesture on the page. A 1000 Hz mouse
delivers several pointermoves per frame, so the drag pays N_groups forced layouts and synchronous
renders per event plus one more per frame from the observer, and the minimap flicker the
window-resize fix describes is still present on exactly this gesture.

**Fix.** Delete the explicit editor.layout() from the drag path and let automaticLayout's observer do it, or
coalesce layoutChanged into one requestAnimationFrame per drag with a flush on pointerup if the
observer proves to miss the final frame. Toggle document.body.classList('live-resize') around the
drag the way the window-resize handler does.

**Size.** N_groups forced layouts and synchronous renders per pointermove removed (N_groups is 1 in the common case). Settle with a performance.mark around apply() during a two-second drag with a 4,000-line module open, before and after. Effort small, risk low, confidence verified.

#### C11. engine/test/freshness.mjs, the only headless guard on the memo that saves 446 ms of a 476 ms pass, is run by neither the gate nor CI

**Where.** engine/package.json:12 (chains four suites under 'test'); tools/verify.ps1:229-244 (runs only
test/language.mjs); .github/workflows/build.yml:69, :85, :94 (smoke, language, positions - not
freshness); engine/src/dispatcher.ts:104-107, :875, :944-955

**What.** Nothing runs `npm test` in engine/. The gate's only engine step runs language.mjs; CI runs three of
the four and skips freshness. The two matches for 'freshness' in tools and .github are the -Live
suite tools\harness\analysis-freshness.mjs, which needs an open editor. The dispatcher's own comment
records that freshness.mjs passes with the describe() Map handling REMOVED, so the suite is loose as
well as unrun.

**Why it matters.** The diagnostics memo is the largest single perf decision in the engine and is correct only while
crossModuleFingerprint is exact; when it stops being exact the failure is silent - no squiggle
appears where one should. Two of the three memo defects already recorded in dispatcher.ts were found
by confusion rather than by a check. This is also the cheap alternative to building memo hit
counters (item 20): one line of gate wiring against a new engine method plus a route plus docs.

**Fix.** Add `node test/freshness.mjs` beside the language.mjs line in verify.ps1's engine step, and after
Positions in the CI engine job. Extend freshness.mjs with a case that fails when describe() degrades
to plain JSON.stringify, since the dispatcher records that today it does not. Note that
freshness.mjs spawns dist/engine.cjs, the same assumption language.mjs already runs under: on a gate
run where nothing was stale, build.mjs is never invoked and the bundle is whatever the last build
left.

**Size.** One engine process over one pipe doing seed/ask/reseed/ask; comparable to smoke.mjs, which CI already affords on both runners. Effort small, risk low, confidence verified.

#### C12. One caret movement republishes diagnostics to every open document and rebuilds the whole problems panel, with no equality guard on either side of the wire

**Where.** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:7084-7101 (per-document filter and unconditional
ShowDiagnostics), :1144-1159 (both hold call sites), :6445-6457 (whole findings set reserialised);
src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:979-988; ui/editor/src/shell.ts:880-884 and :1076-1162;
contrast ui/editor/src/explorer.ts:418-425 and ui/editor/src/workspace.ts:323-335

**What.** Two findings, one gesture. On the shim: an active-line hold transition re-filters the entire
_findings list once per open document and calls ShowDiagnostics unconditionally, and
ActiveLineHold.Hides requires the module to match, so the other N-1 messages are provably
byte-identical to what the page already holds; the same call sites also reserialise the whole
findings set. On the page: setFindings assigns and calls renderPanel with no comparison, and
renderPanel does replaceChildren, a full copy-filter-sort, a second pass building findingHomes, then
five elements and four dataset writes per surviving finding straight into the live pane. Both
siblings on the same message do guard - explorer.setProblemCounts returns early on equal counts,
renderTabs returns on an unchanged renderKey.

**Why it matters.** The hold fires on every line the caret enters or leaves while typing, and it rebuilds the panel even
when the held line carries no findings, so the visible set is identical - which also destroys the
row the developer had focused or was about to click. The per-pause LIVE path is already deduped on
the host, so the unguarded pushes are this one and the project pass.

**Fix.** On the shim, publish markers for the module whose hold changed rather than for every open document -
both hold call sites know the module - keeping the all-documents form for the paths that genuinely
replace the whole set. On the page, give setFindings the renderKey guard renderTabs already uses:
join severity|module|project|line|column|code|message plus the three filter states, compare, return
when equal. Three lines each, in the shape each file already uses.

**Size.** N-1 redundant setDiagnostics messages per line transition (cheap each), one whole-findings serialisation per transition (the larger half), and roughly six DOM node creations per finding per push on the page. Unmeasured. Effort small, risk low, confidence verified.

#### C13. The whole-project index is rebuilt once per module on every context-cache miss, and past eight modules each rebuild re-parses the entire project

**Where.** engine/src/moduleContext.ts:68-89 (cache key) and :131-149 (miss builds a project-wide index);
engine/src/dispatcher.ts:124-160 (a second whole-project index, kept only as a string map);
../xlide_vscode/src/analyzer/parser/parseModule.ts:135-136 (PARSE_CACHE_MAX = 8)

**What.** assembleContext caches per (seeded array identity, module name), so the miss path - which calls
buildLiveVbaProjectIndex over every module of the project, with liveOverride substituting one entry
rather than narrowing the walk - runs once per MODULE per seed. dispatcher.crossModuleFingerprint
separately builds another whole-project index from the same seeded array and keeps only the
resulting Map<string,string>. The analyzer's parse cache is an 8-entry LRU keyed on source, so
building an index over a project with more than eight modules evicts every entry as it goes and the
next module's context build re-parses the project from scratch.

**Why it matters.** Completion, hover, signature help, canonicalCase and semantic tokens all route through
assembleContext, on a pipe that serves one request at a time, so each rebuild also delays whatever
queued behind it. After a reseed - every write-back whose text actually differs - the first request
in each module pays this, and above nine modules it is a full N-module parse each time, not a warm
re-walk. The moduleContext comment claiming the cache turns completion 'from indexing the whole
project into scanning one module' holds for repeat requests in one module and not for the first in
each.

**Fix.** Cache the ProjectIndex itself on the seeded array in a WeakMap the way crossModuleFacts already
caches its fingerprints, and have both callers take it from there. Two real differences must be
preserved rather than assumed away: buildLiveVbaProjectIndex passes ignoreInvalidModules:true and
buildVbaProjectIndex does not, and buildContext's liveOverride replaces the current module's entry
with one carrying moduleKind in place of type/documentType. Keep the live variant as the shared
build and assert the fingerprints are unchanged with engine/test/freshness.mjs (see item 11 - wire
it in first).

**Size.** Unmeasured. Settle by timing textDocument/completion in a not-yet-cached module against a repeat in the same module, on the perf fixture, through the engine's own pipe; engine/test/smoke.mjs already builds that harness. Effort medium, risk medium, confidence supported.

#### C14. For any module under 64,000 characters the page flattens and ships the whole document on every keystroke, and on a For/Next line it flattens it twice

**Where.** ui/editor/src/bridge.ts:1615-1620 (the comment) and :1626-1640 (`if (fullLength < 64_000) {
message.fullText = model.getValue(); }`); src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1379-1399;
ui/editor/src/typing.ts:47, :150-152, :481-500

**What.** Below the threshold the message carries both the changes array and the full text, and the host
prefers fullText, so the incremental path never runs for a typical VBA module. Separately,
maybeLoopSync gates on the current line matching /^[ \t]*(?:For|Next)\b/i and then calls
resolveLoopIteratorSyncEdit(model.getValue(), offset) - a second full flatten of the same model
version on the same keystroke. The vendored helper early-outs on the physical line and only scans
the document when the offset touches an iterator span, so the cost there is the getValue itself, not
the scan.

**Why it matters.** Per keystroke this is an O(document) string build from Monaco's piece tree, a JSON serialisation, a
WebView2 postMessage copy and a JsonDocument parse of the same bytes on the host, for information
already derivable from the ranges in the same message - likely tens of microseconds on a few-KB
module, which is why the bigger payoff is the coverage gap: 64,000 is an unmeasured constant, and
the reconstruction branch (item 3) therefore executes only on modules above it, so it is close to
untested in normal use. No test anywhere exercises either branch.

**Fix.** Measure the two paths with a 60 KB module against a 66 KB one at the same typing rate, then either
move the threshold to the real crossover or drop fullText and keep the fullLength cross-check that
already catches divergence. Do not drop it before the reconstruct branch has a test. Meanwhile share
one flatten between bridge.ts and maybeLoopSync for the same model version rather than taking it
twice; a version-keyed cache does not help across keystrokes, because every keystroke bumps the
version.

**Size.** Up to 64 KB of string build plus JSON plus IPC plus host-side parse per keystroke, doubled on loop-header lines. Typical VBA modules are far smaller than the ceiling. Effort small, risk medium, confidence supported.

#### C15. The gate runs the xunit v3 assembly through dotnet test's VSTest host when the assembly is already a runnable executable

**Where.** tools/verify.ps1:278-286; tests/Xlide.Vbe.Core.Tests/Xlide.Vbe.Core.Tests.csproj:13-15;
artifacts/bin/Xlide.Vbe.Core.Tests/release/Xlide.Vbe.Core.Tests.exe beside testhost.exe

**What.** The step runs `dotnet test $solution -c Release --no-build --nologo -v q`. The project references
Microsoft.NET.Test.Sdk, xunit.v3 and xunit.runner.visualstudio, so the run goes through VSTest, and
the self-hosting executable xunit v3 emits is already on disk in both configurations. The solution
contains exactly one test project.

**Why it matters.** dotnet test pays MSBuild evaluation of the solution, VSTest discovery and a testhost launch before
the first test runs. Because xunit v3's adapter delegates to the test project's own executable
anyway, running it directly removes those three and not the execution - which also means the Smart
App Control constraint that keeps dev.ps1 on Debug test binaries is not a new risk introduced by the
change.

**Fix.** Run artifacts\bin\Xlide.Vbe.Core.Tests\release\Xlide.Vbe.Core.Tests.exe directly, setting
DOTNET_ROOT to the local dotnet directory verify.ps1:44-47 already locates, and parse xunit's own
summary instead of the 'Passed: N' line scraped today. Keep dotnet test in CI, where the trx logger
is wanted.

**Size.** A few seconds inside a roughly twenty-second gate. One timed run of each form settles it; the gate already prints per-step seconds. Effort small, risk low, confidence supported.

#### C16. CI runs neither the route audit nor the variant-as-object shape guard, the two cheapest checks in the local gate

**Where.** .github/workflows/build.yml (219 lines, no audit-routes and no As<object> grep anywhere);
tools/verify.ps1:165-196 and :259-264; the third candidate at :246-253 with
tools/harness/page-probe.mjs:72-80

**What.** The build job is npm ci, spec:check, typecheck, build, npm test, dotnet restore/build/test,
Languages, publish, three pwsh binary checks, upload. The variant guard is a Select-String for
`\.As<\s*object\s*>\s*\(` over src\**\*.cs whose own comment records that this defect has killed
Excel twice and that com-leak.mjs cannot catch either occurrence; audit-routes.mjs reads the routes
out of the shim and fails when one is undocumented or undriven. Neither runs in CI.

**Why it matters.** CI is the only check that runs on a change that did not come from this machine, and the two guards
it lacks need nothing but a checkout and under a second each. One of them stands in for a crash that
ends the host process and that the leak sweep provably cannot see.

**Fix.** Add the variant grep (it can run before dotnet restore) and `node tools/harness/audit-routes.mjs` to
the build job. Treat the five headless page probes as a separate decision: page-probe.mjs launches
msedge.exe from two hardcoded Windows paths and throws when neither exists, so they can only live in
the windows-latest job and they add a browser dependency and a new failure mode (a runner image
change silently breaks five checks) rather than being the zero-dependency addition they look like.

**Size.** One Select-String over src and one node process. Sizeable from the gate's own per-step seconds for those two steps. Effort small, risk low, confidence verified.

#### C17. Object Browser search re-sorts the library and rebuilds both panes synchronously on every keystroke, with two fresh listeners per row

**Where.** ui/editor/src/objectbrowser.ts:538-541 (undebounced input handler), :444-508 (renderTypes copies and
sorts the whole array, then click and keydown per row), :310 and :372-389 (renderMembers);
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:2506-2508 (no cap); precedent at
ui/editor/src/searchwidget.ts:429-430

**What.** The input handler calls renderTypes and renderMembers with no debounce; renderTypes starts with
replaceChildren then iterates [...types].sort(byName), so the array is copied and sorted on every
call, and each matching row builds four elements plus two listeners. The host caps nothing. The
fixed per-keystroke cost is the copy and sort of the whole library; rows are only built for matches,
so the DOM work shrinks as the query grows.

**Why it matters.** Typing in the search box is the palette's main interaction and every character pays a full sort of
the library plus a rebuild of the matching rows, synchronously on the input event. The magnitude
depends on a type count nobody has read - the shim logs it at AddInSession.cs:2507 - and a typelib
is normally in the hundreds, not thousands.

**Fix.** Cheapest first: sort once when the types arrive rather than inside renderTypes, and debounce the
input handler at 150 ms the way searchwidget.ts already does. Then, if a measured row count
justifies it, replace the per-row click/keydown listeners with one delegated listener on typesPane
reading a data attribute. Virtualisation should wait for that measurement.

**Size.** One full array copy and sort per keystroke, plus four elements and two listeners per matching type. Log types.length for the Excel and Office libraries and time renderTypes with performance.now() before going further. Effort small, risk low, confidence supported.

#### C18. The page is one eager iife with no code splitting and no dynamic imports, so the dev surface, six dialogs and the Object Browser document are all parsed before first paint

**Where.** ui/editor/build.mjs:287-295 (format iife, no splitting) and :259-263 (only three defines, one config
for every build); ui/editor/src/main.ts:1-88, :929, :1242-1250;
src/Xlide.Vbe.Shim/Editor/BrowserPalette.cs:96-99; ui/editor/src/objectbrowser.ts:164-171;
src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:2398-2404

**What.** Two findings, one cause: everything ships in one eagerly evaluated bundle. esbuild supports
splitting only for esm, so the iife choice forecloses it, and grepping ui/editor/src for `import(`
returns nothing. main.ts eagerly imports installDevSurface (63 KB of source, called unconditionally
and present in Release too), openSyncDialog, the settings/help/sponsor/references dialogs and
bootObjectBrowserPage - about 145 KB of source for surfaces that start closed. The ?view=objbrowser
branch sits at the bottom of main.ts, after all ~40 Monaco feature registers have already evaluated,
so the palette's WebView2 loads and executes the entire 3.64 MB bundle to draw two lists, then
removes #shell. That cost is once per session, not per open: AddInSession returns early with
Present() when the palette already exists.

**Why it matters.** The largest startup cost is bundle parse and today nothing can be moved out of it - there is no
mechanism for 'load this when the pane opens', which is also the prerequisite for ever deferring
anything on the Monaco side. The app's own share is the smaller term (single-digit percent of the
output after minification); the ~40 eager feature registers are the bulk and no format change defers
them.

**Fix.** Strictly after item 9, so the saving is a number rather than a hope. Then switch the page entry to
format esm with splitting and index.html to a module script - the worker entry is already esm and
the CSP is already script-src 'self' blob: - and convert the dialog openers and installDevSurface to
await import() at their call sites. Give the palette its own entry point and document. Note that a
module script is deferred, which changes when boot() runs relative to the host's ready handshake;
that is the real risk and needs checking.

**Size.** About 145 KB of unminified app source off the first-paint path, plus the whole bundle off the palette's first load, plus the ability to defer anything else later. Exact output bytes unknown until the metafile exists. Effort medium, risk medium, confidence verified.

#### C19. The gate packages and ships the engine executable without ever typechecking engine/src

**Where.** tools/verify.ps1:152 (packages) and :198-205 (the only typecheck, and it is ui/editor);
engine/package.json:9; engine/build.mjs:29-40; .github/workflows/build.yml:58-64

**What.** esbuild strips types without reading them, and nothing calls `npm run check-types` in engine/
outside CI - the only two hits in the repo are the package.json script and the CI engine job, and
that job runs `npm run build`, never --package, so no packaged executable in this project is ever
produced by a typechecked run. dev.ps1:159 packages untypechecked too.

**Why it matters.** A type error in engine/src passes the local gate, gets injected into xlide-engine.exe by that same
gate, and is what the developer then tests against and can publish. The 'engine executable is
current' step makes the executable current and unverified in one move.

**Fix.** Run `npm run check-types` in engine/ inside that step, before the package call, and in dev.ps1
before its package call. Budget properly: engine/tsconfig.json includes only src, but every engine
module imports the analyzer by relative path into ../../../xlide_vscode/src, and those are .ts
sources that skipLibCheck does not skip, so tsc typechecks the whole shared analyzer - seconds to
tens of seconds, and it will go red whenever the neighbouring analyzer tree has a type error. That
is defensible, since that tree is what gets bundled, but the step's failure text must name the
analyzer, because this repo is forbidden to patch it. Ranked here rather than higher because it
makes the gate slower in exchange for a guard, which is the opposite trade from every item above it.

**Size.** tsc --noEmit over engine/src plus the analyzer checkout, on the runs that were going to package anyway. Effort small, risk low, confidence verified.

#### C20. Nothing reports whether the engine's five memos are hitting, so a caching regression shows only as latency against a baseline nobody keeps

**Where.** src/Xlide.Vbe.Shim/Diagnostics/EngineCounters.cs:39-90;
src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1005-1029 (perf), :1794-1836 (stats), :2461-2502
(engine); engine/src/dispatcher.ts:352 (debug/liveSource is the only debug method)

**What.** EngineCounters tallies Calls, WaitTotalMs, CallTotalMs, WaitMaxMs, CallMaxMs, Refused and a 32-slot
ring, and the perf route surfaces exactly those. No route reports lastAnalysis, semanticMemo,
outlineMemo, symbolsMemo or crossModuleFacts state; stats covers process, placement, marshal,
refresh, message and COM-wrapper counters. From the shim's side a hit and a cheap miss are the same
fast call, and a pass that takes the SameSources skip makes no engine call at all, so a correct skip
and a broken pass are both zero.

**Why it matters.** Both memo defects already recorded in dispatcher.ts were silent and would have shown instantly as a
hit rate falling to zero. Ranked last deliberately: both were closed by tests, and the test that
covers the same risk already exists and merely runs nowhere (item 11, one line). This is a new
engine method plus a new route plus a client method plus two doc entries to satisfy
audit-routes.mjs, and the two compete for the same budget.

**Fix.** Only after item 11 is wired in and only if a measured latency regression turns on memo behaviour.
Then add a new engine method beside debug/liveSource - debug/memos - answering per project the entry
counts and hit/miss counters for the five maps, with a reset, and surface it as a NEW debug-api
route beside perf rather than a field on perf, since it is engine state and not shim timing. Its
route entry must say plainly that it stops at the engine layer, because it answers what the render
depends on rather than what the render shows.

**Size.** Five counter pairs and one route; about a dozen numbers per project in the reply. No runtime cost worth naming. Effort medium, risk low, confidence verified.

### Themes

- The skip was added at the consumer and not at the producer. The analyzer's unchanged-sources skip,
  the findings dedupe and the modules change key all fire AFTER the expensive COM read or
  serialisation that feeds them (items 4, 8, 12). A rule worth adopting: the comparison that decides
  whether work was needed belongs upstream of the work, not downstream of it.
- A guard applied to one path and not to its siblings. The window-resize settle exists and three
  splitter drags lack it; openProject prunes four maps and closeProject prunes two; PublishModules
  keys its change and PublishProjects does not; explorer and the tab strip guard and the problems
  panel does not (items 4, 5, 10, 12). Each fix is a copy of code already in the same file.
- Two scripts that should decide identically do not. verify.ps1 has the engine staleness gate, the
  wider include list and the running-host check; dev.ps1 has the assertion but not the gate, a
  narrower glob, and no host check (items 1, 19). The gate and CI have diverged the same way, and the
  checks CI lacks are the cheapest ones (item 16).
- The host thread is the product's latency and nothing measures it. The perf route times engine
  methods and placement passes, so every shim finding here lands with a verified mechanism and an
  unmeasured magnitude (items 2, 4, 6, 7, 8). The same is true of startup on the page (item 9). Adding
  a duration around ProjectReader.ReadAll, ResyncFromModule and the contentChanged handler would
  settle five of these at once.
- The whole module text is materialised somewhere on every keystroke. Under 64,000 characters the page
  flattens and ships it; over 64,000 the host scans and rebuilds it; on a For/Next line the page
  flattens it twice (items 3, 14). Neither side maintains an incremental representation, so the
  threshold only chooses which side pays.
- Name resolution and object lookup are repeated where an index would do. DISPIDs resolved per call,
  components found by walking and comparing names, pane owners re-derived per tick beside a cache that
  already holds them (items 2, 6, 7). The DISPID cache is the shared multiplier: it makes each walk
  cheaper and removes none of them.
- Three engine and gate items exist because a check was written and never wired in, or written for one
  layer only. freshness.mjs runs nowhere, engine/src is never typechecked before being injected into
  the shipped executable, and the route audit and variant guard stop at the local machine (items 11,
  16, 19).

### What this section leaves out

"Merges. Twenty-five surviving findings became twenty items. Five merges, each because two findings
shared one cause: readall-before-the-skip and immediate-line-rereads-every-module (cause:
ProjectReader.ReadAll is unconditional, so every caller pays it - fixing the Immediate call site
alone is the wrong fix, and the adversary showed why); markers-posted-to-every-open-document and
problems-panel-rebuilt-on-every-push (cause: one hold transition republishes unguarded on both sides
of the wire); fulltext-shipped-per-keystroke and loop-sync-reads-whole-document-per-key (cause: a
full model flatten per keystroke, twice on loop lines); bundle-is-unattributed and
page-boot-timings-are-log-only (cause: startup has no retained measurement at either end);
iife-forecloses-lazy-loading and objbrowser-loads-the-whole-editor (cause: one eager bundle serves
every document). In each case the underlying cause is named in the item and fixing it is a different
job from fixing the instances - most obviously item 2, which is the cross-COM multiplier under items
6, 7 and 8 but removes none of their walks.\n\nDropped, two. glyph-hover-decoration-per-mousemove:
real, but the churn is confined to pointer samples inside a roughly 20px glyph margin during a brief
gesture, a few dozen single-decoration deltas, and the adversary reduced it to a churn tidy rather
than a latency item. gate-runs-node-and-dotnet-in-sequence: the proposed mechanism does not exist on
this machine (Windows PowerShell 5.1 has no Start-ThreadJob without an installed module, and
Start-Job costs about a second per group), both groups already saturate the CPU through MSBuild,
esbuild and the AOT linker, and the join ceiling is single-digit seconds against a medium-effort
restructure of the gate's step harness - a poor trade next to items 1 and 15, which take time off
the same loop for a fraction of the work.\n\nWhat I verified myself. I re-read tools/dev.ps1:1-60
and :125-194 and tools/verify.ps1:100-170 (item 1, including that -NoBuild is all-or-nothing so
there is no existing escape), src/Xlide.Vbe.Shim/Com/DispatchObject.cs:10-125 (item 2, including the
class doc's incorrect claim about early-bound hot paths), and
src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1355-1409 (item 3, the reconstruct branch and the
fullText preference). Items 4 through 20 rest on the finders' evidence as re-read by the adversary;
I did not independently reopen those files in this pass, and where the adversary corrected a line
number or a magnitude I carried the correction rather than the original.\n\nHonest limits. Nothing
was executed: no build, no gate, no Excel, no engine, read-only throughout, so every magnitude here
is arithmetic over code and not a measurement. Every shim-side item is therefore 'verified
mechanism, unmeasured cost', which is itself item 9's point. The analyzer under ../xlide_vscode/src
was consulted only where the engine calls into it (the parse cache size behind item 13) and is out
of scope for any change. Not examined at all in this consolidation: the installer, the NativeAOT
publish configuration, WebView2 startup on the host side, the -Live harness suites' own runtime, and
whether any of the proposed page changes affect the com-leak sweep."

---

## Appendix: every standing finding

All 118, grouped by the area that raised them, with the evidence each rests on and the adversary's
reason for letting it through. The ranked sections above are a view of this register, not a
different set: an item can appear there merged with two others.

### Where driving Excel still falls off the api

_8 findings, from the `api-fallbacks` finder._

##### `sync-dialog-only-reachable-by-eval` The import/export dialog has no act action and no ui field, so the gate's module-sync suite drives it with eleven raw DOM eval scripts

- **Where:** `tools/harness/module-sync.mjs:248`
- **Kind:** api-coverage / medium effort, claim observed, confidence verified, severity medium
- **Evidence:** module-sync.mjs opens, fills, switches, reads and applies the dialog entirely through the generic
  escape hatch: line 248 `await api.eval(`(() =>
  document.querySelector("button[data-command=openSync]").click())()`)`, line 250 sets
  `#sync-folder`.value and dispatches a synthetic change event, line 271 clicks
  `.sync-direction[data-direction=import]`, line 280 finds the Apply button by its textContent, line
  289 clicks `#sync-close`. The plan it compares against the route is scraped by the SYNC_ROWS
  constant at line 20: `[...document.querySelectorAll(".sync-item")].map(r =>
  r.querySelector(".sync-item-name").textContent + ":" + r.querySelector(".sync-chip").textContent)`.
  Eleven api.eval calls in total (lines 248, 250, 255, 270, 271, 273, 274, 276, 280, 289, 311). The
  page side has nothing to offer instead: the act table in ui/editor/src/devsurface.ts:514 lists 33
  actions (closeActive, treeMenu, renameModule, dock, press, format, bookmark and so on) and none of
  them is a sync action - a grep for "sync" across devsurface.ts returns only the words "synchronous"
  and "async". UiSnapshot at devsurface.ts:38 carries workspace, explorer, panes, dialogs, waiting,
  focus, settings, emptyViewShown, longTasks, census, search, bookmarks and at; the dialog contributes
  only an id and a title through dialogsUp() (devsurface.ts:182 maps `[aria-modal="true"]` to `{id,
  title}`), so the rows the render computed are invisible to `ui`.
- **Why:** This is the suite that exists to prove the dialog and the route leave the same state, and it is the
  one place in the gate where the UI half of that comparison is a scrape. Every selector in it
  (button[data-command=openSync], #sync-folder, .sync-direction, .sync-item-name, .sync-chip,
  #sync-foot, #sync-close) is a private detail of syncdialog.ts that no compiler or gate step
  protects; a CSS class rename turns a green comparison into a red suite with a fault that is in the
  harness. It also runs twice per gate, once per planner, so the cost lands twice. And the scrape
  cannot tell a stale render from a correct one, which is exactly the defect docs/driving-excel.md:836
  records the suite already tripping over.
- **Change:** Add a sync action to the act table and a sync field to the ui snapshot. act("sync", {open|close}),
  act("sync", {folder}), act("sync", {direction}) and act("sync", {apply:true}) can call the same
  functions the dialog's own listeners call (syncdialog.ts wires plain click and change handlers at
  lines 140, 169, 375, 381 and 753, so no pointer sequence is involved), and ui.sync = {open, folder,
  direction, mode, rows:[{name, status, selected}]} would carry the plan from the same array the row
  builder at syncdialog.ts:471 renders. The suite then compares route against dialog through two typed
  answers rather than through a string join.
- **Size:** 11 api.eval calls and 7 private CSS selectors removed from one gate suite that runs twice per gate
- **Adversary:** Every cited line reads as claimed. tools/harness/module-sync.mjs has exactly 11 api.eval calls, at
  248, 250, 255, 270, 271, 273, 274, 276, 280, 289, 311 (grep of the file), and SYNC_ROWS at line
  20-22 is the .sync-item/.sync-item-name/.sync-chip scrape. The seven selectors are all private
  details of ui/editor/src/syncdialog.ts (sync-card at 90, sync-close at 105, .sync-direction at 121,
  sync-folder at 164, sync-foot at 365, .sync-item at 471, .sync-chip at 510) and nothing else in the
  repo references them except this suite and one prose line in docs/driving-excel.md:837. A
  case-insensitive grep for 'sync' across ui/editor/src/devsurface.ts returns only 'synchronous' and
  'async', so there is no sync act action and no sync field on UiSnapshot (devsurface.ts:38-93 lists
  exactly the 13 fields named; dialogsUp at 182-192 yields only {id,title}). The only other harness
  reference to openSync is tools/harness/menu-bar.mjs:159, which merely asserts the button is on the
  toolbar, so no api path opens the dialog. verify.ps1:422 runs 'module-sync.mjs xlide' and
  'module-sync.mjs builtIn', so the cost does land twice per gate.
- **Correction applied:** The act table holds 34 actions, not 33: the 30 from closeActive (devsurface.ts:525) through press
  (1289) plus insert, backspace, format and bookmark (1342-1413). The stale-render note is
  docs/driving-excel.md:835-841, not 836 exactly.

##### `objbrowser-live-probe-uses-cdp-for-what-eval-serves` The live Object Browser probe opens a DevTools WebSocket to script the palette page although eval already targets that surface and the probe sends no trusted input

- **Where:** `tools/harness/objbrowser-live-probe.mjs:25`
- **Kind:** api-coverage / medium effort, claim observed, confidence verified, severity low
- **Evidence:** The probe connects a CDP socket (`const socket = new WebSocket(wsUrl)`, line 25), attaches to the
  palette by window title (`const palette = await attachToPage(send, "Object Browser")`, line 122) and
  runs its whole interaction through `Runtime.evaluate`. What it evaluates needs no trusted input at
  all: the toolbar summons is `button.click()`, the library picker is `picker.dispatchEvent(new
  Event('change', {bubbles:true}))`, the row pick is `target.click()` and the navigate leg is
  `member.dispatchEvent(new MouseEvent('dblclick', {bubbles:true, cancelable:true}))`. All four are
  synthesised DOM events, which is what the eval route already delivers. The shim can reach that page:
  RunPageScriptOnce (src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:221) picks the browser with
  `surface == "palette" ? _browserPalette?.Browser : _editorSurface?.Browser`, the eval route reads
  `surface` from the query at line 1102, and the client exposes it as `eval(script, surface)` at
  tools/harness/xlide-api.mjs:666. Test-SplitWorkspace.ps1:484 already proves it works live:
  `Invoke-RestMethod "$api/eval?surface=palette" -Method Post -Body 'document.title'`. The one thing
  eval cannot do for the palette is describe it: the ui route hardcodes the surface
  (AddInSession.DebugApi.cs:1054, `RunPageScript($"window.xlideUi.state({arguments})", null, ...)`),
  so ui can only ever answer for the editor page.
- **Why:** The only live coverage the Object Browser has is a probe that needs a second door open, needs the
  devtoolsPort, and speaks the DevTools protocol by hand for 60 lines of plumbing (connect,
  attachToTarget, evaluate, exceptionDetails). That plumbing is why the probe is not in the gate.
  docs/driving-excel.md:1272 tells the next reader that CDP is what the palette needs, which sends
  them down the same road; the actual boundary is narrower than the doc says.
- **Change:** Replace the CDP transport in this probe with api.eval(script, "palette"), which deletes
  connect/attachToPage/evaluate and the --cdp argument, then let ui answer for the palette: honour an
  optional surface query on the ui route (`RunPageScript(..., which, ...)` instead of the hardcoded
  null) and install a devsurface state on the palette page carrying libraries, module rows, member
  rows and the details pane. Keep CDP in the doc only for genuinely trusted input, which nothing live
  currently needs.
- **Size:** about 60 lines of hand-written DevTools plumbing and one --cdp argument removed from the only live Object Browser probe
- **Adversary:** The transport claim holds. tools/harness/objbrowser-live-probe.mjs:24-76 hand-rolls
  connect/attachToPage/evaluate over a DevTools WebSocket, and all four interactions it drives are
  synthesised DOM events (button.click() at 106, picker.dispatchEvent(new Event('change')) at 133,
  target.click() at 139, member.dispatchEvent(new MouseEvent('dblclick')) at 145). The shim can serve
  those: RunPageScriptOnce picks the palette browser at AddInSession.DebugApi.cs:220, the eval route
  reads surface at 1101, the client exposes eval(script, surface) at tools/harness/xlide-api.mjs:666,
  and RunPageScript awaits a promise via a ticket-and-poll wrapper (AddInSession.DebugApi.cs:78-161),
  so the probe's async IIFEs would work. The ui route does hardcode the surface:
  AddInSession.DebugApi.cs:1054 passes null.
- **Correction applied:** Two citations are wrong and one inference does not hold. (1) The live proof is
  tools/harness/Test-DebugApi.ps1:484, not Test-SplitWorkspace.ps1:484; Test-SplitWorkspace.ps1 is 477
  lines and contains no palette reference at all. (2) docs/driving-excel.md never mentions the
  palette. Section 6 'What needs CDP' begins at line 1271 and covers trusted mouse input and
  shadow-root reads only; the only 'object browser' string in that file is in the ui.dialogs comment
  at line 296. So the doc does not send the next reader down that road. (3) 'That plumbing is why the
  probe is not in the gate' is unsupported: Test-ObjectBrowser.ps1 launches its own Excel and throws
  when Application.VBE is null (lines 162-177), which is the harder blocker, and
  objbrowser-page-probe.mjs, which verify.ps1:247 does run, itself speaks CDP (its header line 9).
  Also note a straight port needs an explicit waitMs: the eval route's budget defaults to 5000ms
  (AddInSession.DebugApi.cs:1101-1103) while the probe's inner polls run to 10000ms.

##### `no-route-shows-or-hides-the-editor-window` Nothing can hide or re-show the VBE main window through the api, which is the sole reason one harness script still requires the AccessVBOM trust setting

- **Where:** `tools/harness/Test-ObjectBrowser.ps1:179`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** Test-ObjectBrowser.ps1 states the dependency itself at lines 171-176: "This one test genuinely needs
  the VBA project object model: it HIDES and re-shows the editor, and there is no ungated equivalent
  of that - Excel's ribbon command opens the editor but does not close it", then throws if `$null -eq
  $excel.VBE` and drives `$excel.VBE.MainWindow.Visible = $true` at 179, `= $false` at 208 and `=
  $true` again at 213. Test-CloseVbe.ps1:59 needs the same call to reopen the frame after closing it
  with WM_SYSCOMMAND. The api can only look: the windows route
  (src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2610) enumerates the editor's Windows collection
  and reports `window.GetInt32("Type")`, `GetString("Caption")` and `GetBool("Visible")`, and there is
  no setter anywhere in the switch. docs/driving-excel.md:1160 records the consequence in its own
  trust table: "Verifying WITHOUT the add-in | yes" and "New-RenameFixture.ps1 | yes" are listed, but
  the object browser probe's dependency is not, so the doc understates what still needs the box
  ticked.
- **Why:** The add-in is handed the VBE object at OnConnection and needs no trust setting to touch MainWindow,
  so this dependency exists only because the capability was never given a door. It costs the harness
  its best property: a probe that cannot run with AccessVBOM off is a probe the developer skips, and
  Test-ObjectBrowser is already outside the gate (verify.ps1:247 runs objbrowser-page-probe.mjs
  headless instead). The palette's whole lifecycle story - hides with the editor, stays away, returns
  on a summons - therefore has no coverage anyone runs.
- **Change:** Add a route that sets the editor frame's visibility from inside, for example `frame?visible=0|1`
  answering the same shape the windows route reports, driving VBE.MainWindow.Visible on the host
  thread. Test-ObjectBrowser.ps1 then drops its GetActiveObject attach and its trust-setting throw
  entirely, and Test-CloseVbe.ps1 keeps only the WM_SYSCOMMAND close, which is genuinely a window
  operation.
- **Size:** 4 AccessVBOM-gated call sites across 2 probes; removes the last trust-setting dependency in tools/harness
- **Adversary:** tools/harness/Test-ObjectBrowser.ps1:170-177 carries the quoted comment and the '$null -eq
  $excel.VBE' throw verbatim, and drives $excel.VBE.MainWindow.Visible at 179, 208 and 213 to test the
  palette lifecycle. Test-CloseVbe.ps1:53 does the same to reopen the frame. I listed every 'case
  "..."' in the route switch (AddInSession.DebugApi.cs, 60 cases): there is no frame,
  window-visibility or show/hide setter anywhere, and the windows route at 2610-2628 only reads Type,
  Caption and Visible off the Windows collection. VbeCommands.ForName (VbeCommands.cs:229-254) maps no
  window command either, so the command route cannot reach it.
- **Correction applied:** 'Removes the last trust-setting dependency in tools/harness' is false. Five other harness scripts
  reach Application.VBE and so need AccessVBOM: Test-CloseVbe.ps1:36, Test-CloseHiddenPane.ps1:15,
  Test-GhostLocalsPanel.ps1:8, Test-ResizeFollow.ps1:60 and Test-WatchPanel.ps1:60. The doc
  understatement is correspondingly bigger than claimed: the trust table at
  docs/driving-excel.md:1155-1163 names only New-RenameFixture.ps1 and 'Verifying WITHOUT the add-in',
  omitting all six. Separately, the proposed route is itself an undesignated deviation from api rule
  (a): no UI action hides the VBE main window without closing it, so 'frame?visible=0' must be
  designated and documented in the route table, the client method and the code, exactly as the
  pane-close finding proposes for its own route.

##### `no-route-closes-a-native-code-pane` pane close only drives the tab X gate, so the host-originated pane close - the direction two probes are written to test - can only be reached through Application.VBE

- **Where:** `tools/harness/Test-CloseHiddenPane.ps1:37`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** The pane route (src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2919) closes only through the
  surface: `case "close": ... OnModuleCloseRequested(paneModule, DisplayFromProjectId(paneOwner),
  closeAnswer)`, described in its own comment as "Through the same gate the tab's own X uses".
  Test-CloseHiddenPane.ps1 needs the other direction and says so - "Close the HIDDEN pane exactly as
  the host's close path does: through its window" - and reaches it with COM at lines 34-37: `foreach
  ($pane in $vbe.CodePanes) { ... $pane.Window.Close() }`. Test-ResizeFollow.ps1:103-105 needs the
  same thing to reach its second state: `foreach ($window in $vbe.Windows) { if ($window.Type -eq 0) {
  try { $window.Close() } catch {} } }`. Both scripts open with
  `[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')` and `$app.VBE`, so both
  require AccessVBOM. Neither asserts through a route: Test-CloseHiddenPane greps the shim log for a
  `setModules` line that does not name the closed module (line 60).
- **Why:** A native pane closing underneath the surface is a real host event with a real fix behind it (the
  pane tracker only holds panes it can match, so a hidden pane's close changed nothing and the strip
  kept a dead tab). It is the exact class of defect editor parity exists to catch, and the only driver
  for it sits outside the api, needs the trust setting, and asserts by reading a log line instead of
  the tab strip.
- **Change:** Add an action to the existing pane route - `pane?action=closeNative&module=...` - that calls Close
  on the code pane's Window from inside, marked in the route table, the client method and the code as
  the deliberate deviation it is: it is the HOST's close, not the tab's, and that is the point. The
  two probes then drop COM entirely and assert through ui().workspace tabs, which is what the
  developer would see.
- **Size:** 2 probes, both currently AccessVBOM-gated, both asserting through shim-log greps
- **Adversary:** The pane route (AddInSession.DebugApi.cs:2885-2936) offers only open and close, and close routes
  through OnModuleCloseRequested with the comment 'Through the same gate the tab's own X uses' (2922).
  Nothing else in the 60-case switch closes a code pane window.
  tools/harness/Test-CloseHiddenPane.ps1:33-40 carries the quoted comment and '$pane.Window.Close()',
  and Test-ResizeFollow.ps1:103-105 does the '$window.Type -eq 0' close. Both open with
  GetActiveObject plus $app.VBE (CloseHiddenPane:14-15, ResizeFollow:59-60), so both need AccessVBOM,
  and neither asserts through a route.
- **Correction applied:** Two line numbers are off. The pane close case is at AddInSession.DebugApi.cs:2920, not 2919.
  Test-CloseHiddenPane.ps1 is 52 lines long, so its setModules log grep is at line 47, not line 60;
  the substance of the claim (it concludes from a shim-log line rather than the tab strip) is exact.

##### `no-route-creates-a-watch` The Watch panel can be read and asserted but never populated, so its only live coverage fills the native modal with WM_CHAR from a spawned PowerShell helper

- **Where:** `tools/harness/immediate-watch.mjs:165`
- **Kind:** api-coverage / medium effort, claim observed, confidence verified, severity medium
- **Evidence:** immediate-watch.mjs declares the gap in its header at line 10 - "The Watch panel is read-only from
  here by design: watches are added through the editor's own dialog" - and its watch section proves
  it: line 165 `const added = await api.command("addWatch")`, checked only for "the add-watch command
  reaches the editor", then lines 171-177 dismiss whatever it opened with `api.dismiss("Cancel")`. So
  the suite asserts a list exists and never puts a row in it. Test-WatchPanel.ps1 is the only thing
  that does, and it spawns a second PowerShell process that finds the "Add Watch" dialog with
  EnumWindows, grabs `GetDlgItem($dialog, 4853)`, sends one WM_CHAR per character ("Typed, not set:
  the dialog answers 'Empty watch expression' to text planted by WM_SETTEXT") and posts BM_CLICK to
  the OK button, then greps the shim log for "watch: 1 row(s)". The observation half already exists
  first-class: the watches route at AddInSession.DebugApi.cs:816 and the assert vocabulary's
  `watchHas` at line 2127. The action half stops one step short: DialogWatch
  (src/Xlide.Vbe.Shim/Diagnostics/DialogWatch.cs) already enumerates a dialog's children and reads
  their class and text (Dismiss walks ChildrenOf(dialog) at line 145), but it can only post BmClick to
  a Button.
- **Why:** The Watch panel is a shipped surface whose value-tracking behaviour is only ever proven by a probe
  nobody runs, that needs AccessVBOM, that spawns a helper process, and that concludes from a log line
  rather than from the panel. Meanwhile the suite that IS in the gate reaches the honest limit of the
  api and cancels. The route table's watchHas assertion has no way to be made true from a script.
- **Change:** Give DialogWatch a fill capability and expose it as a new route - `fill?caption=Add
  Watch&field=<dialog control id or preceding label>&text=Counter` - typing WM_CHAR into the edit the
  way the helper already proves works, then dismiss("OK"). It stays inside the existing envelope:
  DialogWatch is #if DEBUG and already posts synthetic BM_CLICK, so no product code gains synthetic
  input. immediate-watch.mjs then adds a real watch and asserts through watches() and assert watchHas,
  and Test-WatchPanel.ps1 retires.
- **Size:** 1 gate suite that currently cancels rather than asserts; 1 out-of-gate probe with a spawned helper process and ~45 lines of Win32
- **Adversary:** tools/harness/immediate-watch.mjs:162-168 runs api.command("addWatch") and checks only that the
  command reaches the editor, then 170-177 dismisses whatever it opened with api.dismiss("Cancel"), so
  the suite never puts a row in the panel. Test-WatchPanel.ps1 is the only thing that does: it writes
  a helper script to TEMP and spawns it (93-95), and the helper finds the dialog by EnumWindows (26),
  grabs GetDlgItem($dialog, 4853) (45), sends one WM_CHAR per character with the quoted 'Typed, not
  set' comment (48-51) and posts BM_CLICK 0x00F5 to OK (55); the verdict is a log grep for 'watch:
  [1-9]\d* row' at 135. The observation half exists: the watches route at AddInSession.DebugApi.cs:816
  and watchHas at 2127. DialogWatch.cs is '#if DEBUG' at line 1, walks ChildrenOf(dialog) in Dismiss
  at 143, and can only Win32.PostMessage(child, BmClick, 0, 0) at 153. No other route fills a native
  dialog: the 'type' route at 1675-1733 types into the page's monaco editor via trigger("keyboard",
  "type"), not into a host window.
- **Correction applied:** 'The route table's watchHas assertion has no way to be made true from a script' is overstated.
  Test-WatchPanel.ps1 does make it true from a script (lines 85-99: set the caret, execute control
  1820, let the spawned helper type and click OK); what has no way to make it true is the API. Also
  DialogWatch's child walk in Dismiss starts at line 143, not 145. Note for the owner: the proposal
  moves WM_CHAR keystroke synthesis into shim source. DialogWatch is compiled out of Release, so it is
  not the shipped product, but the 'no synthetic input' rule makes this a judgment call rather than a
  free extension of an existing envelope.

##### `ghost-locals-probe-never-retired` Test-GhostLocalsPanel.ps1 builds its fixture through VBProject, steps with CommandBars.FindControl and concludes from a log grep, all of which routes now do

- **Where:** `tools/harness/Test-GhostLocalsPanel.ps1:10`
- **Kind:** dead-code / small effort, claim derived, confidence verified, severity low
- **Evidence:** Every step of the probe has had a route for some time. It builds the fixture with COM at line 10,
  `$components = $app.ActiveWorkbook.VBProject.VBComponents` then `$components.Add(1)` and
  `AddFromString`, which component?action=add plus writeModule now do from inside
  (AddInSession.DebugApi.cs:2707, and docs/driving-excel.md:1157 lists "Building a fixture | no"
  against the trust setting). It steps with `function Get-Control([int]$id) {
  $vbe.CommandBars.FindControl(1, $id) }` and `(Get-Control 188).Execute()` three times, which is
  command?name=stepInto - VbeCommands.ForName maps "stepInto", "stepOver", "stepOut" and "reset" at
  src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:238-241. It concludes by counting log lines: `$pushes =
  @($since | Where-Object { $_ -match 'locals: \d+ row' })`, where the locals route
  (AddInSession.DebugApi.cs:797) and the assert vocabulary's localsHas (line 2118) answer the same
  question about the panel itself. tools/harness/debugger-features.mjs and step-into-features.mjs
  already drive break, step and parity entirely through the client. Nothing references this file
  except docs/locals-break-investigation.md, which cites its 2026-08-04 result.
- **Why:** It is the same coverage gap seen from the far end: the api arrived and the probe was never retired.
  Left in tools/harness it reads as the way to test the Locals panel, it needs AccessVBOM to run at
  all, and its log-grep assertion passes on a session whose panel renders nothing, because a pushed
  row and a drawn row are different claims. Keeping it also keeps a VBProject call site in the harness
  that the trust-setting story would otherwise be free of.
- **Change:** Fold its one distinct claim - locals values changing across steps, not merely being pushed - into
  debugger-features.mjs using component/writeModule for the fixture, command?name=stepInto for the
  steps and locals()/assert localsHas for the assertion, then delete the script.
  Get-EditorScreenshot.ps1 is the same case and can go the same way: it is referenced by nothing, it
  builds BrokenModule and CleanModule through `$excel.ActiveWorkbook.VBProject.VBComponents` at line
  252, and the picture it exists for is now the capture route, wrapped by Get-Shot.ps1.
- **Size:** 2 orphaned scripts, about 500 lines, carrying 2 of the harness's remaining VBProject call sites
- **Adversary:** Test-GhostLocalsPanel.ps1 reads exactly as described: VBProject fixture build at 10-27, Get-Control
  at 29 with (Get-Control 188).Execute() at 45, 47 and 49, reset 228 at 52-53, and the verdict from
  log greps for 'locals: \d+ row' and 'setLocals' at 63-68. The routes that replace those steps exist:
  component?action=add at AddInSession.DebugApi.cs:2707, VbeCommands.cs maps reset at 236 and
  stepInto/stepOver/stepOut at 238-240, locals at 797 and localsHas at 2118. step-into-features.mjs:92
  already drives api.command("stepInto") and reads api.locals() at 124; debugger-features.mjs enters
  break through api.breakpoint plus a run (84-102) and reads locals at 114. A repo-wide grep finds
  Test-GhostLocalsPanel.ps1 named only by docs/locals-break-investigation.md (lines 24 and 163), and
  Get-EditorScreenshot.ps1 named only inside its own help text (29, 33); its VBProject fixture build
  is at line 252 and Get-Shot.ps1:60 wraps the capture route.
- **Correction applied:** Two details. The probe also uses $app.OnTime(...) at line 36 to provoke the break, which is Excel
  automation rather than a route or a VBE call; the fold-in must replace it with the
  breakpoint-plus-run path debugger-features.mjs already uses, not only the step and locals calls. The
  two scripts total 447 lines (69 + 378), not 'about 500'. VbeCommands maps the four commands at
  236-240, not 238-241.

##### `no-act-opens-or-closes-a-panel` The act table can move a panel between docks but cannot open or close one, and ui.panes omits which side a pane is on, so the split-workspace probe clicks the panes menu through eval

- **Where:** `tools/harness/Test-SplitWorkspace.ps1:270`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** DevSurfaceParts exposes exactly two pane capabilities (ui/editor/src/devsurface.ts:127): `list():
  {name, title, open, permanent}[]` and `moveTo(name, side)`, and the act table has only `dock`
  (devsurface.ts:1235) which calls moveTo. So Test-SplitWorkspace.ps1 opens and closes panels by
  driving the toolbar menu as DOM: line 270 `document.querySelector('#toolbar
  [data-command="openPanes"]').click();` then line 271 `[...document.querySelectorAll("#panes-menu
  .panes-menu-item")].find(i => i.dataset.pane === "problems")`, and line 258 closes one with
  `group.querySelector(".dock-close")`. It also scrapes the answer to "which side is this pane on":
  line 305 `["left","right","top","bottom"].find(s => [...document.getElementById("dock-" +
  s).querySelectorAll(".panel-tab")].some(t => t.dataset.panel === pane))`. The snapshot cannot say:
  `panes: parts.panes.list().map(({name, title, open}) => ({name, title, open}))` at devsurface.ts:461
  drops even the permanent flag, and never had a side. The side is known internally -
  PaneDocks.findPane walks the four sides and returns `{side, group}`
  (ui/editor/src/paneldocks.ts:270) - and the layout route's LayoutScript does answer it, by scraping
  `document.getElementById("dock-" + side)` from C# (AddInSession.DebugApi.cs:444).
- **Why:** Opening and closing a panel is an ordinary thing a developer does and the only way a script can do
  it is by knowing the toolbar's data-command value, the menu's element id and its item class. That is
  three private details for one action, in a probe the gate runs. The observation half is worse than
  missing: it exists twice, once as a DOM scrape in the probe and once as a DOM scrape inside the
  shim's layout script, and neither reads the structure PaneDocks actually holds.
- **Change:** Add act("pane", {name, open:true|false}) calling the same method the panes menu item calls, and
  carry the side on the existing snapshot: `panes: [{name, title, open, permanent, side}]`, taken from
  PaneDocks.findPane rather than from the DOM. The probe then loses six selectors, and the layout
  route's scrape has a typed answer to fall back on.
- **Size:** 6 private selectors in one gate probe; ui.panes gains 2 fields the page already holds
- **Adversary:** DevSurfaceParts.panes exposes exactly list() and moveTo() at ui/editor/src/devsurface.ts:127-131,
  and dock at 1235-1251 is the only act that touches panes; it calls parts.panes.moveTo. The snapshot
  at devsurface.ts:461 is 'parts.panes.list().map(({ name, title, open }) => ({ name, title, open
  }))', dropping permanent and never carrying a side, while PaneDocks.findPane at
  ui/editor/src/paneldocks.ts:269-278 holds {side, group} internally. Test-SplitWorkspace.ps1 does
  close a panel with group.querySelector(".dock-close") at 258, reopen it through '#toolbar
  [data-command="openPanes"]' and '#panes-menu .panes-menu-item' at 270-271, and scrape the side at
  305. That probe is in the gate: verify.ps1:369 runs it.
- **Correction applied:** The observation half is not simply missing. The layout route already answers which side each pane is
  on: LayoutScript at AddInSession.DebugApi.cs:443-455 walks dock-left/right/top/bottom and reports
  each group's tabs with {pane, active} per side. So the probe's line-305 scrape duplicates an answer
  a route already gives, and the genuinely absent capability is the open/close action, not the side.
  Adding side to ui.panes remains worthwhile because the layout route gets it by scraping the DOM from
  C# rather than from PaneDocks, but it is a de-duplication, not a new answer.

##### `route-audit-runs-one-way-only` audit-routes.mjs proves every route is documented and driven, but nothing proves the harness is not going around the api, so each fallback above can be added without the gate noticing

- **Where:** `tools/harness/audit-routes.mjs:124`
- **Kind:** api-coverage / small effort, claim derived, confidence supported, severity low
- **Evidence:** The audit asks three questions of each route and no question of the harness: `if
  (!inReference(route))`, `if (!inDriving(route))`, `if (!inClient(route))`, then `const driven =
  isDriven(route)` matched against a corpus of tools/harness (lines 148-164). Its one exemption list,
  NOT_DRIVEN_ON_PURPOSE at line 124, holds a single entry (drainfinalizers) with a written reason, and
  the audit fails if an excused route turns out to be driven - a good design, applied in only one
  direction. The corpus it builds at line 137 is every .mjs and .ps1 in tools/harness, so the raw
  material for the reverse question is already read into memory: a grep of that same corpus for
  GetActiveObject, `.VBE`, VBProject, AccessibleObjectFromWindow, `new WebSocket` and json/list is
  what produced the findings above.
- **Why:** Every gap in this report is invisible to the gate. A new suite can reach Excel through COM or CDP,
  or scrape a value the render already computed, and twelve green steps say nothing about it - which
  is how Test-GhostLocalsPanel and Get-EditorScreenshot stayed in the tree after their routes arrived,
  and how the sync dialog came to be driven by eleven eval calls in a gate suite. The audit's own
  exemption design is the fix already proven here: a fallback that must be named with a reason cannot
  rot quietly.
- **Change:** Extend audit-routes.mjs with a second pass over the corpus it already reads: flag GetActiveObject,
  Application.VBE, VBProject, AccessibleObjectFromWindow, raw DevTools use (new WebSocket, /json/list,
  Runtime.evaluate) and api.eval/api.ask calls, each allowed only by an OFF_THE_API list naming the
  file and the reason, in the same shape as NOT_DRIVEN_ON_PURPOSE. Seed it with the entries this
  report establishes, so the list starts honest and shrinks as routes land. I did not read every suite
  in tools/harness (see coverage), so the seed list should be built from the audit's own first run
  rather than from this report alone.
- **Size:** unmeasured; the audit already reads the whole harness corpus, so the second pass is a regex list and an exemption map
- **Adversary:** I read tools/harness/audit-routes.mjs end to end. It asks three questions per route (inReference,
  inDriving, inClient at 150-152) plus isDriven at 154, and nothing else; there is no pass over the
  corpus asking what the harness does off the api. NOT_DRIVEN_ON_PURPOSE holds one entry,
  drainfinalizers, with a written reason (128-135), and the symmetric check at 162-164 fails an excuse
  that turns out to be driven. The corpus at 137-140 is every .mjs and .ps1 in tools/harness except
  the client and the audit itself, so the raw material for a reverse question is already in memory.
  Nothing else enforces it: the gate steps in verify.ps1 do not include any harness-side policy check.
- **Correction applied:** The proposal is only workable on the COM and CDP axis. GetActiveObject appears in exactly six
  harness files (Test-CloseHiddenPane, Test-CloseVbe, Test-GhostLocalsPanel, Test-ObjectBrowser,
  Test-ResizeFollow, Test-WatchPanel), which is a six-entry seed. Flagging api.eval/api.ask/Page is
  not: the harness holds roughly 60 such calls, 29 in Test-SplitWorkspace.ps1 alone, 11 in
  module-sync.mjs, 9 in Test-DiscardProblems.ps1, 5 in Test-Churn.ps1, and the eval route's own
  contract invites the one-off use (AddInSession.DebugApi.cs:1094-1099, 'Reach for this when the
  question is genuinely new'). An OFF_THE_API list over eval would start with dozens of entries and
  would not shrink, which is the failure mode the drainfinalizers design exists to avoid. Scope the
  second pass to GetActiveObject, Application.VBE, VBProject, AccessibleObjectFromWindow and raw
  DevTools use.

### Page features with no driver or no observer

_12 findings, from the `api-ui-surface` finder._

##### `search-scoped-matches-unobserved` NO OBSERVER: ui.search.matches reports 0 for a project-scope search, reading a field the results list does not use

- **Where:** `ui/editor/src/searchwidget.ts:473`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** state() builds the ui route's search block from the module-scope match array: `matches:
  this.matches.length,` (searchwidget.ts:473) and `current: this.current,` (:474). For any scope other
  than module, scopeChanged() empties exactly those fields: `this.decorations.clear(); this.matches =
  []; this.current = -1; this.counter.textContent = "";` (:570-573). The rows the user actually sees
  in that mode are built by showSearchResults(), which returns early for module scope (`if (id <
  this.pendingSearchId || this.scope() === "module") { return; }`, :199) and appends its own DOM count
  note `${matches.length}${truncated ? "+" : ""} match...` (:217-219) without ever touching
  this.matches. debug-api.md:69 advertises the ui route as carrying "the find/replace widget's query,
  scope and match count".
- **Why:** docs/driving-excel.md:306 shows the canonical drive as `api.act("search", { query: "Recalculate",
  scope: "project" })`. Any assertion of the form search.matches > 0 after that call is asserting on a
  field that is unconditionally 0 in that mode, and any assertion of matches === 0 passes whatever the
  panel shows. A project search that returned nothing, returned the wrong module's hits, or never
  rendered its rows is indistinguishable from one that worked.
- **Change:** Extend SearchWidget.state() with the scoped result the panel is holding: the count showSearchResults
  was handed, whether it was truncated, the replaced count, and the grouped rows (module, workbook,
  line). Add them as new fields on UiSnapshot["search"] so no existing field changes meaning, and read
  them from the same values passed into showSearchResults rather than from the .search-row DOM.
- **Size:** unmeasured; one added record on the existing search snapshot, roughly 15 lines in searchwidget.ts plus the UiSnapshot type
- **Adversary:** Read ui/editor/src/searchwidget.ts:466-477: state() returns `matches: this.matches.length` and
  `current: this.current`. scopeChanged() at :563-576 takes the non-module branch and does
  `this.decorations.clear(); this.matches = []; this.current = -1; this.counter.textContent = ""`, and
  open() calls scopeChanged() at :529, so any project/workbook-scope open empties exactly those two
  fields. queryChanged() at :557-561 only searches when scope()==="module", and showSearchResults() at
  :195-255 returns early for module scope and builds its own DOM note without touching this.matches.
  So for a non-module scope, ui.search.matches is structurally 0 and ui.search.current is -1
  regardless of what the panel drew. docs/debug-api.md:69 does advertise the ui route as carrying the
  widget's match count, and docs/driving-excel.md:300 (`ui.search; // open, query, scope, matches,
  current`) and :306 (`api.act("search", { query: "Recalculate", scope: "project" })`) are the
  canonical pairing. Verified end to end; no other listener on findInput exists (only
  searchwidget.ts:319).

##### `search-replace-no-driver` NO DRIVER: Replace, Replace All and Find All cannot be triggered, and act search never runs a non-module search at all

- **Where:** `ui/editor/src/devsurface.ts:943`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity high
- **Evidence:** The whole search vocabulary is open/find/close: `search: (args) => { const shut = flag(args.close,
  false); if (shut) { parts.search.close(); ... } const query = ...; if (query === null) {
  parts.search.open(...) } parts.search.find(query, {...})` (devsurface.ts:943-965), and
  DevSurfaceParts.search declares only `state`, `find`, `open`, `close` (:120-125). The widget's four
  other commands are DOM-only: `this.replaceButton.addEventListener("click", () =>
  this.replaceCurrent()); this.replaceAllButton.addEventListener("click", () => this.replaceAllRun());
  this.findAllButton.addEventListener("click", ...)` and the prev/next buttons
  (searchwidget.ts:322-333). find() ends with `this.findInput.dispatchEvent(new Event("input"))`
  (:503), and the input handler is `queryChanged(): void { if (this.scope() === "module") {
  this.findInModule(false); } }` (:557-561) - nothing else. A scoped search only starts from Enter or
  Find All (:326-333, :350-357). Nothing in tools/harness or docs mentions replaceAll; the only repo
  hits for it are the page->shim message case at src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1904 and
  the handler behind it.
- **Why:** Replace All is the most destructive thing on the surface: it rewrites text across every module of a
  project through handlers.replaceAll, and it has no driver, no route and no suite. The scoped search
  behind it is equally unreachable, so `act search scope=project` types a query and returns did:true
  having searched nothing - a probe that then reads matches (also 0, see
  search-scoped-matches-unobserved) concludes the feature works.
- **Change:** Widen the search act rather than adding routes: accept
  `run=find|findAll|next|previous|replace|replaceAll` and a `replacement` argument, and have
  DevSurfaceParts.search expose the widget methods those press (runScoped, showModuleResults, next,
  previous, replaceCurrent, replaceAllRun) so each goes through the same method the button's click
  handler calls. Add the argument to the act row in docs/debug-api.md, the act list in
  docs/driving-excel.md, and the client's act helper.
- **Size:** unmeasured; 6 methods to expose, about 30 lines in devsurface.ts and searchwidget.ts
- **Adversary:** devsurface.ts:943-962 is the whole search act: close, open, or
  find(query,{scope,matchCase,wholeWord}). DevSurfaceParts.search at :120-125 declares only
  state/find/open/close. find() ends at searchwidget.ts:503 with an `input` event, whose only handler
  is queryChanged() (:319, :557-561), which searches only in module scope; a scoped search starts only
  from Enter (:350-357) or Find All (:326-333), and replaceCurrent/replaceAllRun hang off click
  listeners at :324-325 with no method exposed. The act keys in devsurface.ts (listed at 525-1467)
  contain no replace/findAll/next/previous. Repo grep for replaceAll finds only bridge.ts:301/1380,
  main.ts:417, searchwidget.ts and the shim side; nothing in tools/harness or docs. Verified the shim
  end is real state change, not a stub: EditorSurface.cs:1904-1928 raises ReplaceAllRequested, wired
  at AddInSession.cs:1104 to OnReplaceAllRequested (AddInSession.cs:2742). So a cross-module rewrite
  has no driver, no route and no suite, and `act search scope=project` genuinely types a query and
  searches nothing while answering did:true.

##### `tab-geography-no-driver` NO DRIVER: tab drag, the tab context menu and middle-click close are unreachable; moveTab is private where movePaneTo is public

- **Where:** `ui/editor/src/workspace.ts:996`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity low
- **Evidence:** The drop handler's target is `private moveTab(id: DocumentId, from: EditorGroup, destination: {
  group?: EditorGroup; index?: number; split?: { of: EditorGroup; direction: Exclude<DropZone,
  "center"> } }): void` (workspace.ts:996-1000), reached only from installTabDrag (:1177). The strip
  also owns two other user paths with no method behind them: `strip.addEventListener("auxclick", ...)`
  closing on middle button (:1126-1136) and `strip.addEventListener("contextmenu", ...)` offering `{
  label: "Close" }`, `{ label: "Close Others", enabled: others.length > 0, ... }`, `{ label: "Close
  All", ... }` (:1138-1163). The act vocabulary reaches tabs only through closeActive, activate,
  cycleTab and split (devsurface.ts:525, 609, 648, 656). Panes got the opposite treatment:
  DevSurfaceParts.panes carries `moveTo(name, side): boolean` described as "Moves a pane to a dock
  side, through the method a real drop calls" (devsurface.ts:129-130).
- **Why:** Tab geography is the surface's richest gesture - reorder within a strip, move to another group,
  split by dropping on a zone - and none of it can be driven, so every defect in it is found by hand.
  Close Others across several unsaved modules is a designed queueing behaviour (shell.ts:769-775
  comments the one-at-a-time queue) that no test can even start. The dock act exists because
  synthesising the drag tests the synthesiser; the same argument applies here and was not followed
  through.
- **Change:** Make moveTab public (or add a thin `dropTab(id, {group, index, split})` beside it) and add an act
  that calls it, mirroring dock; add a `tabMenu` act that fires contextmenu on the tab row the way
  treeMenu does for tree rows, so the existing chooseMenuItem can pick Close Others/Close All.
  Document both in the act row of docs/debug-api.md and the act list in docs/driving-excel.md.
- **Size:** unmeasured; one visibility change plus roughly 40 lines of act code
- **Adversary:** workspace.ts:996-1000 is `private moveTab(...)`; the strip's auxclick close is at :1126-1136 and the
  contextmenu with Close / Close Others / Close All at :1138-1164; devsurface.ts exposes only
  closeActive (:525), activate (:609), cycleTab (:648) and split (:656) for tabs, while panes get
  moveTo at :127-131 driven by the dock act at :1235-1251. Two details in the evidence are wrong, and
  one consequence is overstated.
- **Correction applied:** moveTab is NOT reached only from installTabDrag: workspace.ts:774-779 splitActive() calls it, and
  the `split` act (devsurface.ts:656-660) drives that, so the split-of-own-group destination is
  exercised. Middle-click close is not wholly unreachable either -
  tools/harness/close-confirm-page-probe.mjs:63-64 dispatches a synthetic auxclick against the built
  page and asserts the dirty-close question, and the same probe deliberately stands in for the Close
  Others queue (its comment at :14-16, checks at :50-57), so the one-at-a-time queueing shell.ts:766
  describes IS covered, by a stand-in rather than by the menu. What genuinely has no driver anywhere:
  the drop destinations that carry {group, index} or a split of ANOTHER group, and the tab context
  menu itself (its three items are built inline in the listener and no act right-clicks a tab the way
  treeMenu right-clicks a tree row).

##### `toolbar-commands-no-driver` NO DRIVER: no act presses a toolbar button, so two suites hand-roll querySelector clicks to reach one

- **Where:** `ui/editor/src/toolbar.ts:35`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** COMMANDS holds about thirty entries dispatched by id - `{ id: "openSync", target: "editor", ... }`,
  `{ id: "openPanes", ... }`, `{ id: "openHelp", ... }`, `{ id: "objectBrowser", target: "host", ...
  }`, plus the comment suite, indent/outdent, fold/unfold, gotoLine and quickCommand
  (toolbar.ts:35-92), each wired with `button.addEventListener("click", () => run(command))` (:166)
  into the shell handler at main.ts:468-508. The act vocabulary has no entry that takes a command id;
  only settings and sponsors are wired individually (devsurface.ts:903-911), and format/undo hard-code
  two monaco action ids (:1392-1394, :1206-1208). The consequence is in the harness: `await
  api.eval('(() => document.querySelector("button[data-command=openSync]").click())()')`
  (tools/harness/module-sync.mjs:248) and `const button =
  document.querySelector('[data-command="objectBrowser"]')`
  (tools/harness/objbrowser-live-probe.mjs:105).
- **Why:** The toolbar is the product's main command surface and it is drivable only by raw eval, which is the
  practice devsurface.ts was written to end (its header names hand-rolled querySelectorAll as the most
  expensive habit in the repo). A test that clicks a button it found itself keeps passing after the
  button is renamed, removed, disabled by the needsBreak gate, or scrolled out of the strip, because
  it never went through buildToolbar's dispatch.
- **Change:** Add an act that takes a command id, looks it up in COMMANDS, and calls the same `run(command)`
  callback buildToolbar hands the button - refusing a command that is not on the strip or is disabled,
  and answering which. That one action unlocks openSync, openPanes, openHelp, objectBrowser and the
  whole editor-action cluster in one place. Document it in both api docs and give it a client method.
- **Size:** unmeasured; about 25 lines in devsurface.ts plus a shell accessor for the command callback
- **Adversary:** toolbar.ts:35-92 is COMMANDS, about thirty entries including openSync (:87), openPanes (:88),
  openHelp (:90), objectBrowser (:85) and the editor-action cluster (:68-83); the click wiring is
  `button.addEventListener("click", () => run(command));` at toolbar.ts:165. No act takes a command id
  - the act keys in devsurface.ts are closeActive, answerCloseConfirm, activate, cycleTab, split,
  expandWorkbook, unfoldModule, treeMenu, renameModule, treeAdd, menuBar, chooseMenuItem,
  answerRemoveConfirm, settings, sponsors, closeDialogs, key, focusEditor, search, hover, completions,
  signature, quickFixes, timeFeature, references, definition, rename, undo, dock, press, insert,
  backspace, format, bookmark. `press` (devsurface.ts:1289-1326) is a keyboard key, not a button, and
  the shim's `command` route is VbeCommands.ForName, the NATIVE editor's command, not a page toolbar
  id. The two hand-rolled clicks are real: tools/harness/module-sync.mjs:248 and
  tools/harness/objbrowser-live-probe.mjs:105.
- **Correction applied:** The wiring line is toolbar.ts:165, not :166. Also worth adding to the evidence:
  tools/harness/menu-bar.mjs:157 scrapes `#toolbar [data-command]` to assert the strip's contents, so
  three harness sites reach the toolbar by selector and none through its dispatch.

##### `statusbar-no-observer` NO OBSERVER: the status bar - caret readout, module name and the notice line - is in no snapshot

- **Where:** `ui/editor/src/shell.ts:747`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** The shell owns three status fields (`statusPosition`, `statusModule`, `statusNotice`,
  shell.ts:155-157, 200-202), written by `notify(text) { this.statusNotice.textContent = text;
  this.statusNotice.classList.add("visible"); ... }` (:747-760) and `this.statusPosition.textContent =
  `Ln ${line}, Col ${column}`` (:877). UiSnapshot declares workspace, explorer, panes, dialogs,
  waiting, focus, settings, emptyViewShown, longTasks, census, search, bookmarks, at
  (devsurface.ts:38-93) and nothing else; state() matches that list (:458-475). notify is the only
  feedback for a whole class of outcomes: `bridge.shell?.notify(message.text)` for host notices
  (bridge.ts:1125), the tree rename summary (main.ts:526-528), the F2 rename summary
  `bridge.shell?.notify(renameSummary(answer, newName))` (main.ts:833), and Undo Rename's result
  (main.ts:1127-1129).
- **Why:** notify is documented in the code as the reply to actions that were legitimately declined - the
  deliberate alternative to a dialog. Nothing can read it, so "the refusal was shown to the developer"
  is untestable, and the three page-local notify calls never cross the transport either, so even the
  messages route cannot see them. A rename that silently reports nothing looks identical to one that
  reported correctly.
- **Change:** Add a `status: { position, module, notice }` block to UiSnapshot, read from the shell's own last-set
  values rather than from textContent, plus a flag for whether the notice is still within its
  five-second visible window. Mention the new field in the ui row of docs/debug-api.md.
- **Size:** unmeasured; three fields, about 15 lines
- **Adversary:** shell.ts:155-157 declares statusPosition/statusModule/statusNotice, bound at :200-202; notify() at
  :747-760 sets the notice text, adds .visible and clears both after 5000ms; setPosition() at :876-878
  writes `Ln x, Col y`. UiSnapshot (devsurface.ts:38-93) and state() (:458-475) carry workspace,
  explorer, panes, dialogs, waiting, focus, settings, emptyViewShown, longTasks, census, search,
  bookmarks, at - no status block. The four notify call sites are bridge.ts:1125 (host notice),
  main.ts:526-530 (tree rename summary), main.ts:833 (F2 rename summary) and main.ts:1127-1129 (Undo
  Rename).
- **Correction applied:** One layer of the claim is too strong. Host-originated notices DO cross the transport as a `notice`
  message (EditorSurface.cs:839 sends NoticeMessage("notice", text); bridge.ts:1124-1126 receives it),
  and the `messages` route snapshots traffic in both directions (AddInSession.DebugApi.cs:555-564), so
  that class of refusal is observable at the transport. What no route can see is (a) that the shell
  actually rendered it and (b) the three page-local notify calls, which never leave the page.

##### `pane-open-close-no-driver` NO DRIVER: panes cannot be shown or hidden, though shell.paneVisibility already offers setOpen and the immediate pane's visibility gates host polling

- **Where:** `ui/editor/src/devsurface.ts:127`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** DevSurfaceParts.panes declares only `list(): {name, title, open, permanent}[]` and `moveTo(name,
  side): boolean` (devsurface.ts:127-131), and the only pane act is dock, which calls
  `parts.panes.moveTo(pane, side)` (:1247). The object main.ts actually passes carries more:
  paneVisibility() returns `{ list, moveTo, setOpen: (name, open) => { if (open) {
  this.docks.open(name); } else { this.docks.close(name); } } }` (shell.ts:349-362) over
  PanelDocks.open/close (paneldocks.ts:222, 247). Six seats exist - explorer, properties, problems,
  immediate, locals, watch (shell.ts:268-278). Their visibility is not cosmetic on the host side:
  `OnPanelChanged(string name, bool open) { if (name == "immediate") { _watchingImmediate = open;
  UpdatePolling(); } }` (src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3659-3666).
- **Why:** Showing and hiding a pane is an everyday action with a host-side consequence - the immediate poll
  interval - and no api can perform it, so the transition that starts and stops that polling is never
  exercised. The locals and watches routes read shim state, so a suite can assert on rows the
  developer cannot see because the pane is closed, and never notice.
- **Change:** Add `setOpen` to the DevSurfaceParts.panes interface (the concrete object already has it) and an act
  that calls it, answering the pane's open state afterwards from panes.list() so the answer is the
  outcome rather than the request. Add it to the act row in docs/debug-api.md, the act list in
  docs/driving-excel.md, and the client.
- **Size:** unmeasured; one interface line plus about 15 lines of act code
- **Adversary:** DevSurfaceParts.panes (devsurface.ts:127-131) declares only list() and moveTo(), and dock
  (:1235-1251) is the only pane act. shell.paneVisibility() at shell.ts:349-362 does return a third
  member, `setOpen(name, open)` over docks.open/close, and PanelDocks.close/open are real at
  paneldocks.ts:222-243 and :246-265. The six seats are at shell.ts:268-278. The host consequence is
  real: shell.ts:293 wires visibilityChanged to handlers.panelChanged, main.ts:518 to
  bridge.panelChanged, and AddInSession.cs:3659-3666 flips _watchingImmediate and calls UpdatePolling.
- **Correction applied:** The strongest consequence is overstated. Visibility is emitted from every dock render
  (paneldocks.ts:427-428 calling emitVisibility at :936-950), so the existing `dock` act already flips
  a pane's visible/invisible state and therefore CAN start and stop the immediate polling indirectly -
  for instance by docking another pane into the immediate pane's group. What has no driver is the
  pane's own open/close (setOpen, and the Panes menu behind the openPanes toolbar button, itself
  undrivable per toolbar-commands-no-driver).

##### `renamemodule-act-skips-shell` DEVIATION: act renameModule calls the bridge directly, skipping the shell handler that reports the result, while the route table says it is the path the tree's Rename item runs

- **Where:** `ui/editor/src/devsurface.ts:762`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** The act is `const answer = await bridge.requestModuleRename(module, workbook, newName);`
  (devsurface.ts:762). The tree's Rename item runs the shell handler main.ts wires: `renameModule:
  (name, workbook, newName) => { void bridge.requestModuleRename(name, workbook,
  newName).then((answer) => { bridge.shell?.notify(answer.refused ?? `Renamed ${name} to ${newName}:
  ${answer.replaced} mention...`); }); }` (main.ts:524-530). So the act reproduces the request and not
  the handler around it. Its own comment claims otherwise - "The prompt is the only thing skipped ...
  Everything after it is the code the menu item runs" (devsurface.ts:745-746) - and so does the route
  table: "renameModule is the product's rename ... the one the tree's Rename item runs"
  (docs/debug-api.md:70).
- **Why:** The gap is exactly the half a user sees: the notice saying how many mentions were replaced. A test
  driving renameModule reports did:true with the summary in `detail`, taken from the promise the act
  awaited itself, so a shell handler that stopped notifying - or was never called - passes. The doc
  claiming full parity is what makes the false negative silent; a reader has no reason to add the
  missing check.
- **Change:** Either route the act through the shell's renameModule handler (pass it into DevSurfaceParts the way
  openSettings is) so the notice fires, or keep the direct call and designate it in all three places:
  the act's own comment, the act row in docs/debug-api.md, and the client method, each saying that the
  status notice is not raised.
- **Size:** unmeasured; one handler reference or three doc edits
- **Adversary:** devsurface.ts:751-772: the act awaits `bridge.requestModuleRename(module, workbook, newName)`
  directly and formats its own detail string. The tree's Rename runs the shell handler at
  main.ts:524-530, which calls the same bridge request and then `bridge.shell?.notify(...)` with the
  refusal or the mention summary. So the act does reproduce the request without the handler wrapped
  round it, and its own comment at devsurface.ts:745-746 ("The prompt is the only thing skipped ...
  Everything after it is the code the menu item runs") and the route table at docs/debug-api.md:70
  ("the one the tree's Rename item runs") both claim more than the code does. The deviation is real
  and undesignated.
- **Correction applied:** The practical loss is smaller than stated: the notice is unreadable by any route today (see
  statusbar-no-observer), so routing the act through the shell handler would not by itself make the
  missing half assertable. The defect that stands on its own is the false parity claim in the comment
  and the route table.

##### `references-dialog-no-driver` DEVIATION and NO DRIVER: act references answers the lookup but never opens the references dialog, and the deviation is designated only in the code

- **Where:** `ui/editor/src/devsurface.ts:1117`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** The act ends at the lookup: `const answer = await parts.referencesAt({ line, column }); ... return {
  did: answer.found.length > 0, detail: `${answer.found.length} reference(s) to ${answer.word}`, data:
  answer.found };` (devsurface.ts:1117-1134), over the wiring `referencesAt: async ({ line, column })
  => { ... return model ? referencesAt(bridge, model, new monaco.Position(line, column)) : null; }`
  (main.ts:943-947). The user action does one thing more: the Shift+F12 action and its context-menu
  entry run showReferences (main.ts:1139-1146), which calls the same lookup and then
  `openReferencesDialog(word?.word ?? "this symbol", found, { navigate: ... })` (main.ts:210-213). The
  act's comment designates the difference ("The dialog is opened by the same lookup",
  devsurface.ts:1114-1115), but `references` appears in neither the do= list of the act row in
  docs/debug-api.md:70 nor any act example in docs/driving-excel.md (which lists dock, press, hover,
  completions, quickFixes, definition, rename and the rest at lines 304-379, 1108-1109).
- **Why:** Find All References exists because the editor's own peek window cannot render a module with no tab
  open (main.ts:1135-1138), so the dialog's rendering IS the feature. Nothing opens it, nothing reads
  its rows, and the act that carries its name leaves no dialog on screen - so ui.dialogs after a
  references act shows the same surface as before, and a broken dialog passes. Its absence from both
  docs means a reader cannot tell the act stops short.
- **Change:** Add an `open=1` argument to the references act that calls the same showReferences the menu entry
  runs, so the dialog stands afterwards and ui.dialogs sees it; keep the data-only default. List
  `references` in the act row of docs/debug-api.md and in the act examples of docs/driving-excel.md,
  naming which layer each form stops at.
- **Size:** unmeasured; one argument plus one parts entry, about 20 lines
- **Adversary:** devsurface.ts:1117-1135 ends at the lookup and returns the found list as data; parts.referencesAt is
  wired at main.ts:944-948. The user path does more: main.ts:1139-1147 registers
  xlide.findAllReferences (Shift+F12, context menu) whose run calls showReferences, and showReferences
  at main.ts:203-213 calls the same referencesAt and then openReferencesDialog. Repo grep for
  openReferencesDialog returns only referencesdialog.ts:29 and main.ts:73/210 - nothing in tools or
  docs opens it. The doc gap checks out: `references` is absent from the do= list in
  docs/debug-api.md:70 and from docs/driving-excel.md (its only hit there is :296, the word inside a
  ui.dialogs comment). audit-routes.mjs audits ROUTES, not act actions
  (tools/harness/audit-routes.mjs:34-58 parses `switch (request.Route)` cases), so nothing in the gate
  catches this.

##### `sync-dialog-no-driver` NO DRIVER and NO OBSERVER: the import/export dialog can only be driven by hand-rolled eval, and the suite doing it has already gone red for that reason

- **Where:** `ui/editor/src/syncdialog.ts:708`
- **Kind:** api-coverage / large effort, claim observed, confidence verified, severity medium
- **Evidence:** The dialog is opened from the toolbar only - `if (command.id === "openSync") { openSyncDialog((args,
  body) => bridge.requestSync(args, body), ...) }` (main.ts:478-483) - and its Apply composes its own
  request: `const apply = document.createElement("button"); ... apply.addEventListener("click", () =>
  void applyPlan());` (syncdialog.ts:377-381) and `request({ action: "apply", direction, folder, mode:
  modeSelect.value }, ...)` over the ticked rows (:708-726, per-row ticks at :481-486). No act names
  it and ui reports only that some aria-modal is up (devsurface.ts:182-192). What the harness does
  instead: `await api.eval('(() =>
  document.querySelector("button[data-command=openSync]").click())()')`, then sets the folder by
  assigning `.value` and dispatching a change event, then clicks
  `.sync-direction[data-direction=import]` and finally finds the Apply button by its text
  (tools/harness/module-sync.mjs:248-289), with a comment recording that waiting on scraped rows made
  the suite "go red on one planner and green on the other for a reason that was purely the harness
  (2026-08-10)".
- **Why:** The sync route is honestly the same HandleSync call the dialog reaches (AddInSession.cs:254, 491),
  so everything the route proves is about the planner and nothing is about the dialog: which rows it
  ticked, whether it redrew for the new direction, what its status line said. That half is covered by
  a scrape that has already produced one false failure, and a scrape cannot tell a stale render from
  the state.
- **Change:** Give the dialog the treatment the tree got: an act that opens it through the toolbar command
  callback, acts for setting folder/direction/mode through the dialog's own handlers rather than by
  assigning input values, an act for ticking a row by id, and one for pressing Apply. Report the drawn
  plan (rows, ticks, status text) as a new block on the ui snapshot read from the dialog's own item
  list. Then module-sync.mjs drops its eval scripts.
- **Size:** unmeasured; a 768-line dialog with roughly 6 controls to expose
- **Adversary:** openSyncDialog is reached only from the toolbar branch at main.ts:478-483; Apply is built and wired
  at syncdialog.ts:375-380 and composes its own request at :707-711 (`request({ action: "apply",
  direction, folder, mode: modeSelect.value }, [...ticked].join("\n"))`) over per-row ticks wired at
  :481-492. No act names the dialog and ui reports only the aria-modal presence
  (devsurface.ts:182-192). The harness does exactly what the finding says: module-sync.mjs:248 clicks
  the toolbar button by selector, :250-251 assigns #sync-folder.value and dispatches change, :271
  clicks .sync-direction[data-direction=import], :280 finds Apply by textContent, and the comment at
  :258-268 records the scrape-caused red/green split on 2026-08-10.

##### `properties-pane-no-coverage` NO DRIVER and NO OBSERVER: the Properties pane edits real component state and nothing in the api can drive it, read it, or name it

- **Where:** `ui/editor/src/shell.ts:539`
- **Kind:** api-coverage / medium effort, claim observed, confidence verified, severity medium
- **Evidence:** The pane writes through two controls - `this.handlers.editProperty(this.propertiesComponent,
  property.name, select.value)` for a choice (shell.ts:492) and
  `this.handlers.editProperty(this.propertiesComponent, property.name, input.value)` for a text field
  (:539) - reaching bridge.editProperty (main.ts:522), the page-to-shim case `editProperty`
  (src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1943-1953) and the wired handler
  `_editorSurface.PropertyEditRequested = OnPropertyEdit;`
  (src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1114). A repo-wide grep for editProperty or
  selectComponent across tools/harness, docs and src/Xlide.Vbe.Shim/AddIn returns nothing outside that
  wiring: no route, no act, no client method, no doc row, no suite. UiSnapshot has no properties block
  (devsurface.ts:38-93).
- **Why:** This is a whole pane, one of six seats (shell.ts:270-273), whose edits change component state on the
  host, and it is the only user-visible surface in the product with zero api presence in either
  direction. A regression here - a row that stops rendering, an edit that never crosses, a value
  written to the wrong component - is invisible to every automated check that exists.
- **Change:** Report the pane in the ui snapshot from the rows the shell holds (component, and for each property
  name, value, kind, editable), and add an act that calls the same handlers.editProperty a control's
  change reaches. Document the field in the ui row and the action in the act row of docs/debug-api.md,
  and add examples to docs/driving-excel.md.
- **Size:** unmeasured; one snapshot block plus one act, about 40 lines
- **Adversary:** shell.ts:492 and :539 both call this.handlers.editProperty(this.propertiesComponent, ...) from the
  select's and the input's handlers; main.ts:522-523 wires editProperty and selectComponent to the
  bridge (bridge.ts:647-653), the shim receives them at EditorSurface.cs:1583 and :1943, and
  AddInSession.cs:1114 wires PropertyEditRequested. A repo-wide grep for editProperty, selectComponent
  and propertiesComponent across ui, tools, docs, src, installer and .github returns nothing outside
  that wiring - no route in the DebugApi switch (I listed all ~65 cases), no act key, no client method
  in tools/harness/xlide-api.mjs, no doc row, no suite. UiSnapshot has no properties block.

##### `palette-surface-no-ui-act` NO OBSERVER and NO DRIVER: ui and act are hard-wired to the editor surface, so the object browser palette is reachable only by raw eval

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1054`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity medium
- **Evidence:** Both routes pass a null surface: `var ui = RunPageScript($"window.xlideUi.state({arguments})", null,
  WaitMilliseconds(request, 5000));` (:1054) and `var act =
  RunPageScript($"window.xlideUi.act({quotedName}, {arguments.ToJsonString()})", null,
  WaitMilliseconds(request, 8000));` (:1080-1083), while RunPageScriptOnce resolves the target as `var
  browser = surface == "palette" ? _browserPalette?.Browser : _editorSurface?.Browser;` (:220). Only
  eval and capture take the surface argument (:1100-1103, the capture route's window argument). The
  palette page is a 568-line surface with its own search box, tree, member list, detail pane, Escape
  handling and arrow-key resizing (ui/editor/src/objectbrowser.ts:274-297, 384, 500), and the probe
  that exercises it reaches in with document.querySelector through eval
  (tools/harness/objbrowser-live-probe.mjs:98-111).
- **Why:** The second-largest page in the product has no snapshot and no action vocabulary, so every assertion
  about it is a scrape of the render - the exact failure mode devsurface.ts was built to remove, and
  the one that cannot distinguish a stale render from the state.
- **Change:** Accept the existing `surface` argument on the ui and act routes (default unchanged) and install a
  small xlideUi on the palette page with its own state() and act() covering the search box, the
  selected node and the detail pane. Route shapes stay stable: this is a new argument, not a rename.
- **Size:** unmeasured; two argument reads in the shim plus a palette-side devsurface
- **Adversary:** AddInSession.DebugApi.cs:1054 passes null for the surface on ui, and :1080-1083 does the same on
  act; RunPageScriptOnce picks the target at :220 with `surface == "palette" ?
  _browserPalette?.Browser : _editorSurface?.Browser`. The palette really is a separate document with
  no xlideUi: main.ts:1242-1244 routes `?view=objbrowser` to bootObjectBrowserPage BEFORE boot, and
  installDevSurface (which installs globalThis.xlideUi at devsurface.ts:1462) runs only inside boot.
  objbrowser-live-probe.mjs:98-111 reaches the palette by document.querySelector through eval, as
  claimed.
- **Correction applied:** `await` also takes the surface argument (AddInSession.DebugApi.cs:1119,
  `request.Query.TryGetValue("surface", out var awaitSurface)`, documented in the await row of
  docs/debug-api.md), so the surface-aware routes are eval, await and capture, not just eval and
  capture.

##### `undorename-menu-item-no-driver` NO DRIVER: the Undo Rename context-menu action is unreachable because the route reverses the rename in the shim instead

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2938`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** The route body is `case "undoRename": UndoRename(0); return ...DebugCommandReply(true, 0)...` with
  the comment "The same path the editor's own Undo Rename takes. Here so a probe can prove a rename is
  reversible without driving the page" (:2938-2946). The user's path is an editor action:
  `editor.addAction({ id: "xlide.undoRename", label: "Undo Rename", contextMenuGroupId:
  "1_modification", ..., run: async () => { const answer = await bridge.requestRenameUndo();
  bridge.shell?.notify(answer.refused ?? `Rename put back: ...`); } })`
  (ui/editor/src/main.ts:1120-1131). No act runs an editor action by id, so nothing exercises that
  run() body. The route row in docs/debug-api.md:91 describes what it does but does not say it
  bypasses the page.
- **Why:** The whole point of this command is that the editor's own undo would leave a half-renamed project, so
  it is the safety net for the product's most cross-cutting edit. The registration, the menu placement
  and the notice all live in the page and none of them are ever run by a test; only the shim function
  is. A broken menu entry passes the suite that claims to cover rename reversal.
- **Change:** Keep the route (it is a legitimate shim-level primitive) and designate it as a page bypass in its
  route row and the client method, then add the missing UI driver - either the toolbar-command act
  above generalised to run an editor action by id, or an explicit act that calls
  editor.getAction("xlide.undoRename").run() the way format calls getAction.
- **Size:** unmeasured; one act plus two doc lines
- **Adversary:** AddInSession.DebugApi.cs:2937-2945 is `case "undoRename": ... UndoRename(0); return
  DebugCommandReply(true, 0)` with the comment "Here so a probe can prove a rename is reversible
  without driving the page". The user path is the editor action at main.ts:1120-1131, whose run()
  awaits bridge.requestRenameUndo() and then notifies; that message reaches the shim at
  EditorSurface.cs:1772. No act runs an editor action by id (the act list has no such entry, and
  `undo` at devsurface.ts:1205-1208 triggers the built-in undo/redo by name, not getAction), and the
  two suites that touch it - rename-features.mjs:116 and three-copies.mjs:231 - both call
  api.undoRename(), the route. So the registration, the menu placement and the notice are never run by
  anything.
- **Correction applied:** The doc gap is in two places, not one: neither the route row at docs/debug-api.md:91 nor the client
  row at docs/driving-excel.md:210 says the route bypasses the page. Severity is low because the
  reversal logic itself is covered by two live suites; what is uncovered is the menu entry's
  registration and its notice, and the notice is unobservable anyway (see statusbar-no-observer).

### Shim capabilities with no driver or no observer

_5 findings, from the `api-shim-surface` finder._

##### `component-rename-strands-bookkeeping` component?action=rename skips AdoptRename, so a scripted rename strands breakpoints and the write baseline under the old name

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2861`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** The route body is `target.SetString("Name", newName); var readBack = target.GetString("Name") ??
  newName; Log.Info($"component: renamed {componentName} to {readBack}"); ComponentsChanged();`
  (AddInSession.DebugApi.cs:2861-2865). The UI path for the same operation is the Properties grid's
  "(Name)" row: EditorSurface.cs:1951 raises `PropertyEditRequested`, AddInSession.cs:3236-3239 does
  `found.SetString("Name", value); ... AdoptRename(component, actual);`. AdoptRename
  (AddInSession.cs:3344-3384) is what carries the record across: it re-keys `_writtenModules` (`if
  (_writtenModules.Remove(WrittenKey(oldName, owner), out var baseline)) {
  _writtenModules[WrittenKey(newName, owner)] = baseline; }`), re-keys `_breakpoints`
  (`_breakpoints[WrittenKey(newName, display)] = moving with { Module = newName };`), moves
  `_propertiesTarget`, reloads the surface if the shown module was the renamed one, and only then
  calls ComponentsChanged(). The sibling action in the same route already learned this lesson for
  removal: RemoveComponent (AddInSession.cs:6841-6845) does `foreach (var key in new[] {
  WrittenKey(removed, owner), WrittenKey(removed, null) }) { _writtenModules.Remove(key);
  _breakpoints.Remove(key); }`, and docs/debug-api.md:85 designates it: "**`remove` runs the product's
  own removal** ... the bare COM call it used to make left all three behind ... a harness removing a
  component left a different machine than a developer removing the same one". No such designation
  exists for `rename` in that row, in tools/harness/xlide-api.mjs:494, or in the code comment, which
  designates only the difference from the page's engine-backed `renameModule`.
- **Why:** This is the design rule (a) failure the `remove` action was already fixed for, still live in
  `rename`. After a scripted rename the session's breakpoint record still says the old module name, so
  GET breakpoints reports a breakpoint on a module that no longer exists and none on the module that
  now carries it; the write baseline is filed under the old key too, so ModuleDiffersFromSaved
  (AddInSession.cs:6564) answers false for the renamed module and the tab close gate silently stops
  offering Save/Don't Save. A debugger or close-gate suite that builds its fixture with this route is
  testing a machine no developer can produce, and the divergence reads as a product bug rather than a
  harness one.
- **Change:** Call AdoptRename(componentName, readBack) after the SetString instead of ComponentsChanged()
  (AdoptRename ends by calling ComponentsChanged itself, so nothing else changes). If the
  fixture-primitive behaviour is wanted deliberately, keep it behind an explicit argument and
  designate the deviation in all three places `remove` already designates its: the docs/debug-api.md
  row, the xlide-api.mjs method comment, and the code.
- **Size:** two dictionaries (_writtenModules, _breakpoints) left unmigrated per rename; one route action; the fix is roughly a one-line swap plus a docs row
- **Adversary:** Read the route myself: src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2852-2867 is exactly `using
  var target = FindComponent(componentName, componentOwner, out _); ... target.SetString("Name",
  newName); var readBack = target.GetString("Name") ?? newName; Log.Info(...); ComponentsChanged();`
  with nothing else. ComponentsChanged (AddInSession.cs:3401-3406) is only PublishModules +
  PublishProjects + _analysis?.Reanalyse(), so it cannot re-key anything. _breakpoints is a plain
  shim-side Dictionary keyed by name (AddInSession.cs:2214) and every read goes through
  BreakpointKey/WrittenKey (2234, 2907, 2941, 2976; the GET breakpoints route reads
  _breakpoints.Values at DebugApi.cs:2959); the only re-key sites in the whole src tree are
  AdoptRename (AddInSession.cs:3352-3363) and RemoveComponent's drop (6841-6845), which I read. The UI
  equivalent, the Properties grid "(Name)" row, does call AdoptRename (AddInSession.cs:3230-3240). The
  designation the finding says is missing really is missing: docs/debug-api.md:85 designates only
  `remove`, docs/debug-api.md:70 and the code comment at DebugApi.cs:2840-2842 designate only the
  difference from `renameModule` (reference rewriting), and tools/harness/xlide-api.mjs:472-494 says
  nothing about bookkeeping.
- **Correction applied:** Two corrections. (1) The proposed one-line swap is not safe as written: AdoptRename re-keys against
  `DisplayFromProjectId(_shownProject)` (AddInSession.cs:3349, 3360), while the route resolves its
  target from `componentOwner = ProjectIdFromDisplay(componentProject) ?? _shownProject`
  (DebugApi.cs:2718), so renaming a component in a non-shown workbook would migrate the keys under the
  wrong workbook - the owner has to be threaded into AdoptRename, the way RemoveComponent already
  computes `owner` at AddInSession.cs:6822. (2) Impact today is latent, not active: the only caller of
  this action in the repo is tools/harness/com-leak.mjs:301-303, which renames and immediately renames
  back, so the keys land back where they started and no current suite is misled. Also unverified by me
  and worth checking with the fix: RemoveComponent additionally calls
  `_editorSurface?.DiscardEdits(removed, owner)` (AddInSession.cs:6825); whether the surface's own
  per-module document/edit table re-keys on a rename I did not establish.

##### `properties-panel-has-no-driver-or-observer` The Properties panel - the shim's only property-write path into the object model - has no route to read it and no way to drive an edit

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3216`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity medium
- **Evidence:** `private void OnPropertyEdit(string component, string name, string value)` (AddInSession.cs:3216)
  writes into the host: `found.SetString("Name", value)` for the "(Name)" row (3234), otherwise
  `properties?.GetItem(name)` plus `WriteProperty(property, value, out var complaint)` (3251-3259),
  notifying and republishing on refusal. Its only caller is EditorSurface.cs:1951, reached from the
  page message `case "editProperty":` (EditorSurface.cs:1943). The page raises it only from DOM
  handlers on the panel's own controls: ui/editor/src/shell.ts:492
  `this.handlers.editProperty(this.propertiesComponent, property.name, select.value)` and shell.ts:539
  for the text input, through bridge.ts:647. There is no `editProperty` entry in the act table
  (ui/editor/src/devsurface.ts:514-1481; the full action list is closeActive, answerCloseConfirm,
  activate, cycleTab, split, expandWorkbook, unfoldModule, treeMenu, renameModule, treeAdd, menuBar,
  chooseMenuItem, answerRemoveConfirm, settings, sponsors, closeDialogs, key, focusEditor, search,
  hover, completions, signature, quickFixes, timeFeature, references, definition, rename, undo, dock,
  press, insert, backspace, format, bookmark). The read side is equally absent: PublishProperties
  (AddInSession.cs:3074-3130) computes the rows, the component kind and the allowed-name list and
  sends them as SurfacePropertyEntry[] (EditorMessages.cs:323), but no case in the route switch
  answers about properties (grep of `case "` over AddInSession.DebugApi.cs), and the `ui` route's page
  snapshot carries only `panes: parts.panes.list().map(({ name, title, open }))` (devsurface.ts:461),
  never the grid's rows.
- **Why:** A test cannot set a property, cannot read what the grid is showing, and cannot check that a refused
  write produced the complaint WriteProperty returns. The path matters more than a settings grid
  usually would: the "(Name)" row renames a component and runs AdoptRename with all its re-keying
  (AddInSession.cs:3239), and ClassifyDocument gates which rows a document component may even edit
  (3117). The only way to exercise any of it today is to synthesise input events on shell.ts's
  controls through `eval`, which is the probe-testing-the-probe pattern the act route was built to
  end.
- **Change:** Add a `properties` route: GET (optional `component`) answers the rows PublishProperties computed -
  component, kind, and each entry's name, value and whether it is editable - so the observation reads
  the same fields the panel renders; POST with `component`, `name`, `value` calls OnPropertyEdit, the
  exact entry point the page's message reaches, and returns the complaint when the write is refused.
  Document it in both api docs and give xlide-api.mjs a properties() / setProperty() pair.
- **Size:** one route, two verbs; unmeasured beyond that
- **Adversary:** I checked all four reachability paths the rules demand. (1) Route switch: I listed every `case "` in
  src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs (518-3327) - log, messages, capture, immediate,
  locals, watches, problems, drainfinalizers, history, assert, journal, perf, ui, act, eval, await,
  console, inspect, bench, trip, layout, reload, dialogs, compile, type, mark, guard, dismiss, stats,
  state, doctor, engine, native, windows, menus, outline, sync, component, pane, undoRename,
  breakpoints, settings, projects, project, documents, command, breakpoint, module, caret, placement -
  no properties case. (2) act table: read the action list in ui/editor/src/devsurface.ts:514-1481, no
  editProperty or property action. (3) ui snapshot: grep for 'propert' over devsurface.ts returns
  nothing at all. (4) Every editProperty reference in the repo: EditorSurface.cs:1943 (the page
  message), bridge.ts:278/647/2143, main.ts:522, shell.ts:67/492/539 - and bridge.ts:2143 is the
  browser DEMO host echoing a notice, not a driver. Harness grep for 'propert' finds only unrelated
  hits plus Test-SplitWorkspace.ps1:338 and Test-DebugApi.ps1:558, which treat 'properties' as a dock
  PANEL NAME, never its rows. `inspect` (DebugApi.cs:1218-1290) is a CSS/DOM/computed-style reader,
  not a read of the panel's model.
- **Correction applied:** One qualifier on the stated impact: the gap is a coverage gap, not a violation of api rule (a) or
  (b) - there is no route making a wrong claim, there is simply no route. The practical cost is
  concentrated in the "(Name)" row, because that is the UI entry point for AdoptRename
  (AddInSession.cs:3230-3240), which finding component-rename-strands-bookkeeping shows is the
  behaviour the rename route is missing; today neither the route nor the UI path has a test.

##### `locals-route-cannot-say-which-layer-it-stopped-at` GET locals collapses four different failures into an empty list: no ghost palette, an unreadable one, a dead reader thread, and a break with no variables

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:797`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** The route is `var snapshot = _ghostReaders?.Locals; var rows = snapshot is null ? [] : new
  SurfaceLocalRow[snapshot.Rows.Count]; ... return ... new DebugLocalsReply(snapshot?.Context, rows)`
  (AddInSession.DebugApi.cs:797-814). `_ghostReaders` is null whenever GhostReaderThread.Start found
  no windows: `if (localsWindow == 0 && watchWindow == 0) { return null; }`
  (GhostReaderThread.cs:54-57). The thread can start and still never read: `locals = _localsWindow !=
  0 ? LocalsReader.Create(_localsWindow) : null;` with `Log.Info(locals is null ? "locals: the ghost
  palette could not be read; the panel sits idle" : ...)` (GhostReaderThread.cs:95-101). It can also
  die mid-session: `catch (Exception ex) { Log.Error("ghost reader: the reading thread died", ex); }`
  (GhostReaderThread.cs:130-135), after which Locals keeps answering the last Volatile.Read or null
  after ClearReadings. All four reach the caller as {context: null, locals: []}. The assert vocabulary
  repeats it: `case "localsHas": var rows = _ghostReaders?.Locals?.Rows; ... names.Length == 0 ? "(no
  locals)" : ...` (AddInSession.DebugApi.cs:2118-2126). The watches route carries one bit more - `new
  DebugWatchesReply(_inBreak, rows)` (816-831) - and locals carries none.
- **Why:** This is the observation rule (b) failure in its exact form: the answer does not say which layer it
  stopped at. The 2026-08-05 episode this thread exists for (24 sessions of unreadable elements) would
  present today as a green-looking empty locals list; a debugger suite asserting localsHas would fail
  with "(no locals)" and point the reader at the debugger rather than at the accessibility reader, and
  the only evidence separating the cases is a log line.
- **Change:** Widen DebugLocalsReply with what the session already knows: whether _ghostReaders exists, whether
  the LocalsReader attached (GhostReaderThread can hold the Create result as a bool), whether the
  reading thread is still alive, the ghost window handle, and _inBreak, the way watches already
  reports it. Mirror the same fields on watches. No new route, no shape change to the existing fields.
- **Size:** one reply record plus two bools tracked on GhostReaderThread
- **Adversary:** The code reads as quoted. src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:797-813 returns `new
  DebugLocalsReply(snapshot?.Context, rows)` with rows empty whenever `_ghostReaders?.Locals` is null,
  and watches at 816-830 does carry `_inBreak` where locals carries nothing. The file is
  src/Xlide.Vbe.Shim/Editor/GhostReaderThread.cs (not Debugging/), and every cited line is right
  there: Start returns null when both handles are 0 (54-57), Create can return null with the log line
  "locals: the ghost palette could not be read; the panel sits idle" (95-101), the run loop's catch
  logs "ghost reader: the reading thread died" (130-135), and Locals is just `Volatile.Read(ref
  _locals)` (46) which ClearReadings nulls (77-81). The assert claim localsHas collapses the same way
  (DebugApi.cs:2118-2125).
- **Correction applied:** Three cases, not four, and the headline overstates the blindness. The no-palette case IS already
  separable one call away: `doctor` reports `GhostReadersUp: _ghostReaders is not null` and pushes the
  finding "the ghost readers are not attached, so Locals and Watch cannot fill"
  (DebugApi.cs:2419-2422, 2455). `_inBreak` is likewise reachable via the watches route and the
  `stopped` assert claim (DebugApi.cs:829, 2095-2099). And `snapshot?.Context` is already a partial
  layer indicator: Context is only populated when a read actually produced a context row
  (LocalsReader.cs:35, 205), so a non-null context proves the reader attached and read. What genuinely
  cannot be separated is LocalsReader.Create having returned null versus the reading thread having
  died versus a break whose read produced no context - all three answer {context: null, locals: []}.
  Rule (b)'s primary clause is also satisfied rather than violated: an empty list is exactly what the
  panel renders in every one of these states, so this is a diagnosability improvement, not a false
  observation.

##### `native-route-thin-for-background-panes` The native parity route gives background panes only a process-local hash: no caret, no text under text=1, and no key that compares outside the one reply

- **Where:** `src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:884`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity low
- **Evidence:** `public sealed record DebugNativePaneRow([property: JsonPropertyName("module")] string Module,
  [property: JsonPropertyName("project")] string? Project, [property: JsonPropertyName("hostContent")]
  string? HostContent, [property: JsonPropertyName("surfaceContent")] string? SurfaceContent);`
  (DebugServer.cs:884-893). The route fills those rows from ReadOpenModules() and nothing else
  (AddInSession.DebugApi.cs:2570-2589), while caret and text are computed only for the active pane:
  `activePane.InvokeInt32s("GetSelection", selection); caretLine = selection[0]; caretColumn =
  selection[1];` (2523-2527) and `wantText ? nativeText : null, wantText ? surfaceText : null`
  (2604-2605), where nativeText/surfaceText are the active pane's. The reduction both content fields
  use is ContentKey: `var normalised = text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
  return $"{normalised.Length}:{normalised.GetHashCode(StringComparison.Ordinal)}";`
  (AddInSession.cs:5731-5740). The session already has the per-pane machinery: FindCodePane(component,
  projectId) at AddInSession.cs:7014.
- **Why:** The route's own comment says a background tab going stale is the case nothing notices, but when a
  background row disagrees the reply carries nothing to act on: no text to diff, and a key built from
  String.GetHashCode, which is randomised per process, so it cannot be compared against anything
  outside that single reply - not a stored expectation, not the page's own copy read through eval.
  Recovering the two texts means two more `module` calls with a different reduction, and the third
  layer, what the page's monaco model actually holds, never appears in a native comparison at all.
  Per-pane caret is missing too, which is what decides where Run and Step land after a tab switch.
- **Change:** Add caretLine/caretColumn to DebugNativePaneRow, read through the existing FindCodePane; accept
  module= alongside text=1 so the pair of texts can be pulled for a named background pane; and reduce
  content with a stable hash (FNV or SHA over the normalised text) so a key means the same thing in
  two replies and can be computed by a harness. Optionally carry the page's model text for each row as
  a third column so native, surface and page compare in one reply.
- **Size:** four new fields on one row record plus one query argument; ContentKey is 6 lines
- **Adversary:** The route body is as described: src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2569-2590 builds
  each pane row from ReadOpenModules() as `new DebugNativePaneRow(pane.Name, pane.Project,
  ContentKey(hostText), ContentKey(_editorSurface?.TextOf(pane.Name, pane.Project)))` and nothing
  more, while caret comes only from ActiveCodePane's GetSelection (2528-2533) and `text=1` carries
  only the active pane's two texts (2605-2606). The record has exactly the four fields quoted
  (DebugServer.cs:884-893), and ContentKey is
  `$"{normalised.Length}:{normalised.GetHashCode(StringComparison.Ordinal)}"`
  (AddInSession.cs:5731-5740), whose hash is per-process randomised.
- **Correction applied:** The description is right but three of the supporting claims are wrong, and they carry the finding's
  weight. (1) "cannot be compared against anything outside the one reply" is false: String.GetHashCode
  is stable for the LIFETIME OF THE PROCESS, so successive `native` replies from the same session
  compare fine - which is the stated use, "asked after every step of a randomised walk"
  (DebugApi.cs:2550-2551) - and the key is prefixed with the normalised LENGTH, which is portable
  across processes. The real limit is only a stored cross-session expectation or a harness-computed
  key. (2) "Recovering the two texts" is already a documented two-call operation for ANY named pane,
  background or not: `module?name=X` returns the host's source (DebugApi.cs:3285-3292) and
  `module?name=X&live=1` returns the surface's copy (3273-3282), the same `_editorSurface.TextOf` the
  pane row hashes - so nothing is unreachable, it is merely not in one reply. (3) The per-pane caret
  rationale is wrong: Run, Step and ToggleBreakpoint act on ActiveCodePane, whose caret the reply
  already carries and which tools/harness/debugger-features.mjs:109-110 and 130-132 already assert on.
  What survives is a convenience/portability improvement, not a coverage hole.

##### `no-driver-for-shutdown-or-cancelled-shutdown-revival` Session teardown and the cancelled-shutdown revival - the path that once left the add-in dead inside a living Excel - cannot be driven or observed at all

- **Where:** `src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs:116`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity medium
- **Evidence:** OnBeginShutdown calls `_session?.Stop();` then arms `_watchdog ??=
  ShutdownWatchdog.Create(OnWatchdogTick);` (XlideAddIn.cs:124-132). OnWatchdogTick revives from
  retained pointers after two consecutive enabled-and-visible frame ticks: `_session = new
  AddInSession(editor, addIn); _session.Start(); _session.HostStartupComplete();`
  (XlideAddIn.cs:205-223), guarding the failure recorded at 252-259 ("observed 2026-08-02 ... the
  developer cancelled, and the editor came back with the add-in dead and nothing listening"). Stop()
  is what must leave the host whole: `_debugServer?.Dispose(); _debugServer = null;` first
  (AddInSession.cs:7605-7606), then `SetNativeChromeBands(visible: true);` (7632),
  RestoreLocalsPalette/RestoreWatchPalette (7644-7645), palette and chrome disposal (7654-7661).
  Nothing in the route switch reaches any of it (no session, stop, shutdown or revive case in the grep
  of `case "` over AddInSession.DebugApi.cs), and the harness never exercises it: grep of
  tools/harness for revive|revival|watchdog|OnBeginShutdown|shutdown returns one unrelated hit, `await
  call("shutdown", {})` against the engine's own JSON-RPC in tools/harness/engine-live-probe.mjs:167.
  The server is per-session and rewrites its discovery file on start (DebugServer.Start at
  AddInSession.cs:1026; `var discoveryPath = Path.Combine(directory,
  $"debug-api-{Environment.ProcessId}.json")` at DebugServer.cs:73), so a revived session would be
  rediscoverable.
- **Why:** Two release-blocking properties live here and neither has a test. Stop() is the only place the
  native menu and toolbar bands, the ghosted Locals and Watches palettes and the host title bar are
  put back, so a regression that leaves the editor menu-less after unload is invisible until a user
  hits it. The revival is worse: it is a documented field failure with a subtle two-tick guard against
  reviving during the save prompt, and the only way to reach it today is to close Excel by hand and
  press Cancel. The COM wrapper counters that make the leak gate meaningful (ComWrappersLive in the
  stats route, AddInSession.DebugApi.cs:1833) are never checked across a teardown, which is where a
  leak would actually show.
- **Change:** Add a Debug-only lifecycle route, for example session?action=beginShutdown and
  session?action=disconnect, that calls the same XlideAddIn entry points the host calls (reachable
  through a static reference set in OnConnection) rather than a private teardown of its own, so the
  state left behind is the state the host's own call leaves. The server dies with the session by
  design; the client reconnects by re-reading debug-api-{pid}.json, which the revived session
  rewrites. That makes three assertions possible for the first time: the native chrome came back, a
  cancelled shutdown revives, and ComWrappersLive returns to its pre-stop level.
- **Size:** one route with two actions plus a static hook in XlideAddIn; unmeasured
- **Adversary:** Every cited line reads as claimed. src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs:116-140 is OnBeginShutdown
  calling `_session?.Stop()` then arming the watchdog; OnWatchdogTick:151-230 revives with `_session =
  new AddInSession(editor, addIn); _session.Start(); _session.HostStartupComplete();` behind the
  two-consecutive-enabled-and-visible-ticks guard (194-203), and OnDisconnection:250-264 carries the
  2026-08-02 field-failure note. I enumerated the whole route switch myself (DebugApi.cs:518-3327) -
  there is no session, stop, shutdown, disconnect or revive case. My own grep over tools and docs for
  watchdog|revive|revival|beginshutdown|ondisconnection returns only prose in docs/architecture.md:98,
  docs/lessons.md:349-360 and docs/decisions.md:259, the last of which is a DIFFERENT watchdog (a
  modal-dialog dismissal guard), and no harness reference at all. The nearest existing coverage,
  tools/harness/Test-CloseVbe.ps1, closes the VBE FRAME with WM_SYSCOMMAND SC_CLOSE and reopens it
  through the object model; it never reaches OnBeginShutdown or Stop().
- **Correction applied:** One mechanical problem with the proposed route that has to be designed around, not a defect in the
  finding: Stop() disposes the DebugServer FIRST (AddInSession.cs:7605-7606, which is also the
  listener serving the in-flight request), so a route that calls it synchronously cannot write its own
  reply - the caller sees a dropped connection. The teardown has to be posted to the host thread after
  the response is flushed, and the route documented as answer-less, otherwise every suite using it
  fails on a transport error rather than on an assertion. Also note the finding's own suggestion
  satisfies rule (a) only if the route calls XlideAddIn's entry points rather than AddInSession.Stop
  directly, since it is OnBeginShutdown that arms the watchdog (XlideAddIn.cs:130-132) and the revival
  is the thing under test.

### The api as a testing instrument

_14 findings, from the `api-quality` finder._

##### `command-route-always-ran-true` `command` answers `ran: true` for a command the editor refused, because ExecuteEditorCommand discards VbeCommands.Execute's result

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3197`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** Route: `ExecuteEditorCommand(command); return ... new DebugCommandReply(true, command)`
  (DebugApi.cs:3197-3199). The reply field is literally named `ran` (DebugServer.cs:709-711:
  `[property: JsonPropertyName("ran")] bool Ran`). `ExecuteEditorCommand` is `private void`
  (AddInSession.cs:1722) and does `var ran = VbeCommands.Execute(_editor, command);` (1756) then uses
  `ran` only to Notify for three specific ids (1762-1779); the value never leaves the method.
  `VbeCommands.Execute` returns false in several ways: no CommandBars (VbeCommands.cs:100), `if
  (control is null) { Log.Info($"command: {commandId} is not present in this host"); return false; }`
  (109-114), `if (!control.GetBool("Enabled")) { ... return false; }` (118-122), and on exception
  (128-132). So `POST command?name=stepOver` outside break mode logs "currently disabled" and answers
  `{"ran":true}`. The one .mjs consumer can only check the transport:
  tools/harness/immediate-watch.mjs:165-168 `const added = await api.command("addWatch")...;
  check("the add-watch command reaches the editor", added.error === undefined, added.error ??
  `ran=${added.ran}`)` with the comment "The command is asserted to REACH the editor".
- **Why:** `command` is the harness's only way to drive Run, Step, Reset, Compile, Save and the watch commands,
  and it cannot distinguish "executed" from "the menu item was greyed". Every debugger suite step
  (debugger-features.mjs:99,127,154; step-into-features.mjs:81,92) is a call whose reply is a
  constant, so a regression that disables a command reads as a passing call followed by a state
  assertion that fails somewhere else entirely.
- **Change:** Make ExecuteEditorCommand return the bool it already computes (or a small record carrying the
  decline reason it already builds for Notify) and put it in `ran`. No route or field is renamed;
  `ran` starts meaning what its name says. Add a `detail` field carrying "not present" / "disabled" so
  a caller can tell the two apart.
- **Size:** one bool plumbed out of a void method; 9 existing call sites of ExecuteEditorCommand, only 3 in the debug api
- **Adversary:** Read src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3197-3199: `ExecuteEditorCommand(command);
  return ... new DebugCommandReply(true, command)` - the constant is literal.
  src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:709-711 names the field `ran`.
  src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1722 is `private void ExecuteEditorCommand(int command)`
  and at 1755 does `var ran = VbeCommands.Execute(_editor, command);`, using `ran` only for the three
  Notify blocks at 1760-1780. src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:90-132 returns false on no
  CommandBars (97-101), control not found (109-114 `Log.Info($"command: {commandId} is not present in
  this host")`), `if (!control.GetBool("Enabled")) { Log.Info($"command: {commandId} is currently
  disabled"); return false; }` (118-122) and on exception (127-131). docs/debug-api.md:96 documents
  only `name` and `keep`, so nothing designates `ran` as meaning 'the request was posted'.
  tools/harness/immediate-watch.mjs:164-168 checks only `added.error === undefined`.
- **Correction applied:** The size figure is wrong: ExecuteEditorCommand has 7 call sites, not 9 (AddInSession.cs:1664, 4156,
  4288, 5363 and DebugApi.cs:712, 1643, 3197); the '3 in the debug api' part is right. Also soften the
  blast radius: debugger-features.mjs never reads `.ran` from these calls (it waits for break mode at
  line 101 and re-reads api.native() at 130), so a disabled command costs a delayed and misattributed
  diagnosis rather than a silent pass.

##### `compile-reports-clean-when-it-never-ran` `compile` answers `compiled: true` when the compile command was refused or never started, and always burns the whole waitMs

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1670`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** `var command = VbeCommands.ForName("compile"); compileSurface.RunOnHostThread(() =>
  ExecuteEditorCommand(command));` (DebugApi.cs:1642-1643) - the result is discarded exactly as in the
  `command` route. The verdict is derived only from dialogs seen: `new DebugCompileReply(said.Count ==
  0, [.. said], ...)` (1670-1671). The watch loop has no early exit: `var settle =
  Environment.TickCount64 + WaitMilliseconds(request, 6000); while (Environment.TickCount64 < settle)
  { Thread.Sleep(150); foreach (var raised in DialogWatch.Dialogs()) ... }` (1646-1668). So three
  different states collapse onto `compiled: true`: the project compiled, the Compile item was disabled
  and nothing ran, and the error box arrived at second 7. Consumer:
  tools/harness/debugger-features.mjs:74-77 `// It has to compile, or every command below is really a
  test of the dialog guard.` / `const compiled = await api.compile(); check("the fixture compiles",
  compiled.compiled !== false, ...)`.
- **Why:** This is the precondition gate for the whole debugger suite, and the comment above it states what the
  gate is protecting against. A gate that passes when the compile never happened turns the suite it
  guards into the thing the comment says it must not become. It also costs a fixed 6s per call because
  a clean compile is indistinguishable from a compile still running.
- **Change:** Take the `ran` bool out of ExecuteEditorCommand (see the previous finding) and add `started` to
  DebugCompileReply, so `compiled: true, started: false` is expressible. For the timing half, add an
  `elapsedMs` field and stop early once the project's Mode has settled, or at minimum document that
  `compiled` means "no dialog appeared within waitMs" rather than "the project compiles".
- **Size:** 6000ms fixed per call; one new bool field
- **Adversary:** DebugApi.cs:1642-1643 `var command = VbeCommands.ForName("compile");
  compileSurface.RunOnHostThread(() => ExecuteEditorCommand(command));` discards the result exactly as
  the `command` route does. DebugApi.cs:1646-1668 is a `while (Environment.TickCount64 < settle)` loop
  with `Thread.Sleep(150)` and no break on success, and 1670-1671 derives the verdict solely from
  dialogs seen: `new DebugCompileReply(said.Count == 0, [.. said], ...)`. WaitMilliseconds default is
  6000 at 1647. tools/harness/debugger-features.mjs:74-77 is the precondition gate with the comment
  'It has to compile, or every command below is really a test of the dialog guard.'
- **Correction applied:** The route table is not silent about the semantics: docs/debug-api.md:82 already says '`compiled` is
  false when anything appeared', which designates the dialog-derived meaning in one of the three
  required places. What is undesignated is the 'never started' case (the client row
  docs/driving-excel.md:204 says only 'compiles; errors as DATA, modal cleared', and the code comment
  at 1630-1636 says nothing about it), and the fixed cost is the default of a caller-overridable
  `waitMs`, not a hard 6s.

##### `parity-compares-host-against-shim-reconstruction` The parity instrument (`native`, inSync, parityAll) compares the live host pane against the shim's RECONSTRUCTION of the page, and never asks the page for its text

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2588`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity medium
- **Evidence:** `native` builds both sides of every comparison from the shim:
  `ContentKey(_editorSurface?.TextOf(pane.Name, pane.Project))` (DebugApi.cs:2588) and `var
  surfaceText = _editorSurface?.Text;` (2560), keyed at 2604. Those read the shim's own dictionary:
  `public string? Text => ActiveDoc?.Text;` and `public string? TextOf(...) =>
  _docs.TryGetValue(DocKey(...), out var doc) ? doc.Text : null;` (EditorSurface.cs:949-953). That
  field is a reconstruction, not a copy of the page's model: the page sends incremental `changes` and
  the shim applies them to its own string - `parsedEdits = ParseChanges(editedDoc.Text, changeSet);
  updated = parsedEdits is null ? null : ApplyEdits(editedDoc.Text, parsedEdits);` then
  `editedDoc.Text = updated;` (EditorSurface.cs:1388-1406). A drift detector exists and goes nowhere:
  `if (... lengthElement.TryGetInt32(out var expectedLength) && updated.Length != expectedLength) {
  Log.Error($"surface: reconstructed {updated.Length} character(s) where the page holds
  {expectedLength}"); }` (1394-1399) - grep for "reconstructed" finds it in no route, no reply, and no
  counter. The client presents this as page parity: `inSync()` is documented "True when the native
  pane, the surface and the page all name the same module" and its page contribution is a NAME only,
  `const page = ui.focus.model ? ui.focus.model.split("/").pop() : null` (xlide-api.mjs:387), while
  the content comparison is `below.nativeContent === below.surfaceContent` (398-399).
  debugger-features.mjs:188 asserts on it as "the native editor and the page still agree". `documents`
  (DebugApi.cs:3179) and `module?live=1` (3275) read the same reconstruction and say nothing about
  which layer they stopped at.
- **Why:** Editor parity is a definition-of-done property here and this is the instrument that certifies it,
  but the page is never asked what its model holds. A reconstruction bug - a mis-parsed change range,
  a dropped message - shows up as `agreed: true` on every suite that calls parity(), which is most of
  them. The one check that would catch it already runs on every edit and only reaches the log.
- **Change:** Add fields, do not change existing ones: give `native` a `pageContent`/`pageContentKey` read through
  RunPageScript from the model the render uses, so the three-way comparison is three-way. Separately
  expose the reconstruction-mismatch count (the Log.Error at EditorSurface.cs:1398) as a counter in
  `stats`, so a suite can assert it is zero without reading the log.
- **Size:** one page round trip added to `native`; one counter
- **Adversary:** Verified the substance. No route asks the page for its model text: DebugApi.cs:2560 `var surfaceText
  = _editorSurface?.Text;` and 2588 `ContentKey(_editorSurface?.TextOf(pane.Name, pane.Project))`,
  both resolving to EditorSurface.cs:949 `public string? Text => ActiveDoc?.Text;` and 951-953 TextOf.
  That string is rebuilt from incremental edits at EditorSurface.cs:1387-1391 (`parsedEdits =
  ParseChanges(...); updated = ... ApplyEdits(...)`) and stored at 1406 `editedDoc.Text = updated;`.
  The page does send `fullLength` (ui/editor/src/bridge.ts:1626,1633), so the drift check at
  EditorSurface.cs:1394-1399 does fire, and its only outlet is `Log.Error($"surface: reconstructed
  {updated.Length} character(s) where the page holds {expectedLength}")` - grep for 'reconstructed'
  across the repo finds it in no route, reply, counter or doc. ui/editor/src/devsurface.ts:38-78
  confirms the UiSnapshot carries no text or length, only `focus.model` as a URI.
- **Correction applied:** Two evidence errors must not stand. (1) 'native builds both sides of every comparison from the shim'
  is false: the host side is read live off the object model - DebugApi.cs:2554 `nativeText =
  ProjectReader.ReadSource(component)` and 2573-2575 `FindComponent(...)` then
  `ProjectReader.ReadSource(found)`. The comparison is host vs shim, as the title says. (2) The layer
  stops ARE designated, so api rule (b) is satisfied for the routes as documented:
  docs/debug-api.md:50 says the native reply carries 'what the surface believes it is showing', :83
  says `documents` is 'the documents the surface holds TEXT for', :86 says `module?live=1` is 'the
  SURFACE's copy', and the client names its fields `surfaceContent`/`nativeContent` and says of the
  page only that it 'names the same module' (xlide-api.mjs:384-403). What genuinely survives is
  narrower: the shim-to-page link is the one hop nothing asserts, and the detector that already
  watches it (EditorSurface.cs:1398) reaches only the log, so no suite can assert it is zero. The
  `pageContent` field is a nice-to-have; the counter in `stats` is the load-bearing half.

##### `breakpoint-and-caret-report-goto-silence` `breakpoint` reports the state it was asked for rather than the one that happened, and a failed navigation makes it toggle a breakpoint on the previously shown module

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3237`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** Route: `GoTo(module, breakLine, 1, project); var alreadySet = _editorSurface?.Module is { }
  shownModule && BreakpointsFor(shownModule).Contains(breakLine); ... if (shouldSet != alreadySet) {
  ToggleBreakpoint(breakLine); } return ... new DebugCommandReply(shouldSet,
  VbeCommands.Command.ToggleBreakpoint)` (DebugApi.cs:3217-3238). Three silences feed it. (a) GoTo
  swallows a missing pane: `using var pane = FindCodePane(component, projectId); if (pane is null) {
  Log.Info($"navigate: no pane for {component}"); return; }` (AddInSession.cs:1359-1364) and wraps
  everything in `catch (Exception ex) { Log.Error(...); }` (1404-1407). (b) After that silent return
  `_editorSurface.Module` is still the PREVIOUS module, and ToggleBreakpoint acts on exactly that:
  `var module = _editorSurface?.Module; if (module is null || line < 1) { return; }`
  (AddInSession.cs:2931-2935). (c) ToggleBreakpoint refuses silently on a non-executable line: `if
  (!clearing && !CanBreakOn(_editorSurface?.LineAt(line))) { Log.Info($"breakpoint: {module}({line})
  is not an executable statement"); return; }` (2943-2951) and again when its own pane lookup fails
  (2963-2967). Meanwhile the `ran` field carries `shouldSet`, so `state=off` on a clean line answers
  `ran:false` on success. The same GoTo silence sits under `caret`, which answers `new
  DebugCommandReply(true, 0)` unconditionally (DebugApi.cs:3310-3315). Callers already pay for it:
  debugger-features.mjs:84-91 does `await wait(1200); const recorded = (await
  api.breakpoints()).breakpoints ?? []; check("the breakpoint was recorded against Runner", ...)`.
- **Why:** Setting a breakpoint on a module whose pane is not open silently sets one somewhere else and answers
  ok - the exact class of defect the route's own comment says `state=on|off` was added to prevent. And
  because `ran` carries the requested state, a caller cannot use it as a success flag in either
  direction, so every breakpoint assertion costs a sleep plus a second `breakpoints` round trip.
- **Change:** Make GoTo return a complaint string the way WriteModule does, and have `caret` and `breakpoint`
  answer it as `error`. Add `set` (the observed state, re-read from BreakpointsFor after the toggle)
  and `module` (the module actually acted on) to the breakpoint reply so `ran` stops carrying two
  meanings and one call answers what two calls answer now.
- **Size:** one return type on GoTo, 2 routes, 2 new fields; removes one 1200ms sleep plus a round trip per breakpoint assertion
- **Adversary:** DebugApi.cs:3217-3238 read as quoted: `GoTo(module, breakLine, 1, project);` then `alreadySet` from
  `_editorSurface?.Module`, then `ToggleBreakpoint(breakLine)`, then `new DebugCommandReply(shouldSet,
  VbeCommands.Command.ToggleBreakpoint)` - so `ran` carries the requested state and `state=off`
  answers `ran:false` on success. GoTo is `private void` (AddInSession.cs:1354) with the silent early
  return at 1359-1364 and a swallowing `catch (Exception ex) { Log.Error(...); }` at 1404-1407.
  ToggleBreakpoint (AddInSession.cs:2929-2935) acts on `_editorSurface?.Module`, which after a failed
  GoTo is still the previously shown module, and refuses silently at 2943-2951 and again at 2963-2967.
  The caret route (DebugApi.cs:3298-3315) calls the same GoTo and returns `new DebugCommandReply(true,
  0)` unconditionally.
- **Correction applied:** The mechanism sentence in why_it_matters is wrong. FindCodePane creates a pane that was never opened
  - AddInSession.cs:7023-7025, 'Reading CodePane on a module that has never been opened creates the
  pane' - so 'a module whose pane is not open' is NOT the failing case. GoTo returns null-pane only
  when the component does not resolve: a misspelt module name, or a name that exists only in another
  workbook when `project` is wrong. That is the case in which the breakpoint silently lands on the
  previously shown module.

##### `pane-close-always-ok` `pane action=close` answers ok even when the close was declined, which is the exact lie its sibling `open` was fixed for

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2926`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** `OnModuleCloseRequested(paneModule, DisplayFromProjectId(paneOwner), closeAnswer); return ... new
  DebugCommandReply(true, 0);` (DebugApi.cs:2925-2928). The callee is `private void` and has three
  paths that do not close: a failed save - `if (display is null || !SaveWorkbookOf(display)) {
  Log.Warn($"close: {display ?? \"the workbook\"} would not save; {component}'s tab stays"); return;
  }` (AddInSession.cs:6485-6491); an unanswered confirm - `if (display is not null &&
  ModuleDiffersFromSaved(component, display)) { ... _editorSurface?.ConfirmClose(component,
  projectDisplay); return; }` (6545-6550); and a discard whose revert write is refused, which only
  calls `_editorSurface?.Notify(...)` (6512). Twenty lines above, the `open` branch carries the
  comment "And its answer is the show's answer: opening a module that is not there replied ok, which
  is the same lie the write route told about a module that is not there (2026-08-09)"
  (DebugApi.cs:2910-2913) and returns `showed`'s error. docs/debug-api.md:89 records the open fix and
  says nothing about close.
- **Why:** Close is the more interesting half - it is the path with the save/discard/cancel gate the tab's own
  X uses - and a caller cannot tell "the tab closed" from "a Save box is now standing and the tab is
  still there". A probe testing tab-strip behaviour after a close is testing a state it has not
  established.
- **Change:** Give OnModuleCloseRequested a `string?` complaint return like WriteModule and RemoveComponent
  already have, and answer it. Where it raised ConfirmClose rather than closing, answer that as a
  distinct non-error outcome (`closed: false, awaiting: "confirm"`) so a caller can drive
  `act("answerCloseConfirm")` next instead of guessing.
- **Size:** one return type, three existing early-return sites
- **Adversary:** DebugApi.cs:2919-2928: `OnModuleCloseRequested(paneModule, DisplayFromProjectId(paneOwner),
  closeAnswer); return ... new DebugCommandReply(true, 0);`, twenty lines under the open branch whose
  comment at 2909-2912 says opening a module that is not there 'replied ok, which is the same lie the
  write route told' and which now returns `showed`'s error (2914-2918). OnModuleCloseRequested is
  `private void` at AddInSession.cs:6472; the failed-save return is at 6484-6491 (`Log.Warn($"close:
  {display ?? \"the workbook\"} would not save; {component}'s tab stays"); return;`) and the
  unanswered-confirm return at 6544-6550 (`_editorSurface?.ConfirmClose(component, projectDisplay);
  return;`). docs/debug-api.md:89 records the open fix and says nothing about the close reply.
- **Correction applied:** Two non-closing paths, not three. The discard branch whose revert write is refused
  (AddInSession.cs:6506-6512) still falls through to `CloseModule(component, projectDisplay)` at 6540
  - the tab does close, the module simply keeps the abandoned text and the page gets a Notify. Drop it
  from the list; it is a different defect (a silent failed revert), not a false close.

##### `undorename-always-ok` `undoRename` answers `ran: true` unconditionally; its failures go to the page and the log, so the only .mjs caller prints the reply and then sleeps 3 seconds

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2943`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** `UndoRename(0); return ... new DebugCommandReply(true, 0);` (DebugApi.cs:2943-2945). `UndoRename` is
  `private void` (AddInSession.cs:5030) and every failure is reported elsewhere: `if (_undoableRename
  is not { } undo) { surface.ShowRenamed(requestId, null, null, [], 0, "There is no rename to undo.");
  return; }` (5038-5042) - that message goes to the page for request id 0, which the route invented; a
  refused write stops the undo part-way with `stopped = $"'{target}' could not be written, so the undo
  stopped there. {refused}"` (5058-5062) and `catch (Exception ex) { stopped = ...; Log.Error(...);
  break; }` (5067-5072), leaving some modules restored and some not. Consumer:
  tools/harness/rename-features.mjs:116-118 `const undone = await api.undoRename(); console.log(`
  ${JSON.stringify(undone).slice(0, 200)}`); await wait(3000);` - the reply is printed, never checked,
  and the suite then re-reads two modules to find out what happened. The forward half of the same
  feature answers properly: `const said = await api.act("rename", ...); check("the rename was
  accepted", said.did, said.detail)` (82-84).
- **Why:** Rename reversibility is the property this route exists to let a probe assert, and the reply cannot
  express "there was nothing to undo" or "it stopped half way" - the worst outcome, a partially
  reverted multi-module rename, is indistinguishable from a clean one. The two halves of one feature
  answer in two shapes, so a suite can assert on the rename and only on the side effects of the undo.
- **Change:** Return the same shape the page gets: modules restored, the component's restored name, and the
  `stopped` reason when there is one. `act("rename")`'s `{did, detail}` is the precedent to match.
- **Size:** one reply record; removes a 3000ms sleep and two read-backs from rename-features.mjs
- **Adversary:** DebugApi.cs:2938-2945: `UndoRename(0); return ... new DebugCommandReply(true, 0);`. UndoRename is
  `private void UndoRename(int requestId)` at AddInSession.cs:5030; the nothing-to-undo path is
  `surface.ShowRenamed(requestId, null, null, [], 0, "There is no rename to undo."); return;` at
  5038-5042, and the partial-revert paths set `stopped = $"'{target}' could not be written, so the
  undo stopped there. {refused}"` and `break` at 5058-5062 and 5066-5072, leaving `restored` short.
  None of it reaches the route. tools/harness/rename-features.mjs:115-118 prints the reply and sleeps
  3000; the forward half at :82-84 does `check("the rename was accepted", said.did, said.detail)`.

##### `immediate-doc-contradicts-code` docs/debug-api.md says `immediate` only schedules and waits ten seconds; the handler waits, and its default is fifteen

- **Where:** `docs/debug-api.md:207`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** docs/debug-api.md:207-208 still reads "`immediate` only SCHEDULES. A statement that hits a
  breakpoint does not return until the developer continues, so an api that waited for it would jam its
  own connection." The same document at line 99 contradicts it - "`ran: false` means the evaluation
  did not finish inside the ten second wait" - and the code contradicts both: the route waits on the
  evaluation with `var deadline = Environment.TickCount64 + WaitMilliseconds(request, 15000);` and
  polls `evaluated.Wait(120)` while draining dialogs (DebugApi.cs:756-780), then answers `new
  DebugImmediateReply(ran, outcome, failed)` (792-794). The route's own header comment records the
  change: "THE OUTCOME, not the request. This posted the line to the host thread and answered `{ran:
  true}` without waiting" (DebugApi.cs:642-657). A stale copy of the old claim survives in the code
  too, in the summary of the method that dispatches every route: "The immediate route only schedules -
  a statement that hits a breakpoint does not return until the developer continues, and an api that
  waited on it would jam" (DebugApi.cs:496-499).
- **Why:** A reader deciding whether `immediate` can be used to run a statement that might break is told the
  opposite of what the code does, twice, by the reference document - and by the dispatcher's own
  summary, which is where the next person to touch this file will look. The ten-versus-fifteen second
  figure sets the wrong client timeout: xlide-api.mjs:934 hardcodes 20000, which happens to still
  clear 15000, but only by accident.
- **Change:** Delete the "only SCHEDULES" paragraph at docs/debug-api.md:207-208, correct "ten second wait" to the
  actual fifteen second default (and name `waitMs` as the override), and fix the stale sentence in the
  AnswerDebugRequest summary at DebugApi.cs:497-499.
- **Size:** three sentences across two files
- **Adversary:** docs/debug-api.md:207-208 reads '`immediate` only SCHEDULES. A statement that hits a breakpoint does
  not return until the developer continues, so an api that waited for it would jam its own
  connection.' The code waits: DebugApi.cs:756 `var deadline = Environment.TickCount64 +
  WaitMilliseconds(request, 15000);` with the poll at 758-778 and `new DebugImmediateReply(ran,
  outcome, failed)` at 792-794, under a header comment at 642-657 that records the change ('THE
  OUTCOME, not the request'). docs/debug-api.md:99 says 'ran: false means the evaluation did not
  finish inside the ten second wait' against a 15000 default. The dispatcher summary at
  DebugApi.cs:497-499 still says 'The immediate route only schedules ... an api that waited on it
  would jam.' xlide-api.mjs:934 is `timeout: 20000`.

##### `assert-shownmodule-reads-our-record` `assert that=shownModule` is satisfied by the shim's own field, not by the native pane or the page

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2104`
- **Kind:** api-coverage / small effort, claim observed, confidence supported, severity low
- **Evidence:** EvaluateClaim: `case "shownModule": { var shown = _editorSurface?.Module; return (shown is not null
  && (expected is null || shown.Equals(expected, StringComparison.OrdinalIgnoreCase)), shown ??
  "(none)"); }` (DebugApi.cs:2104-2110). `Module` is the shim's own record of the active document:
  `public string? Module => ActiveDoc?.Module;` (EditorSurface.cs:119). The same session has a live
  answer for the same question and does not use it here: the `native` route reads
  `_editor.GetObject("ActiveCodePane")` and walks to `component?.GetString("Name")`
  (DebugApi.cs:2517-2542). The route's own docstring claims the opposite of what it does - "Read from
  the snapshots the reader thread publishes and from fields the host thread writes, so a claim can be
  tested while that thread is busy" (2086-2089) - which is true of `localsHas`/`watchHas` and not of
  this one. `surfaceReady` (`_surfaceShown`, 2101) and `stopped`/`running` (`_inBreak`, 2095-2099) are
  the same shape; only `noDialogs` (2112) and `responsive` (2142) read something outside the session.
- **Why:** `shownModule` is the natural wait after `pane open` or `caret?module=`, and it holds the moment the
  shim updates its own field - which is precisely the state where the native pane and the surface can
  disagree, the disagreement the `native` route was added for on 2026-08-08. A guard comparing our
  record against our record cannot see the host move a pane underneath us.
- **Change:** Add a claim beside it rather than changing this one (the route shapes are stable): `nativeModule`,
  evaluated from ActiveCodePane the way `native` does. Failing that, say in the route table that
  `shownModule` asks the surface's record and point at `native` for the host's answer.
- **Size:** one new case in EvaluateClaim plus two doc rows
- **Adversary:** DebugApi.cs:2104-2110 is exactly `case "shownModule": { var shown = _editorSurface?.Module; ... }`,
  and EditorSurface.cs:119 is `public string? Module => ActiveDoc?.Module;`. The same file answers the
  host's version elsewhere - DebugApi.cs:2517 `_editor.GetObject("ActiveCodePane")` walked to
  `component?.GetString("Name")` at 2542. Checked the designation requirement:
  docs/debug-api.md:305-307 lists the claim names with no statement of which layer any of them reads,
  and the client docstring is no more specific, so the layer stop is undesignated under api rule (b).
- **Correction applied:** Drop the 'the route's own docstring claims the opposite of what it does' argument. The docstring at
  DebugApi.cs:2086-2089 says the claims are read 'from the snapshots the reader thread publishes and
  from fields the host thread writes' - `_editorSurface.Module` IS a field the host thread writes, so
  the docstring is accurate about thread-safety and simply silent about authority. The finding stands
  on the undesignated layer stop alone.

##### `client-cannot-send-keep-or-crop` The client exposes neither `keep` on `command` nor `selector`/`pad` on `capture`, so the only callers that use them build the URL by hand

- **Where:** `tools/harness/xlide-api.mjs:426`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** `command: (name) => call(`command${query({ name })}`, { method: "POST" }),` (xlide-api.mjs:426) - no
  `keep`. The mechanism is generic and documented: every host-thread route honours it
  (`RememberRaisedDialogs(standingBefore, keep: request.Query.ContainsKey("keep"))`, DebugApi.cs:1894)
  and docs/debug-api.md:96 lists `keep` as an argument of `command`, with the rule at line 279 "A
  request that means to open a dialog passes `keep=1` and what it opens is exempt." The only users are
  PowerShell probes hand-building the URL: `Invoke-RestMethod "$api/command?name=references&keep=1"`
  (Test-DebugApi.ps1:466) and `"$api/command?name=run&keep=1"` (514). A .mjs suite that opens a dialog
  on purpose therefore cannot protect it - tools/harness/immediate-watch.mjs:165 runs
  `api.command("addWatch")` and then has to sweep up after the guard: `const afterCommand = await
  api.dialogs(); if ((afterCommand.dialogs ?? []).length > 0) { ... await api.dismiss("Cancel") ... }`
  (172-177). Same shape on capture: `capture: (window) => call(`capture${query({ window })}`, { raw:
  true, timeout: 20000 }),` (xlide-api.mjs:424) drops the crop the route supports
  (DebugApi.cs:585-637) and the reference documents (docs/debug-api.md:66), so Get-Shot.ps1:55 appends
  `"&selector=$([uri]::EscapeDataString($Selector))&pad=$Pad"` itself and Test-DebugApi.ps1:416 does
  the same. Grepping the .mjs corpus for `.capture(` returns zero calls.
- **Why:** `keep` is a safety argument: without it a .mjs suite cannot open a dialog deliberately, so the
  deliberate-modal path is only ever exercised from PowerShell. And the crop is the mode that makes
  `capture` useful for a UI defect - a whole frame is a picture in which a 54-pixel drop zone cannot
  be seen, per the route's own comment - so the .mjs suites do not screenshot at all.
- **Change:** Add optional arguments to the two existing client methods: `command(name, { keep })` and
  `capture(window, { selector, pad })`. Both are pass-through query values; no route changes. Add
  `keep` to the `command` row in docs/driving-excel.md:203, which currently documents only
  `command(name)`.
- **Size:** two client signatures, one doc row
- **Adversary:** xlide-api.mjs:426 is `command: (name) => call(\`command${query({ name })}\`, { method: "POST" }),`
  and :424 is `capture: (window) => call(\`capture${query({ window })}\`, { raw: true, timeout: 20000
  }),`. Both arguments exist on the routes: `keep` is honoured generically at DebugApi.cs:1894
  `RememberRaisedDialogs(standingBefore, keep: request.Query.ContainsKey("keep"))` and again at 1903,
  and documented at docs/debug-api.md:96 and :279; `selector`/`pad` are implemented at
  DebugApi.cs:583-600 and documented in the capture row at docs/debug-api.md:66. The only users build
  the URL by hand: Test-DebugApi.ps1:466 and :514 for keep, Get-Shot.ps1:54-55 (`$query +=
  "&selector=$([uri]::EscapeDataString($Selector))&pad=$Pad"`) and Test-DebugApi.ps1:416 for the crop.
  Grepping tools/harness/*.mjs for `.capture(` returns nothing. immediate-watch.mjs:164-177 does the
  sweep-up the finding describes. docs/driving-excel.md:203 is `| \`command\` | \`command(name)\` |
  any editor command by name |` and :201 is `| \`capture\` | \`capture(window)\` | a BMP of the
  window, through PrintWindow |`.
- **Correction applied:** The capture client row is docs/driving-excel.md:201, not only the command row at :203; both need the
  new arguments.

##### `clientfor-not-exported` There is no exported way to build a client from a known base URL, so a probe handed `--api` re-implements the transport without the error check

- **Where:** `tools/harness/xlide-api.mjs:271`
- **Kind:** api-coverage / small effort, claim observed, confidence supported, severity low
- **Evidence:** `function clientFor(entry) { ... }` (xlide-api.mjs:271) is module-private; the exports are
  `whyDidItDie`, `wait`, `waitFor`, `waitUntilStable`, `discover` and `open`, and all of them route
  through discovery. A probe launched with an already-known base therefore writes its own:
  tools/harness/objbrowser-live-probe.mjs:11-22 `const apiBase = args[args.indexOf("--api") + 1]; ...
  async function api(route) { const reply = await fetch(`${apiBase}/${route}`, { method:
  route.startsWith("command") ? "POST" : "GET" }); return reply.json(); }`. That copy loses everything
  the real client's `call` does: the error-field check `if (answer && typeof answer === "object" &&
  "error" in answer) { throw new Error(...) }` (xlide-api.mjs:296-298), the abort timeout (274-276),
  and the host-death diagnosis (282-289). An error reply is a 200 with an `error` field
  (DebugServer.cs:216 writes "200 OK" for every answer including DebugError), so in that probe a
  failed route silently becomes an object whose fields are all undefined.
- **Why:** The probe with the weakest transport is the one driving the live page through CDP, where a route
  failing quietly reads as the page misbehaving. It also means the method verbs, timeouts and
  `keep`/`raw` handling that the client encodes have a second, divergent implementation that nothing
  keeps in step.
- **Change:** Export `clientFor` (or a thin `openAt(base)` that takes `http://127.0.0.1:port/token` and returns
  the same object), and have objbrowser-live-probe.mjs use it in place of its local `api()`.
- **Size:** one export plus a 4-line deletion in the probe
- **Adversary:** xlide-api.mjs:271 is `function clientFor(entry) {` with no export; the exported symbols are
  whyDidItDie (48), wait (110), waitFor (132), waitUntilStable (164), discover (180) and open (224).
  tools/harness/objbrowser-live-probe.mjs:12 takes `--api` and :19-22 re-implements the transport as
  `const reply = await fetch(\`${apiBase}/${route}\`, { method: route.startsWith("command") ? "POST" :
  "GET" }); return reply.json();`, which has none of the client's error-field throw
  (xlide-api.mjs:295-298), abort timeout (272-276) or death diagnosis (281-289).

##### `history-omits-polled-routes` `history` silently drops five routes from the "every request" transcript, and the reply does not say so

- **Where:** `src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:237`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** `RecordRequest` opens with `if (route is "history" or "log" or "journal" or "state" or "dialogs") {
  // The routes a probe polls would otherwise be the whole transcript. return; }`
  (DebugServer.cs:235-241). docs/debug-api.md:59 says "every request this door has served" and line
  316 repeats "hands back every request the door has served, plus a runnable script of them"; the
  route's own comment says "The session as a script ... this hands it back ready to run, so a bug
  found by hand becomes a probe by copying" (DebugApi.cs:885-889). Nothing in DebugHistoryReply
  (DebugServer.cs:1034) carries the exclusion list or a dropped count.
- **Why:** The replay script is missing exactly the calls that make a hand-driven session reproducible: the
  `log?match=...&waitMs=` waits and the `state` polls that sequence everything else. Copy the script
  out and the steps run back to back with no waits, so a session that only worked because of its waits
  replays as a race, and the transcript gives no hint that anything was removed.
- **Change:** Keep the exclusion (the reason for it is sound) but make it visible: add a `notRecorded` field
  listing the five route names, and emit a `# waited on log/state here` marker in the generated script
  where a polled route was skipped. Correct the two "every request" sentences in docs/debug-api.md.
- **Size:** one field plus two doc sentences
- **Adversary:** DebugServer.cs:235-241: `if (route is "history" or "log" or "journal" or "state" or "dialogs") { //
  The routes a probe polls would otherwise be the whole transcript. return; }`. DebugHistoryReply
  (DebugServer.cs:1034-1039) carries only requests, script and routeCosts - no exclusion list, no
  dropped count. docs/debug-api.md:59 says 'every request this door has served' and :316 repeats
  'hands back every request the door has served, plus a runnable script of them'. The script builder
  at DebugApi.cs:891-902 emits one Invoke-RestMethod line per recorded request with nothing between
  them, so the omitted `log?waitMs=` waits leave no marker.

##### `build-fixture-retry-unreachable` build-fixture's write-retry can no longer run for the refusal it was written for, because the `module` POST route learned to report it and the client turns that into a throw

- **Where:** `tools/harness/build-fixture.mjs:64`
- **Kind:** api-coverage / small effort, claim derived, confidence supported, severity low
- **Evidence:** The retry's stated purpose (build-fixture.mjs:40-49): "The first write to a module that is freshly
  added AND currently shown fails - the editor answers 'Invalid procedure call or argument' - and the
  write route reports that in the log rather than in its reply ... A second write always takes." The
  loop is `for (let attempt = 1; attempt <= 3; attempt++) { try { await api.writeModule(name, code); }
  catch (error) { if (!/did not answer in time/.test(String(error?.message))) { throw error; } ... }
  ... }` (53-69). That premise no longer holds: the route now answers the writer's complaint - `var
  complaint = WriteModule(moduleName, request.Body, projectId, hostRewrite: true); return complaint is
  null ? ...DebugCommandReply(true, 0) : ...DebugErrorReply(complaint)` (DebugApi.cs:3258-3263) - and
  WriteModule returns a string for every failure path including the AddFromString refusal (`return
  restored ? $"{component} was not written: {refusal.Message}. What it held before is back." : ...`,
  AddInSession.cs:1946-1954) and its outer catch (`return $"{component} could not be written:
  {ex.Message}";`, AddInSession.cs:2062). The client throws on any error field
  (xlide-api.mjs:296-298), and "Invalid procedure call or argument" does not match `/did not answer in
  time/`, so attempt 1 rethrows and attempts 2 and 3 are unreachable.
- **Why:** The fixture builder now hard-fails on the one transient it was built to absorb, instead of retrying.
  Fixtures are the precondition for the live suites, so this surfaces as a whole `-Live` pass that
  cannot start, with an error message about a module rather than about a retry that stopped working.
- **Change:** Widen the catch to retry on a refusal as well as on the door timeout - retry when the message
  matches the timeout OR the writer's refusal wording, rethrow otherwise - and update the comment at
  lines 40-49, which now describes behaviour the route no longer has.
- **Size:** one regex in one catch, plus a stale comment
- **Adversary:** tools/harness/build-fixture.mjs:38-49 states the premise ('the write route reports that in the log
  rather than in its reply'), and :53-66 is the loop whose catch rethrows anything not matching `/did
  not answer in time/`. The premise is stale: DebugApi.cs:3258-3263 now answers `WriteModule`'s
  complaint as a DebugErrorReply, WriteModule returns a string for the AddFromString refusal
  (AddInSession.cs:1946-1954) and from its outer catch (AddInSession.cs:2061-2063), and
  xlide-api.mjs:295-298 throws on any error field. The route's own comment at DebugApi.cs:3253-3257
  confirms the change and even names the fixture builder's line-count workaround as the thing that
  should not have been necessary.
- **Correction applied:** 'Attempts 2 and 3 are unreachable' is too strong. The loop has a second retry path that is
  untouched: the line-count poll at build-fixture.mjs:78-105 falls through to `console.log(... write
  ${attempt} did not take; asking again)` at :104 and loops, so a write that fails SILENTLY is still
  retried. Only a write the route now reports as an error is terminal on attempt 1. Scope the fix and
  the comment to that case.

##### `audit-isdriven-matches-console-log` audit-routes' driven-check for `log` is satisfied by `console.log(`, so the coverage guarantee it exists to enforce is vacuous for that route

- **Where:** `tools/harness/audit-routes.mjs:144`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** `const isDriven = (route) => [...(methods.get(route) ?? [])].some((method) => new
  RegExp(`\\.${method}\\s*\\(`).test(corpus)) || new
  RegExp(`/${escaped(route)}[?"'\`\\s]`).test(corpus);` (audit-routes.mjs:143-146). The corpus is
  every harness file except the client and the audit (137-140). For route `log`, methodsByRoute walks
  back from `call(`log${query({ since, match, max, waitMs })}`...)` to the nearest 4-space property,
  which is `log:` (xlide-api.mjs:421-422), so the method set is {log} and the test is `/\.log\s*\(/`.
  Counted over that corpus: 208 lines match `.log(`, of which 203 are `console.log(` and 5 are
  `api.log(`. The header this check was added under says the opposite is the point: "DOCUMENTED AND
  REACHABLE IS NOT COVERED ... A route with a doc row and a client method can still be a route nothing
  has ever called" (audit-routes.mjs:9-12). The exemption list itself is sound - nothing in the corpus
  calls `drainFinalizers` or `/drainfinalizers`, so the one NOT_DRIVEN_ON_PURPOSE entry still holds.
- **Why:** The gate would not notice if all five real `api.log(...)` calls were deleted, because 203
  `console.log(` lines keep the check green - and `log?waitMs` is the mechanism the harness's whole
  no-sleep policy rests on. The same bare-method-name match is one common identifier away from doing
  this to another route.
- **Change:** Anchor the method match to a receiver: test `/\b(?:api|client)\w*\.method\s*\(/` rather than
  `\.method\s*\(`, and print the file each route was found driven in so a vacuous match is visible in
  the ok line rather than only in a failure.
- **Size:** one regex; 203 false matches on one route today
- **Adversary:** tools/harness/audit-routes.mjs:143-146 is `[...(methods.get(route) ?? [])].some((method) => new
  RegExp(\`\\\\.${method}\\\\s*\\\\(\`).test(corpus))`, over the corpus built at :137-140 from every
  harness .mjs/.ps1 except the client and the audit. methodsByRoute (:100-121) walks back from the
  `call(` line to the nearest 4-space property, which for `call(\`log${query(...)}\`)` at
  xlide-api.mjs:421-422 is `log:`, so the test is `/\.log\s*\(/`. Counted over that exact corpus: 203
  occurrences of `console.log(` against 5 of `api.log(` (analysis-freshness.mjs:225,242;
  menu-bar.mjs:126,137; write-rollback.mjs:116). The file header at :9-14 states the check exists
  precisely to catch a route nothing calls. The drainfinalizers exemption at :28-37 is sound.

##### `assert-unknown-claim-looks-like-a-failed-claim` `assert` treats a misspelt claim name as a claim that is not yet true, polling it for the full timeout and answering the same shape as a real failure

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:924`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** The route polls until the deadline whatever the claim is: `while (true) { (held, saw) =
  EvaluateClaim(claim, expected); if (held || Environment.TickCount64 >= deadline) { break; }
  Thread.Sleep(150); }` (DebugApi.cs:924-933), with `timeout` defaulting to 10000 (916-917).
  EvaluateClaim's fallback is `default: return (false, $"unknown claim {claim}");` (2148-2149) - a
  value, not a refusal. So `assert?that=shownModul` sleeps ten seconds and answers
  `{"held":false,"claim":"shownModul","expected":"(none)","saw":"unknown claim shownModul"}`, which
  the client returns as a normal answer (`assert` is not error-shaped, so `call`'s error check at
  xlide-api.mjs:296 does not fire). The vocabulary is a string list in three places that can drift
  apart - the switch (2093-2150), docs/debug-api.md:306-307, and the client's docstring
  (xlide-api.mjs:592-594).
- **Why:** An assertion instrument that reports a typo as a failed assertion sends the reader to look at the
  product. The only distinguishing signal is prose inside `saw`, which a `check(..., answer.held,
  ...)` call site never reads, and each occurrence costs the caller its whole timeout.
- **Change:** Evaluate the claim once before entering the poll and answer `{"error":"unknown claim X; known claims
  are ..."}` immediately when it is unrecognised - the same treatment `bench` already gives an unknown
  `what` ("unknown benchmark {what}; try tabswitch, layout, or type", DebugApi.cs:1368) and `trip`
  gives an unknown scenario (1493-1498).
- **Size:** one pre-check; saves 10s per typo
- **Adversary:** DebugApi.cs:919-933: `var deadline = Environment.TickCount64 + timeout;` (timeout clamped, default
  10000, at 916-917) then `while (true) { (held, saw) = EvaluateClaim(claim, expected); if (held ||
  Environment.TickCount64 >= deadline) { break; } Thread.Sleep(150); }`, with no pre-check.
  EvaluateClaim's fallback at 2147-2148 is `default: return (false, $"unknown claim {claim}");` - a
  value, so the reply is a normal DebugAssertReply (935-937) and the client's error check
  (xlide-api.mjs:295-298) does not fire. The precedent the finding cites is real: DebugApi.cs:1368
  `return DebugError($"unknown benchmark {what}; try tabswitch, layout, or type");` and 1492-1498 for
  trip.

### AddInSession.cs

_8 findings, from the `complexity-session` finder._

##### `hostapp-wrapper-survives-stop` _hostApp holds a live DispatchObject on Excel's Application that Stop() and Dispose() never release, so a session torn down inside a living process abandons a COM wrapper to the finalizer thread

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5664`
- **Kind:** perf / small effort, claim observed, confidence verified, severity high
- **Evidence:** Field: `private DispatchObject? _hostApp;` (AddInSession.cs:5664), populated by `_hostApp ??=
  HostApplication.Find();` at 5685 (WorkbookSaved) and 6583 (SaveWorkbookOf). `grep -n "_hostApp"
  src/Xlide.Vbe.Shim/AddIn/*.cs` returns exactly 11 hits and the only two Dispose calls are inside
  catch blocks: line 5708 (`// The application answer went stale ...` then `_hostApp?.Dispose();
  _hostApp = null;`) and line 6609 (same pair after `Log.Error($"close: {display} could not be saved",
  ex);`). I read Stop() in full (7593-7669): it disposes _debugServer, _analysis, _frameSubclass,
  _immediateReader, _ghostReaders, _typeLibraries, _browserPalette, _codePanes, _hostChrome,
  _editorSurface, and restores both ghost palettes; _hostApp appears nowhere. Dispose() (7733-7740) is
  `Stop(); _addIn?.Dispose(); _editor.Dispose();`. I also confirmed _hostApp is the only cached
  DispatchObject-typed field besides _editor/_addIn (grep for `private ... DispatchObject` in the file
  returns only 27, 39, 5664 as fields). DispatchObject.Dispose
  (src/Xlide.Vbe.Shim/Com/DispatchObject.cs:745) documents what an undisposed one costs: "Releasing
  one from the finalizer thread is not slow or untidy, it is invalid: it read as an access violation
  inside `Marshal.Release`, which ahead-of-time compilation cannot throw and so turns into a FailFast
  that takes the whole of Excel with it."
- **Why:** WorkbookSaved is called from PublishModules on every poll tick, so _hostApp is populated in
  essentially every session. At a real host shutdown the process dies and nothing is noticed, but
  Stop()+Dispose() also run when the process keeps living: XlideAddIn.cs:208-210 tears the old session
  down on a watchdog revival (`var stopped = _session; _session = null; stopped?.Dispose();`) after a
  cancelled Excel shutdown, and OnDisconnection with a non-HostShutdown mode does the same. In both
  cases an abandoned wrapper on Excel's Application object sits on the heap until some later GC
  finalizes it on the finalizer thread, which is the exact FailFast the codebase already paid for
  three times. com-leak.mjs cannot catch this: it compares stats.comWrappersLive across routes while
  the session is alive, and this wrapper is only orphaned at teardown, when nothing is left to answer
  the debug api.
- **Change:** Add `_hostApp?.Dispose(); _hostApp = null;` to Stop(), beside the other automation-reference
  releases (after _codePanes, before _hostChrome). Two lines.
- **Size:** 2 lines added; removes one leaked COM wrapper per non-shutdown session teardown
- **Adversary:** Verified end to end. src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5664 declares `private DispatchObject?
  _hostApp;`, populated by `_hostApp ??= HostApplication.Find();` at 5685 and 6583. A repo-wide grep
  for _hostApp returns hits only in AddInSession.cs, and the only Dispose calls are the two
  stale-answer catch blocks at 5708 and 6609. I read Stop() (7593-7666) in full: it releases
  _debugServer, _analysis, _frameSubclass, _immediateReader, _ghostReaders, _typeLibraries,
  _browserPalette, _codePanes, _hostChrome, _editorSurface and both ghost palettes; _hostApp is
  absent. Dispose() at 7733-7740 is `Stop(); _addIn?.Dispose(); _editor.Dispose();`.
  src/Xlide.Vbe.Shim/Interop/HostApplication.cs:49 returns `window?.GetObject("Application")` with the
  doc saying "The caller owns the result", so this is an owned wrapper.
  src/Xlide.Vbe.Shim/Com/DispatchObject.cs:743-757 has no finalizer and Dispose is the only route to
  ComRuntime.GiveBackWrapper; the inner wrapper is a CreateObjectFlags.UniqueInstance ComObject
  (src/Xlide.Vbe.Shim/Com/ComRuntime.cs:70) whose Finalize releases on the finalizer thread, the exact
  FailFast documented at DispatchObject.cs:724-740. Non-shutdown teardown in a living process is real:
  XlideAddIn.cs:208-210 `var stopped = _session; _session = null; stopped?.Dispose();` on watchdog
  revival, and OnDisconnection at 250-252 for a non-HostShutdown mode. The harness cannot see it:
  tools/harness/com-leak.mjs:103 and 424 both read api.stats() while the session is alive, and
  _hostApp is taken once per session (steady, not growing) and only orphaned at teardown.

##### `getitem-leaks-on-throw` Two collection loops bind GetItem to a plain local instead of a using, so an exception between the take and the manual Dispose leaks the wrapper

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:6010`
- **Kind:** perf / small effort, claim derived, confidence verified, severity medium
- **Evidence:** `grep -n "GetItem(" AddInSession.cs | grep -v "using var"` returns exactly two sites.
  FindProjectByDisplayName, 6008-6031: `var project = projects!.GetItem(i);` then `var matches =
  wantedId is not null ? string.Equals(ProjectReader.Identity(project).Id, wantedId, ...) :
  string.Equals(WorkbookDisplayName(project), displayName, ...); if (matches) { return project; }
  project.Dispose();` -- and the whole loop sits inside a `try { ... } catch (Exception ex) {
  Log.Info($"project: '{displayName}' could not be looked up ..."); }` at 6028, which swallows the
  throw and never disposes the in-hand item. FindComponent, 6997-6007: `var candidate =
  components.GetItem(j); if (candidate?.GetString("Name") == component) { foundProject = identity;
  return candidate; } candidate?.Dispose();` -- here there is no try at all, so the throw escapes and
  the wrapper is simply dropped. Both reads can throw: ProjectReader.Identity
  (src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:92) calls `project.GetString("Name")` outside its own
  try, and WorkbookDisplayName (AddInSession.cs:5969) does the same after its FileName guard; the
  surrounding comments ("Unsaved: the property raises rather than answering empty") establish that
  these property reads do raise.
- **Why:** The manual-dispose pattern is correct on the happy path and only correct there. FindComponent is the
  single most-called COM helper in the file (it backs the write path, rename, close, the object
  browser scan, replace-all and PublishDocument), so any transient refusal from the editor
  mid-enumeration -- exactly the state UpdateDebugState's own catch is written for ("refused while the
  editor runs the developer's code") -- drops a live wrapper. One dropped wrapper is a FailFast
  waiting for the next GC, per DispatchObject.Dispose's own note, not a byte of wasted memory.
- **Change:** In both loops, take the item into a `using var` and, on the match branch, hand ownership out
  deliberately -- either add a `DispatchObject.Detach()`/`Release()`-style transfer, or keep the plain
  local and wrap the body in `try { ... } catch { candidate.Dispose(); throw; }`. The second form is
  about three lines per site and changes no happy-path behaviour.
- **Size:** ~6 lines across 2 sites
- **Adversary:** Both sites are exactly as described. `grep -n "GetItem(" AddInSession.cs | grep -v "using var"`
  returns precisely two lines, 6010 and 6999. At AddInSession.cs:6010 `var project =
  projects!.GetItem(i);` is followed by ProjectReader.Identity(project) or
  WorkbookDisplayName(project) and a manual `project.Dispose();` at 6023, all inside the try whose
  catch at 6027-6030 swallows the throw without disposing the item in hand. At 6999 `var candidate =
  components.GetItem(j);` then `candidate?.GetString("Name")` then `candidate?.Dispose();` with no try
  at all, so a throw drops the wrapper and propagates. The reads can throw:
  src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:92 `var name = project.GetString("Name") ??
  "VBAProject";` sits outside Identity's own try, AddInSession.cs:5969 does the same after
  WorkbookDisplayName's FileName guard, and DispatchObject.GetString (Com/DispatchObject.cs:111) goes
  through InvokeCore, which throws on hr<0 (DispatchObject.cs:715-719). The consequence is the
  finalizer-thread release documented at DispatchObject.cs:724-740.
- **Correction applied:** The frequency claim is unverified: nothing I read establishes that FindComponent is "the single
  most-called COM helper in the file", and no observed throw from GetString("Name") mid-enumeration
  exists in the repo. The finding is a latent leak-on-throw window, not a reproduced one. Size is
  right: two sites, about six lines.

##### `engine-roundtrip-shape-x13` Thirteen copies of the same capture-hop-deadline-map-marshal block, seven of them opening with a byte-identical guard line

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:4363`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity low
- **Evidence:** `grep -n "if (surface is null || module is null || source is null || _analysis is not { }
  analysis)"` returns seven identical lines: 4369, 4419, 4463, 4512, 4564, 4609, 5285. `grep -n
  "CancellationTokenSource(TimeSpan.FromSeconds"` returns thirteen inside the request handlers (4381,
  4431, 4475, 4530, 4576, 4621, 4667, 4724, 4788, 4931, 5211, 5254, 5314). I read
  OnCompletionRequested (4363-4406), OnHoverRequested (4413-4454), OnSignatureHelpRequested
  (4457-4499), OnSmartEnterRequested (4506-4552), OnCanonicalCaseRequested (4558-4597),
  OnLoopSyncRequested (4603-4642) and OnCodeActionsRequested (4650-4691) in full; each is `var surface
  = _editorSurface; var module = surface?.Module; var source = surface?.Text;` then the guard with an
  empty reply, then `_ = Task.Run(async () => { T result = default; try { using var deadline = ...;
  var answered = await analysis.XAsync(...).ConfigureAwait(false); if (answered is not null) { result
  = map(answered); } Log.Info($"x: {module}@{offset} -> ..."); } catch (Exception ex) { Log.Info($"x:
  {module}@{offset} failed ({ex.GetType().Name})"); } surface.RunOnHostThread(() =>
  surface.ShowX(requestId, result)); });`. The only per-handler differences are the analysis method,
  the projection lambda, the empty value, the log verb and the deadline.
- **Why:** Any change to the shared policy -- the deadline, the ConfigureAwait, the failure log shape, the
  empty-reply contract, or adding cancellation when a newer request supersedes an older one -- has to
  be made thirteen times, and a handler that misses one drifts silently because each answers empty on
  failure and therefore never complains. The drift has already started: OnCodeActionsRequested (4650)
  omits `source` from its guard and its Log.Info sits inside `if (actions.Length > 0)` while
  OnCompletionRequested logs unconditionally, so the log tells you nothing consistent about which
  requests were served.
- **Change:** One generic helper on the session, e.g. `private void AnswerFromEngine<T>(int requestId, T empty,
  TimeSpan deadline, string verb, Func<AnalysisService, string, string, CancellationToken, Task<T>>
  ask, Action<int, T> reply)` doing the capture, guard, Task.Run, deadline, catch, log and
  RunOnHostThread. Each handler then shrinks to its guard-specific empty value, its ask lambda and its
  projection. Generics are fine under NativeAOT here -- every instantiation is a closed type known at
  compile time, no reflection.
- **Size:** roughly 200 of about 400 lines
- **Adversary:** The greps reproduce: the identical guard line appears at AddInSession.cs:4369, 4419, 4463, 4512,
  4564, 4609 and 5285, and there are thirteen CancellationTokenSource(TimeSpan.FromSeconds sites in
  the handlers (4381, 4431, 4475, 4530, 4576, 4621, 4667, 4724, 4788, 4931, 5211, 5254, 5314). I read
  OnCompletionRequested (4363-4406) and OnCodeActionsRequested (4650-4691) in full and they are the
  same capture-guard-Task.Run-deadline-map-catch-RunOnHostThread block differing only in the analysis
  call, the projection, the empty value, the log verb and the deadline. The drift claim also checks
  out: 4655 guards on `surface is null || module is null || _analysis is not { } analysis` without
  source, and its Log.Info at 4678 is inside `if (actions.Length > 0)` while 4688 logs failures
  unconditionally.
- **Correction applied:** The count of genuinely-fungible handlers is seven, not thirteen, and the seven identical guards are
  not the seven handlers listed. 5285 is OnLiveAnalysisDue (5279-5340), which has no requestId, no
  reply and an extra staleness re-check on the host thread, so it does not fit the proposed helper;
  OnCodeActionsRequested (4650) is fungible but has the different guard. The other five deadline sites
  diverge materially: OnOutlineRequested (4699) carries a `failed` flag and passes projectId with
  source:null, OnSemanticTokensRequested (5254) likewise, OnRenameRequested (4788) and
  OnModuleRenameRequested (4931) return six-field outcomes and do host-thread write-backs,
  OnNavigationRequested (5211) captures the workbook. The proposed signature
  Func<AnalysisService,string,string,CancellationToken,Task<T>> fits only the seven. Size is roughly
  60-80 boilerplate lines across about 260, not "roughly 200 of about 400".

##### `publishmodules-per-tick-workbook-walk` PublishModules walks Excel's Workbooks collection once per distinct workbook on every poll tick, and the change-key that would make it free is computed only after that work is done

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5574`
- **Kind:** perf / medium effort, claim derived, confidence supported, severity low
- **Evidence:** PollDebugState calls `PublishModules();` unconditionally at 3934, with the comment "Every tick, not
  only the resync ones ... The change-key inside makes an unchanged strip cost a read and no message."
  The tick interval comes from UpdatePolling (3703-3707): `_resyncPanePolls > 0 ?
  ClosingPollMilliseconds : _pollsRemaining > 0 ? DebugPollMilliseconds : ...`, i.e. 16ms after a tab
  close (3726) and 150ms through a debug episode (2244). Inside PublishModules the key is built at
  5647 and compared at 5649, AFTER line 5641 `bool[] dirty = [.. modules.Select(m => DirtyOf(m.Name,
  m.Project))];`. DirtyOf (5587) memoises per call in `savedByProject` (5574) but that dictionary is
  local to the call, so each tick re-runs WorkbookSaved once per distinct workbook. WorkbookSaved
  (5681-5713) is `_hostApp.GetObject("Workbooks")`, `books.GetInt32("Count")`, then per workbook
  `books.GetItem(i)` + `book.GetString("Name")` until it matches, plus `book.GetBool("Saved")`. Ahead
  of that, ReadOpenModules (5773-5841) costs roughly eight IDispatch invokes per open pane: GetItem,
  CodeModule, Parent, Name, Collection, Parent, ProjectReader.Identity (itself Name + FileName).
  Everything else on the same tick is small by comparison: ProjectCount is two calls (3819-3820),
  UpdateDebugState is about seven (3543, 3593, 3600, 3608-3610).
- **Why:** With P open panes, D distinct workbooks among them and W workbooks in Excel, each tick costs about
  8P + D*(2 + 2W) IDispatch invokes on the VBE host thread before it discovers that nothing changed,
  against roughly ten for everything else on the tick combined. At the 16ms close-resync interval --
  chosen precisely because the work there was believed cheap -- that is the same walk sixty times a
  second while the developer watches a tab disappear. The comment at 3930-3933 states the cost as "a
  read", which understates it by an order of magnitude, so anyone tuning the tick is reasoning from a
  wrong number.
- **Change:** Read Excel's Workbooks once per PublishModules call into a name -> Saved map and have DirtyOf
  consult that, turning D collection walks into one. That also collapses the duplicated walk:
  WorkbookSaved (5691-5702) and SaveWorkbookOf (6589-6602) are the same
  `_hostApp.GetObject("Workbooks")` / Count / GetItem / match-on-Name loop written twice, differing
  only in what they do on the match. Before changing the tick itself, add a PerfCounters stamp around
  PublishModules the way PlacementFast/PlacementFull are stamped (7251, 7336) -- there is no counter
  for it today, which is why the cost is asserted rather than measured.
- **Size:** unmeasured; structurally 8P + D*(2+2W) IDispatch invokes per tick vs ~10 for the rest of the tick
- **Adversary:** The structure is exactly as claimed. AddInSession.cs:3934 calls PublishModules() unconditionally
  each tick. Inside, `bool[] dirty = [.. modules.Select(m => DirtyOf(m.Name, m.Project))];` is at 5641
  and the change-key is only built at 5646-5648 and compared at 5649, so the walk always precedes the
  cheap exit. DirtyOf (5587-5637) memoises in savedByProject, declared at 5574 and local to the call,
  so WorkbookSaved runs once per distinct workbook per tick. WorkbookSaved (5681-5713) is
  GetObject("Workbooks") + Count + a GetItem/GetString("Name") pair per workbook until match +
  GetBool("Saved"). ReadOpenModules (5773-5841) is about eight invokes per pane (GetItem, CodeModule,
  Parent, Name, Collection, Parent, Identity's Name plus FileName). The duplicated walk is real:
  WorkbookSaved 5691-5702 and SaveWorkbookOf 6589-6601 are the same
  Workbooks/Count/GetItem/match-on-Name loop, differing only in Saved versus Invoke("Save"). And the
  cost is genuinely unmeasured: src/Xlide.Vbe.Shim/Diagnostics/PerfCounters.cs exposes PlacementFull,
  PlacementFast, Marshal, WindowEvent, Refresh, Follow, LogLine, Poll and Beat, and nothing for
  publish.
- **Correction applied:** The hotness rhetoric is wrong. UpdatePolling (AddInSession.cs:3703-3707) sets interval 0 unless a
  resync burst, a debug episode, immediate watching or an empty workspace is in play, so this is not a
  continuous background poll. ClosingPollMilliseconds=16 (3726) applies to "a handful of ticks over
  about a tenth of a second, only after a close" (comment at 3699-3701), i.e. roughly six ticks, not
  "sixty times a second while the developer watches a tab disappear". The sustained worst case is
  DebugPollMilliseconds=150 (2244) for up to 20s while stopped, then 300 (2247) and 1000 (3728).
  Severity is accordingly low, and the per-tick invoke arithmetic is derived from counting call sites,
  not measured - a PerfCounters stamp around PublishModules is what would settle whether it matters.

##### `import-reads-module-twice` The import path reads the whole module twice before writing it, and the second read is provably the same text as the first

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:1903`
- **Kind:** perf / small effort, claim derived, confidence verified, severity low
- **Evidence:** WriteModule line 1903: `var wasHoldingBefore = keepEveryCharacter ? ProjectReader.ReadSource(found)
  : null;` with the comment "One read, on the import path only, which is already the slow one." Line
  1915: `var wroteDiff = baseline is not null && TryWriteLineDiff(module, baseline, text);`. Line
  1928, inside `if (!wroteDiff)`: `var wasHolding = ProjectReader.ReadSource(found);`. `grep -rn
  "keepEveryCharacter" src/` shows exactly one caller passing true: line 395, the sync/import path.
  TryWriteLineDiff (2110-2193) returns false only at 2114 (`baseline == text`, before any call) and
  2142 (`oldWindow > LargestDiffLines || newWindow > LargestDiffLines`, also before any call); every
  mutation is below that guard and any failure there rethrows (2188) rather than returning false. So
  on the import path with no baseline, wroteDiff is false, nothing has touched the module between 1903
  and 1928, and the two reads return identical text. The cost of one such read is measured in the
  code's own comment at 1922-1924: "3ms of a 1,037ms write at 1,002 lines, 66ms of a 12,594ms write at
  40,002."
- **Why:** Every module of a first import pays one extra whole-module COM read on the VBE host thread, at the
  size the author already measured: 3ms per thousand-line module, 66ms per forty-thousand-line one. A
  fifty-module repository import is roughly 150ms of pure duplication in an operation the user is
  already waiting on. It is free to remove and the correctness argument is settled by
  TryWriteLineDiff's own structure.
- **Change:** Replace line 1928 with `var wasHolding = wasHoldingBefore ?? ProjectReader.ReadSource(found);`. One
  line. The comment above it should keep its measurement but note that the import path already has the
  copy in hand.
- **Size:** 1 line; measured 3ms per 1,002-line module, 66ms per 40,002-line module, per import
- **Adversary:** Verified line by line. AddInSession.cs:1903 `var wasHoldingBefore = keepEveryCharacter ?
  ProjectReader.ReadSource(found) : null;` and 1928 `var wasHolding =
  ProjectReader.ReadSource(found);` inside `if (!wroteDiff)` at 1917. `grep -rn keepEveryCharacter
  src/` shows the only caller passing true is line 395, the import path. Between the two reads sit
  only a dictionary lookup for writtenKey/baseline and the TryWriteLineDiff call at 1915.
  TryWriteLineDiff (2110-2190) has exactly one `return false`, at 2142, guarded by oldWindow/newWindow
  against LargestDiffLines and reached before any COM call - the first COM touch is
  GetInt32("CountOfLines") at 2154 and the first mutation is DeleteLines at 2161. Its other exits
  return true (2116 on baseline==text with no call, 2190 after a completed write) and its failure path
  rethrows at 2186. So on the import path with wroteDiff false the module is untouched between 1903
  and 1928 and the two reads are identical. The cost is the code's own measurement at 1922-1924: 3ms
  of a 1,037ms write at 1,002 lines, 66ms of a 12,594ms write at 40,002.
- **Correction applied:** One detail to keep in mind when applying the fix: wasHoldingBefore is not spare, it already backs
  the character-loss restore at 1982 (`PutModuleBack(module, wasHoldingBefore)`), so reusing it at
  1928 is a read elision only and changes no restore semantics. Severity is low - import is not a
  frequent operation, and 3ms per thousand-line module is half a percent of the write it accompanies.

##### `readsource-reimplemented-three-times` ProjectReader.ReadSource is reimplemented inline at three sites that already have the component in hand

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:2536`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** The shared helper exists and is used: ResyncFromModule (3777) `var stored = found is null ? null :
  ProjectReader.ReadSource(found);`, and WriteModule at 1903, 1928 and 1967. Its body
  (src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:199-216) is `using var code =
  component.GetObject("CodeModule"); ... var lines = code.GetInt32("CountOfLines"); return lines <= 0
  ? string.Empty : code.GetStringIndexed("Lines", 1, lines) ?? string.Empty;` inside a try. Three
  sites hand-roll the same thing after the same `FindComponent(...)`: ScanProjectMembers 2536-2539
  (`using var component = FindComponent(module, projectId, out _); using var code =
  component?.GetObject("CodeModule"); var count = code?.GetInt32("CountOfLines") ?? 0; var source =
  count > 0 ? code!.GetStringIndexed("Lines", 1, count) : null;`), PublishDocument 4073-4082 (same
  shape, `?? string.Empty`), and CaptureBefore 5005-5012 (same shape, `?? string.Empty`). `grep -rn
  'GetStringIndexed("Lines"' src/` confirms these are the only whole-module reads outside
  ProjectReader.
- **Why:** ReadSource's doc records the edge case the copies each re-guard by hand -- "An empty module has no
  lines at all, and asking such a module for line one raises rather than returning nothing." All three
  copies happen to get it right today, so this is a maintenance cost rather than a live bug: the next
  hand-rolled copy, or a change to how an empty module is represented, has four places to keep in step
  instead of one.
- **Change:** Replace each of the three inline reads with `ProjectReader.ReadSource(component)`. Note that
  ReadSource logs via Log.Error while two of the three sites currently log at Verbose on failure, so
  decide once whether a failed module read is an error or a verbosity and let the helper own it.
- **Size:** ~9 lines removed across 3 sites
- **Adversary:** `grep -rn 'GetStringIndexed("Lines"' src/` returns exactly five hits: the helper at
  Engine/ProjectReader.cs:210, a single-line read at AddInSession.cs:2829, and the three whole-module
  copies at AddInSession.cs:2539 (ScanProjectMembers, 2534-2540), 4082 (PublishDocument, 4073-4083)
  and 5012 (the rename undo capture, 5005-5012). Each copy is the same FindComponent ->
  GetObject("CodeModule") -> GetInt32("CountOfLines") -> GetStringIndexed("Lines", 1, count) sequence
  that ProjectReader.ReadSource (199-216) already owns, including the empty-module guard its doc
  explains.
- **Correction applied:** The replacement is not purely mechanical at two of the three sites. At AddInSession.cs:4076-4080 a
  null CodeModule produces a distinct `document: {moduleName} could not be found to publish` log and
  an early return, whereas ReadSource collapses "no CodeModule" and "the read threw" into a single
  null; the same distinction drives the `continue` at 5007. ReadSource also logs at Log.Error
  (ProjectReader.cs:214) where all three sites log at Log.Verbose. About nine lines removed, with
  those two decisions made once.

##### `misattached-xml-summaries` Ten members carry two <summary> blocks, the first belonging to a different member which is left undocumented

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:6420`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** `awk '/<\/summary>/{prev=NR} /<summary>/{if(NR==prev+1) print NR}'` over the file returns ten sites:
  131, 1247, 3720, 4007, 4974, 5720, 6420, 6875, 6932, 7224. I opened seven. 6420: the block
  "Publishes every finding to the surface's panel, across all modules - except the ones the
  active-line hold is keeping back" sits immediately above `private void DropFindingsFor(...)` at 6425
  and describes PublishFindingsToSurface at 6445. 5720: "What the developer calls a project, from its
  identity alone" sits above `private static string? ContentKey(string? text)` at 5731 and describes
  DisplayFromProjectId at 5742. 6932: "Finds a component by name across every open project" sits above
  `private string? ActivePaneOwner(string component)` at 6937 and describes FindComponent at 6969.
  7224: "Recomputes where the surface belongs, now ... a menu item has just opened or closed a native
  window" sits above `private void PlaceSurfaceFast()` at 7230 and describes RefreshSurfacePlacement
  at 7281, which has no doc at all. 3720: "How often the pane picture is re-derived in the moments
  after a close" and "How often the project tree is refreshed while the editor has no panes" are
  stacked above `ClosingPollMilliseconds` at 3726, leaving `EmptyWorkspacePollMilliseconds` at 3728
  bare. 1247 and 4974 are the same shape. 6875 is different: both summaries describe ShowModule, the
  second being a returns contract in a summary tag.
- **Why:** These doc comments are the map of a 7,741-line file and they are unusually load-bearing -- they
  carry the measurements, the dates and the failure histories that justify the code. A reader who
  hovers DropFindingsFor is told it publishes every finding across all modules; a reader who hovers
  ContentKey is told it names workbooks. Ten members are lying about what they do, and ten more have
  lost their explanation. Nothing catches it: the build does not generate a documentation file, so
  CS1571 (duplicate tag) never fires despite TreatWarningsAsErrors being true in
  Directory.Build.props.
- **Change:** At each of the ten sites, move the leading summary down onto the member it describes; at 6875 turn
  the second summary into a <returns>. Zero code lines change. Consider turning on
  GenerateDocumentationFile for this project so the duplicate-tag warning keeps it from recurring.
- **Size:** 10 doc blocks relocated, 0 code lines
- **Adversary:** I ran the same detection and it returns exactly the ten line numbers given: 131, 1247, 3720, 4007,
  4974, 5720, 6420, 6875, 6932, 7224. Spot-checked three. At 6415-6424 the summary "Publishes every
  finding to the surface's panel, across all modules - except the ones the active-line hold is keeping
  back" sits directly above DropFindingsFor at 6425 and describes PublishFindingsToSurface at 6445,
  which has no doc. At 6931 the one-line summary "Finds a component by name across every open project,
  or null when there is none." sits above ActivePaneOwner at 6937 and belongs to the FindComponent
  overload at 6960, which is bare. At 3715-3726 the summary about the project tree refresh with no
  panes stacks above the ClosingPollMilliseconds summary, leaving EmptyWorkspacePollMilliseconds at
  3728 undocumented. Directory.Build.props:8 does set TreatWarningsAsErrors true and does not set
  GenerateDocumentationFile, so no XML doc file is produced today.
- **Correction applied:** The proposed guard would not work. There is no C# diagnostic for a duplicated <summary> tag: CS1571
  is "XML comment has a duplicate param tag" and CS1710 is the typeparam equivalent, neither of which
  these sites trip. Turning on GenerateDocumentationFile would emit both summaries into the XML and
  warn about nothing, so the recurrence prevention has to be a review habit or a custom check, not a
  compiler warning. The ten relocations themselves stand.

##### `watchdog-ticks-write-only` _watchdogTicks is a counter for a patience budget that was removed, and two comments still refer to the budget as if it exists

- **Where:** `src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs:38`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** `grep -n "_watchdogTicks" src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs` gives declaration at 38, resets at
  130 and 260, `_watchdogTicks++;` at 170, and one read at 171 inside `Log.Verbose($"watchdog: no
  frame yet, tick {_watchdogTicks}");`. It gates nothing. The comment it belongs to says the opposite
  of a budget: "A dying process takes the watchdog with it, and in a living one the tick costs a
  window enumeration, so there is no budget to spend - standing down here is what would strand a
  cancelled shutdown whose editor window did not survive" (164-168). A second comment at 186-187 then
  says the enabled-tick wait "does not spend the patience budget above, because a dialog can sit
  unanswered for minutes" -- referring to a budget that the comment fourteen lines earlier says does
  not exist. Repo-wide `grep -rn "_watchdogTicks" src tools docs` returns 5 hits, all in this file.
- **Why:** Small, but it is a live trap in the one file whose correctness the codebase calls "the single most
  important line": a maintainer reading OnWatchdogTick sees a counter and a comment promising a
  give-up budget, goes looking for the threshold, and finds none. The counter itself is legitimate as
  verbose diagnostics under the standing verbose-logging directive; the stale cross-reference is not.
- **Change:** Keep the counter if the verbose line is wanted, but rename the comment's "patience budget above" to
  what is actually above it (an unbounded wait for the editor frame), or drop the field and log the
  enabled-tick count that does gate the revival. Either way, one comment must stop promising a
  mechanism that was deliberately removed.
- **Size:** 1 field and 2 comment sentences; few lines but misleading
- **Adversary:** `grep -rn _watchdogTicks src tools docs` returns five hits, all in
  src/Xlide.Vbe.Shim/AddIn/XlideAddIn.cs: declaration at 38, resets at 130 and 260, increment at 170,
  and the single read at 171 inside `Log.Verbose($"watchdog: no frame yet, tick {_watchdogTicks}");`.
  It gates nothing - the revival is gated by _watchdogEnabledTicks at 202. The stale cross-reference
  is real: the comment at 186-187 says the enabled-tick wait "does not spend the patience budget
  above, because a dialog can sit unanswered for minutes", and there is no such budget.
- **Correction applied:** Two things in the finding are wrong. The dimension is not dead code: the field is read at line 171
  by a live Log.Verbose call, which the standing verbose-dev-logging directive protects, so nothing
  should be deleted. And it is one stale comment, not two: the comment at 164-168 explicitly states
  "there is no budget to spend", which is the correct description, so only the sentence at 186-187
  promises a removed mechanism. The fix is one sentence.

### AddInSession.DebugApi.cs and Diagnostics

_10 findings, from the `complexity-debugapi` finder._

##### `guard-fallthrough-to-host` The pool-side route switch has no default, so an argument-rejected route is marshalled to the host thread and answered by the host's timeout machinery

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1837`
- **Kind:** complexity / medium effort, claim derived, confidence verified, severity low
- **Evidence:** The first switch opens at line 516 (`switch (request.Route)`) and closes at 1837 (`}`) with no
  `default:` label anywhere between them - `grep -n "^\s*default:"` over the file returns only 1492
  (inside `trip`), 2148 (EvaluateClaim), 2870 (component), 2931 (pane) and 3323 (the on-host switch).
  Eleven of the cases in that range are guarded on their arguments, counted with `grep -cE '^\s+case
  "[a-zA-Z]+"( |$).*when'` over lines 1..2288. Two of them are the routes that exist for a blocked
  editor: line 1783 `case "dismiss" when request.Query.TryGetValue("button", out var button) &&
  button.Length > 0:` and line 909 `case "assert" when request.Query.TryGetValue("that", out var
  claim) && claim.Length > 0:`. When a guard does not hold, control leaves the switch and reaches line
  1839 `var host = _editorSurface;` / 1842 `return DebugError("the surface is not up yet");`, then
  1861 `host.RunOnHostThread(...)` and 1878 `var answered = done.Wait(TimeSpan.FromSeconds(3));`. The
  comment at 503-513 states the design intent this breaks: "dialogs, dismiss and guard all return
  before reaching it, and those are exactly the routes a caller uses while something is standing."
- **Why:** `dismiss` with a missing or empty `button` while a modal is standing does not answer "dismiss needs
  a button". It falls out of the switch, is queued on the host thread that the modal owns, waits the
  full three seconds, and answers `{"error":"the host thread did not answer in time"}` - the exact
  diagnosis the route exists to avoid, delivered for an argument mistake. Same for `assert` with an
  empty `that`. With no surface up, any argument-rejected pool route answers "the surface is not up
  yet", which sends the reader at the wrong layer. The file already paid for this class of confusion
  once: the default at 3323 carries a six-line apology written on 2026-08-07 after `caret?line=-1` was
  answered "unknown route caret".
- **Change:** Validate arguments once, before the switch, from a static route-to-required-arguments table (route
  name, argument name, parser, constraint), and answer the precise refusal from there. That lets the
  seventeen `when` clauses become plain `case` labels, lets the pool switch fall through to the host
  marshal only for routes that genuinely belong there, and lets the default at 3323 stop hedging.
- **Size:** net roughly zero lines (about 25 added for the table, 17 `when` clauses and the 6-line hedging default simplified); the change is behavioural, not size
- **Adversary:** The structure is as described. src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:516 opens `switch
  (request.Route)` and 1837 closes it; `grep -n "^\s*default:"` over the file returns 1492 (inside the
  trip sub-switch), 2148 (EvaluateClaim), 2870 and 2931 (inside on-host cases) and 3323 (the on-host
  default) - none in 516..1837. I listed every `case "` label in that range: 11 carry `when` guards
  (909 assert, 1062 act, 1092 eval, 1111 await, 1218 inspect, 1297 bench, 1425 trip, 1515 layout, 1675
  type, 1735 mark, 1783 dismiss), and an unheld guard leaves the switch and reaches 1839 `var host =
  _editorSurface;` / 1842 `return DebugError("the surface is not up yet");` and then 1861
  RunOnHostThread / 1878 `done.Wait(TimeSpan.FromSeconds(3))`. The consequence claim is where it
  overstates: I read 3323-3344 and the on-host default already answers `$"no route '{request.Route}'
  accepted this request. Either there is no such route, or there is and its required arguments were
  missing or rejected..."` with the given arguments listed. So in the ordinary case (surface up, host
  answering) an argument-rejected pool route gets an ACCURATE diagnosis today, just after a needless
  host crossing. The claimed `{"error":"the host thread did not answer in time"}` requires the host
  thread to be genuinely blocked, and the file's own measured note at 1846-1850 says the common
  blocker does not do that: "a VBA modal PUMPS messages, so marshaled work still runs and no timeout
  ever comes (measured 2026-08-06 ...)".
- **Correction applied:** Two of the three claimed harms do not occur in the ordinary case; the on-host default at 3323
  already names the argument possibility. What is real and unconditional: a route that answers WITHOUT
  the host thread by design (assert, dismiss, guard) loses that property the moment its guard does not
  hold, and with `_editorSurface` null it answers "the surface is not up yet" (line 1842) instead of
  naming the argument. The proposed fix as written breaks a live behaviour: `layout` appears twice,
  guarded at 1515 (`when ... resetFlag != "0"`) and plain at 1549, so the `when` clauses are not all
  argument validation - at least one pair is overload dispatch, and a pre-switch
  route-to-required-arguments table would swallow plain `layout`. The minimal correct change is a
  pool-side default (or a pre-switch guard limited to the host-free routes) that returns the same
  message the on-host default already composes, leaving the layout pair alone.

##### `host-marshal-written-twice` The host-thread marshal-and-wait scaffold is written twice, and the second copy silently drops the counter and the dialog attribution the first one performs

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:963`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** Lines 963-987 inside `case "journal"`: `using var ready = new ManualResetEventSlim(false);
  journalHost.RunOnHostThread(() => { try { sessionState = AnswerDebugRequestOnHost(new
  DebugServer.DebugRequest("state", request.Query, string.Empty)); } catch (Exception ex) {
  sessionState = $"{{\"error\":\"{ex.GetType().Name}\"}}"; } finally { ready.Set(); } }); if
  (!ready.Wait(TimeSpan.FromSeconds(3))) { sessionState = null; }`. Lines 1859-1878, the main
  dispatch: `using var done = new ManualResetEventSlim(false); ... host.RunOnHostThread(() => { try {
  answer = AnswerDebugRequestOnHost(request); } catch (Exception ex) { answer =
  System.Text.Json.JsonSerializer.Serialize(new DebugErrorReply($"{ex.GetType().Name}: {ex.Message}"),
  DebugJsonContext.Default.DebugErrorReply); } finally { done.Set(); } }); var answered =
  done.Wait(TimeSpan.FromSeconds(3));`. Same event, same callback shape, same hardcoded three-second
  deadline. The main copy then does two things the journal copy does not: line 1882
  `PerfCounters.Marshal(Environment.TickCount64 - marshalStarted);` and line 1894
  `RememberRaisedDialogs(standingBefore, keep: ...)`.
- **Why:** Two consequences, both live. The journal's host crossing is invisible to the `perf` and `stats`
  routes even though the comment at 1880 says "Every marshaled request doubles as a probe of the host
  thread's responsiveness" - journal requests are the exception nobody wrote down. And the journal
  copy shapes its error by hand as `{"error":"TypeName"}` rather than through DebugErrorReply, so one
  route on this door emits an error object built by string concatenation while every other one goes
  through the serializer. A third caller added later copies whichever of the two it lands on.
- **Change:** One private helper - `(bool Answered, string? Answer) OnHost(DebugServer.DebugRequest request)` -
  holding the event, the callback, the catch, the deadline and the PerfCounters.Marshal sample. The
  main dispatch keeps the dialog attribution and the blocked-reply path around it; journal calls the
  helper with its synthetic `state` request.
- **Size:** about 20 lines removed from the journal case, and the perf/error divergence goes with them
- **Adversary:** Read both. src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:965-987 in `case "journal"` is `using
  var ready = new ManualResetEventSlim(false); journalHost.RunOnHostThread(() => { try { sessionState
  = AnswerDebugRequestOnHost(new DebugServer.DebugRequest("state", request.Query, string.Empty)); }
  catch (Exception ex) { sessionState = $"{{\"error\":\"{ex.GetType().Name}\"}}"; } finally {
  ready.Set(); } }); if (!ready.Wait(TimeSpan.FromSeconds(3))) { sessionState = null; }`, and
  1858-1882 is the same event, callback shape and hardcoded 3s deadline. The main copy alone carries
  1882 `PerfCounters.Marshal(Environment.TickCount64 - marshalStarted);` under the comment at
  1880-1881 ("Every marshaled request doubles as a probe of the host thread's responsiveness"), so the
  journal's crossing really is invisible to perf and stats. `grep -n RunOnHostThread` over the file
  shows these are the only two marshal-and-wait sites (216, 712, 732, 741, 1643 are fire-and-forget).
- **Correction applied:** The dropped dialog attribution is not a defect: the journal marshals a synthetic read-only `state`
  request, which raises no dialog, so RememberRaisedDialogs would have nothing to attribute. The
  hand-built error is not a shape divergence either - DebugErrorReply is `[property:
  JsonPropertyName("error")] string Error` at src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:729-730,
  so `{"error":"TypeName"}` is the same wire shape, and it is embedded as a string in
  DebugJournalReply.State, escaped by the outer serializer. The one real divergence is the missing
  PerfCounters.Marshal sample.

##### `host-error-reply-boilerplate` The on-host switch has no DebugError equivalent, so twenty-one error paths spell out the serializer call the pool side does in one line

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2747`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** Line 33 defines the pool-side helper: `private static DebugServer.DebugReply DebugError(string
  error) => DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(new
  DebugErrorReply(error), DebugJsonContext.Default.DebugErrorReply));` and it is used 25 times.
  `AnswerDebugRequestOnHost` at 2289 returns `string` rather than `DebugReply`, so it cannot use it:
  `awk 'NR>=2289 && /new DebugErrorReply\(/'` counts 21 sites, each written out in full.
  Representative, lines 2747-2751: `return System.Text.Json.JsonSerializer.Serialize(new
  DebugErrorReply($"kind '{kindText}' is not one of 1/module/standard, " + "2/class, 3/form"),
  DebugJsonContext.Default.DebugErrorReply);`. The same shape recurs at 2760, 2785, 2811, 2828, 2847,
  2855, 2871, 2879, 3104, 3194, 3289, 3338 and eight more. The success reply is duplicated the same
  way: `new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply` appears verbatim
  at 2916, 2928, 2945, 3261, 3315 and 3321.
- **Why:** Three lines of ceremony per error path is what makes this switch 1,058 lines, and the ceremony is
  where the divergence hides: the `component` case at 2879 wraps its message as `$"{componentAction}
  failed: {ex.Message.Trim()}"` while the outer catch at 1869 uses `$"{ex.GetType().Name}:
  {ex.Message}"` and the journal at 975 uses raw string concatenation. A client parsing errors sees
  three shapes from one door because nobody could see all three at once.
- **Change:** Add `private static string HostError(string error)` and `private static string HostOk(int command =
  0)` beside DebugError, both returning the serialized string the on-host switch already returns.
  Twenty-seven call sites collapse to one line each.
- **Size:** about 45 lines removed (21 error sites x 2 lines, 6 ok sites x 1)
- **Adversary:** Counts check out. src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:33-35 defines `DebugError`, and
  `AnswerDebugRequestOnHost` at 2288 returns `string`, so it cannot use it. `awk 'NR>=2289 && /new
  DebugErrorReply\(/'` gives exactly 21 sites: 2499, 2670, 2686, 2699, 2748, 2761, 2786, 2812, 2829,
  2848, 2856, 2872, 2880, 2918, 2933, 3105, 3194, 3263, 3278, 3289, 3339. `new DebugCommandReply(true,
  0)` appears 6 times: 2916, 2928, 2945, 3261, 3315, 3321. I read 2745-2752 and it is the three-line
  form quoted.
- **Correction applied:** The finding's error-line numbers are each one low (it cited the `return System.Text.Json...` line,
  not the `new DebugErrorReply(` line): 2761, 2786, 2812, 2829, 2848, 2856, 2872, 2880, 3105, 3194,
  3289, 3339. The 'three shapes from one door' claim is wrong and should be dropped: every one of
  these sites serializes DebugErrorReply, whose only property is `error` (DebugServer.cs:729), and the
  journal's hand-built object matches it. What differs is the human message text (2880
  `$"{componentAction} failed: {ex.Message.Trim()}"` vs 1869 `$"{ex.GetType().Name}: {ex.Message}"`),
  which a client does not parse. The finding stands as duplication only, not as a client-visible
  defect.

##### `page-eval-reply-shaping-x4` Four page-script routes repeat the identical error-or-DebugEvalReply shaping

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1085`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** `ui` at 1055-1059, `act` at 1085-1089, `eval` at 1104-1108 and `type` at 1728-1732 are the same five
  lines with the local renamed: `return act.Error is { } actError ? DebugError(actError) :
  DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(new
  DebugEvalReply(act.Answered, act.ErrorCode, act.Result, Unwrap(act.Result)),
  DebugJsonContext.Default.DebugEvalReply));`. All four consume the same tuple returned by
  RunPageScript, declared at line 68 as `(bool Answered, int ErrorCode, string Result, string?
  Error)`. `grep -n "new DebugEvalReply("` returns exactly those four lines.
- **Why:** Four copies of the decision "an Error means the transport failed, anything else is an answer even
  when the page said no". A fifth page-driving route is written by copying one of them, and if the
  Unwrap step or the error/answer split ever needs to change, it changes in four places or in three.
- **Change:** `private static DebugServer.DebugReply PageReply((bool Answered, int ErrorCode, string Result,
  string? Error) run)` next to RunPageScript. Each of the four returns become `return
  PageReply(act);`.
- **Size:** about 12 lines removed
- **Adversary:** `grep -n "new DebugEvalReply("` over src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs returns
  exactly 1058, 1088, 1107 and 1731, and I read all four blocks. Each is the identical `X.Error is { }
  err ? DebugError(err) : DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(new
  DebugEvalReply(X.Answered, X.ErrorCode, X.Result, Unwrap(X.Result)),
  DebugJsonContext.Default.DebugEvalReply))` over the tuple RunPageScript declares at line 69 as
  `(bool Answered, int ErrorCode, string Result, string? Error)`. The only per-route variation (the
  script text, the surface argument, the wait budget) is above the shaping, not inside it, so a single
  PageReply helper is a clean lift.

##### `workspace-wait-loop-twice` The wait-for-the-page-to-come-back loop is copied verbatim between the layout-reset and reload routes

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1532`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** Lines 1532-1543 (`layout?reset=1`): `var back = Environment.TickCount64; var restored = false; while
  (Environment.TickCount64 - back < WaitMilliseconds(request, 20000)) { Thread.Sleep(150); var probe =
  RunPageScript("!!(window.xlideBridge && window.xlideBridge.workspace)", null, 1500); if (probe.Error
  is null && probe.Result.Trim() == "true") { restored = true; break; } }`. Lines 1580-1589 (`reload`)
  are the same loop with `startedAt`/`ready`/`reloadBudget` in place of the three names: same 150ms
  tick, same probe script, same 1500ms budget, same success test. Both routes reach it after a
  `RunPageScript` that reloads the page (1530 and 1574).
- **Why:** Two copies of the definition of "the page is back", and they are already drifting in shape: the
  reset copy re-evaluates `WaitMilliseconds(request, 20000)` on every iteration of the loop condition,
  the reload copy hoists it to a local at 1578. Any change to what readiness means - a different
  bridge property, a different tick - has to be made twice, in two routes a probe uses interchangeably
  for cleanup.
- **Change:** `private bool WaitForWorkspace(int budgetMs)` holding the loop and the probe script; both routes
  call it with their own budget.
- **Size:** about 12 lines removed
- **Adversary:** src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1532-1543 and 1579-1589 are the same loop: same
  150ms Thread.Sleep, same probe `RunPageScript("!!(window.xlideBridge &&
  window.xlideBridge.workspace)", null, 1500)`, same `probe.Error is null && probe.Result.Trim() ==
  "true"` test, same 20000ms default budget. `grep -n "xlideBridge && window.xlideBridge.workspace"`
  returns 1537 and 1583 in this file (plus 457, a different script). The shape drift is real too: 1534
  re-evaluates `WaitMilliseconds(request, 20000)` in the loop condition while 1578 hoists it to
  `reloadBudget`.
- **Correction applied:** The re-evaluation at 1534 is a shape difference, not a cost: WaitMilliseconds is a dictionary lookup
  and an int.TryParse, next to a 150ms sleep and a cross-process script call. Do not sell the
  extraction on that.

##### `bench-trip-quantile-shaping` The percentile-and-reply shaping is written twice in this file and a third time in EngineCounters; the trip route is a switch with one live case

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1501`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** Lines 1411-1422 (`bench`) and 1501-1512 (`trip`) build DebugBenchReply the same way: sort, then
  `ordered[0]`, `ordered[ordered.Length / 2]`, `ordered[Math.Min(ordered.Length - 1,
  (int)(ordered.Length * 0.95))]`, `ordered[^1]`. The same p95 expression appears a third time at
  src/Xlide.Vbe.Shim/Diagnostics/EngineCounters.cs:123 over a long[]. Inside `trip`, the switch at
  1472 has exactly one case - `case "pagecall":` at 1474 - plus a default at 1492 that explains why
  there are no others; the comment at 1452-1457 records that the only other scenario, `hostcall`, was
  removed as unmeasurable from inside a route body.
- **Why:** Three copies of one quantile convention means the door can report p95 by two different definitions
  without anyone noticing (the harness has its own copies at tools/harness/xlide-api.mjs:830 and :871,
  so the count is really five). And the `trip` route is 88 lines of switch scaffolding around one
  twelve-line scenario that the comment says will not be joined by another, because the constraint
  that removed `hostcall` applies to every candidate.
- **Change:** One `private static DebugBenchReply Quantiles(string what, List<double> samples, string detail)`
  used by both routes. Reduce the `trip` switch to `if (tripWhat is not "pagecall") { return
  DebugError(...); }` and keep the explanatory message.
- **Size:** about 14 lines removed here; the EngineCounters copy can stay, it is over a different element type
- **Adversary:** Read both. src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1411-1422 (bench, case at 1297) and
  1500-1512 (trip, case at 1425) build DebugBenchReply with the identical `ordered[0]`,
  `ordered[ordered.Length / 2]`, `ordered[Math.Min(ordered.Length - 1, (int)(ordered.Length *
  0.95))]`, `ordered[^1]`. The third copy is real:
  src/Xlide.Vbe.Shim/Diagnostics/EngineCounters.cs:122-123 `sorted[sorted.Length / 2]` /
  `sorted[Math.Min(sorted.Length - 1, (int)(sorted.Length * 0.95))]`. The harness copies are at
  tools/harness/xlide-api.mjs:830 and :871, both `ordered[Math.min(ordered.length - 1,
  Math.floor(ordered.length * 0.95))]`. The trip sub-switch at 1472 does have one case (1474 `case
  "pagecall"`) and a default at 1492.
- **Correction applied:** "88 lines of switch scaffolding" is wrong and would mislead whoever does the work. The trip case
  runs 1425-1513; of that, 1427-1457 is a 31-line comment recording two measured constraints (the
  host-thread pump deadlock, and why hostcall was removed), 1492-1498 is the default's 7-line message,
  1500-1512 is the quantile shaping. The actual switch scaffolding is about 6 lines. Collapsing it to
  an `if` saves those 6 lines and nothing else; the value in this finding is the quantile duplication,
  which is 2 copies here plus EngineCounters plus 2 in the harness. PerfCounters.cs has no copy - it
  exposes raw rings and mentions quantiles only in a comment at line 131.

##### `orphan-doc-and-dead-retried-field` AnswerBlockedRequest's doc comment is stranded above EvaluateClaim, and it documents a retry that the code does not perform

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2070`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity medium
- **Evidence:** Lines 2070-2084 are a complete `<summary>` beginning "What to say - and do - when the host thread
  did not answer" and ending "A dialog that was NOT standing when this request began was raised by
  this request, and answering it is undoing our own mess, so it is dismissed and the request retried
  once." It is immediately followed at 2085 by a second `<summary>` ("Whether a named claim holds
  right now...") and then by `private (bool Held, string Saw) EvaluateClaim(...)` at 2091 - two
  summary elements on one member. The method the first summary describes, `AnswerBlockedRequest`, is
  at 2239 and lines 2237-2238 carry no comment at all. The retry it promises does not happen: 2272 is
  `var completed = pressed is not null && done.Wait(TimeSpan.FromSeconds(3));`, which waits for the
  work already queued, and both DebugBlockedReply constructions pass `Retried: false` (2259 and 2285).
  `grep -rn Retried` over src/, tools/harness and the two api docs finds only those two sites and the
  record field at src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:744 - nothing ever sets it true and
  nothing reads it.
- **Why:** The blocked-request policy is the hardest thing in this file to reason about, and its only
  explanation is attached to the wrong method, where a reader looking at AnswerBlockedRequest will not
  find it and a reader of EvaluateClaim gets a paragraph about modal dialogs. The explanation is also
  now false, which is worse than absent: it says the request is retried, and the wire field `retried`
  that a client would use to check that is a constant false.
- **Change:** Move lines 2070-2084 to sit above line 2239, correct "retried once" to what the code does (waits out
  the work already queued), and either set Retried honestly or drop it from DebugBlockedReply.
- **Size:** about 4 lines moved, one dead wire field removed
- **Adversary:** Read src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2070-2091: a complete `<summary>` beginning
  "What to say - and do - when the host thread did not answer" is immediately followed at 2085 by a
  second `<summary>` ("Whether a named claim holds right now...") and then `private (bool Held, string
  Saw) EvaluateClaim(...)` at 2091 - two summary elements on one member. AnswerBlockedRequest is at
  2239 with no comment above it. The retry claim at 2080 ("it is dismissed and the request retried
  once") is false: 2272 is `var completed = pressed is not null &&
  done.Wait(TimeSpan.FromSeconds(3));` under its own comment at 2270-2271 saying "the work this
  request asked for was queued before the dialog appeared, so it may complete on its own" - a wait,
  not a retry. Both DebugBlockedReply constructions pass `Retried: false` (2259, 2285). A repo-wide
  grep over src/, tools/, docs/, ui/editor/src and engine/src for `Retried|retried` finds only those
  two, the record field at src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:744, and unrelated prose;
  nothing sets it true and nothing reads it. It is absent from docs/debug-api.md, so `retried` is on
  the wire, always false, and undocumented.

##### `two-boolean-flag-conventions` Query flags are parsed by two incompatible rules in the same switch, so the same argument value means opposite things on different routes

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:1014`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity medium
- **Evidence:** Six sites treat any value other than the literal "0" as true: 1014
  `request.Query.TryGetValue("reset", out var perfReset) && perfReset != "0"`, 1231 `rulesFlag !=
  "0"`, 1515 `resetFlag != "0"`, 2469 `engineText != "0"`, 2561 `wantsText != "0"`, 3273 `liveFlag !=
  "0"`. Two sites require an affirmative word: 1764 `_guardEverything = wanted is "1" or "true" or
  "yes" or "on";` and 2984 `asked is "1" or "true" or "yes" or "on"`.
- **Why:** `perf?reset=false` resets the analyzer counters; `guard?on=false` correctly leaves the guard off.
  `problems?rules=off` returns the rules. A caller who learns the convention from one route gets the
  opposite behaviour from the next, and the failure is silent - the counters are already cleared by
  the time anyone notices. The two conventions are 750 lines apart in one switch, which is why neither
  author saw the other.
- **Change:** One `private static bool Flag(DebugServer.DebugRequest request, string name, bool fallback = false)`
  used by all eight sites, accepting 1/true/yes/on as true and 0/false/no/off as false and answering
  `fallback` for anything else.
- **Size:** 8 sites to one helper, about 6 lines; the risk is that any harness call relying on `=false` meaning true changes behaviour, which is the point
- **Adversary:** All eight sites are as cited (grep over src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs): `!= "0"`
  at 1014 (perf reset), 1231 (problems rules), 1515 (layout reset), 2469 (native text), 2561 (engine
  text), 3273 (module live); `is "1" or "true" or "yes" or "on"` at 1764 (guard on) and 2984
  (settings). The concrete failure is real: `perf?reset=false` satisfies `perfReset != "0"` and clears
  the engine counters. And the cross-contamination path is documented - docs/debug-api.md:81 teaches
  `guard | POST | on=true|false`, so a caller does meet `true|false` on this door before meeting a
  `=1` toggle.
- **Correction applied:** The two conventions are not arbitrary: they carry different semantics. The six `!= "0"` sites are
  per-request toggles the docs only ever spell as `=1` (docs/debug-api.md lines 50, 56, 65, 75, 76);
  the two word-list sites are SETTERS of persisted state where `false` must mean false -
  `guard?on=false` has to turn the guard off, and the settings route at 2982-2985 already has a local
  `bool Flag(string name, bool current)` doing exactly the tri-state the finding proposes. So the fix
  is to lift the existing local Flag out and adopt it, not to invent one, and it must keep
  absent-means-unchanged for guard and settings while absent-means-false for the toggles.

##### `component-enumeration-duplicated` The projects and project routes each walk VBComponents with their own copy of the scratch-module filter

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3058`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity low
- **Evidence:** Lines 3058-3070 in `case "projects"`: `using (var list = project.GetObject("VBComponents")) { var
  total = list?.GetInt32("Count") ?? 0; for (var c = 1; c <= total; c++) { using var component =
  list!.GetItem(c); if (component?.GetString("Name") is { Length: > 0 } name &&
  !IsScratchComponent(name)) { components++; } } }`. Lines 3110-3122 in `case "project"`: `using (var
  components = project.GetObject("VBComponents")) { var count = components?.GetInt32("Count") ?? 0;
  for (var i = 1; i <= count; i++) { ... using var component = components!.GetItem(i); if
  (component?.GetString("Name") is not { Length: > 0 } name || IsScratchComponent(name)) { continue; }
  ...` - the same enumeration, the same wrapper lifetimes, the same filter written in the negative.
- **Why:** The scratch module's exclusion rule is a fixture-correctness property - the comment at 3054-3056
  says a fixture that counts it counts wrong - and it is now stated twice, in two routes a fixture
  uses together. The two loops also differ in their failure handling (3080 logs at Info, 3142 logs at
  Verbose) for the same class of unreadable component, so a component that fails to read is visible in
  one route's log and effectively invisible in the other's.
- **Change:** One `private IEnumerable<ComPtr> LiveComponents(ComPtr project)` (or a callback form, to keep the
  wrapper lifetimes obvious under ComRuntime) holding the count loop, the name read, the scratch
  filter and the per-entry catch. Both routes consume it.
- **Size:** about 12 lines removed; medium risk only because the wrappers must keep going through TakeWrapper/GiveBackWrapper unchanged
- **Adversary:** Read both. src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:3057-3070 (projects) and 3110-3122
  (project) are the same VBComponents count-loop with the same scratch filter written positively then
  negatively, the same `using var component = ...GetItem(i)` lifetimes. The failure handling differs
  as claimed (3080 Log.Info, 3142 Log.Verbose).
- **Correction applied:** Two things are understated. First, the failure-handling difference is not just log level: in
  `projects` the inner component loop has NO try of its own - the try opens at 3047 around the whole
  project entry and `found.Add(new DebugProjectRow(...))` at 3071 sits inside it, so one unreadable
  component makes the ENTIRE project disappear from the reply, where `project` (per-item try at
  3114-3143) skips the component and returns the rest. Second, the walk is written three times, not
  two: src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5919-5926 is the same count-loop plus
  `component?.GetString("Name") is { Length: > 0 } name && !IsScratchComponent(name)` in PRODUCTION
  code building the surface's project tree. That makes the helper more worthwhile and also fixes its
  placement: it cannot live inside the `#if DEBUG` region of AddInSession.DebugApi.cs if the
  production walk is to share it.

##### `log-collapse-marker-misattributes-thread` The duplicate-collapse line in the log claims level info and thread host regardless of what was actually repeated

- **Where:** `src/Xlide.Vbe.Shim/Diagnostics/Log.cs:145`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** The collapse key at line 130 is `string.Concat(level, "|", origin, "|", message)`, so a suppressed
  run carries its own level and origin. The flush at 144-145 hardcodes both:
  `text.Append(CultureInfo.InvariantCulture, $"{DateTime.Now:HH:mm:ss.fff} [info] [host] ... last line
  repeated {_suppressed} more time(s){Environment.NewLine}");`. The line immediately below, 151-152,
  formats real lines with the actual `{level}` and `{origin}`.
- **Why:** The comment at 124-126 states why the thread is on every line: "a whole class of defect is right
  call, wrong thread, and a log that hides the thread hides the defect". A verbose burst from a pool
  thread - which is what the window-event storms this collapsing exists for actually are - is
  summarised by a line saying [host], so the one reading that tells you which thread was busy is the
  one that is wrong about it. Reading a collapsed burst as host-thread work sends a stall
  investigation at the wrong thread.
- **Change:** Carry the suppressed run's level and origin alongside `_lastKey` and format the flush line with
  them.
- **Size:** about 3 lines
- **Adversary:** Read src/Xlide.Vbe.Shim/Diagnostics/Log.cs:127-152. The key at 130 is `string.Concat(level, "|",
  origin, "|", message)`, so a suppressed run is per level and per origin, and the flush at 144-145
  hardcodes `$"{DateTime.Now:HH:mm:ss.fff} [info] [host] ... last line repeated {_suppressed} more
  time(s)"` while the real line two statements later at 151-152 uses `[{level}] [{origin}]`. The
  comment at 124-126 is quoted correctly.
- **Correction applied:** The consequence is much smaller than stated. The collapse only ever suppresses lines identical to
  one already written, so the marker is always preceded in the file by that same line carrying its
  true level and origin - a reader has the right attribution one line up, and nothing machine-reads
  the origin field (no harness script or api doc parses `[host]`/`[tN]`; the log route filters on a
  free-text `match`). Worse for the finding's framing: the storm the collapsing exists for is
  host-thread work, not pool-thread - the comment at 102-106 names a resize drag, and
  AddInSession.cs:1058 puts that path "In the frame's message chain, so a resize re-places the surface
  synchronously", i.e. the host UI thread, so `[host]` is correct in the motivating case. The bug
  bites only for a burst from an engine or pool thread.

### The rest of the shim and the core library

_10 findings, from the `complexity-shim-rest` finder._

##### `locals-watch-reader-twins` WatchReader is a copy of LocalsReader: Connect and Dispose are byte-identical, and each spins up its own CUIAutomation client on the same thread

- **Where:** `src/Xlide.Vbe.Shim/Editor/WatchReader.cs:52`
- **Kind:** complexity / medium effort, claim derived, confidence verified, severity medium
- **Evidence:** Diffed the two files region by region. `diff <(sed -n '236,249p' WatchReader.cs) <(sed -n '303,316p'
  LocalsReader.cs)` returns nothing - the two Dispose bodies are byte-identical:
  
   if (_condition != 0) { Marshal.Release(_condition); _condition = 0; }
   _element?.Dispose(); _element = null;
   _automation?.Dispose(); _automation = null;
  
  `diff` of Connect (WatchReader 52-92 against LocalsReader 68-110) reports exactly two added comment
  lines and, after normalising the log prefix, nothing else - the same
  CoCreateInstance(UiAutomationIds.AutomationClass, 0, Win32.ClassContextInProcessServer,
  UiAutomationIds.Automation), the same ComHandle<IUIAutomation>.Own, the same ElementFromHandle, the
  same CreateTrueCondition. The field blocks match too (WatchReader 25-33, LocalsReader 37-49:
  `_window`, `_automation`, `_element`, `_condition`, `_consecutiveFailures`, `_retryAt`), as do the
  Create factories (41-50 vs 57-66) and the backoff catch (`_retryAt = Environment.TickCount64 +
  5000;` with the first-of-streak log, in both). WatchReader's own comments admit it: "Failure streak
  and backoff, the same manner as the Locals reader's" (line 31), "the same manner as the Locals
  reader" (121), "Backed off rather than stopped, the same manner as the Locals reader" (174). The
  Read loops (WatchReader 95-185, LocalsReader 119-249) run the same FindAll -> GetLength ->
  GetElement -> ControlTypeProperty -> NameProperty walk with the same per-element try/catch and
  `poisoned` counter; they diverge only in which control types are accepted, LocalsReader's `stage`
  string, its Edit/Pane context branch, and ParseRow. Both readers are constructed on one thread:
  GhostReaderThread.Run lines 95 and 103 call LocalsReader.Create then WatchReader.Create inside a
  single CoInitializeEx(MTA), so the process holds two CUIAutomation clients where one would serve
  both windows.
- **Why:** About 74 lines are duplicated with no textual difference beyond a log prefix, and the duplication is
  not inert: this is the code path that died on 2026-08-05 and whose fix (the sized variant
  out-parameter, the separate thread, the per-element try/catch, the 5000 ms backoff) had to be
  understood once and then transcribed. The next correction to the accessibility walk has to be made
  twice, and a reader who fixes only the file they were looking at leaves the other panel on the old
  behaviour. The two automation clients are a second, smaller cost: CoCreateInstance runs twice at
  ghost-reader start and two provider connections exist for the life of the session.
- **Change:** Give the two readers one base, or one class parameterised by what differs. The stable part is
  exactly Create/Connect/Dispose plus a `ForEachListItem(Action<ComHandle<IUIAutomationElement>, int
  controlType, string name>)` that owns the FindAll walk, the per-element try/catch, the poisoned
  count and the failure backoff. LocalsReader keeps its stage strings and its Edit/Pane context rule
  inside the callback; WatchReader keeps ParseRow. Hoist the CUIAutomation instance and the true
  condition into GhostReaderThread, which already owns the apartment both readers run in, and hand
  each reader the shared IUIAutomation rather than letting each create one.
- **Size:** 74 lines byte-identical (Dispose 14, Connect 41, fields 9, Create 10), plus a ~55-line Read loop that differs in three places; one redundant CUIAutomation instance per session
- **Adversary:** Read both files. src/Xlide.Vbe.Shim/Editor/WatchReader.cs:52-92 Connect and
  src/Xlide.Vbe.Shim/Editor/LocalsReader.cs:68-110 Connect are the same body (same
  CoCreateInstance(UiAutomationIds.AutomationClass, 0, Win32.ClassContextInProcessServer,
  UiAutomationIds.Automation), same ComHandle<IUIAutomation>.Own, same ElementFromHandle, same
  CreateTrueCondition); the only textual differences are the log prefix ('watch:' vs 'locals:') and
  two extra comment lines in LocalsReader. I ran diff on WatchReader.cs:236-249 against
  LocalsReader.cs:303-316 and it returned no output, so the two Dispose bodies are byte-identical.
  Field blocks (WatchReader 25-33 / LocalsReader 37-49) and Create factories (41-50 / 57-66) match,
  and both catch blocks set `_retryAt = Environment.TickCount64 + 5000;` with the first-of-streak log
  (WatchReader 172-184, LocalsReader 235-245). Read loops share the FindAll -> GetLength -> GetElement
  -> ControlTypeProperty -> NameProperty walk with the same per-element try/catch and `poisoned`
  counter, diverging only in accepted control types, LocalsReader's `stage` string, its Edit/Pane
  context branch and the row parser. src/Xlide.Vbe.Shim/Editor/GhostReaderThread.cs:95 and :103 do
  construct both inside one CoInitializeEx(MTA) at line 88, so two CUIAutomation clients exist on one
  thread. Not a preference: this is the walk that killed the host on 2026-08-05, and its hardening now
  has to be applied twice.
- **Correction applied:** The second CUIAutomation client is real but costs one extra CoCreateInstance at ghost-reader start
  and one extra provider connection, not anything on a poll tick; it is the weakest half of the
  finding. The duplication count is right (Dispose 14 identical, Connect ~41 differing only in a log
  prefix and two comments, fields 9, Create 10).

##### `execute-script-handler-never-unrooted` WebView2Surface.ForgetScriptHandler is never called, so every debug-api eval permanently roots a handler and its COM callback wrapper

- **Where:** `src/Xlide.Vbe.Shim/WebView/WebView2Surface.cs:143`
- **Kind:** dead-code / small effort, claim derived, confidence verified, severity medium
- **Evidence:** Line 143 declares `internal void ForgetScriptHandler(object handler) =>
  _pendingScripts.Remove(handler);`. `grep -rn "ForgetScriptHandler" --include=*.cs src` returns that
  line and nothing else - no caller anywhere in the shim. The list it drains is filled unconditionally
  in ExecuteScript (line 132): `_pendingScripts.Add(handler);` under the comment "The handler is
  rooted until it fires: nothing else holds it, and a collected one would be a callback into freed
  memory" - but ExecuteScriptCompletedHandler.Invoke (lines 1065-1077) only calls `_completed(...)`
  and returns HResult.Ok; it never un-roots. Dispose (942-998) nulls `_environmentHandler`,
  `_controllerHandler`, `_navigationHandler`, `_messageHandler` and `_acceleratorHandler` and never
  touches `_pendingScripts`. Each entry also carries a CCW: CreateCallback (line 931) does
  `ComRuntime.Wrappers.GetOrCreateComInterfaceForObject(handler, CreateComInterfaceFlags.None)`, so
  the COM identity stays alive as long as the managed handler is reachable. That path does NOT go
  through ComRuntime.TakeWrapper / GiveBackWrapper (ComRuntime.cs 62 and 129), which are the only two
  doors WrappersTaken / WrappersGivenBack / WrappersDisposed count, so `stats` reads balanced while
  the list grows. The whole block is inside `#if DEBUG` (line 84 opens it, 203 closes it; the handler
  class is inside a second `#if DEBUG` at 1056-1079), and the only caller is the debug api's script
  round trip at AddInSession.DebugApi.cs:221 `browser.ExecuteScript(script, (code, json) => ...)`.
- **Why:** The eval route is how the harness observes the page, so a live suite makes hundreds of these calls
  in one session and each one adds a managed object plus a CCW that is never released for the life of
  the surface. It is Debug-only, so it will not ship, but Debug is exactly where the leak instrument
  is supposed to be trustworthy: this growth is invisible to WrappersLive by construction, which means
  a session can be accumulating COM identities while the stats route says the accounting is clean. The
  un-root method that would fix it was written and then never wired up, so the code reads as if the
  lifetime were managed.
- **Change:** Wrap the caller's callback in ExecuteScript so completion removes the entry: capture the handler,
  and in the wrapper invoke `completed(code, json)` then `ForgetScriptHandler(handler)`. Clear
  `_pendingScripts` in Dispose alongside the other handler fields, so a surface torn down with calls
  in flight does not hold them. If instead the intent is that a handler outlives its call, delete
  ForgetScriptHandler and say so at the list, because a public un-root nobody calls reads as a
  lifetime that is managed.
- **Size:** one leaked handler plus one CCW per eval call, unbounded over a session; 1 dead method
- **Adversary:** src/Xlide.Vbe.Shim/WebView/WebView2Surface.cs:143 `internal void ForgetScriptHandler(object handler)
  => _pendingScripts.Remove(handler);` and `git grep ForgetScriptHandler` over the whole repo returns
  only that line plus the list at :141 and the Add at :132. ExecuteScriptCompletedHandler.Invoke
  (WebView2Surface.cs:1065-1077) calls `_completed(...)` and returns HResult.Ok with no un-root.
  Dispose (942-998) nulls _environmentHandler, _controllerHandler, _navigationHandler, _messageHandler
  and _acceleratorHandler and never touches _pendingScripts. The list therefore grows for the life of
  the surface. The block is Debug-only (#if DEBUG at :84, #endif at :203; handler class inside a
  second #if DEBUG at 1056-1079), so nothing ships.
- **Correction applied:** The size is understated and the mechanism is slightly off. Growth is one handler per
  RunPageScriptOnce call, not per `eval` route call:
  src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:221 is the only ExecuteScript caller, but
  RunPageScript's backoff poll loop calls RunPageScriptOnce again on every tick (DebugApi.cs:143-148,
  'polled about a dozen times' per its own comment at :137-139), and ui (:1054), act (:1080), run
  (:1103), wait (:1163), layout (:1555) and a dozen more routes all go through RunPageScript. A live
  suite accumulates thousands of entries, not hundreds. What leaks is the managed handler plus its
  closure (which pins the already-disposed ManualResetEventSlim from DebugApi.cs:213) plus the
  ComWrappers table entry; the native CCW itself is torn down once the browser drops its reference,
  since CreateCallback (:929-940) releases both references it takes. The claim that WrappersLive
  cannot see it stands: ComRuntime.TakeWrapper/GiveBackWrapper are the only counted doors and
  GetOrCreateComInterfaceForObject at :931 is not one of them.

##### `vbecommands-describe-dead` VbeCommands.Describe has no caller anywhere in the repository and was superseded by VbeMenus.Describe and the menus route

- **Where:** `src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:322`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** `grep -rn "VbeCommands.Describe" --include='*' . --exclude-dir=node_modules --exclude-dir=.git`
  returns nothing. The only reference to the name inside the file is its own doc comment on the
  Command class (line 30): "<see cref=\"Describe\"/> writes each menu item's identifier and caption to
  the log". The method itself (322-361) walks CommandBars and logs `Log.Info($"
  {control.GetInt32(\"Id\")} '{control.GetString(\"Caption\")}'")` for every control on every bar, and
  its own summary says "It is not called during normal operation." The job it describes is now done by
  a route: VbeMenus.Describe (VbeMenus.cs:401) returns `(int Index, int Id, string Caption, bool
  Popup, bool Enabled, bool Suppressed)[]`, is called from AddInSession.DebugApi.cs:2652, and its
  comment states the succession outright - "until this existed there was no way to ask the running
  editor for a menu's ids at all - they were measured once by hand and written into a comment, which
  is how a table of numbers goes quietly out of date." docs/debug-api.md line 86 documents the `menus`
  route that serves it.
- **Why:** Forty lines of live COM enumeration that nothing can reach, presented by the doc comment on Command
  as the way to re-establish the command identifiers against a disagreeing host. A maintainer who
  follows that instruction finds a method with no call site and has to work out for themselves that
  the answer now comes from the menus route, which reports the ids AND whether each one is suppressed
  - strictly more than the log dump gives.
- **Change:** Delete Describe (322-361) and repoint the Command class doc comment at the `menus` debug route and
  VbeMenus.Describe, which is the surviving instrument and reports suppression as well as identity.
- **Size:** 40 lines removable
- **Adversary:** src/Xlide.Vbe.Shim/Editor/VbeCommands.cs:322-361 `public static void Describe(DispatchObject
  editor)`. A repo-wide `git grep Describe` (excluding node_modules) returns no call site for it: the
  only other hit in that file is its own doc comment at :29 on the Command class, and every other
  Describe in the tree belongs to a different type (VbeMenus.Describe,
  ModuleText.LostCharacter.Describe, WindowEvent.Describe, FolderLock.Describe, VBA fixtures).
  VbeMenus.Describe (src/Xlide.Vbe.Shim/Editor/VbeMenus.cs:401) is live from
  src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs:2652 serving the `menus` route, documented at
  docs/debug-api.md:86. So the method is unreachable without editing code, while its doc comment tells
  a maintainer to 'Re-run it if a host build ever disagrees'.
- **Correction applied:** The 'superseded, strictly more' framing is wrong in one respect that changes the fix.
  VbeMenus.Describe reaches controls only from the MENU BAR: ControlsAt (VbeMenus.cs:480-498) starts
  at FindMenuBar (:505-522, selecting the bar whose Type == MenuBarType) and walks a position chain.
  VbeCommands.Describe enumerates EVERY CommandBar and its controls, which is where VbeCommands.Find
  looks (PreferredBars = ['Debug','Edit'], VbeCommands.cs:23). Deleting it therefore loses the
  whole-CommandBars id dump the menus route cannot produce. Either delete it and repoint the :29
  comment at the menus route while accepting that gap, or add a bar selector as a NEW field/route on
  the menus door and then delete.

##### `set-theme-message-dead` SetThemeMessage is declared and registered on the AOT serializer context but the shim never constructs one

- **Where:** `src/Xlide.Vbe.Shim/Editor/EditorMessages.cs:43`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** `grep -rn "SetThemeMessage|\"setTheme\"" --include='*.cs' src` returns exactly two hits: the record
  at EditorMessages.cs:43 and `[JsonSerializable(typeof(SetThemeMessage))]` at EditorMessages.cs:522.
  There is no `new SetThemeMessage(...)` and no C# site that emits the string "setTheme", so no
  EditorSurface method sends one - contrast every live message shape, each of which has a Post/Send
  call site (for example ShowCompletions at EditorSurface.cs:249 `Post(JsonSerializer.Serialize(new
  CompletionResultMessage("completionResult", requestId, items),
  EditorMessageContext.Default.CompletionResultMessage))`). The page half is real and still live:
  ui/editor/src/bridge.ts:62 declares `| { type: "setTheme"; theme: XlideTheme }` and
  bridge.ts:1162-1164 handles it by calling `monaco.editor.setTheme(...)`, and ui/editor/README.md:50
  documents it as a host message.
- **Why:** The message contract says the host can pin the page's theme and it cannot: the record exists, the
  page listens, and no code path produces the message. Anyone reading EditorMessages.cs to find out
  what the shim can tell the page gets a wrong answer, and the README documents a host-to-page message
  the host has no way to send. It also costs a source-generated serializer for a type never
  serialised, which under NativeAOT is generated metadata for nothing.
- **Change:** Decide the direction and make the file say it. Either drop the record and its JsonSerializable line
  (and the README row, since nothing on the host side drives it), or add the EditorSurface method that
  sends it and the settings path that calls that method. Do not leave the shape declared with no
  producer.
- **Size:** 6-line record plus 1 serializer registration, or one missing send method
- **Adversary:** `git grep -rn 'setTheme|SetThemeMessage'` over the whole repo returns exactly six hits: the record
  at src/Xlide.Vbe.Shim/Editor/EditorMessages.cs:43-45, its registration at EditorMessages.cs:522,
  ui/editor/README.md:50, and three in ui/editor/src/bridge.ts (:62 the union member, :1162-1164 the
  handler calling monaco.editor.setTheme, :1353 inside applyOsTheme). No C# site constructs one and no
  C# site emits the string, so the host has no way to send it. The page half is live.
- **Correction applied:** Worth adding to the fix decision: nothing else on the host carries a theme either - `git grep theme`
  over EditorMessages.cs and EditorSurface.cs returns only the SetThemeMessage record and its
  JsonPropertyName - and the page sets its own theme from the OS (bridge.ts:1349 applyOsTheme). So
  ui/editor/README.md:50 documents pinning the theme against prefers-color-scheme, a capability the
  host lacks entirely, not merely an unused message.

##### `debug-capture-duplicate-interop` DebugCapture re-declares Rect, GetWindowRect, SelectObject and DeleteObject that Interop/Win32 already declares

- **Where:** `src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:489`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** DebugServer.cs 489-495 declares `[StructLayout(LayoutKind.Sequential)] private struct Rect { public
  int Left; public int Top; public int Right; public int Bottom; }`. Win32.cs 6-13 already declares
  `[StructLayout(LayoutKind.Sequential)] internal struct Rect` with those four int fields in that
  order. The imports repeat too: DebugServer.cs:465 `private static partial bool GetWindowRect(nint
  window, out Rect rect)` against Win32.Events.cs:88 `public static partial bool GetWindowRect(nint
  window, Rect* rect)`; DebugServer.cs:479 `private static partial nint SelectObject(nint dc, nint
  gdiObject)` against Win32.cs:326 `public static partial nint SelectObject(nint deviceContext, nint
  gdiObject)`; DebugServer.cs:483 `private static partial bool DeleteObject(nint gdiObject)` against
  Win32.cs:323 `public static partial bool DeleteObject(nint gdiObject)`. The class comment (lines
  450-453) justifies raw GDI - "the shim is ahead-of-time compiled and carries no drawing library" -
  which explains reaching for GDI, not declaring a second copy of imports the same assembly already
  has. GetDC, ReleaseDC, PrintWindow, CreateCompatibleDC, DeleteDC and CreateDIBSection are genuinely
  new and belong here.
- **Why:** Two Rect structs in one assembly means a call site can pick the wrong one and only find out at the
  marshalling boundary, and the three duplicated imports generate a second set of LibraryImport
  marshalling stubs for functions already bound. It also breaks the rule the Interop folder exists to
  enforce - one declaration per native entry point - so the next person adding a GDI call has two
  places to look and no way to tell which is canonical.
- **Change:** Have DebugCapture use Interop.Win32.Rect, Win32.GetWindowRect, Win32.SelectObject and
  Win32.DeleteObject, and keep only the six imports that are genuinely new (GetDC, ReleaseDC,
  PrintWindow, CreateCompatibleDC, DeleteDC, CreateDIBSection). GetWindowRect differs in shape
  (`Rect*` against `out Rect`); the pointer form works from a local with `&rect`, which is how
  CodePaneTracker.ReadBounds already calls GetClientRect.
- **Size:** 1 duplicate struct plus 3 duplicate P/Invokes, about 20 lines
- **Adversary:** The duplication is exactly as described. src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs:488-495
  declares a private `Rect` with Left/Top/Right/Bottom; src/Xlide.Vbe.Shim/Interop/Win32.cs:6-13
  already declares `internal struct Rect` with the same four ints in the same order.
  DebugServer.cs:463-465 GetWindowRect duplicates Win32.Events.cs:88 (`public static partial bool
  GetWindowRect(nint window, Rect* rect)`); DebugServer.cs:478-479 SelectObject duplicates
  Win32.cs:326; DebugServer.cs:481-483 DeleteObject duplicates Win32.cs:323. GetDC, ReleaseDC,
  PrintWindow, CreateCompatibleDC, DeleteDC and CreateDIBSection are genuinely new, as claimed.
- **Correction applied:** Both stated harms are wrong, which drops this to a pure tidiness item. (1) No call site can pick the
  wrong Rect: DebugCapture.Rect is a PRIVATE NESTED struct, invisible outside the class, and its
  layout is identical to Interop.Win32.Rect anyway. (2) There is no shipped stub cost: DebugServer.cs
  is wholly inside `#if DEBUG` (line 1 `#if DEBUG`, line 1256 `#endif`), so none of these imports
  exist in Release, and an unreferenced LibraryImport partial is not rooted by ILC in any case. What
  remains is one duplicate struct and three duplicate imports in a Debug-only file, about 20 lines.

##### `window-text-and-class-readers-duplicated` DialogWatch and CodePaneTracker each carry a private copy of the same two window-text helpers

- **Where:** `src/Xlide.Vbe.Shim/Diagnostics/DialogWatch.cs:223`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** DialogWatch.cs 223-229:
  
   private static string ReadClass(nint window)
   { const int capacity = 64; var buffer = stackalloc char[capacity];
   var length = Win32.GetClassName(window, buffer, capacity);
   return length <= 0 ? string.Empty : new string(buffer, 0, length); }
  
  CodePaneTracker.cs 910-916 has the same body under the name ReadClassName, same capacity 64, same
  guard. DialogWatch.ReadText (215-221) and CodePaneTracker.ReadWindowText (918-924) are likewise the
  same body with capacity 256 over Win32.GetWindowText. Both are private, so neither can see the
  other.
- **Why:** Two copies of the buffer size for a class name and two copies of the truncation rule. A 64-char
  class buffer is fine for MDIClient and VbaWindow, but the decision is now recorded in two places and
  a caller who needs a longer one will lengthen whichever file they happen to be in. Both are also the
  natural home for the next window-inspection helper, so the split invites a third copy.
- **Change:** Move ReadClassName and ReadWindowText onto Interop.Win32 as internal statics next to the
  GetClassName and GetWindowText imports they wrap, and have both callers use them. They take no state
  and belong with the imports, not with either consumer.
- **Size:** 14 duplicated lines across two files
- **Adversary:** src/Xlide.Vbe.Shim/Diagnostics/DialogWatch.cs:215-221 ReadText (capacity 256, Win32.GetWindowText,
  `length <= 0 ? string.Empty : new string(buffer, 0, length)`) and :223-229 ReadClass (capacity 64,
  Win32.GetClassName, same guard) are the same bodies as
  src/Xlide.Vbe.Shim/Editor/CodePaneTracker.cs:918-924 ReadWindowText and :910-916 ReadClassName. Both
  sets are private, so neither can see the other, and the imports they wrap sit together at
  src/Xlide.Vbe.Shim/Interop/Win32.Events.cs:51-52 and :60-61.
- **Correction applied:** Three sites, not two, and one of the cited two never ships.
  src/Xlide.Vbe.Shim/Editor/HostChrome.cs:200-206 ReadCaption is a third copy of the window-text read
  with capacity 512 that returns null rather than string.Empty. DialogWatch.cs is entirely inside `#if
  DEBUG` (line 1), so it is a Debug-only copy. A consolidation onto Interop.Win32 therefore has to
  reconcile three capacities (64/256/512) and two empty conventions (string.Empty vs null), which is
  more than a lift-and-move.

##### `editor-frame-located-four-ways` CodePaneTracker finds the editor frame by two different mechanisms and repeats one of them three times

- **Where:** `src/Xlide.Vbe.Shim/Editor/CodePaneTracker.cs:135`
- **Kind:** complexity / medium effort, claim derived, confidence verified, severity low
- **Evidence:** Three copies of the same walk. MainWindow (135-150): `var ours = Win32.GetCurrentProcessId(); nint
  frame = 0; while ((frame = Win32.FindWindowEx(0, frame, FrameClass, null)) != 0) {
  Win32.GetWindowThreadProcessId(frame, out var owner); if (owner == ours) { return frame; } }`.
  VisiblePanes (625-640) opens with the identical four lines and then enumerates children instead of
  returning. FindByCaption (681-696) opens with the identical four lines, guarded by `_captionFound ==
  0`, and enumerates children. A fourth, different implementation answers the same question:
  FindFrameByEnumeration (884-908) walks GetTopWindow/GetWindow(GwHwndNext), filters on `owner !=
  self` and returns the first window whose class is FrameClass, and it is what both FindEditorFrame's
  fallback (878) and the public FindFrame (882) return. Both entry points are live from outside:
  AddInSession.cs:522 and 1186 call MainWindow, AddInSession.cs:1315 and XlideAddIn.cs:161 call
  FindFrame.
- **Why:** Two public members of one class answer "where is the editor frame" by different native mechanisms,
  and a caller has no way to tell which to use - FindWindowEx enumerates by class in creation order,
  GetTopWindow walks z-order, so with two frames they can disagree and nothing in either doc comment
  mentions the other exists. The three-times-copied prologue means the process-ownership filter (whose
  comment at 131-134 explains it exists because "a second host running the same editor has one too")
  is written out three times and can be fixed in one.
- **Change:** Keep one frame walk. Make MainWindow the single locator, have FindFrame and FindEditorFrame's
  fallback call it, and give VisiblePanes and FindByCaption a shared `ForEachOwnedFrame(Action<nint>)`
  that holds the FindWindowEx loop and the process filter once. If the two mechanisms genuinely differ
  in what they should return, the doc comments have to say which one a caller wants and why.
- **Size:** 3 copies of an 8-line prologue plus a redundant 24-line second locator
- **Adversary:** All four sites read as described in src/Xlide.Vbe.Shim/Editor/CodePaneTracker.cs. MainWindow
  (135-150) is `var ours = Win32.GetCurrentProcessId(); nint frame = 0; while ((frame =
  Win32.FindWindowEx(0, frame, FrameClass, null)) != 0) { Win32.GetWindowThreadProcessId(frame, out
  var owner); if (owner == ours) { return frame; } }`. VisiblePanes (625-640) and FindByCaption
  (681-696) each open with that identical prologue and then EnumChildWindows instead of returning.
  FindFrameByEnumeration (884-908) is a separate implementation over
  GetTopWindow(0)/GetWindow(GwHwndNext) filtered on `owner != self` and `ReadClassName(window) ==
  FrameClass`, and it is what both FindEditorFrame's fallback (878) and the public FindFrame (882)
  return. Call sites confirmed by grep: AddInSession.cs:522 and :1186 call MainWindow,
  AddInSession.cs:1315 and XlideAddIn.cs:161 call FindFrame.
- **Correction applied:** The divergence rationale is wrong and should not be used to justify the work. FindWindowEx with a
  null parent enumerates the desktop's children in Z ORDER, the same order
  GetTopWindow(0)+GetWindow(GW_HWNDNEXT) walks, and both filter to this process and class
  wndclass_desked_gsk (CodePaneTracker.cs:49), so MainWindow() and FindFrameByEnumeration() return the
  same handle even with two frames. The defect is redundancy only: a 24-line second implementation of
  an identical function plus three copies of the 8-line ownership prologue whose reason is written
  once at 131-134. Nothing can currently disagree, so severity is maintenance, not correctness.

##### `hresult-ole-era-constants-unused` Ten of HResult's eighteen constants have no reference, including four OLE embedding codes from a design the shim no longer has

- **Where:** `src/Xlide.Vbe.Shim/Com/HResult.cs:12`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** Counted `HResult.<name>` occurrences across all of src: Ok 19, InvalidArg 5, Fail 3, False 1,
  ClassNotAvailable 1, NoAggregation 1, DispMemberNotFound 1, DispUnknownName 1. The other ten return
  zero references: NotImplemented (12), NoInterface (13), OutOfMemory (16), Unexpected (17),
  ClassNotRegistered (20), DispBadParamCount (24), OleUseRegistry (28), OleInvalidVerb (31),
  OleAdviseNotSupported (33), OleNoConnection (34). No call site uses the raw literals either - `grep
  -rn "0x80004002|0x80004001|0x8000FFFF|0x80040154"` over Com/ and AddIn/XlideAddIn.cs matches only
  the declarations in HResult.cs, and every HRESULT returned by ClassFactory.cs, Exports.cs and
  XlideAddIn.cs comes from the named constants (ClassFactory.cs:26 `return HResult.NoAggregation;`,
  Exports.cs:79 `DllCanUnloadNow() => HResult.False;`, XlideAddIn.cs:333 `return resolvedAll ?
  HResult.Ok : HResult.DispUnknownName;`). The four Ole* entries are the tell: OLE_S_USEREG,
  OLEOBJ_S_INVALIDVERB, OLE_E_ADVISENOTSUPPORTED and OLE_E_NOCONNECTION are the vocabulary of an
  in-place-activated embedded object, which this add-in is not.
- **Why:** The file's summary calls these "The HRESULT values this server returns or inspects", and ten of them
  it does neither. The four OLE codes are worse than inert: they describe an interface surface the
  shim once had and no longer implements, so a reader inferring the COM contract from this file infers
  an embedding-capable object.
- **Change:** Delete the ten unreferenced constants. If the OLE four are being kept deliberately against a future
  in-place-activation path, say that in a comment naming the path, because right now they read as live
  contract.
- **Size:** 10 of 18 constants, about 12 lines
- **Adversary:** Read src/Xlide.Vbe.Shim/Com/HResult.cs in full (35 lines). Checked the ten named constants three
  ways: `git grep -w -c 'HResult.<name>'` over src, tests and tools returns nothing for
  NotImplemented, NoInterface, OutOfMemory, Unexpected, ClassNotRegistered, DispBadParamCount,
  OleUseRegistry, OleInvalidVerb, OleAdviseNotSupported and OleNoConnection; a bare-identifier grep
  over src and tests matches only their own declarations at lines 12, 13, 16, 17, 20, 24, 28, 31, 33,
  34; and `git grep 'using static'` over src returns nothing, so no unqualified path exists. The
  raw-hex grep likewise matches only HResult.cs. The OLE-era reading is supported: `git grep -in
  'IOleObject|OLE_S_USEREG|embedding'` over docs and src finds no IOleObject implementation anywhere,
  only the HResult.cs comment and two docs lines about Excel's automation embedding mode, which is a
  different subject.
- **Correction applied:** The cost is readability only. `const int` fields are compile-time and carry no runtime or binary
  cost, so 'about 12 lines' is the whole size. The file summary at :3-5 ('The HRESULT values this
  server returns or inspects') is the thing that is actually false; the four Ole* names are the ones
  worth removing, while NotImplemented and NoInterface are the standard return vocabulary a
  hand-written QueryInterface would need and are cheap either way.

##### `win32-unused-imports-and-constants` Two P/Invoke declarations and thirteen message and style constants in Interop/Win32 have no call site

- **Where:** `src/Xlide.Vbe.Shim/Interop/Win32.cs:129`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** Counted whole-word occurrences across src, tests, installer and tools; each of these appears exactly
  once, at its own declaration. Imports: Win32.cs:129 `public static partial nint SetParent(nint
  child, nint parent);` and Win32.cs:132 `public static partial nint GetForegroundWindow();` -
  confirmed separately with `grep -rn "SetParent|GetForegroundWindow" --include='*.cs' src`, which
  returns only those two lines. Constants in Win32.cs: WsVisible (114), CsDblClks (136), WmCreate
  (140), WmDestroy (141), WmSetFocus (143), ColorButtonFace (154), WmSysCommand (227), ScClose (230).
  Constants in Win32.Events.cs: GaParent (34), LrDefaultSize (70), GwlStyle (201), WsMaximize (203),
  WmMdiRestore (206). None of these is a string-dispatch target: they are C# identifiers consumed only
  by C#, so the identifier grep is the whole answer for them.
- **Why:** SetParent in particular is a loaded declaration to leave sitting in the interop surface of a product
  whose entire placement strategy is a layered overlay that deliberately does not reparent the host's
  windows; its presence suggests reparenting is an available move. WmSysCommand with ScClose, and
  GwlStyle with WsMaximize and WsVisible, similarly read as live window-management vocabulary that
  nothing exercises. Under NativeAOT the two LibraryImport declarations also generate marshalling
  stubs for entry points nothing binds.
- **Change:** Delete the two imports and the thirteen constants. Anything genuinely reserved for imminent work
  should carry a one-line comment naming that work, so the next dead-code pass does not have to
  re-derive the same answer.
- **Size:** 2 P/Invokes and 13 constants, about 25 lines across two files
- **Adversary:** Ran a whole-word count for each of the fifteen names over src, tests, installer, tools, ui and docs.
  Every one returns exactly one file with one hit, its own declaration: SetParent
  (src/Xlide.Vbe.Shim/Interop/Win32.cs:129), GetForegroundWindow (Win32.cs:132), WsVisible (114),
  CsDblClks (136), WmCreate (140), WmDestroy (141), WmSetFocus (143), ColorButtonFace (154),
  WmSysCommand (227), ScClose (230), and in Win32.Events.cs GaParent (34), LrDefaultSize (70),
  GwlStyle (201), WsMaximize (203), WmMdiRestore (206). I also checked WsVisible is not folded into
  the composite WsOverlappedWindow at Win32.cs:120-121 (it is not:
  Caption|SysMenu|ThickFrame|MinimizeBox|MaximizeBox). None of these is a string-dispatch target, so
  the identifier grep is the whole answer.
- **Correction applied:** The NativeAOT size claim is wrong: an unreferenced LibraryImport partial is not rooted by ILC and
  never reaches the binary, and consts vanish at compile time, so the removable size is 0 bytes
  shipped and about 25 source lines. The real cost is the one the finding names second - SetParent
  sitting in the interop surface of a product whose placement strategy is a non-reparenting overlay,
  and GwlStyle/WsMaximize/WmMdiRestore/WmSysCommand/ScClose reading as live window management. Judge
  it as a readability finding, not a footprint one.

##### `vbemenus-is-fully-live` Nothing in VbeMenus is left over: all four of its public members are reachable, and the wrench routes through the same addressing the retired bar used

- **Where:** `src/Xlide.Vbe.Shim/Editor/VbeMenus.cs:199`
- **Kind:** dead-code / small effort, claim derived, confidence verified, severity low
- **Evidence:** Checked every public member against the whole repo. Read (199) is called from AddInSession.cs:4207
  `var items = VbeMenus.Read(_editor, path);`. ControlAt (459) from AddInSession.cs:4233 `using var
  control = VbeMenus.ControlAt(_editor, path);`. Describe (401) from AddInSession.DebugApi.cs:2652,
  serving the `menus` route documented at docs/debug-api.md:86. FindMenuBar (505) from
  AddInSession.cs:7196, and its own comment says why it survives the bar's retirement - "Also
  consulted for its height, which is where the loader's ground begins." The private machinery is
  load-bearing for the wrench: Read's `path.Length == 0` branch (256-265) appends the single synthetic
  item `new SurfaceMenuItem(XlideMenuPosition, "VBA", ..., Popup: true, ..., Icon: "wrench")`,
  ReadXlideMenu (274) composes Tools (30007) and Add-Ins (30038) into it, Resolve (335) turns `rank *
  SourceStride + realPosition` back into a real position chain, and PositionOf (374) finds each source
  menu by id. The Replaced set (136-163) is still consulted on every read, by both Read (232) and
  ReadXlideMenu (307), and Describe reports membership per control so the table is readable from the
  running editor. tools/harness/menu-bar.mjs:30 mirrors the synthetic position and drives it.
- **Why:** This is the negative answer the audit asked for, and it is worth stating so nobody spends the sweep
  here: the menu replication machinery survived the v0.6.0 retirement intact because the wrench reuses
  all of it - the position-chain addressing, the live read, the suppression filter and the property
  reads that tolerate refusal. Deleting any of it breaks the one menu that is left. The leftovers from
  that change are in VbeCommands, not here.
- **Change:** No change. Recorded so the next reader does not re-open it. The one thing worth doing is unrelated
  to dead code: the Options-522 comment (150-158) flags Require Variable Declaration as a capability
  the suppression dropped with no replacement, and calls it "a real gap, not a rehoming" - that is a
  product item for the settings dialog, not a cleanup.
- **Size:** unmeasured
- **Adversary:** Independently verified the negative. Grepping member declarations in
  src/Xlide.Vbe.Shim/Editor/VbeMenus.cs gives Read (199), Describe (401), ControlAt (459) and
  FindMenuBar (505) as the only non-private members, and `git grep 'VbeMenus.'` finds a caller for
  each: AddInSession.cs:4207 (Read), AddInSession.DebugApi.cs:2652 (Describe, serving the `menus`
  route at docs/debug-api.md:86), AddInSession.cs:4233 (ControlAt), AddInSession.cs:7196
  (FindMenuBar). The private machinery is wired too: Read:203-205 routes XlideMenuPosition to
  ReadXlideMenu (274), which is also called from Describe:413; Resolve (335) is called at 208, 428 and
  471; PositionOf (374) at 280 and 349; ExtractShortcut (531) at 238 and 311; ControlsAt (480) at 208,
  428 and 472; MenuBarType and PopupControlType are read at 513 and 237/310/450.
  tools/harness/menu-bar.mjs:30 mirrors the 900 position. Nothing in the file is orphaned.
- **Correction applied:** Two nits. FindMenuBar is `internal`, not public, so the count is three public members plus one
  internal. And this is a negative result, not a finding: it proposes no change and changes nobody's
  next action beyond saving a second look, so it should be carried as a coverage note rather than an
  audit item.

### The Monaco page

_9 findings, from the `complexity-page` finder._

##### `syncengine-dropped-by-page-update` The page's updateSettings omits syncEngine, so the settings dialog cannot set it and every other setting change silently resets it to "xlide"

- **Where:** `ui/editor/src/bridge.ts:672`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** bridge.ts:672-681 posts the whole settings object minus one key: `updateSettings(settings:
  EditorSettings): void { this.transport.post({ type: "updateSettings", blockLayout: ...,
  continueCommentOnNewline: ..., mirrorCommentSpacing: ..., treeFollowsEditor: ..., formatIndentSize:
  settings.formatIndentSize, }); }` - no syncEngine. The outbound message type at bridge.ts:300-308
  has no syncEngine member either, while the inbound `setSettings` type at bridge.ts:91-99 does.
  settingsdialog.ts:56 declares `key: "syncEngine" as const, kind: "choice"` and its handler at
  settingsdialog.ts:347-349 is `update({ ...currentSettings(), [option.key]: select.value })`.
  main.ts:471 and main.ts:570 both wire that callback straight to `bridge.updateSettings(next)`. On
  the host, EditorSurface.cs:1623-1626 reads a MISSING syncEngine as the literal default rather than
  the stored value: `var syncEngine = document.RootElement.TryGetProperty("syncEngine", out var
  engineValue) && engineValue.ValueKind == JsonValueKind.String ? engineValue.GetString() ?? "xlide" :
  "xlide";` and EditorSurface.cs:1628-1636 then builds a fresh ProductSettings from those six values
  and fires SettingsChangeRequested, which persists it. Contrast the debug-api route,
  AddInSession.DebugApi.cs:2997-2999, which preserves it: `SyncEngine =
  request.Query.TryGetValue("syncEngine", out var planner) ? ... : settings.SyncEngine`. The only
  thing in the repo that ever exercises the setting is tools/harness/module-sync.mjs:35 `await
  api.settings({ syncEngine: planner })`, which takes the api path.
- **Why:** Picking "Built into the add-in" in the settings dialog does nothing at all - the host writes "xlide"
  back and the echo re-renders the select at its old value. Worse, changing any unrelated setting
  (indent size, block layout) also resets the planner to xlide, so a developer who set it through the
  api or by hand loses it the next time they touch the dialog. This is exactly the api-mirrors-the-UI
  rule inverted: the api route reaches a state the UI cannot, so the whole harness passes over a
  broken UI path, and docs/debug-api.md:90 asserts the opposite of what is true - "the page's own
  update takes the whole object".
- **Change:** Add `syncEngine: string` to the `updateSettings` member of the client message union
  (bridge.ts:300-308) and pass `settings.syncEngine` in bridge.ts:672-681. Separately, make the host's
  absent-field branch fall back to the stored value the way the debug-api route does, so a page that
  omits a field can never reset one. Drive the dialog's select through the page, not through
  api.settings, in whichever live suite covers settings.
- **Size:** one omitted field; one setting unreachable from the UI and silently reset by five other UI actions
- **Adversary:** Read every link myself. ui/editor/src/bridge.ts:672-681 posts exactly five keys (blockLayout,
  continueCommentOnNewline, mirrorCommentSpacing, treeFollowsEditor, formatIndentSize) and no
  syncEngine; the outbound union member at bridge.ts:300-308 also lacks it while the inbound
  setSettings at bridge.ts:91-99 has `syncEngine?: string`. ui/editor/src/settingsdialog.ts:56 does
  declare `key: "syncEngine" as const, kind: "choice"` and the shared select handler at
  settingsdialog.ts:346-348 calls `update({ ...currentSettings(), [option.key]: select.value })`;
  main.ts:471 and main.ts:570 wire that update straight to `bridge.updateSettings(next)`. On the host,
  EditorSurface.cs:1623-1626 turns a missing syncEngine into the literal "xlide" and
  EditorSurface.cs:1628-1636 builds a fresh ProductSettings from the six values and invokes
  SettingsChangeRequested. The debug-api route at AddInSession.DebugApi.cs:2997-2999 does preserve it
  (`: settings.SyncEngine`). Only harness driver is tools/harness/module-sync.mjs:34-36,85, which uses
  the api path.
- **Correction applied:** The finding understates the blast radius: it is not just a session reset. AddInSession.cs:861-869
  OnSettingsChanged assigns `_settings = updated.Normalized()` wholesale and immediately
  `File.WriteAllText(path, _settings.ToJson())`, so any settings-dialog change PERSISTS
  SyncEngine="xlide" to the settings file, overwriting a planner choice made earlier through the api.
  Otherwise the finding is accurate line for line.

##### `settings-shape-written-five-times` The six settings keys are spelled out structurally in five places, and nothing makes the copies agree

- **Where:** `ui/editor/src/bridge.ts:91`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** The same key list appears as: settings.ts:10-27 `interface EditorSettings`; bridge.ts:91-99 the
  inbound `{ type: "setSettings"; blockLayout: string; continueCommentOnNewline: boolean;
  mirrorCommentSpacing: boolean; treeFollowsEditor: boolean; formatIndentSize?: number; syncEngine?:
  string; }`; bridge.ts:300-308 the outbound `{ type: "updateSettings"; ... }` with five of the six;
  bridge.ts:672-681 the hand-written post body; bridge.ts:1333-1340 the hand-written `applySettings({
  ... })` adopt. settings.ts:49-58 `applySettings` normalises all six a sixth time. bridge.ts never
  imports the `EditorSettings` type into either message shape - it restates the members inline, so the
  compiler cannot notice that the outbound shape is one member short of the inbound one.
- **Why:** This is the mechanism that produced the syncEngine bug and will produce the next one: adding a
  setting means five coordinated edits in two files, and the type system checks none of them against
  each other because every copy is structurally independent. The two shapes that must agree - what the
  page sends and what the page receives - are the two that diverged.
- **Change:** Define the wire payload once as `Omit<EditorSettings, never>` (or a `SettingsPayload` type in
  settings.ts) and use it in both message members: `{ type: "setSettings" } & SettingsPayload` and `{
  type: "updateSettings" } & SettingsPayload`. Then `updateSettings` can post `{ type:
  "updateSettings", ...settings }` and the adopt case can call `applySettings(message)` directly,
  leaving the coercion in the one place that already does it (settings.ts:49).
- **Size:** 5 declarations of the same 6 keys collapse to 1; roughly 35 lines removed from bridge.ts
- **Adversary:** All copies read and present: settings.ts:10-27 (interface EditorSettings), bridge.ts:91-99 (inbound
  setSettings, members restated inline with two optional), bridge.ts:300-308 (outbound updateSettings,
  five of six members), bridge.ts:672-681 (hand-written post body), bridge.ts:1332-1340 (hand-written
  applySettings adopt with its own coercions), settings.ts:49-58 (applySettings normalising all six
  again). Neither message member references the EditorSettings type, so nothing makes the inbound and
  outbound shapes agree - which is precisely how finding 1 became possible.
- **Correction applied:** Two numbers to fix. It is six structural spellings on the page side, not five (interface, inbound
  member, outbound member, post body, adopt block, normaliser), and three more on the host side that
  the finding does not count: EditorMessages.cs:84 (SetSettingsMessage carries syncEngine),
  DebugServer.cs:953 (DebugSettingsReply), and the field-by-field parse at EditorSurface.cs:1600-1636.
  Also, the proposed fix as written would change wire tolerance: the inbound member deliberately marks
  formatIndentSize and syncEngine optional so a host that predates them still type-checks, so the
  shared payload type has to stay optional on the inbound side (`{ type: "setSettings" } &
  Partial<SettingsPayload>`) while the outbound one is total.

##### `settheme-dead-end-to-end` The setTheme message is handled by the page and registered with the AOT serializer, but nothing in the repo ever sends it

- **Where:** `ui/editor/src/bridge.ts:1162`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** bridge.ts:62 declares `| { type: "setTheme"; theme: XlideTheme }` and bridge.ts:1162-1165 handles
  it: `case "setTheme": ... monaco.editor.setTheme(message.theme === THEME_LIGHT ? THEME_LIGHT :
  THEME_DARK);`. On the host, EditorMessages.cs:43-45 declares `public sealed record
  SetThemeMessage([property: JsonPropertyName("type")] string Type, [property:
  JsonPropertyName("theme")] string Theme);` and EditorMessages.cs:522 registers it:
  `[JsonSerializable(typeof(SetThemeMessage))]`. `grep -rn "SetThemeMessage" src --include=*.cs`
  returns exactly those two lines - the record is never constructed. A repo-wide grep for the string
  `setTheme` across .cs, .ts, .mjs, .md and .ps1 (excluding node_modules and dist) returns only
  bridge.ts:62/1162, the monaco call at bridge.ts:1353, and README.md:50. demoTransport
  (bridge.ts:1936) does not send it either.
- **Why:** A whole protocol path - page handler, host record, source-generated serializer entry, and a README
  row - exists for a message no code produces. The JsonSerializable registration is not free under
  NativeAOT: it emits a converter for a type that will never be serialized. Anyone tracing how the
  theme is chosen follows this into a dead end instead of finding theme.ts:118 preferredTheme /
  watchPreferredTheme, which is what actually decides it.
- **Change:** Either delete the path (bridge.ts:62 and 1162-1165, EditorMessages.cs:43-45 and 522, README.md:50)
  or give it a sender if host-pinned theming is still wanted. The `themePinned` flag at bridge.ts:548
  only ever reads false today, so deleting is the smaller correct change.
- **Size:** about 12 lines across three files, plus one AOT serializer registration
- **Adversary:** git grep for setTheme and SetThemeMessage across the repo excluding dist and node_modules returns
  exactly: EditorMessages.cs:43 (the record declaration), EditorMessages.cs:522
  ([JsonSerializable(typeof(SetThemeMessage))]), README.md:50, bridge.ts:62 (union member),
  bridge.ts:1162-1165 (the handler), bridge.ts:1353 (the unrelated monaco call inside applyOsTheme).
  The record is never constructed anywhere in src. Every host-to-page message goes through typed
  `Post(JsonSerializer.Serialize(...))` calls in EditorSurface.cs (about 28 of them, lines 249-1193),
  so there is no raw-JSON side door that could carry the kind, and no debug-api route or devsurface
  act names the string either.
- **Correction applied:** The dead surface is slightly larger than stated. Because nothing ever sends setTheme, `themePinned`
  (bridge.ts:477) is only ever assigned true at bridge.ts:1163 inside the dead case, so the guard at
  bridge.ts:1350 in applyOsTheme is never taken, and the public getter `isThemePinned` at
  bridge.ts:548-550 has zero callers anywhere in the repo. Deleting the message should take the pin
  flag and the getter with it.

##### `applyedit-dead-on-the-page` applyEdit is a fully implemented host-to-page path with no sender anywhere and no host record at all

- **Where:** `ui/editor/src/bridge.ts:1572`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** bridge.ts:61 `| { type: "applyEdit"; revision: number; changes: HostTextChange[] }`,
  bridge.ts:1159-1161 dispatches it, and bridge.ts:1572-1597 implements it in full -
  `pushEditOperations`, the applyingHostEdit guard, and the revision adoption
  `this.revisions.set(docKeyOf(shown.module, shown.project), revision)`. A repo-wide grep for
  `applyEdit` across .cs, .ts, .mjs, .md, .ps1 (excluding node_modules and dist) finds no sender: the
  only C# hits are the unrelated private helper `ApplyEdits` at EditorSurface.cs:2037, and there is no
  ApplyEditMessage record in EditorMessages.cs. The private `applyEdit` at bridge.ts:1572 has exactly
  one caller, the dead case at 1160. The echo-suppression flag it sets is NOT dead - applyingHostEdit
  is also set at bridge.ts:1496 and 1541 - so only this branch goes.
- **Why:** Twenty-six lines of the most delicate code in the bridge (undo-stack semantics, echo suppression,
  revision authority transfer) sit unreachable and untested. It reads as the mechanism by which the
  host writes into a document, so the next person implementing a host-side edit will wire to it rather
  than to syncDocument (bridge.ts:1565), which is what the host actually uses. The revision-adoption
  semantics described in README.md:70-73 describe this dead branch, not shipped behaviour.
- **Change:** Delete bridge.ts:61, 1159-1161 and 1572-1597, and the README paragraph at lines 70-73 that describes
  revision adoption after applyEdit. If a host write path is planned, note it in docs/decisions.md
  rather than leaving a stub the compiler cannot tell is unreachable.
- **Size:** about 30 lines of unreachable code in the page's largest file
- **Adversary:** bridge.ts:61 declares the union member, bridge.ts:1159-1161 dispatches to
  `this.applyEdit(message.revision, message.changes)`, and bridge.ts:1572-1597 is the full
  implementation (pushEditOperations with forceMoveMarkers, the applyingHostEdit guard in try/finally,
  then `this.revisions.set(docKeyOf(shown.module, shown.project), revision)`). git grep for applyEdit
  across the repo excluding dist and node_modules returns no sender: the only other hits are
  engine/test/smoke.mjs's unrelated local `applyEdits` helper and the README rows. There is no
  ApplyEditMessage record in EditorMessages.cs, and the serializer context at
  EditorMessages.cs:515-530 has no entry for one, which under NativeAOT means the host could not
  serialize such a message even if someone wrote the call. The host's real write path is syncDocument
  (bridge.ts:1565).

##### `readme-protocol-table-stale` ui/editor/README.md's host-to-page protocol table documents a message that does not exist and covers 7 of the 38 kinds the page dispatches

- **Where:** `ui/editor/README.md:48`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** README.md:48 is the table's first row: `| \`loadDocument\` | \`moduleName: string\`, \`text:
  string\` | Swaps in a model at \`xlide:/<moduleName>\`, resets revision to 0. No echo. |`.
  `loadDocument` appears nowhere in ui/editor/src - grep finds it only at README.md lines 48, 70 and
  73. The page's actual dispatch (bridge.ts:1115-1340) has 38 cases and the equivalent one is
  `openDocument` (bridge.ts:1115). The table lists seven rows: loadDocument (nonexistent), applyEdit
  (no sender), setTheme (no sender), setDiagnostics, setCurrentLine, setBreakpoints, revealLine. Three
  of the seven are unreachable and 32 real kinds are absent.
- **Why:** This is the page's only written protocol reference. More than half of it is wrong and the half that
  is right is a small minority of the surface, so anyone coding a host message against it starts from
  a name the page will hit `default:` on and log as unhandled (bridge.ts:1341-1344). The 2026-08-09
  lesson about reading source instead of trusting a probe applies here too: the doc is worse than no
  doc because it is confidently specific.
- **Change:** Either regenerate the two tables from the ClientMessage and HostMessage unions in bridge.ts, or cut
  them down to a pointer at bridge.ts:40-100 and 260-310 as the authority. Whichever way, remove the
  loadDocument row and the revision paragraph at lines 70-73 alongside the applyEdit deletion.
- **Size:** 3 of 7 documented host messages do not exist; 32 real ones undocumented
- **Adversary:** Read ui/editor/README.md:44-73. The host-to-page table has seven rows and its first is
  `loadDocument`, which git grep finds nowhere in ui/editor/src - only at README.md:48, 70 and 73. The
  equivalent live kind is `openDocument` (bridge.ts:1115). applyEdit and setTheme are the two rows
  verified dead above, and the revision paragraph at README.md:70-73 describes applyEdit and
  loadDocument, neither of which exists as shipped behaviour. Unknown kinds fall through to `default:`
  at bridge.ts:1341-1344 and only console.warn. This README is the only protocol table in the repo
  (git grep "Host to page" over all .md returns just README.md:44).
- **Correction applied:** The counts are off. The dispatch at bridge.ts:1115-1332 has 39 case labels, not 38 (three of them -
  obLibrariesResult, obTypesResult, obMembersResult - share one deliberate no-op at 1326-1331). Four
  of the seven documented rows are real, so about 35 real kinds are undocumented, not 32. The
  page-to-host table (README.md:60-66) is worse than the finding says: four rows against roughly 50
  client message kinds at bridge.ts:260-310.

##### `ten-pending-request-tables` bridge.ts carries ten hand-copied request/response tables where one generic helper would do

- **Where:** `ui/editor/src/bridge.ts:405`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity low
- **Evidence:** Ten `Map<number, { resolve; timer }>` fields at bridge.ts:405-462 (pendingCompletions,
  pendingHovers, pendingSignatures, pendingCanonicalCases, pendingCodeActions, pendingSemanticTokens,
  pendingNavigations, pendingRenames, pendingOutlines, pendingSyncs), ten counters at
  bridge.ts:464-473 (nextSyncId ... nextOutlineId), ten request bodies that are the same eight lines
  with the names changed - bridge.ts:710-722 is `const id = this.nextCompletionId++; return new
  Promise((resolve) => { const timer = setTimeout(() => { this.pendingCompletions.delete(id);
  resolve([]); }, 2000); this.pendingCompletions.set(id, { resolve, timer }); this.transport.post({
  type: "completion", id, offset }); });` and bridge.ts:728-740, 746-758, 764-776, 782-804, 825-840,
  844-856, 866-880, 896-935, 1050-1062 repeat it verbatim with a different map, counter, post type and
  empty value. Then ten resolve blocks that are the same six lines: bridge.ts:1200-1207 `const waiter
  = this.pendingCompletions.get(message.id); if (waiter) { this.pendingCompletions.delete(message.id);
  clearTimeout(waiter.timer); waiter.resolve(message.items); }`, repeated at 1209, 1218, 1227, 1236,
  1245, 1258, 1267, 1277, 1290. The 2000ms deadline is written out ten times as a bare literal.
- **Why:** Roughly 400 lines of the repo's largest page file carry one idea. Every new host round trip costs
  four coordinated edits in four separate regions of a 2356-line file, and a missed clearTimeout or a
  wrong map in the resolve block is a leaked timer or a promise that resolves empty after two seconds
  - both of which look like a slow host rather than a bug. The 2000ms budget cannot be tuned or read
  off in one place.
- **Change:** One `private readonly requests = new RequestTable()` with `ask<T>(kind: string, body: object,
  fallback: T, timeoutMs = 2000): Promise<T>` and `settle(kind: string, id: number, value: unknown):
  void`. The ten dispatch cases collapse to a handful of lines mapping result kind to payload field,
  the ten maps and counters to one map of maps and one counter, and the deadline becomes a single
  named constant.
- **Size:** about 400 lines carrying one pattern; 10 maps + 10 counters + 10 request bodies + 10 resolve blocks
- **Adversary:** Counted and read the pattern myself. Ten Map fields at bridge.ts:405-462 and ten counters at
  bridge.ts:464-473, exactly as listed. Twelve promise-returning request methods (bridge.ts:710, 728,
  746, 764, 782, 824, 843, 862, 896, 910, 924, 1045) each repeat the same eight-line body - allocate
  id, setTimeout that deletes from the map and resolves a fallback, map.set, transport.post - and ten
  resolve blocks in the dispatch (1200, 1209, 1218, 1227, 1236, 1245, 1258, 1267, 1277, 1290) repeat
  the same get / delete / clearTimeout / resolve. The duplication and the four-coordinated-edits cost
  are real, not a length complaint.
- **Correction applied:** The deadline claim is wrong and it matters, because it is the part of the proposed fix that would
  change behaviour. `2000` appears at five sites only (bridge.ts:717, 735, 753, 771, 793).
  requestOutline (843), requestNavigation (862) and requestSemanticTokens (1045) use 8000; the three
  rename methods (896, 910, 924) use 30000; requestSync (824) uses 120000 with a comment saying why.
  The fallback values differ too ([] / null / {error: ...} / a refused HostRenameAnswer). Also it is
  twelve request methods over ten tables, since the three rename methods share pendingRenames. A
  shared helper is still the right shape, but the timeout and the fallback must stay per-call
  arguments, not a single named constant.

##### `modal-plumbing-six-copies` Six dialogs hand-roll the same backdrop, Escape handler and dismiss, and five re-declare CSS a general .modal-backdrop rule already provides

- **Where:** `ui/editor/src/helpdialog.ts:76`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity medium
- **Evidence:** helpdialog.ts:76-90 and sponsordialog.ts:59-73 are the same code with one word changed - `const
  existing = document.getElementById("help-card"); if (existing) {
  existing.querySelector<HTMLElement>("#help-close")?.focus(); return; } const backdrop =
  document.createElement("div"); backdrop.id = "help-backdrop"; const card =
  document.createElement("div"); card.id = "help-card"; card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");` against the identical block with "sponsor". Their
  dismissals match too: helpdialog.ts:195-215 and sponsordialog.ts:169-190 both define `dismiss` that
  does `document.removeEventListener("keydown", onKey, true); backdrop.remove(); closed?.();`, the
  same `onKey` that preventDefaults Escape, the same backdrop mousedown check `event.target ===
  backdrop`, and the same captured `document.addEventListener("keydown", onKey, true)`.
  settingsdialog.ts:410-429, referencesdialog.ts:142-163, syncdialog.ts:740-760 and shell.ts:819-821
  are four more copies of that dismissal. On the CSS side the general mechanism already exists:
  styles.css:1877-1886 `.modal-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.35);
  display: flex; align-items: flex-start; justify-content: center; padding-top: 18vh; z-index: 95; }`
  with .modal-card, .modal-title, .modal-detail, .modal-buttons, .modal-button beside it - and only
  shell.ts:796 and shell.ts:950 use it. The other five declare byte-identical geometry under their own
  ids: styles.css:1190 (#help-backdrop, padding-top 12vh, z-index 90), :1360 (#sponsor, 12vh, 90),
  :1660 (#settings, 12vh, 90), :2798 (#references, 10vh, 90), :2922 (#sync, 6vh, 90).
- **Why:** Six copies means the things a modal should do get written zero times: no dialog in the page
  implements a focus trap, so Tab walks straight out of an aria-modal="true" card into the Monaco
  surface behind it (grep for Tab handling across the five dialog modules finds only an unrelated
  comment at syncdialog.ts:281). It also means the modals disagree visibly - five different vertical
  offsets (6, 10, 12, 12, 12 vs 18vh) and two stacking levels (90 vs 95), so the close-confirm at
  shell.ts:796 sits above the others and the others sit below it, which is the wrong way round when a
  close-confirm is raised from a dialog.
- **Change:** Promote the shell.ts pattern to `openModal({ id, label, build }): { dismiss }` in one module: it
  creates the .modal-backdrop/.modal-card pair, sets role and aria-modal, installs the captured Escape
  and the backdrop mousedown, traps Tab inside the card, restores focus on dismiss, and returns
  dismiss. Point the five dialogs at it and delete the five #*-backdrop rules, keeping only whatever
  per-dialog width and padding genuinely differs.
- **Size:** 6 copies of ~20 lines of TS plus ~50 lines of duplicated CSS; one focus trap written once instead of six times
- **Adversary:** Opened all six. helpdialog.ts:76-90 and sponsordialog.ts:59-73 are the same block with the word
  changed (same getElementById guard, same focus-what-is-open early return, same backdrop/card
  creation, same role and aria-modal). The dismissals match at helpdialog.ts:195-215,
  sponsordialog.ts:169-190, settingsdialog.ts:408-429, referencesdialog.ts:142-163 (as
  closeReferencesDialog plus a module-level onKey), syncdialog.ts:739-761 and shell.ts:819-826: each
  removes the captured keydown listener, removes the backdrop, calls closed, and each re-declares the
  same Escape handler and the same `event.target === backdrop` check. CSS verified line by line:
  styles.css:1877-1886 .modal-backdrop is used only at shell.ts:796 and shell.ts:950, while
  #help-backdrop (1190), #sponsor-backdrop (1360), #settings-backdrop (1660), #references-backdrop
  (2798) and #sync-backdrop (2922) each repeat position/inset/background
  rgba(0,0,0,0.35)/flex/flex-start/center verbatim, differing only in padding-top (12/12/12/10/6vh)
  and z-index 90. The focus-trap gap is real: grep for Tab key handling across ui/editor/src returns
  only a comment in devsurface.ts:1285, so six cards carry aria-modal="true" with nothing keeping Tab
  inside them.
- **Correction applied:** One sub-claim in the rationale is backwards and should be dropped: .modal-backdrop is z-index 95
  against the dialogs' 90, so the close-confirm raised from a dialog correctly renders ABOVE it - that
  is the right way round, not the wrong one. The differing padding-top values are also defensible per
  card height (the sync dialog is the tallest and sits highest), which the proposed fix already allows
  for. What survives is the duplicated TS plumbing, the five redundant backdrop rules, and the missing
  focus trap - the last being the only one that costs a user anything.

##### `split-tree-algebra-three-copies` The split-tree prune and same-direction absorb are written three times, and only the copy in docktree.ts is tested

- **Where:** `ui/editor/src/workspace.ts:946`
- **Kind:** complexity / large effort, claim observed, confidence verified, severity low
- **Evidence:** docktree.ts is the extracted, tested version - its header at docktree.ts:1-15 says "Both layouts are
  the same shape ... It lives here, separately and purely, so it can be tested by calling it", and
  test/docktree.mjs:60-186 exercises prune, splitBeside, firstGroup, allGroups, groupHolding and
  indexAtMidpoints. workspace.ts imports only `resizeAt` from it (workspace.ts:29) and re-implements
  the rest: workspace.ts:946-970 is docktree.ts:78-105's prune with `leaf` for `group`, and
  workspace.ts:895-929 is docktree.ts:128-162's `replace` including the same-direction absorb - `if
  (child.kind === "split" && child.direction === node.direction) { ... sizes.push((node.sizes[index]
  ?? 1) * (child.sizes[inner] ?? 1)); }`. The copies have already drifted: docktree.ts:160-161
  renormalises after absorbing (`const total = sizes.reduce(...) || 1; ... sizes.map((s) => s /
  total)`) and workspace.ts:928 does not - it returns `{ ...node, children: flattened, sizes }` raw.
  paneldocks.ts imports `prune` at paneldocks.ts:22 and uses it at paneldocks.ts:804, then defines a
  THIRD local `const prune` at paneldocks.ts:308-338 that shadows the import inside pruneUnknown; its
  lines 321-337 are the imported function's body verbatim, wrapped around a tab filter.
- **Why:** The module written specifically so this arithmetic could be tested by calling it covers one of its
  three users. The bug class its own header names - "removing a pane can empty a group, which can
  leave a split with one child, which must collapse into its parent, and the sizes have to stay a
  partition of one" - is live in two untested copies, and the divergence is already there in the
  renormalise. The shadowed name in paneldocks.pruneUnknown is worse than the duplication: a reader
  who sees `prune` in that method and looks at the import at line 22 reads the wrong function.
- **Change:** Have workspace.ts adopt docktree's TreeNode by holding its DOM element in a side map keyed by node,
  then call the shared prune and splitBeside instead of its private copies. For
  paneldocks.pruneUnknown, split the tab filter from the collapse: map the tree's groups through the
  known/placed filter first, then call the imported prune, and rename or remove the local so nothing
  shadows it.
- **Size:** 3 implementations of the same two functions, about 80 duplicated lines, 2 of 3 untested
- **Adversary:** Read all three. docktree.ts:78-105 prune and docktree.ts:128-162 splitBeside's replace are the
  tested originals (test/docktree.mjs exercises prune, the renormalise case, splitBeside and the
  rest). workspace.ts imports only resizeAt (workspace.ts:29) and re-implements both:
  workspace.ts:946-970 is prune with `leaf` for `group`, and workspace.ts:893-929 is the same replace
  including the same-direction absorb with the identical `sizes.push((node.sizes[index] ?? 1) *
  (child.sizes[inner] ?? 1))` arithmetic. paneldocks.ts imports prune at paneldocks.ts:21 and calls it
  at paneldocks.ts:803, then declares a local `const prune` at paneldocks.ts:308-338 that shadows the
  import inside that method; its lines 322-338 are the imported body verbatim wrapped around a
  known/placed tab filter.
- **Correction applied:** The drift is narrower than stated. workspace.ts's prune DOES renormalise (workspace.ts:968-969:
  `const total = ...; sizes: sizes.map((size) => size / total)`); the only divergence is the absorb
  path at workspace.ts:928, which returns `{ ...node, children: flattened, sizes }` where
  docktree.ts:160-161 renormalises. And that divergence is arithmetically a no-op while sizes
  partition one, since a child split's shares are multiplied by its parent share and sum back to it -
  so this is a latent robustness difference, not an observed layout bug, and the finding should not be
  sold as one. Scope the fix accordingly: the paneldocks.ts:308 shadow is the cheap, low-risk half;
  adopting docktree's TreeNode in workspace.ts is a large refactor because LayoutNode carries a DOM
  `element` on every split node, which docktree's pure tree has no place for.

##### `devsurface-ships-on-a-wrong-number` devsurface.ts ships in the production bundle, which is deliberate, but the comment justifying it understates its size by two orders of magnitude

- **Where:** `ui/editor/src/devsurface.ts:18`
- **Kind:** complexity / small effort, claim derived, confidence verified, severity low
- **Evidence:** How it was established: main.ts:72 is a plain static import, `import { installDevSurface } from
  "./devsurface.js";`, called unconditionally at main.ts:929. build.mjs has one esbuild config used
  for both entry points with `minify: true` and no `drop`, no `pure`, no conditional define and no
  per-file exclusion - the only alias is `xlide-spec` to vendor/. And the strings are in the shipped
  artifact: grepping dist/editor.js finds `xlideUi`, `closeActive`, `longTasks` and `emptyViewShown`.
  So it ships, and there is no mechanism by which it could not. The justification at
  devsurface.ts:17-19 reads "Shipped in every build: it is a few hundred bytes of read-only reporting
  over objects the page already holds". The file is 63,412 bytes and 1,483 lines - the second largest
  in the page after bridge.ts.
- **Why:** The decision to ship the debug surface in Release is a good one and the reasoning behind it holds -
  a door only present in Debug is a door nobody trusts. What does not hold is the number it was
  justified on. Anyone revisiting the tradeoff reads "a few hundred bytes" and closes the question;
  the real figure is a file two thirds the size of the bridge, including a PerformanceObserver
  (devsurface.ts:220-225) that runs in every user's session. The comment should carry the number the
  decision actually costs.
- **Change:** Correct devsurface.ts:17-19 to state the real source size and what it costs minified, and note the
  always-on PerformanceObserver as part of what ships. If the cost turns out to matter, the split is
  between the read-only `state()` reporting (which is the part that must ship) and the `act()` driving
  surface (which only the harness calls) - but measure before splitting.
- **Size:** 63,412 bytes of source, 1,483 lines, confirmed present in dist/editor.js; minified contribution unmeasured because I am read-only and could not build
- **Adversary:** Verified each leg. main.ts:72 is a static import and main.ts:929 calls installDevSurface
  unconditionally. ui/editor/build.mjs has one `common` config with `minify: true` (build.mjs:266), no
  drop, no pure, and a define block (build.mjs:259-263) carrying only version stamps, shared by both
  entry points (build.mjs:290, 294) - no mechanism could exclude the module. dist/editor.js contains
  xlideUi, emptyViewShown and long-animation-frame. devsurface.ts is 63,412 bytes and 1,483 lines
  against its own comment at devsurface.ts:17-19 saying "a few hundred bytes of read-only reporting".
  watchLongTasks (devsurface.ts:220) is called unconditionally from installDevSurface at
  devsurface.ts:364, so the PerformanceObserver pair does run in every session.
- **Correction applied:** The comment was never accurate, so "understates by two orders of magnitude" should be "was wrong
  when written and is now 5.4x wronger": git log -S shows the "a few hundred bytes" line arrived in
  d92b25a, the commit that created the file, where devsurface.ts was already 11,724 bytes. The second
  half of the sentence is also stale - the module is no longer "read-only reporting", it carries the
  act() driving surface. The minified contribution to editor.js remains unmeasured, so any size figure
  in the corrected comment must be labelled as source bytes. This is a comment-accuracy fix; the
  ship-in-Release decision itself stands.

### Engine wrapper, harness, build and release tooling

_15 findings, from the `complexity-tooling` finder._

##### `audit-routes-counts-unrun-drivers` audit-routes.mjs counts every file in tools/harness as a driver, so four routes pass the gate's coverage check while nothing that runs drives them

- **Where:** `tools/harness/audit-routes.mjs:137`
- **Kind:** api-coverage / medium effort, claim derived, confidence verified, severity medium
- **Evidence:** The driver corpus is the whole directory, run or not: `const corpus = readdirSync(join(root,
  "tools/harness")).filter((file) => /\.(mjs|ps1)$/.test(file) && file !== "xlide-api.mjs" && file !==
  "audit-routes.mjs").map(...).join("\n")` (line 137-140), and `isDriven` just regexes that blob (line
  144-148). I extracted the 50 route names from src/Xlide.Vbe.Shim/AddIn/AddInSession*.cs with the
  same parser the audit uses, rebuilt methodsByRoute from tools/harness/xlide-api.mjs the same way,
  then intersected each route's driver set against the set verify.ps1 actually invokes (the probe list
  at verify.ps1:247, the PowerShell list at verify.ps1:369, the fixture plan at verify.ps1:420-427,
  and com-leak.mjs at verify.ps1:469). Four routes have drivers, and none of those drivers is ever
  run: `compile` -> debugger-features.mjs; `guard` -> debugger-features.mjs, step-into-features.mjs;
  `mark` -> write-rollback.mjs; `trip` -> perf-scaling.mjs. Confirmed by grep:
  `debugger-features.mjs:75: const compiled = await api.compile();`, `step-into-features.mjs:50: await
  api.guard(true);`, `write-rollback.mjs:86: await api.mark(...)`, `perf-scaling.mjs:37:const floor =
  await api.trip("pagecall", { n: 10 });` and no other harness file names any of the four.
- **Why:** The gate's last line is `${routes.length - excused.length} are driven by a probe`, and
  verify.ps1:263 prints it as the step's verdict. For compile, guard, mark and trip that sentence is
  false in the way that matters: the route can break and every gate stays green. guard is the one that
  hurts, because it is the break-mode guard the two debugger suites arm and disarm; a regression there
  is exactly the class of defect the live gate exists for.
- **Change:** Give audit-routes.mjs the list of files the gate runs (or read it out of verify.ps1) and split the
  verdict into driven-by-something-that-runs versus driven-only-by-an-unrun-script, failing on the
  second unless the file is named with a reason the way NOT_DRIVEN_ON_PURPOSE already names routes.
  Then either put debugger-features.mjs, step-into-features.mjs, write-rollback.mjs and
  perf-scaling.mjs into the -Live plan, or excuse them by name.
- **Size:** 4 of 50 routes reported as covered are not
- **Adversary:** tools/harness/audit-routes.mjs:137-140 builds the corpus from readdirSync over the whole harness
  directory and isDriven (144-148) regexes that blob. I listed every step in tools/verify.ps1 (grep of
  'Step ' plus the bodies at 246-252, 369, 420-427, 469): page probes are the five *-page-probe.mjs,
  live probes are Test-DebugApi/Test-SplitWorkspace/Test-DiscardProblems/Test-Churn, suites are
  import-guard, immediate-watch, analysis-freshness, menu-bar, module-sync x2, format-positions,
  three-copies, colouring, settings-bite, plus com-leak. grep over tools/harness for
  api.compile|api.guard|api.mark|api.trip returns only debugger-features.mjs:65,75,170,
  step-into-features.mjs:50,135, write-rollback.mjs:86,95, perf-scaling.mjs:37, and no .ps1 builds
  /compile, /guard, /mark or /trip by URL. None of those four files is in any verify.ps1 step.
- **Correction applied:** Two details. (1) The audit's own failure text is 'nothing in tools/harness drives it', so the check
  enforces driver-exists, not gate-runs-it; the misleading part is the summary line at
  audit-routes.mjs:174 and its use as the step verdict. (2) Of the four routes, only compile and guard
  carry behaviour worth a regression check; mark is a log marker and trip is a timing helper, and all
  four suites are documented hand-run tools (docs/driving-excel.md:446, 564, 674, 737). Size is better
  stated as 2 meaningful routes with no automated coverage, not 4 of 50 broken.

##### `engine-freshness-test-runs-nowhere` engine/test/freshness.mjs is in no runner: CI runs three of the four engine tests and verify.ps1 runs one

- **Where:** `engine/package.json:12`
- **Kind:** dead-code / small effort, claim derived, confidence verified, severity medium
- **Evidence:** `"test": "node test/smoke.mjs && node test/language.mjs && node test/positions.mjs && node
  test/freshness.mjs"` names all four. .github/workflows/build.yml runs them individually - `run: node
  test/smoke.mjs` (line 69), `run: node test/language.mjs` (line 85), `run: node test/positions.mjs`
  (line 94) - and never freshness. The `npm test` at build.yml:139 is under `working-directory:
  ui/editor`, so it is the page's test script, not the engine's. verify.ps1's 'engine language matrix'
  step runs `node test/language.mjs` alone (verify.ps1:240). A repo-wide grep for freshness.mjs finds
  only engine/package.json lines 12-13, the CI-less script alias `test:freshness`, and
  docs/lessons.md:1439 ("Five checks in engine/test/freshness.mjs").
- **Why:** 12 KB of checks over the re-analysis-on-callee-change invariant, the same invariant
  tools/harness/analysis-freshness.mjs guards at the live layer, never execute. The engine half can
  rot silently, and when the live suite goes red there is no cheap host-free test to bisect against.
- **Change:** Add `node test/freshness.mjs` as a step in the engine CI job beside Positions, or make the engine
  job run `npm test` so package.json stays the single list. Prefer CI over verify.ps1: it needs no
  Excel and both runners already build the bundle.
- **Size:** 11,989 bytes / 5 checks that never run
- **Adversary:** engine/package.json:12 lists all four tests; .github/workflows/build.yml:69, 85, 94 run smoke,
  language and positions individually and never freshness; build.yml:139 npm test is under
  working-directory ui/editor (line 138). tools/verify.ps1:229-244 runs only node test/language.mjs,
  and verify.ps1:219-227 npm test runs in $pageRoot. Repo-wide grep for freshness.mjs outside
  node_modules/dist finds engine/package.json:12-13 and docs/lessons.md:1439 only. ls -l engine/test
  shows freshness.mjs at 11,989 bytes.

##### `didclose-route-never-sent` The engine's module/didClose method is routed and documented but nothing has ever sent it, and the dispatcher says so in its own comment

- **Where:** `engine/src/dispatcher.ts:286`
- **Kind:** dead-code / small effort, claim observed, confidence verified, severity low
- **Evidence:** `case 'module/didClose': this.analysis.handle({ kind: 'forget', docKey: this.require<{ documentKey:
  string }>(params).documentKey }); return null;`. Its replacement is thirty lines below in
  closeProject: "It keeps an incremental parse per document and drops one when it is told that
  document closed - `module/didClose`, which this product has never sent, because a module's TAB
  closing is not the module leaving the project... Released here rather than by starting to send
  didClose" (lines 449-457), followed by the loop that forgets every seeded module. A whole-repo grep
  for didClose outside node_modules and engine/dist returns four hits: engine/README.md:36 (`|
  module/didClose | Drops per-document incremental state. |`), the case, and the two comment lines. No
  .cs, .ts, .mjs or .ps1 file sends it.
- **Why:** engine/README.md advertises it as part of the protocol, so the next person wiring up tab-close
  plumbing has a documented hook that duplicates what closeProject already does correctly and at the
  right moment. Two ways to forget a document, one of them untested and unreachable.
- **Change:** Delete the case and the README row, and move the closeProject comment's explanation of why didClose
  is not sent up to where the reader looks for it. If it is kept deliberately as a future hook, say so
  in the README row rather than describing it as live protocol.
- **Size:** 3 lines in dispatcher.ts plus one README row
- **Adversary:** engine/src/dispatcher.ts:286-288 is the case; the closeProject comment at 448-457 states the product
  has never sent it and explains why the release happens in closeProject instead. Repo-wide grep for
  didClose outside node_modules and engine/dist returns exactly engine/README.md:36 (a protocol table
  row), dispatcher.ts:286, and the two comment lines at 450 and 456. No .cs, .ts, .mjs or .ps1 sends
  it.

##### `engine-test-scaffolding-x4` All four engine tests carry their own copy of the same 78-line spawn, connect, frame and call scaffolding

- **Where:** `engine/test/positions.mjs:23`
- **Kind:** complexity / medium effort, claim derived, confidence verified, severity low
- **Evidence:** positions.mjs lines 23-100 are imports, `const here/dist`, `const pipeName =
  \`xlide-engine-positions-${process.pid}\``, the spawn with `stdio: ['ignore','pipe','pipe']`, the
  stderr drain, the await-for-'listening' promise with a 30s reject,
  `net.connect(\`\\\\.\\pipe\\${pipeName}\`)`, the `nextId`/`pending`/`inbox` newline framer,
  `function call(method, params)` with its 30s timeout, and `let passed = 0; const failures = [];
  function check(what, body)`. I diffed that block against the equivalent block in each sibling with
  difflib: freshness.mjs (lines 22-99) 0.997, language.mjs (lines 26-112) 0.895, smoke.mjs (lines
  7-80) 0.846. Only the pipe name suffix and, in smoke/language, an optional `--exe` branch differ. No
  shared module exists under engine/test.
- **Why:** About 300 lines of transport that exist four times. A protocol change - a second framing character,
  a different readiness line on stdout, a longer connect budget - is four edits, and the copies have
  already drifted: smoke.mjs and language.mjs accept `--exe` so they can be run against the packaged
  executable, positions.mjs and freshness.mjs hardcode `process.execPath` with engine.cjs and cannot
  be pointed at the thing the add-in actually launches.
- **Change:** Extract engine/test/harness.mjs exporting `startEngine({ label, useExe })` returning `{ call, stop
  }` plus the `check`/`passed`/`failures` reporter, and have the four tests import it. The --exe
  branch then reaches all four for free.
- **Size:** about 300 duplicated lines across 4 files
- **Adversary:** engine/test/positions.mjs:23-50 and engine/test/freshness.mjs:22-49 are the same imports, here/dist,
  pipeName template, spawn(process.execPath, [join(dist,'engine.cjs'),'--pipe',pipeName]), stderr
  drain, listening promise with the same 30_000 reject text, net.connect and framer, differing only in
  the pipe label. smoke.mjs:15-19 and language.mjs:34-38 carry a useExe branch selecting
  dist/xlide-engine.exe that positions.mjs:33 and freshness.mjs:33 do not, so the drift claim holds.
  Each file defines its own call() (smoke:68, language:84, positions:77) and check() (smoke:124,
  language:98, positions:91). No shared module exists in engine/test (ls shows only the four tests).

##### `mjs-suite-epilogue-duplication` Nine copies of the same check helper, eight of the same finally-epilogue, and nine local wait definitions shadowing the one xlide-api.mjs exports

- **Where:** `tools/harness/colouring.mjs:30`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity low
- **Evidence:** colouring.mjs:28-33 is `let passed = 0; const failures = []; const check = (what, ok, detail) => {
  if (ok) { passed += 1; console.log(\`ok ${what}\`); } else { failures.push(what); console.log(\`FAIL
  ${what}${detail ? \`\\n ${detail}\` : ""}\`); } };`. The identical `const check = (what, ok, detail)
  => {` line opens in nine files: colouring, debugger-features, import-guard, language-features,
  menu-bar, rename-features, settings-bite, step-into-features, write-rollback. The epilogue
  `console.log(\`\\n${passed} passed, ${failures.length} failed\`)` followed by the failure loop and
  `process.exitCode = failures.length === 0 ? 0 : 1;` appears in eight: analysis-freshness:344,
  colouring:200, com-leak:432, format-positions:435, immediate-watch:197, import-guard:108,
  settings-bite:152, three-copies:246. Four of those (colouring:198-202, import-guard:106-110,
  settings-bite:150-155, and the same shape in three-copies/format-positions) also repeat the
  identical teardown `await api.pane("close", { module: name, project: project.projectId, answer:
  "discard" }).catch(() => {}); await api.component("remove", { name, project: project.projectId
  }).catch(() => {});`. And `xlide-api.mjs:110` exports `export const wait = (ms) => ...` while nine
  suites define their own: com-leak:35, debugger-features:20, immediate-watch:25,
  language-features:25, language-live-probe:24, perf-scaling:24, rename-features:25,
  step-into-features:17, surface-walk:44.
- **Why:** verify.ps1:434-443 parses each suite's output with `Select-String '(\d+) passed, (\d+) failed'` and
  `Select-String '^\s*FAIL'`, so the gate's ability to read a verdict depends on nine hand-copied
  console.log formats staying byte-identical. Two suites already differ: module-sync.mjs uses `let
  failed = 0` (line 26) and write-rollback.mjs uses `const passed = []; const failed = []` (lines
  56-57), so a change to the reporter has to be made in three shapes, not one. The scratch-module
  teardown is the other risk: a suite that copies the block and forgets the `.catch(() => {})` turns a
  failed assertion into a teardown stack trace, which is exactly what three-copies.mjs:243 and
  analysis-freshness.mjs:333 wrote comments to warn about.
- **Change:** xlide-api.mjs already absorbed wait/waitFor/waitUntilStable; give it the rest of the same layer.
  Export a `reporter()` returning `{ check, done }` where done prints the one verdict line verify.ps1
  greps for and sets exitCode, and a `scratchModule(api, project, name)` returning a disposable that
  does the add plus the pane-close/component-remove teardown. Then delete the nine local `wait`
  consts.
- **Size:** 9 check copies, 8 epilogue copies, 9 wait copies, 4+ teardown copies
- **Adversary:** grep -l over tools/harness/*.mjs: the exact line 'const check = (what, ok, detail) =>' opens
  colouring, debugger-features, import-guard, language-features, menu-bar, rename-features,
  settings-bite, step-into-features, write-rollback (9). The '${failures.length} failed' epilogue
  appears in analysis-freshness, colouring, com-leak, format-positions, immediate-watch, import-guard,
  settings-bite, three-copies (8). Nine local wait consts (com-leak:35, debugger-features:20,
  immediate-watch:25, language-features:25, language-live-probe:24, perf-scaling:24,
  rename-features:25, step-into-features:17, surface-walk:44) shadow the exported one at
  xlide-api.mjs:110. The pane-close plus component-remove teardown pair appears verbatim in
  colouring:196-197, import-guard:102-103, settings-bite:148-149.
- **Correction applied:** The stated failure mode is overstated: all three reporter shapes still emit the line verify.ps1:439
  greps (module-sync.mjs:450 '${passed} passed, ${failed} failed', write-rollback.mjs:128
  '${passed.length} passed, ${failed.length} failed', colouring.mjs:200). Nothing is currently
  unparseable; the finding is duplication plus latent divergence, not a live gate defect.

##### `fixture-generator-driver-x4` The four New-*Fixture.ps1 generators repeat the same three-phase driver verbatim

- **Where:** `tools/New-RenameFixture.ps1:211`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity low
- **Evidence:** New-RenameFixture.ps1 lines 208-255: `$harness = Join-Path $PSScriptRoot 'harness'`, `Write-Host '1.
  Making an empty macro workbook.'`, the `New-Object -ComObject Excel.Application` / `SaveAs($Path,
  52)` / `finally { $maker.Quit(); ReleaseComObject }` block, `Write-Host '2. Opening it with the
  editor...'` calling `Start-Excel.ps1 -Workbook $Path -Fresh`, the `$plan = @{ modules = @(...);
  sheetCode = ...; openAtEnd = ... }` hashtable, the BOM-free
  `[System.IO.File]::WriteAllText($planPath, ($plan | ConvertTo-Json -Depth 5), (New-Object
  System.Text.UTF8Encoding $false))`, and `& node (Join-Path $harness 'build-fixture.mjs') $planPath`
  in a try/finally that removes the temp file. The identical sequence starts at
  New-DebugFixture.ps1:116, New-LanguageFixture.ps1:193 and New-PerfFixture.ps1:107 - grep for the
  literal `Write-Host '1. Making an empty macro workbook.'` matches all four - and the surrounding
  lines match on `New-Object -ComObject Excel.Application`, `SaveAs($Path, 52)`, `Start-Excel.ps1`,
  `ConvertTo-Json -Depth 5`, `build-fixture.mjs`. Only the temp filename differs (xlide-fixture /
  xlide-debug-fixture / xlide-language-fixture / xlide-perf-fixture) and New-LanguageFixture.ps1:187
  admits the copy: "Three phases, exactly as New-RenameFixture.ps1 does them and for the same
  reasons".
- **Why:** About 40 lines x 4. The BOM comment above the WriteAllText exists because PowerShell 5.1's
  `-Encoding utf8` writes a BOM that JSON.parse refuses; that lesson is now recorded four times and
  has to be preserved four times. A fifth fixture is a copy-paste, and a change to how
  build-fixture.mjs is invoked is four edits with no test that would catch missing one.
- **Change:** Move the driver into a shared `tools/New-Fixture.ps1` (or a dot-sourced tools/FixtureDriver.ps1)
  taking `-Path`, `-Modules`, `-SheetCode`, `-OpenAtEnd`, leaving each New-*Fixture.ps1 as its module
  sources plus one call. The VBA bodies are the part that genuinely differs and should stay per-file.
- **Size:** about 160 duplicated lines across 4 generators
- **Adversary:** The literal "1. Making an empty macro workbook." appears at New-DebugFixture.ps1:116,
  New-LanguageFixture.ps1:193, New-PerfFixture.ps1:107 and New-RenameFixture.ps1:211, each followed
  within about forty lines by the same SaveAs($Path, 52) block (122/200/113/218), the same
  Start-Excel.ps1 -Workbook $Path -Fresh call (131/209/122/227), the same ConvertTo-Json -Depth 5
  written through UTF8Encoding $false (146/226/137/246) and the same build-fixture.mjs invocation in
  try/finally (151/231/142/251). I read New-RenameFixture.ps1:200-256 in full and the BOM comment is
  there as described.

##### `powershell-discovery-x8` Eight PowerShell scripts each rebuild the debug-api discovery path, token and base URL by hand, and they do not agree on how to pick the host

- **Where:** `tools/harness/Test-DebugApi.ps1:20`
- **Kind:** complexity / medium effort, claim observed, confidence verified, severity low
- **Evidence:** `$discoveryPath = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"` then `$api
  = "http://127.0.0.1:$($d.port)/$($d.token)"`. The same two lines, with the same string literal,
  appear in Get-Shot.ps1:47, Open-VbeIn.ps1:103, Test-Churn.ps1:34, Test-DiscardProblems.ps1:36,
  Test-ObjectBrowser.ps1:183, Test-SplitWorkspace.ps1:23, tools/page.ps1:70 and tools/verify.ps1:352.
  Six of them get the pid from `Get-Process EXCEL | Select-Object -First 1`; Open-VbeIn.ps1 takes a
  `$ProcessId` parameter; page.ps1:66-70 enumerates all Excels and refuses to guess ("Never guesses
  between two Excels: a page deployed into the wrong session is a confusing hour"). The node client
  does none of this by hand: xlide-api.mjs:180-217 `discover()` reads every debug-api-*.json in the
  directory, proves each is alive by calling state(1500), and `open({ pid, workbook })` at line 224
  "Throws with the list when the choice is ambiguous, because guessing which Excel to drive is how a
  test writes into the wrong workbook."
- **Why:** The node side already learned that first-Excel-wins writes into the wrong workbook, and six
  PowerShell probes still do exactly that - including four the gate runs (Test-DebugApi,
  Test-SplitWorkspace, Test-DiscardProblems, Test-Churn). None of them proves the discovery file
  belongs to a live process, so a stale file left by a killed Excel points them at a dead port. Any
  change to the discovery file's name, location or fields is eight edits.
- **Change:** Add tools/harness/XlideApi.psm1 exporting `Get-XlideApi [-ProcessId] [-Workbook]` that mirrors
  discover()/open(): enumerate the directory, prove liveness with a /state call, and throw with the
  candidate list rather than picking the first. Convert the eight call sites. Do not change route
  shapes - this is only how the base URL is found.
- **Size:** 8 hand-rolled copies; 6 of them pick the wrong Excel when two are open
- **Adversary:** grep for 'debug-api-$' across tools/*.ps1 returns nine sites, all building the same
  LOCALAPPDATA\xlide_vbide\debug-api-<pid>.json path: Get-Shot.ps1:47, Open-VbeIn.ps1:103,
  Test-Churn.ps1:34, Test-DebugApi.ps1:20, Test-DiscardProblems.ps1:36, Test-ObjectBrowser.ps1:183,
  Test-SplitWorkspace.ps1:23, page.ps1:70, verify.ps1:352. Six take the pid from 'Get-Process EXCEL
  ... Select-Object -First 1' (Get-Shot:44, Test-Churn:31, Test-DebugApi:17, Test-DiscardProblems:33,
  Test-SplitWorkspace:20, verify.ps1:340). page.ps1:66-90 enumerates all Excels, proves liveness with
  /state and throws on ambiguity; xlide-api.mjs:179-217 discover() and 224+ open() do the same on the
  node side.
- **Correction applied:** Nine sites, not eight: tools/verify.ps1:352 is the ninth. Also the stale-discovery-file scenario is
  weak: all six derive the pid from a currently running EXCEL process, so a corpse file only misleads
  after pid reuse. The real defect is first-Excel-wins plus no /state liveness proof, and
  Test-DebugApi.ps1:21-24 at least fails loudly when no file exists for the chosen pid.

##### `unrun-wrapper-probes-lose-legs` Test-CloseConfirm.ps1 and Test-ObjectBrowser.ps1 are three-leg wrappers the gate never calls, so their seam legs and their only live/engine legs run nowhere

- **Where:** `tools/harness/Test-CloseConfirm.ps1:31`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity medium
- **Evidence:** Test-CloseConfirm.ps1's header names three legs: "1. Seams... 2. Page behaviour -
  close-confirm-page-probe.mjs... 3. Engine behaviour - engine-live-probe.mjs walks the built engine
  over its own pipe through the stale-problems story (2026-08-05)". It defines `function Test-Seam` at
  line 31 and calls it eight times (lines 51-80) across ui/editor/src/shell.ts, bridge.ts,
  EditorMessages.cs, EditorSurface.cs and AddInSession*.cs. verify.ps1:247 runs
  `close-confirm-page-probe.mjs` directly and never runs the wrapper, so legs 1 and 3 execute nowhere.
  A repo-wide grep for engine-live-probe.mjs finds exactly one invoker, Test-CloseConfirm.ps1.
  Test-ObjectBrowser.ps1 has the same shape (header legs 1/2/3, leg 3 being objbrowser-live-probe.mjs)
  and verify.ps1:247 likewise runs only objbrowser-page-probe.mjs; grep finds
  objbrowser-live-probe.mjs invoked only from Test-ObjectBrowser.ps1 and mentioned in
  docs/debug-api.md. Neither wrapper appears in verify.ps1's probe list at line 369.
- **Why:** The stale-deploy tripwire is gone: docs/testing.md:15-18 keeps seam checks specifically for "a
  rebuilt page that never reached the publish tree has bitten more than once", and the only Test-Seam
  implementation in the repo is in a file nothing runs. So is the engine leg of the stale-problems
  story, which docs/testing.md:38-41 uses as the worked example of a probe that could not tell two
  builds apart. The gate looks like it covers close-confirm and the object browser; it covers one
  third of each.
- **Change:** Either add Test-CloseConfirm.ps1 and Test-ObjectBrowser.ps1 to verify.ps1's live probe list at line
  369 (they emit the RESULT: PASS line that loop already parses), or promote the two legs the gate is
  missing - run engine-live-probe.mjs beside the page probes at verify.ps1:247, since it needs no
  Excel, and move Test-Seam into a step of its own.
- **Size:** 8 seam checks plus 2 live probes running nowhere
- **Adversary:** Test-CloseConfirm.ps1:7-15 names the three legs; Test-Seam is defined at line 31 and called at 51,
  53, 55, 57, 59, 61, 67, 80, 83 and 88 (ten calls, including the published-bundle stale-deploy check
  at 88), and Invoke-NodeProbe 'engine' 'engine-live-probe.mjs' is at line 125. verify.ps1:247 runs
  close-confirm-page-probe.mjs directly and verify.ps1:369 lists only Test-DebugApi,
  Test-SplitWorkspace, Test-DiscardProblems, Test-Churn, so neither wrapper is ever invoked. grep
  shows engine-live-probe.mjs invoked only from Test-CloseConfirm.ps1 and objbrowser-live-probe.mjs
  only from Test-ObjectBrowser.ps1 (plus a docs/debug-api.md:355 mention).
- **Correction applied:** There are two Test-Seam implementations, not one: Test-CloseConfirm.ps1:31 and
  Test-ObjectBrowser.ps1:29. Both files are unrun, so the conclusion that no seam check executes in
  the gate stands. Test-CloseConfirm makes ten seam calls, not eight.

##### `superseded-and-orphan-harness-scripts` Five harness scripts have no runner and no document: one is superseded, four are the only cover their behaviour has

- **Where:** `tools/harness/rename-features.mjs:1`
- **Kind:** dead-code / small effort, claim derived, confidence verified, severity medium
- **Evidence:** For each of the fifteen tools/harness scripts verify.ps1 does not run, I grepped docs/, README.md,
  tools/, engine/ and .github/ for the filename. Five have zero references anywhere outside
  themselves: rename-features.mjs, Test-CloseVbe.ps1, Test-CloseHiddenPane.ps1, Test-ResizeFollow.ps1,
  Get-EditorScreenshot.ps1. rename-features.mjs IS superseded - its header says "Rename and Go to
  Definition, driven through the providers... Run against the rename fixture" and three-copies.mjs,
  which verify.ps1:426 runs on RenameFixture.xlsm, drives rename and `api.undoRename()` on the same
  fixture with the workbook/surface/engine comparison rename-features.mjs does not have.
  Get-EditorScreenshot.ps1 (16 KB) is superseded by Get-Shot.ps1 (2.8 KB), which does the same job
  through the api's /capture route and is the one docs/debug-api.md names. The other three are not
  superseded: Test-CloseVbe.ps1 closes the frame via WM_SYSCOMMAND SC_CLOSE three times and checks
  Excel survives ("Guards the 2026-08-04 crash (lesson 27)"), Test-CloseHiddenPane.ps1 checks a hidden
  pane's close still removes its tab, Test-ResizeFollow.ps1 checks the overlay and the browser child
  both follow a frame resize. `grep -l resize *.mjs` and `grep -l
  'closeFrame|SC_CLOSE|frameVisible|hidePane' *.mjs` in tools/harness both return nothing.
- **Why:** Two files are ballast that a reader has to disprove. The other three are worse: they are the only
  checks in the repo for three named, already-shipped crashes and layout faults, and nothing runs
  them, so the guard reads as present and is not. docs/testing.md:24 tells the reader that
  `tools\harness\Test-*.ps1` are "gated behind verify.ps1 -Live", which is true of four of the
  fifteen.
- **Change:** Delete rename-features.mjs (three-copies.mjs covers it) and Get-EditorScreenshot.ps1 (Get-Shot.ps1
  covers it). For Test-CloseVbe.ps1, Test-CloseHiddenPane.ps1 and Test-ResizeFollow.ps1, decide
  explicitly: add them to verify.ps1:369 - they already print PASS/FAIL, though the loop needs the
  RESULT: line they may not emit - or record in docs/testing.md that they are hand-run crash
  reproductions, so the gap is a decision rather than an accident.
- **Size:** 2 files superseded, 3 guards for shipped crashes running nowhere
- **Adversary:** grep across *.md, *.ps1, *.mjs, *.yml, *.cs for the five names returns only self-references
  (Get-EditorScreenshot.ps1:29,33 in its own examples; rename-features.mjs:6 in its own usage line)
  and nothing at all for Test-CloseVbe.ps1, Test-CloseHiddenPane.ps1, Test-ResizeFollow.ps1.
  Test-CloseVbe.ps1:1-17 does close the frame by WM_SYSCOMMAND and cite the 2026-08-04 crash, and grep
  -l over tools/harness for resize or SC_CLOSE/closeFrame/frameVisible/hidePane returns no .mjs at
  all, so those three behaviours are covered nowhere else.
- **Correction applied:** The supersession half is wrong. rename-features.mjs is NOT covered by three-copies.mjs:
  three-copies.mjs:224-232 does one rename plus undoRename with three-copy agreement, while
  rename-features.mjs:105, 109, 112 and 132 check the ambiguous bare call being left alone, Rival's
  own Recalculate untouched, the prefix-sharing HelpersExtra untouched, and Go to Definition from a
  call site, and grep for Rival/HelpersExtra finds those cases only in rename-features.mjs,
  three-copies.mjs, format-positions.mjs, language-live-probe.mjs and Test-Language.ps1 fixture text.
  Deleting it would drop unique coverage. Get-EditorScreenshot.ps1:1-34 also launches and closes its
  own host, which Get-Shot.ps1 (attach-to-running, /capture) does not do. Finally, there are twelve
  Test-*.ps1 files, not fifteen, so docs/testing.md:24 is true of four of twelve.

##### `page-loop-two-scripts` tools/page.ps1 and tools/Update-Page.ps1 are the same page loop written twice, and the operational guide points at the older one

- **Where:** `tools/Update-Page.ps1:38`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** Update-Page.ps1 lines 38-65: `$publish = Join-Path $repoRoot
  "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64"`, `$served =
  Join-Path $publish 'ui\editor\dist'`, npm run build unless -NoBuild, Copy-Item into $served, then `&
  node (Join-Path $repoRoot 'tools\harness\reload-page.mjs')`. page.ps1 lines 53-56 computes the same
  target - `$publishRoot = Join-Path $repoRoot
  "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64\ui\editor\dist"` -
  and its synopsis is "Rebuild the editor page and put it into the RUNNING editor, without restarting
  anything" against Update-Page.ps1's "Builds the editor page and puts it in front of a running
  editor, without restarting the host". page.ps1 is the superset: -Watch, -Reset, -NoTypecheck, and
  its own Get-LiveApi that refuses to guess between two Excels. Both are current: docs/status.md:57
  and README.md:100 name page.ps1; docs/driving-excel.md:134 and :139 name Update-Page.ps1 in the
  what-to-run-after-changing-what table. Nothing in verify.ps1, dev.ps1 or CI runs either.
- **Why:** docs/driving-excel.md is the operational guide and is under a standing rule to stay current with the
  harness; its change-to-command table sends the reader to the script without the typecheck, without
  -Watch, and with the reload path that does not disambiguate between two running Excels. Two scripts
  computing the same publish path also means a change to the publish layout silently fixes one loop
  and breaks the other.
- **Change:** Delete tools/Update-Page.ps1 and repoint docs/driving-excel.md:134 and :139 at tools\page.ps1,
  mapping -NoBuild to -NoDeploy/-NoTypecheck as appropriate. If the thin version is wanted for
  scripting, make it `page.ps1 -NoTypecheck -NoReload` rather than a second file.
- **Size:** 66-line script duplicating a 200-line one; one stale doc row
- **Adversary:** tools/Update-Page.ps1:41-42 computes artifacts\publish\Xlide.Vbe.Shim\<config>_win-x64 then
  ui\editor\dist, builds unless -NoBuild (44-48), copies (55-56) and shells reload-page.mjs (64).
  tools/page.ps1:53-56 computes the same served path and page.ps1:16-28 documents -Reset, -Watch,
  -NoDeploy. docs/status.md:57 and README.md:100 name page.ps1; docs/driving-excel.md:134 and :139
  name Update-Page.ps1. grep finds no runner for either in verify.ps1, dev.ps1 or CI.
- **Correction applied:** The disambiguation criticism is refuted: Update-Page.ps1 reloads through
  tools/harness/reload-page.mjs, which uses discover() and deliberately reloads every live editor
  (reload-page.mjs:1-3, 16-33), so it does not pick a wrong single Excel; and both scripts write to
  the same fixed publish path, so neither disambiguates the deploy target. The real difference is
  page.ps1's typecheck, -Watch and -Reset, plus the duplicated publish-path computation and the split
  doc pointers.

##### `twinfixture-has-no-generator` surface-walk.mjs and the driving guide both require TwinFixture.xlsm, and no generator makes it

- **Where:** `tools/harness/surface-walk.mjs:21`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** surface-walk.mjs:20-22: "Run it against a session with TWO workbooks open, or the interesting half
  does not exist:" followed by a Start-Excel.ps1 line naming RenameFixture.xlsm and TwinFixture.xlsm,
  and the suite's own reason at lines 15-18: "A run that never holds two modules of the same name
  passes every label check vacuously... That line is how the `pane` route's dropped project argument
  was found." docs/driving-excel.md:618 repeats the same command and lines 346, 371 and 372 use
  TwinFixture.xlsm in worked examples. tools/ holds New-DebugFixture.ps1, New-LanguageFixture.ps1,
  New-PerfFixture.ps1 and New-RenameFixture.ps1 and nothing else matching New-*Fixture; a repo-wide
  grep for TwinFixture across .ps1, .mjs and .md finds only those five consumer mentions plus a
  comment in xlide-api.mjs:1167.
- **Why:** The one suite written to catch cross-workbook ambiguity - the class of defect that produced the
  dropped project argument on the pane route - cannot be run as documented by anyone who does not
  already have the file, and it is not in artifacts/. Run without a twin it still exits zero,
  reporting collision=0, which the header calls out as the vacuous pass it was built to avoid.
- **Change:** Add tools/New-TwinFixture.ps1 producing a workbook whose module names deliberately collide with
  RenameFixture.xlsm's (Helpers, HelpersExtra), using the same three-phase driver the other four use.
  If the twin is meant to be a second copy of the rename fixture, say that in surface-walk.mjs's
  header and give New-RenameFixture.ps1 a -Path default that makes the copy obvious.
- **Size:** 1 suite and 4 doc examples blocked on a missing fixture
- **Adversary:** ls tools/*.ps1 shows New-DebugFixture, New-LanguageFixture, New-PerfFixture and New-RenameFixture
  and no twin generator. surface-walk.mjs:15-21 states the two-workbook requirement and the
  collision=0 vacuous pass, and docs/driving-excel.md:346, 371, 372, 618 use TwinFixture.xlsm.
  .gitignore:2 ('artifacts/') and :55 ('artifacts/fixtures/') exclude the fixture tree and 'git
  ls-files artifacts' is empty, so no clone can obtain it.
- **Correction applied:** artifacts/fixtures/TwinFixture.xlsm does exist on this machine (ls of artifacts/fixtures lists
  Debug, Language, Perf, Rename and Twin), so the suite is runnable here; the defect is that the file
  is untracked and unreproducible, not that it is absent everywhere.

##### `formfeed-in-runnable-command` A literal form feed byte sits where a backslash-f belongs in the driving guide's two-workbook command, so the sample cannot be copied and run

- **Where:** `docs/driving-excel.md:618`
- **Kind:** complexity / small effort, claim observed, confidence verified, severity low
- **Evidence:** The raw bytes of docs/driving-excel.md line 618 are `tools\harness\Start-Excel.ps1 -Workbook
  artifacts\x0cixtures\RenameFixture.xlsm,artifacts\x0cixtures\TwinFixture.xlsm` - 0x0C where `\f` of
  `artifacts\fixtures` was intended. tools/harness/surface-walk.mjs line 21 carries the identical
  corruption. A scan of every .md, .mjs, .ps1, .cs, .ts, .yml and .json in the repo (excluding .git,
  node_modules, dist, obj, bin, artifacts) for 0x0C finds these two files and no others, two
  occurrences each.
- **Why:** docs/driving-excel.md is under a standing rule that its samples are meant to be run rather than
  read, and this one silently resolves to a path that does not exist. It also renders as
  `artifactsixtures` in most viewers and in grep output, so a reader searching for the fixtures
  directory will not find this line, and someone fixing the path in the doc will not know the same
  bytes are in surface-walk.mjs.
- **Change:** Replace the 0x0C byte with a literal backslash in both files. It almost certainly came from a shell
  or editor interpreting `\f`, so whatever writes these paths should emit them single-quoted.
- **Size:** 4 bytes across 2 files
- **Adversary:** od -c of docs/driving-excel.md line 618 shows 'a r t i f a c t s \f i x t u r e s' twice, i.e. byte
  0x0C where a backslash plus f belongs. cat -v of tools/harness/surface-walk.mjs line 21 shows
  'artifacts^Lixtures' twice. A grep -lP for \x0c across .md/.mjs/.ps1/.cs/.ts/.yml/.json outside
  node_modules and dist matches exactly these two files.

##### `initialize-handshake-discarded` The engine's initialize handshake answers a protocol version and the shim throws the answer away

- **Where:** `src/Xlide.Vbe.Shim/Engine/EngineClient.cs:114`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** `await CallAsync("initialize", new Dictionary<string, object>(),
  cancellation).ConfigureAwait(false);` - the result is awaited and not assigned, and it is the last
  statement of the connect path. The engine answers something worth reading:
  engine/src/dispatcher.ts:272-274, `case 'initialize': this.initialized = true; return { engine:
  'xlide', protocol: 1 };`. Grepping src/ for the strings `"initialize"` and for any read of an
  `engine` or `protocol` field off an engine response finds nothing beyond this one call site.
- **Why:** verify.ps1:94-102 exists because of exactly the failure this field would catch: "an engine change
  can be built, tested against the bundle, committed, and published, while the thing that actually
  runs is hours old and refuses every new method as unknown... That happened on 2026-08-06, and the
  only reason it surfaced was a live session's log." The gate's timestamp check only covers the
  developer's own machine; an installed user running a mismatched engine gets MethodNotFound per
  feature with nothing naming the cause. The handshake that would name it in one line is already on
  the wire and is discarded.
- **Change:** Read the result, compare `protocol` against a constant in the shim, and log at Warn (or refuse to
  start the engine) when it does not match, naming both numbers. Bump the engine's `protocol` whenever
  a method is added or a shape changes. This is additive to an existing response, not a new route or a
  rename.
- **Size:** 1 discarded response carrying the version field
- **Adversary:** src/Xlide.Vbe.Shim/Engine/EngineClient.cs:114 is 'await CallAsync("initialize", new
  Dictionary<string, object>(), cancellation).ConfigureAwait(false);' and it is the last statement of
  the connect path (the method ends at 115). engine/src/dispatcher.ts:272-274 answers { engine:
  'xlide', protocol: 1 }. grep for 'protocol' across src/*.cs finds only unrelated comments and
  EngineProtocol.cs serialisation, so nothing reads the field.
- **Correction applied:** The fix is bigger than reading a field. protocol is a hard-coded literal 1 at dispatcher.ts:274 that
  has never been bumped, and grep for engineVersion/engineBuild/EngineVersion across src/ and
  engine/src returns nothing, so there is no version or build stamp anywhere on the wire today.
  Comparing protocol against a shim constant would catch a stale engine only once the engine starts
  bumping it; answering a build stamp from the packaged exe would catch the 2026-08-06 case directly.

##### `release-skipgate-drops-engine-freshness` release.ps1 -SkipGate removes the only check that the installer's engine is not stale, and nothing else in the release path re-checks it

- **Where:** `tools/release.ps1:53`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** `if (-not $SkipGate) { ... & (Join-Path $PSScriptRoot 'verify.ps1') ... }` at lines 53-57, then
  installer\build.ps1 at line 60 and `gh release upload $Tag $installer --clobber` at line 74. The
  engine freshness check lives only in verify.ps1's 'engine executable is current' step (lines
  94-163), which compares engine\dist\xlide-engine.exe's LastWriteTimeUtc against both engine\src and
  the neighbouring xlide_vscode\src checkout and repackages when anything is newer.
  installer/build.ps1 does not repeat it: greping it for engine shows only existence checks - line 78
  `Get-ChildItem $enginePublish -File -Filter '*.exe'` and line 84 `if ($engine.Count -eq 0 -and -not
  $WithoutEngine) { throw ... }`. It asks whether an engine exists, never how old it is. release.ps1
  also never hashes or otherwise identifies the exe it just uploaded; its final report prints only
  asset names and sizes from `gh release view` (lines 81-85).
- **Why:** -SkipGate is documented for the ordinary case ("for when it has already been run in this session"),
  and a session where the gate ran an hour ago and the analyzer checkout was pulled since is exactly
  the shape of the 2026-08-06 incident the freshness step was written for. The result ships to users,
  and the release report gives nobody a way to tell after the fact which engine went out.
- **Change:** Move the freshness comparison into a small function verify.ps1 and release.ps1 both call, and run it
  in release.ps1 unconditionally, after -SkipGate. Add the uploaded exe's SHA-256 and its
  LastWriteTimeUtc to the closing report so the shipped engine is identifiable from the release page.
- **Size:** 1 skippable gate step standing between a stale engine and an upload
- **Adversary:** tools/release.ps1:53-57 gates verify.ps1 behind -not $SkipGate, then runs installer\build.ps1 at 60
  and 'gh release upload $Tag $installer --clobber' at 74; the closing report at 81-85 prints only
  asset name and MB from gh release view. The freshness comparison exists only at verify.ps1:94-162.
  installer/build.ps1:76-91 only asks whether any *.exe exists in engine\dist and throws when none
  does; nothing there compares timestamps.
- **Correction applied:** verify.ps1's step is not only a check: at lines 139-162 it repackages the engine when any source is
  newer, refusing first if EXCEL or xlide-engine holds the exe. So -SkipGate skips an automatic
  repackage as well as the comparison, which makes the gap slightly worse than stated.

##### `testing-doc-misdescribes-the-live-gate` docs/testing.md describes the live gate as the Test-*.ps1 probes and never mentions the ten .mjs suites that are most of it

- **Where:** `docs/testing.md:24`
- **Kind:** api-coverage / small effort, claim derived, confidence verified, severity low
- **Evidence:** Line 24: "**Live probes** (`tools\harness\Test-*.ps1`, gated behind `verify.ps1 -Live`). Drive a
  real editor through the debug api." verify.ps1:369 runs four of the fifteen Test-*.ps1 files. The
  bulk of -Live is the step at verify.ps1:395-443, whose own comment opens "THE NODE SUITES, which
  existed and passed and which nothing ran", running import-guard.mjs, immediate-watch.mjs,
  analysis-freshness.mjs, menu-bar.mjs, module-sync.mjs twice, format-positions.mjs, three-copies.mjs,
  colouring.mjs and settings-bite.mjs (lines 420-427). docs/testing.md names none of them; its only
  .mjs mentions are the `*-page-probe.mjs` family (line 20) and com-leak.mjs (line 67). Line 15
  documents "Seam checks (`Test-Seam` in the harnesses)" as one of the four kinds, and the sole
  Test-Seam definition is tools/harness/Test-CloseConfirm.ps1:31, which nothing runs.
- **Why:** This is the document that answers "which kind of check should I write". It points a reader at the
  PowerShell probe pattern for live work when the project moved to node suites against xlide-api.mjs,
  and it presents seam checks as a live kind when their only implementation is unreachable. A new
  suite written from this document arrives in the wrong language, in a file the gate does not run,
  with a verdict line verify.ps1's parser does not recognise.
- **Change:** Add a fifth kind for the node suites - what they are, that they report `N passed, M failed` because
  verify.ps1 parses that, that they import open/waitFor from xlide-api.mjs and must put the fixture
  back in a finally - and correct line 24 to say which four Test-*.ps1 the gate runs. Fix or remove
  the seam-checks paragraph depending on what happens to Test-CloseConfirm.ps1.
- **Size:** 10 of the 14 live-gate suites undocumented
- **Adversary:** docs/testing.md:24 reads '**Live probes** (`tools\harness\Test-*.ps1`, gated behind `verify.ps1
  -Live`)'. verify.ps1:369 runs four of the twelve Test-*.ps1 files, while verify.ps1:395-443 runs ten
  node suite invocations (import-guard, immediate-watch, analysis-freshness, menu-bar, module-sync
  twice, format-positions, three-copies, colouring, settings-bite) and verify.ps1:464-478 runs
  com-leak.mjs. grep for 'mjs' across the 155 lines of docs/testing.md returns only line 20
  (*-page-probe.mjs) and line 67 (com-leak.mjs). Line 15 presents seam checks as a live kind.
- **Correction applied:** Two counts. There are twelve Test-*.ps1 files, not fifteen, and the gate runs four of them.
  Test-Seam has two implementations, Test-CloseConfirm.ps1:31 and Test-ObjectBrowser.ps1:29, both in
  files the gate never invokes, so the seam-check paragraph describes a kind that currently never runs
  rather than one with a single unreachable implementation.

### The add-in's own cost

_7 findings, from the `perf-shim` finder._

##### `dispid-resolved-per-call` Every object-model property read costs two cross-COM calls, because the member name is resolved through GetIDsOfNames before every Invoke and no DISPID is ever cached

- **Where:** `src/Xlide.Vbe.Shim/Com/DispatchObject.cs:101`
- **Kind:** perf / small effort, claim observed, confidence verified, severity medium
- **Evidence:** GetProperty is the funnel every read goes through:
  
   public ComVariant GetProperty(string name)
   {
   var dispId = GetDispId(name); // line 101
   ...
   return InvokeCore(dispId, InvokeKind.PropertyGet, []);
   }
  
  and GetDispId is an unconditional round trip:
  
   var hr = _dispatch.GetIDsOfNames(NullGuid, (nint)namePointers, 1, 0, (nint)dispIds); // line 93
  
  There is no memo of any kind: `grep -n "GetDispId"` in DispatchObject.cs returns 16 call sites
  (lines 101, 165, 220, 237, 258, 295, 308, 328, 343, 372, 400, 420, 446, 462, 511) and every one
  calls straight through. A repo-wide grep for GetDispId/DispId outside DispatchObject.cs finds only
  XlideAddIn.cs's inbound IDispatch, so nothing caches on the caller's side either. GetItem is the
  worst case because it re-resolves a constant:
  
   public DispatchObject? GetItem(int index)
   {
   var dispId = GetDispId("Item"); // line 446, inside every loop iteration
   ...
   }
- **Why:** Every host-thread loop over a VBE collection pays a name lookup per item. FindComponent
  (AddInSession.cs:6969) over one 60-component project runs 60 x GetIDsOfNames("Item") + 60 x Invoke +
  60 x GetIDsOfNames("Name") + 60 x Invoke, so about 240 cross-COM calls where 122 would do.
  ReadOpenModules runs on every poll tick, PublishProjects walks every component of every project
  twice, and ProjectReader.ReadAll runs on the host thread on every analysis pass. All of it is on the
  thread that freezes the editor when it is busy.
- **Change:** Give DispatchObject a small per-instance `Dictionary<string,int>` populated on first resolve. It is
  exactly safe (same object, therefore same type, therefore same DISPIDs) and needs no reflection or
  codegen, so NativeAOT is unaffected. That alone removes the repeated GetIDsOfNames("Item") from
  every collection loop and every repeated read on a held object (Count, Name, Mode on the same
  object). A second tier keyed on the object's type info would also cover the
  fresh-object-per-iteration case (component.GetString("Name")); measure the per-instance tier first,
  since it is free.
- **Size:** Halves the cross-COM call count for repeated reads on one object; for the FindComponent example above, 240 calls to 181 with the per-instance cache alone (-25%), to about 122 (-49%) if a type-keyed tier is added. Per-call microsecond cost unmeasured - the shim has no counter for object-model calls.
- **Adversary:** src/Xlide.Vbe.Shim/Com/DispatchObject.cs:81-96 GetDispId is an unconditional
  IDispatch::GetIDsOfNames round trip with no memo; the class has only _pointer and _dispatch fields
  (lines 23-24), so no cache exists anywhere on the type. GetProperty:101, GetStringIndexed:165,
  SetProperty:511 and both GetItem overloads (446, 462) call straight through, and GetItem re-resolves
  the constant "Item" inside every loop iteration. The loops are real:
  src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:6976-7007 (FindComponent) does GetItem(j) +
  GetString("Name") per component, src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:162-183 does the same
  per component per pass, and AddInSession.cs:5781-5814 (ReadOpenModules) does it per pane on every
  poll tick. A repo-wide grep for GetIDsOfNames finds only the inbound dual-interface implementation
  in XlideAddIn.cs:310 and the typelib/UIA vtable declarations, so nothing caches on the caller side.
- **Correction applied:** The class doc at DispatchObject.cs:14-17 defends the design by claiming hot per-keystroke paths use
  early-bound interfaces instead. That is false for the VBE object model: the only early-bound
  wrapper, ComHandle<T>, is used solely for UI Automation (ImmediateReader.cs:29-30,
  LocalsReader.cs:39-40) and WebView2, never for a VBE member. Every VBE read in the product goes
  through DispatchObject. Also worth stating plainly, which the finding's own arithmetic does but its
  title does not: the per-instance cache removes only the repeated GetIDsOfNames("Item") on the
  collection object; the per-item GetString("Name") is on a fresh component object each iteration and
  is untouched by it (240 calls to 181 for the 60-component case, not to 122).

##### `keystroke-rebuilds-whole-module` A keystroke in a module over 64,000 characters scans and re-allocates the module's entire text on the VBE host thread

- **Where:** `src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1970`
- **Kind:** perf / medium effort, claim derived, confidence verified, severity high
- **Evidence:** The page sends only the change ranges once a module passes a threshold
  (ui/editor/src/bridge.ts:1636):
  
   if (fullLength < 64_000) {
   message.fullText = model.getValue();
   }
  
  so above 64,000 characters the shim reconstructs. contentChanged (EditorSurface.cs:1354) does:
  
   parsedEdits = ParseChanges(editedDoc.Text, changeSet);
   updated = parsedEdits is null ? null : ApplyEdits(editedDoc.Text, parsedEdits);
  
  ParseChanges opens with a full scan of the module (line 1970):
  
   var lineStarts = TextPositions.LineStarts(text);
  
  and LineStarts (Xlide.Vbe.Core/Engine/EngineProtocol.cs:255) walks every character and builds a
  List<int> then copies it to an array. ApplyEdits then allocates a whole new copy of the module per
  edit (line 2048):
  
   updated = string.Concat(updated.AsSpan(0, edit.Start), edit.Text, updated.AsSpan(edit.End));
  
  Nothing is cached between keystrokes: ParseChanges is static and takes the text fresh each time.
  This runs on the host UI thread - WebView2Surface.OnWebMessageReceived (line 811) calls the handler
  inline, and its own comment at line 837 says "this runs on the host user interface thread". The page
  does not debounce: onModelContentChanged posts on every Monaco change event.
- **Why:** The repo's own Massive fixture is 64,802 lines / about 1.42 MB (docs/lessons.md, finding 48/53).
  Typing one character in it costs, per keystroke and on the thread that draws the editor: a
  1.4-million-character scan, a ~65,000-entry List<int> grown and then copied to an int[] (about 520
  KB of garbage), and a 2.84 MB string copy that lands on the Large Object Heap. The 11,000-line perf
  fixture (~256 KB) is over the threshold too. This is precisely the latency the page comment at
  bridge.ts:1616 says it is avoiding by not sending full text - the cost was moved to the host thread
  rather than removed.
- **Change:** Hold the reconstruction state on OpenDoc instead of rebuilding it from scratch: cache the line-start
  table beside doc.Text and invalidate it when the text is replaced, and apply the bottom-up edits
  into a reusable char buffer (or a StringBuilder held per document) so one keystroke costs one splice
  rather than one full copy. Materialise the string only when something actually needs it - the
  write-back debounce, the live-analysis pass, HasUnwritten comparisons - not on every change message.
- **Size:** Per keystroke in the 1.42 MB fixture: ~1.4M char reads plus ~3.4 MB of allocation, of which 2.84 MB is LOH. Wall-clock unmeasured; PerfCounters has no counter on this path, which is why it has never shown up.
- **Adversary:** ui/editor/src/bridge.ts:1626-1638 attaches fullText only when getValueLength() < 64_000, and the
  poster is wired to model.onDidChangeContent at bridge.ts:491 with no debounce or throttle.
  src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1384-1390 therefore takes the reconstruct branch for
  anything larger: ParseChanges then ApplyEdits. ParseChanges opens with
  TextPositions.LineStarts(text) at EditorSurface.cs:1970, and LineStarts
  (src/Xlide.Vbe.Core/Engine/EngineProtocol.cs:255-270) scans every character into a List<int> and
  then copies it with [.. starts]. ApplyEdits at EditorSurface.cs:2048 allocates a fresh whole-module
  string per edit via string.Concat. Both are static and take the text fresh, so nothing survives
  between keystrokes. The thread is the host UI thread: WebView2Surface.OnWebMessageReceived calls
  handler(message) inline at src/Xlide.Vbe.Shim/WebView/WebView2Surface.cs:846 and its own comment at
  line 834 says so. Sizes check out: docs/lessons.md:1241 records the 64,802-line / 1.42 MB Massive
  module, and docs/lessons.md:1234 puts the 11,000-line fixture's last procedure near offset 256,000,
  both over the threshold.
- **Correction applied:** Line numbers: the contentChanged case begins at EditorSurface.cs:1357 (1354 is still inside
  selectionChanged); the reconstruct is at 1388-1389. The finding understates the case rather than
  overstating it: EditorSurface.cs:1460 shows the engine is fed only the parsed edits for a large
  module (full is null), so the shim's full-text rebuild is the one place in the whole per-keystroke
  path that still materialises the entire module, and 2.84 MB of LOH garbage per keystroke will force
  gen2 collections during typing, on the same thread.

##### `readall-before-the-skip` Every analysis pass reads every module's full source over COM on the host thread before the unchanged-sources skip can decide the pass is unnecessary

- **Where:** `src/Xlide.Vbe.Shim/Engine/AnalysisService.cs:797`
- **Kind:** perf / medium effort, claim observed, confidence verified, severity medium
- **Evidence:** AnalyseEverythingAsync's first act is the host-thread read (line 845):
  
   var snapshots = await ReadProjectsAsync(generation).ConfigureAwait(false);
  
  and ReadProjectsAsync marshals the whole read onto the host thread (line 797):
  
   read.TrySetResult(ProjectReader.ReadAll(_editor, generation));
  
  ProjectReader.ReadAll (ProjectReader.cs:39) loops every project, and Read (line 143) loops every
  component calling ReadSource, which pulls the whole module text across (line 210):
  
   return lines <= 0 ? string.Empty : code.GetStringIndexed("Lines", 1, lines) ?? string.Empty;
  
  Only afterwards, at AnalysisService.cs:889, does the skip get a chance:
  
   var held = _seeded.TryGetValue(snapshot.ProjectId, out var was) ? was : null;
   if (held is not null && SameSources(held.Sources, snapshot.Modules))
  
  so the comparison that saves the engine's work is fed by a read that already happened.
- **Why:** docs/lessons.md finding 44 records that the skip took a byte-identical write-back to "nothing at all
  - no seed, no diagnostics, not one call" on the engine side. The host-thread side was not covered:
  the pass still marshals in, reads every module of every open workbook, and blocks the VBE thread for
  the duration. Reanalyse has ten call sites (finding 45), coalesced to at most two passes per burst,
  so a burst of write-backs still pays two full source reads of every open workbook. On the Massive
  fixture that is two 1.42 MB BSTR transfers plus a .NET string each, on the thread the developer is
  typing on.
- **Change:** Gate the expensive part of the read. CountOfLines is one property read per component and is already
  fetched inside ReadSource; reading it first and skipping GetStringIndexed for components whose line
  count AND last-known text are both unchanged would turn most passes into a cheap walk. Where
  correctness needs the text (a macro rewriting a module without changing its line count), the session
  already tracks _writtenModules and ResyncFromModule already re-reads open documents, so the risk is
  confined to unopened modules changed by another add-in. Measure it before choosing: add a duration
  around ProjectReader.ReadAll and expose it on the existing perf route, which today reports engine
  method timings and placement/marshal samples but nothing about the host-thread read.
- **Size:** One full source read of every open workbook per pass, up to two passes per write-back burst. Duration unmeasured - the perf route (AddInSession.DebugApi.cs:1005) has no counter for it.
- **Adversary:** src/Xlide.Vbe.Shim/Engine/AnalysisService.cs:845 is AnalyseEverythingAsync's first act after the
  generation bump, and ReadProjectsAsync marshals ProjectReader.ReadAll onto the host thread at line
  797. ReadAll (src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:39-67) walks every project, Read (143-193)
  walks every component, and ReadSource (199-217) pulls the whole module text with
  code.GetStringIndexed("Lines", 1, lines) at line 210. The unchanged-sources skip is only reached at
  AnalysisService.cs:893-904, after the read has already happened for every project.
  docs/lessons.md:1110-1137 confirms the skip was designed and measured against the engine's work only
  ("A byte-identical write-back now costs the engine nothing at all"); the host-thread read is outside
  what it covers.
- **Correction applied:** Frequency is narrower than "every analysis pass" suggests to a reader. Reanalyse fires only from
  write-backs and adoption events, not per keystroke or per tick: AddInSession.cs:3806 is guarded by
  `if (adopted)`, and RunPassesAsync (AnalysisService.cs:317-331) coalesces a burst to at most two
  passes. So the cost is one full source read of every open workbook per write-back burst, not a
  steady-state drain. The proposed CountOfLines gate is also less free than it reads: CountOfLines is
  already the first read inside ReadSource (ProjectReader.cs:209), so the saving is only the
  GetStringIndexed transfer, at the price of missing a same-line-count rewrite in an unopened module.

##### `tree-republished-with-no-change-key` PublishProjects enumerates every project and component twice and posts the whole tree unconditionally, including once a second while the workspace is empty

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5850`
- **Kind:** perf / small effort, claim observed, confidence verified, severity medium
- **Evidence:** The method walks the project collection once for display names (line 5879):
  
   for (var i = 1; i <= projectCount; i++)
   {
   using var project = projects!.GetItem(i);
   displays.Add(project is null ? "VBAProject" : WorkbookDisplayName(project));
   }
  
  and again for the tree itself (line 5906), reading Name and Type for every component and calling
  ProjectReader.Identity(project) once more per project. It ends with an unconditional send (line
  5941):
  
   surface.ShowProjects([.. tree]);
  
  EditorSurface.ShowProjects (line 1019) serialises and posts with no comparison of its own. Contrast
  PublishModules, which does keep one (line 5644): `if (key == _lastModulesKey) { return; }`. Call
  sites include FollowActivePane (line 1604, every pane-substance change) and PollDebugState's
  empty-workspace branch (line 3847), which runs at EmptyWorkspacePollMilliseconds = 1000.
- **Why:** With no module open - a fresh workbook, or after the last tab closes - the shim re-reads every
  component of every open workbook and posts a full setProjects message once a second, forever,
  whether or not anything changed. That is the same once-a-second push that docs/lessons.md finding 20
  traced a visibly cycling tree to; the page was made idempotent, the source of the push was not. It
  also means a tab switch pays a double enumeration of the whole object model on the host thread
  before the new module is shown.
- **Change:** Give PublishProjects the same treatment PublishModules already has: compose the tree, build a change
  key from the project displays and each project's component names and types, and return before the
  send when it matches the last one. Then fold the two enumerations into one pass - the display names
  can be collected in the same loop that builds the members, with the numbering applied afterwards,
  since numbering only needs the name list.
- **Size:** Removes one full setProjects serialisation and one full component enumeration per second in the empty-workspace state; halves the COM enumeration on every tab switch. Absolute cost unmeasured - no counter exists on this path.
- **Adversary:** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5880-5884 enumerates the project collection for display
  names and 5907-5939 enumerates it again for the tree, reading Name and Type per component and
  calling ProjectReader.Identity(project) a second time at 5936. The send at 5941 is unconditional,
  and EditorSurface.ShowProjects (src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1019-1026) serialises and
  Sends with no comparison. The once-a-second claim holds end to end: PollDebugState calls
  PublishProjects unguarded at AddInSession.cs:3847 whenever _watchingEmpty, and UpdatePolling
  (3703-3707) resolves the interval to EmptyWorkspacePollMilliseconds = 1000 (line 3728) whenever
  _watchingEmpty is set and nothing faster is watching, with no decay - only _pollsRemaining decays
  (3946-3950). _watchingEmpty is set on the last-pane close at 5556. The non-empty branch by contrast
  is properly gated on a project-count change (3869-3875).
- **Correction applied:** The PublishModules change key the finding contrasts against is at AddInSession.cs:5647-5652, and the
  `if (key == _lastModulesKey)` line is 5649, not 5644.

##### `resync-reads-every-open-document` Switching tabs re-reads the entire source of every open document over COM, each preceded by a name scan of the project's components

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3777`
- **Kind:** perf / medium effort, claim observed, confidence verified, severity medium
- **Evidence:** ResyncFromModule loops every open document and, per document (lines 3777-3778):
  
   using var found = FindComponent(module, projectId, out _);
   var stored = found is null ? null : ProjectReader.ReadSource(found);
  
  FindComponent (line 6969) has no index - it walks the projects, calls ProjectReader.Identity on
  each, then walks VBComponents comparing Name:
  
   var candidate = components.GetItem(j);
   if (candidate?.GetString("Name") == component)
  
  and ReadSource pulls the whole module text (ProjectReader.cs:210). Its own comment at line 3760 says
  "this runs on every pane follow", and the call site is FollowActivePane at line 2778... the
  substance-changing branch (line 1610), reached whenever the tracked pane set changes by anything
  other than geometry - which is what a tab switch is.
- **Why:** With eight tabs open in a 60-component workbook, one tab switch costs eight component name scans
  (about 240 cross-COM calls each with the current no-DISPID-cache arithmetic) plus eight full module
  text transfers, synchronously on the VBE host thread, before the new module appears. The
  per-document loop already caches project identities within one call; the component lookup and the
  text read are not cached at all. docs/lessons.md finding 59 put the felt threshold for a tab gesture
  at about 40ms and this path has no measurement.
- **Change:** Two independent cuts. First, resolve the component once per document and keep the DispatchObject or
  at least a (projectId, module) to index memo, invalidated on the component add/remove/rename paths
  the session already owns. Second, gate the text read: CodeModule.CountOfLines is one property read,
  and a module whose line count matches its _writtenModules baseline's line count needs no full
  transfer to establish it is probably unchanged - fall through to the full read only when the counts
  differ or the module has no baseline. Instrument it first: a duration around ResyncFromModule on the
  existing perf route would say whether this is the tab-switch cost or a rounding error.
- **Size:** O(open tabs) full module reads plus O(open tabs x components) property reads per tab switch. Unmeasured; no counter exists.
- **Adversary:** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:3777-3778 is exactly as quoted: FindComponent per open
  document, then ProjectReader.ReadSource on whatever it found. FindComponent (6969-7011) has no index
  and walks VBProjects, calls ProjectReader.Identity per project, then walks VBComponents comparing
  GetString("Name") per item. ReadSource pulls the whole module text (ProjectReader.cs:210). The path
  is a tab gesture: FollowActivePane gates its tail on a substance key at AddInSession.cs:1570-1577
  and calls ResyncFromModule at 1610, past that gate, so a pane-set change - which is what a tab
  switch is - pays it. The comment at 3751-3754 confirms the author knows the resolve walks the
  project collection on every pane follow. No memo exists for the component lookup or the text read;
  only the project identities are cached, within one call, at 3770-3775.
- **Correction applied:** The evidence's call-site sentence is garbled: ResyncFromModule is called at AddInSession.cs:1610
  (FollowActivePane, past the substance gate at 1573) and at 2778 (the search-and-replace path), not
  "at line 2778... the substance-changing branch (line 1610)". The per-switch count is also slightly
  lower than stated: documents with unwritten edits are skipped first (3762), so it is one full read
  per open tab that is not currently being typed in.

##### `pane-project-identity-re-derived-per-tick` The poll tick re-derives each open pane's project identity from scratch, including a thrown-and-caught exception per unsaved workbook

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5814`
- **Kind:** perf / small effort, claim observed, confidence verified, severity medium
- **Evidence:** PollDebugState calls PublishModules on every tick (line 3934), and PublishModules calls
  ReadOpenModules (line 5536), which for each pane walks CodePanes -> Item -> CodeModule -> Parent ->
  Name and then, per pane (line 5814):
  
   using var collection = component.GetObject("Collection");
   using var project = collection?.GetObject("Parent");
   if (project is not null)
   {
   owner = DisplayFromProjectId(ProjectReader.Identity(project).Id);
   }
  
  ProjectReader.Identity (ProjectReader.cs:90) reads Name, then reads FileName inside a try, and for
  an unsaved workbook that read raises - its own comment says so: "Unsaved: the property raises rather
  than answering empty." The tick interval is DebugPollMilliseconds = 150 while stepping and
  ImmediatePollMilliseconds = 300 whenever the Immediate panel is open (line 2244/2247), and the
  Immediate panel is a panel a developer leaves open. Nothing memoises a pane's owner between ticks;
  the change key at line 5644 is compared only after all the reads have happened.
- **Why:** Steady state with the Immediate panel open and five tabs: about 42 property reads (84 cross-COM
  calls at today's two-calls-per-read) every 300 milliseconds, on the VBE host thread, to compute an
  answer that is almost always identical - plus one .NET exception thrown and caught per unsaved
  workbook per pane per tick. The dirty-flag half of PublishModules genuinely has to be polled (a
  host-side Ctrl+S fires no event we hear, per the comment at line 3928); the pane-to-project mapping
  does not, because it only changes when a pane opens, closes, or is renamed, and CodePaneTracker
  already invalidates on exactly those events (CodePaneTracker.cs:206, `_openComponents = null` on
  create/destroy, and line 180 on a rename).
- **Change:** Reuse the tracker's cache instead of re-deriving. CodePaneTracker.ReadPaneComponents already
  produces OpenComponent(Name, ProjectIdentity) and already holds it in _openComponents with correct
  invalidation; expose it and have ReadOpenModules take the owner from there, falling back to the
  current derivation only when the tracker has no entry for a pane. Separately, cache Identity per
  project object for the lifetime of a DispatchObject so the FileName raise happens once rather than
  once per pane per tick.
- **Size:** Removes roughly 6 property reads and up to one throw/catch per open pane per tick, at 3.3 ticks/second while the Immediate panel is open. Per-read cost unmeasured.
- **Adversary:** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:5777-5814 is as quoted: per pane, CodePanes.Item ->
  CodeModule -> Parent -> Name, then Collection -> Parent -> ProjectReader.Identity, roughly eight
  property reads per pane. ProjectReader.Identity (src/Xlide.Vbe.Shim/Engine/ProjectReader.cs:90-104)
  reads Name then FileName inside a try whose comment says an unsaved project raises rather than
  answering empty, so the throw per unsaved workbook per pane per tick is real. PublishModules runs
  every tick (AddInSession.cs:3934) and its change key is only compared at 5647-5652, after every
  read. Intervals confirmed: DebugPollMilliseconds = 150 at line 2244, ImmediatePollMilliseconds = 300
  at 2247, selected at 3704-3705. The proposed reuse is better grounded than the finding claims:
  CodePaneTracker.ReadPaneComponents (src/Xlide.Vbe.Shim/Editor/CodePaneTracker.cs:502-543) performs
  the identical walk down to ProjectIdentityOf and is memoised in _openComponents (line 340) with
  invalidation on rename (182) and pane create/destroy (208), so the tick's walk is a literal
  duplicate of an already-cached one.
- **Correction applied:** One objection the finding does not anticipate, and it survives: ReadOpenModules' own doc at
  AddInSession.cs:5767-5771 forbids deriving this list "from the tracker's window map", because with
  maximised panes only the active pane has a window. That prohibition is about FindPaneWindows, not
  about _openComponents, which is read from the CodePanes collection itself - so the proposed fix does
  not contradict it. The throw/catch is the smaller half of the cost; the six-to-eight cross-COM
  property reads per pane per tick are the substance.

##### `markers-posted-to-every-open-document` A caret leaving a typed line re-filters the whole findings list and posts a diagnostics message for every open document, when only one module's markers can have changed

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:7084`
- **Kind:** perf / small effort, claim observed, confidence supported, severity low
- **Evidence:** PublishMarkersToSurface loops every open document and filters the full findings set per document
  (line 7084):
  
   foreach (var (module, project) in surface.OpenDocuments)
   {
   var markers = _findings
   .Where(f => string.Equals(f.Module, module, ...)
   && ... && !_activeLineHold.Hides(f.Module, f.StartLine, f.EndLine))
   .Select(...)
   .ToArray();
  
   surface.ShowDiagnostics(module, project, markers); // line 7101
   }
  
  ShowDiagnostics (EditorSurface.cs:979) serialises and posts unconditionally - it keys the
  held-message slot per document but does not compare against what was last sent. It is always called
  paired with PublishFindingsToSurface (line 6445), which walks _findings again and serialises the
  whole set. The trigger is the active-line hold (lines 1149 and 1157): LineTyped calls both when
  Begin changes the hold, CaretLineSettled calls both when Release does. ActiveLineHold.Hides
  (Xlide.Vbe.Core/Editor/ActiveLineHold.cs:70) can only ever affect findings whose Module equals the
  single held module.
- **Why:** Moving the caret off a line the developer just typed on costs, on the host thread: two walks of the
  entire findings list, N+1 JSON serialisations, and N+1 postMessage calls, where N is the number of
  open tabs. N-1 of those diagnostics messages are byte-identical to the ones the page already has,
  because the hold cannot touch a module that is not the held one. This fires roughly twice per line
  of typing. I read the shim side end to end; I did not read the page's setDiagnostics handler, so
  what a redundant message costs once it lands is not established here.
- **Change:** Publish markers for the module whose hold changed rather than for every open document - the two hold
  call sites both know the module. Keep the all-documents form for the paths that genuinely replace
  the whole finding set (the analysis pass at line 1189, ShowModuleInSurface at line 5441). If the
  all-documents form has to stay for safety, add a per-document last-sent key inside ShowDiagnostics
  the way PublishModules keys the strip, so an identical marker set sends nothing.
- **Size:** N-1 redundant serialise-and-post per line transition, N = open tabs, roughly twice per line typed. Serialisation cost per document depends on the finding count for that module, which is unmeasured.
- **Adversary:** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:7084-7101 loops every open document and re-filters the
  whole _findings list per document, then calls ShowDiagnostics unconditionally.
  EditorSurface.ShowDiagnostics (src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:979-988) keys the held
  slot per document but never compares against what was last sent, and Send
  (EditorSurface.cs:1248-1264) posts immediately once loaded. Both hold call sites pair it with
  PublishFindingsToSurface (AddInSession.cs:1149-1150 and 1157-1158), and both only fire when the hold
  actually changed (ActiveLineHold.Begin/Release, src/Xlide.Vbe.Core/Editor/ActiveLineHold.cs:30-55).
  Hides (ActiveLineHold.cs:70-74) requires Module to match, so a hold transition provably cannot
  change any other module's marker set, making the other N-1 messages byte-identical to what the page
  holds.
- **Correction applied:** The "analysis pass" call site is AddInSession.cs:1483, inside the findings callback's
  RunOnHostThread hop, not 1189; line 1190 is the page-ready republish and 1189 is
  PublishFindingsToSurface. Size: the redundant part is N-1 small setDiagnostics messages, each
  carrying only one module's markers, so it is cheap per tab. The larger of the two costs at a hold
  transition is the PublishFindingsToSurface serialisation of the entire findings set
  (AddInSession.cs:6445-6457), which happens once per call regardless of tab count and which the
  finding's proposed fix does not address.

### The page

_10 findings, from the `perf-page` finder._

##### `splitter-drag-relayouts-every-editor` Every splitter drag runs a synchronous Monaco layout per pointermove, on top of automaticLayout, the exact doubling main.ts already diagnosed and fixed for window resize

- **Where:** `ui/editor/src/paneldocks.ts:874`
- **Kind:** perf / small effort, claim derived, confidence verified, severity medium
- **Evidence:** paneldocks.ts:874-890 buildPaneSplitter's apply: `const total = node.direction === "row" ?
  container.clientWidth : container.clientHeight;` ... `cells.forEach((cell, cellIndex) => {
  cell.style.flex = ... });` then `this.handlers.layoutChanged();`. paneldocks.ts:820-825 does the
  same for the dock splitter (`style.setProperty("--dock-size", ...)` then
  `this.handlers.layoutChanged()`), and workspace.ts:845 for the editor-area splitter. Each is called
  from `const move = (moved: PointerEvent) => { ... apply(position - last); ... }` registered at
  paneldocks.ts:914 / 849 / workspace.ts:869 - one call per pointermove event, not per frame.
  layoutChanged is `() => workspace?.editors().forEach((editor) => editor.layout())` (main.ts:438, and
  again at main.ts:467 for the Shell), and editors() is `this.groups.map((group) => group.editor)`
  (workspace.ts:470-472), so it is every open group. In monaco 0.56,
  esm/vs/editor/browser/widget/codeEditor/codeEditorWidget.js:1080 `layout(dimension,
  postponeRendering = false) { this._configuration.observeContainer(dimension); if
  (!postponeRendering) { this.render(); } }`, and
  esm/vs/editor/browser/config/elementSizeObserver.js:82-92 `measureReferenceDomElement(true,
  dimension)` -> `observedWidth = this._referenceDomElement.clientWidth;`. So each call is a forced
  style+layout read immediately after a style write, plus a synchronous view render. Every editor is
  created with `automaticLayout: true` (main.ts:327), which already installs a ResizeObserver on the
  same container. main.ts:444-447 records the identical defect on the other path: 'It waits for the
  resize to pause - running it per event doubled every layout of a drag, which read as latency and
  churn (2026-08-05)'. That fix (a 150ms settle plus the body.live-resize class that quiets the
  minimap, main.ts:453-461) was applied only to `window.addEventListener("resize", ...)`; the three
  splitter drag paths were left calling layout per event.
- **Why:** Dragging any dock or group splitter is the most common layout gesture on the page. Each pointermove
  writes a style and then forces a layout read and a full synchronous Monaco render, once per open
  editor group, while the ResizeObserver that automaticLayout installed is going to render the same
  frame anyway. A 1000Hz mouse delivers several pointermoves per frame, so the drag pays N_groups
  synchronous renders per event and 1 more per frame from the observer. body.live-resize is also never
  applied during a splitter drag, so the minimap flicker main.ts:449-452 describes (canvas repainting
  a frame behind the layout that moved it) is still present on exactly this gesture.
- **Change:** Delete the explicit editor.layout() from the drag path and let automaticLayout's observer do it, or,
  if the observer proves to miss the final frame, coalesce layoutChanged into one
  requestAnimationFrame per drag and flush once on pointerup. Add/remove
  document.body.classList.toggle("live-resize") around the drag the way the window-resize handler
  does, so the minimap gets the same quiet.
- **Size:** N_groups forced layouts + synchronous renders per pointermove eliminated (N_groups is 1 in the common case, up to the number of split groups). Unmeasured in absolute ms; the settling measurement is a performance.mark around apply() during a 2-second drag with a 4000-line module open, compared with the explicit layout() removed.
- **Adversary:** Read all three drag paths myself. paneldocks.ts:820-825 (dock splitter apply: sets --dock-size then
  this.handlers.layoutChanged()), paneldocks.ts:874-890 (pane splitter apply: writes cell.style.flex
  per cell then this.handlers.layoutChanged()), workspace.ts:831-846 (editor-area splitter, same
  shape). Each apply() is called from the pointermove listener registered at paneldocks.ts:849,
  paneldocks.ts:914 and workspace.ts:869, one call per event with no rAF or settle. layoutChanged
  resolves to editors().forEach(editor => editor.layout()) at main.ts:438 (Workspace handlers) and
  main.ts:467 (Shell handlers, which shell.ts:294 passes straight through to PaneDocks), and editors()
  is this.groups.map(g => g.editor) at workspace.ts:470-472. In the installed monaco 0.56
  (ui/editor/package.json devDependency, node_modules present) codeEditorWidget.js:1080
  layout(dimension, postponeRendering=false) { this._configuration.observeContainer(dimension); if
  (!postponeRendering) { this.render(); } } and elementSizeObserver.js:81-92 observe(dimension) ->
  measureReferenceDomElement(true, dimension) -> observedWidth =
  this._referenceDomElement.clientWidth, so each call is a forced measure followed by a synchronous
  render(true) (codeEditorWidget.js:1249-1256). Every editor is created with automaticLayout: true
  (main.ts:327), whose ResizeObserver already covers the same resize, and main.ts:444-461 documents
  the identical defect being fixed for window resize only (150ms settle plus body.live-resize). grep
  for live-resize finds it only in main.ts:455/459 and styles.css:834, so no splitter drag applies it.
  Nothing in the proposed fix touches AccessVBOM, synthetic input, NativeAOT or the api rules.

##### `bundle-is-unattributed` The build emits no metafile, so nothing in the repo can say what the 3.64 MB first-paint bundle is made of

- **Where:** `ui/editor/build.mjs:253`
- **Kind:** perf / small effort, claim observed, confidence verified, severity low
- **Evidence:** build.mjs:253-285 `const common = { ... bundle: true, ... minify: true, sourcemap: false,
  legalComments: "none", ... logLevel: "warning" }` - no `metafile: true`. build.mjs:287-295 runs the
  two builds and build.mjs:301-314 reports only file sizes from readdir/stat, so the only number
  produced is the total. `ls -la ui/editor/dist` gives editor.js 3,642,564 bytes, editor.css 184,668
  bytes, editor.worker.js 304,346 bytes. tools/verify.ps1:212-214 checks only that the bundle builds
  and that warnings are 0; the only other reference (line 309) checks that dist\index.html and
  dist\editor.js exist. There is no size budget and no per-module attribution anywhere in the repo.
- **Why:** The recorded startup breakdown puts nearly all of ~2.1s in bundle fetch and parse, and every
  decision about what to cut from that 3.64 MB (which of the ~40 monaco feature registers at
  main.ts:6-61 is expensive, whether referenceSearch or quickCommand or folding is worth its bytes)
  currently rests on guesswork. esbuild already computes the answer during the build and throws it
  away. Without it, any feature removal is a shot in the dark and any regression - a new import that
  drags in another 300 KB of monaco - lands silently.
- **Change:** Add `metafile: true` to the common esbuild options and, after the build, print the top 20 inputs by
  bytesInOutput from results[0].metafile (and write the full metafile to dist/metafile.json for
  esbuild's analyzer). Then consider a size assertion in tools/verify.ps1 alongside the existing
  warnings check, so a bundle that grows past a stated ceiling fails the gate rather than showing up
  as slower startup weeks later.
- **Size:** Zero runtime cost; it is the measurement that sizes every other startup finding. Two lines in build.mjs.
- **Adversary:** build.mjs:253-285 sets bundle/minify/target/alias/loader/logLevel with no metafile key;
  build.mjs:287-295 builds the two entry points and build.mjs:301-317 reports only readdir/stat byte
  counts and a warnings count. ls of ui/editor/dist gives editor.js 3,642,564 bytes, editor.css
  184,668, editor.worker.js 304,346 - the finding's figures are exact. tools/verify.ps1:207-217 ('page
  build') only fails on a non-zero exit or a non-zero warning count, and verify.ps1:309 only checks
  that dist\index.html and dist\editor.js exist in the publish. The only size assertion in the repo is
  a lower bound on the worker (ui/editor/test/smoke.mjs:60-61, size > 50_000). No metafile, no
  per-input attribution, no ceiling anywhere.
- **Correction applied:** The '~2.1s startup, nearly all bundle fetch and parse' in why_it_matters is not sourced anywhere in
  the repo - docs/status.md and the handoffs carry no startup figure - so treat it as a remembered log
  line, not evidence. That does not change the finding: adding metafile:true is two lines with zero
  runtime cost.

##### `objbrowser-loads-the-whole-editor` The Object Browser window loads and executes the entire 3.64 MB editor bundle, every Monaco feature included, to draw two lists

- **Where:** `ui/editor/src/main.ts:1239`
- **Kind:** perf / medium effort, claim derived, confidence verified, severity low
- **Evidence:** main.ts:1239-1244: '// One bundle, two documents: the editor surface, and the Object Browser palette
  the host // opens in its own floating window. The palette wants none of the editor's machinery - //
  no Monaco boot, no shell, no bridge - so it takes its own door before any of that starts.' followed
  by `const entry = new URLSearchParams(window.location.search).get("view") === "objbrowser" ?
  bootObjectBrowserPage : boot;`. That branch is at the BOTTOM of main.ts, after the top-level
  side-effect imports at main.ts:1-88 (`import "monaco-editor/features/bracketMatching/register.js"`
  through `wordPartOperations`, plus goToCommands, standaloneReferenceSearch, documentSemanticTokens,
  suggestController, and `import "./styles.css"`). Those all evaluate before the branch is reached, so
  the palette pays the full fetch, parse and module-init of every Monaco contribution and never uses
  one. build.mjs:290 declares a single page entry point (`entryPoints: { editor: "src/main.ts" }`) and
  one index.html (build.mjs:128-223), so there is no smaller document to serve. The host opens it as a
  brand new WebView2 in a new top-level window: src/Xlide.Vbe.Shim/Editor/BrowserPalette.cs:96-99
  `palette._browser = WebView2Surface.Start(handle, new PixelRect(...), "?view=objbrowser");`.
  bootObjectBrowserPage (objectbrowser.ts:164-171) then throws the shell markup away:
  `document.getElementById("shell")?.remove();`. It also never calls bridge.start, so unlike the
  editor page it reports no timings at all (main.ts:984-997 is inside boot()).
- **Why:** Opening the Object Browser costs a second full bundle parse, in a fresh JS context, for a page that
  uses none of it - the same order of latency as the editor's own ~2.1s startup, minus whatever the
  HTTP cache saves on the fetch half. It is also invisible: the palette emits no startup timing, so
  nobody can even see how long it takes.
- **Change:** Give the palette its own esbuild entry point (src/objbrowser.ts importing only objectbrowser.ts, its
  transport from bridge.ts, and its slice of styles) and its own objbrowser.html, then change
  BrowserPalette.cs:99 from "?view=objbrowser" to that page. build.mjs already builds two entry
  points, so a third is mechanical; the CSS split is the only real work. Have the palette report its
  own scriptMs/totalMs the way boot() does, so the saving is observable.
- **Size:** Removes essentially the whole 3.64 MB parse from the palette's load. Measure by having bootObjectBrowserPage post the same timings message boot() does, before and after.
- **Adversary:** main.ts:1242-1250 picks bootObjectBrowserPage vs boot from ?view=objbrowser at the very bottom of
  the module, after the eager side-effect imports at main.ts:1-88, so every monaco feature register
  has already run by then. build.mjs:290 declares one page entry (editor: src/main.ts) and
  build.mjs:128-223 writes one index.html whose body ends with <script src="./editor.js">, so there is
  no smaller document. BrowserPalette.cs:96-99 starts a separate WebView2 in its own top-level window
  with "?view=objbrowser", and objectbrowser.ts:164-171 then removes #shell and builds its own DOM; it
  constructs a PaletteHost and never calls bridge.start, so no timings message is posted
  (main.ts:984-997 is inside boot()).
- **Correction applied:** The palette is created ONCE per session, not per open: AddInSession.cs:2398-2404 returns early with
  _browserPalette.Present() when it already exists, and BrowserPalette.cs:163-180 Hide()/Present()
  just ShowWindow the same window. So the full-bundle fetch, parse and module-init is a one-time cost
  on the FIRST Object Browser open of a session, not a cost per open. Its absolute size is unmeasured
  (the ~2.1s comparison has no source in the repo), which is precisely why the metafile/timing work
  should land before the entry-point split.

##### `problems-panel-rebuilt-on-every-push` The problems panel rebuilds its whole DOM on every findings push with no equality guard, while both sibling lists have one

- **Where:** `ui/editor/src/shell.ts:881`
- **Kind:** perf / small effort, claim observed, confidence verified, severity low
- **Evidence:** shell.ts:881-883 `setFindings(findings: ShellFinding[]): void { this.findings = findings;
  this.renderPanel(); }` - no comparison against what is already on screen. renderPanel
  (shell.ts:1076-1162) then does `this.panelList.replaceChildren();` (1092), a full copy-filter-sort
  of every finding (1097-1103), a second full pass to build findingHomes (1107-1114), and per
  surviving finding creates five elements, four dataset writes and an appendChild straight into the
  live pane (1116-1150). The other two lists on the same message do guard: explorer.ts:418-424
  `setProblemCounts` opens with '// Identical counts arrive constantly and must change nothing:
  neither a redraw of a large // unfolded list, nor a re-parse of its module.' and returns early when
  equal; workspace.ts:324-336 renderTabs builds a renderKey string and returns when `renderKey ===
  this.lastTabsKey`. The cadence is stated by the page's own code: explorer.ts:655 '// An unchanged
  list redraws nothing: refreshes arrive with every analysis pass, and most // of them confirm what is
  already on screen.' The host sends the whole project each time:
  src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1159-1166 '/// <summary>Replaces the panel's contents,
  across every module.</summary> public void ShowFindings(SurfaceFinding[] findings)'.
- **Why:** Every analysis pass while typing rebuilds the entire problems list from scratch, and by the
  explorer's own comment most passes confirm what is already there. On a project with a few hundred
  findings that is over a thousand element creations and appendChilds into a laid-out container,
  discarded a moment later, on the same tick that the explorer and the tab strip correctly decide to
  do nothing. It also destroys the row the developer had focused or was about to click.
- **Change:** Give setFindings the renderKey guard workspace.renderTabs already uses: join
  severity|module|project|line|column|code|message plus the three filter states, compare with the last
  one, and return when equal. Same shape, same file conventions, three lines.
- **Size:** Eliminates roughly 6 DOM node creations per finding per analysis pass in the common (unchanged) case. Measure with a performance.mark around renderPanel and a counter of guarded-vs-rebuilt pushes over a minute of typing.
- **Adversary:** shell.ts:880-884 setFindings assigns and calls renderPanel with no comparison. renderPanel
  (shell.ts:1076-1162) does panelList.replaceChildren() at 1092, a full copy-filter-sort at 1097-1103,
  a second full pass building findingHomes at 1107-1114, then per surviving finding creates 5 elements
  (row, mark, body, message, where) with 4 dataset writes and appends into the live pane at 1116-1150.
  Both siblings on the same message do guard: explorer.ts:418-425 setProblemCounts returns early on
  equal counts with the comment 'Identical counts arrive constantly and must change nothing', and
  workspace.ts:323-335 renderTabs returns when renderKey === this.lastTabsKey. The push path is
  bridge.ts:1153-1154 case 'setFindings' from EditorSurface.cs:1160-1166.
- **Correction applied:** The cadence is smaller than 'every analysis pass while typing'. The per-pause LIVE path is already
  deduped on the host: AddInSession.cs:5334-5339 'Unchanged findings are not republished' with if
  (existing.SequenceEqual(findings)) return. The genuinely unguarded pushes are (a) the project pass,
  AddInSession.cs:1460-1485 FindingsReady -> PublishFindingsToSurface unconditionally, fired from
  AnalysisService.cs:1014 which is deliberately unconditional, and triggered by discrete Reanalyse()
  calls rather than by keystrokes, and (b) the active-line hold at AddInSession.cs:1144-1159, which
  republishes on every line the caret enters or leaves while typing and rebuilds the panel even when
  the held line carries no findings, so the visible set is identical. The guard is still worth having;
  the win is on line changes and repeated project passes, not on every keystroke.

##### `objbrowser-filters-by-full-rebuild` Object Browser search re-sorts and re-renders both panes synchronously on every keystroke, undebounced, unvirtualised, two listeners per row, over an uncapped type list

- **Where:** `ui/editor/src/objectbrowser.ts:538`
- **Kind:** perf / medium effort, claim observed, confidence verified, severity low
- **Evidence:** objectbrowser.ts:538-541 `search.addEventListener("input", () => { renderTypes(); renderMembers();
  });` - no debounce. renderTypes (444-508) starts `typesPane.replaceChildren();` then `for (const
  type of [...types].sort(byName))`, so the whole type array is copied and sorted on every keystroke,
  and each surviving type builds four elements plus `item.addEventListener("click", ...)` (498) and
  `item.addEventListener("keydown", ...)` (499-504) before `typesPane.appendChild(item)` (506).
  renderMembers does the same from membersPane.replaceChildren() at 310. The list is uncapped at the
  source: src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:2506-2508 `var types =
  TypeLibraries().TypesOf(library) ?? []; ... return [.. types.Select(row => new ObTypeRow(row.Name,
  row.Kind))];` - no Take, no page size. scopePick and the member-arrival callback (431-432, 544-545)
  go through the same two functions.
- **Why:** Typing into the Object Browser's search box is the palette's main interaction, and each character
  re-sorts the whole library and rebuilds up to every row of it, attaching two fresh closures per row.
  On a full type library that is thousands of element creations and thousands of listener
  registrations per keystroke, all synchronous on the input event, which is felt directly as typing
  lag in the box.
- **Change:** Three independent fixes, cheapest first: (1) sort once when the types arrive rather than inside
  renderTypes; (2) debounce the input handler the way searchwidget.ts:430 already does
  (`this.refindTimer = window.setTimeout(() => this.findInModule(true), 150)`); (3) replace the
  per-row click/keydown listeners with one delegated listener on typesPane reading a data attribute,
  which also removes the per-rebuild listener churn. Virtualisation is a larger change and should wait
  for a measured row count.
- **Size:** Per keystroke: one full array sort plus ~4 elements and 2 listeners per visible type, eliminated or deferred. Unmeasured in ms; measure by logging types.length for the Excel and Office libraries and timing renderTypes with performance.now().
- **Adversary:** objectbrowser.ts:538-541 search.addEventListener('input', () => { renderTypes(); renderMembers(); })
  with no debounce; scopePick (543-546), pickType (437-442), adoptScope (512-529) and the members
  arrival (430-432) all re-enter the same two functions. renderTypes (444-508) starts
  typesPane.replaceChildren() then iterates [...types].sort(byName), so the array is copied and sorted
  on every call, and each surviving row builds 4 elements plus a click listener (498) and a keydown
  listener (499-504). renderMembers does the same from membersPane.replaceChildren() at 310 with
  click/dblclick/keydown per row (372-389). The host caps nothing: AddInSession.cs:2506-2508 returns
  TypeLibraries().TypesOf(library) projected whole, no Take. The debounce precedent the fix cites is
  real: searchwidget.ts:429-430 window.setTimeout(() => this.findInModule(true), 150).
- **Correction applied:** The magnitude is unsupported. 'Thousands of element creations and thousands of listener
  registrations per keystroke' rests on a type count nobody has read - the shim logs it
  (AddInSession.cs:2507 'object browser: {library} -> {types.Count} type(s)') and a typelib is
  normally in the hundreds. Worse, the count SHRINKS as the query grows, because renderTypes skips
  non-matching types before creating anything (465-474): the fixed per-keystroke cost is one array
  copy plus sort of the library's types, and rows are only built for matches. Sorting once on arrival
  plus the 150ms debounce are the two cheap fixes; delegated listeners and virtualisation should wait
  for a measured row count.

##### `fulltext-shipped-per-keystroke` Every keystroke ships the module's entire text to the host for anything under 64 KB, in the same function whose comment calls that 'what typing latency is made of'

- **Where:** `ui/editor/src/bridge.ts:1636`
- **Kind:** perf / small effort, claim observed, confidence supported, severity low
- **Evidence:** bridge.ts:1615-1620 '// A small module travels whole, which is simplest. A large one travels as its
  changes: // building and shipping the full text per keystroke is what typing latency is made of, //
  and the host reconstructs the same text from the ranges.' Then bridge.ts:1626-1640: `const
  fullLength = model.getValueLength();` ... the message is built with the full `changes` array already
  (1610-1613), and `if (fullLength < 64_000) { message.fullText = model.getValue(); }` before
  `this.transport.post(message);`. The host prefers fullText whenever present and only reconstructs
  otherwise: src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:1379-1390 `var updated =
  document.RootElement.TryGetProperty("fullText", out var text) ? text.GetString() : null;` and the
  ParseChanges/ApplyEdits branch runs only `if (updated is null ...)`. So for every module under 64 KB
  - which is most VBA modules - the incremental path the comment describes is never taken, and each
  keystroke posts the whole document text plus the changes that describe the same edit.
- **Why:** Per keystroke this is an O(document) string build from Monaco's piece tree, a JSON serialisation of
  it, a WebView2 postMessage copy, and a JsonDocument parse of the same bytes on the host side, for
  information the host can already derive from the ranges that ride in the same message. The threshold
  reads as a value nobody measured: the code contradicts its own stated rule up to 64 KB. It has a
  correctness consequence too - the reconstruction path only executes on modules above 64 KB, so it is
  close to untested in normal use.
- **Change:** Measure the two paths, then either lower the threshold to where the crossover actually is or drop
  fullText entirely and keep the fullLength cross-check (which is already what catches divergence,
  EditorSurface.cs:1394-1399). Whichever way it goes, put the measured number in the comment so the
  next reader does not have to re-derive it.
- **Size:** Up to 64 KB of string build + JSON + IPC + host-side parse removed per keystroke; typical VBA modules are far smaller. Settle it by timing the contentChanged handler with a 60 KB module (fullText path) against a 66 KB one (changes path) at the same typing rate.
- **Adversary:** bridge.ts:1599-1641 onModelContentChanged runs per content event; it builds the changes array at
  1610-1613, reads fullLength at 1626, and at 1636-1638 adds message.fullText = model.getValue()
  whenever fullLength < 64_000, then posts. The host prefers it: EditorSurface.cs:1379-1390 reads
  fullText first and only runs ParseChanges/ApplyEdits when updated is null, with the length
  cross-check at 1391-1399. So for any module under 64 KB the incremental path never executes and the
  whole document rides every keystroke alongside the ranges that already describe the edit. The
  untested-branch half is verified too: grepping tools, docs, ui/editor/test and tests for
  fullText/fullLength/ParseChanges/ApplyEdits finds no test of either branch, and tests/ holds only
  Xlide.Vbe.Core.Tests.
- **Correction applied:** 'The code contradicts its own stated rule' is unfair - bridge.ts:1615-1617 states a deliberate
  TWO-path design ('A small module travels whole, which is simplest. A large one travels as its
  changes'), and EditorSurface.cs:1377-1379 says the same. What is actually wrong is that 64_000 is an
  unmeasured constant, and the confidence should stay 'supported': the per-keystroke cost of a
  getValue plus JSON on a typical few-KB VBA module is likely tens of microseconds, so the real payoff
  of the finding is the coverage gap, not the latency. Do not drop fullText outright before the
  reconstruct branch has a test, since today it is exercised only by modules over 64 KB.

##### `iife-forecloses-lazy-loading` The page is one iife with no code splitting and zero dynamic imports, so every dialog, the dev surface and the object browser page are parsed before first paint

- **Where:** `ui/editor/build.mjs:290`
- **Kind:** perf / medium effort, claim observed, confidence verified, severity low
- **Evidence:** build.mjs:287-295 builds the page as `build({ ...common, entryPoints: { editor: "src/main.ts" },
  format: "iife" })` with no `splitting` option - esbuild supports splitting only for format esm, so
  the iife choice forecloses it. `grep -n "import(" ui/editor/src/*.ts` returns nothing: there is not
  one dynamic import in the page. main.ts:69-88 therefore pulls in, eagerly and unconditionally,
  installDevSurface (devsurface.ts, 63,412 bytes of source), openSyncDialog (syncdialog.ts, 26,929),
  bootObjectBrowserPage (objectbrowser.ts, 19,340), openSettingsDialog/openPanesMenu
  (settingsdialog.ts, 17,252), openHelpDialog (6,876), openSponsorDialog (6,428) and
  openReferencesDialog (5,471) - about 145 KB of source for surfaces that all start closed. There is
  no build-time gate on any of them: build.mjs:259-263 defines only __XLIDE_BUILD__, __XLIDE_VERSION__
  and __XLIDE_BUILD_NUMBER__, and one config serves every build, so devsurface.ts ships in Release
  too.
- **Why:** The largest single startup cost is bundle parse, and today nothing can be moved out of it: the
  format decision means there is no mechanism for 'load this when the pane opens'. The app's own share
  is the smaller term (roughly 145 KB of 700 KB source against a 3.64 MB output, so single-digit
  percent after minification), but the mechanism is also the prerequisite for ever deferring anything
  on the Monaco side, and for the object browser split above.
- **Change:** Switch the page entry to `format: "esm", splitting: true` and index.html to `<script type="module"
  src="./editor.js">` - the worker entry is already esm (build.mjs:294) and the CSP is already
  `script-src 'self' blob:` (build.mjs:133), so both allow it. Then convert the six dialog openers and
  installDevSurface to `await import()` at their call sites in main.ts. Do the metafile change first
  so the saving is a number rather than a hope; a module script is also deferred, which changes when
  boot() runs relative to first paint and needs checking against the host's ready handshake.
- **Size:** Roughly 145 KB of source (unminified) moved off the first-paint path, plus the ability to defer anything else later. Exact output bytes are unknown until the metafile exists.
- **Adversary:** All the facts hold. build.mjs:287-290 builds the page as format 'iife' with no splitting option;
  grep for 'import(' across ui/editor/src returns nothing, so there is not one dynamic import.
  main.ts:69-88 imports installDevSurface, openSyncDialog, bootObjectBrowserPage,
  openSettingsDialog/openPanesMenu, openHelpDialog, openSponsorDialog and openReferencesDialog
  eagerly, and installDevSurface is called unconditionally at main.ts:929 - build.mjs:259-263 defines
  only __XLIDE_BUILD__, __XLIDE_VERSION__ and __XLIDE_BUILD_NUMBER__ and one config serves every
  build, so devsurface.ts (63 KB of source) is in the Release bundle too. The CSP at build.mjs:133 is
  script-src 'self' blob:, which permits a module script, and the worker entry at build.mjs:294 is
  already esm.
- **Correction applied:** The perf payoff is unmeasured and probably small, and the finding says so itself: ~145 KB of
  unminified app source against a 3.64 MB output that is overwhelmingly monaco, pulled in by the ~40
  eager feature registers at main.ts:6-61 which no format change defers. Treat this as a
  mechanism/prerequisite item strictly downstream of the metafile finding, not as a startup saving in
  its own right; switching index.html to a module script also changes when boot() runs relative to the
  host's ready handshake (EditorSurface ready path), which is the real risk.

##### `loop-sync-reads-whole-document-per-key` Typing inside a loop header builds the entire document string on every character

- **Where:** `ui/editor/src/typing.ts:490`
- **Kind:** perf / small effort, claim observed, confidence verified, severity low
- **Evidence:** typing.ts:150-152 `if (!/[\r\n]/.test(change.text)) { this.maybeLoopSync(model, change); }` on every
  single-change content event. maybeLoopSync (typing.ts:481-500) gates cheaply on the current line -
  `if (!LOOP_LINE.test(model.getLineContent(line))) { return; }` - and then, when the gate passes,
  typing.ts:490 `const edit = resolveLoopIteratorSyncEdit(model.getValue(), offset);`.
  model.getValue() flattens Monaco's whole piece tree into one string, per keystroke, and
  resolveLoopIteratorSyncEdit (vendored from the spec at ui/editor/vendor/xlide-spec) is then handed
  the whole document to scan.
- **Why:** Every keystroke typed on a For/Do line in a large module allocates the full module text and hands it
  to a scanner, on the input event, before the character is painted. The gate keeps it off ordinary
  lines, so this is narrow rather than pervasive - but loop headers are exactly where a developer
  types slowly and watches the line, which is the worst place to add per-character O(document) work.
- **Change:** Cache the flattened text against model.getVersionId() so repeated keystrokes in one line reuse one
  string, or check whether the spec helper can be given a bounded window around the offset. The
  vendored spec sources are not this repo's to change, so the caching route is the one available here.
- **Size:** One full-document string allocation per keystroke on loop lines. Measure by typing a 30-character loop header in a 4000-line module with a performance.mark around maybeLoopSync; the cost of resolveLoopIteratorSyncEdit itself is unread and could dominate the getValue.
- **Adversary:** typing.ts:150-152 calls maybeLoopSync for every single-change event whose text has no newline, and
  maybeLoopSync at typing.ts:481-500 gates on the current line then calls
  resolveLoopIteratorSyncEdit(model.getValue(), offset) at 490 - a full flatten of the model per
  qualifying keystroke, before the caret's character is painted.
- **Correction applied:** Three corrections. (1) The gate is For/Next only, not For/Do: typing.ts:47 LOOP_LINE = /^[
  \t]*(?:For|Next)\b/i. (2) The scanner does NOT scan the whole document on every call -
  vendor/xlide-spec/vbaSmartEnter.ts:205-210 early-outs on physicalLineAt(source, offset) and only
  reaches physicalLines(source) when the offset actually touches an iterator span, so the cost being
  paid is the getValue allocation itself, not a document scan. (3) The proposed fix as written does
  not work for the stated reason: every keystroke bumps model.getVersionId(), so a version-keyed cache
  never hits across keystrokes. It helps only by sharing one flatten with bridge.ts:1637, which calls
  getValue() on the SAME version for modules under 64 KB - which also means that on typical modules
  this finding is a doubling of an already-paid cost, and the case where it is the only full flatten
  is a module over 64 KB, where bridge.ts skips fullText.

##### `page-boot-timings-are-log-only` The page's startup breakdown reaches the log and nowhere else, so the one number that governs the startup work cannot be asserted by any test

- **Where:** `src/Xlide.Vbe.Shim/Editor/EditorSurface.cs:2119`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity low
- **Evidence:** The page measures carefully: main.ts:106 `const scriptMs = performance.now();`, main.ts:442 `const
  createMs = performance.now();`, and main.ts:984-997 posts scriptMs, createMs, totalMs, fetchMs,
  transferBytes, requestStartMs and htmlMs to the host. The host's DescribeTimings
  (EditorSurface.cs:2119-2149) reads each field into a local, formats `$" in {total}ms (bundle
  {script}ms, editor {create}ms{detail}...)"` and returns the string - nothing is retained on the
  surface or the session. Grepping src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs,
  tools/harness/xlide-api.mjs and docs/debug-api.md for scriptMs, bootMs or timings finds only the
  unrelated `stats` row about placement-pass timings (docs/debug-api.md:53).
- **Why:** Startup is named as the highest-value page path and the page already computes the full breakdown,
  but the only consumer is a log line a human reads. No harness suite can assert it, tools/verify.ps1
  cannot regress on it, and a change that adds 400 ms to bundle parse ships without anything failing.
  Every other finding here that claims a startup saving has to be verified by reading a log by hand.
- **Change:** Retain the parsed timings on the surface as a small record and answer them from a new debug-api
  route (a new noun, not a change to an existing one), documented in both docs/debug-api.md and
  docs/driving-excel.md and driven by a harness check, per the repo's own rule that every route
  appears in both and is driven by something. That makes the startup number an observation the gate
  can hold to a ceiling.
- **Size:** No runtime cost; it turns an unmeasurable path into a measurable one. The fields already exist end to end.
- **Adversary:** The page posts the full breakdown (main.ts:984-997: scriptMs, createMs, totalMs, build, fetchMs,
  transferBytes, requestStartMs, htmlMs) and the host consumes it into a string and drops it:
  EditorSurface.cs:2119-2150 DescribeTimings reads each field into a local and returns the formatted
  suffix, used only at EditorSurface.cs:1307 Log.Info($"editor surface: ready{DescribeTimings(...)}").
  Only the build stamp survives (EditorSurface.cs:1309 PageBuildStamp, Debug-only). Grepping
  AddInSession.DebugApi.cs, tools/harness/xlide-api.mjs and docs/debug-api.md for
  scriptMs/bootMs/timings finds nothing; the totalMs at DebugServer.cs:1031 is DebugRouteCost, the
  door's own per-route cost, unrelated. A new noun for it obeys the repo's rules (new capability as a
  new route, documented in both docs, driven by something).
- **Correction applied:** 'No harness suite can assert it' overstates the gap. Two partial observations already exist:
  docs/debug-api.md:54 'log' takes match and waitMs and BLOCKS until a matching line appears, so a
  suite can already pull the formatted 'editor surface: ready in Nms (bundle ...)' line, and the
  reload route already answers an elapsed reload-to-ready in ms (AddInSession.DebugApi.cs:1573-1600,
  150ms polling granularity). The real defect is narrower and still worth fixing: the breakdown is not
  retained as structured data, so nothing can hold scriptMs to a ceiling without parsing a log string.

##### `glyph-hover-decoration-per-mousemove` Hovering the glyph margin rewrites a model decoration on every mousemove rather than once per line crossed

- **Where:** `ui/editor/src/bridge.ts:1750`
- **Kind:** perf / small effort, claim observed, confidence verified, severity low
- **Evidence:** bridge.ts:537-543 wires `editor.onMouseMove((event) => { if (editor === this.ed()) {
  this.onMouseMove(event, hover); } else { hover.clear(); } })`. onMouseMove (bridge.ts:1733-1766)
  returns early unless the target is GUTTER_GLYPH_MARGIN and the line has no breakpoint, then calls
  `lineCanCarryBreakpoint(this.model()?.getLineContent(line) ?? "")` and `hover.set([...])` at
  bridge.ts:1750 - with no memo of the line it last previewed. The non-margin case is genuinely free:
  monaco's EditorDecorationsCollection.clear() early-returns when empty
  (esm/vs/editor/browser/widget/codeEditor/codeEditorWidget.js:1830-1836 `if
  (this._decorationIds.length === 0) { return; }`). set() has no such guard (1837-1843): it always
  runs changeDecorations + deltaDecorations, which fires onDidChangeModelDecorations and re-renders
  the glyph margin.
- **Why:** Moving the pointer down the glyph margin - the gesture a developer makes when hunting for a place to
  set a breakpoint - removes and re-adds the same decoration at every pointer sample instead of at
  every line boundary, each one a model decoration change and a margin re-render. Narrow gesture, but
  the fix is a two-line guard and it removes decoration churn from a path that also drives breakpoint
  clicks.
- **Change:** Hold the last previewed line number alongside the hover collection and return early when the target
  line is unchanged; clear it when the pointer leaves the margin or the model changes.
- **Size:** One deltaDecorations plus one glyph-margin render per pointermove reduced to one per line crossed.
- **Adversary:** bridge.ts:537-543 wires editor.onMouseMove to this.onMouseMove(event, hover); bridge.ts:1733-1766
  returns early with hover.clear() unless the target is GUTTER_GLYPH_MARGIN and the line carries no
  breakpoint, then calls hover.set([...]) at 1750 with no memory of the last previewed line. The
  monaco asymmetry is real in the installed 0.56: codeEditorWidget.js:1830-1836 clear() { if
  (this._decorationIds.length === 0) { return; } ... } but set() at 1837-1848 always runs
  changeDecorations + deltaDecorations. So moves over text are free and moves inside the margin
  re-delta the same decoration per pointer sample.
- **Correction applied:** Size is smaller than the write implies. The churn is confined to pointer samples inside the ~20px
  glyph margin during a brief hunt for a breakpoint line - a few dozen to a couple of hundred
  single-decoration deltas over the gesture, each cheap - so this is a two-line correctness-of-churn
  tidy, not a latency item. Guarding it is still right because the same collection feeds the
  breakpoint click path.

### The language engine and the checking loop

_10 findings, from the `perf-engine-gate` finder._

##### `close-project-keeps-live-text` project/close prunes two of the four per-module maps, so every closed workbook's live module text and outlines stay in the engine for the life of the process

- **Where:** `engine/src/dispatcher.ts:464`
- **Kind:** perf / small effort, claim observed, confidence verified, severity medium
- **Evidence:** closeProject (engine/src/dispatcher.ts:444-480) prunes exactly two maps by prefix and one by
  project:
  
   const prefix = `${params.projectId}\0`;
   for (const key of this.lastAnalysis.keys()) {
   if (key.startsWith(prefix)) { this.lastAnalysis.delete(key); }
   }
   for (const key of this.semanticMemo.keys()) {
   if (key.startsWith(prefix)) { this.semanticMemo.delete(key); }
   }
   this.symbolsMemo.delete(params.projectId);
   forgetProjectWords(params.projectId);
  
  openProject, for modules that vanished from a reseed, prunes FOUR (dispatcher.ts:427-434):
  
   this.lastAnalysis.delete(key);
   this.semanticMemo.delete(key);
   this.outlineMemo.delete(key);
   this.liveSources.delete(key);
  
  I grepped every mention of both missing maps in the file (liveSources at lines 199, 382, 431, 651,
  808, 820, 824, 835, 842, 849; outlineMemo at 206, 431, 782, 796). The only deletes are the
  openProject prune at 431/432, which is scoped to the CLOSING project's own prefix and so never
  revisits a project that has gone, and the desync delete at 835. Nothing else clears either map.
  liveSources holds one full module source string per module (dispatcher.ts:199, "The live text of
  modules being edited"); outlineMemo holds a source string plus its OutlineResult (206).
- **Why:** The engine is one process per Excel session and outlives every workbook. Closing a workbook
  currently releases the analyzer's own per-document state (the fix documented at
  dispatcher.ts:447-457, "every module a session ever analysed stayed held for the life of the
  engine") and the analysis/semantic memos, but keeps a full copy of the text of every module the
  developer typed in, plus its outline. A developer moving between workbooks all day accumulates all
  of it. This is the same defect the closeProject comment was written to close, two maps short.
- **Change:** Extend the existing prefix loop in closeProject to cover outlineMemo and liveSources, the same four
  maps openProject already prunes. Cheapest correct form is one loop over a list of the four maps
  rather than four copies of the same loop, so the next map added cannot be missed the same way.
- **Size:** unmeasured; the retained bytes are the sum of every closed project's module sources. Measurable today without new code by driving project/open and project/close through the engine and reading process.memoryUsage(), or by adding the count of liveSources/outlineMemo entries to the debug/liveSource reply.
- **Adversary:** Read engine/src/dispatcher.ts:444-480 myself: closeProject prunes lastAnalysis (465-469) and
  semanticMemo (471-475), deletes symbolsMemo (477) and calls forgetProjectWords (478). outlineMemo
  and liveSources are not touched. openProject's prune at 427-434 does delete all four. liveKey at
  dispatcher.ts:75-77 is `${projectId}\0${moduleName.toLowerCase()}`, the exact prefix closeProject
  already loops on, and outlineMemo (780-796) and liveSources (817-849) are both keyed with it, so the
  omission is a real leak and the proposed one-loop fix is mechanically correct. liveSources holds one
  whole module source per module (199); outlineMemo holds a source plus its OutlineResult (206).
- **Correction applied:** Two details in the evidence are off. (1) The openProject prune is scoped to the OPENING project's
  prefix and additionally only removes keys whose module name is absent from the reseed, so it
  reclaims a closed project's entries only if that same projectId is reopened later with those modules
  gone; a workbook closed and never reopened is never reclaimed at all, which is stronger than the
  finding states. (2) The desync delete at 835 clears one module's live text on a mismatch, not a
  project's. Severity is memory growth, not correctness: nothing serves stale text out of these maps
  after a close, because seededModules and generations are dropped at 445 and 462 and every read path
  goes through a project that is no longer known.

##### `dev-loop-repackages-engine-always` tools/dev.ps1 rebuilds and repackages the 90 MB engine executable on every run, including for pure C# changes, and only then checks whether it needed to

- **Where:** `tools/dev.ps1:143`
- **Kind:** perf / small effort, claim observed, confidence verified, severity medium
- **Evidence:** dev.ps1:143-189, the step named 'Build the engine (bundle, then executable)', runs unconditionally
  whenever -NoBuild is absent:
  
   node build.mjs --package
   ...
   if ($LASTEXITCODE -ne 0) { throw "The engine build failed ($LASTEXITCODE)." }
  
  The staleness comparison it would need to skip that is computed immediately AFTER, as an assertion
  (dev.ps1:172-185):
  
   $newestEngineSource = Get-ChildItem $engineSources -Recurse -Include *.ts |
   Sort-Object LastWriteTime -Descending | Select-Object -First 1
   $engineExe = Get-Item $enginePath ...
   if ($newestEngineSource -and $newestEngineSource.LastWriteTime -gt $engineExe.LastWriteTime) {
  throw ... }
  
  verify.ps1 already made exactly this conditional, and states the cost (verify.ps1:126-140):
  "Packaging only when something IS newer, rather than on every run: the injection writes a 90 MB
  executable and takes the better part of a minute, which is most of this gate's whole runtime to
  spend on the common case where nothing changed." Its condition is the same comparison
  (verify.ps1:113-122), over engine/src plus the neighbouring analyzer checkout, which dev.ps1:172-174
  also assembles.
- **Why:** dev.ps1 is the inner loop: build, register, launch. A shim-only change pays the full SEA injection
  every round. verify.ps1 judged that cost to be most of a twenty-second gate and removed it; the
  script run far more often still pays it, and the data needed to skip it is already gathered four
  lines later.
- **Change:** Hoist the comparison at dev.ps1:172-185 above the `node build.mjs --package` call and package only
  when a watched source is newer than engine\dist\xlide-engine.exe, keeping the post-package read-back
  assertion for the case where it did run. This is the shape verify.ps1:117-160 already uses, so the
  two scripts would decide identically instead of one skipping and one not.
- **Size:** verify.ps1:136 puts the injection at 'the better part of a minute'; saved on every dev.ps1 run that touches no .ts. Measurable by timing `node build.mjs --package` once.
- **Adversary:** tools/dev.ps1:143-189 read in full. `node build.mjs --package` runs at 159 with no staleness
  condition in front of it, inside `if (-not $NoBuild)` at 134; the newest-source comparison is built
  at 172-185 and used only to throw afterwards. verify.ps1:107-162 does the same comparison first and
  packages only when $newer.Count -gt 0, and its own comment at 135-137 puts the injection at 'the
  better part of a minute'. engine/build.mjs:29-70 confirms --package does the SEA blob plus a full
  copy of node.exe, so the cost is real and unconditional.
- **Correction applied:** Two adjustments to the proposed change. dev.ps1:176 globs only *.ts while verify.ps1:117 watches
  *.ts, *.mjs, *.js; hoisting the comparison must adopt verify's include list or the two scripts will
  still disagree, which is the defect being fixed. And dev.ps1 has no equivalent of
  verify.ps1:145-150, the check that refuses to package while EXCEL or xlide-engine holds the
  executable, so the conditional should carry that guard across too rather than only the comparison.

##### `project-index-rebuilt-per-module` The whole-project index is rebuilt once per module on every context-cache miss, and a second whole-project index is built separately from the same seeded array for the diagnostics fingerprint

- **Where:** `engine/src/moduleContext.ts:139`
- **Kind:** perf / medium effort, claim derived, confidence supported, severity medium
- **Evidence:** buildContext (moduleContext.ts:131-149) builds a project-wide index inside the per-module cache
  miss:
  
   const inputs: VbaProjectModuleInput[] = entries.map(...);
   const project = buildLiveVbaProjectIndex(inputs, {
   moduleName: current.name, moduleKind: context.moduleKind, source: current.source,
   });
   const symbols = projectEditorSymbolContextForModule(project, current.name);
  
  assembleContext caches per (seeded array identity, module name) (moduleContext.ts:68-89), so this
  runs once per MODULE per seed, not once per seed. `entries` is every module of the project (97-117),
  and buildLiveVbaProjectIndex walks all of them: ../xlide_vscode/src/vbaProjectAnalysis.ts:106-128
  shows buildLiveVbaProjectIndex delegating to buildVbaProjectIndex, whose body is `for (const mod of
  modules) { applyProjectModule(setModule, mod, liveOverride) }` - the liveOverride substitutes a
  single module, it does not narrow the walk.
  
  Separately, dispatcher.ts:128-134 builds another whole-project index from the same seeded array and
  throws it away after extracting strings:
  
   const index = buildVbaProjectIndex(seeded.map((module) => ({...})));
   const procedures = projectProcedureSignatures(index);
   ... describe(projectAnalysisOptionsForModule(index, module.moduleName, procedures))
  
  Only the resulting Map<string,string> is cached on the seeded array (crossModuleFacts,
  dispatcher.ts:93); the ProjectIndex itself is not.
- **Why:** Completion, hover, signature help, canonicalCase and semantic tokens all route through
  assembleContext. After a reseed - which is every write-back that changes any text
  (AnalysisService.cs:906) - the first request in each module pays a whole-project index build.
  Working across ten modules is ten of them, on a pipe that serves one request at a time
  (EngineClient.cs:672, "One request is outstanding at a time"), so each one also delays whatever
  queued behind it. The moduleContext comment at line 66 claims the cache turns completion "from
  indexing the whole project into scanning one module", which holds for repeat requests in one module
  and not for the first request in each.
- **Change:** Cache the ProjectIndex itself on the seeded array in a WeakMap the way crossModuleFacts already
  caches its fingerprints, and have both callers take it from there: dispatcher.crossModuleFingerprint
  and moduleContext.buildContext. Per-module work then reduces to
  projectEditorSymbolContextForModule(project, name). Two real differences must be preserved rather
  than assumed away: buildLiveVbaProjectIndex passes ignoreInvalidModules:true and
  buildVbaProjectIndex does not, and buildContext's liveOverride replaces the current module's entry
  with one carrying moduleKind in place of type/documentType. Both are settleable by keeping the live
  variant as the shared build and asserting the fingerprints are unchanged with
  engine/test/freshness.mjs.
- **Size:** unmeasured. Settled by timing textDocument/completion in a not-yet-cached module against a repeat request in the same module, on the perf fixture, through the engine's own pipe (engine/test/smoke.mjs already builds that harness).
- **Adversary:** engine/src/moduleContext.ts:68-89 caches per (seeded array identity, lowercased module name); the
  miss path at 131-149 calls buildLiveVbaProjectIndex over `entries`, which is every module of the
  project (97-117). ../xlide_vscode/src/vbaProjectAnalysis.ts:110-121 is `for (const mod of modules) {
  applyProjectModule(...) }` with liveOverride substituting one entry rather than narrowing the loop,
  and ProjectIndex.setModule (../xlide_vscode/src/analyzer/symbols/projectIndex.ts:425-437) calls
  buildModuleSymbols, which calls parseModule. dispatcher.ts:124-160 does build a second whole-project
  index for the fingerprint and keeps only the Map<string,string>. assembleContext is imported by
  completion.ts:25, hover.ts:15, onType.ts:159, semantic.ts:20 and signature.ts, so this is on the
  interactive request path, behind a pipe that serves one request at a time.
- **Correction applied:** The size is worse than the finding allows and can be stated from code rather than left unmeasured.
  ../xlide_vscode/src/analyzer/parser/parseModule.ts:135-136 sets PARSE_CACHE_MAX = 8 on an LRU keyed
  by source value. Building an index over a project with more than eight modules evicts every entry as
  it goes, so the next module's context build re-parses the whole project from scratch: for N > 8
  modules the per-module cost is a full N-module parse, not a warm re-walk. Below nine modules the
  second build is mostly cache hits and the finding is close to free. Also correct the trigger: a
  reseed is not every write-back, it is every write-back whose text actually differs, because
  AnalysisService.cs:894 skips OpenProjectAsync at 906 when SameSources holds.

##### `gate-runs-node-and-dotnet-in-sequence` The gate's four post-page-build node steps and its four dotnet steps are independent of each other and run strictly in sequence

- **Where:** `tools/verify.ps1:219`
- **Kind:** perf / medium effort, claim derived, confidence verified, severity low
- **Evidence:** Every step goes through `Step $name { ... }` (verify.ps1:56-78), which runs the scriptblock inline
  and records elapsed seconds; there is no job, runspace or Start-Job anywhere in the file. Order as
  written: vendored spec (80), engine executable is current (94), no variant is read as an object
  (165), page typecheck (198), page build (207), page tests (219), engine language matrix (229), page
  probes headless (246), debug api documented and driven (259), solution build Release (272), unit
  tests (278), Release carries no debug api (288), native publish (302).
  
  The candidate is CONFIRMED WITH ONE DEPENDENCY. The dotnet chain is NOT independent of the whole
  node chain: src/Xlide.Vbe.Shim/Xlide.Vbe.Shim.csproj:83 globs the page bundle into the build,
  
   <None Include="$(MSBuildThisFileDirectory)..\..\ui\editor\dist\**\*"
   LinkBase="ui\editor\dist"
  
  and verify.ps1:309-312 asserts the published output carries ui\editor\dist\index.html and editor.js.
  So 'solution build (Release)' depends on 'page build'. Nothing downstream of the page build does:
  'page tests' (npm test in ui/editor), 'engine language matrix' (node engine/test/language.mjs, which
  spawns its own engine per engine/test/language.mjs:37-38), 'page probes (headless)' (five node
  scripts) and 'debug api is documented and driven' (node audit-routes.mjs, which reads sources) touch
  no dotnet output, and the dotnet steps touch none of their inputs. 'no variant is read as an object'
  (165) is a Select-String over src\**\*.cs and depends on nothing at all.
- **Why:** The gate is the thing a developer runs before every commit, so its wall clock is paid many times a
  day. Half of it is idle CPU: while dotnet builds and tests, four node steps that could have been
  running are waiting, and vice versa.
- **Change:** Keep steps 1-5 (through 'page build') serial, then run two groups concurrently: {page tests, engine
  language matrix, page probes, audit routes} and {solution build, unit tests, Release carries no
  debug api, native publish}, joining before the summary. Start-ThreadJob or two Start-Job handles are
  enough; the Step function already stores per-step Ok/Seconds/Detail into an ordered hashtable
  (verify.ps1:54-78) so the summary table needs only the results merged back in order. Keep the FAILED
  line per step so a failure still names itself.
- **Size:** unmeasured here, but the gate already prints it: the summary at verify.ps1:485-490 emits per-step seconds on every run, so the saving is min(sum of the node group, sum of the dotnet group) read straight off one run.
- **Adversary:** Verified the structure: tools/verify.ps1:56-78 runs each scriptblock inline with `& $work` and a
  stopwatch, and a grep of the file for Start-Job, ThreadJob, Runspace and -Parallel returns nothing.
  Step order is as listed (80, 94, 165, 198, 207, 219, 229, 246, 259, 272, 278, 288, 302). The one
  dependency the finding names is real: src/Xlide.Vbe.Shim/Xlide.Vbe.Shim.csproj:83-86 globs
  ..\..\ui\editor\dist\**\* into the build and publish, and verify.ps1:309-313 asserts the published
  tree carries it.
- **Correction applied:** The saving is smaller than the framing implies and the proposed mechanism does not exist on this
  shell. docs/status.md:48 puts the twelve non-publish steps at 'about twenty seconds' in total, so
  the ceiling on the join is single-digit seconds unless the Release publish (verify.ps1:302, skipped
  by -Quick) is in the dotnet group, where it dominates and the node group finishes long before it.
  The environment is Windows PowerShell 5.1, where Start-ThreadJob is not present unless the ThreadJob
  module has been installed, leaving Start-Job, which spins a whole PowerShell process per group and
  costs about a second each before any work starts. Both groups are also already CPU-parallel
  (MSBuild, esbuild, the AOT linker), so overlapping them does not add throughput on a saturated
  machine. Worth doing only as -Quick's problem, not the full gate's.

##### `unit-tests-through-vstest` The gate runs the xunit v3 test assembly through dotnet test's VSTest host when the assembly is already a runnable executable

- **Where:** `tools/verify.ps1:279`
- **Kind:** perf / small effort, claim observed, confidence supported, severity low
- **Evidence:** verify.ps1:278-286 runs:
  
   $out = dotnet test $solution -c Release --no-build --nologo -v q 2>&1
  
  tests/Xlide.Vbe.Core.Tests/Xlide.Vbe.Core.Tests.csproj references `xunit.v3` 3.1.0 alongside
  `Microsoft.NET.Test.Sdk` 17.14.1 and `xunit.runner.visualstudio` 3.1.4, so the run goes through
  VSTest. xunit v3 also emits a self-hosting executable, and it is on disk:
  artifacts/bin/Xlide.Vbe.Core.Tests/release/Xlide.Vbe.Core.Tests.exe, beside
  artifacts/bin/Xlide.Vbe.Core.Tests/release/testhost.exe (the VSTest host the current path launches).
  Both exist for debug and release. The solution (xlide_vbide.slnx) contains exactly one test project,
  so `dotnet test $solution` resolves to that one assembly.
- **Why:** dotnet test pays MSBuild evaluation of the solution, VSTest discovery and a testhost process launch
  before the first test runs. Running the assembly directly skips all three and reports the same
  results. In a gate that the whole team's pre-commit loop waits on, that is dead wall clock on every
  run.
- **Change:** Replace the dotnet test call with a direct run of
  artifacts\bin\Xlide.Vbe.Core.Tests\release\Xlide.Vbe.Core.Tests.exe, exporting
  DOTNET_ROOT=$env:LOCALAPPDATA\Microsoft\dotnet first (verify.ps1:44-47 already locates that
  directory for PATH, so the value is in hand) since the apphost will not otherwise find a .NET 10
  outside Program Files. Parse the pass count from xunit's own summary rather than the 'Passed: N'
  line the current step scrapes at verify.ps1:282-284. Keep dotnet test in CI where the trx logger is
  wanted (.github/workflows/build.yml:151).
- **Size:** unmeasured. One run of each form, timed, settles it; the gate's own per-step seconds line for 'unit tests' is the before figure.
- **Adversary:** verify.ps1:279 is `dotnet test $solution -c Release --no-build --nologo -v q`.
  tests/Xlide.Vbe.Core.Tests/Xlide.Vbe.Core.Tests.csproj:13-15 references Microsoft.NET.Test.Sdk
  17.14.1, xunit.v3 3.1.0 and xunit.runner.visualstudio 3.1.4, so the VSTest path is what runs. Listed
  artifacts/bin/Xlide.Vbe.Core.Tests/release and both Xlide.Vbe.Core.Tests.exe and testhost.exe are on
  disk, alongside xunit.v3.runner.inproc.console.dll, so the self-hosting runner the proposal would
  call directly exists.
- **Correction applied:** The mechanism is one hop deeper than described, and it removes the objection the finding did not
  raise. xunit v3's VSTest adapter delegates to the test project's own executable as a child process
  (inferred from the package set, not read from the adapter's source here), so
  Xlide.Vbe.Core.Tests.exe is already what runs today under dotnet test. That means the Smart App
  Control constraint on this machine, which blocks fresh unsigned Release test binaries and is why
  dev.ps1:140 runs -c Debug, is not a new risk introduced by running the exe directly. What the change
  actually removes is MSBuild evaluation of xlide_vbide.slnx plus VSTest discovery plus testhost.exe,
  not the test execution, so the ceiling is a few seconds inside a roughly twenty-second gate.

##### `gate-packages-untypechecked-engine` The gate packages and ships the engine executable without ever typechecking engine/src; esbuild strips types without reading them

- **Where:** `tools/verify.ps1:152`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** verify.ps1:152 packages the engine as part of the 'engine executable is current' step:
  
   npm run package --prefix (Join-Path $repoRoot 'engine') 2>&1 | Out-Host
  
  engine/package.json defines the two as separate scripts: "check-types": "tsc --noEmit" and
  "package": "node build.mjs --package". engine/build.mjs:30-40 is an esbuild.build call with no type
  checking in front of it. I grepped the whole repo for check-types: the only two hits are
  engine/package.json:9 and .github/workflows/build.yml:60. Neither tools/verify.ps1 nor tools/dev.ps1
  calls it, and both of them run `node build.mjs --package` (verify.ps1:152, dev.ps1:159).
  verify.ps1's only typecheck is `npm run typecheck` inside $pageRoot (verify.ps1:201), which is
  ui/editor.
- **Why:** A type error in engine/src passes the local gate, gets injected into xlide-engine.exe by that same
  gate, and is what the developer then tests against and can publish. The gate's own 'engine
  executable is current' step exists precisely so the running executable matches the source; it makes
  the executable current and unverified in one move. CI catches it afterwards, which is the wrong side
  of the commit for the machine that just shipped the binary.
- **Change:** Run `npm run check-types` in engine/ inside the same step, before the package call at
  verify.ps1:152, and again in dev.ps1 before dev.ps1:159. tsc --noEmit over 14 files with
  skipLibCheck (engine/tsconfig.json) is a second or two, and only pays it on the runs that were going
  to package anyway.
- **Size:** tsc --noEmit over engine/src (14 files, 3,850 lines, skipLibCheck true); the cost is bounded by one tsc start-up.
- **Adversary:** verify.ps1:152 is `npm run package --prefix ...engine` with no typecheck anywhere in that step;
  verify.ps1:198-205 typechecks $pageRoot, which is ui/editor (verify.ps1:40). engine/package.json:9
  defines check-types and engine/build.mjs:29-40 is a bare esbuild.build with nothing in front of it.
  A grep of tools, .github, engine and package.json for check-types returns exactly
  engine/package.json:9 and .github/workflows/build.yml:60. build.yml:58-60 typechecks the engine only
  in the engine job, and that job runs `npm run build` (64), never --package, so no packaged
  executable in this project is ever produced by a typechecked run.
- **Correction applied:** The cost estimate is wrong by a large factor. engine/tsconfig.json includes only 'src', but every
  engine module imports the analyzer by relative path (engine/src/dispatcher.ts:10-19,
  completion.ts:17, moduleContext.ts:13-20, navigation.ts:14-22 and six more all resolve to
  ../../../xlide_vscode/src/*), and those are .ts sources, which skipLibCheck does not skip. tsc
  --noEmit therefore typechecks the whole shared analyzer as well, which is why CI gives it a
  dedicated job with the companion repo checked out. Expect seconds to tens of seconds, and expect the
  local gate to go red whenever the neighbouring analyzer working tree has a type error in it -
  defensible, since that tree is what gets bundled into the executable, but it puts a gate failure on
  a file this repo is forbidden to patch, so the step needs to name the analyzer in its failure text.

##### `freshness-test-runs-nowhere` engine/test/freshness.mjs, the headless guard on the diagnostics memo, is run by neither the default gate nor CI

- **Where:** `engine/test/freshness.mjs:1`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** engine/package.json's "test" script chains four suites (smoke, language, positions, freshness), but
  nothing calls `npm test` in engine/. verify.ps1's only engine step runs one of them
  (verify.ps1:240): `node test/language.mjs`. The CI engine job runs three
  (.github/workflows/build.yml:69 smoke.mjs, :85 language.mjs, :94 positions.mjs) and not freshness.
  Grepping *.ps1 and *.yml for 'freshness' returns only verify.ps1:397 and :421, which are the LIVE
  suite tools\harness\analysis-freshness.mjs behind -Live, not this one.
  
  What is unguarded: freshness.mjs's own header states it exists for the memo whose fingerprint could
  not see a signature change because `JSON.stringify(new Map([...]))` is `"{}"`, and the dispatcher
  comment at engine/src/dispatcher.ts:104-107 says "test/freshness.mjs passes with this handling
  REMOVED, because the other fields of the options object happen to move when a signature does". So
  the describe() helper at dispatcher.ts:112-118 has a suite written for it that runs on no machine
  automatically, and that suite is admittedly not tight enough to fail if describe() is reverted.
- **Why:** The diagnostics memo at dispatcher.ts:944-955 is the largest single perf decision in the engine
  ("446ms of an 476ms pass", dispatcher.ts:875) and it is only correct while crossModuleFingerprint is
  exact. When it stops being exact the failure is silent: no squiggle appears where one should, which
  is the header's own account of an hour lost. The -Live gate covers the invariant end to end, but
  that needs an open editor; the headless version needs nothing and runs nowhere.
- **Change:** Add `node test/freshness.mjs` beside the existing language.mjs line in verify.ps1's 'engine language
  matrix' step (or as its own step), and add it to the CI engine job after Positions
  (.github/workflows/build.yml:94). Separately worth doing: extend freshness.mjs with a case that
  fails when describe() degrades to plain JSON.stringify, since the dispatcher comment records that
  today it does not.
- **Size:** freshness.mjs spawns one engine over one pipe and does seed/ask/reseed/ask; comparable to smoke.mjs, which CI already affords on both runners.
- **Adversary:** engine/package.json:12 chains smoke, language, positions and freshness under 'test', and nothing
  runs npm test in engine/. verify.ps1's only engine step is 229-244, running `node test/language.mjs`
  and nothing else. I read .github/workflows/build.yml end to end: the engine job runs smoke (69),
  language (85) and positions (94), never freshness. A grep across tools and .github for 'freshness'
  returns only verify.ps1:397 and :421, which are the -Live suite
  tools\harness\analysis-freshness.mjs. The dispatcher's own comment at
  engine/src/dispatcher.ts:104-107 states that freshness.mjs passes with the describe() Map handling
  removed, so the suite is admittedly loose as well as unrun.
- **Correction applied:** One practical detail the proposal needs: engine/test/freshness.mjs:31-32 spawns dist/engine.cjs, not
  the packaged exe, and verify.ps1's engine step packages only when something is stale.
  engine/build.mjs:29-42 writes engine.cjs before the SEA injection, so the bundle does exist after
  any packaging run, but on the common run where nothing was newer the gate never invokes build.mjs at
  all and dist/engine.cjs is whatever the last build left. That is the same condition
  test/language.mjs already runs under at verify.ps1:240, so adding freshness beside it inherits an
  existing assumption rather than creating one, but the two suites share the risk of testing a bundle
  older than the sources when the exe was current.

##### `ci-missing-cheap-source-guards` CI runs neither the route audit nor the variant-as-object shape guard nor the headless page probes, all of which are cheap and all of which are in the local gate

- **Where:** `.github/workflows/build.yml:137`
- **Kind:** api-coverage / small effort, claim observed, confidence verified, severity medium
- **Evidence:** I read build.yml end to end (219 lines). The `build` job goes: npm ci, spec:check, typecheck, build,
  npm test, dotnet restore, build, test, Languages, publish, three pwsh binary checks, upload.
  Grepping the file for 'audit-routes' and 'As<' returns nothing.
  
  The three local-gate steps with no CI counterpart:
  - verify.ps1:165-196, 'no variant is read as an object': a Select-String for
  `\.As<\s*object\s*>\s*\(` over src\**\*.cs. Its own comment records that this defect "has killed
  Excel twice, months apart in code terms" and that com-leak.mjs "cannot catch either one".
  - verify.ps1:259-264, 'debug api is documented and driven': node tools\harness\audit-routes.mjs,
  which reads the routes out of the shim and fails when one is undocumented or undriven.
  - verify.ps1:246-253, 'page probes (headless)': five node scripts (close-confirm, objbrowser, tree,
  boot-error, sole-workbook) that need no host and no dotnet output.
- **Why:** CI is the only check that runs on a change that did not come from this machine, and the three
  cheapest guards in the gate are exactly the ones it does not run. The variant guard is a regex over
  C# source: it needs no build, no runner state and under a second, and it stands in for a crash that
  ends the host process and that the leak sweep provably cannot see.
- **Change:** Add the variant-as-object grep and `node tools/harness/audit-routes.mjs` to the build job (the grep
  can run before dotnet restore, since it needs only a checkout), and the five headless page probes
  after 'Check the page bundle', where the bundle they read has just been built. All four are
  pure-node or pure-text and add no runner dependencies the job does not already have.
- **Size:** the variant grep is one Select-String over src; audit-routes is one node process; the five probes are five short node processes. Sizeable from the gate's own per-step seconds for those three steps.
- **Adversary:** Read .github/workflows/build.yml in full (219 lines). The build job is npm ci, spec:check,
  typecheck, build page, npm test, dotnet restore/build/test, Languages, publish, three pwsh checks,
  upload. No audit-routes, no .As<object> grep, no page probes. The three local-gate steps cited exist
  as described: verify.ps1:165-196 (Select-String for '\.As<\s*object\s*>\s*\(' over src, with the
  comment recording two host-killing regressions and that com-leak.mjs cannot see them),
  verify.ps1:259-264 (node tools\harness\audit-routes.mjs), verify.ps1:246-253 (five probes).
- **Correction applied:** The page probes are not pure node. tools/harness/page-probe.mjs:72-80 locates msedge.exe at two
  hardcoded Windows paths and launches it headless over the DevTools protocol, throwing at :161 when
  neither exists. They can therefore only go in the windows-latest build job, never the ubuntu half of
  the engine matrix, and they add a browser dependency to a job that currently has none - which is
  fine on windows-latest, where Edge is preinstalled, but it is a new failure mode (a runner image
  change silently breaks five checks) rather than the zero-dependency addition claimed. The
  variant-as-object grep and audit-routes.mjs are exactly as described, need only a checkout, and are
  the two that should go in first.

##### `immediate-line-rereads-every-module` Every Immediate-window line triggers a full-project re-read across COM on the VBE host thread, after the module it was correcting for has already been removed

- **Where:** `src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:4193`
- **Kind:** perf / medium effort, claim derived, confidence supported, severity low
- **Evidence:** AddInSession.cs:4192-4193, immediately after evaluating one Immediate line:
  
   // Evaluating adds and removes a module, which the analyzer would otherwise report on.
   _analysis?.Reanalyse();
  
  Reanalyse (AnalysisService.cs:274-308) starts AnalyseEverythingAsync, whose first act is
  ReadProjectsAsync (AnalysisService.cs:845), which marshals ProjectReader.ReadAll onto the host
  thread (AnalysisService.cs:793-821). ReadAll walks every project and every component and reads each
  module's whole text over COM (ProjectReader.cs:39-67, 162-183, and ReadSource at 199-217:
  `code.GetStringIndexed("Lines", 1, lines)`). ProjectReader's own header (lines 17-20) says "Every
  call here crosses into the host... Reading a large project is measured in the tens of milliseconds,
  which is fine when a project opens and not fine per keystroke".
  
  The temp module is already gone before Evaluate returns: ImmediateEvaluator.cs:27 says "Removed
  after every evaluation", and Remove(components) is called in the finally at
  ImmediateEvaluator.cs:282 (and defensively before the add at :179). So by the time the pass reads,
  SameSources (AnalysisService.cs:1126-1142) reports the project unchanged and the engine-side work is
  skipped (AnalysisService.cs:894-904) - the COM read is the whole cost.
- **Why:** Immediate-window use is rapid-fire during a debug session, and this is the one path where a full
  snapshot read lands on the host thread per gesture rather than per structural change. The skip added
  at AnalysisService.cs:878-892 removed the analyzer cost of an unchanged project but not the read
  that decides it is unchanged, so the developer pays the read for a pass that will conclude nothing
  happened.
- **Change:** Two options, both in scope. Narrower: have ImmediateEvaluator report whether its cleanup Remove
  actually removed anything unexpected, and call Reanalyse only then, keeping the correction for the
  case the comment is really about (a pass that ran while the temp module existed). Broader and more
  generally useful: give AnalyseEverythingAsync a cheap pre-check on the host thread - component count
  plus CountOfLines per component - and skip ReadSource for a project whose shape and line counts are
  unchanged, falling back to the full read whenever anything differs. The second changes no observable
  behaviour and pays off for every Reanalyse caller, not just this one.
- **Size:** unmeasured. ProjectReader.cs:19 puts a large project's read at 'tens of milliseconds'; measurable now by timing ProjectReader.ReadAll on the perf fixture, or from the debug api's marshalMs samples around an Immediate line.
- **Adversary:** src/Xlide.Vbe.Shim/AddIn/AddInSession.cs:4192-4193 reads exactly as quoted, unconditional after
  Evaluate at 4180. AnalysisService.cs:274-308 starts a pass; AnalyseEverythingAsync:845 awaits
  ReadProjectsAsync, which at 810 marshals ReadOnHost -> ProjectReader.ReadAll onto the host thread,
  and ProjectReader's header at 17-20 says a large project's read is 'tens of milliseconds ... not
  fine per keystroke', with ReadSource pulling each module's whole text through CountOfLines plus
  GetStringIndexed("Lines", 1, lines). The scratch module is removed in the finally at
  Editor/ImmediateEvaluator.cs:282 before Evaluate returns, so the pass that follows almost always
  hits the SameSources skip at AnalysisService.cs:894-904 and makes no engine call, leaving the COM
  read as the entire cost. I compared the other nine Reanalyse call sites: 2051 and 5296 are the
  typing-driven ones and both sit behind FullAnalysisQuietMilliseconds, so 4193 really is the
  unthrottled per-gesture one.
- **Correction applied:** Two corrections. The cost is not paid inline: Reanalyse at 306-307 sets _passWanted and does
  Task.Run, and the read comes back to the host thread later through HostMarshal, so what the
  developer meets is a host-thread work item of tens of milliseconds shortly after the line runs, not
  a stall inside Evaluate. More importantly, the narrower of the two proposed fixes is wrong.
  Reanalyse at 4193 is not correcting for the scratch module still existing at the end; it corrects a
  pass that READ while it existed, and that is reachable because VBA execution pumps messages (the
  same property recorded at AddInSession.DebugApi.cs:1846-1850, where marshalled work runs while a VBA
  modal owns the editor), so a marshalled ReadOnHost can land in the middle of Evaluate. Gating on
  whether Remove removed something unexpected would not see that case. Only the broader fix - a cheap
  host-thread pre-check on component count plus CountOfLines before ReadSource - is admissible.

##### `no-memo-hit-visibility` The perf route reports what each engine method cost but nothing reports whether the engine's five memos are hitting, so a caching regression is only visible as latency without a baseline

- **Where:** `src/Xlide.Vbe.Shim/Diagnostics/EngineCounters.cs:57`
- **Kind:** api-coverage / medium effort, claim observed, confidence verified, severity low
- **Evidence:** EngineCounters.Record (EngineCounters.cs:57) keeps per method: Calls, WaitTotalMs, CallTotalMs,
  WaitMaxMs, CallMaxMs, Refused, and a 32-slot Recent ring. The perf route surfaces exactly those
  (AddInSession.DebugApi.cs:1005-1029) and the client reshapes them
  (tools/harness/xlide-api.mjs:629-642: method, calls, totalMs, waitMs, callMs, medianMs, p95Ms,
  worstMs, refused). Split into queued and served, which is the split that matters on a serialised
  pipe (EngineClient.cs:653-666).
  
  What it cannot answer: the engine holds five reuse mechanisms - lastAnalysis (dispatcher.ts:241),
  semanticMemo (214), outlineMemo (206), symbolsMemo (230) and crossModuleFacts (93) - and a hit and a
  cheap miss are the same fast call from the shim's side. The shim's own skip is invisible for a
  stronger reason: when SameSources says a project is unchanged (AnalysisService.cs:894-904) the pass
  makes NO engine calls at all, so a pass that correctly skipped everything and a pass that wrongly
  analysed nothing both show as zero counters. The engine exposes debug/liveSource
  (dispatcher.ts:802-814) for what it holds of one module's text, and nothing equivalent for what it
  is reusing.
- **Why:** Two of the three memo defects already recorded in this file were silent - the Map fingerprint blind
  spot (dispatcher.ts:98-107) and the rename that hit a stale symbolsMemo (dispatcher.ts:217-229) -
  and both would have shown instantly as a hit rate falling to zero. Today they show as callMs rising
  against a baseline nobody keeps, which is why one of them was found by an hour of confusion rather
  than by an instrument.
- **Change:** Add a new engine method (not a rename of an existing one) alongside debug/liveSource - debug/memos -
  answering per project: entry counts and hit/miss counters for the five maps, with a reset. Surface
  it as a new debug-api route beside perf rather than a new field on perf, since it is the engine's
  state and not the shim's timing, and give it a client method plus entries in docs/debug-api.md and
  docs/driving-excel.md so audit-routes.mjs is satisfied. It answers what the render depends on rather
  than what the render shows, so its route entry should say plainly that it stops at the engine layer.
- **Size:** five counter pairs and one route; the reporting cost is a JSON object of about twelve numbers per project.
- **Adversary:** The gap is real. EngineCounters.cs:39-90 tallies Calls, WaitTotalMs, CallTotalMs, WaitMaxMs,
  CallMaxMs, Refused and a 32-slot Recent ring and nothing else; I checked the route switch for an
  existing answer under another name, which is what would refute it. The 'stats' route
  (AddInSession.DebugApi.cs:1794-1836) reports process, placement, marshal, refresh, message and
  COM-wrapper counters, no engine cache state. The 'engine' route (2461-2502) answers one module's
  held text and line count against the surface, which is debug/liveSource - the only debug method the
  engine exposes (engine/src/dispatcher.ts:352). Nothing anywhere reports lastAnalysis, semanticMemo,
  outlineMemo, symbolsMemo or crossModuleFacts hit rates, and the finding's stronger point holds too:
  a pass that takes the SameSources skip at AnalysisService.cs:894-904 makes no engine call, so a
  correct skip and a broken pass are the same zero.
- **Correction applied:** Reframe the priority rather than the facts. Both historical defects it cites were closed by tests,
  not by instruments, and one of those tests already exists and simply runs nowhere -
  engine/test/freshness.mjs, the subject of the freshness-test-runs-nowhere finding. Wiring that into
  the gate and CI is a one-line change that covers the same silent-memo risk, where this proposal is a
  new engine method plus a new route plus a client method plus two doc entries to satisfy
  audit-routes.mjs. Build the counters only if a measurement shows a latency regression whose cause
  turns on memo behaviour; the ordering matters because these two findings compete for the same budget
  and one of them is a line.

---

## What was not read

Each finder reported its own coverage honestly, including what it did not reach. A named gap is
worth more than a false all-clear, so they are reproduced here rather than summarised away.

**Where driving Excel still falls off the api.** Read in full or substantially: the tools/harness listing and tools/*.ps1 listing; module-sync.mjs
(header, planner setup, sections 5 and 6, all eval sites); menu-bar.mjs; objbrowser-live-probe.mjs
(whole); immediate-watch.mjs (watch section and closing checks); audit-routes.mjs (whole);
page-probe.mjs header and the four headless probes' headers (close-confirm, sole-workbook,
boot-error, objbrowser); Test-ObjectBrowser.ps1 live section; Test-CloseHiddenPane.ps1,
Test-GhostLocalsPanel.ps1, Test-CloseVbe.ps1, Test-Language.ps1, Get-Shot.ps1 (whole);
Test-WatchPanel.ps1 header and helper; Test-ResizeFollow.ps1 lines 1-140; Get-EditorScreenshot.ps1
lines 200-300; Test-CloseConfirm.ps1 header; verify.ps1 live section (lines 325-470);
docs/driving-excel.md sections 4, 4a, 5, 6 and 7 in full plus targeted greps of the route tables;
the debug-api route switch by case list, with ui, act, eval, await, type, mark, guard, dismiss,
pane, command, windows, menus, capture, placement, sync and RunPageScript/RunPageScriptOnce read in
full; DialogWatch.cs (Dialogs, Dismiss); VbeCommands.ForName; PerfCounters.PlacementSnapshot;
devsurface.ts UiSnapshot, DevSurfaceParts, dialogsUp, the act table's names, and the dock, treeMenu,
press and closeActive actions; paneldocks.ts by grep.

Not read, so not covered: Test-DebugApi.ps1 (29 KB, greps only - it drives the api directly and its
eval uses looked like route self-tests rather than fallbacks, but I did not confirm each),
Test-SplitWorkspace.ps1 (greps plus the eval blocks quoted; the drag-compass sections at lines
290-370 drive synthesised pointer events at .drop-petal-* and I did not establish whether
act(\"dock\") is a faithful stand-in for a real drop, so I made no claim about it), Test-Churn.ps1
and Test-DiscardProblems.ps1 (headers only), Invoke-VbeLoadCheck.ps1, Get-EditorWindowTree.ps1,
Open-VbeIn.ps1, Open-VbeEditor.ps1, New-ScratchWorkbook.ps1, Start-Excel.ps1 (greps only -
Start-Excel is the documented and correct way in, so I looked no further), tools/New-*Fixture.ps1
(greps only; New-RenameFixture.ps1 is documented at driving-excel.md:1163 as still needing the trust
setting but its header at line 15 says it no longer does, and I did not resolve that contradiction),
tools/dev.ps1 and tools/page.ps1, and the suites com-leak.mjs, three-copies.mjs,
format-positions.mjs, colouring.mjs, settings-bite.mjs, analysis-freshness.mjs, perf-scaling.mjs,
write-rollback.mjs, import-guard.mjs, rename-features.mjs, language-features.mjs,
language-live-probe.mjs, engine-live-probe.mjs, build-fixture.mjs, tree-page-probe.mjs and
surface-walk.mjs, which I covered only by grepping for COM, CDP, filesystem and eval use.
language-live-probe.mjs:81 has one api.eval I did not chase to a missing field. I ran no code and
opened no Excel.

**Page features with no driver or no observer.** Read in full or near-full: ui/editor/src/toolbar.ts, contextmenu.ts, devsurface.ts (header,
UiSnapshot, DevSurfaceParts, state(), sendKey and the whole actions record), and the route-name
inventory of src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs (all 60 case labels) with deep reads
of act, ui, eval, await, layout, layout?reset, immediate, pane, undoRename, settings, command,
module, caret, breakpoint, documents and the LayoutScript constant. Read the load-bearing parts of
searchwidget.ts (wire, state, find, open, close, queryChanged, scopeChanged, runScoped,
findInModule, showSearchResults, showModuleResults), workspace.ts (snapshot, strip wiring, tab
context menu, moveTab), shell.ts (seat list, paneVisibility, notify, status fields, properties
writes), main.ts (toolbar dispatch, Ctrl+W/Ctrl+backslash, devsurface wiring, rename provider,
undoRename and findAllReferences actions, showReferences), paneldocks.ts (public method outline
only), and the full route table of docs/debug-api.md plus the act examples in docs/driving-excel.md.
Checked drivers against tools/harness/xlide-api.mjs, audit-routes.mjs (which audits routes, not act
actions - act coverage is unenforced) and grepped the suite directory for each gap.

Not read, so gaps there are not covered: menubar.ts, explorer.ts, objectbrowser.ts, syncdialog.ts
and settingsdialog.ts beyond targeted greps; helpdialog.ts, sponsordialog.ts, referencesdialog.ts,
bookmarks.ts, format.ts, typing.ts, dragcompass.ts, edgescroll.ts, theme.ts, documents.ts,
docktree.ts, settings.ts, vba.ts and bridge.ts were not read end to end. I did not read the assert
vocabulary cases (DebugApi.cs:2095-2142) closely, so gaps in what can be asserted about the page are
unexamined. I did not verify at run time that a synthesised KeyboardEvent through act key reaches
monaco's keybinding service, so a few gaps I call NO DRIVER (Find All References via Shift+F12,
Enter in the find box) may be awkwardly reachable through act key with an explicit key argument;
each finding's evidence rests on the direct path, not on that possibility. Nothing was executed,
nothing was written.

**Shim capabilities with no driver or no observer.** Read in full: VbeCommands.cs, GhostReaderThread.cs, XlideAddIn.cs, ProductSettings.cs, and the
DebugServer.cs reply records for state, native, locals, watches, immediate and windows. Read by
section: AddInSession.DebugApi.cs (RunPageScript and its surface routing 68-268; log, messages,
capture 518-660; immediate, locals, watches, problems 658-850; act 1062-1090; stats and the
host-thread marshal 1794-1900; the assert vocabulary 2095-2150; state and doctor head 2293-2340;
native, windows, menus, outline 2504-2705; component rename 2838-2886; pane, undoRename,
breakpoints, settings 2885-3035; project and documents 3089-3188; command, breakpoint, module,
caret, placement and the default arm 3188-3348) and AddInSession.cs (ExecuteEditorCommand 1725-1800,
PublishProperties 3074-3130, OnPropertyEdit and WriteProperty head 3216-3300, AdoptRename and
ComponentsChanged 3344-3405, debug-mode publishing and UpdateDebugState 3460-3660,
AttachImmediateReader and OnDebugOutput 3958-4030, OnModuleCloseRequested and ModuleDiffersFromSaved
6472-6580, RemoveComponent tail 6820-6870, Stop 7593-7675, plus the full method index).
Cross-checked every candidate against the route case list, the client method list in
tools/harness/xlide-api.mjs, the act table in ui/editor/src/devsurface.ts:514-1481, and the
docs/debug-api.md rows for state, ui, problems, watches, breakpoints and component. Confirmed that
eval?surface=palette and capture?window=palette do reach the Object Browser palette
(BrowserPalette.Browser at BrowserPalette.cs:33, selected at AddInSession.DebugApi.cs:220), so the
object browser and the type-library catalog behind it are not reported as gaps; likewise notices
(EditorSurface.Notify) and the setCurrentLine frame are visible in the messages tap, and the tab
dirty dot is visible through ui (workspace.ts:493-500), so those were dropped rather than reported.
NOT covered: HandleSync and the sync planners (AddInSession.cs:138-443, ModuleSyncService.cs,
Core/Sync/ModuleSync.cs) beyond confirming the route accepts plan, apply and settings - the
import/export planners and the folder lock are unaudited; CodePaneTracker, WindowEventHook,
FrameSubclass, OverlayWindow, WebView2Surface and the placement and cutout machinery were read only
through their call sites in AddInSession, so window-lifecycle observability (split, activate,
PoliceNativeToolWindows, DarkenTitleBar) is a gap in my coverage rather than an all-clear;
TypeLibraryCatalog, VbeMenus.Describe, DialogWatch internals, ImmediateEvaluator, LocalsReader and
WatchReader parsing, ProjectReader and EngineClient's transport were not read, and engine/src was
out of my path entirely. I did not open the harness suites, so I cannot say which of the routes I
judged sufficient are actually driven today - audit-routes.mjs enforces that separately. No Excel
was launched and nothing was written.

**The api as a testing instrument.** Read in full: src/Xlide.Vbe.Shim/AddIn/AddInSession.DebugApi.cs (all 3348 lines, both route switches
and every helper), tools/harness/audit-routes.mjs, docs/debug-api.md, tools/harness/xlide-api.mjs.
Read src/Xlide.Vbe.Shim/Diagnostics/DebugServer.cs lines 1-448 (server, request parsing, history,
durations, reply writing) and the reply-record declarations by name; I did NOT read the DebugCapture
GDI section (455-660) or every one of the ~50 reply records field by field - I read
DebugCommandReply, DebugErrorReply, DebugBlockedReply and DebugDialogRow/Reply and skimmed the rest.
Followed every claim to the implementing method: VbeCommands.Execute, ExecuteEditorCommand, GoTo,
ToggleBreakpoint, WriteModule, UndoRename, OnModuleCloseRequested, EditorSurface.Text/TextOf and the
incremental-change reconstruction. Of docs/driving-excel.md (1373 lines) I read the route-to-client
table (196-245) and the waiting/sleeps section (783-872) but not the operational walkthroughs, so a
doc-vs-code drift living in its prose rather than its table would have been missed. Ran nothing: no
builds, no harness, no Excel, no git state changes; the counts I cite (208 vs 5 `.log(` matches,
zero `.capture(` calls, zero drainFinalizers callers) come from read-only greps over tools/harness.
Gaps: HandleSync (AddInSession.cs:254 onward) was not read, so the `sync` route's four actions and
their error shapes are unaudited; ui/editor/src/devsurface.ts was not opened, so the drift between
the `do=` enumeration at docs/debug-api.md:70 and the page's real action list is unverified and
deliberately not reported; the ghost-reader snapshots behind `locals`/`watches` were not traced to
their source, so I cannot say whether those two observations are live or recorded; and I did not
audit the ~47 fixed sleeps in the .mjs suites one by one against the retained-by-design set named at
docs/driving-excel.md:823-825 (import-guard's three, three-copies' one), so the only sleeps I cite
are ones I read in context and confirmed are not in that set.

**AddInSession.cs.** Read fully: XlideAddIn.cs (all 412 lines). Built a complete member map of AddInSession.cs (160
private/internal methods, all fields, all consts) and ran repo-wide reference counts for every one
of those names across src, tools, ui, engine and docs -- no private method, field or const in either
file is unreferenced, so there are no dead members to report; the string-dispatch risk was moot
because nothing came back at zero. Deep-read in AddInSession.cs: the settings/sync region heads
(27-260), AddInSession/Start/HostStartupComplete, EnsureSurfaceForFrame (1035-1259),
TryShowEmptyWorkspace, FollowActivePane, OnSurfaceKey head, ExecuteEditorCommand, WriteModule
(1853-2065), PutModuleBack head, TryWriteLineDiff, SplitPhysicalLines, ScanProjectMembers,
ReplaceMatches head, UpdateDebugState, OnPanelChanged, WatchDebugState, UpdatePolling, the poll
constants, ResyncFromModule, ProjectCount, PollDebugState, the eight engine round-trip handlers
(4363-4711), OnMenuRequested/OnMenuExecuteRequested/RouteMenuCommand/ForgetBreakpoints,
CaptureBefore, PublishModules, ReadOpenModules, WorkbookSaved, ContentKey head, WorkbookDisplayName,
FindProjectByDisplayName, HideReplacedWindows head, the ghost palette region (6102-6266),
HideNativeToolbars head, DarkenTitleBar, HideImmediateWindow head, SaveWorkbookOf, ActivePaneOwner,
FindComponent, FindCodePane, IsShownComponent, ProjectIdFromDisplay, the placement region
(7105-7395), PoliceNativeToolWindows, Stop, ReportEnvironment, ReportOpenProjects, Dispose. Also
read for corroboration: Com/DispatchObject.cs Dispose and GetString, Com/ComRuntime.cs counters,
Engine/ProjectReader.cs Identity and ReadSource, Editor/VbeCommands.cs command ids and
RoutesThroughSession. NOT read, and therefore not audited: AddInSession.DebugApi.cs (another agent's
area, deliberately untouched); and within AddInSession.cs the bodies of HandleSync (254-443),
SharedPlan/SharedPlanFrom/DiffFrom (550-760), the scratch-module and breakpoint regions (898-1001,
2226-2335), ScanModuleMembers and the object-browser scanners (2436-2725),
RunSearchAsync/ReplaceInLine (2752-2900),
PublishBreakpoints/PublishProperties/OnPropertyEdit/WriteProperty/AdoptRename/ComponentsChanged
(2998-3427), PublishLocals/PublishWatches (3428-3538),
AttachImmediateReader/OnDebugOutput/OpenExternal/EvaluateImmediate (3958-4202),
OnRenameRequested/OnModuleRenameRequested/UndoRename/ApplyModuleRename (4764-5192),
OnNavigationRequested/OnSemanticTokensRequested/OnLiveAnalysisDue (5192-5354),
RunCommand/SyncCaretToPane/ShowModuleInSurface/RefreshWindowTitle (5354-5522), PublishProjects body
(5850-5947), the close/insert/remove region (6472-6882), PublishMarkersToSurface and TrackCodePanes
(7076-7592). The perf claim about PublishModules is derived from counting IDispatch invokes on the
tick, not from a measurement; no counter exists for it today and I did not run anything.

**AddInSession.DebugApi.cs and Diagnostics.** Read in full: src/Xlide.Vbe.Shim/Diagnostics/PerfCounters.cs (195 lines) and EngineCounters.cs
(181), both entirely inside #if DEBUG with no Release reachability and nothing dead in them. Read in
AddInSession.DebugApi.cs: 1-70 (header and helpers), 490-670 (dispatch head, log, messages,
capture), 909-1200 (assert, journal, perf, ui, act, eval, await), 1297-1620 (bench, trip, layout,
reload, dialogs), 1758-1904 (guard, dismiss, stats, and the host marshal that closes the pool
switch), 2060-2290 (EvaluateClaim, AnswerBlockedRequest), 2289-2350 (state, doctor head), 2704-2950
(sync, component, pane), 3034-3348 (projects, project, documents, command, breakpoint, module,
caret, placement, default). Read in DebugServer.cs: 1-260 (Start, discovery sweep, Loop, Serve,
ReadRequest head) plus a full member and record map by grep, and the 52 JsonSerializable entries.
Read Log.cs 100-200. Whole-file greps that back the counting claims: every `case "..."` label with
line numbers, every method declaration, every `default:`, every `new
DebugErrorReply(`/`DebugCommandReply(true, 0)`/`new
DebugEvalReply(`/`RunPageScript(`/`request.Query.TryGetValue` site, and repo-wide searches for each
Debug* record type, `Retried`, `CropBmp`, and the nine assert claim names.

Not read, and therefore not judged: AddInSession.DebugApi.cs lines 670-909 (immediate, locals,
watches, problems, drainfinalizers, history), 1200-1297 (console, inspect), 1620-1758 (compile,
type, mark), 1919-2060 (the dialog-watch internals and RememberRaisedDialogs body), 2350-2704
(doctor body, engine, native, windows, menus, outline), 2948-3034 (breakpoints, settings).
DebugServer.cs 260-460 (RecordRequest/RecordDuration bodies, ReadRequest, WriteReply, Dispose) and
460-1260 (the GDI capture code and the reply record definitions) were mapped but not read line by
line. DialogWatch.cs (231 lines) was not opened at all - the dismiss/safe-button logic that
AnswerBlockedRequest depends on is unexamined.

On the two questions I could not settle: the Debug boundary is clean by construction - every
Diagnostics file except Log.cs is wholly inside #if DEBUG, and the Release solution build in the
gate is what proves the call sites are guarded, since an unguarded PerfCounters or EngineCounters
reference would not compile. I found no Debug-only code leaking into a type that ships. The gate
step at tools/verify.ps1:288 searches the published Release DLL for four literals -
`__xlideConsole`, `unknown benchmark`, `the pending result was lost`, `debug-api-` - which live in
DebugApi.cs and DebugServer.cs, so it proves those two files contributed no strings; it does not and
cannot prove the absence of Debug-only code that has no distinctive literal, and I did not attempt
to establish that any such code exists. Two things I checked and dropped rather than report: the
ManualResetEventSlim disposed while a late host callback may still Set it (Set is documented not to
throw after Dispose, and Dispose holds the same lock the setter takes, so it is safe), and the
per-component COM reads in the `projects` route on the host thread inside the three-second deadline
(I could not measure it, so I make no perf claim). I also verified the file's own note that
seventeen cases are guarded on their arguments: it is accurate, eleven pool-side and six on-host.

**The rest of the shim and the core library.** Read in full: VbeMenus.cs, VbeCommands.cs, WatchReader.cs, LocalsReader.cs, ImmediateReader.cs,
GhostReaderThread.cs, Com/ComHandle.cs, Com/ComRuntime.cs, Com/HResult.cs, Interop/TypeLibrary.cs,
Core/Engine/EngineProtocol.cs, Core/Sync/SyncSettings.cs, Shim/Sync/SyncMessages.cs. Read in part:
EditorSurface.cs (member list plus 232-411, the Show* block), CodePaneTracker.cs (120-190 and
600-940), DebugServer.cs (400-500 only), WebView2Surface.cs (100-170, 925-1000, 1050-1085, plus the
#if DEBUG map), DialogWatch.cs (150-230), Win32.cs (1-30 plus targeted greps), Win32.Events.cs
(targeted greps only). Verified dead-code claims by grepping the whole repo excluding node_modules
and .git, and for the C#-only identifiers (HResult, Win32 constants, command ids) by whole-word
occurrence counts across src, tests, tools and installer; for message shapes and command names I
also checked ui/editor/src, docs/debug-api.md and tools/harness, since those dispatch on strings.
NOT READ, so this is not a clean bill for them: Engine/AnalysisService.cs (1247 lines),
Engine/EngineClient.cs (747), Engine/ProjectReader.cs, WebView/WebView2Interop.cs (688),
WebView/LoopbackPageServer.cs, Editor/TypeLibraryCatalog.cs (611), Editor/OverlayWindow.cs (786),
Editor/BrowserPalette.cs, Editor/ImmediateEvaluator.cs, Editor/HostChrome.cs,
Editor/FrameSubclass.cs, Editor/WindowEventHook.cs, Com/DispatchObject.cs (760, greps only),
Core/Sync/ModuleSync.cs (887), Shim/Sync/ModuleSyncService.cs (575), Diagnostics/Log.cs,
PerfCounters.cs (member list only), EngineCounters.cs, AddIn/ShutdownWatchdog.cs,
AddIn/XlideAddIn.cs, Interop/UiAutomation.cs, KillOnCloseJob.cs, FolderPicker.cs, ShimModule.cs,
HostApplication.cs, and Core/Editor/*.cs beyond their usage counts. Two things I looked for and did
not establish: I found no duplicated caret or selection conversion in the Editor folder (the
arithmetic lives in Core TextPositions and I did not find a second copy, but I did not read
EditorSurface end to end, so treat that as unchecked rather than clear); and on the Core-split
question I checked every Core type's callers and found ActiveLineHold and CodePaneCaption have
exactly one shim caller each - I am NOT reporting that as a finding, because both have unit tests
under tests/Xlide.Vbe.Core.Tests that the NativeAOT shim assembly could not host, which is what the
split buys. I made no performance claims: nothing I read gave me a path cost I could name against
what else runs on that path.

**The Monaco page.** Read in full: docktree.ts, dragcompass.ts, settings.ts, edgescroll.ts, build.mjs, package.json,
tsconfig.json, and the dialog open/dismiss regions of helpdialog.ts, sponsordialog.ts,
settingsdialog.ts, referencesdialog.ts, syncdialog.ts, shell.ts. Read in depth but not front to
back: bridge.ts (message unions 40-100 and 260-310, pending tables 405-473, request methods 700-940
and 1050-1062, the whole dispatch switch 1115-1345, the edit/sync/marker region 1490-1690),
workspace.ts (layout types, splitLeaf, dissolveEmptyGroups, tabAfter, the method inventory),
paneldocks.ts (pruneUnknown and the drag/zone sites), devsurface.ts (header, UiSnapshot, dialogsUp,
watchLongTasks), main.ts (imports, dialog wiring, the boot region by grep). Traced into the shim
where a page claim needed settling: EditorMessages.cs, EditorSurface.cs message switch,
AddInSession.DebugApi.cs settings route, tools/harness/module-sync.mjs, docs/debug-api.md.

CSS came back clean and I want that on the record rather than dressed up as a finding. I enumerated
all 199 distinct class selectors and all 102 id selectors in styles.css and tested each against
every .ts plus build.mjs's inline HTML, twice - once as a substring and once with token boundaries,
both giving the same 13 candidates. All 13 are false positives: drop-overlay-join/new and the five
drop-petal-* are built by concatenation at dragcompass.ts:51 and :149, split-column at
workspace.ts:805/826 and paneldocks.ts:438/869, tab-edge-start/end and toolbar-edge-start/end at
edgescroll.ts:109 from the className argument (workspace.ts:149 passes "tab-edge"), and
.glyph-margin is Monaco's own, styled as an override at styles.css:2532. The only unmatched ids were
hex colour literals. No @keyframes are defined at all, and no CSS custom property is declared
without a var() or TS reference. Every getElementById target in the page exists in build.mjs's
INDEX_HTML or is created in TS. All six settings keys have readers.

Two things constrained the dead-code haul and are worth knowing: tsconfig.json sets noUnusedLocals
and noUnusedParameters, so module-local dead code cannot survive `npm run typecheck` - dead code
here can only be exported-but-unused symbols, unreachable-at-runtime branches, or DOM/CSS
mismatches, and I checked all three. And I verified every exported runtime value in ui/editor/src
has at least one importer or in-module caller; there are no orphan exports.

What I did not settle. I found no perf finding I could substantiate and am deliberately not
reporting the one I chased: bridge.ts:527 posts selectionChanged on every cursor move with no
coalescing, and there is no debounce, throttle or requestAnimationFrame anywhere in ui/editor/src,
but I read the host handler at EditorSurface.cs:1332-1355 and it is a field assignment plus an
event, with StartAnalyseTimer gated on the line actually changing - so I have no basis for calling
it expensive and did not. I did not read explorer.ts, searchwidget.ts, objectbrowser.ts, menubar.ts,
typing.ts, format.ts, vba.ts or bookmarks.ts closely, so duplication and dead branches inside those
are unaudited. I did not audit the ~180 lines of styles.css that are compound-selector rules
(descendant and state combinations) for whether the combination can ever match - only that each
class token is produced somewhere. I did not measure devsurface.ts's minified contribution because
building is outside my authority; that is the one number in finding
devsurface-ships-on-a-wrong-number that is unmeasured, and it is marked as such.

**Engine wrapper, harness, build and release tooling.** Read in full or near-full: tools/verify.ps1 (the parts that name scripts, the whole -Live section,
the engine-freshness and variant-guard steps), tools/release.ps1, tools/Update-Page.ps1,
tools/page.ps1 head, tools/commands.txt, tools/harness/audit-routes.mjs (lines 1-175),
tools/harness/page-probe.mjs import surface, tools/harness/Get-Shot.ps1, .github/workflows/build.yml
engine and page test steps, engine/package.json, engine/build.mjs head, engine/src/dispatcher.ts
(the handle switch at 270-357, closeProject at 444-480, the member index),
src/Xlide.Vbe.Shim/Engine/EngineClient.cs connect path, docs/testing.md in full. Skimmed heads and
epilogues of every .mjs in tools/harness and all four engine/test files, plus the heads of the
fifteen Test-*.ps1 and the four New-*Fixture.ps1. Verified the dead-code claims with whole-repo
greps across src, ui/editor/src, engine/src, engine/test, tools, tests, docs, installer and .github
rather than by identifier alone, and reconstructed audit-routes.mjs's own route/method/corpus logic
to prove the four undriven routes rather than inferring them. NOT covered: I did not read
engine/src/navigation.ts (41 KB), protocol.ts (20 KB), moduleContext.ts, onType.ts or sync.ts line
by line, so my dead-code claim for engine/src rests on export-level call counts (every one of the 21
exports has at least one caller) plus the didClose route - unread protocol FIELDS that nothing
consumes are a real possibility I did not rule out, and finding them needs a field-by-field pass
over protocol.ts against the shim's EngineProtocol.cs. I did not read the bodies of
Test-DebugApi.ps1 (29 KB), Test-SplitWorkspace.ps1 (23 KB), Test-Churn.ps1 or module-sync.mjs (24
KB) past their headers, so overlap between those four and the .mjs suites at the check level is
unmeasured; my PowerShell-versus-mjs overlap findings are about the wrapper structure and the
discovery scaffolding, which I did verify. installer/build.ps1 was grepped, not read. I ran nothing:
no build, no gate, no Excel, no harness, so every timing or pass/fail claim here is read out of the
source and never observed executing.

**The add-in's own cost.** Read in full: docs/lessons.md sections 15, 16, 20, 21, 23, 24, 44, 45, 46, 48, 53, 59, 63 plus the
headings index; docs/status.md; the numbers and open-items sections of
docs/handoff-2026-08-10-1030.md; Diagnostics/PerfCounters.cs; Diagnostics/Log.cs; Com/ComRuntime.cs;
Engine/ProjectReader.cs; Core/Editor/ActiveLineHold.cs; Editor/GhostReaderThread.cs;
Core/Engine/EngineProtocol.cs TextPositions. Read the load-bearing regions of the big files rather
than every line: AddInSession.cs (Start/HostStartupComplete/EnsureSurfaceForFrame 1001-1230,
FollowActivePane 1490-1620, StartAnalysis 1416-1495, WriteModule 1853-1935, TypeLibraries 2320-2390,
PublishLocals/PublishWatches/UpdateDebugState 3430-3650,
UpdatePolling/ResyncFromModule/ProjectCount/PollDebugState 3690-3960,
OnLiveAnalysisDue/ShowModuleInSurface 5270-5450,
PublishModules/WorkbookSaved/ReadOpenModules/PublishProjects 5460-6010, PublishFindingsToSurface
6445, FindComponent/ProjectIdFromDisplay 6955-7075,
PublishMarkersToSurface/SurfaceBounds/placement/PoliceNativeToolWindows 7076-7470,
ReportEnvironment/ReportOpenProjects 7675-7740); EditorSurface.cs (Create 470-530, Follow 582-630,
ShowDiagnostics/ShowProjects/ShowModules 975-1030, FlushEdits 1198-1240, OnMessage head and the
contentChanged branch 1285-1470, ParseChanges/SingleLineTypedIn/ApplyEdits 1963-2050);
CodePaneTracker.cs (OnWindowEvent through ReadAllComponentNames, 140-620); OverlayWindow.cs (timer
ids, RunOnHostThread/DrainActions 200-300, WindowProc 470-560); WebView2Surface.cs
OnWebMessageReceived 805-850; DispatchObject.cs (Attach/GetDispId/GetProperty/GetString/GetInt32
55-130, GetObject/CallObject 286-310, GetItem 437-472); AnalysisService.cs
ReadProjectsAsync/AnalyseEverythingAsync 780-960; the perf and stats routes in
AddInSession.DebugApi.cs; ui/editor/src/bridge.ts onModelContentChanged 1595-1645 (to establish the
keystroke cadence and the 64,000-character threshold).

Not read, so not covered: WebView2Surface.cs beyond the message handler (browser start-up, SetBounds
internals, the loopback server), LoopbackPageServer.cs, VbeMenus.cs, VbeCommands.cs,
ImmediateEvaluator.cs, LocalsReader.cs and WatchReader.cs internals (only their call sites),
BrowserPalette.cs, TypeLibraryCatalog.cs, HostChrome.cs, OverlayWindow's paint path,
Interop/Win32.cs and UiAutomation.cs, Sync/ModuleSyncService.cs and Core/Sync/ModuleSync.cs,
DebugServer.cs, and the ~50-route switch in AddInSession.DebugApi.cs beyond perf and stats. The
startup budget in docs/status.md (about 2.1s) is therefore only partly attributed here: I
established that TypeLibraryCatalog is lazy and ReportOpenProjects is cheap, but I did not read the
WebView2 environment creation, HideReplacedWindows, HideNativeToolbars or PrepareGhostPalette, which
is where the remainder plausibly sits. No finding below rests on those.

No timings of my own: I was read-only and Excel was not running, so every duration in this report is
either quoted from the repo's own measurements in docs/lessons.md or explicitly marked unmeasured
with the measurement that would settle it. Nothing here is a claim that a path IS slow on the clock;
each is a claim about what work runs on which path and how often, which I read.

**The page.** Read in full or near-full: ui/editor/package.json, tsconfig.json, build.mjs (all 324 lines,
including the generated index.html and boot.js), typing.ts around the
content/cursor/recase/loop-sync paths, objectbrowser.ts boot and both render functions, explorer.ts
render/setProblemCounts, shell.ts renderPanel/setFindings, workspace.ts
render/editors/renderTabs/splitter drag, paneldocks.ts both splitter apply paths, main.ts lines
1-170, 290-360, 415-475, 955-1002 and 1225-1250. Cross-checked the host side where a page claim
depended on it: EditorSurface.cs contentChanged (1355-1430), ShowFindings (1159-1167),
DescribeTimings (2118-2149), BrowserPalette.cs window creation (60-120), AddInSession.cs BrowseTypes
(2475-2509). Verified three claims against monaco 0.56 in node_modules rather than assuming:
codeEditorWidget.js layout() and EditorDecorationsCollection.clear/set, elementSizeObserver.js
measureReferenceDomElement. Grepped all of ui/editor/src for layout reads, high-frequency handlers,
dynamic imports, whole-list rebuilds and monaco imports.

Not covered, and each could hold something: bridge.ts is 2,356 lines and I read roughly 400 of them
(the model-change, marker, decoration, selection and mouse paths plus the request helpers) - the ~50
message handlers in its switch and the thirteen setTimeout-guarded request/reply pairs around lines
700-1060 were not audited for redundant round trips, so the "message traffic" part of the brief is
only partly answered. devsurface.ts (1,483 lines) was sampled at its entry only; I established it is
eagerly bundled and shipped in Release but did not audit what watchLongTasks or its act/observe
handlers cost. styles.css (77 KB) was not read at all, so no claim about selector cost or paint.
searchwidget.ts, menubar.ts, format.ts, settingsdialog.ts, syncdialog.ts and the
docktree/dragcompass/edgescroll trio were greped but not read. I did not read the vendored spec
helpers under ui/editor/vendor/xlide-spec, so the cost of resolveLoopIteratorSyncEdit itself is
unknown and that finding sizes only the getValue around it. No build was run and no page was loaded,
so every size and timing here is either a byte count from disk or an explicitly unmeasured estimate
- the metafile finding exists precisely because the attribution I would have wanted does not exist
yet.

**The language engine and the checking loop.** READ IN FULL: engine/src/dispatcher.ts (all 1073 lines), engine/src/moduleContext.ts,
engine/src/main.ts, engine/src/semantic.ts, src/Xlide.Vbe.Shim/Engine/ProjectReader.cs,
tools/verify.ps1, tools/dev.ps1, .github/workflows/build.yml,
tests/Xlide.Vbe.Core.Tests/Xlide.Vbe.Core.Tests.csproj, xlide_vbide.slnx, engine/package.json,
engine/tsconfig.json.

READ IN PART: src/Xlide.Vbe.Shim/Engine/AnalysisService.cs (lines 260-400, 786-1060, 1110-1200, plus
a member listing of the rest - I did NOT read the ~30 per-feature Async wrappers at 409-780 line by
line, only their signatures); src/Xlide.Vbe.Shim/Engine/EngineClient.cs (lines 570-747 in full, the
rest via a grep of every await/serialize/lock, which is enough to establish the one-call-at-a-time
contract but not enough to have audited each payload builder); Diagnostics/EngineCounters.cs (first
80 lines); AddIn/AddInSession.DebugApi.cs (the perf and ui route bodies only, not the other ~48
routes); ui/editor/src/bridge.ts and main.ts (the semantic-token paths only); engine/build.mjs
(first 60 lines); engine/test/freshness.mjs (header only); ImmediateEvaluator.cs (a structural grep,
not a read); ../xlide_vscode/src/vbaProjectAnalysis.ts lines 106-182, read only to establish that
buildLiveVbaProjectIndex walks every module - nothing there is proposed for change.

NOT REACHED, and these are real gaps: engine/src/navigation.ts (1075 lines) and
engine/src/protocol.ts (571 lines) were not read at all, so nothing here covers the
rename/definition/references implementations or the wire types; engine/src/completion.ts, hover.ts,
signature.ts, onType.ts, outline.ts, search.ts, codeActions.ts, sync.ts were not read, so I cannot
say whether any of them re-scan a module a sibling feature just scanned; tools/page.ps1 got a grep
for its build commands only; tools/harness was not audited beyond audit-routes.mjs and
xlide-api.mjs's perf helpers; the shim's src/Xlide.Vbe.Core/Engine/EngineProtocol.cs was not opened.

CLAIMS I DELIBERATELY DID NOT MAKE. I did not report the pipe's one-request-at-a-time serialisation
(EngineClient.cs:661-674) as a defect: the engine's dispatcher is synchronous JavaScript on one
thread (main.ts:114-141 answers inline except for the async sync/plan), so pipelining would reorder
waiting, not add parallelism, and the existing counters already attribute the queueing correctly. I
did not report the CI 'Languages' step re-running tests the previous step ran; build.yml:153-160
designates it and prices it at about a second. I noticed but could not settle within budget whether
a CallAsync that times out after writing its request but before reading the reply leaves the reader
one line out of step for the rest of the session (EngineClient.cs:670-704 releases the semaphore in
the finally and marks nothing dead) - that is a correctness question rather than a perf one and
belongs to whoever owns the transport; I flag it here rather than as a finding because I did not
read the reconnect path.

TWO NAMED CANDIDATES, ANSWERED. xunit v3 as a direct executable: CONFIRMED, the exe is on disk at
artifacts/bin/Xlide.Vbe.Core.Tests/release/Xlide.Vbe.Core.Tests.exe and the project references
xunit.v3 3.1.0; DOTNET_ROOT is indeed needed and verify.ps1:44-47 already computes that path. Node
and dotnet chains concurrent: CONFIRMED WITH ONE DEPENDENCY - the shim csproj globs ui/editor/dist
(Xlide.Vbe.Shim.csproj:83), so the dotnet chain must follow 'page build', but the four node steps
after it are independent of the four dotnet steps. Neither saving is measured; verify.ps1's own
summary table already prints the per-step seconds that would size both.

