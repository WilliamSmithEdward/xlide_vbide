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

**Boolean arguments read the same way everywhere.** `1`, `true`, `yes` and `on` are true;
`0`, `false`, `no` and `off` are false; anything else, and anything absent, is the route's own
default. There were two rules until 2026-08-11 and they disagreed about the same text: six sites
took anything that was not the literal `0` as true, so **`perf?reset=false` cleared the counters**
and `native?text=no` asked for the text, while two other sites accepted only an affirmative word,
so `module?live=1` worked and `module?live=yes` did not. Both readings are defensible and no
caller can hold both at once.

| Route | Method | Arguments | Answers |
| --- | --- | --- | --- |
| `state` | GET | | shown module and project, debug mode, unwritten edits, engine up, frame and document-area handles and rects, **the frame's caption as the title bar actually reads it**, **whether the frame is on screen** (`frameVisible`, read off the window; it is where `frame?action=close` 's posted outcome is observed), palette open and visible, surface ready, DevTools port. `frameCaption` is read off the window, not out of the session's record of what it last wrote there, and that is the only reason it is worth having: the host rewrites that window underneath the add-in. Pressing Reset a second time does it, and the caption sat saying "Microsoft Visual Basic for Applications" with nothing able to observe it (2026-08-09). `engineUp` means the engine is ANSWERING, not that the session got as far as constructing the service that talks to it. It was the second thing until 2026-08-08, which made it true from start-up to shutdown whatever the engine did: killing the engine process left this reporting `true` while the editor drew squiggles from the last pass that had run. It is now false BEFORE the engine connects too, which is a real state a launcher should wait through rather than report |
| `windows` | GET | | every editor window: type, caption, visible |
| `native` | GET | `text=1` | the HOST's own editor, underneath the surface that covers it: the ACTIVE CODE PANE's module, project and caret, every pane it holds open, and - in the same reply - what the surface believes it is showing. Run, Step, Compile and ToggleBreakpoint all act on the native pane and its caret rather than on the page, so a disagreement means a Run that executes where the developer is not looking and a breakpoint on the wrong line, with nothing on screen to say so. Carries the pane's CONTENT and the surface's, each reduced the same way - line endings normalised, trailing blanks dropped - so a single changed character registers while the host's CRLF does not. Names agreeing is NOT parity: a surface holding an empty document for a module the host has 42 lines of passes every name comparison and shows a blank editor. `text=1` carries both texts for the run that fails. EVERY open pane carries the workbook's content and the surface's side by side, not only the active one: a background tab holds a copy the developer is not looking at, so a module written from outside while its tab sits behind another goes stale with nothing to notice until it is clicked. `parityAll()` on the client reports which |
| `engine` | GET | `module`, `text=1` | what the ENGINE is holding for a module, against what the surface holds: each text's line count, each reduced to a comparable form, and both texts themselves with `text=1`. There are THREE copies of a module in play, not two. The engine keeps its own live text per module, fed by the page's edits, and prefers it over the copy the project was seeded with whenever a request carries no source of its own. Every finding is measured in whichever it chose, so when a squiggle is drawn away from the word it is about, this is the call that says which side drifted. Nothing could ask it until 2026-08-08, and until then a misplaced finding could only be guessed at |
| `projects` | GET | | EVERY open workbook: name, id, component count, and whether the surface is showing one of its modules. `project` answers about ONE, the one named or the active one, so with two workbooks open there was no way to discover the other's name from the host at all: a probe either knew it in advance or read the page's tree, which is the surface's view rather than the object model's. Two workbooks holding a module of the same name is a designed case here and three separate defects have lived in it, so a suite that cannot name the workbook it means cannot test any of them. The language suite failed exactly there. Cheap on purpose: names and counts, not contents |
| `stats` | GET | | uptime, managed and working memory, handle count, GC counts, placement pass counts and timings, host-thread marshal count and timings, log lines, poll interval, message totals, and the COM wrapper counters. `comWrappersGivenBack` and `comWrappersDisposed` are counted SEPARATELY and must always be equal: the first is how many wrappers were handed back, the second how many were actually released. They diverged for the whole life of the product because the release was `(wrapper as IDisposable)?.Dispose()` and the wrapper is a `ComObject`, which does not implement `IDisposable` - so the cast failed silently and every wrapper went to the finalizer thread, which ends the host. `disposed: 0` against `givenBack: 1882` is what finally named it. See [com-wrapper-release.md](com-wrapper-release.md) |
| `log` | GET | `since`, `match`, `max`, `waitMs` | a slice of the shim log and the next byte offset. With `waitMs` it BLOCKS until a matching line appears, so an event can be awaited rather than slept for |
| `doctor` | GET | | the start-of-session checklist: shim and bundle build times, the page's own build stamp, engine and ghost readers attached, dialogs standing. **When the page never reported a stamp it names the CAUSE**, not the silence: boot.js installs ahead of the bundle and pushes any load-time throw to the host, so the finding reads "it THREW while loading: ReferenceError ... at editor.js:1:7" rather than "it may not have finished booting". It distinguishes a page that recorded nothing (still starting) from one that could not be asked (not responding), because those call for opposite next steps. `findings` is empty when nothing is wrong. An engine that is present but NOT ANSWERING is its own finding, and the dangerous one: the engine is started once and never restarted, so the findings on screen are from the last pass that ran and will never change again. That state was called healthy until 2026-08-08 |
| `perf` | GET | `reset=1` | recent raw placement and marshal durations, **and the analyzer's cost per method**: calls, time spent QUEUED behind another call, time spent ON the pipe, median, p95, and the individual calls over 120ms. The engine serves one request at a time, so a diagnostics pass over a large module delays every keystroke's completion behind it and a combined figure blames completions. `reset=1` forgets the engine figures first, so an experiment measures what it provokes |
| `journal` | GET | `lines` | one capture of a whole moment: state, dialogs, counters, recent log, recent page traffic |
| `drainfinalizers` | GET | | Forces a collection and WAITS for the finalizers, so a leaked COM wrapper fails NOW instead of minutes later in a stack that names nothing about what created it. A wrapper nothing disposed is released by the finalizer thread, and for the editor's objects that is an access violation the runtime cannot throw: it FailFasts and ends Excel. Run an operation, call this, and if it does not answer, THAT operation left a wrapper behind. It is a BISECTING tool, not a health check: the last route that tried to be both reported a clean bill of health with 8,734 wrappers pending, because it measured the heap rather than the outcome. The count still comes from `stats.comWrappersLive` |
| `history` | GET | | every request this door has served, a script that replays them, and **`routeCosts`**: each route's count, total and worst time, slowest total first. The door is the instrument every other measurement is taken through, so a route that has quietly become slow makes everything look slow |
| `assert` | POST | `that`, `value`, `timeoutMs` | states an expectation and waits for it, answering with what was actually seen |
| `messages` | GET | `last` | recent page traffic both directions, per surface |
| `problems` | GET | `module` | the analyzer's current findings. **Not the same thing as the squiggles.** These are the analyzer's findings for the project as last analysed; the underline a developer sees also comes from the LIVE analysis of unwritten text, which this does not carry. Measured 2026-08-08: an undeclared variable showed an error underline on screen while this answered an empty list. To check a squiggle, ask `ui` with `word` or `line`+`column` and read `at.squiggles`, which is monaco's markers - exactly what draws the underline |
| `locals` | GET | | the Locals panel's context and rows |
| `watches` | GET | | whether stopped, and the Watch panel's rows |
| `module` | GET | `name`, `project`, `live=1` | a module's text. Through the session's reader by default; with `live=1` the SURFACE's copy, which is what the developer has typed but not yet written back. That window is where every typing behaviour lives - smart Enter, comment continuation, auto-indent - so without it those could only be checked by eye |
| `capture` | GET | `window=frame\|palette`, `selector`, `pad` | a BMP of the window, through PrintWindow. With `selector` it is cropped to that element |
| `module` | POST | `name`, `project`, body = text | writes the module through the session's writer, with the baseline and engine corrections a host rewrite carries. Answers `error` when the write did not happen: a module that is not there, or text the editor refuses. It used to answer ok to both |
| `dialogs` | GET | | native dialogs standing now, with their buttons, how long the host thread has been quiet, and **`recentlyCleared`: what the guard took off the screen just before you asked**. That last field is there because this route observes something another part of this door actively removes, and an instrument that changes what it measures has to say so in its own answer: an empty `dialogs` means either nothing opened or something opened and was cancelled a second ago, and reading the second as the first has already produced a wrong conclusion here (2026-08-09, twice in one hour). Needs no host thread, so it answers while the editor is stuck |
| `ui` | GET | `waitMs`, `line`+`column`, `word` | the surface as the PAGE describes it: tabs with the labels the strip drew and the MRU order a close falls through, the tree's expansion and unfolded module, panes, page-side dialogs, what text has not arrived yet, the caret, **the main-thread stalls over 50ms** (which is what jank is, and which no host counter can see), and the model/document census a leak shows up in. `dialogs` is asked of the DOM by `aria-modal`, not read off a list, so a dialog added later is reported without anything remembering to add it. Reported from the fields the render reads, so it cannot mistake a stale render for the state, the find/replace widget's query, scope and match count, and the bookmark lines of the model on screen (read from its live decorations, which ride the text as it is edited, rather than from where the marks were put). With `line`+`column` or `word` it also answers `at`: the word there, the COLOUR actually painted on it (read off the rendered span, not derived from a token type and a theme map), its weight/style, and every marker covering it - which is the squiggle. **`statusNotice`** is what the status line is saying, empty when it is saying nothing: the status bar was in no snapshot at all until 2026-08-11, so the explanation a declined action gives could be driven and never observed. It carries both kinds - the ordinary notice that clears itself after five seconds, and a HELD one that stays until the condition ends, of which the only one today is the engine's cold start. `statusPosition` and `statusModule` complete the bar (2026-08-12): the caret readout ("Ln 4, Col 7") and the module name, read off the elements the render writes - the readout is the developer's own answer to where a Run would land, so debugger-features holds it to the native caret while stopped |
| `act` | POST | `do=closeActive\|answerCloseConfirm\|search\|bookmark\|format\|backspace\|hover\|completions\|signature\|quickFixes\|definition\|rename\|activate\|cycleTab\|split\|expandWorkbook\|unfoldModule\|treeMenu\|treeAdd\|menuBar\|chooseMenuItem\|answerRemoveConfirm\|renameModule\|settings\|sponsors\|closeDialogs\|key\|focusEditor\|editorAction\|editProperty\|toolbar`, plus that action's arguments | drives the surface through the methods a click reaches. Answers `{did, detail}`; `did: false` means the page declined or the action did not land, which is an answer and not an error. **`closeActive` reports the OUTCOME, not the request**: closing a module with unsaved changes raises a Save / Don't Save / Cancel box and the tab stays until it is answered, so `did` is false and `detail` names what is standing. `answerCloseConfirm` with `answer=save\|discard\|cancel` answers it \| `search` takes `query` (plus `scope`, `matchCase`, `wholeWord`, `replacement`, `run`) or `close=1`, and TYPES the query so the widget's own input handler runs - assigning an input's value raises no event and the widget would never search. **TYPING IS NOT SEARCHING outside module scope**, and `run` is the press that follows it: `findAll`, `next`, `previous`, `replace`, `replaceAll`, each going through the method that button's own click listener calls. The `input` event only searches when the scope is `module`, so without `run` a project search types a query and searches nothing - and it answered `did: true` while doing so until 2026-08-11, agreeing with a `matches` field that is structurally 0 for every non-module scope. **Read `scopedMatches`, not `matches`, for any scope but module**; `matches` and `current` describe the module-scope engine only. `replaceAll` rewrites text across every module of the project and had no driver of any kind before this. `bookmark` takes `which=toggle\|next\|previous\|clear` - **not `do`**, which is the action selector itself; an argument of that name overwrites it and the door answers "unknown action". `format` runs Format Module, or Format Selection with `selection=1`, and WAITS for it. `dock` takes `pane` and `side=left\|right\|top\|bottom` and moves a pane through THE METHOD A REAL DROP CALLS. The docking gestures had no coverage of any kind before it: `layout` could report the arrangement and `layout?reset=1` could restore it, and nothing in between could move anything, so the whole surface was drivable only by hand. Synthesising the drag was never the answer, because the drop arms on pointerdown and completes on pointerup against a compass that follows the pointer, so a synthetic sequence tests the synthesiser. Answers `did: false` and names the pane when there is no such pane, rather than pretending. `backspace` presses the key `times` over, through `deleteLeft`, which is what the key is bound to: the `type` route cannot express one, because it drives `trigger("keyboard", "type")` and that only ever inserts. It is worth being able to send because in a line's LEADING whitespace Backspace does not delete one character: with `useTabStops` on it takes back a whole indent level, which is the behaviour that makes indenting with spaces bearable now that VBA's refusal to store a tab has removed the choice. Answers the caret and the line's text in `data`. `hover`, `completions`, `signature` and `quickFixes` take `line`+`column` or `word` and answer in `data` by calling the PROVIDER OBJECTS monaco calls, with the arguments monaco passes - so they answer what the developer's editor would, including each provider's refusal to answer for anything but the host-active module. Not the bridge request underneath (that skips the refusal and reports healthy while the screen is dead), and not by scraping a widget (monaco reuses an open one, so a scrape returns the previous answer in every state). `definition` answers where F12 would go WITHOUT moving the caret, which matters because moving it is what cancels a symbol-navigation request. **`rename` CHANGES STATE**: it is the call F2 makes, so the host applies the edits across every module using the symbol and the surface republishes - undo with `undoRename`. It answers the provider's own refusal, word for word as the rename box shows it. `treeMenu` takes `module` or `workbook` and RIGHT-CLICKS that row, answering the menu it produced as a ` \| ` separated list, which is the only way to those menus: they hang off a DOM listener with no method behind them, and what a menu offers is a claim about the class of the thing clicked, so reading it is the test. A collapsed workbook has no component rows, so `expandWorkbook` comes first. `chooseMenuItem` takes `label` and fires `pointerup`, which is what the menu listens for; a trailing ellipsis need not be spelled. **`answerRemoveConfirm` is the only thing that removes a component through the product path**, with `answer=remove\|cancel`, and cancel is what anything unrecognised gets. `treeAdd` takes `workbook` and presses that row's plus; the plus is only VISIBLE on hover and no script can provoke :hover, so this presses the control rather than the pixel and leaves whether it appears to a probe that can move a pointer. `menuBar` opens the surface's own menu, the wrench, and is the one menu whose items come from the HOST - so it waits for them and says so when they never arrive, because an empty menu and a menu that never came are different failures. `chooseMenuItem` drives all three: the menu bar and the context menus both activate on `pointerup`. **`renameModule` is the product's rename**, taking `module`, `newName` and an optional `workbook`: the one the tree's Rename item runs, which renames the component AND every mention of it across the project, and answers how many mentions in how many modules. It is not `component?action=rename`, which is the fixture primitive and leaves every reference pointing at a name nothing answers to. Nothing could drive this until 2026-08-10, because the only route to it was a person typing into `window.prompt`. **`editorAction` runs any registered editor action by `id`**, awaited when it answers a promise. Several of this product's features are only an `editor.addAction` reachable from a context menu or a key, and until this existed the door covered them by calling the shim underneath, which proves the operation works and not that the menu item does: `editorAction id=xlide.undoRename` is the Undo Rename item itself, where `undoRename` is the shim's own undo. Undo and redo are registered as COMMANDS rather than actions, so `getAction` finds nothing for them; use the `undo` action. **`editProperty` takes `name` and `value` and CHANGES REAL COMPONENT STATE** through the object model's own property bag - `(Name)` renames the component, and there is no undo behind it. It edits whichever component the panel is showing, so open or select one first and read `ui.properties` to see which. It answers the HOST's outcome, not the post: the panel's controls set the row as they send, so the value reading back proves only that it was asked for, and the act waits for the host to republish before comparing. A refused write answers `did: false` with what the row actually holds and whatever the status line said about it. Together with `ui.properties` this closes the last user-visible surface that had no api presence in either direction. **`toolbar` presses a strip command by `command` id** - `openSync`, `objectBrowser`, `openPanes`, `openHelp`, the indent and fold cluster, about thirty of them - through the button's own click, which is why a disabled one is refused as disabled and says whether it needs break mode. With no argument it answers the strip's contents in `data.commands` as `{id, label, disabled}`. This is the driver half of the sync dialog, the Object Browser and the Panes menu: `press` is a keyboard key and the `command` route below is the NATIVE editor's command by name, so until 2026-08-11 the page's own command surface could only be reached by finding a button by its `data-command` attribute and clicking it, which three harness files did and which keeps passing after the button is renamed, disabled, or left out of the build because its editor action was not bundled. **`settings` with `key` and `value` changes a setting THROUGH THE PAGE**, which is a different path from the `settings` route below and the only one that can carry this class of defect: the route replaces one field of the stored settings on the host side, the page posts the whole object as it believes it to be. Its post was one field short of the six for months, so every change through the dialog reset the developer's sync planner. Bare, `settings` still just opens the dialog |
| `trip` | GET | `what=pagecall`, `n` | times a round trip ACROSS the boundary, wall clock from asking to observable. `bench` times the page and `perf` reports the host; this is the crossing neither measures. Read `pagecall` alongside the rest, since it is the floor every other figure contains. **A route body holds the host thread, and a message POSTED to the page is delivered by that thread's pump, so nothing whose effect arrives that way can be observed from inside a request at all** - those are measured across requests, in the client (`tripCaret()`) |
| `eval` | POST | `surface`, `waitMs`, body = script | runs script in the live page and returns its result as JSON. A PROMISE is awaited, so `(async () => ...)()` works. Prefer `ui` and `act` where they fit |
| `await` | POST | `surface`, `waitMs`, body = predicate | polls a predicate IN the page until it is true, answering `met` and how long it took. One request instead of a caller's poll loop |
| `layout` | GET | | the whole visible arrangement: dock sections and their groups, editor groups and their tabs, sizes, open documents, whether a drag is live |
| `layout` | POST | `reset=1`, `waitMs` | puts the arrangement back to the default and waits for the page, so a probe that dragged panes about can clean up in one line |
| `inspect` | GET | `selector`, `styles`, `rules=1`, `max` | elements matching a selector: box, classes, hidden, the computed values of the styles asked for, and which CSS rules claim them |
| `bench` | GET | `what=tabswitch\|layout\|type`, `n` | times a scenario in the page and answers min, median, p95, max, and the raw samples |
| `console` | GET | `last` | what the page said to itself: a ring of console lines, installed by **boot.js ahead of the bundle** rather than at page ready. That ordering is the whole value: a bundle that throws while its modules initialise never reaches ready, and the ring used to be created there, so this route answered `{"installed": false, "lines": []}` at exactly the moment somebody was asking why the screen was blank (2026-08-09). The JavaScript context survives a module that throws, so the ring can be read out of a page that never booted. It carries uncaught errors and unhandled rejections as well as console output |
| `reload` | POST | `waitMs` | reloads the page and waits for it to come back, answering with the bundle stamp it is now running and whether that is behind the one on disk |
| `dismiss` | POST | `button`, `caption` | clicks a dialog button by name. Explicit, so it will press OK if asked |
| `guard` | POST | `on=true\|false`, `forget` | turns the dialog guard on or off and answers what it has cleared. While ON, a NOTICE this door did not raise is cleared too - see below. Off by default |
| `compile` | POST | `waitMs` | compiles the project and answers its errors as DATA, clearing the modal it raises. **`compiled` means no dialog appeared within `waitMs` AND the command actually ran** - the editor has no positive report to give, it answers a compile with a modal or with nothing. `started` is the second half and the reason this is two fields: a greyed Compile item ran nothing, so no dialog appeared, so the reply said the project compiled. The debugger suite reads this as its precondition, which made every check behind it a test of the dialog guard |
| `documents` | GET | | the documents the surface holds TEXT for, with line counts, unwritten flags, and which is active. Not the same list as the tabs |
| `sync` | GET/POST | `direction=export\|import`, `action=plan\|apply\|settings\|browse`, `folder`, `mode`, `select=checked\|all`, `project`; POST body names the rows, one id per line | import and export, through the SAME call the dialog's button goes through, so a plan read here is the plan drawn there and an apply leaves the project where the button would have left it. Without `action` it answers the plan and changes nothing. `settings` reads what the project remembers, and writes it when a value is named. `browse` raises the system's folder chooser and BLOCKS until somebody answers it, so it is the dialog's, not a harness's |
| `component` | POST | `action=add\|remove\|rename`, `kind=1\|module\|standard`, `2\|class`, `3\|form`, `name`, `newName`, `project` | adds, renames or removes a component FROM INSIDE, so a fixture can be built without "Trust access to the VBA project object model". `name` is read back from the component; a refused name adds nothing. A kind that is neither a number nor a word it knows is REFUSED, not defaulted. **`remove` runs the product's own removal**, the one the tree's Remove item runs, so it also drops the page's unwritten edits for that module, the baseline this session last wrote, and its breakpoints - the bare COM call it used to make left all three behind, pointing at a module that no longer existed, and a harness removing a component left a different machine than a developer removing the same one. It skips only the confirmation, which is the page's; a document module is refused here as it is there. **`rename` now adopts the rename the way the Properties grid does**, re-keying the write baseline and the breakpoint record onto the new name: setting the name on the component renames it in the collection and nowhere else, so a breakpoint the developer could see was one this session could no longer find. It is still the fixture primitive and still renames nothing that MENTIONS the module - that is `act renameModule` |
| `menus` | GET | `path` | THE EDITOR'S OWN MENUS, with their ids, including every control the surface suppresses. `path` is the position chain, comma separated; absent means the menu bar. `suppressed` is this session's answer about each control, so the suppression table can be read back rather than inferred - it is a list of magic numbers, and before this the only way to learn one was to enumerate the bar by hand and write it into a comment. `path=900` is the composed **xlide** menu rather than one of the editor's, and answers each synthetic position with the id of the real control it resolves to, which is the one place that mapping is visible |
| `project` | GET | `project` | what the VBA project CONTAINS: every component with its kind, line count, and whether a pane is open on it, plus the project's execution mode. The object model's own answer, read from inside |
| `outline` | GET | `module`, `project` | a module's procedures, from the analyzer: name, kind, line. For asserting on shape without parsing the text again in the caller |
| `pane` | POST | `action=open\|close\|closeNative`, `module`, `project`, `answer=save\|discard` | opens or closes a module's TAB. A close goes through the same gate the tab's own X uses, so unwritten edits raise the question unless `answer` settles it. An open answers `error` when nothing carries that name; it used to answer ok and do nothing. **A close answers `{closed, detail, awaiting}`** for the same reason: a save that would not save, a revert the module refused, and a confirm now standing on screen all left the tab where it was and all replied ok. `awaiting: "confirm"` is the third answer, neither success nor failure - answer it with `act answerCloseConfirm`. `closeNative` is the OTHER direction and **a designated deviation**: it closes the HOST's own pane window through the editor's pane list, which is the developer's click on the native child's close box - the direction the 2026-08-04 dead-tab defect lived in, where a HIDDEN pane's close fired no Changed and the strip kept the tab. The native close box asks no unwritten-edits question (that question belongs to the page's tab X), so none is asked; the tab disappears because the tracker notices the window go. Until this existed the host-originated close could only be produced through `Application.VBE` from outside, which project trust gates |
| `settings` | GET/POST | any setting name | the developer's settings, and a POST changes only what it names. the page's own update takes the whole object. `syncEngine=xlide\|builtIn` chooses which planner decides an import or export |
| `palette` | POST | `action=hide` | puts the Object Browser palette away, the way its own close box does: its WM_CLOSE handler HIDES with state intact, so the next summons (the `objectBrowser` command, which means summon and not toggle) presents the same page. Answers `{did, detail, visible}`; `did` is false when no palette exists yet. Existed from 2026-08-12, because every probe that summoned the palette either left it standing in front of the next probe or reached for a window message from outside the process - and the summoning command cannot put it away |
| `frame` | POST | `action=close\|show` | the editor window itself: the one pair of gestures every developer makes daily and no route could produce. `close` POSTS the developer's own X click (`SC_CLOSE` delivered by the pump after the reply), so the editor runs its whole close path - the hide, the palette follow, the placement retreat - outside the request; the reply's `visible` still reads true and **the outcome is observed on `state.frameVisible`**, the rule every posted effect on this door lives by. `show` reaches the same end state as Excel's Developer > Visual Basic button, synchronously, so its reply is the outcome. Closing the editor shipped a crash once (lesson 27); until this route the only automated driver was a probe sending window messages from outside |
| `undoRename` | POST | | puts the last rename back: every module it touched and the component's old name. The editor's own undo cannot, because a rename spans modules and the undo stack is per model. Answers `{undone, from, to, modules, stopped}`, where `stopped` is the half that matters: an undo can restore four modules and be refused the fifth, leaving the project in neither state. It answered a bare true until 2026-08-11, and the sentence saying which of those had happened went to the page under a request id this route invented, where no caller could read it. **A DESIGNATED DEVIATION**: it reaches the session's undo directly, so it proves the operation is reversible and NOT that the context-menu item works. Drive the menu item with `act editorAction id=xlide.undoRename` when that is the question |
| `breakpoints` | GET | | every recorded breakpoint as `{module, project, lines}`, and the project mode. **`project` is part of the identity**: two open workbooks can each hold a module of the same name, and the record was keyed by the name alone until 2026-08-08, so a breakpoint set in one was reported against the other and a run that should have stopped did not |
| `type` | POST | body = text, `waitMs` | types into the editor through its own KEYBOARD pipeline (`trigger("keyboard", "type")`), so smart Enter, comment continuation and auto-indent all run. A `\n` is sent as a real Enter, because that keystroke is what the block layout hangs off, and the text is typed in segments with a turn of the event loop between them - smart Enter defers its work to a microtask, so a script that types a newline and the next line synchronously computes the continuation against a line that already has the next one on it, which is how this route reported comment continuation broken when it was not (2026-08-07). Setting the text instead would go around every handler that makes typing feel like anything. Use `act press` for a single Enter, Tab, Backspace, Delete or Escape |
` is an Enter, and each segment is given a turn of the loop - typing has gaps, and a script without them tests nothing |
| `mark` | POST | `text` | writes a labelled line into the shim log and answers the byte offset it landed at, so `log?since=` can return exactly the slice a step produced |
| `command` | POST | `name`, `keep` | runs an editor command by name (`VbeCommands.ForName`). `keep=1` exempts any dialog it opens from the guard below. **`ran` is the editor's answer, not the door's**: it used to be the constant true, because the session method that executes a command computed the answer and returned void, so pressing Step Over outside break mode was reported as a step. `detail` says why a false came back - `not present in this host`, `currently disabled`, or the exception's name - because those want different responses from a caller |
| `caret` | POST | `line`, `column`, `module`, `project` | puts the caret there, navigating first when a module is named. **A navigation that finds no pane is an `error` now**, not an ok: it left the caret in the module already shown, and a caller's next Run, Step or breakpoint then landed there instead |
| `breakpoint` | POST | `module`, `line`, `project`, `state=on\|off` | goes to the line and sets, clears, or toggles a breakpoint. `ran` is the record READ BACK after the toggle, so a line the editor will not carry one on answers false and says so; it used to answer the state that was asked for. A navigation that did not arrive answers `error` rather than toggling against whichever module was already shown |
| `immediate` | POST | `text` | evaluates a line in the Immediate window and answers WHAT IT CAME TO: the result text and whether it failed. It used to answer `{ran: true}` without waiting, so a caller learned that an evaluation had been ASKED FOR and nothing else, and what the expression evaluated to went only to the page. That is the reason this panel had a route and no suite: nothing could read what it said. The same rule the rest of this door follows, where `closeActive` reports whether the tab actually closed and `compile` answers errors as data. WITHOUT `text` it reads instead, answering the whole window as it stands. `ran: false` means the evaluation did not finish inside the ten second wait; the line still went in |
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
trip per tick and invites a `Start-Sleep` chosen by guess - the loops written during the
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
needed a dozen ad-hoc `eval` measurements per question - and it earned itself the day it
landed, answering "why is the Problems pane on the left?" in one call. (It was a probe that
had dragged it there and not put it back.)

`reload` reloads the page and WAITS for it to come back, answering with how long that took,
the build stamp the page is now running, and whether that stamp is behind the bundle on
disk. The manual version - reload, sleep a guess, hope - was run a dozen times in one
afternoon, and a guess that is too short reports on the page that is going away.

`layout?reset=1` (POST) puts the arrangement back and waits for the page. A probe that drags
panes about is testing the right thing and leaving the wrong thing behind - the layout is
persistent state, and clearing its storage key does not undo what the page already holds in
memory.

### Seeing why the page looks like that

`inspect` answers with the elements a selector matches: their boxes, classes, whether they
are hidden, the computed value of any styles named, and - with `rules=1` - which CSS rules
claim those properties, spelled out.

```powershell
Invoke-RestMethod "$api/inspect?selector=.dock-split&styles=align-items,display&rules=1"
```

The rule list is the point. This page shares a document with a large bundled stylesheet, and
a structural class of ours once inherited `align-items: baseline` from an unrelated rule,
collapsing every pane to its tab strip's height. It read as a flex bug in our own code and
cost an hour; the loop that eventually found it - walk every stylesheet, keep the rules this
element matches - is now this route.

`console` answers with what the page said to itself. Only UNCAUGHT errors reach the shim log,
deliberately, because forwarding every line would drown it - so a handled `console.error` or
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
is NOT the document area it is a child of - the surface draws the menu bar and toolbar too,
so it is taller - and a first cut that used the parent landed tens of pixels high, returning
the toolbar when asked for a pane header.

### Numbers for what a developer feels

`bench` times a named scenario in the page and answers min, median, p95, max, and the raw
samples. The counters elsewhere say what the HOST spent; this says what the surface costs,
which is where the risk moved when the workspace learned to split and dock.

| `what` | measures |
| --- | --- |
| `tabswitch` | putting another open document's model on screen - the live-model claim, in milliseconds |
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
is answerable for - including the routes that need no host thread, which is where the sweep
belongs, because `dialogs`, `dismiss` and `guard` are exactly what a caller reaches for while
something is stuck. It presses one SAFE button, and what counts as safe depends on what is
being asked:

- A **notice** - every button an acknowledgement (OK, Help, Close, Continue) - is answered with
  OK. It is reporting, not asking, so pressing it decides nothing. This case used to be
  missed: the policy was Cancel/Close/No only, so a compile error offering OK and Help matched
  nothing and stood with the host thread behind it.
- A **question** is only ever declined: Cancel, then No, then Close. Never OK, Yes, Save,
  Delete, or Run - a dialog nobody read must not be agreed with.

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
`shownModule`, `nativeModule`, `noDialogs`, `localsHas`, `watchHas`, `problemFree`,
`responsive`. It answers with what it SAW as well as whether the claim held, which is the half
a bare false leaves out - `{"held":false,"claim":"shownModule","expected":"NoSuchModule","saw":"BrokenModule"}`
diagnoses itself.

`shownModule` and `nativeModule` are deliberately two claims. The first asks the surface what it
believes it is showing, which is this session's own record; the second reads the host's
ActiveCodePane, which is what Run, Step and ToggleBreakpoint actually act on. Asking the first
about where a command will land is asking our record to confirm our record.

**A NAME THAT IS NOT A CLAIM IS AN ERROR**, answered immediately. It used to be polled like a
claim that had yet to come true, so a typo cost the full ten second timeout and then answered in
the same shape a real failure answers in, which reads as the product being broken.

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
