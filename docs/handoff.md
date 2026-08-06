# Handoff

Everything needed to pick this up cold. Rewritten at the end of the session that made the surface
the whole visible editor; the previous version described a foundation, this one describes a product
with a short list of holes.

Read this, then [lessons.md](lessons.md) for the long-form findings with evidence,
[architecture.md](architecture.md) for the design, and [decisions.md](decisions.md) for choices
that would be expensive to reverse.

## START NEW SESSION HERE — 2026-08-04 evening

**2026-08-05 LATE — THE WATCH PANE IS VERIFIED AGAINST A REAL WATCH (the last unverified
seam of the debug panels).** The developer's manual test showed "No watches" with a watch
defined; the reader was reading cleanly and parsing zero rows. MEASURED (design-mode UIA
dump of the watch ghost — safe there, the reader only wakes during break polls): a real
watch row's accessible name has NO "Expression" header word — " counter Value &lt;Out of
context&gt; Type Empty Context BreakProbe.BreakHere " — the leading space is the watch-type
icon column's empty cell. WatchReader.ParseRow now anchors on the first " Value " (watch
expressions can contain spaces, unlike locals) and the last " Type "/" Context " pair.
`Test-WatchPanel.ps1` is the standing PASS/FAIL probe and it PASSES: it drives the native
Add Watch dialog (command 1820) from a helper process — the modal blocks Execute — typing
the expression as WM_CHAR keystrokes into edit 4853 and clicking OK (control 1). TRAP:
the dialog answers "Empty watch expression" to text planted by WM_SETTEXT; it only
believes keystrokes. The probe breaks, steps to change the value, and asserts from the
LOG ONLY (a second non-empty push is the tracking proof; never dump the ghost mid-break,
lesson 33). Watches are session-scoped: they die with Excel, and deleting a watch's
module drops the watch.**

**2026-08-05 NIGHT — ONE SEARCH UI (developer: "move the search pane into the search popup
window, so we're not duplicating UIs").** The bottom panel's Search tab is GONE and so is
Monaco's find widget from view; `ui/editor/src/searchwidget.ts` is the one search UI — a
floating widget where Monaco's find sat (absolute in #container, z 40). MODULE scope is
live: find-as-you-type via model.findMatches, painted with Monaco's own findMatch/
currentFindMatch decoration classes plus overview-ruler marks, Enter/F3 cycling, "n of m"
counter; Replace All is ONE executeEdits (one undo step, flows the normal didChange path —
better than the panel's old host-side ReplaceLine for the shown module). WORKBOOK/ALL
scopes use the panel's old engine protocol UNCHANGED (search/replaceAll/searchResult ids,
monotonic acceptance) with the grouped results rendered inside the widget; rows navigate
the problem-row route. Actions registered by the widget: xlide.search.open (Ctrl+F),
.replace (Ctrl+H), .workbook (Ctrl+Shift+F), .next/.previous (F3/Shift+F3), Escape claimed
in-editor behind a context key (xlideSearchOpen). TRAP MET: the toolbar keeps only
commands that resolve as actions AT BUILD TIME, so the widget must be constructed BEFORE
the Shell (it registers the actions the toolbar's find/replace buttons name). KNOWN
RESIDUE: Monaco's own find cannot be unbundled — the multicursor feature imports
findController.js, which self-registers actions at import — so the F1 palette still lists
Monaco's Find/Replace entries and they open the native find widget; our keybindings
outrank Monaco's (verified: Ctrl+F opens ours, Monaco's stays hidden) and Ctrl+D
multicursor still works. Page-only change; shell.ts lost its whole search section
(ShellHandlers.search/replaceAll, ShellSearchMatch, showSearchResults), bridge routes
searchResult to bridge.searchWidget. Demo-verified end to end: tab gone, live count 1 of
3 → cycle → scope flip clears decorations, engine answers render (2 rows), module Replace
All 3-in-one-undo, Escape closes and refocuses, toolbar buttons open the widget.
FOLLOW-UP same night (developer request): a Find All button in the find row — module scope
renders every match as clickable line-and-preview rows in the results region (clicks
select and reveal, the counter tracks, the table rides every re-find live and resets when
the query empties; capped at 500 rows with a note); wider scopes run the engine search,
restoring the old panel's explicit button. And a left-edge chevron expander folds the
widget to find-only (the fresh state); Ctrl+H forces it open with replace focused, the
toggle is sticky otherwise.**

**2026-08-05 EVENING — LOCALS TRACKING IN BREAK IS FIXED (lesson 33 has the full story).**
The developer's report: in a live break the Locals panel said "Not stopped. Variables
appear here in break mode." while the code pane showed the stopped line. Root cause found
by splitting in-proc vs out-of-proc reads: `UiVariant`/`ComVariantBlock` were declared
Size = 16, but an x64 VARIANT is 24 bytes — every GetCurrentPropertyValue wrote 8 bytes
past the buffer. Release layouts survived it; the SAC-forced switch to Debug publishes
(2026-08-05 morning) moved a live slot into the overhang and every read died in a frameless
NullReferenceException ("skipped 4 unreadable element(s)", 24 sessions straight). Fixed:
both structs are Size = 24. THREE MORE THINGS SHIPPED WITH IT: (1) the readers moved off
the host thread onto `Editor/GhostReaderThread.cs` — an MTA thread owning both LocalsReader
and WatchReader; the host only RequestRead()s and reads published snapshots, never blocks
(bounded 500ms join on dispose, readings cleared at break exit so a new break starts
empty); in-proc UIA clients on the provider's own thread were never a supported shape.
(2) Honest panel states: `setLocals` gained a `stopped` flag (shim record + EditorSurface +
bridge + shell + demo); in a break with nothing readable the panel says "No variables to
show.", never "Not stopped". The readers also kept the discarded-patch armor, rebuilt: per-
element fault isolation, stage-named first-failure logging, 5s backoff, recovery lines.
(3) The context strip: the context box is a PANE to UIA (not an Edit — that was the old
"context reads null" nit) and reads EMPTY in-proc even though out-of-proc sees
"VBAProject.Module.Proc"; the reader accepts a dotted pane name and normalises empty to
null so the strip hides rather than showing blank. VERIFIED: quiet-break probe (no external
COM) — break enters, honest empty first tick, `locals: 3 row(s)` next tick, break holds;
`Test-GhostLocalsPanel.ps1` PASS with three distinct row pushes tracking two steps
(counter 1→2, label alpha→beta) and a clean clear at exit; 72/72 Core tests; page rebuilt
and mirrored. TRAP FOR PROBES: dumping the ghost's UIA from OUTSIDE while the in-proc
reader is alive can RESET THE PROJECT mid-break (measured twice) — never do both at once.
Watch rows' accessible-name shape against a REAL watch remains unverified (needs the
native Add Watch dialog — developer's first watch is the test). THE BIGGER MEANING: the
2026-08-05 crash storm that forced the reset to v0.1.0 is ROOT-CAUSED by this same
variant — the branch `post-v010-experiments` (tip 269a2b3, the debug api + CDP doors +
eleven routes) concluded from those crashes that the provider was unsafe from anywhere
in-process and only an out-of-process reader could work; that conclusion is RETIRED. The
crashes were our own 8-byte overwrite, differently dressed per stack layout. When
re-landing the branch, rebase it onto this fix first, and re-land commit by commit with
the developer testing between steps as planned.**

**2026-08-05 THE OBJECT BROWSER ENDGAME — A FLOATING XLIDE PALETTE (developer-chosen,
confirmed live; supersedes every OB delta below).** After the hole shipped, the developer
asked for theming and true outside-the-canvas floating — both impossible for the native
window — and chose, by explicit question, the xlide typelib browser rehomed into a real
top-level window of ours. The native Browser retires for good; the canvas is purely xlide
with NO exceptions. SHIPPED AND VERIFIED LIVE: F2 / the toolbar button intercept command
473 BEFORE the native execute and summon `Editor/BrowserPalette.cs` — a registered
"XlidePalette" WS_OVERLAPPEDWINDOW owned by the editor's frame, DarkenTitleBar'd, hosting
a SECOND WebView2Surface (Start grew an entryQuery parameter) navigating the same bundle
with `?view=objbrowser`; main.ts branches to bootObjectBrowserPage() (drops the editor's
#shell skeleton), and objectbrowser.ts is now that page — its own PaletteHost transport
speaking id-matched obLibraries/obTypes/obMembers plus navigate and close. The window
hides on close and re-presents on the next summons, state intact; Escape posts close.
THE WHITE-WINDOW LESSON: the surface creates its browser INVISIBLE until Reveal (the
editor's loader discipline); calling Reveal right after Start is a call on a controller
that does not exist yet, and the window stays white forever — the palette reveals on the
page's FIRST MESSAGE instead. Session side: BrowseLibraries lists open projects (kind
"project", display-name currency) ahead of the typelib catalog (kind "library");
BrowseTypes answers a project with its VBComponents; BrowseMembers scans the project
module's CodeModule text by declaration line (Sub/Function/Property collapsed per name,
Const, Field with comma splits, Event, WithEvents, Enum/Type blocks, Declare) so project
members carry REAL LINE NUMBERS — ObMemberRow grew `line`, ObLibraryRow grew `kind` —
and navigate goes SetForegroundWindow(frame) + GoTo. The editor surface's ObX plumbing is
REMOVED (the palette owns the protocol; bridge requestOb* and the editor demo's canned ob
answers went with it). The cutout machinery is EXCISED from AddInSession
(PlaceObjectBrowserWindow, FindObjectBrowserChild, ObjectBrowserCutout, WindowRectIn,
_watchingCutouts, the 200ms tier and every branch); OverlayWindow.SetCutouts stays as
dormant API; the police pass hides type 2 again. DETAILS PANE (developer request): the
one-line strip became a stacked pane under both lists behind its own splitter (pointer
drag + ArrowUp/Down, 48px..60% clamp) — monospace signature, "Member of X.Y, line N"
context, description paragraph. UI-TEXT RULE (developer, STANDING): product UI strings
use plain ASCII punctuation — no em dashes, no ellipsis glyphs ("Loading...", periods,
commas, parentheses); source comments keep house style. VERIFIED LIVE: palette opens dark
and floating from the toolbar click; scratch.xlsm project (4 modules, members from code:
"Sub Test(), line 14"); the developer live-drove VBA (35 types, ColorConstants,
Conversion), stdole (10), Office (318). NOT yet observed: the navigate leg (synthetic
double-click does not register in Chrome; the wiring is the proven GoTo path — the
developer's first double-click on a project member is the test). Parity backlog (told):
global cross-library search results pane, Back/Forward history, hidden-members toggle,
F1 help.

**2026-08-05 POSTSCRIPT — THE TWO-BUILDS MORNING, TRUE MECHANISM (corrected by the
developer's own diagnosis).** After the palette shipped, the developer kept seeing the
OLD in-canvas browser card and a white float at startup while every scripted session
showed the new palette. The REAL cause: the agent's shell runs in a sandbox whose HKCU
registry (and parts of AppData, including the local .NET 10 SDK) are a PRIVATE OVERLAY.
Agent-side registrations never reach the developer's real hive; agent-launched Excels
inherit the overlay and always look healthy. The developer's hive still pointed at
`release_win-x64` — the 07:21 FloatFrame-era build, whose bundle strings were never even
committed — so HIS Excel faithfully loaded that, and when the agent deleted the stale
release publish (believing it unregistered, having read only the overlay), his add-in
broke outright: "cannot be loaded", then an emptied Addins64 key. RESOLUTION: the
developer re-registered HIMSELF with a pure-registry PowerShell block (no .NET — his
machine has only .NET 8/9 for real; the repo's local SDK exists only in the overlay),
pointing at `debug_win-x64`, and confirmed working. STANDING RULES: anything that must
change the developer's registry or run on his side is handed to him as a plain command
(registry writes, or a NativeAOT tool on F:\ — F:\ is shared both ways); the shim log's
`serving from` + `build` lines are ground truth for what any session runs; and ANY
change to the registered DLL path must be flagged to the developer immediately and
clearly, with the ready-to-paste fix (his explicit directive). Also fixed in the same
stretch: the palette hides when the editor frame hides (`df9b29e` — owned windows do not
follow a hidden owner), and search gained its scope dropdown with the whole-group pull
in All mode (`fbffa1d`).

**2026-08-05 SUPERSEDED — THE HOLE RESOLUTION (`c49595e`, worked and was confirmed, then
replaced by the palette above).** The developer wants THE NATIVE
Browser, and the native Browser paints ONLY as a child inside the editor's own tree —
every alternative was measured dead: LinkedWindows.Remove registers in the layout without
moving the live window (a fresh session opens it floating, mid-session it stays docked;
Visible-toggling recreates it docked); a reparent-adopted window NEVER paints, record
visible or not; and the editor's reconciliation closes an adopting frame in both message
spellings. SHIPPED: the cutout hole came back for EXACTLY ONE TENANT. F2 / the toolbar
button execute native 473; the window is restored from maximised and placed centred at
~78%x82% of the document area (PlaceObjectBrowserWindow) so it reads as a floating
window; ObjectBrowserCutout punches its parent-clipped hole; the 200ms hole-watch,
per-event full pass, and settle retry are scoped to the hole being open; its own close box
works through the hole and the canvas goes whole again. The police ignores type 2; the
startup HideReplacedWindows hides a remembered one (no blank apparition at launch — the
VBE saves it open otherwise). The purely-xlide rule stands everywhere else, with this one
licensed exception. The xlide typelib BROWSER VIEW was built, then UNSANCTIONED by the
developer (built without agreement — process error, owned): its UI has no entry points;
objectbrowser.ts sits unimported; the typelib CATALOG (Interop/TypeLibrary.cs +
Editor/TypeLibraryCatalog.cs + obLibraries/obTypes/obMembers protocol) stays as DORMANT
#10 infrastructure — References + LoadTypeLibEx + ITypeLib walking are proven live (VBA,
Excel, stdole, Office loaded). Do not resurface that UI without the developer asking.

**2026-08-05 THE XLIDE OBJECT BROWSER IS REAL (`44deb00`).** The native one is retired
(below); its replacement ships: F2 / the toolbar button open a themed view whose LIBRARY
PICKER holds each open workbook's project AND every referenced type library — VBA, Excel
16.0, stdole, Office all confirmed loading live through `VBProject.References` (works from
the add-in without the trust setting) + `LoadTypeLibEx`. New pieces:
Interop/TypeLibrary.cs (ITypeLib/ITypeInfo as GeneratedComInterface with exact vtable
order + TYPEATTR/FUNCDESC/VARDESC/ELEMDESC/TYPEDESC as sequential structs) and
Editor/TypeLibraryCatalog.cs (types filtered of hidden/restricted/underscored; coclass
members read through the default impl; dual-interface HRESULT+retval collapsed; Property
Get/Let pairs collapsed to one row; VBA-spelled signatures with VT map, VT_PTR deref,
VT_USERDEFINED resolution, optional [param] brackets; enum/module constants with values).
Protocol: obLibraries/obTypes/obMembers request-response pairs (id-matched, cached
host-side per library). Page: objectbrowser.ts — picker, types pane, members pane, detail
strip with the full signature, search within the scope, project members Enter/dblclick
NAVIGATE to their definition. Demo transport serves a canned mini-Excel for the lab.
VERIFIED: demo full flow; live open + 4 libraries loaded + project browsing + detail
strip. NOT yet exercised live: a real library's type/member reads (the picker flip
resisted synthetic driving) — failure modes are bounded (per-item try/catch, HRESULT
checks, empty lists + log counts "Excel -> N type(s)"), and the developer's first
dropdown flip is the test. Backlog from here: <All Libraries> scope, member search across
unloaded types, engine-side completion backfill from the same catalog (#10's other half).

**2026-08-05 PURELY-XLIDE CANVAS (developer directive, STANDING): the cutout-hole
machinery is GONE, and the native OBJECT BROWSER is RETIRED with it.** "I'd like our
canvas to be purely xlide" — nothing native shows through the surface, ever. Every native
tool window now has its home: Immediate, Project Explorer, Properties are panels; Locals
and Watches are ghost palettes; the Object Browser has NONE until its xlide replacement
ships from the typelib model (#10 — now the priority route). The full story is lesson 32,
every step measured: the editor cannot float it (MDI document window; LinkedWindows.Remove
is a silent no-op); ADOPTING its window into a frame of ours (FloatFrame, built then
deleted same day — git history has it) hit an unwinnable pair — OM-hidden, the editor
stops servicing the window and the Browser ships BLANK (the ghost-feed rule cutting the
other way); OM-visible, the editor's reconciliation closes the frame every few seconds,
arriving as both bare WM_CLOSE and WM_SYSCOMMAND/SC_CLOSE, indistinguishable from the
developer's close box. A window whose owner draws its insides can only be replaced, never
stolen. TODAY: the toolbar button answers with a status notice ("being replaced; on the
roadmap"), the native command is never executed, and the police pass hides any type-2
window that appears. Cutout removal stands: NativeToolWindowCutouts, SetCutouts
(surface+overlay), `_watchingCutouts`, the 200ms hole poll, Test-CutoutHoles.ps1 — all
deleted; `PoliceNativeToolWindows` on the settle full pass keeps the canvas pure (hides
2/6/7, hides docked ghost strays 3/4). Call stack button greys outside break
(`setDebugState` mode publish; declined click → status notice). MACHINE NOTE: Smart App
Control began BLOCKING the fresh RELEASE shim publish mid-morning (CodeIntegrity 3033/
3077 events; LoadBehavior silently reset to 0, no shim log at all — that pair is the
diagnostic signature) — the dev loop now publishes `-Configuration Debug` (SAC-clean,
registered at debug_win-x64; ui dist mirrored there too). Recovery: LoadBehavior back to
3 + relaunch on the Debug publish; NEVER change SAC itself (standing).

**2026-08-05 menu curation (developer directives, STANDING):** the end goal is the menu
bar stripped to what only menus can reach. TODAY: the Edit menu lost its editing half —
Undo/Redo, Cut/Copy/Paste/Clear/Select All, Find/Find Next/Replace, Indent/Outdent, and
the five IntelliSense items (List Properties/Methods, List Constants, Quick Info,
Parameter Info, Complete Word) — all duplicative of the toolbar, the find widget, or the
engine, and all acting on the COVERED native pane besides (traps as much as duplicates).
Only Bookmarks survives (nothing on the surface does its job — a future surface feature).
Also gone: the View menu's Toolbars popup itself — its id, 30045, was the old "unknown
popup id" nit, now MEASURED and suppressed alongside its 761 children. QUEUED: the View
menu goes ENTIRELY once Watch + Call Stack port via the ghost-palette route; the UserForm
items (Toolbox 548, Tab Order 469) move to the designer backlog (#14), not back into a
menu. The full menu-bar id map was measured 2026-08-05 (every suppressed id verified
unique across the bar; 746 is the repeating generic id and must never be suppressed) —
the enumeration one-liner lives in the transcript; suppression is `VbeMenus.Replaced`.
Verified live: WM-click into the render widget opened Edit showing exactly [Bookmarks].
COMPLETED same day (`51547d9`): both menus are GONE (top-level ids 30003, 30004). The
menu bar now reads File Insert Format Debug Run Tools Add-Ins Window Help. What replaced
the survivors: BOOKMARKS are the surface's own (ui/editor/src/bookmarks.ts — Monaco
decorations that ride edits, captured per model at onWillDispose and restored by URI on
return; Toggle Ctrl+Alt+K, Next/Prev Ctrl+Alt+N/P, Clear All Ctrl+Alt+Shift+K, all in the
palette and context menu; session-lifetime like the native ones); the WATCH WINDOW is a
panel fed by its own ghost palette (PrepareWatchGhost type 3 + WatchReader four-column
parse Expression/Value/Type/Context + PublishWatches on the Locals cadence — both ghosts
confirmed live: "locals/watch: ghost palette ... feeding the panel"); CALL STACK is a
toolbar debug-cluster button (native dialog 620, a scriptable-dialog port later);
DEFINITION (939, Shift+F2) and LAST POSITION (1822, Ctrl+Shift+F2) are claimed keys plus
editor context actions through the caret-synced execute path. The demo drove bookmarks
end to end and the Watch panel rows; live capture shows the pruned bar, the Watch tab,
and the Call stack button. NOTE: the WatchReader row parse mirrors the Locals pattern but
its exact accessible-name shape is unverified against a real watch (adding one needs the
native Add Watch dialog) — if the panel stays empty during a break with watches set, dump
the ghost's UIA names first. Bundle-size trap met en route: import monaco from
"monaco-editor/editor/editor.api.js" like every other page module — the bare
"monaco-editor" specifier bundles a second megabyte of language contributions.

**2026-08-05 resize arc (`c0bce55`..`91b412b`, developer-accepted):** the minimap/scrollbar
"pushed off canvas" report unwound into three findings, lessons 30-31. (1) THE RATCHET: a
grid cell cannot shrink below its content, and the editor sets its own pixel width — grow
tracked, shrink never did; `#container { min-width: 0; overflow: hidden }` lets the cell
shrink first and the editor's observer follows. Repro requires grow-THEN-shrink in one
session; the hidden browser pane cannot verify it (no frames → no ResizeObserver/rAF —
assert stylesheet rules + class toggles, prove visuals with PrintWindow captures of the
real VBE). (2) THE SETTLE DISCIPLINE: a 60-step drag storm ran 60 full placement passes
(three event routes each treating geometry ticks as news); now every route follows BOUNDS
synchronously per tick and one full pass runs at a 150ms settle — pane events gate on
SUBSTANCE (`_lastFollowSubstance`: pane list + active module + workbook), with two
event-synchronous exceptions (hiding frame per lesson 27; open cutout holes). Same storm
after: ONE pass. (3) THE MINIMAP REST: its canvas repaints a frame behind layout, so it
fades out under `body.live-resize` while events stream and returns at the settle (page
settle timer toggles the class). Also hardened en route: `Follow` asserts browser bounds
every pass (change-guarded, `webview: bounds` verbose line), and Test-ResizeFollow gained
a third column — the Chromium child must match the frame client too. Storm driver pattern
lives in the transcript: SetWindowPos loop + log-line counting between markers.

**2026-08-05 morning delta:** closing a dirty tab now asks Save / Don't Save / Cancel in a
themed modal (`66f269a`, developer-confirmed live). The host gates every close route (tab X,
middle-click, Ctrl+W both sides, tab menu) on the module's OWN text vs its saved baseline —
not the workbook dot — and holds the close with a `confirmClose` message; the page asks one
question at a time, dedupes repeat asks, and answers back through `closeModule`'s `action`.
Save = `SaveWorkbookOf` (trust-free `Workbook.Save`) then close, keep the tab on failure;
Don't Save = `WriteModule(baseline, owner)` revert (+`DiscardEdits` so the debounced old
text cannot chase it) then close; no baseline = close keeping text, honestly logged.
FOLLOW-UP the live test caught (`7d96cb6`): the engine's live copy of a module (didChange)
OUTRANKS its seeded copy in every answer, so the revert left the engine diagnosing the
discarded text — stale problems survived close AND reopen. Any HOST REWRITE that bypasses
the page (revert, Replace All) now sends a corrective full-source didChange and forces the
full pass (`WriteModule(..., hostRewrite: true)`, `NotifyLiveText(+projectId)`); Replace
All also adopts read-backs into `_writtenModules`. PINNED by `Test-CloseConfirm.ps1`
(PASS/FAIL), three legs: seam tripwires incl. the published-bundle stale-deploy trap;
`close-confirm-page-probe.mjs` — a dependency-free headless-Edge DevTools driver walking
the built page's whole flow (ask, Escape, dedupe, queue, all three answers, both routes);
and `engine-live-probe.mjs` — drives the BUILT engine over its own named pipe through the
stale-problems story (live outranks seed; reseed alone cannot heal; corrective didChange
does). Both probe patterns are NEW and reusable: page behaviours and engine contracts can
now be pinned the same way.

**Full context dump, written at the developer's request at session end (2026-08-04). HEAD
was `47da928`; Excel was left running the latest build. Read this block, then the day's
detailed history below it, then the rest of the document.**

**Where the product stands.** The surface IS the visible editor: menus, toolbar (with a
comment suite and Clear-all-breakpoints), tabs (dirty dots, pointer-armed close, drag,
collision labels), explorer (no header label), Problems/Immediate/Locals/Search panels,
settings (tabs indent by default now), find that floats, tooltips that land on their
buttons, breakpoints that preview honestly (dim dot where one can go, orange X on hover
where one cannot, no tooltip on the X) and whose record shifts with edits and always
answers a clearing click. The native editor never retreats — cutout holes for the
unreplaced windows (Object Browser auto-restored from maximised so it always has a close
box), empty-region-silenced chrome bands, synchronous WM_SIZE placement with a bounds-only
fast path. Locals is fully ours, fed by the invisible layered ghost palette. Search/replace
runs engine-side over live text, scoped module/workbook/all. The View menu sheds items as
their windows are ported (2554/2555/2557/222 suppressed today; the STANDING RULE is to keep
shedding as ports land).

**The evening's endgame, not yet in the history below:**
- The Search panel (#38) shipped end to end: engine `workspace/search` (engine/src/search.ts
  + dispatcher case; protocol in both repos' shapes), EngineClient/AnalysisService
  SearchAsync, session handlers (FlushEdits first; scope resolves module/project from
  `_shownProject`), grouped clickable results, monotonic-id acceptance (equality raced a
  synchronous transport — the demo — and dropped its own answer). Replace All searches then
  rewrites per line via CodeModule ReplaceLine, bottom-up per module, re-matched at write
  time; it then RESYNCS THE SHOWN MODULE IMMEDIATELY (waiting for a pane event left the
  editor showing old text while the panel claimed the count — the developer's "replace is
  not working"), arms the resync polls for the rest, reanalyses, and answers with the count
  plus an EMPTY list ("No matches remain") rather than re-listing what it destroyed.
- Unsaved dots (#39) shipped through three corrections, each developer-reported:
  (1) "always showing" — the host was RIGHT (`modules: publish … dirty [False]` in the log);
  the codicon base class ties a single-class `.tab-dirty` display rule on specificity and
  wins on bundle order, so the dot rendered regardless — the selector now carries an extra
  class (`.tab .tab-dirty`); note the close box always used opacity to dodge exactly this.
  (2) Save timing and coverage — `Workbook.Saved` flips a beat AFTER the save command
  returns, and Excel-side saves flip it with no event we hear: PublishModules now diffs
  against what it last sent (`_lastModulesKey`), runs on EVERY poll tick, logs
  `modules: publish […] dirty […]` on change, and Save arms a short resync burst.
  (3) "revert should lift the dot" — the host flag never un-dirties, so the host keeps
  SAVED-TEXT BASELINES (`_savedBaselines`, keyed workbook\0module, snapshotted whenever a
  workbook is KNOWN clean) and, while the flag says dirty, the dot shows only if some
  module's known text differs from its snapshot; unknown keeps the flag's verdict. Probed
  live: save→False, insert line→True, delete it→False with the flag still dirty. SEMANTIC
  NOTE: the dot now means "code differs from saved" — cell-only changes no longer dot the
  tabs, deliberately.
- Deploy-order gotcha met today: package the ENGINE only after killing Excel — the running
  sidecar holds `engine/dist/xlide-engine.exe` and the copy fails EBUSY, leaving a stale
  engine deployed (a stale engine answers workspace/search with MethodNotFound and the
  panel sits at "Searching...").

**Known nits and edges, all small, none blocking:**
- Locals panel: the context strip stays hidden — the context box is a PANE to UIA (never an
  Edit; measured 2026-08-05) and its name reads EMPTY in-proc even though an outside client
  sees "VBAProject.Module.Proc"; the reader takes a dotted pane name when one ever arrives
  and normalises empty to null. External-command steps outrun the poll cadence; real steps
  go through our command path which arms the fast watch.
- Residual canvas flicker during resize drags is the browser compositor catching up — the
  dark ground keeps it subtle; accepted for now.
- `_writtenModules` is keyed by bare module name — two same-named modules in different
  workbooks can cross-talk in the dirty comparison's current-text lookup and in replace's
  baseline; the project-qualified rework is a known small refactor.
- Search matching is literal with match-case/whole-word only (no regex); replace re-matches
  host-side with the same rules. The results list caps at 500 with a truncated flag.
- The Toolbars POPUP menu item id is unknown (its child toggles 761 are suppressed) — find
  it via VbeCommands.Describe someday and suppress the popup itself.
- Baselines only exist for modules whose text we have seen (shown once, or written); other
  modules of a dirty workbook keep the flag's verdict until the next save snapshots them.
- The hidden-pane caption-match question stands (below, tab-X bullet): why the tracker
  cannot match hidden panes — cosmetic now that resync covers it.

**Standing probes (tools\harness), each self-describing:** Test-CutoutHoles,
Test-ResizeFollow, Test-CloseVbe, Test-CloseHiddenPane (PASS/FAIL), Test-GhostLocalsPanel
(PASS/FAIL), Test-CloseConfirm (PASS/FAIL; drives the built page headless via
close-confirm-page-probe.mjs), Probe-FloatLocals, Probe-GhostLocals, Probe-LocalsLifecycle. The page demo
serves via `.claude/launch.json` (npx http-server :8123) for browser-side verification —
the demo transport answers search, settings, locals, and modules with canned data.

**Suggested next, in order:** the Watch window by the ghost-palette route (same API chain
as Locals — float via LinkedWindows.Remove, layer to alpha 0, park off-screen, read by
UIA; the dormant-panel pattern is proven); a live click-through of Search on a real
workbook (the engine path is typed and packaged but the only full verify was the demo);
the Locals context-strip nit; then the Object Browser replacement (ties #10 typelibs),
which retires the biggest remaining hole.

**The day's detailed history (still accurate):**
- **Afternoon batch** (commits `eb31423`..`HEAD`, each message carries its story): tabs
  indent by default; the engine finally calls `resolveTypeCompletions` (types after `As`);
  toolbar gained Clear-all-breakpoints and a comment/uncomment/toggle suite; breakpoints
  tell the truth (hover previews dim-dot-or-orange-X via a page mirror of CanBreakOn, click
  draws nothing on refusal, clears never validity-gated, and the record SHIFTS with edits —
  the ghost dot was line-drift the log named); a maximised Object Browser is restored to a
  window by the cutout pass (it had no close box and no way back); WM_SIZE places bounds
  only (PlaceSurfaceFast — the lag was OM work per drag tick; developer confirms mostly
  gone, residual canvas flicker = compositor catch-up, known); the find widget floats
  instead of pushing line 1 down. QUEUED with specs in tasks #38/#39: the scoped search
  panel (module/workbook/all — engine-side search over its seeded live text) and per-
  workbook unsaved tab dots (verify the all-modules-save-together model via Workbook.Saved
  first; the developer believes it but asked it be checked).
- **The tab X's REAL mechanism, found third: the tracker cannot see a hidden pane close.**
  The pointer-path and render-skip fixes below were sound but fixed a different (real,
  demo-proven) race; the live bug was host-side, and the per-pass verbose instrumentation
  named it: the pane tracker's list only ever holds the pane windows it can match — the
  ACTIVE one, in practice — so closing a HIDDEN pane leaves its picture identical
  ("pass saw 1 [CleanModule] unchanged" straight through the close), Changed never fires,
  PublishModules never runs, and the strip keeps a dead tab whose next click REOPENS the
  module. Closing the active tab always worked, hence "only when not focused". Fixes: any
  window destroy (a dying window's class reads empty, so a pane cannot be told from a
  tooltip) arms a moment of polls that re-read the object model's open list and republish
  (the page's render-key skip makes identical republishes free); the tracker queues events
  that land mid-refresh instead of dropping them, with a bounded trailing loop that
  declares itself Stale when a burst outruns it; and refresh failures 2..199 log at
  verbose instead of silently. `Test-CloseHiddenPane.ps1` is the standing PASS/FAIL probe.
  OPEN QUESTION for later: WHY hidden panes fail the tracker's caption match (maximised
  MDI captions?) — the WM_CLOSE fallback in CloseModule iterates tracker panes and so
  cannot reach hidden panes either; the OM path covers them today.
- **The tab X died a second death and got a second, structural fix.** "Clicking a tab's X
  doesn't close it if it's not focused": a click needs pointerdown and pointerup on the
  same LIVE element, and pressing an unfocused surface stirs the host into a setModules
  echo whose renderTabs rebuild destroyed the pressed X mid-press — the click never fired.
  Now the X is armed at pointerdown as DATA (module + workbook) and fired at pointerup
  against whatever twin sits under the release point, so no rebuild can eat it; sliding
  off still cancels; and renderTabs skips entirely when its render key (order, identity,
  active pair, badge counts) is unchanged, so echo rebuilds stop happening at all. All
  three behaviours verified with synthetic pointer events against the demo, including a
  deterministic mid-press clone-rebuild. UI pinned: Test-ResizeFollow.ps1 and
  Test-CloseVbe.ps1 joined Test-CutoutHoles.ps1 in tools\harness; .claude/launch.json
  serves the page demo (npx http-server, port 8123) for browser-side verification.
- **Two live-test regressions of the morning's work, both fixed** (lesson 27). Closing the
  VBE window crashed Excel: the frame-HIDE event drove placement, whose cutout pass called
  the object model INSIDE the editor's close handling; placement now gates on
  `IsWindowVisible(_frame)` — hidden frame → surface hides with it, zero OM work — and
  `Probe-CloseVbe`-style close/reopen cycles leave Excel standing. Resizing was "slippery":
  `File.AppendAllText` per verbose line opened the file (and woke the antivirus) thousands
  of times a second during event storms; the log now holds its file open and flushes per
  line, move events log only for editor-relevant window classes, and `SetChrome` sends only
  on change.
- **Resize follows everywhere now.** The developer's "resizing window no longer adjusts
  canvas": placement only ever followed PANE events, and the empty workspace has none, so
  nothing re-placed the surface. Two routes fixed it: the tracker's new `FrameChanged`
  (frame/MDIClient window events → RefreshSurfacePlacement, works in every state) and
  `FrameSubclass` (comctl32 subclass on the frame; WM_SIZE re-places the surface
  SYNCHRONOUSLY, before the native layout paints — that ordering is the fix for the
  developer's follow-up "on resize, the old native UI flickers and bleeds through"). Three
  more anti-bleed layers: the overlay paints the theme ground wherever the browser child
  has not caught up (PaintGround; WM_PAINT when not loading); the browser's own idle
  colour is dark via ICoreWebView2Controller2.DefaultBackgroundColor; and the native
  menu/toolbar band windows are SILENCED BY EMPTY WINDOW REGION (SetNativeChromeBands) —
  they paint without clipping siblings, so covering never stopped them, and hiding them
  lost to the editor re-showing them per resize; a zero-region window cannot paint and
  nothing in Office resets it (lesson 26). Regions are cleared when not covering and on
  session stop, so a bare editor keeps its menus. Probe-verified pixel-exact in both pane
  states; the flicker feel awaits the developer's live test. The explorer's PROJECT header
  label is gone (developer request).
- **Verbose dev logging is a standing directive** (developer, 2026-08-04: "as verbose as
  possible, for dev, so that you can discern as much as possible during my live tests").
  `Log.Verbose` is ON by default (`XLIDE_VERBOSE=0` quiets; #15 flips the ship default),
  and the logger collapses consecutive duplicate lines so event storms read as one line
  with a count. Instrumented seams: window events (kind/class/hwnd), placement decisions
  and skips, cutout candidates per pass, page traffic BOTH directions (type + length),
  debug-state transitions, follow decisions, watchdog ticks, frame-subclass hits. When
  adding features, instrument as you go — dedupe-friendly single lines per state change.
- **A cancelled Excel shutdown no longer strands the add-in.** The 08-02 log showed
  OnBeginShutdown AND OnDisconnection(HostShutdown) both land BEFORE the save prompt;
  cancel then left a corpse (the watchdog used to be retired in OnDisconnection). The
  watchdog and retained pointers now survive a HostShutdown-mode disconnection; revival
  requires the frame standing + VISIBLE + enabled for two 1.5s ticks (visible: a real
  teardown hides windows before destroying them, and a cancellation with the editor window
  closed just waits until the developer reopens it). The no-frame patience stand-down is
  gone — the watchdog waits as long as it takes; a dying process takes it along.

**2026-08-02 late delta:** the developer's "opening non xlide windows reverts
the toolbar" is FIXED and probe-guarded — the surface never retreats; it punches
`SetWindowRgn` holes where native tool windows sit (`tools\harness\Test-CutoutHoles.ps1`,
lesson 24). All docked native toolbars are hidden at start-up. Then the Locals-panel
replacement was built end to end and deliberately REVERTED after five probes proved the
editor only feeds an on-screen Locals window (machinery dormant in-tree; lesson 25; open
thread 4 has the detail). Cutout state: tools `[2,3,4,6,7]` get holes; Immediate stays
hidden-and-mirrored. The developer has since asked for maximum creativity on getting
Locals (and the debug windows generally) into OUR UI regardless — and on 2026-08-04 the
GHOST PALETTE shipped the same day: **Locals is now OURS** — the themed panel is live,
fed by the native window floated, made WS_EX_LAYERED at ALPHA ZERO, click-through, and
parked at -20000,-20000. The feed gate was never about being SEEN: a layered window
renders into its own surface regardless of position or occlusion, so the invisible ghost
is fed through every break and step (`Probe-GhostLocals.ps1` tracked counter 1→4 across
steps at alpha 0 off-screen; lesson 29). `PrepareLocalsGhost` (AddInSession) floats it via
`LinkedWindows.Remove`, ghosts it, and attaches the once-dormant LocalsReader;
PublishLocals pushes on the debug watch; the page's Locals tab is visible and 2555 routes
to it; `RestoreLocalsPalette` on Stop gives a bare editor back a normal window.
`Test-GhostLocalsPanel.ps1` is the standing PASS/FAIL probe (PASSED: 3 rows at entry,
clear on run-end). Known nits: the context strip reads null through the COM reader (the
Edit element's name — panel-header polish), and external-command steps outrun the poll
cadence (real steps go through our command path, which arms the fast watch). NEXT: the
Watch window by the identical route, then Call Stack. The earlier hole-based findings
below stand as history; the CAPTIVE PALETTE route (visible, caption-stripped, in-panel)
remains proven (`tools\harness\Probe-FloatLocals.ps1`):
`window.LinkedWindowFrame.LinkedWindows.Remove(window)` UNDOCKS the Locals window through
pure object model (its linked frame is type 11); geometry is then settable on the WINDOW
(`.Left/.Top/.Width/.Height`, points — a 260x160-point request produced a 173x107-pixel
`VBFloatingPalette` top-level); and the tiny floating window FEEDS PERFECTLY, tracking
break entry AND EVERY STEP live (counter 1 → 1-pending → 2 across two steps) — the
step-freshness that was impossible covered is fully alive floating, because floating is
genuinely on-screen. BUILD PLAN (next stretch): float it at session start, strip
WS_CAPTION/WS_THICKFRAME/WS_SYSMENU via SetWindowLong + SWP_FRAMECHANGED, position it
exactly over the panel's Locals tab body (page ships the body rect in physical pixels —
multiply by devicePixelRatio — host adds the overlay's screen origin), re-place on
FrameChanged and page layout changes (owned top-levels do not follow their owner), show
only while the Locals tab is selected (Visible toggle; show-during-break refreshes,
proven), un-hide the page's Locals tab, route 2555 to it, and on session Stop RESTORE the
caption styles and leave it hidden-floating so a bare editor keeps a usable window. Being
top-level and owned by the frame it naturally floats ABOVE the surface — no holes, no
z-fights; native expansion and in-place value editing come free inside our panel chrome.
The dormant themed-table machinery stays for a later fully-drawn skin. Watch window: same
route, same API, afterwards. Everything below still holds.

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
   e. DONE both sides. Requests: explorer rows carry their workbook, and
      `activateModule`/`outline`/`navigate`/`closeModule` send an optional `project` (resolved
      host-side via `ProjectIdFromDisplay`); the accordion is a (module, workbook) pair. Push:
      `setModules` carries the workbook beside every module and the active one; tabs are
      (module, workbook) pairs through ordering, activation, cycling, closing, dragging, and
      the menu, labelled "Module1 — Book2" only on collision; close matches the pane's own
      project. This was also the developer-reported bug — two same-named modules could not be
      open in two tabs — fixed the moment it was reported. The last residue is DONE too:
      findings reach the page with their workbook, tab badges count their own module only,
      tree counts file under (workbook, module) via `problemCountKey`, panel rows name their
      workbook on collision, and clicking a row navigates to the right workbook.
   #24 IS COMPLETE. `Test-CollidingModules.ps1` (session scratchpad) is the acceptance probe;
   re-create it from the handoff's recipe if the scratchpad is gone: harness -KeepOpen, attach,
   Workbooks.Add, add a colliding module with its own defect, read the log.
   Beware NUL bytes: string-separator literals must be the `\0` ESCAPE, never the raw byte —
   four files carried raw bytes and git/ripgrep treated them as binary (fixed in eac2be2).
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
4. THE PHASE-OUT PROGRAM (developer directive, 2026-08-02 evening): remove or supersede every
   remaining native VBE UI element. Already done: native toolbars hidden and their menu
   toggles suppressed (shared id 761), the native window list suppressed (shared id 830) —
   the tab strip is the window list — and Tools > Options (id 522) routed to the product's
   settings dialog; Project Explorer, Properties, and Immediate were already replaced and
   routed. ALSO DONE (2026-08-02 late, probe-verified): the surface never retreats for a
   native window any more — "opening non xlide windows reverts the toolbar" was the retreat
   itself; now it keeps the whole client and punches region holes where native tool windows
   sit (Test-CutoutHoles.ps1 is the standing probe; lesson 24 has the maximised-MDI phantom
   strip). The Locals REPLACEMENT was then built end to end — LocalsReader (UIA rows
   `Expression <n> Value <v> Type <t>` + context Edit), SetLocals message, a page panel,
   route 2555 — and REVERTED the same evening after five probes: the editor only feeds an
   ON-SCREEN Locals window. Hidden it never fills; visible-but-covered it fills at some break
   entries, late or never, and never on a step; any genuinely visible part (half the window
   through a hole) tracks perfectly; a 3px sliver does not. Lesson 25 has the whole
   investigation including the step-semantics trap that faked half the evidence. All the
   machinery STAYS DORMANT in the tree (LocalsReader.cs, SetLocals/SurfaceLocalRow, the
   page's hidden Locals tab + PublishLocals in AddInSession) for when the data has a
   reliable source — the #14 debugger era, or restyling the native window positioned inside
   the surface panel (SetWindowLong caption strip on a floating palette is the sketched
   route). Remaining in the program, in rough order: Watch window and Call Stack (same
   feed problem, same conclusion — native through holes for now); Object Browser (ties to
   #10's typelib model, replaceable WITHOUT the debugger); the Window menu's arrangement
   items; native dialogs the object model can faithfully rebuild (References and Macros are
   scriptable; Project Properties); the UserForm designer and its Toolbox (#14, the
   largest); Help items last. The rule stands: until a replacement exists, the native
   window stays reachable — through a hole, not by retreat.
5. Then the standing backlog: #20 right-click curation (needs the developer), #22 split
   groups, #13 tests panel, #14 debugging/forms designer, #10 typelib backfill, #9 the C#
   analyzer port. (#12 settings is DONE: six choices, gear dialog, settings.json, formatter
   wired.)

**Before touching anything, know these:**
0. Smart App Control turned ON on this machine mid-day 2026-08-02. It blocks freshly built
   unsigned RELEASE managed test assemblies from running (0x800711C7, surfacing as xUnit's
   "Test process did not return valid JSON"); Debug test assemblies and the NativeAOT shim
   itself load fine. dev.ps1's test gate therefore runs Debug. If a Release-built THING ever
   fails to start with that error, suspect the policy before the code. Turning SAC off is the
   developer's decision alone — it cannot be re-enabled without reinstalling Windows.
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
  panel, and the project explorer item answers with a notice. A native tool window opening does
  NOT make the surface retreat (retreating handed the native menu and toolbar rows back — "the
  toolbar reverts", 2026-08-02): the surface keeps the frame's whole client and punches
  window-region holes exactly where each visible native tool window sits
  (NativeToolWindowCutouts -> OverlayWindow.SetCutouts, SetWindowRgn RGN_DIFF). The window is
  live inside its hole — painting, input, dragging all work — and everything around it stays
  ours. Hole rects come from the object model's window captions (FindChildByCaption, class-free
  because the Object Browser has its own class), clipped to each window's parent (a maximised
  MDI child reports a phantom caption strip above the document area; lesson 24). While any hole
  is open a 200ms poll re-derives placement, because the Object Browser moves without a word to
  the pane tracker; the overlay skips identical states so the quiet case costs nothing. All
  docked native toolbars are hidden at session start (HideNativeToolbars; their commands are
  all in the menus, their toggles suppressed). Verified by tools\harness\Test-CutoutHoles.ps1:
  our chrome intact in every capture, Locals and Object Browser each live inside their holes,
  "surface: N native hole(s) cut [...]" then "whole again" in the log when they close.
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
   them would remove features). While any is open the surface no longer retreats: it keeps the
   whole client and shows the window through a punched region hole, chrome intact
   (Test-CutoutHoles.ps1 guards this; hole rects logged as "surface: N native hole(s) cut").
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
  replacement exists, the native window stays reachable — shown through a punched hole in the
  surface, never by retreating the surface (the toolbar-revert bug, fixed 2026-08-02); the menu
  routing table in RouteMenuCommand is where "open ours instead" gets decided per window.
