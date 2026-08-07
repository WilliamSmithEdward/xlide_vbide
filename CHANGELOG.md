# Changelog

## Unreleased

The analyzer has been answering more than anything asked it. Three of those answers now reach the
editor.

### Added

- Quick fixes. A finding that can be fixed offers to fix it, from the lightbulb or Ctrl+period:
  the missing `Set`, the parentheses a `Call` wants, `Option Explicit` at the top of a module, a
  stub for a procedure that does not exist yet, a placeholder for an argument that was left out.
  Every finding also offers to silence itself on the line above.
- Semantic colouring. A name that denotes a type is coloured as the kind of type it is — class,
  enum, or user-defined — and a host global is tinted apart from an ordinary identifier, unless
  something in the module has declared that name itself, in which case it is the developer's and
  is coloured as one.
- Go to Definition and Find All References, on Ctrl+click, F12, Shift+F12, and the editor's
  right-click menu. Both answer across every module of the workbook and stop at its edge: two open
  workbooks can each hold a `Module1` and a `Recalculate`, and they are unrelated. A definition in
  a module with no tab open is still reached — the module opens on the way.

### Changed

- Go to Definition and Last Position are xlide's rather than the host's. Shift+F2 and
  Ctrl+Shift+F2 still do what a VBA developer's hands expect; what answers is the analyzer, which
  crosses modules, resolves members reached through a receiver, and reads the text as typed rather
  than as last written back. Only what the host alone can do is still sent to it: running a
  procedure, and the breakpoints the debugger owns.

### Known limits

- The references window lists only modules that already have a tab open. The answer behind it
  covers the whole workbook; the window can only draw a module it has the text for.
- Renaming a symbol across modules is not wired yet. It is the one item here that changes code
  rather than describing it, and what it does to a module with no tab open has to be decided
  before it can be trusted.

## v0.3.0 (2026-08-06)

The editor window becomes xlide's: its title bar, its corner, and a command strip that survives a
narrow pane. Uninstalling becomes something that finishes, or says why it did not.

### Added

- The window is titled `XLIDE - Book1.xlsm` and carries the product's icon. Only the product name
  is replaced; the workbook and module the editor names are still named. The editor rewrites its
  own caption as the active project changes, so it is retaken on every rename, and both the title
  and the icon are put back when the add-in unloads.
- A wordmark and version sit at the top right of the menu bar, quiet enough to disappear while
  working.
- An About dialog, from the question mark beside the gear: the version, the build number, when the
  surface was built, where the add-in is loaded from, and the keys worth knowing. Every question
  that begins "it used to work" is really a question about which build is loaded.
- `docs/testing.md`: which of the four kinds of check to reach for, and the traps this codebase has
  already charged for.

### Changed

- The command strip scrolls instead of clipping. When the pane is too narrow for every command, an
  edge appears at each end that has more past it: full height, square, tinted, with the commands
  fading beneath. Held down it keeps scrolling, and a vertical wheel moves it sideways.
- Split Right and Split Down leave the tab right-click. Splitting is a placement, and the two
  direct ways of placing a tab, dragging it to the edge you mean and Ctrl+\\, both show you the
  result before you commit to it.
- The cursor is a closed hand for the whole page while a drag is live, including over the editor,
  which was still showing a text caret under a dimmed page.
- The unsaved dot belongs to the tab that earned it. It was a workbook fact worn by every tab the
  workbook owned, which marked four untouched tabs for one edit.

### Fixed

- An uninstall started from the installed-apps list failed on its own executable, having already
  removed the entry from that list: files on disk, the editor still loading them, and no way left
  to remove it through Windows. The process handing over now exits at once instead of waiting for a
  keypress while the copy tries to delete it, deletion catches the access denial Windows actually
  raises, and the entry is removed last and only when the files are really gone.
- Deleting the installation by hand stranded that entry pointing at an executable that was no
  longer there. The uninstaller now lives outside the folder it removes, so Uninstall still works.
- Excel being open no longer ends an install or an uninstall. Both offer to wait while you close
  it, or to force close it.
- An install replaces the previous one rather than writing over the top of it, so nothing an older
  version shipped survives into a folder that is otherwise this one.
- xlide appeared in the installed-apps list with a generic icon.
- A tab opened on a workbook that has never been saved showed the unsaved dot immediately, and
  could never lose it: baselines were only recorded while the workbook was clean, and that moment
  never comes for a workbook with no saved text.
- Closing a tab and declining to save left the Problems panel reporting errors in code that no
  longer existed.
- Dragging a pane to the editor's edge did nothing until a module had been opened. With none open
  the editor area is not drawn and measures nothing, so the drop zones were skipped entirely.
- Assemblies and the surface now read one version, from one file.

### Known limits

- The tool panes' arrangement survives a session; the editor's own splits do not. Reopening brings
  the open modules back as tabs in one group.
- Still not code-signed, so Windows warns before running the installer.
- The engine is 90 MB of Node runtime wrapping 2.2 MB of analyzer. Compression hides most of that
  in the download, not on disk. The C# port that removes it is in progress.

## v0.2.1 (2026-08-06)

A packaging release. The installer is 28 MB, down from 102 MB, and it now arrives attached to
the release instead of staying on the machine that built it.

### Added

- `tools\release.ps1` builds the installer and attaches `xlide-setup.exe` to a tag, so a
  release hands you the thing that installs rather than the source it was built from. It runs
  the gate first and refuses to release a build whose engine or page is missing.
- The installer carries the product's icon, drawn at every size Windows asks for rather than
  downscaled by the shell from one bitmap.

### Changed

- The payload was embedded raw, and 90 MB of it is the language runtime inside
  `xlide-engine.exe`. It ships compressed now and unpacks on install, which takes the download
  from 102 MB to 28 MB. Every file was checked byte for byte on the way back out. Compression
  is cached by content, so only a payload that actually changed pays for it.
- The page's smart-editing helpers live in this repository instead of being read from a
  neighbouring checkout, so the page builds on any machine. CI had never once built it, and
  had been publishing whatever bundle happened to be committed (decision 14).

### Fixed

- The installer packaged the entire `ui` folder, so `node_modules` travelled inside it:
  monaco's sources, esbuild's binaries, typescript. 146 MB of build tooling that nothing on a
  user's machine ever runs. Only the built page ships.
- The installer would build without a language engine and mention it in grey, producing
  something that installs an editor with no diagnostics, completion, or hover. It refuses now.
  An editor-only build has to be asked for, and is named so it cannot be handed over as the
  product.
- Assemblies reported version 0.1.0 whatever release they shipped in.

## v0.2.0 (2026-08-06)

The workspace becomes arrangeable: every module open at once, editors side by side, and tool
panes that dock wherever the developer puts them.

### Added

- Editor groups. Split right or down (Ctrl+\\, the tab menu, or dragging a tab to a group's
  edge) and work in two modules at once. Tabs move between groups by dragging, and a group
  dissolves when its last tab leaves. Which modules are open stays the editor's truth; where
  their tabs sit is the developer's, and survives every host update.
- Every open module is live. The surface holds one editor model per open module instead of one
  for the module on screen, so switching tabs is instant, undo history and scroll position
  survive a switch, and a background module's squiggles update while another is being edited.
  Two workbooks' Module1 are two documents (decision 12).
- Dockable panes. Explorer, Properties, Problems, Immediate, Locals, and Watch dock in four
  sections around the editor (left, right, top, bottom), each a split of tabbed groups. Drag a
  pane by its title: a five-zone compass appears over the region under the pointer, the centre
  tabs it onto that group, an edge splits beside it, and an editor edge docks it against the
  editor. Dragging along a strip reorders the tabs in it. The preview outlines a section it
  would join, or dashes the space it would create, and the drag ends if the window loses focus.
  Layout persists across sessions.
- A Panes menu, beside Settings on the toolbar: one checkable row per pane. Every pane also
  carries an X on its group. The explorer cannot be closed, because with every tab shut it is
  the only route back to a module, and its row says so.

### Fixed

- Backspace and Tab acted on the wrong editor once a second group existed: every editor shares
  one keybinding service, so two identically-scoped rules matched and the later one won
  everywhere. Each editor's rules are now scoped to that editor (lesson 34).
- A page reload came back with its models but no tabs, and later without its properties: the
  "what changed" caches still described the conversation before the reload, so the republish
  sent nothing. A second ready clears them first (lesson 35).

### Changed

- The Window menu is gone; the bar reads File, Insert, Run, Tools, Add-Ins, Help. Its window
  list was already the tab strip's job, and Split, Tile, and Cascade managed native windows the
  surface covers. The editor groups are that job done where the developer looks.
- Every divider that resizes something wears the same centred grip.

### Development

- The debug api's `eval` now AWAITS a promise instead of answering `{}`, so an async probe
  works; `await` polls a predicate in the page and answers when it comes true, replacing
  caller-side poll loops; `layout` returns the whole visible arrangement in one request, and
  `layout?reset=1` puts a rearranged one back; `reload` reloads the page, waits for it, and
  says which bundle came back and whether it is behind the one on disk; `inspect` answers what
  a selector matches with the CSS rules that claim each property; `bench` times tab switching,
  layout, and typing in the page; and `console` keeps what the page said to itself, which the
  log deliberately does not carry. See `docs/debug-api.md`.
- `capture?selector=` crops a screenshot to one element, so a widget can be looked at rather
  than inferred from numbers; `tools\harness\Get-Shot.ps1` fetches one and writes a PNG.
- `tools\page.ps1` is the page loop in one command: typecheck, build, deploy into the running
  shim, reload the live page, and prove the running build is the one just made. About a second,
  with no Excel restart. `-Watch` repeats it on every save.
- `tools\verify.ps1` is the whole local gate in one command: page typecheck, build, bundle
  checks, headless probes, Release build, unit tests, and the Release-carries-no-debug-api
  check. `-Live` adds the standing probes against an open editor.
- CI now builds and typechecks the page. It never did: `dist/` is not committed, so every run
  published a shim carrying NO editor bundle, and the artifact would have shown the native
  pane. A TypeScript error merged just as green. The workflow now installs the page's
  dependencies, typechecks, builds, runs the bundle checks, and asserts that the published shim
  actually carries the page and that Release carries no debug api.
- `Test-Churn.ps1` is a leak probe for the churn this release introduced: splitting and
  dissolving, docking and undocking, a dozen times each, asserting that editors, models, dock
  groups and DOM come back to where they started. Counts rather than megabytes, because an
  exact number that must return to its starting value is a far better leak detector than a heap
  figure nobody can interpret.
- The split tree's arithmetic moved into `docktree.ts` (pruning, collapsing, same-axis
  absorption, divider arithmetic), and the docking layout now runs on it, so the 12 unit tests
  (under a second) cover the code that actually ships rather than a copy of it. The live drags
  test the gesture; this tests the algebra, where a mis-collapse used to show up as a pane that
  vanished three drags later. The editor grid keeps its own tree for now: its splits carry DOM
  references, so sharing this one is a separate change.
- `docs/ui-lessons.md`: what building this surface inside a host that moves under it taught us.
  Pointer gestures that survive rebuilds, one keybinding service across many editors, drag
  targets a person can aim at, and probing a live page honestly.

### Known limits

- Still not signed, so no binaries are attached; build from source with `tools\dev.ps1`.
  Signing and update plumbing remain the next release-engineering milestone (decision 8).
- Panes dock in the four sections around the editor; floating and OS-popout windows are not
  built (decision 13). A tool pane cannot join the editor's own tabs, by design.
- The editor grid and the pane docks keep separate split trees. Only the docks' tree is the
  unit-tested pure one; the grid's carries DOM references and shares nothing yet.
- The F1 command palette still lists Monaco's own Find entries, which open the retired native
  find widget; every key, button, and menu route leads to the surface's own search.
- The debugger integration and the UserForm designer are still the two large milestones.

## v0.1.5 (2026-08-06)

A fix for tabs that showed above an empty canvas, page errors that no longer disappear, and a
local diagnostic door for development that no shipped build contains.

### Fixed

- Opening the editor could show module tabs above the "No module is open" view, and closing
  every tab could leave the code on screen with no tabs. The tab strip and the canvas were
  answering to different authorities; the object model now decides both.

### Added

- An uncaught error or unhandled rejection in the editor page is written to the shim log, with
  its message, source location, and stack. Previously these were invisible without developer
  tools attached, which left a misbehaving surface with nothing to report. Bounded at twenty
  per session so a fault in a loop cannot flood the log.

### Development

- A local HTTP api into a running session, for diagnosis and automated testing: session state,
  native windows, analyzer findings, locals and watches, module read and write, commands by
  name, breakpoints, caret placement, page script, window capture, an awaitable log,
  performance samples, a one-request evidence capture, and a session replay script. It detects
  a modal dialog holding the editor, clears one it raised itself using only safe buttons, and
  never touches a dialog opened by hand. See `docs/debug-api.md`.
- Debug builds only. A Release publish contains none of it, verified by inspecting the built
  binary.
- `tools\dev.ps1` waits for the shim to be unlocked before publishing and refuses to report a
  publish older than its own sources, so a stale build cannot be tested by accident.
- Three investigation writeups: `docs/locals-break-investigation.md`,
  `docs/watch-window-investigation.md`, and `docs/working-with-modals.md`.

## v0.1.4 (2026-08-05)

The Watch panel works against real watches, and it carries its own controls. Two more menus
leave the bar, which is down to seven.

### Fixed

- The Watch panel stayed empty while watches were defined. Its row parse expected a header word
  that real watch rows do not carry; measured against a watch made through the editor's dialog,
  the parse now reads expression, value, type, and context correctly and tracks values as you
  step. `tools\harness\Test-WatchPanel.ps1` pins it.
- Building from source: `tools\dev.ps1` failed at registration on machines with Smart App
  Control enabled, leaving a published shim unregistered. The registration tool now builds
  Debug, like the test gate. The screenshot harness also follows Excel's launch handoff, which
  had been leaving it waiting on a process Excel replaced.

### Added

- The Watch panel has Add, Edit, and Quick buttons above its table, so watches are managed where
  they are read. They open the editor's own watch dialogs by design: those dialogs are modal,
  and driving one invisibly can hang the editor with no window left to dismiss (decisions.md,
  11).
- The toolbar gained Compile, Run to Cursor, Set Next Statement, and Show Next Statement. The
  last three grey outside break mode, like the Call Stack button.

### Changed

- The Format and Debug menus are gone; the bar reads File, Insert, Run, Tools, Add-Ins, Window,
  Help. Every Debug command has a home on the toolbar or in the Watch panel. Format held only
  UserForm layout commands, which return with the designer that gives them meaning.

## v0.1.3 (2026-08-05)

One search UI. The bottom panel's Search tab and Monaco's find widget duplicated each other;
both give way to a single floating search widget where the find widget sat.

### Changed

- Search is one floating widget with a scope dropdown. Module scope finds live as you type,
  with match highlights, an n-of-m counter, and Enter/F3 cycling; its Replace All applies every
  match as one edit, so a single undo reverts the whole operation. Workbook and All-workbooks
  scopes run the engine search and render the grouped, clickable results inside the widget.
- Keys: Ctrl+F opens find (seeded from the selection), Ctrl+H opens with replace unfolded and
  focused, Ctrl+Shift+F opens in Workbook scope, F3 and Shift+F3 cycle and reopen the last
  search, Escape closes from anywhere. The toolbar's Find and Replace buttons open the widget.

### Added

- Find All: in module scope it opens a table under the widget listing every match as a
  line-and-preview row; a click selects and reveals the match in the editor and the counter
  follows. The table tracks edits and query changes live, and caps at 500 rows with the count
  naming the cap. In wider scopes the button runs the engine search.
- A left-edge chevron folds the widget to find-only; the replace row unfolds on the chevron or
  Ctrl+H and the state sticks while you work.

### Known limits

- The F1 command palette still lists Monaco's own Find entries, which open the retired native
  find widget; every key, button, and menu route leads to the new widget.

## v0.1.2 (2026-08-05)

Debug panel polish from the first live day of the restored Locals tracking.

### Changed

- The Locals and Watch tables no longer spread their columns across the panel: columns hug the
  longest content up to a cap and cluster left, header and rows staying aligned, while hover
  bands and the sticky header still run the full width. The Watch table also gains the scroll
  and monospace styling only Locals had.

### Added

- The Immediate panel has a clear button beside its prompt: it empties the panel's history and
  returns focus to the input. Clearing is page-local by design, because the hidden native window
  keeps its text and the mirror appends only what is new, so cleared output does not return.

## v0.1.1 (2026-08-05)

Locals and Watch track the debugger again. The panels had gone dark during breaks, and taken a
day of crashes with them, because the buffer behind every accessibility property read was sized
for 32-bit Windows in a 64-bit process, corrupting eight bytes of stack per read. Release builds
happened to survive the overwrite; the Debug builds the dev loop switched to did not.
[docs/locals-break-investigation.md](docs/locals-break-investigation.md) has the full story.

### Fixed

- The Locals panel fills at break entry and tracks every step, and the Watch panel rides the
  same pipeline. Root cause: `UiVariant`/`ComVariantBlock` were `Size = 16`; an x64 VARIANT is
  24 bytes.
- In a break with nothing readable, the Locals panel says "No variables to show." instead of
  claiming "Not stopped": the message now knows the debugger's state.
- An unreadable element or a failed read no longer silences the panel for the session: faults
  are isolated per element, logged once per streak with the failing stage named, and retried
  after a five-second backoff.

### Changed

- The ghost palette readers moved off the editor's own thread onto a dedicated reading thread
  (`GhostReaderThread`): the host never blocks on a read, readings are cleared at break exit so
  a new break starts empty, and the client now has the shape the accessibility framework
  actually supports.

### Known limits

- The context strip (the broken procedure's name) stays hidden: the context box reads empty
  through the in-process channel even though the data is visibly there from outside.
- The Watch row parse is unverified against a real watch; setting one through the native Add
  Watch dialog is the outstanding test.

## v0.1.0 (2026-08-05)

First tagged milestone. xlide turns the Visual Basic Editor into a modern development surface: a
native COM add-in (ahead-of-time compiled, no runtime required on the user's machine) hosts a
browser-based editor over the VBE's document area, with an analysis engine running as a sidecar
process.

### The editor surface

- Monaco-based editing over every code pane: VBA syntax highlighting, folding, multi-cursor,
  find, format-on-demand, and the module tab strip with drag reorder, close affordances, and
  unsaved indicators per workbook.
- Language intelligence from the engine: live diagnostics with severity filters, completions,
  hover, and signature help, plus typing ergonomics (auto-casing, block closers, smart Enter and
  Tab, auto-indent).
- Closing a dirty tab asks Save, Don't Save, or Cancel; Don't Save reverts the module to its
  last saved text everywhere, including the analyzer's copy.
- Bookmarks with toggle and next/previous navigation, kept per module.

### The shell

- Project explorer rooted at each open workbook, with context menus, a properties panel, and
  project-qualified addressing so two workbooks can both hold a Module1.
- Panels: Problems, Immediate (a live mirror of the native window), Locals, Watch, and Search
  with find/replace scoped to module, workbook, or all workbooks.
- The native menu bar, Edit menu, and View menu are replaced by the surface's own menus and
  toolbar; debug-only commands grey out outside break mode.

### The Object Browser

- A floating themed window of xlide's own, outside the canvas, carrying the editor's icon: a
  library picker covering the open workbook projects and every referenced type library, read
  directly through LoadTypeLibEx.
- Types and members with VBA-spelled signatures; members of project modules carry their line
  and jump to the definition on double click.
- One search bar with a scope dropdown: Group filters the type list, Object filters the selected
  type's members, All searches both and pulls a whole matched group in.
- A stacked details pane behind a resizable splitter: signature, location, and the library's
  description when it documents the member.

### Under the hood

- The shim is a NativeAOT COM server; registration is per user and needs no administrator
  rights. The engine speaks over a named pipe and dies with its host.
- Regression harnesses pin the close-confirm flow, resize behaviour, and the whole Object
  Browser arc (34 checks: source seams, a headless page drive, and a live editor walk).

### Known limits

- Binaries are not yet code-signed, so no installer is attached to this release; build from
  source with `tools\dev.ps1`. Signing and update plumbing are the next release-engineering
  milestone.
- The debugger integration and the UserForm designer are on the roadmap.
- The analyzer port is in progress; diagnostics today cover the implemented layers.
