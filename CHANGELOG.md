# Changelog

## v0.1.2 (2026-08-05)

Debug panel polish from the first live day of the restored Locals tracking.

### Changed

- The Locals and Watch tables no longer spread their columns across the panel: columns
  hug the longest content up to a cap and cluster left, header and rows staying aligned,
  while hover bands and the sticky header still run the full width. The Watch table also
  gains the scroll and monospace styling only Locals had.

### Added

- The Immediate panel has a clear button beside its prompt: it empties the panel's
  history and returns focus to the input. Clearing is page-local by design - the hidden
  native window keeps its text, and the mirror appends only what is new, so cleared
  output does not return.

## v0.1.1 (2026-08-05)

Locals and Watch track the debugger again. The panels had gone dark during breaks - and
taken a day of crashes with them - because the buffer behind every accessibility property
read was sized for 32-bit Windows in a 64-bit process, corrupting eight bytes of stack per
read. Release builds happened to survive the overwrite; the Debug builds the dev loop
switched to did not. [docs/locals-break-investigation.md](docs/locals-break-investigation.md)
has the full story.

### Fixed

- The Locals panel fills at break entry and tracks every step, and the Watch panel rides
  the same pipeline. Root cause: `UiVariant`/`ComVariantBlock` were `Size = 16`; an x64
  VARIANT is 24 bytes.
- In a break with nothing readable, the Locals panel says "No variables to show." instead
  of claiming "Not stopped" - the message now knows the debugger's state.
- An unreadable element or a failed read no longer silences the panel for the session:
  faults are isolated per element, logged once per streak with the failing stage named,
  and retried after a five-second backoff.

### Changed

- The ghost palette readers moved off the editor's own thread onto a dedicated reading
  thread (`GhostReaderThread`): the host never blocks on a read, readings are cleared at
  break exit so a new break starts empty, and the client now has the shape the
  accessibility framework actually supports.

### Known limits

- The context strip (the broken procedure's name) stays hidden: the context box reads
  empty through the in-process channel even though the data is visibly there from outside.
- The Watch row parse is unverified against a real watch; setting one through the native
  Add Watch dialog is the outstanding test.

## v0.1.0 (2026-08-05)

First tagged milestone. xlide turns the Visual Basic Editor into a modern development
surface: a native COM add-in (ahead-of-time compiled, no runtime required on the user's
machine) hosts a browser-based editor over the VBE's document area, with an analysis
engine running as a sidecar process.

### The editor surface

- Monaco-based editing over every code pane: VBA syntax highlighting, folding,
  multi-cursor, find, format-on-demand, and the module tab strip with drag reorder,
  close affordances, and unsaved indicators per workbook.
- Language intelligence from the engine: live diagnostics with severity filters,
  completions, hover, and signature help, plus typing ergonomics (auto-casing, block
  closers, smart Enter and Tab, auto-indent).
- Closing a dirty tab asks Save, Don't Save, or Cancel; Don't Save reverts the module
  to its last saved text everywhere, including the analyzer's copy.
- Bookmarks with toggle and next/previous navigation, kept per module.

### The shell

- Project explorer rooted at each open workbook, with context menus, a properties
  panel, and project-qualified addressing so two workbooks can both hold a Module1.
- Panels: Problems, Immediate (a live mirror of the native window), Locals, Watch,
  and Search with find/replace scoped to module, workbook, or all workbooks.
- The native menu bar, Edit menu, and View menu are replaced by the surface's own
  menus and toolbar; debug-only commands grey out outside break mode.

### The Object Browser

- A floating themed window of xlide's own, outside the canvas, carrying the editor's
  icon: a library picker covering the open workbook projects and every referenced
  type library, read directly through LoadTypeLibEx.
- Types and members with VBA-spelled signatures; members of project modules carry
  their line and jump to the definition on double click.
- One search bar with a scope dropdown: Group filters the type list, Object filters
  the selected type's members, All searches both and pulls a whole matched group in.
- A stacked details pane behind a resizable splitter: signature, location, and the
  library's description when it documents the member.

### Under the hood

- The shim is a NativeAOT COM server; registration is per user and needs no
  administrator rights. The engine speaks over a named pipe and dies with its host.
- Regression harnesses pin the close-confirm flow, resize behaviour, and the whole
  Object Browser arc (34 checks: source seams, a headless page drive, and a live
  editor walk).

### Known limits

- Binaries are not yet code-signed, so no installer is attached to this release;
  build from source with `tools\dev.ps1`. Signing and update plumbing are the next
  release-engineering milestone.
- The debugger integration and the UserForm designer are on the roadmap.
- The analyzer port is in progress; diagnostics today cover the implemented layers.
