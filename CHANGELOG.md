# Changelog

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
