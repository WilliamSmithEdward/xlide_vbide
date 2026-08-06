# Build status

Updated 2026-08-06, after v0.1.5 (the workspace rearrangement is unreleased).

A short, current snapshot. The living documents are [handoff.md](handoff.md) for what happened
and what is next, [decisions.md](decisions.md) for choices that would be expensive to reverse,
and [lessons.md](lessons.md) for the findings behind them.

## What the product is

A native COM add-in (NativeAOT, so no .NET runtime is required inside Excel) that loads into
the Visual Basic Editor and replaces its visible surface with a WebView2 page running Monaco.
The native editor keeps running underneath as the text of record, the compile target, and the
debugger; an out-of-process engine supplies diagnostics, completions, and hover.

## What is proven, and how

- **It loads and renders in a real editor.** `tools\dev.ps1` builds, registers, and launches;
  `tools\harness\Get-EditorScreenshot.ps1` captures the running editor at true DPI. The shim
  log's `editor surface: ready` line itemises every start-up stage, so a regression names its
  own cause. Recent measurements: ready in about 180ms.
- **The surface is the whole visible editor.** Menu bar (now File, Insert, Run, Tools,
  Add-Ins, Help), toolbar, module tabs, and six dockable panes — Explorer, Properties,
  Problems, Immediate, Locals, Watch. The Object Browser is a floating themed window of
  xlide's own. Nothing native shows through the canvas.
- **The workspace is arrangeable.** Editor groups split right and down and pass tabs
  between themselves; tool panes dock in four sections around the editor, each a split of
  tabbed groups, dragged by one gesture with a five-zone compass (decisions 12 and 13).
  Every open module holds a live model, so switching tabs keeps undo, scroll, and
  squiggles. `Test-SplitWorkspace.ps1` is 18 checks over the whole of it.
- **Locals and Watch are fed from the editor's own windows**, floated and made invisible, read
  through UI Automation on a dedicated thread. Both verified against a live break, with
  standing probes: `Test-GhostLocalsPanel.ps1` and `Test-WatchPanel.ps1`.
- **Search is one floating widget** covering module, workbook, and all-workbook scopes, with a
  match table whose rows jump the editor.
- **Unit tests**: 120 pass (`dotnet test`), covering registration, layout arithmetic, edit
  reconciliation, protocol handling, and the analyzer.
- **A debug api** (Debug builds only, compiled out of Release and verified absent from the
  published binary) exposes the running session for diagnosis and automated testing:
  [debug-api.md](debug-api.md). `Test-DebugApi.ps1` is 41 checks. Beyond reading state it
  can await a condition in the page, answer the whole visible layout and reset it, say which
  CSS rule set a property, time the surface, and keep the page's console.

## What is not done

- **Not signed.** No binaries are attached to releases; the product is built from source.
  Signing and update plumbing are the next release-engineering milestone (decision 8).
- **The debugger and the UserForm designer** are the two large remaining milestones. Break
  mode is now reachable from the harness, which is the regression net the debugger needs
  before it starts.
- **The analyzer port** is in progress; diagnostics cover the implemented layers.
- **Watch management stays with the editor's own dialogs** by decision (11), and the watch
  row parse is verified against a real watch.

## Standing probes

In `tools\harness`, each self-describing and PASS/FAIL where it can be:
`Test-DebugApi`, `Test-SplitWorkspace`, `Test-WatchPanel`, `Test-GhostLocalsPanel`,
`Test-CloseConfirm`, `Test-CloseHiddenPane`, `Test-ObjectBrowser`, `Test-ResizeFollow`,
`Test-CloseVbe`. `xlide-api.mjs` drives a live session from the command line or a script.

The page also runs against its own loopback host in a plain browser — two documents, live
tabs, the close-confirm loop — which is where a layout change is exercised without an
Excel: `.claude/launch.json`'s `editor-demo` serves `ui/editor/dist`.
