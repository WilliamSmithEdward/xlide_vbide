# Build status

Updated 2026-08-10, at v0.6.0.

A short snapshot, and deliberately shorter than it was: this is the one document whose only job
is to be true today, and the version of it that described v0.3.0 was still claiming a menu bar
four days and three releases after the fact. Everything that belongs to another document has been
left there. The newest handover (the highest-dated `docs/handoff-*.md`) is what happened and what
is next, [decisions.md](decisions.md) is the choices that would be expensive to reverse,
[lessons.md](lessons.md) is the findings behind them, and [debug-api.md](debug-api.md) is the
door. If a fact lives in one of those, it does not live here.

## What the product is

A native COM add-in (NativeAOT, so no .NET runtime is required inside Excel) that loads into the
Visual Basic Editor and replaces its visible surface with a WebView2 page running Monaco. The
native editor keeps running underneath as the text of record, the compile target, and the
debugger; an out-of-process engine supplies diagnostics, completions, and hover.

## Where it is

- **The surface is the whole visible editor.** A toolbar, module tabs, and six dockable panes:
  Explorer, Properties, Problems, Immediate, Locals, Watch. The Object Browser is a floating
  themed window of xlide's own. Nothing native shows through the canvas.
- **There is no menu bar.** All ten of the editor's menus are suppressed; a wrench at the head of
  the toolbar holds the five dialogs that are genuinely the editor's own. Everything the other
  menus carried has a home: modules are added from a plus on each workbook row in the tree and
  removed from the module's own right-click menu, import and export are the sync dialog, and
  running, stepping, compiling and Design Mode are toolbar buttons.
- **It installs from one executable.** `installer\build.ps1` produces `xlide-setup.exe`, 29.2 MB,
  per user, no administrator rights, nothing required beforehand. It refuses to build without a
  language engine or a built page. `tools\release.ps1` attaches it to a tag.
- **The window is xlide's**, caption and icon, retaken whenever the editor rewrites its own and
  put back when the add-in unloads.

## What is not done

- **Not signed.** The installer carries no code signature, so Windows warns before running it.
  Signing and update plumbing are the next release-engineering milestone (decision 8).
- **The debugger and the UserForm designer** are the two large remaining milestones.
- **Require Variable Declaration** left with the native Options dialog and has no equivalent in
  xlide's settings. A real gap, recorded where the suppression is.
- **The native pane's rendering is unobserved.** Parity compares text, so a pane that draws wrong
  while holding the right text passes every check there is.

## How it is checked

`tools\verify.ps1` is the whole local gate in one command, twelve steps in about twenty seconds:
vendored spec, engine currency against the analyzer checkout, the variant-as-object guard, page
typecheck, build and tests, the engine language matrix, the headless page probes, the debug api's
route audit, the Release build, the unit tests, and the Release-carries-no-debug-api check.
`-Live` adds the suites that need an open editor.

Counts move, so they are given as of this line rather than as standing facts: 237 unit tests,
50 api routes all documented and driven, 4 headless page probes.

`tools\page.ps1` is the page loop: typecheck, build, deploy into the running shim, reload, and
prove the running build is the one just made, in about a second and with no restart.
`tools\harness\xlide-api.mjs` drives a live session from a command line or a script, and
[driving-excel.md](driving-excel.md) is how.
