# Build status

Updated 2026-09-05, at v0.13.0.

A short snapshot, and deliberately shorter than it was: this is the one document whose only job
is to be true today, and the version of it that described v0.3.0 was still claiming a menu bar
four days and three releases after the fact. Everything that belongs to another document has been
left there. The newest handover (the highest-dated `docs/handoff-*.md`) is what happened and what
is next, [decisions.md](decisions.md) is the choices that would be expensive to reverse,
[lessons.md](lessons.md) is the findings behind them, and [xlide-api.md](xlide-api.md) is the
door. If a fact lives in one of those, it does not live here.

## What the product is

A native COM add-in (NativeAOT, so no .NET runtime is required inside Excel) that loads into the
Visual Basic Editor and replaces its visible surface with a WebView2 page running Monaco. The
native editor keeps running underneath as the text of record, the compile target, and the
debugger; an out-of-process engine supplies diagnostics, completions, and hover.

## Where it is

- **The surface is the whole visible editor.** A toolbar, module tabs, and eight dockable panes:
  Explorer, Properties, Problems, Immediate, Locals, Watch, Tests and Changes. The Object Browser
  is a floating themed window of xlide's own. Nothing native shows through the canvas.
- **There is no menu bar.** All ten of the editor's menus are suppressed; a wrench at the head of
  the toolbar holds the five dialogs that are genuinely the editor's own. Everything the other
  menus carried has a home: modules are added from a plus on each workbook row in the tree and
  removed from the module's own right-click menu, import and export are the sync dialog, and
  running, stepping, compiling and Design Mode are toolbar buttons.
- **It installs from one executable.** `installer\build.ps1` produces `xlide-setup.exe`, 31.6 MB,
  per user, no administrator rights, nothing required beforehand. It refuses to build without a
  language engine or a built page. `tools\release.ps1` attaches it to a tag, refusing to
  ship an engine older than the analyzer checkout it was built from, and hashing the uploaded
  asset against the local one.
- **The window is xlide's**, caption and icon, retaken whenever the editor rewrites its own and
  put back when the add-in unloads.
- **The explorer has two layouts.** Tree is the flat list by kind; Folders groups modules by the
  `'@Folder("Parent.Child")` comment at the top of each one, the Rubberduck convention, read
  from every module opened or not and moved on the typing pause. Both follow the editor under
  the one setting. The status bar names the procedure the caret is in, held to the editor's own
  ProcOfLine line by line, and the tree marks its row (#23).
- **Break mode prints a local by name.** `? counter` answers from the Locals window; an
  expression still declines, because evaluating one adds a procedure and resets the debugger
  (#21).
- **Annotations control the hidden attributes.** `'@PredeclaredId`, `'@ModuleDescription`,
  `'@Exposed`, `'@Description`, `'@DefaultMember`, `'@Enumerator`, `'@ExcelHotkey` and
  `'@VariableDescription` (the Rubberduck convention) are read from the code, compared with what
  the saved module carries, and the drift is filed in the Problems pane with quick fixes: apply
  the module's annotations, add the missing annotation, or take the attribute away. Applying
  exports, rewrites only the attribute lines, and imports the module back (decision 17); the
  analyzer hears about a predeclared class the moment it is applied. [attributes.md](attributes.md).

## What is not done

- **Not signed.** The installer carries no code signature, so Windows warns before running it.
  Signing and update plumbing are the next release-engineering milestone (decision 8).
- **Five refactorings**, three of them landed on 2026-09-03.

  - `rename` and `renameModule`, the two this product started with.
  - **Extract Method**: selected statements lifted into a Private procedure below the one they
    came from, its signature worked out from the analyzer's reference kinds and its refusals
    covering every way the meaning could move. [extract-method.md](extract-method.md) is the
    design and what shipped against it.
  - **Implement Interface**: the `IFace_Member` stub for every member a class promised with
    `Implements` and has not written, each signature copied from the interface's own source text
    rather than rebuilt, each body raising rather than silently doing nothing.
  - **Encapsulate Field**: a public module variable becomes a private one behind a property pair
    that keeps its name, so nothing that used it is rewritten - a `Property Let` for a value, a
    `Property Set` for what VBA assigns with one.
  - **Extract Variable**: a selected expression declared and assigned above the statement it came
    from. The declared type and whether the assignment needs `Set` are the analyzer's answers
    (`resolveExpressionType`, filed as xlide_vscode#61 and landed in 6.1.0), because VBA has no
    assignment form that works for both an object and a value.

  - **Inline Variable**: the pair to the one above - a local replaced by what it was assigned,
    its declaration and assignment taken away. Only an ATOMIC value, because VBA's brackets are
    an evaluation rather than grouping: `Foo (x)` passes by value where `Foo x` passes by
    reference, so parenthesising for precedence would change how every call site binds.

  - **Move to Module**: a procedure taken out of one standard module and put into another, with
    every qualified call site repointed and unqualified ones left alone, because VBA resolves
    those across the project. It refuses a procedure that touches anything Private to the module
    it would leave, naming what would be stranded.

  - **Introduce Parameter**: a local becomes a `ByVal` parameter and every call site is given the
    value it used to be assigned. It refuses when that value names anything a caller cannot see -
    a local, a parameter, or something Private to the module - because the expression reads
    perfectly where it was written and means nothing where it would have to be written instead.

  That is the set. What is left in this area is not more refactorings but the ones that need a
  capability nothing has yet: a rename that follows a symbol into a form's designer, and anything
  that has to reason about what a call does rather than what it names.
- **A project's own conditional compilation arguments are read and honoured.** They live in the
  VBE's Project Properties box, which the object model has no property for, so they come out of
  the saved package's `dir` stream (MS-OVBA record `0x000C`) and ride the project seed. Without
  them every `#If MY_FLAG` is undecidable: both arms analysed, declarations from dead code in the
  table, and findings reportable from an arm the compiler never sees.
- **The UserForm designer's milestones M1 to M6 have all landed**, so it is no longer a remaining
  milestone; [userform-designer.md](userform-designer.md) stays the ground truth for what it does
  and what it deliberately does not.
- **The debugger** is no longer a milestone either. v0.10.0 closed the cluster where a run that
  never happened looked like one - design mode, the current-statement marker, breakpoint state,
  the Tests pane, and the commands that grey.
- **The native pane's rendering is unobserved.** Parity compares text, so a pane that draws wrong
  while holding the right text passes every check there is.

## How it is checked

`tools\verify.ps1` is the whole local gate in one command, in three tiers. Bare, it is
19 headless steps in about ninety seconds: vendored spec, engine currency against the analyzer checkout, the
variant-as-object guard, page and engine typecheck, build and tests, the engine language matrix
and its host-supplied facts, knowledge routes, inline comment features, generated module casing,
the headless page probes, the xlide api's route audit, the Release build, the unit tests, the
check that Release ships the api shut, and the native publish. `-Live` adds four steps that need
an open editor; `-Deep` adds four more and is the pre-release tier, the one to run before a
release rather than before a commit. All 27 take about eight and a half minutes.

Counts move, so they are given as of this line rather than as standing facts: 405 unit tests,
66 api routes of which 64 are driven by one of the 54 suites the gate runs, 2 left out on
purpose.

`tools\page.ps1` is the page loop: typecheck, build, deploy into the running shim, reload, and
prove the running build is the one just made, in about a second and with no restart.
`tools\harness\xlide-api.mjs` drives a live session from a command line or a script, and
[driving-excel.md](driving-excel.md) is how.
