# Handoff

Everything needed to pick this up cold. Written at the end of the session that built the
foundation, for whoever continues it.

Read this, then [status.md](status.md) for what is proven, [architecture.md](architecture.md) for
the design, and [decisions.md](decisions.md) for the choices that would be expensive to reverse.

## 1. What this is, in one paragraph

An add-in that upgrades the VBA editor from inside it. A native COM in-process server loads into the
editor, hosts web-based UI in the editor's own docked tool windows, overlays a modern editing
surface on the code panes, and talks to a language engine running in a separate process. It installs
from one executable, per user, with no runtime on the machine.

Repository: `F:\GitHub\xlide\xlide_vbide`, published at
<https://github.com/WilliamSmithEdward/xlide_vbide>.

## 2. Machine setup

Two things are installed but not on the system path, and forgetting either produces a confusing
failure rather than a clear one.

```powershell
# The SDK. Not on PATH; every script prepends it.
$env:PATH = "$env:LOCALAPPDATA\Microsoft\dotnet;$env:PATH"

# Required for ahead-of-time publishing: the native linker is located through vswhere, which is
# also not on PATH. Without it the build fails with the linker's error text embedded in a command
# line, which reads as a linker problem rather than a discovery problem.
$env:PATH = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer;$env:PATH"

# Required to RUN anything built framework-dependent. A built program launches through a small
# native host that finds the runtime through the registry or this variable, never through PATH.
# Without it a program reports that .NET is not installed at all.
$env:DOTNET_ROOT = "$env:LOCALAPPDATA\Microsoft\dotnet"
```

Installed during that session: .NET 10.0.302 SDK (user-local), Visual Studio Build Tools 2026 with
the C++ workload. Node 24.18 and npm 11.16 were already present. Excel 365 x64, VBA 7.1.

## 3. Running things

```powershell
tools\dev.ps1                                  # build, test, register, verify in a real editor
tools\dev.ps1 -Reuse                           # repeat check against an already-open host, ~0.24s
tools\dev.ps1 -Unregister                      # leave the machine clean

tools\harness\Invoke-VbeLoadCheck.ps1          # does the add-in load and connect
tools\harness\Invoke-AnalysisCheck.ps1         # does it report a real defect and not a false one
tools\harness\Get-EditorScreenshot.ps1         # open the editor and capture it, about 2.6s
tools\harness\Get-EditorWindowTree.ps1         # dump the editor's window structure

tools\Compare-Lexers.ps1                       # differential gate for the analyzer port
installer\build.ps1                            # produce xlide-setup.exe

cd engine; npm run build; npm test             # the language engine and its smoke test
cd ui\editor; node build.mjs                   # the editing surface bundle
```

## 4. Lessons that cost real time

Each of these was found by running something, not by reading. Several look like one kind of failure
while being another, which is why they are written down.

### The host does not load add-ins when started through automation

A host created with automation runs in embedding mode and loads no add-ins at all. Registration can
be perfect and the result is silence: no library load, no class activation, nothing in any log. The
harness launches the executable as a process and then attaches.

### A failed connect disables the add-in permanently

When an add-in fails to connect, the editor rewrites its `LoadBehavior` to 0. Every later run then
lists it but never activates it, so the second failure looks nothing like the first and appears to
be caused by whatever changed in between. Every harness restores it to 3 before running.

### Source-generated COM interop does not supply IDispatch

Interop generated from `[GeneratedComInterface]` exposes exactly the interfaces a class declares.
The editor asks an add-in for `IDispatch` as well as the extensibility interface and refuses,
silently, when it cannot get it. Declaring both fixed it. Because the extensibility interface is
dual, its first four members are the dispatch members, so one implementation satisfies each.

### Attaching through the running object table costs tens of seconds

The host publishes itself there lazily, ten to forty seconds after it is visibly usable, and waiting
for that was ninety-eight percent of a check. Asking a worksheet window for its native object model
answers in well under a second and names the instance precisely. See `WorkbookWindowOf` in the
harness scripts.

Two bugs made that technique look unworkable when first tried, and both fail invisibly: a window
class read marshalled as ANSI, so every comparison failed while appearing to return text; and
`0xFFFFFFF0` parsed as a signed value by the scripting runtime, so the call threw before reaching
the API. If a documented technique appears not to work, suspect the marshalling before the
technique.

### Add-ins load when the editor initialises, not when the host starts

A harness that starts the host and waits for the add-in without opening the editor waits forever.

### Terminating the host poisons the next run

The host treats termination as a crash and offers document recovery on the next start, before it is
drivable. Every harness clears the resiliency keys first.

### Reading the screen captures the wrong window

Screen capture cannot block, which makes it tempting, but it captures whatever is actually in front,
and a background process is not permitted to reliably raise a window. The failure is silent and
produces a flawless capture of a different application; it happened, and nothing in the output said
so. Ask the window to render itself instead, with the full-content flag so composited surfaces
appear, and check responsiveness first because rendering can block on a window that is not pumping.

### Static initialiser order across partial classes is undefined

Frozen lookup tables built from field initialisers read generated arrays declared in the other half
of a partial class, and got null. Building them inside a nested type forces the outer type to
initialise first, which is guaranteed rather than incidental.

### A property assignment through dispatch carries a named argument

The value being assigned is identified by a reserved dispatch identifier, not by position. A setter
that passes it positionally fails every time, with no message. This was got wrong here and the
failures were read as the editor refusing to allow those properties to be set, which is what
retired the tool window; see lessons.md 10. The panels live in the editing surface for reasons that
do not depend on it, but do not repeat the inference: a failed call with no message says nothing
about whose fault it is.

### A surface among the panes loses a race it cannot win

The editor raises a pane whenever it activates one, before anything outside the process can react.
The surface is a child of the frame, positioned on the document area, so activating a pane cannot
reorder it. Structural, rather than a faster reaction.

### The editor rewrites the source it is handed

Giving a module its text respells keywords and normalises spacing, so what it holds afterwards is
not what was sent. Read it straight back and adopt its version, or the two drift apart from the
first keystroke.

### Whether a command is enabled is not the execution state

`Reset` is enabled in design mode as well, so it cannot mean "stopped". `VBProject.Mode` reports it
directly. And running does not block the call that starts it, so the state has to be watched for a
while afterwards rather than checked once.

## 5. Editor internals, measured

Full detail in [editor-windows.md](editor-windows.md); re-measure with
`tools\harness\Get-EditorWindowTree.ps1`.

- The editor frame is class `wndclass_desked_gsk`.
- A code pane is class `VbaWindow`. So is the Immediate window. The caption separates them, and
  captions are localised, so the object model stays the authority for which components have panes
  open while window enumeration supplies the handle it does not expose.
- Panes are document children with live splitters between docked panes, so a pane can be resized by
  a change nowhere near it. Anything drawn over one must follow window events rather than sample a
  rectangle once.
- The registry path for add-ins is confirmed by the string embedded in the editor's own library:
  `Software\Microsoft\VBA\VBE\6.0\Addins64` for a 64-bit host.

## 6. How the pieces fit

```text
EXCEL.EXE
  Xlide.Vbe.Shim.dll         native, no runtime loaded into the host
    AddInSession             owns everything released before shutdown, and the editing contract
    CodePaneTracker          which panes exist, where they are, what they show
    EditorSurface            the window over the document area, and the protocol with the page
    OverlayWindow            the window itself, and the two timers the session runs on
    VbeCommands              runs the editor's own commands through its command bars
    AnalysisService          owns the engine and converts offsets to line and column
  WebView2 process           the editing surface

xlide-engine.exe             language engine, named pipe, JSON-RPC one object per line

ui/editor/src
  main.ts                    boots the editor and the shell
  shell.ts                   tab strip, problems panel, splitters
  explorer.ts                the project tree
  toolbar.ts                 the command bar, and which commands belong to whom
  format.ts                  indentation and canonical keyword spelling
  bridge.ts                  every message in both directions
  vba.ts, theme.ts           the language and the themes
```

Everything the developer sees is in the surface. The editor's own project explorer and properties
windows are closed at start-up, because a hidden window cannot be uncovered by anything the editor
does later, and closing them gives the document area their space.

Two rules hold the whole thing together. Nothing expensive runs in the host process. Every COM
object has exactly one owner that releases it exactly once, and everything is released before the
host begins tearing down.

## 7. The analyzer port

The engine currently in use is the TypeScript analyzer from
[the editor extension](https://github.com/WilliamSmithEdward/xlide_vscode), reused rather than
rewritten, packaged as one executable. It works and is validated, and it ships a language runtime
that is ninety megabytes against an add-in of two. The port removes that.

The port is staged and gated. Each layer must agree with the reference implementation on a shared
corpus before the next begins, and the reference stays in the repository as the oracle. The lexer is
done and agrees on all 175 files of the corpus; run `tools\Compare-Lexers.ps1` to confirm.

Five defects were found by that comparison that 48 unit tests missed, and every one was a genuine
misreading of the language. The most expensive: a hash only opens a date literal when the body
between the pair reads as a date. File statements write a file number as a hash followed by an
expression, so pairing it with the next hash on the line swallowed quoted text and commas.

The parser contract is fully mapped and is the next layer: about 3,400 lines, with the boundary
between structured nodes and the raw-statement fallback enumerated exhaustively. Ask for that map
again if it is not in the repository; it was produced by reading
`xlide_vscode\src\analyzer\parser\`.

## 8. The rule that decides the design

**The module is the source of truth. The surface is a view of it.**

The compiler, the debugger, the workbook that gets saved and the analyzer all read the module, and
nothing reads the surface, so an edit that has not reached the module has not happened. Everything
in the editing path follows from this: edits are written back, the module is read straight back
afterwards because the editor rewrites what it is given, and when the two differ the module wins.
The one exception is an edit the developer has not finished, which is never overwritten.

Anything added here should be able to answer: what happens when these two disagree, and how does it
get back in step without the developer losing work.

## 9. Open and known

- **The editing surface takes about two seconds to appear**, of which the whole cost is fetching,
  parsing and evaluating a 3.2 MB bundle; constructing the editor is fifty milliseconds of it. The
  page reports this itself, so the number is in the log on every run. Options not yet tried:
  splitting the bundle, warming the surface during host start-up, or serving it in a form the
  browser can cache compiled.
- **Panels cannot be rearranged.** The layout is a fixed explorer, editor and panel with draggable
  dividers. Dragging one panel to another edge, or stacking panels as tabs, is not built.
- **The Immediate, Locals and Watch windows are still the editor's own**, so they do not match the
  rest. They are left alone deliberately: there is nothing to replace them with yet, and taking
  them away would remove the feature rather than restyle it.
- **The forms designer is untouched.**
- Breakpoints are remembered rather than read: the editor exposes no way to enumerate them. Every
  route to toggling one goes through the same bookkeeping, so the surface and the editor agree in
  practice, but a breakpoint set some other way would be real and undrawn.
- Nothing is signed, so the installer draws a Windows Security warning. This blocks public release
  and nothing else.

## 10. What to do next, in order

1. Cut the surface start-up. Two seconds is the first thing anyone will notice.
2. Make the panels rearrangeable: drag to an edge, stack as tabs.
3. Replace the Immediate window, driving the editor's own hidden one as the evaluator.
4. Continue the port: the parser, then symbols, then the rules, each behind the differential gate.
5. Settings, so the formatting options are the developer's rather than the defaults.
6. Completion and tooltips, backfilled from referenced type libraries.
7. The forms designer, after a spike.
8. Signing and an update mechanism, before any public release.

## 10. Conventions

Plain ASCII everywhere, no em dashes, no emoji. Comments explain constraints and non-obvious
behaviour, never narration. Prose wraps at 100 characters.

Commit messages say what changed and why, and name defects found and what they looked like, because
that is the part nobody can reconstruct later.

Never mention any other add-in product in anything that ships or is published. This is a clean-room
implementation built on documented interfaces.

Report status literally. A check that passes because it did not look hard enough is worse than no
check: the harness in this repository once reported success while the add-in was not connected, and
finding that was worth more than the feature being tested.
