# xlide_vbide architecture

This document is the design contract for the project. Code that disagrees with it is either a
bug or a reason to change this document in the same commit.

## 1. What this is

A VBA development environment that runs inside the Excel Visual Basic Editor. The developer
opens the VBE the way they always have; every surface they then interact with is ours.

## 2. The total ownership doctrine

The VBE contains several excellent engines wrapped in a 1998 interface: a compiler, a debugger,
an expression evaluator, a forms designer, and a project model. It also has no editor
extensibility API of any kind.

The doctrine is: **keep every native engine, replace every native surface.**

Native components are never removed and never bypassed. They are made invisible and driven
programmatically, so they remain the authority on correctness while we own the experience.

| Native component | Kept as | Replaced by |
| --- | --- | --- |
| Code pane | Text buffer of record, compile target, debug anchor | Monaco editor surface |
| UserForm designer | Authoritative control model, `.frm`/`.frx` writer | WYSIWYG canvas over the MSForms designer object model |
| Immediate window | Expression evaluator | Our console UI, results marshalled back |
| Locals / Watch | Variable inspection | Our data grid over the same values |
| Debugger | Breakpoints, stepping, break state | Our gutter, our controls, our current-line rendering |
| Compiler | Ground truth for errors | Our diagnostics surface, reconciled against it |
| Project model | Component lifecycle | Our project explorer |

A fallback is a defect. Where a mechanism is not yet proven, the plan names a spike that proves
it, not a downgrade that avoids it. The one honest boundary is stated in section 12.

## 3. Process and component topology

```text
EXCEL.EXE  (single STA, hosts the VBE)
  |
  +-- xlide_vbide shim            native DLL, no runtime dependency, loaded by the VBE
  |     +-- add-in entry point    IDTExtensibility2
  |     +-- window services       WinEvent hooks, subclassing, layout tracking
  |     +-- tool window host      minimal ActiveX control the VBE can site
  |     +-- WebView2 host         one browser surface per window
  |     +-- VBE bridge            COM access to the VBE object model, single owner per object
  |
  +-- WebView2 processes          Chromium, out of process by construction
        +-- editor surface        Monaco, code panes
        +-- designer surface      UserForm canvas
        +-- panels                diagnostics, explorer, tests, console

xlide-engine.exe   separate process, language engine, named pipe transport
```

Three properties follow from this shape:

1. Nothing expensive runs in Excel's process. The shim marshals, hosts, and hooks. Parsing,
   symbol indexing, and analysis happen elsewhere and cannot stall the UI thread or leak into
   Excel's working set.
2. Rendering is already isolated. WebView2 runs Chromium in its own processes, so UI crashes
   cannot take down Excel.
3. The engine is reusable. The same process serves the VS Code extension, so analyzer work
   benefits both products.

## 4. The shim

### 4.1 Runtime choice

The shim is C# compiled ahead of time to a native library with NativeAOT. No .NET runtime is
loaded into Excel. This matters for three reasons: load time is a DLL load rather than a runtime
start, there is no interaction with whatever runtime another add-in has already loaded, and the
failure surface inside the host process is small enough to reason about.

COM is implemented with source-generated interop (`[GeneratedComInterface]`,
`[GeneratedComClass]`), which is the supported path for COM under NativeAOT.

### 4.2 Loading

The VBE loads add-ins listed under `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64` (and `Addins`
for 32-bit hosts). The subkey name is the ProgID; `LoadBehavior` of 3 means load at startup.
ProgID resolves to CLSID, and CLSID resolves to our DLL through `InprocServer32`. Because the
shim is native, `InprocServer32` points at it directly.

Registration is per-user, so installation needs no administrator rights.

The VBE calls `OnConnection` and hands us the root `VBE` object. That is the only handle we need;
the object model is reachable from it.

### 4.3 COM object lifetime

Every COM object obtained from the VBE is owned by exactly one wrapper that releases it exactly
once. The rules, which are not negotiable:

- A wrapper releases only what it created. Objects handed to a caller are not released by the
  caller.
- Enumeration produces wrappers, never raw references left to the garbage collector.
- Event sinks are attached through a single registry that can detach all of them in one call.
- `OnBeginShutdown` detaches every hook, every subclass, and every event sink before the host
  begins tearing down, then releases wrappers in reverse acquisition order.

Teardown crashes in a managed add-in inside an unmanaged host are the best documented failure
mode in this category of software. The mitigation is ownership discipline enforced at the type
level, not care at the call site.

### 4.4 Window services

The VBE object model raises almost no events. Selection changes, focus changes, window creation,
and keystrokes are all invisible to it. These are recovered with Win32:

- `SetWinEventHook` scoped to the VBE process for window creation, destruction, location change,
  focus, name change, and caret movement.
- Subclassing of code pane windows for input interception.
- Layout tracking so our surfaces follow their host windows exactly.

All hook callbacks arrive on the host UI thread and do only enough work to post a message. Real
work happens off that thread.

## 5. The editor surface

### 5.1 Composition

For each code pane the VBE creates, we create a WebView2-backed surface positioned over the
pane's client area and keep it aligned as the MDI child moves, resizes, scrolls into view, or
is activated. The native pane still exists, still holds text, and still participates in
compilation and debugging. It is simply never the thing the developer looks at.

Monaco provides the editing experience: completion, hover, signature help, inline diagnostics,
semantic highlighting, code actions, multi-cursor, folding, find and replace, minimap, bracket
matching, and a real undo stack. These are properties of the surface, so each one costs
configuration rather than a Win32 project.

### 5.2 Synchronisation

The native `CodeModule` is the buffer of record. Monaco is a view that is always converged with
it.

Outbound (developer types in Monaco): edits are coalesced and applied to `CodeModule` through
line-oriented operations. Applying an edit suppresses the inbound path for that revision so a
change cannot echo.

Inbound (something else changes the module): the VBE's own text changes are detected through
window events and content revision checks, then diffed into Monaco as a minimal edit so the
cursor and undo stack survive.

Every applied edit carries a revision number. If revisions ever diverge, the surface resynchronises
from `CodeModule` wholesale rather than guessing. Divergence is logged as a defect, not absorbed.

### 5.3 Input

The surface has focus during authoring, so keystrokes belong to Monaco and no keyboard hook is
needed for typing. A small set of VBE-level shortcuts (compile, run, step, toggle breakpoint) is
forwarded from Monaco to the corresponding VBE command, so muscle memory keeps working.

## 6. The UserForm designer

The MSForms designer is reachable as an object model: a form component exposes a designer, the
designer exposes a controls collection, and each control exposes its properties. That model is
the authority. It is also what writes `.frm` and `.frx`, which means a form edited through it is
byte-compatible with a form edited natively, and forms remain editable by developers who do not
have this tool installed.

So the designer is built the same way as the editor: our canvas renders the control tree, our
inspector edits properties, and every mutation goes through the native model. We gain what the
native designer never had, including snapping and alignment guides, multi-select with distribute
and align, a searchable property inspector with typed editors, z-order and tab-order tools,
undo built on recorded property transactions, keyboard nudging, zoom, and a live outline of the
control tree.

Two mechanisms need proving before this is considered solved, and each has a spike: creating and
deleting controls through the designer model across the full control set, and reading enough
appearance detail (fonts, pictures, borders) to render a faithful preview. Where a property is
readable but not renderable, the canvas shows the control's real bounds and identity rather than
an approximation that lies.

[userform-designer.md](userform-designer.md) carries this from strategy to ground: what the
component plumbing already does (measured 2026-08-13), the full spike list with instruments,
the milestone order, and the page-side architecture question a non-monaco tab raises.

## 7. Debugging

Breakpoints, stepping, and break state are not exposed as methods on the VBE object model. They
are exposed as commands, and commands are reachable.

- Toggle breakpoint: position the caret through `CodePane.SetSelection`, then execute the VBE's
  own toggle-breakpoint command. Breakpoints are additionally recorded by us, which gives the
  feature the VBE never had: breakpoints that survive a session.
- Step into, step over, step out, continue, break, reset: the same command mechanism, driven
  from our gutter and toolbar.
- Break state detection: the VBE main window's caption gains a break marker, and debug command
  availability changes. Both are observable, the first through a name-change window event.
- Current statement: on entering break the VBE moves the caret to the statement about to run.
  We read the position and render the highlight in Monaco.

Immediate, Locals, and Watch remain alive as evaluation engines and become invisible. Our console
sends an expression to the native Immediate window and reads the result back; our inspector reads
the same values the Locals window reads. The evaluator stays native, so results are exactly what
the VBE would print.

Editing during break must match native behaviour, including the cases where the VBE itself resets
the project. Parity is the requirement; refusing to edit is not.

## 8. The language engine

The engine is the XLIDE analyzer, already validated against the real VBE compiler across a corpus
of accepted and rejected cases, already meeting per-keystroke latency budgets on modules of
tens of thousands of lines. It ships as a self-contained executable, so no Node installation
exists on the user's machine.

The transport is a named pipe carrying JSON-RPC with LSP-shaped methods. The protocol is a
boundary, not an implementation detail: it is what allows the engine to be rewritten in another
language later without touching the shim, and what allows the VS Code extension to share it.

Non-negotiable properties inherited from the analyzer's design:

- A diagnostic is reported only when it can be proven. Ambiguity produces silence.
- Severity tells the truth. Red means the compiler will reject this.
- One lex and parse per pass, shared by every rule.
- One incremental project index, not a cache per feature.
- Per-procedure incremental reanalysis, with a full pass only when a declaration changes.

## 9. Installation

One executable, `xlide-setup.exe`, which carries the product inside it and is itself compiled ahead
of time. A user who has nothing installed can run it.

It installs under the user's own profile and writes only to the user's registry hive, so it never
asks for administrator rights. That is not only a convenience: the editor resolves class
registration through the user hive, so per-user is the correct scope rather than a reduced one.

The installer reuses the same registration description that the product and its tests use, so there
is no second copy of the registry layout to drift. This is why the installer is written in the same
language as the product rather than in a packaging tool: correctness by construction beats a test
that compares two descriptions.

Uninstall is complete. It removes both class registrations, both program identifiers, the add-in
entry, the entry in the installed programs list, the installed files, and the log and browser cache
folders. Nothing of the user's is touched: their code lives in their workbooks and is never
modified by installation or removal. The installer refuses to install or uninstall while the host
is running, because overwriting a loaded library produces an installation that appears to succeed
and silently does nothing until the next restart.

Because this is not a managed package format, the product checks for updates itself and can apply
one without the user hunting for a download.

## 10. Testing

Three layers, in descending order of how often they run.

1. Engine tests. Pure, fast, no Excel, no COM. This is where correctness of the language service
   is established, and it is the layer that runs on every commit.
2. Shim unit tests. The parts of the shim that can be isolated from COM: layout arithmetic, edit
   coalescing and diffing, protocol handling, revision reconciliation, registration key
   construction.
3. Integration tests against real Excel. Driven by a harness built on the process-safety
   mechanics proven in the pyVBAharness work: create an owned instance rather than attaching to
   the developer's Excel, tie it to a kernel job object so it dies with the harness, record
   process identity so a stale instance can never be mistaken for ours, watch for modal dialogs
   with a policy that never guesses, and enforce a deadline from outside the blocked apartment.

The integration layer also records every VBE interaction as a structured trace. Traces are
replayed in ordinary unit tests and checked against lifecycle invariants, which is how COM
behaviour gets CI coverage without CI having Excel.

## 11. Development loop

`tools/dev.ps1` builds, registers, launches Excel with a scratch workbook, and opens the VBE, so
a change can be seen in the real host in one command. Unregistering is the same script with a
flag. The shim logs to a file that the harness reads, so a failed load is diagnosable without a
debugger attached.

## 12. The one honest boundary

Everything above is achievable with documented mechanisms. One thing is not fully in our control:
the VBE owns modal states, and while one is up, the host is not ours to drive. This is a property
of the host, not a gap in the design, and the native VBE has exactly the same behaviour. Our
obligation is to never create an ambiguous modal state of our own, and to detect the host's.

## 13. Legal position

This project is a clean-room implementation built on Microsoft's documented interfaces: the VBE
extensibility model, the MSForms designer object model, Win32, and published binary format
specifications. Its language engine is the author's own prior work. No code, structure, or
expression is taken from any other add-in.
