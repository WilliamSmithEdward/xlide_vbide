# Findings

Behaviour discovered by running against real Excel, with the evidence that established it. Each
entry cost real debugging time and is not obvious from documentation.

For a shorter list aimed at someone starting work, see [handoff.md](handoff.md). This file is the
long form, with the evidence.

Environment for all entries: Excel 365 x64 build 16.0.20228.20124, VBA 7.1, Windows 11, .NET 10.

## 1. A host started through automation does not load add-ins

A harness that creates Excel with automation gets an instance running in embedding mode, and that
mode suppresses add-in loading entirely. The add-in registration can be perfect and the result is
still silence: no library load, no class activation, nothing in any log.

Evidence: with identical registration, creating Excel through automation produced an empty add-in
list and no library load, while launching `EXCEL.EXE` as a process and then attaching produced both.

Consequence: the integration harness launches the executable and then attaches, verifying by process
identity that it attached to the instance it started. See `tools/harness/Invoke-VbeLoadCheck.ps1`.

## 2. A failed connect disables the add-in permanently

When an add-in fails to connect, the editor rewrites its `LoadBehavior` value to 0. Every subsequent
run then lists the add-in but never activates it, so the second failure looks completely different
from the first and appears to be caused by whatever was changed in between.

Evidence: the registered value was 3 before the first run and 0 after it, with no code path of ours
writing that value.

Consequence: the harness restores `LoadBehavior` to 3 before every run. Without that, a developer
chases a phantom regression after the first genuine failure.

## 3. Source-generated COM interop does not supply IDispatch

Interop generated from `[GeneratedComInterface]` exposes exactly the interfaces a class declares.
The editor asks an add-in for `IDispatch` as well as the extensibility interface, and refuses to
connect when it cannot get it. The refusal is silent.

Evidence: an isolated probe outside the host showed `QueryInterface` for the extensibility interface
returning `S_OK` while `IDispatch` returned `E_NOINTERFACE`. Declaring `IDispatch` on the class
changed the add-in from `connect=False` to `connect=True` with no other change.

Consequence: the add-in class declares `IDispatch` explicitly. Because the extensibility interface
is dual, its first four members are the dispatch members, so one implementation satisfies both.

## 4. Terminating the host poisons the next run

Terminating Excel is normal for a harness, and Excel treats it as a crash. The next start offers
document recovery and may disable items it blames, both of which appear before the instance can be
driven and cause an unrelated-looking failure.

Consequence: the harness clears the resiliency state before each run.

## 5. Ahead-of-time compilation to a native server works, and is small

The shim publishes to a 1.91 MB native library exporting `DllGetClassObject` and `DllCanUnloadNow`,
with no runtime deployed alongside it and no runtime loaded into the host process. Class activation,
interface negotiation, and calls into the editor object model all work from that library.

Building it requires the C++ toolchain for the native linker, and the linker lookup expects
`vswhere.exe` on `PATH`. Without it the build fails with the linker error text embedded in a command
line, which reads as a linker problem rather than a discovery problem.

## 6. A documented technique that appears not to work is usually marshalled wrong

Reaching the host through its own window rather than the running object table was tried, appeared to
fail completely, and was abandoned for a design that cost forty seconds per check. It had two bugs,
neither of which reported anything:

`GetClassName` was declared without specifying character width, so class names were read as ANSI
from UTF-16 data. Every comparison failed while the function appeared to return text.

`0xFFFFFFF0` was parsed by the scripting runtime as a signed value, so the call threw on conversion
before reaching the API at all.

Fixing both made the technique work first time, at 0.68 seconds. The lesson generalises: when a
documented interop technique produces nothing, suspect the declaration before the technique.

## 7. Reading the screen captures whatever is in front, not what was asked for

Screen capture cannot block, which makes it the tempting way to take a screenshot. It captures
whatever is actually on top, and a background process is not permitted to reliably raise a window,
so the request to bring the target forward quietly does nothing.

Evidence: a capture of the editor produced a flawless image of a web browser. Nothing in the output
indicated a problem.

Consequence: the window is asked to render itself, with the full-content flag so composited surfaces
appear rather than leaving a hole, after checking that it is responding, because rendering blocks on
a window that is not pumping messages.

## 8. Inferring what a surface should show put the wrong document in it

A browser surface chose its content by looking for the editing bundle on disk and falling back to
the panel document. Once the bundle existed, the docked panel showed the editing surface.

The failure presented as a rendering fault: a small, black, empty panel. It was the wrong document,
correctly rendered. Content is now stated by whoever creates the surface.

## 9. Late binding removes a class of failure at the boundary

Calls into the editor object model go through dispatch by name rather than through compiled vtable
offsets. This avoids any dependency on member ordering in a particular typelib build, which is a
failure mode that produces memory corruption rather than an error. The cost is irrelevant for
control-plane calls. Paths that run per keystroke will use early binding, measured rather than
assumed.
