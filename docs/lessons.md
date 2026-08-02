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

## 10. A property assignment carries a named argument, and forgetting it looks like a refusal

Assigning a property through dispatch is the one call shape that needs a named argument: the value
is identified by a reserved dispatch identifier rather than by position. A setter that passes it
positionally fails every time, and the failure is an ordinary failed call with no message.

This was got wrong, and the wrong conclusion was drawn from it. Two setters passed the value
positionally; every assignment through them failed; and the failures were recorded here as the
editor refusing to allow those properties to be set. **That entry was wrong and has been removed.**
The editor was refusing a malformed call.

What was actually measured about tool windows still stands: a docked one was six pixels high with a
client area measuring a negative height, and its contents did not follow when the window was
resized. Whether it can be sized once asked properly is now unproven either way, because the tool
window was removed before the setter was fixed.

The panels stayed in the editing surface regardless, for reasons that do not depend on this: the
surface can have tabs, splitters, and one consistent layout, and the product owns all of it.

The general lesson is the one worth keeping. A failure that arrives with no message says nothing
about whose fault it is, and "the host refuses to do this" is a conclusion that needs more evidence
than a failed call. Compare against something already known to work: the boolean setter beside these
two had the named argument and had always worked, and the difference between them was the answer.

## 11. A surface among the panes loses a race it cannot win

The editor raises a code pane whenever it activates one, which is every time the user picks a
different module. A surface that is a sibling of the panes has to be raised again afterwards, and
the editor gets there first: the pane being activated is painted, scrollbars and all, before
anything outside the process can react.

Consequence: the surface is a child of the frame, positioned on the document area. Activating a
pane reorders the document area's children and leaves the frame's alone, so nothing comes between
them. This is structural rather than a faster reaction.

## 12. The editor rewrites what it is given

Handing a module its source respells keywords and normalises spacing, so the text it holds
afterwards is not the text that was sent. Every later comparison then sees a difference that is the
editor's doing rather than the developer's.

Consequence: the module is read straight back after every write and its version adopted, applied as
one edit so the undo stack and the caret survive. The module is the source of truth; the surface is
a view of it.

## 13. Whether a command is available is not the execution state

`Reset` looked like a reliable reading of whether execution was stopped, and it is enabled in
design mode as well, so the stopped-line marker appeared before anything had run. `VBProject.Mode`
reports it directly and is neither localised nor inferred.

Running also does not block the call that starts it: the command returns and the code runs
afterwards, so the state at the moment of the command is always "not running yet". A single check
found nothing every time. The state is watched for a while after instead.

## 14. A capture process that is not scaling-aware silently crops

The harness measured the editor at two thirds of its real size on a scaled display and made a
bitmap that size; the window then rendered into it at full size. Every capture lost its right and
bottom edges, which is where the panel is, and the missing panel was read as a panel that did not
work rather than as an image that had been cut.

Consequence: the capture declares per-monitor awareness before measuring anything. It also takes
the largest visible window of the frame class rather than the first, because the editor owns small
windows of that class which are not the frame.

## 15. Reading a property runs its getter, and some getters do real work

The properties panel read the value of everything in a document component's property collection.
That collection is the entire host object surface, roughly one hundred and seventy properties on a
workbook, and it includes the hidden mail-integration ones. Reading a value calls its getter, and
reading a workbook's mail session getter starts a mail logon: on a machine with a mail client and
no profile, selecting a worksheet opened the mail system's "create new profile" wizard. The panel
was also reading every value twice, once for its type and once for its text.

The editor's own Properties window never does this because it filters to the properties the type
library marks browsable, and the dangerous getters are all marked hidden. Until the type library
is read directly, the browsable sets for the known document kinds are spelled out as allowlists,
an unknown document kind falls back to its name alone, names are enumerated before any value is
read (enumerating names runs nothing), and each shown value is read exactly once.

## 16. One dying pane, one give-up, and one create-on-read lined up behind a tab that would not close

Closing a tab closed the real pane and the tab stayed. The module list aborted wholesale when one
pane refused to answer, and the pane that was just closed is exactly such a pane; the tracker had
no retry of its own, so once the editor started refusing refreshes, the window events that
normally drive them all failed identically, one log line each, two thousand in a minute; and a
second click on the dead tab found the pane THROUGH THE COMPONENT, and reading a component's pane
creates one when none is open, so every extra click created a pane in order to destroy it.

Consequences: collection reads that feed the UI tolerate individual members refusing; anything
that gives up on a picture must own a way back to it (the poll timer retries a stale tracker
until the editor answers); repeated identical failures are counted, not repeated, and recovery is
announced; and nothing on a closing path may go through an accessor that creates.

## 17. The development environment's registry writes were a mirage, and every verification read it back

The add-in never loaded in Excel the developer launched, while loading perfectly in every
harness launch, across a full day of registry theories: Click-to-Run masking, App-V overlay
branches, package-data corruption. The real explanation was simpler and worse: the agent tool
shell runs commands in a sandbox that virtualizes registry writes. The per-user registration
was only ever written into the sandbox's private copy-on-write layer. Reads from the same
environment saw the phantom key, so verification always passed; Excel launched from that
environment inherited the same view, so the harness always loaded the add-in. The real user
hive never contained the key at all, which the developer proved in one screenshot of regedit.

Evidence: regedit showed `HKCU\Software\Microsoft\VBA` containing only `7.1`, `Forms3` and
`Trusted` while the same path read from the tool shell showed a fully populated `VBE\6.0`
subtree. A probe module injected into the developer-launched Excel read the key as missing
(0x80070002, which WScript renders misleadingly as "Invalid root in registry key") while the
same probe in a harness-launched Excel read it fine. Elevated child processes escaped the
sandbox, which is why HKLM writes were real and per-user writes were not.

Consequence: registration state must be established from the developer's own context, never
from inside the agent environment — `tools\Register-DevShim.ps1` exists for exactly that, and
verifying persistence means the developer's regedit, not the sandbox's Test-Path. More
generally: when an integration works in every harness run and fails in every human run, suspect
the harness's environment before the product, and prove any "the system is corrupt" theory
against a view of the system the harness cannot have contaminated.

## 18. A posted message that is never dispatched, proven by a line that never appeared

Completion answers were marshalled back to the host thread with PostMessage of an app-range
message to the overlay window, and the handler logged one line when an answer crossed. That line
appears in zero logs across every session ever recorded: the host's component-manager message
loop swallows app-range messages posted to windows it does not manage, so the marshal never ran
once, and the failure was perfectly silent because the drop happened between a successful
PostMessage and a handler that never fired. The same sessions prove WM_TIMER always arrives -
the write debounce and the poll rode it throughout - so the marshal now queues actions and
nudges a zero-delay window timer, which is legal from any thread.

The companion defect: analysis findings were published straight from the engine's reader
thread, and the browser refuses any thread but its own with UI_E_WRONG_THREAD (0x802A000C), so
mid-typing refreshes died and the problems panel served stale findings until a module switch
republished from the host thread - which is also why the defect looked intermittent.

Evidence: grep of every shim log for the delivery line (zero hits ever) against grep for
answered requests (hits in twenty sessions); 0x802A000C on every push since the first
post-ready analysis; both gone after the timer marshal, confirmed live by the developer.

Consequence: a marshal is only trusted when its success is observable - log on delivery and on
drop, never only on send. In this host, cross-thread work rides window timers, not posted
messages. And when a feature works in a demo transport and no live log ever shows its final
hop, the feature has never worked, whatever the demo shows.
