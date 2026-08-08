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

## 19. Three grand theories about a gray strip, and the arithmetic that outlived them

A band of native chrome survived at the bottom of the start-up loader through three
increasingly clever explanations: an aggressive placement heartbeat (shipped, changed
nothing), a DWM border theory (the border was already dark), and a startup-HMENU theory
(built on the shim's own log showing the client 37px shorter during loading — a real
observation, wrongly attributed). The truth was a one-line bug in the fix itself: PixelRect
is four EDGES, and the loading rectangle passed a height where the bottom edge belongs. It
was the first rectangle in the tree with a nonzero top, which is exactly where the
edges-versus-size conventions stop agreeing, so every existing call had been silently
compatible with the wrong reading.

Evidence: a six-second watcher polling the live loading phase five times a second — menu
handle, window rect, client rect, overlay rect, all physical pixels. It showed menu=0
throughout, the client constant, and the overlay bottom exactly one menu-height short; the
2x shortfall in the logged height named the constructor mix-up directly. After the fix the
same watcher showed the overlay bottom equal to the client bottom on every row.

Consequence: when a rectangle type exists, learn whether it is edges or origin-and-size
before constructing one — and treat a helper that works everywhere else as suspect anyway if
yours is the first call with a different shape. And when screenshots breed theories, switch
to numbers: a short polling watch of the live window settles in seconds what reasoning about
compositors cannot settle at all.

## 20. A tree that cycled once a second, and the log line that said the data never changed

The explorer's unfolded class appeared to collapse and expand in a loop. Two theory-led fixes
went out first — stop the accordion following unchanged active-module pushes, stop clearing
fetched outlines on identical project pushes — both real hardenings, neither the cause. The
third pass started from the log instead and ended the hunt in one line: the outline request was
answering every second with exactly 1,565 procedures, every time. The data never changed. The
cycle was the drawing: two update paths redrew the tree unconditionally on the host's
once-a-second pushes, a redraw wipes and rebuilds the rows, and rebuilding a 1,565-row list
resets its scroll — which the eye reads as collapse and expand.

Evidence: the shim log's outline lines, identical count at one-second cadence; the two
screenshots showing both "states" with the chevron open in each.

Consequence: every sink that receives pushed state must be idempotent — identical input changes
nothing, not even a repaint — and any rebuild of a scrolling surface must put the scroll back.
And when a UI loops, read the data cadence out of the log before theorising about state: if the
data is constant, the bug is in the drawing.

## 21. An answer that never came is not an answer of "nothing", and a line being typed is not ready for a verdict

Two defects from the same afternoon, one root shape: treating an absence as a statement.

The first: expanding a large class in the tree flashed its 1,565 procedures and then blanked
them. The outline request's timeout resolved as an empty list, and while the editor spends
seconds absorbing 918KB of module, a timed-out empty could land after the real answer and
replace it. The fix is a vocabulary correction, throughout the pipeline: a timeout or a host
failure resolves as null — "no answer" — and only a real answer, a real empty included, may
replace what an unfolded list already shows. One request per module in flight with at most one
trailing refresh, so answers cannot come home out of order at all. The engine memoises the
outline against the exact source string, and the host stopped shipping the module's whole text
with a request about text the engine's live copy already holds.

The second: the analyzer red-squiggled `MsgBox ` for its argument count while the arguments
were still being typed. The editor extension holds syntax-category findings on the caret's
line, but a semantic verdict about a half-typed line is the same wrong in a different
category. The model the VBE itself uses is the right one: a line is validated when the caret
leaves it. `ActiveLineHold` (Core, pinned by its own tests) is the publish-side of that
contract — typing on a line hides verdicts touching it from the squiggles, the panel, and
the badges at once, because they publish from one filter point; the caret settling anywhere
else republishes from the unfiltered cache with no re-analysis involved.

Evidence: the shim log for the first (outline answers every second, then a
show-transition where the late answers straddled the 2s timeout); the screenshot of
`argument-count` on the line mid-keystroke for the second.

Consequence: give "no answer" its own value the moment a channel gets a timeout, and never
let it share a spelling with "empty". Hold verdicts about text mid-keystroke until the line
is left — the VBE had this right for thirty years. And filter at the single point every
surface publishes through, so squiggles, panel, and badges cannot disagree.

## 22. Two absences that had to be read from the windows themselves

Neither the Immediate window nor a cancelled shutdown announces itself. Both defects came from
inferring the missing announcement badly; both fixes came from reading the state Windows itself
maintains.

The Immediate window exposes no handle, shares its window class with the code panes, and its
caption is localised, so it was identified by closing it and diffing which pane stopped being
visible. The hide is asynchronous: whenever the window list had not caught up, the diff
answered "0 windows changed", the reader never attached, and Debug.Print silently went nowhere
— while evaluation itself worked perfectly, which read as a broken Immediate to anyone typing
in it. The identification that survives timing: the object model NAMES the localised caption,
the handle CARRIES the same caption, visible or hidden, and matching the two cannot lose a
race. The diff remains only as a fallback, and a failed attachment retries on the first
evaluation, by which time the window certainly exists.

The cancelled shutdown: the host asks about unsaved changes AFTER OnBeginShutdown, and Cancel
abandons the shutdown with no callback. The watchdog inferred cancellation from its timer
ticking — but an app-modal dialog pumps timers too, so the session revived while the save
dialog was still on screen, painted the surface over an undecided shutdown, and when the
developer chose Save, the real teardown ripped through a seconds-old session mid-start and
crashed the host. The signal that actually distinguishes the states: an app-modal dialog
DISABLES the application's windows. A disabled frame is a question still open; an enabled
frame held across consecutive ticks is the cancellation.

Evidence: the session logs — "immediate: 0 windows changed when it closed" with "immediate:
ok" answers around it; and the crash log's timeline, revive at +1.5s into the dialog,
OnDisconnection four seconds later, then nothing.

Consequence: when a component gives no notification, do not manufacture one from timing —
find the state the OS already maintains about it (captions, enabled bits, ownership) and read
that. And treat "my timer ticked" with suspicion inside any window of modality: modal loops
pump everybody's timers.

## 23. Two seconds billed to the wrong suspect, and the ready line that itemised it

The surface took 2.1s to boot and the obvious suspects were the obvious ones: 3.4MB of
JavaScript must be slow to compile, or the cache must not be helping. Both were wrong, and one
log line convicted the real cost. The page was taught to itemise its own ready message from
the browser's resource timeline — document arrival, request departure, bytes complete,
compile-and-run — and it read: html 3ms, request 0ms, fetch 2025ms, compile+run 77ms. The
document was instant, the request left at time zero, the compile was noise. The two seconds
were the WebView2 folder mapping brokering the bundle through the browser's host pipe at
about two megabytes a second, every boot, from local disk.

The fix followed from the diagnosis rather than preceding it: serve the same bytes over a
loopback socket (ephemeral port, GET/HEAD only, per-session path token, one directory) and
keep the mapping as the fallback. Boot went from 2164ms to 181ms, fetch from 2025ms to 49ms,
and nothing else changed — same page, same CSP, same everything.

Evidence: the before and after ready lines, in the log, four minutes apart.

Consequence: when a duration needs cutting, make the thing being cut report its own stages
first — the browser already keeps the resource timeline; asking it is one message field. And
keep the itemised line in the log permanently: the next regression in any stage then names
itself. Beware fetch-time labels too — "transfer: cache" from an app-scheme response meant
nothing; the byte count over the socket was the number that could be believed.

## 24. The surface that retreated, and the phantom strip inside the hole

The developer opened any native tool window (Locals, the Object Browser) and the product
visibly reverted: our menu bar and toolbar vanished, the native rows returned. The cause was
the design, not a defect in it: CanCoverChrome said no whenever a native tool window was
visible, and the surface retreated to the document area so the window could be seen. The
retreat gave back the entire chrome to show one window.

The replacement inverts the concession. The surface never retreats once its page is ready: it
keeps the frame's whole client area and punches window-region holes (SetWindowRgn, RGN_DIFF)
exactly where each visible native tool window sits. A hole is real absence, not transparency:
painting, hit-testing and the browser's own child windows stop at the region boundary, so the
native window inside it is fully live while everything around it stays ours. Hole rectangles
come from the object model: each window names its localised caption, the caption finds the
handle (class-free matching, because the Object Browser has its own window class), the handle
has the rectangle. While any hole is open, a 200ms poll re-derives placement, because the
Object Browser moves without a word to the pane tracker; the overlay compares against the
last-applied state and an unchanged layout costs a read and no repaint.

The first capture with the Object Browser open exposed a subtlety worth keeping: the native
menu band appeared INSIDE the hole, above the browser. A maximised MDI child reports a window
rectangle a caption-height TALLER than the document area it sits in - classic MDI merges the
child caption into the menu row, and the phantom strip where the caption used to be extends
above the MDI client, normally clipped invisible by the parent. A hole cut to the raw
rectangle faithfully exposed that strip, and the native band lives there. The fix is the
general truth, not a special case: a child window is clipped to its parent, so the part of it
that can be seen is its rectangle intersected with its parent client area. Clip every hole to
the found window parent.

Evidence: tools\harness\Test-CutoutHoles.ps1 captures - chrome intact with Locals and the
Object Browser each live inside their holes, and the healed capture identical to a plain
boot; the log lines "surface: 2 native hole(s) cut [2,27,638,122;0,130,640,409]" then
"surface: whole again, no native holes".

Consequence: when a cover and a native window contest the same pixels, do not move the cover -
subtract the window from it. Regions make the subtraction exact, the object model names what
to subtract, and clipping to the parent keeps phantom geometry out of it. And when a hole must
track a window nobody announces, poll while the hole exists and make the idle poll free by
comparing before applying.

## 25. The Locals window only lives on screen, and the step that lied about it

The plan was the Immediate pattern a second time: hide the native Locals window, read it with
UI Automation, show the rows in a surface panel. The reader, the message, the panel and the
menu route were all built and the first live break pushed three correct rows. Every probe
after that disagreed with the last one, and the investigation that followed is worth more
than the feature.

What the probes established, five runs deep: the editor feeds the Locals window only when the
window can actually be seen. Hidden, it reads <No Variables> through every break. Visible but
fully covered by the surface, it is fed unreliably - at some break entries, seconds late or
not at all, and never on a step. Any genuinely visible part is enough: half the window
through a region hole tracked every step perfectly. A 3-pixel sliver was not enough, and
neither InvalidateRect, UpdateWindow, nor RedrawWindow with ALLCHILDREN through that sliver
changed anything: the gate is upstream of painting, in whatever the editor consults before
deciding to refresh its debug windows at all.

Two traps inflated the confusion. First, an attach-state machine nobody documents: a window
shown for the first time is fed at the next break; hidden, it detaches immediately; re-shown
while idle it backfills the PREVIOUS break's rows without reattaching; re-shown during a
break it reattaches. Second, and worse: VBA's Step Into stops BEFORE executing the line, so
"step, then check whether counter incremented" reads the OLD value when everything is
working. Half the "stale" evidence was correct values misread through an off-by-one about
step semantics. A probe that asserts freshness must model what the debugger actually executed.

The feature was reverted the same evening: a debug panel that is sometimes right is worse
than the native window through a cutout, and the cutout arrangement already works. The
machinery stays dormant in the tree - reader, message, hidden panel tab - for when the data
has a source that does not depend on the editor's willingness to paint.

Consequence: before building a mirror of a window, probe whether the thing feeding it cares
that it is on screen. A text buffer (Immediate) keeps its content wherever it is; a VIEW
(Locals) can be dead the moment nobody can see it, and no amount of forced painting revives
what the owner never wrote. And when a probe contradicts the last one, suspect the probe's
model of the system before the system: two of five contradictions here were the probe's own
misreadings.

## 26. Covering loses to unclipped painters, hiding loses to the owner, a region wins

The native menu bar kept flickering through the surface on every resize, after the surface
was already covering it and already re-placing itself synchronously inside WM_SIZE. The
bands the menu and toolbars live in - MsoCommandBarDock and their children - paint without
clipping their siblings, so being covered does not stop them: they stamp their pixels
straight over whatever is above them, and the surface repaints a beat later. That is the
whole flicker.

Escalation one, hiding their windows, converted the fight into a different fight: the
editor's own layout shows the docks again on every resize, and the beat between its
ShowWindow and the next placement pass is the same flash. The log told that story plainly -
"chrome bands: hid 1 of 8" repeating once per resize, forever.

Escalation two ended it: give each band window an EMPTY region. A zero-region window owns no
pixels and cannot paint, while its visibility, its layout slot, and everything the editor
believes about it stay exactly as they were - so nothing re-shows, nothing relayouts, and
nothing fights back, because nothing in Office ever sets or resets a region on these
windows. One log line per session: "chrome bands: silenced 8 of 8", set once, holds through
every resize. Restored (SetWindowRgn null) whenever the surface is not covering the chrome
and when the session stops, so a bare editor keeps its menus.

Consequence: to keep a sibling window from showing, pick the mechanism the OWNER does not
manage. Visibility and position belong to the window's owner and it will reassert them;
window regions belong to nobody in Office code, so a region is not a fight but a fact. And
when two mechanisms churn against each other, the deduplicated log names the loser - the
repeated line IS the fight.

## 27. The event that announces a closing window is not an invitation to call its owner

Closing the editor window crashed the host, three times in three minutes, with three
different modules taking the blame - VBE7.DLL, ntdll, the shim itself. The deduplicated
verbose log named the sequence in its last four lines: the frame HIDE event arrived, the
new frame-follow fired placement immediately, placement posted its chrome message, and the
next thing it does - enumerate the editor's windows through the object model for the cutout
pass - never logged. Object-model calls had landed INSIDE the editor's own window-close
handling, three milliseconds after the hide, and the editor faulted under them. The faulting
module varied because a native fault propagating through ahead-of-time frames surfaces
wherever it lands; the trigger never varied.

The fix is a gate, not a retreat: placement checks IsWindowVisible on the frame first. A
hidden frame needs nothing covered, so the surface hides with it and does no other work -
no cutout enumeration, no band regions, no chrome message - and the SHOW event that brings
the frame back re-derives everything. Three scripted close-and-reopen cycles now leave the
host standing.

The same live tests found the cost of the logging that diagnosed this: File.AppendAllText
per line opens the file per line, which invites the antivirus per line, and a host resize
streams thousands of window events per second - the UI thread dragged visibly ("resizing is
slippery"). The log now holds its file open and flushes per line, two orders of magnitude
cheaper with nothing lost in a crash; move events are logged only for the editor's own
window classes; and the chrome message sends only on change.

Consequence: reacting to a teardown announcement by calling the thing being torn down is a
crash with a delay measured in milliseconds - gate on the state the event implies before
doing work in the state that no longer exists. And verbose logging is a feature with a
budget: hold the file open, flush per line, and spend lines on signal, not on another
process's window moves.

## 28. The tracker only holds what it can match, and unchanged is not the same as true

The tab X kept failing for tabs that were not focused, through two correct fixes for two
real races that were not this bug. The third diagnosis started by giving up on theory: one
verbose line per tracker pass, saying exactly what the pass saw. The line repeated through
the entire close - "pass saw 1 [CleanModule] unchanged" before, during, and after closing
BrokenModule - and that repetition WAS the answer. The tracker never contained the hidden
pane at all. Its list holds the pane windows it can match, which in practice is the active
one; closing a hidden pane changes nothing in that picture, the changed event never fires,
nothing republishes the module list, and the strip keeps a dead tab whose next click
reopens the module. Closing the ACTIVE tab always worked, which is exactly why the report
said "if it's not focused".

The durable fix stops trusting change detection for a change the detector cannot represent:
any window destruction arms a moment of polls that re-read the object model's open list and
republish it outright. Destruction strips a window of its class name before the event
arrives, so a dying pane cannot be told from a dying tooltip - every destroy is treated as
"a pane may have closed", and the page's render-key skip makes the republish free when
nothing changed. Two structural repairs rode along: events landing during a refresh are
queued instead of dropped, with a bounded trailing loop that declares the picture stale
when a burst outruns it, and the silent middle of the retry counter now speaks at verbose.

Evidence: the probe's before/after logs - fifteen "unchanged" passes and eternal silence
before; the same passes then "setModules [CleanModule]" 160ms after the destroy, twice,
the second skipped by the page.

Consequence: a cache that cannot represent a state cannot notice entering it. When a
detector says "unchanged" across an event you know happened, ask what the detector CAN see
before asking what went wrong - and when the answer is "not this", re-derive from the
source of truth on a signal cruder than the change itself. And instrument the observer
first: one line per pass, collapsed when identical, named this in one probe run after
theory had missed it three times.

## 29. On screen means a paintable surface, and a layered window always has one

Lesson 25 ended at a wall: the editor only feeds an on-screen Locals window, and on-screen
seemed to mean visible to the developer - which a themed product cannot allow. The wall had
a door. The feed gate is not about being seen; it is about having somewhere to paint. A
window with the layered extended style renders into its own off-screen surface no matter
what covers it or where it sits, so a floated Locals palette with WS_EX_LAYERED at alpha
zero, click-through, parked at -20000,-20000, is fed by the editor as faithfully as one the
developer is staring at. The probe tracked a counter 1 through 4 across four steps with the
window fully invisible and off the virtual screen.

The chain that gets there is all supported surface: LinkedWindowFrame.LinkedWindows.Remove
undocks the window through the object model; geometry is settable on the floating window in
points; the palette is a top-level in our own process, so its styles are ours to edit. The
ghost feeds the accessibility reader, the reader feeds the themed panel, and nothing native
is ever visible. The reverted mirror machinery - reader, message, panel - came back to life
unchanged; only the data source was ever wrong.

Consequence: when a component gates behaviour on visibility, find out WHICH visibility it
means. Style-visible, region-visible, and surface-visible are three different facts, and
the layered style decouples the third from the other two - a window can be impossible to
see and still, as far as its owner can tell, fully on display.

## 30. A grid cell grows freely and shrinks never, when its content sets its own width

The minimap and scrollbar fell off the right edge of the window and stayed there, but only
after the window shrank; growing tracked perfectly. The editor writes an explicit pixel
width onto its own element after every measure, and a grid item's implicit minimum is its
content's width - so the cell could follow the editor up but never come back down below the
editor's last size. A one-way ratchet: every shrink left the editor laid out for a width
the window no longer had, and freeing width elsewhere (narrowing the sidebar) "fixed" it by
letting the propped width fit again. min-width: 0 plus overflow: hidden on the cell lets
the layout shrink first; the editor's own observer then measures and follows down.

Two investigation traps preserved for the future. The host chain was innocent and PROVABLY
so - frame client, overlay, and every Chromium child measured identical while the bug was
on screen - so the divergence had to be inside the page, not the windows. And the browser
pane lab could not reproduce it at first: loading FRESH at the small size never engages a
ratchet whose trigger is shrinking; the repro is grow-then-shrink in one session. The same
lab cannot verify the fix's second half at all: a hidden pane composites no frames, so
ResizeObserver and requestAnimationFrame never fire there - assert stylesheet rules and
class toggles instead, and prove the moving parts with pixel captures of the real window.

Consequence: when a layout follows growth but not shrinkage, suspect min-content propping
before anything in the resize plumbing - and reproduce by walking the same direction the
user walked, not by loading the end state.

## 31. Follow bounds every tick; re-derive the world once, when the ticks pause

Resize latency came back wearing new clothes: a 60-step drag storm logged SIXTY full
placement passes - cutout enumeration, band silencing, module and project publishes, debug
state, resync - because three different event routes (frame subclass, tracker frame events,
tracker pane events) each treated every geometry tick as news. The surface only needs its
BOUNDS per tick, and bounds are cheap and synchronous. Everything else describes WHAT is on
screen, which a drag does not change.

The discipline: every route follows bounds synchronously per event, arms a 150ms settle,
and the full pass runs once when the events pause. Pane events compare what the event
AMOUNTS to - the pane list, the active module, its workbook - and take the fast path when
only rectangles moved. Two exceptions stay event-synchronous on purpose: a hiding frame is
the editor closing (lesson 27 keeps the object model out of that moment), and open cutout
holes must track their native windows exactly. Same storm after: ONE full pass. The
minimap's canvas - repainted a frame behind every layout that moves it - rests during the
storm under a live-resize class and fades back at the settle, so the one artifact the
discipline cannot remove is not seen.

Consequence: in event-driven placement, separate "where its edges are" from "what it is";
the first is per-tick work, the second is per-pause work, and conflating them makes every
mouse move pay for the whole world.

## 32. A document window cannot float, but it can be adopted - if you keep the editor calm

The developer wanted the Object Browser floating beside the surface, and then wanted the
cutout holes gone entirely: a purely xlide canvas. The obvious route - the ghost palettes'
LinkedWindows.Remove - answers the Browser with a SILENT no-op: the call succeeds and
nothing moves, because the Browser is an MDI document window, not a dockable tool window,
and the editor has no floating home for it at all.

The route that works is adoption: reparent the native child into a top-level frame of our
own (owned by the editor's frame, so it rides above like a palette), strip the child's
caption and sizing border, dark-title the frame. Three fights follow, each measured. The
editor notices a "visible" window with no child in sight and closes the adopting frame
about two seconds in - so the object-model window is set not-visible while its HWND lives
on in our frame, the same visible-to-whom split the ghost palettes prove. The hide itself
makes the editor send one reflex WM_CLOSE at the frame - swallowed by a short grace window,
since the developer's own close comes seconds later, not milliseconds. And the native
window's CREATION outruns the command that summons it, so adoption tries at the command,
at the immediate pass, and again at the settle: the flag the object model shows lags the
truth (the same lag the workbook Saved flag has), so gate on the child window existing,
never on Visible.

With the Browser adopted, nothing needed holes any more - Locals and Watches are ghosts,
Project Explorer and Properties are panels - so the cutout machinery went entirely,
replaced by a police pass on the settle: replaced windows re-hidden, a docked Browser
adopted, docked ghost strays hidden. The canvas is purely xlide.

POSTSCRIPT, same day: the adoption LOST. Quieting the record (not-visible) stopped the
closes but shipped a BLANK Browser - the editor only services, fills, and paints a window
its records call on-screen, exactly the ghost-feed rule cutting the other way. Leaving the
record visible kept the content path alive in theory but the reconciliation escalated: the
reflex close also arrives as WM_SYSCOMMAND/SC_CLOSE, indistinguishable from the developer's
own close box, every few seconds. Blank content OR an unwinnable close war - both measured.
The native Object Browser is retired; the toolbar button says so in the status line, and
the real answer is the xlide Object Browser built from the typelib model (#10).

Consequence: when a component refuses to rearrange its own windows, taking the window is
not enough - its OWNER keeps the paint, the layout, and the lifecycle, and it reconciles
against you on its own clock. Adoption works for windows whose content is self-contained;
a window whose owner draws its insides can only be replaced, never stolen.

## 33. An out parameter eight bytes short works until the build changes, then it kills a panel

The Locals panel died on 2026-08-05: every session, every break, "skipped N unreadable
element(s)", zero rows, while the identical accessibility reads from an outside process
returned every row all day. The discarded first fix treated the symptoms - per-element
isolation, backoff instead of a one-fault latch - and the rows still never came.

The crash was one number. UiVariant, the buffer GetCurrentPropertyValue fills, was declared
Size = 16 with a comment calling that "sixteen bytes on x64". Sixteen is the x86 size; a
VARIANT on x64 is TWENTY-FOUR bytes - eight of type tag and padding, then a sixteen-byte
union whose widest member is a record's two pointers. The callee initialises the whole
variant, so every property read wrote eight bytes past the buffer into whatever the caller
kept beside it. Release stack layouts happened to keep dead space there, and the reads
worked for a day of probes and live use. The morning Smart App Control forced the dev loop
onto Debug publishes, the layout changed, a live slot moved into the overhang, and every
read died in a NullReferenceException whose stack pointed at everything and nothing -
memory corruption wearing an innocent exception's clothes.

What found it: not the exception, which named no cause, but the SPLIT. The same reads
out-of-proc (PowerShell UIA client) worked; in-proc they died; raw MSAA in-proc walked the
same tree happily but showed the rows are not in MSAA at all - the grid is served by a
native UIA provider, so the UIA channel is the only channel, and the only difference left
between the working client and the dying one was who marshalled the variant. .NET's client
allocates the real 24; ours allocated 16.

Two consequences. First: when a native fault is build-dependent - works in Release, dies in
Debug, moves between elements across sessions - suspect a buffer the callee writes, and
audit every struct size in the interop layer against the 64-bit ABI, not a comment.
Second: the reads now live on their own thread anyway (GhostReaderThread), because a UIA
client on the provider's own thread inside the provider's process was never a supported
shape - it re-enters the provider mid-call and worked on borrowed luck. The variant fix
made the reads correct; the thread makes them shaped like every client that is supposed to
exist. En route, two smaller truths: the Locals context box is a bare PANE to the
accessibility tree, not an edit, and reads as empty in-proc even when an outside client
sees the full procedure path - the strip stays hidden rather than lying with a blank; and
two UIA clients hitting the editor's provider concurrently (the in-proc reader plus an
out-of-proc probe dump) can kill the BREAK itself - the project resets, panes are
destroyed - so probes must never dump the ghost while the product's reader is alive.

## 34. One keybinding service serves every editor, so a when-clause is the only scoping

Splitting the surface into editor groups meant several standalone Monaco editors on one
page, and the typing automation rebinds Tab and Backspace on each of them. Every editor
created by `monaco.editor.create` shares ONE keybinding service; `editor.addCommand` is not
scoped to the editor it was called on. Two groups produced two identical Backspace rules
with identical when-clauses, both matched, the later registration won everywhere, and its
handler ran `deleteLeft` on ITS editor — so pressing Backspace in the group being typed in
deleted a character in the other group, and the key read as dead.

What confirmed it was the shape of the failure, not the symptom: the live page answered
`prevented: true, handled: false` — the key WAS claimed and the command DID run, so the
binding was alive and pointed somewhere else. A dead key would have been prevented false.

The fix is a context key created on each editor (`editor.createContextKey`), included in
that editor's when-clauses. A context key created this way is true only inside its own
editor's context, so each rule matches exactly the editor it belongs to. Anything else the
surface binds per editor needs the same treatment; the search widget's Escape claim already
had it by accident, because its open flag was per editor for another reason entirely.

Two traps in probing this. A synthetic KeyboardEvent is enough to exercise the keybinding
path — Monaco resolves bindings from the event, and only text INSERTION needs a trusted
event — so a probe can press Backspace without a focused window. But the caret must start
somewhere a Backspace can do work: parked at the start of an empty first line it is
correctly a no-op, and reading that as a fault sent a probe hunting a bug that was not
there. A check that inserts its own text first asserts the key, not the fixture.

## 35. A page that reloads has heard nothing, so every "what changed" cache must reset

The surface holds caches whose whole job is to spare a page that already has the picture:
the last tab list, the last language facts, the last chrome state. They compare what is
about to be sent against what was sent last, and send nothing when the two match.

A page reload — a developer's F5, a devtools reload, a crash recovery — produces a second
`ready` on the same session. The document table survives it (the surface re-opens every
live document from its own copy), but everything else the page draws came from the FIRST
boot's held-message replay, and a second ready has nothing held. Republishing on ready
looked like the fix and was not: `PublishModules` compared its list against the pre-reload
list, matched, and sent nothing. The reloaded page came back with its models and no tabs at
all — and the traffic tap proved it, showing every republish EXCEPT setModules.

So a ready that is not the first must clear the caches before republishing. The rule
generalises: any "has this changed" memo held ACROSS a client that can restart has to be
invalidated when the client restarts, because the memo describes a conversation the new
client was never part of.

## 36. The wrapper takes its own reference, and the finalizer thread is the wrong thread to give it back

Excel died four times on 2026-08-07, wearing three different faces. Two heap corruptions
blamed on `ntdll` at the same fault offset. One access violation blamed on `VBE7.DLL`. One
access violation blamed on `Xlide.Vbe.Shim.dll`. Three libraries, three exception codes,
hours apart, each in the middle of something ordinary: a format, some typing, a probe. They
read as three unrelated instabilities and were nearly filed as "the host is flaky today".

Only the fourth carried a managed stack, and the stack was the whole answer:

```
Marshal.Release(IntPtr)
FreeThreadedStrategy.Release(Void*)
ComObject.Finalize()
__Finalizer.DrainQueue()
```

`GetOrCreateObjectForComInstance` builds a wrapper that takes **its own reference** on the
pointer, on top of the one the caller already holds. With `CreateObjectFlags.UniqueInstance`
that wrapper is not cached, so nothing else will ever give that reference back. Releasing
only the caller's leaves the wrapper alive holding a live editor object until the garbage
collector reaches it, and what runs then is `ComObject.Finalize` **on the finalizer thread**.

The editor's objects are apartment-threaded and belong to the host's thread. Releasing one
from the finalizer thread is not slow or untidy, it is invalid. It reads as an access
violation inside `Marshal.Release`, which ahead-of-time compilation cannot throw, so the
runtime FailFasts and takes the whole of Excel with it. Worse, a release on the wrong thread
corrupts COM's own bookkeeping, and that damage is noticed by whoever touches it next,
which is why the same defect was reported against `ntdll` and `VBE7.DLL` as readily as
against this library, and why nothing connected the reports.

`ComHandle.Dispose` had always done it correctly, and its comment says exactly why. The
missing line was in `DispatchObject.Dispose`, which does nearly all the control-plane work.
One line, absent since the type was written.

**The scale is worth stating.** With the defect restored deliberately, a single `project()`
call leaked **441 wrappers**; reading the native panes leaked 156 per call; moving the caret
29. A short probe run took the live count from 13 to 8,734. Every one of those was a live
editor object queued for release on the wrong thread.

Two lessons, and the second cost more than the first.

**Pair the taking with the giving back so they cannot drift.** The fix is one line, but the
same line can go missing again at the next call site. `ComRuntime.TakeWrapper` and
`GiveBackWrapper` are now the only two doors, and each does its own counting: a caller
cannot dispose without counting or count without disposing. `stats` reports taken, given
back, and live, so a leak is a number during development rather than a crash report later.

**The first instrument was a false-negative machine, and passing built confidence in a
build that was still broken.** Two were tried. A `gc` route that collected and drained the
finalizers on demand, on the theory that it would make the crash deterministic. Measured
against the broken build with 8,734 wrappers pending, it reported completely clean and the
host lived, because releasing an apartment-threaded object from the finalizer thread is only
*sometimes* fatal. And a first version of the counter that incremented beside the disposal
rather than by it, which read perfectly balanced on a build with the disposal deliberately
removed. Both were deleted. **An instrument is not proven by passing on a good build; it is
proven by failing on a bad one**, and the only way to know is to break the code on purpose
and watch. The counter now reports 441 leaked wrappers per call on the broken build and
zero on the fixed one, which is the only reason it is worth having.

## 37. An unattributed crash, recorded rather than explained

Left open deliberately, because a crash nobody can reproduce is still a crash, and the worst thing
to do with one is to quietly decide it was the last bug you fixed.

On 2026-08-07 at 21:01, during a leak sweep, Excel died with an access violation inside
`VBE7.DLL` (`0xc0000005`, fault offset `0x9d0c7`) and **no managed stack**. It is not entry 36's
defect: that one produced a FailFast with `ComObject.Finalize` on the stack, and it was fixed and
proven fixed before this happened. The offset also differs from the other `VBE7.DLL` fault seen
that day (`0x2736ad`, at 19:20).

What was ruled out, and how. The sweep had just gained five state-changing rows, so each was run
alone: pane open and close 16 times, component rename and back 14, breakpoint on and off 12,
module read and write 12, a settings change 12. All survived. The full sweep was then run twice
more, some 500 operations, and survived both. So it is neither a specific operation nor the sweep
as a whole, at least not reliably.

What it might be: a sequence effect, a timing window under sustained load, or a host defect this
product merely provokes. All three are guesses and are recorded as guesses.

**What was actually built instead was the thing that would have attributed it.** `whyDidItDie()`
in the harness client asks the Windows event log the moment a call fails, and `open()` and every
route call now append its answer: the faulting module, the exception code, and the managed stack
when there is one. The plain failure is `fetch failed` and `ECONNREFUSED`, which says only that
nothing is listening; the difference between that and `faulted in VBE7.DLL (0xc0000005)` is the
difference between three crashes filed as "the host is flaky today" and one line that explains
four of them.

The general shape is worth keeping: when a defect resists reproduction, the useful work is often
not another attempt at reproducing it but a better instrument for the next occurrence. This entry
is where the next one gets checked against.

## 38. A suspended frame cannot be unwound by the thread it is suspended on

One mistyped line made the Immediate window useless for the rest of the session, and four
attempts to fix it failed for the same reason before the reason was understood. It is the best
example this codebase has of a bug where every plausible fix is a timing fix and every timing fix
is wrong.

**The defect.** VBA cannot evaluate a string, so a line typed into the Immediate panel is compiled
by writing it into a scratch module and running it by name. `On Error GoTo` inside the generated
procedure catches run-time errors, so `?1/0` comes back as "Division by zero (error 11)" and all
is well. A SYNTAX error is not a run-time error: the project never compiles, the handler is never
installed, and the editor raises its own "Compile error" box instead. Dismissing that box leaves
VBA stopped INSIDE the scratch procedure, with `Application.Run` suspended mid-call. From then on
every evaluation answered "Not available while execution is stopped: evaluating adds a procedure
to the project, which would reset it and end the debugging session" -- blaming a debugging session
the developer had never started, and recoverable only by pressing Reset by hand.

**Why four fixes failed.** A suspended VBA frame unwinds only when the host thread returns to its
message loop. Every attempt ran ON that thread:

1. Check the mode after `Application.Run` and reset. `Run` never returns; the check is beneath the
   suspended frame and is never reached. The log proved it by containing none of its output.
2. Arm a flag in the `finally` for the next evaluation to act on. Same reason, plus a second one:
   `Remove` throws when the project is stopped, so anything after it in that block never ran
   either.
3. Reset on entry to the next evaluation and sleep one second for the mode to settle. The reset
   took 1.24 seconds; the budget expired just before it landed, and the line that provoked the
   recovery was declined anyway.
4. Poll the condition rather than the clock, re-issuing Reset, for eight seconds. It never
   cleared, at any budget.

Attempt 4 is the one that gives the answer away. **No budget can be long enough when the waiting
is itself what prevents the wait from ending.** Holding the host thread in a loop is holding the
one thing that has to happen for the loop to finish.

**The proof, in one measurement.** Issued as its own request -- an ordinary host-thread hop that
returns to the pump -- Reset worked instantly: `break` to `design`, and the next line evaluated to
42. The command was never the problem. The thread it was issued from was.

**The fix.** The recovery moved to the door's own thread, where the same `compile` route already
lives and for the same stated reason: it is the only thread still moving while the editor is
occupied. It asks whether the editor is stopped in the scratch module, posts Reset as its own work
item, waits for the condition on a thread whose waiting costs the host nothing, and only then
posts the evaluation. There is still a deadline, but it bounds a wait for something that can
actually happen rather than substituting for one.

**The general rule, which is bigger than this window.** When a fix looks like it needs a timing
constant, first ask whether the code doing the waiting is holding what it is waiting for. If it
is, no constant is correct, and the constant that appears to work on one machine is the bug
arriving later on somebody else's.

Two smaller truths in passing, both of which cost their own confusion:

- `state.debugMode` is a POLLED value. Read the instant an evaluation returns it reports the poll
  before last: "break" at +0ms and "design" from +500ms, with the evaluation that raised the
  question having already succeeded. A check that samples it is asserting on the clock.
- Not every unresolvable name is a failure. `?ThisDoesNotExist` in a module without Option
  Explicit is a legitimate empty Variant, and the editor's own Immediate window prints exactly the
  same, so a test asserting failure there is testing its author's grasp of VBA rather than the
  product.

## 39. The window being renamed is not the window whose name you took

The product's name was on the editor's title bar at start-up and gone by the time anybody looked
again, for the whole of a session. It was reported twice before it was believed, because every
fresh launch showed it working.

The caption is taken over once when the surface comes up, and retaken whenever the frame window
is renamed, since the editor rewrites its own title as the active project changes. The retake was
guarded on the obvious condition: react when the renamed window IS the frame the chrome owns.
That guard never once held.

Two measurements, and neither is guessable from the code:

- A real rename arrives as `rename event for F10AF8` while the chrome owns `13209A6`. The editor
  retitles its neighbours at the same moments it retitles the frame, and those are the events that
  reach the hook.
- Overwriting `13209A6`'s caption by hand, on the very window the chrome owns and had just
  written to, produces **no event at all**. That window's own renames do not reach this hook.

So the condition being waited for was one that never occurs, and the window that does raise
events was being filtered out for not being the right one.

The fix is to stop asking which window was renamed. Any rename in the editor is a perfectly good
cue to ask whether our title still says what we put there, and `Apply` is built for being asked
often: it reads the caption, compares it against the one it last wrote, and writes nothing when
they match, precisely so a version that wrote unconditionally would not chase the rename it had
just caused. Being asked costs a string compare and answers almost always no.

**A poll was tried first and removed, and the removal is the part worth keeping.** Re-applying on
the session's existing tick looked like the robust answer, and failed twice over: that poll stops
when the editor is idle, so it did not fix the case it was added for; and a caption that changes
at a known moment belongs on that moment's event rather than on a tick that must keep asking a
question whose answer is nearly always no. The developer's instinct, on being shown it, was that a
constant poller is not best practice, and the measurement agreed.

One trap on the way, worth naming because an afternoon went into it. The tracker handles renames
and then returns BEFORE the line that logs window events, so the log contains no rename lines
whether or not any were delivered. That emptiness was read as "the event never arrives", which is
the opposite of the truth. **An absent log line is evidence about the logging, not about the
event.** Instrument the handler, not the search.

## 40. The second wrapper was taken behind the counter's back

Entry 36 ends with a counter that catches a leaked COM wrapper deterministically, proven by
breaking the code on purpose. Excel then died again, on 2026-08-07 at 22:45, with the identical
stack: `ComObject.Finalize`, `Marshal.Release`, FailFast. The counter read balanced throughout, at
its resting 13, before and after.

Both facts were true, and together they say something precise: **a wrapper was reaching the
finalizer that this product never took.** The counter measures wrappers taken through
`ComRuntime.TakeWrapper`. It cannot see one the runtime builds on its own.

The line was this, at the bottom of the variant-to-text conversion:

```csharp
_ => value.As<object>()?.ToString(),
```

For a variant holding `VT_DISPATCH` or `VT_UNKNOWN`, `As<object>()` asks the runtime to build a
managed wrapper over that interface. That wrapper belongs to nobody: not taken through
`ComRuntime`, so the live count never sees it; never disposed, so the finalizer thread releases
it; and the editor's objects are apartment-threaded, so releasing one there is an access violation
the runtime cannot throw. One per property read of an object-valued member.

Interface-valued variants are described now rather than converted, and the numeric and date cases
are converted explicitly so the fallback is reached by almost nothing. Callers that genuinely want
the object go through `FromVariant`, which takes its own counted reference.

**What found it was not the leak instrument.** It was the crash reporter added the hour before,
which reads the Windows event log the moment a suite cannot connect and prints the faulting module
and the managed stack. Without it this was one more "Excel died again" in an afternoon that had
several, and the stack that named the cause would have scrolled past unread.

Two lessons, and the second is the one worth carrying:

**An instrument measures what it was pointed at.** The counter is honest and was proven both ways,
and it is still blind to every wrapper this product does not create. Proving an instrument catches
the bug you built it for says nothing about the bug you have not met.

**So keep the instrument that observes OUTCOMES beside the one that observes MECHANISM.** The
counter watches a mechanism this product controls; the crash reporter watches what actually
happened to the process, whatever the cause. The second one found what the first could not see,
and it cost twenty lines.

## 41. Application.Run picks the active workbook, and the scratch module is in another one

Found while running the Immediate suite against a session holding TWO workbooks, which is a
supported state and one that three separate defects have already lived in.

A line typed into the Immediate panel is compiled into a scratch module added to
`ActiveVBProject`, and then run by name through `Application.Run`. Those are two different
notions of "current": the VBE's active project, and the host's active workbook. With one workbook
open they are always the same and the difference cannot be seen. With two, they can differ, and
the evaluation fails with the host's own words:

> Cannot run the macro 'XlideImmediateScratch.XlideImmediateRun'. The macro may not be available
> in this workbook or all macros may be disabled.

Which is true and unhelpful: the macro exists, in the other workbook.

Fixed by qualifying the name with the workbook the scratch module actually went into:
`'Book.xlsm'!Module.Procedure`. That resolves the ambiguity in favour of THE EDITOR'S active
project, which is the right answer here: a line typed into this product's Immediate panel is about
the workbook the developer is looking at in this product, not about whichever Excel window happens
to be in front. A project that has never been saved has no name to qualify with, and there the
unqualified form is all there is.

The other candidate fix, adding the scratch module to the host's active workbook instead, would
have resolved it the opposite way and evaluated against a workbook the developer is not editing.

Measured: with both fixtures open the suite went from 10 passed and 6 failed to 16 and 0.

## 42. A project that goes clean said nothing, so the errors never went away

Found by driving the api by hand rather than by running a suite, which is worth noting on its own:
the question was "what does `problems()` say right now", and the answer named a module and a line
that could not possibly hold it.

`problems()` reported `HelpersExtra 6:11 Wrong number of arguments to 'Close'` for a module whose
entire text is seven lines and contains no `Close` at all, and `Consumer 17:18 Method or data
member not found: 'HelpersExtra.Thing'` for a `Thing` that is declared right there. Both findings
described text that had been replaced some minutes earlier.

The engine's live copy was CORRECT, which ruled out the obvious explanation and pointed at the
publishing rather than the analysis. The pass had run, the log said so, and it said
`renamefixture.xlsm produced 0 finding(s)`. And there it was:

```csharp
if (findings.Count > 0)
{
    FindingsReady?.Invoke(findings);
}
```

**A project with nothing wrong published nothing, so the last non-empty set stood forever.** Fix
the last error in a workbook and it stays on screen, on a line that no longer holds it, until some
other error happens to be found. The receiver was already written for the empty case; it was never
given one.

A second defect was in the same four lines. The publish was per project, from inside the loop over
projects, and the receiver REPLACES the whole set. With two workbooks open the second project's
findings wiped the first's, so half the errors were invisible depending on which was analysed
last. One workbook and that cannot be seen; two is a supported state.

One list for the whole pass, published once, unconditionally. An empty list is not "nothing to
say", it is "there is nothing wrong any more", and it is the only thing that clears a finding the
developer has just fixed.

**THE REGRESSION CHECK TOOK THREE ATTEMPTS TO BITE, and each failure was a different way of not
reproducing the state.** This is the more useful half of the entry.

1. The module was left OPEN. An open module is re-analysed live on every pause, and that path
   publishes per module and clears its own findings, so the error retired whatever the project
   pass did. The check passed against a build with the defect deliberately restored.
2. The pane was closed, but only the offending LINE was removed. That left the module without the
   procedure its caller expects, so another module complained, the pass was non-empty, it
   published, and the stale finding was cleared by accident. Passed again.
3. The module was restored to its ORIGINAL text, so the whole project went clean. Only then did
   the pass produce zero findings, and only then did the check fail on the broken build, naming
   both stale findings exactly.

The rule: **a check must reproduce the state, not a state that resembles it.** Each of the first
two was a fair description of "the code was fixed" and neither reached the condition the bug needs,
which is a project with NOTHING left to report. When a check will not fail on a build you have
broken on purpose, the useful question is not "is the fix wrong" but "what is different between
what I set up and what actually happened".

## 43. Reveal is not a click, and the action that reported success moved nothing

Found in a minute of driving the api by hand, after a manual sweep printed something odd: seven
tabs open, and the surface holding text for exactly one of them.

`act("activate", {module})` answered `did: true, "revealed Consumer"` while the page, the surface
and the native pane all stayed on the module that was already showing. It did that for every
module, with the project argument, without it, and with the display name instead of the id. It had
never worked, and it had always said it had.

Two separate mistakes, and the second is the one that hid the first.

**The wrong method.** It called `workspace.reveal`, which shows a document THE PAGE ALREADY HOLDS
and tells the host nothing. The page holds a document for a module once it has been activated, not
merely because its pane is open, so against any tab whose text had never been fetched reveal found
nothing to show and returned silently. A tab CLICK goes through `selectTab`, which shows it
page-locally and asks the host to activate the native pane behind it. That is the state a click
leaves, so it is the state the api action has to leave: the prime heuristic, in one line.

**Reporting the request.** `return { did: true }` was unconditional. This repo has fixed exactly
this before, in `closeActive`, whose comment says it reports the OUTCOME because closing a dirty
module raises a box and the tab stays. The lesson did not travel to its neighbour in the same file.

There is a third turn worth keeping. With the right method wired in, the action still reported
`did: false` on the FIRST visit to each module and `did: true` on the second, because a document
the page has never held arrives from the host asynchronously: the check was reading the model that
was there before. The fix is the shape `closeActive` and `format` already use, which is to wait
for the outcome before reporting it. **"Report the outcome" and "report it synchronously" are
different claims, and the first does not imply the second.**

Worth noting how it was found. Nothing was broken on screen, no crash, no failing suite: four
suites passed against this defect for as long as it existed, because they reached the state they
needed through `pane("open")` rather than through the tab strip. What surfaced it was printing
several unrelated readings side by side and noticing that two of them could not both be true.
