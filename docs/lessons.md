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
