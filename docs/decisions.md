# Decisions

Choices that would be expensive to reverse, with the reasoning that produced them. Superseding a
decision means editing its entry, not contradicting it silently elsewhere.

## 1. The add-in is a native server, compiled ahead of time

Status: decided and proven.

A managed add-in normally loads a runtime into the host. That runtime is shared with whatever else
the host has already loaded, it costs start-up time on a path the user is waiting on, and it
enlarges the failure surface inside a process we do not own.

Compiling ahead of time to a native library removes all of that. The measured result is a 1.91 MB
library that the editor loads directly, with no runtime deployed beside it and none loaded into the
host.

The cost is real and accepted: no reflection-heavy libraries, no runtime code generation, no
framework that assumes a runtime, and a C++ toolchain required to build. Every one of those is a
constraint on us rather than on the user.

## 2. Analysis runs outside the host process

Status: decided.

The editor is single threaded and owns the user interface thread. Analysis of a large project is
measured in seconds on a first pass, and a previous generation of this analyzer needed two separate
rounds of optimisation before large real-world modules were usable at all. Running that work in the
host means either blocking the thread the user is typing on or maintaining a marshalling discipline
across every analysis path forever.

A separate process cannot block the host, cannot leak memory into it, and cannot crash it. It is
also reusable: the same engine already serves an editor extension, so analyzer improvements land in
both products.

The cost is a protocol boundary and process lifetime management.

## 3a. The analyzer is being ported to the product's own language

Status: decided, in progress. Supersedes the reuse decision below, which stands as the description
of what the port replaces and why the reused engine remains the reference.

Reuse got a working, validated analyzer running in days rather than months, and that was the right
call to make first. It carries one cost that cannot be optimised away: the engine ships a language
runtime, which is ninety megabytes against an add-in of two. Nothing about that number is the
analyzer's fault and nothing about it can be trimmed, because it is the runtime, not the code.

A port removes it entirely. Compiled ahead of time the engine becomes single-digit megabytes with no
runtime to start, which also removes start-up latency from the first keystroke after a project
opens. Working over spans of the source rather than allocating strings per token is available in a
way it is not in the original language, and matters on a per-keystroke path.

The port is staged and gated, never a rewrite-and-hope. Each layer is ported bottom up, and every
layer must agree with the existing implementation on a shared corpus before the next one starts. The
existing engine stays in the repository as the reference implementation and the differential oracle:
the question is never "does the port look right" but "does it answer identically". Above that sits
the corpus of cases adjudicated against the real compiler, which is what makes the no-false-positive
discipline testable rather than aspirational.

Until a layer is ported and agreeing, the reused engine is what ships. There is no window in which
the product depends on a half-ported analyzer.

## 3. The engine is reused rather than rewritten

Status: superseded by 3a for the long term; still describes what ships today.

The analyzer exists, is validated against the real compiler across a corpus of accepted and rejected
cases, carries a rule set with recorded evidence per rule, and already meets per-keystroke latency
budgets on modules of tens of thousands of lines. Rewriting it in another language would take months
and start from behind on correctness.

It ships as a self-contained executable so the user needs nothing installed.

Reversing this later is cheap by construction: the protocol is the contract, so an engine in another
language can replace it without the add-in changing. Rewriting is a performance decision to make
against measurements, not in advance.

## 3b. Referenced libraries are read from the host, not shipped as data

Status: decided, not yet built.

The analyzer carries a curated model of the spreadsheet object model: a couple of hundred types and
several thousand members, hand-verified, with a subset marked complete enough to say a member does
not exist. It is the reason completion and diagnostics are good on ordinary code.

It says nothing at all about anything else a project references. A project that references a data
access library, a scripting runtime, an XML parser, or another organisation's component gets no
completion, no signature help, no hover, and no member checking against any of it. That is a large
hole, and it is exactly the code most in need of help, because it is the code the developer knows
least well.

The metadata already exists on the machine. Every reference the editor resolves is a type library,
and a type library is a queryable description of its own types, members, parameters, return types,
and documentation strings. The editor reads them to drive its own completion. So rather than curate
more data and ship it, the engine reads what the user's own project actually references, from the
user's own machine, at the version they actually have.

The mechanism: the project's reference collection gives each library's identity and path; the
library is loaded and enumerated for its types and members; the result is projected into the same
shape the curated model uses and handed to the analyzer as an additional source. Nothing about the
analyzer changes, because it already accepts a host model. The curated model keeps precedence where
the two overlap, since it is verified and carries completeness marks that a raw enumeration cannot.

Three properties this has to have, or it makes things worse rather than better:

Reading a library is slow and its result never changes, so it happens once per library version and
is cached on disk, keyed by the library's identity and version. A project opening for the second
time pays nothing.

It must never be on the path of a keystroke. Enumeration happens in the engine's process, in the
background, and completion works without it and improves when it arrives.

Absence is not evidence. A member that is not found in an enumerated library means the enumeration
is incomplete, never that the member does not exist. Only the curated model's completeness marks can
support a diagnostic that says something is missing. This is the same rule the analyzer already
applies to its own model and the reason it can be trusted; extending the data must not weaken it.

## 4. User interface is web-based, hosted in the editor's own windows

Status: decided, hosting in progress.

The editor can site an ActiveX control inside a native docked tool window. That is the documented
extension point and the only way to get a first-class docked panel rather than a floating window
that does not belong to the editor.

What goes inside that control is our choice, and a browser surface is the strongest one available.
It renders out of process, so interface crashes cannot take the host down. It gives a modern layout
and styling system, real accessibility, and a component ecosystem. Most importantly it makes the
editor surface possible at all, because a full editing component can be hosted in it.

The alternative, drawing everything with native controls or custom painting, makes every feature a
separate project. That is the trap that has historically limited what tools of this kind can offer.

## 5. Calls into the editor use late binding at the control plane

Status: decided.

Early binding to an automation interface requires the declared member order to match the type
library exactly. A mismatch does not raise an error; it calls the wrong function through the wrong
signature. Across host versions that is a memory-corruption bug waiting for a user we cannot debug.

Control-plane calls happen once per user action or editor event, so dispatch overhead is not
measurable. Paths that run per keystroke may use early binding, chosen against a measurement, with
the member order taken from the type library rather than from documentation.

## 6. Registration has one source of truth

Status: decided and tested.

Registration decides whether the add-in loads at all, and a wrong key produces silence rather than
an error. There are three consumers: the development script, the installer, and the tests. Any two
of them drifting apart produces a bug that reproduces only on a machine nobody is debugging on.

All three derive from one type. A test asserts the installer authoring matches it in both
directions, so an entry cannot be added to one without the other.

## 7. Install is per user, from a single executable we wrote

Status: decided. Replaces an earlier choice of a packaged installer format.

Writing class registration under the user hive needs no administrator rights, which removes the
single largest obstacle to someone trying the product. It is also the correct scope rather than a
reduced one, because the editor resolves class registration through that hive.

The installer is an ordinary program of ours, compiled ahead of time into one executable that
carries the product inside it. Three things follow that a packaging format does not give:

The registry layout has one definition, used by the product, the tests, and the installer. A
packaging format needs its own copy of that layout in its own language, which then has to be kept
in agreement by a test. Sharing the code removes the class of bug instead of detecting it.

Installation is verifiable the same way everything else is, by running it. There is no separate
toolchain to install before the installer can be built, which also keeps continuous integration
simple.

Self-update, which the product needs because nothing else will do it, is the same code path as
install rather than a second mechanism.

The trade accepted: no packaged-format deployment for administrators who require one, and no
built-in transactional rollback. Per-user installation of a development tool is the case that
matters, and the installation is small enough that its failure modes are enumerable. A packaged
format can be added later around the same payload if a real need appears.

## 8. Releases are signed

Status: decided, not yet implemented. Blocks public release; does not block development.

An unsigned installer triggers a Windows Security warning naming the executable as unpublishable
("we can't confirm who published xlide-setup.exe"), observed on this machine during the first
install round trip. Developers are the audience most likely to take that warning seriously, and
the uninstaller re-launches itself from the temporary folder, which is a pattern reputation
systems watch closely.

The plan is Authenticode signing of the installer, the shim, and the engine executable in the
release pipeline. Azure Trusted Signing is the current fit: subscription-based, no certificate
files to protect, and it accrues SmartScreen reputation. Local development builds stay unsigned;
nothing in the product may behave differently based on whether it is signed.

## 9. The integration harness owns its host instance

Status: decided and proven.

A harness that attaches to whatever host is already running will eventually drive, and then
terminate, the developer's own session. The harness therefore starts its own instance, confirms by
process identity that it is driving what it started, refuses to run when a host is already open
unless told otherwise, and terminates only the identity it recorded.

It also launches the host as a process rather than creating it through automation, because a host
created through automation does not load add-ins at all, and it restores state the host rewrites
after a failed load. Both are recorded in `lessons.md` with the evidence.

## 10. The product never requires "Trust access to the VBA project object model"

That Trust Center setting gates exactly two doors: `Application.VBE` and
`Workbook.VBProject` - the bridges from the host's own object model into the editor's. It does
not gate an editor add-in: the editor hands its object model to `OnConnection` directly, and
everything reachable from that instance works with the setting off.

So the product commits to the ungated paths, permanently: every project, component, and module
read or write goes through the `OnConnection` instance; execution reaches the host through
`AccessibleObjectFromWindow` on a worksheet window and `Application.Run`. A feature that could
only be built on one of the gated bridges is a feature that asks every user to weaken a security
setting first, and it gets redesigned or dropped. The development harness scripts do use
`Workbook.VBProject` to seed fixtures - they run on a development machine, they are not the
product, and they must never migrate into it.

## 11. Watches are managed by the editor's own dialogs, not replaced by ours

The Watch panel is ours: it renders what the ghost palette holds, and it is the only watch
display the developer sees. CREATING, editing, and deleting a watch stays with the editor's
native Add Watch and Edit Watch dialogs.

The mechanism a replacement would need was measured (2026-08-05) and works: both dialogs are
fully drivable by messages - expression edit 4853 (typed keystrokes only; text planted by
WM_SETTEXT is rejected as "Empty watch expression"), procedure and module combos 4856/4857 via
CB_SETCURSEL, the watch-type radios 4850/4851/4852 via BM_CLICK, OK 1, Cancel 2, and Edit
Watch's Delete 4859, which removes the selected watch cleanly. A themed dialog of ours could
collect the inputs and drive one of these invisibly, the way the ghost palettes turn native
windows into machinery.

It is not worth the failure mode. These dialogs are MODAL: opening one blocks the editor's
thread until something dismisses it. The driver has to run on another thread, and every reason
it might not find the controls it expects - a different Office build, a localised caption, a
timing slip - ends with a modal dialog nobody dismisses. Hidden at alpha zero and parked off
screen, as the mechanism requires, the developer cannot even find the window to close it: the
editor simply hangs, mid-session, with their unsaved work inside it. That happened during the
measurement itself when a driver argument was mangled, and it is a hazard no amount of care
inside our code can remove, because the dangerous half is the timing of a dialog we do not own.

Consequences. The commands themselves - Add Watch (1820), Edit Watch (940), Quick Watch (229) -
must stay REACHABLE FOREVER, because they are the only way to manage a watch. Where they are
reachable FROM is ours to choose: the same day this was decided, the developer moved them out
of the menu bar and into the Watch panel's own button row, so the Debug menu's three items are
suppressed like every other ported item. The rule that survives is the one that matters - these
commands may only leave a surface when another surface of ours already carries them, and the
panel's buttons must never be removed without putting them back somewhere. The panel keeps
proving its half through
`Test-WatchPanel.ps1`, which drives the native dialog from a HARNESS process, where a hang
costs a test run rather than a developer's session. If this is ever revisited, the watchdog
comes first: a driver that guarantees dismissal, and that restores the dialog to a visible,
centred window whenever it cannot, before anything is hidden. The measurements behind this
decision, including the full control map of both dialogs and the hang that settled it, are
in [watch-window-investigation.md](watch-window-investigation.md), and the rules for working
with a modal at all are in [working-with-modals.md](working-with-modals.md).

## 12. The surface holds every open module live, not one at a time

Until v0.1.5 the page held one Monaco model: activating a tab was a host round-trip that read
the module and replaced the model, disposing the previous one, and every message about text -
edits out, syncs in, squiggles - implicitly meant "the shown module". That shape cannot show
two modules side by side, so it goes.

The host↔page document protocol is module-addressed. The page keeps a live model per open
module, keyed by workbook AND name (`xlide:/{project}/{module}` - two workbooks' Module1 are
two models, where the old name-only URI would have collided the moment both were alive).
`openDocument` carries the workbook and text; `contentChanged` says which module it edits;
`syncDocument` and `setDiagnostics` name their module. Which modules are OPEN stays the host's
truth (the object model's pane list, published as before); which model each editor shows is
the page's. Tab activation still tells the host - the native pane underneath follows, because
the debugger and the compiler act on the active pane - but it no longer re-sends text the page
already holds.

The host mirrors the same shape: per-document text, unwritten-edit flags, and write-back
baselines keyed by (workbook, module). The name-only baseline dictionary was a latent
corruption: a line-diff computed against the other workbook's Module1 would have written a
merge of two unrelated modules. One document per key, everywhere, is the invariant.

Engine requests (completions, hovers, canonical case) stay offset-only against the active
module. That is safe because typing requires focus, focus posts `activateModule` on the same
ordered channel before any request the typing produces, and WebView2 delivers messages in
order. If a surface ever gains a way to edit an inactive module's text, those requests must
learn module addressing the same day.

## 13. The workspace layout is ours: hand-rolled groups and docks, no docking framework

### Shape (2026-08-06)

Two systems, deliberately separate. TOOL PANES dock in four sections around the editor -
left, right, top, bottom - each section a split tree of tabbed groups. The EDITOR is one
contained unit in the middle whose module tabs split against each other the same way. A tool
pane never enters the editor unit and an editor tab never leaves it: the editor's tabs are
the host's open panes, and its groups answer to that list.

Both use one gesture, the studio's, from one implementation (`dragcompass.ts`): drag by the
title, and a five-zone compass appears over the region under the pointer - the centre tabs
onto that group, an edge splits beside it, and an edge of the editor area docks against the
editor. The page dims while a drag is live, so a drag reads as a mode.

The compass IS the target. The pointer must come to the zone it means, rather than the code
guessing an intent from where the pointer happens to be: over a wide short region "near the
left edge" and "just left of centre" are a few pixels apart, and a guess there is a coin
toss the developer has to undo. A release off every zone drops nothing.

The preview says which kind of change a release makes. Landing in something that already
stands is a JOIN, outlined at that thing's own bounds - an editor edge whose section
already exists previews the SECTION, not a half of the editor the drop would never touch.
Carving space that does not exist yet is NEW, dashed, because the shape is a proposal.

A drag ends when the window does. Losing focus - alt-tab, a screenshot tool, the host
stealing focus, which this host does freely - Escape, or the document being hidden all
abandon it, because a dim and a compass that outlive the gesture leave the surface looking
permanently mid-drag.

Membership is the host's, geography is the developer's. WHICH modules are open and WHICH
panes exist are answered by the object model and by the shell; where each one sits, and how
big, is the developer's arrangement.

What survives a session is narrower than that sentence used to claim, and the difference is worth
stating rather than discovering. The TOOL PANES persist page-locally: which sections stand, the
tree inside each, which panes are tabbed together and which is showing, every section's size, and
which panes have been closed. The EDITOR's own splits do not. Reopen, and the modules the host says
are open come back as tabs in one group, however they were arranged before.

That is a gap rather than a decision. It was recorded here as though both halves persisted, which
is the kind of error that only shows up when someone relies on it (2026-08-06).

The explorer may not be closed: with every tab shut it is the only route back to a module.
Every other pane has an X on its group and a checkable row in the Panes menu, which is also
how a closed one returns.

Editor splits and movable panels are built in this codebase's own idiom - the tab strip,
splitters, and pointer handling that already survived the host's focus-stealing habits -
rather than adopted from a docking library (dockview was the candidate: MIT, framework-free,
and current). Three reasons. The tab strip is not a generic tab strip: its identity model,
badge and dirty rendering, close-confirm interception, and its defenses against host echoes
rebuilding elements mid-press are product behaviour a library would have to be bent around.
The environment is hostile in measured ways - the host steals focus mid-gesture, pointer
streams end in pointercancel, echoes arrive between press and release - and those lessons are
encoded in code we own, not in a dependency that never met this host. And the surface ships
inside a native add-in where every dependency is a supply-chain and size commitment; the frame
around the editor has needed none so far.

The costs accepted: floating and OS-popout windows are out of scope until built, and drop-zone
polish is ours to maintain. Revisit if the layout ambitions outgrow a split tree and two
docks - the protocol work of decision 12 is what a library would sit on either way, so
nothing here forecloses that.

## 14. The spec's typing helpers are vendored here, not read from a neighbouring checkout

### Shape (2026-08-06)

Smart Enter, Smart Tab, and the lexer beneath them belong to the spec repo (`xlide_vscode`),
and the page bundles that code rather than reimplementing it, so that typing in the VBE
surface and typing in the VS Code extension cannot drift apart. What changes here is only
WHERE the bundler reads it from: a copy under `ui/editor/vendor/xlide-spec`, committed to
this repo, instead of `../../../xlide_vscode/src`.

The build must not depend on a directory outside the repository. Reading a sibling checkout
worked on the one machine that had both repos and nowhere else, which meant the page could
only ever be built by hand - CI could build the C# and package whatever bundle happened to
be committed, but it could not build the bundle. That hole stayed invisible until CI was
asked to build the page and could not resolve a single spec import. A dependency that only
one machine can satisfy is not a dependency, it is a local habit.

Nine files, 2,316 lines, no imports of their own outside the set. The copy is byte-identical
to its source and the bundle it produces is byte-identical to the bundle the sibling produced,
apart from the build stamp - checked, not assumed, because "it should be the same code" is
the belief this whole change exists to stop relying on.

The split that was already there stays: behaviour resolves to the vendored sources, types
resolve to hand-written declarations in `src/spec/xlide-spec`, because the spec compiles under
looser settings than this project's and its sources would not survive this typecheck.

Vendoring buys reproducibility with the risk of a stale copy, so the copy is checked rather
than trusted. `npm run spec:check` compares it against a neighbouring checkout when there is
one, and against its own manifest when there is not; the entry points come from the page's
real imports, so importing something never vendored fails rather than quietly working here
and breaking there. The gate runs it, which puts the drift failure on the machine that has
both repos - the only machine that can answer the question.

The cost accepted: a sync is now a deliberate step after the spec changes something the page
uses, and the hand-written declarations still describe the API by hand. Revisit if the spec
becomes a published package, at which point the vendored directory is what gets published and
the check becomes a version bump.

## 15. Sync form-creation deliberately forks from the shared planner

Status: decided (owner, 2026-08-19).

The shared sync planner (`moduleSyncPlan.ts`, imported from the xlide_vscode checkout) refuses
to create a UserForm from a file. That refusal is a capability statement about the EXTENSION'S
applier, hardcoded into the shared head - and this product's applier is the add-in, which hands
a `.frm`/`.frx` pair to the VBE's own importer and gets the whole form back: controls, fonts,
pictures, code. The one-brain design exists so the two products agree about what a sync MEANS;
what an applier can DO was never rightly the planner's to decide.

The difference is architectural, not a version gap (the owner, 2026-08-19). The extension's
applier imports by writing the CLOSED file's container directly - CFB writes into the .xlsm -
which is precisely what can never work here: the workbook is OPEN in the host, the file is
locked, and the disk copy is stale by definition. This product's applier speaks to the open
project through the VBE instead, which is precisely what a closed-file writer cannot reach.
The two capability envelopes will never converge, so the fork is permanent, not patience.

So the behavior forks, deliberately: a `.frm` whose `.frx` sits beside it in the folder is a
CREATE here. The engine's planner stand-in re-marks the row after the shared planner answers
(`engine/src/sync.ts`). The fork is exactly that wide and no wider:

- Document modules stay refused everywhere. No applier can conjure ThisWorkbook or a sheet
  module from a file; they are born with the workbook.
- A `.frm` alone stays refused. The VBE would fail the import with less to say than the
  planner's warning already says.

The known cost: the stand-in reads the shared planner's row shape, which moves under us -
xlide_vscode 4.0.0 did exactly that, twice in one day. The designer suite's form-sync rows are
the tripwire, and they fired within a day both times. If upstream ships a caller-declared
capability (asked as xlide_vscode#27), the mechanism switches to it and the shape coupling
goes away; the behavior does not change either way.

## 16. Conventions and user directives (binding)

Carried here from the handoff log on 2026-08-09, when the dated handovers were consolidated.
These are the developer's own directives rather than design choices, and they were the only
copy: a handover is a narrative and gets superseded, which is the wrong shelf for a rule that
does not expire.

- Never mention the other add-in product in anything public. Clean-room; cite only Microsoft
  specs and documented interfaces.
- ASCII prose, no em dashes, wrap at 100. Comments explain constraints, never narrate. Commit
  messages say what changed, why, and what the defect looked like.
- Report status literally; a check that passes by not looking hard enough is worse than none.
- The user rejects backwards-compatibility hacks - full refactors are fine.
- No synthetic input (SendKeys) in production, ever.
- The whole UI should end up ours: consistently dark, VS-style ergonomics, the VBE alive
  underneath as the engine. The module is the source of truth; typing follows xlide_vscode.
- Every native window should eventually be replaced by the surface (user, 2026-08-01): Locals,
  Watch, Object Browser, Properties, Call Stack, and dialogs wherever the object model allows a
  faithful rebuild (References and Macros are scriptable; parts of Options are not). Until a
  replacement exists, the native window stays reachable - shown through a punched hole in the
  surface, never by retreating the surface (the toolbar-revert bug, fixed 2026-08-02); the menu
  routing table in RouteMenuCommand is where "open ours instead" gets decided per window.
