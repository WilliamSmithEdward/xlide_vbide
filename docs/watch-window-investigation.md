# The Watch window: what a real watch taught us

The developer added a watch, entered a break, and the Watch panel said "No watches." This
document records how that was run down, what the editor's watch surfaces actually look like
underneath, and why a plan to replace the native watch dialogs was measured, half-built, and
then abandoned on purpose.

The short version: the panel was reading the window correctly and throwing every row away,
because the parse demanded a header word that real watch rows do not carry. The deeper
finding came from the same session: every control of both watch dialogs can be driven by
message, which makes a replacement UI look easy, and the reason not to build one is that
those dialogs are modal.

## 1. The seam that had never been exercised

The Watch panel arrived on 2026-08-05 as a twin of the Locals panel: the native Watches
window is floated through the object model, made layered at alpha zero, parked off the
virtual screen, and read through UI Automation, with the themed panel rendering what the
reader reads (lesson 29 has the ghost-palette mechanism). Locals could be proven end to end
by a script, because a break produces locals on its own. Watches could not: there is no
object-model API for watches at all, so a watch can only be created through the editor's Add
Watch dialog, and no automated test had ever created one.

The panel therefore shipped with its row parse written by analogy to the Locals parse, and
the handoff said so plainly: "the WatchReader row parse mirrors the Locals pattern but its
exact accessible-name shape is unverified against a real watch". That note turned out to be
the whole defect, sitting in the tree for a day, waiting for someone to add a watch.

## 2. The evidence

The developer's live log during the break:

```text
watch: ghost palette 980852 feeding the panel
watch: 0 row(s)
```

Both lines matter. The first says the ghost was floated, ghosted, and connected: the reader
existed and had an element to read. The second says the reader ran and produced an empty
list rather than failing. A reader that cannot read logs a fault; this one was reading the
window and finding nothing it recognised as a row. That narrows the fault to the parse
before anything is measured.

## 3. Measuring the real row, safely

The rows had to be dumped from outside the process to see their accessible names. That is
the dangerous measurement: two UI Automation clients hitting the editor's provider while it
is stopped can reset the project mid-break, which was measured twice earlier the same day
(lesson 33). The dump therefore ran in DESIGN MODE, where the shim's reader only wakes on
break polls and the outside client has the provider to itself, using a watch created through
the native dialog moments earlier.

What came back:

```text
descendants: 3
  [0] ControlType.List      name=Watches
  [1] ControlType.Pane      name=
  [2] ControlType.ListItem  name= counter Value <Out of context> Type Empty Context BreakProbe.BreakHere
```

Compare a Locals row, measured the day before:

```text
  [5] ControlType.ListItem  name=Expression counter Value 1 Type Long
```

A Locals row leads with the word "Expression". A watch row leads with a SPACE: the Watches
window's first column is the watch-type icon, and an ordinary watch expression has an empty
cell there. The parse required the header word, found none, and rejected every row. The
value "<Out of context>" and the type "Empty" are what a watch reads outside its scope, which
is exactly right for a design-mode dump and is itself a useful confirmation that the ghost
keeps feeding when nothing is stopped.

## 4. The fix

`WatchReader.ParseRow` now anchors on the first " Value " and the last " Type " and
" Context " instead of a leading header word. The first-occurrence rule for the expression is
deliberate and differs from the Locals parse: a locals expression is a single token, but a
watch expression is arbitrary VBA and routinely contains spaces, as in `counter > 40`, so the
expression is everything before the first " Value " rather than a token that must not contain
one.

## 5. Making it a standing test

`tools\harness\Test-WatchPanel.ps1` creates a real watch and asserts the panel gets it. The
mechanics are worth recording, because every one of them cost something to learn.

**The dialog blocks its own caller.** `CommandBars.FindControl(1, 1820).Execute()` does not
return until the Add Watch dialog is dismissed, so the thing that fills the dialog cannot be
the thing that opened it. The probe starts a helper process first, then executes the command;
the helper waits for the window, fills it, and presses OK, which releases the caller.

**The dialog does not believe planted text.** Setting the expression with `WM_SETTEXT` puts
the text in the field visibly, and then OK answers "Empty watch expression". The dialog
validates against state that only real input maintains. Sending the characters as `WM_CHAR`
keystrokes works.

**Assert from the log, never from the ghost.** During the break the probe reads only the shim
log, because dumping the ghost's accessibility tree while the in-process reader is alive can
reset the project. A second non-empty `watch: N row(s)` push after stepping is the proof that
values track, since pushes only happen when the reading's content changes.

The probe passes:

```text
watch: 0 row(s)      <- break entry, before the first read lands
watch: 1 row(s)      <- the watch, read
watch: 1 row(s)      <- after two steps: the value changed, so the push repeated
setWatches stopped:false   <- cleared at break exit
```

## 6. The map of the watch dialogs

Reconnaissance for a possible replacement UI measured both dialogs completely. Recorded here
because the measurement is the expensive part, and it stands whatever is built later.

Add Watch (command 1820) and Edit Watch (command 940) are the same dialog template:

| Control | Id | Notes |
| --- | --- | --- |
| Expression | 4853 | Edit. Accepts typed characters only. |
| Procedure | 4856 | Combo. Readable and settable; holds "(All Procedures)" plus the module's procedures. |
| Module | 4857 | Combo. Holds "(All Modules)" plus every component. |
| Project | 4858 | Static label. |
| Watch Expression | 4850 | Radio, the default. |
| Break When Value Is True | 4851 | Radio. |
| Break When Value Changes | 4852 | Radio. |
| OK | 1 | |
| Delete | 4859 | Edit Watch only. Removes the watch outright. |
| Cancel | 2 | |
| Help | 4860 | |

Everything responds: combos to `CB_SETCURSEL`, radios to `BM_CLICK` with `BM_GETCHECK`
reading back the change, the Delete button to a click, after which the ghost reported zero
rows. The dialog can also be layered at alpha zero and parked off screen, the ghost-palette
trick, and still be driven. In other words a themed dialog of ours could have collected the
inputs and driven one of these invisibly, and watches would have looked entirely ours.

## 7. Why the replacement was abandoned

Both dialogs are MODAL. Opening one blocks the editor's thread until something dismisses it,
so the driver has to run elsewhere, and every reason the driver might not find what it
expects ends the same way: a modal dialog that nobody dismisses and an editor that never
comes back.

That is not hypothetical. It happened during this very reconnaissance. A helper process
received its arguments through `Start-Process -ArgumentList`, which joins the list with
spaces and quotes nothing, so the dialog title "Add Watch" arrived as two arguments and
shifted every argument after it. The helper never found its window, never pressed anything,
and the developer's screen showed an Add Watch dialog with an empty Expression field and a
frozen editor behind it.

In a probe that costs a test run. In the product, with the dialog hidden at alpha zero off
the edge of the screen exactly as the mechanism requires, it costs the developer their
session: there is no window to find, nothing to press Escape on, and unsaved code inside.
The failing half is the timing of a dialog we do not own, which no amount of care in our own
code removes. The feature was abandoned there, and the reasoning is decision 11.

One question stayed unmeasured because the run that would have answered it is the run that
hung: whether selecting a row in the Watches window decides WHICH watch Edit Watch opens. It
matters only to a replacement UI, which is not being built, and any future attempt should
answer it after building a driver that guarantees dismissal, never before.

## 8. What the developer got instead

The panel keeps the display and gains the triggers. Add, Edit, and Quick buttons sit above
the Watch table and run commands 1820, 940, and 229 through the ordinary host command path,
opening the editor's own dialogs, visibly, the way the editor always has. A declined command
says why rather than doing nothing: Edit Watch needs a selected watch, Quick Watch needs an
expression at the caret. With the panel carrying them, the Debug menu's watch items were
suppressed along with the rest of that menu the same evening.

## 9. Traps recorded

- A modal dialog driven from elsewhere is a deadlock waiting for a mismatch. Build the
  dismissal guarantee before the driver, not after.
- `Start-Process -ArgumentList` joins arguments with spaces and quotes nothing. Any argument
  containing a space silently becomes two, and every later argument shifts. Keep harness
  arguments single tokens.
- The Add Watch dialog rejects `WM_SETTEXT` and accepts `WM_CHAR`.
- Never dump a ghost palette's accessibility tree from outside the process while the
  in-process reader is alive in a break. Design mode is safe; break mode can reset the
  project.
- Step Into (188) reports Enabled in design mode, so "wait until the step button enables" is
  not a way to wait for a break. Wait on the shim log's mode line instead.
- Watches are session-scoped: they die with Excel, and deleting a watch's module drops the
  watch with it.
