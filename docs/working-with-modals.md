# Working with modal dialogs in the VBE

Every bad hour this project has spent on the editor's own dialogs came from the same
misunderstanding, arrived at four different ways. This document is what was learned each
time, the mechanism underneath, and the rules that follow from it.

The one-line version: a modal dialog is a message loop that belongs to somebody else, and
the only safe way to open one is to have already decided how it will close.

## 1. What a modal actually does to this product

The VBE's dialogs (Add Watch, Edit Watch, Macros, References, and the rest) are ordinary
Win32 dialogs of class `#32770`, raised on the editor's UI thread. That thread is the one
the shim calls the host thread: it owns the object model, the code panes, the surface, and
the placement passes.

When a modal goes up, three things become true at once, and mixing them up is the source of
every failure below.

**The thread does not stop.** A modal runs its own message loop and PUMPS. Posted work still
runs. This was measured on 2026-08-06: with the Macros dialog standing, the xlide api's
marshaled requests kept answering normally, and `state` returned in milliseconds.

**The product's own work does stop.** The shim's poll tick does not run inside that loop, so
debug state, module publishes, and the Locals and Watch feeds all freeze. The clearest signal
is `PerfCounters.HeartbeatAgeMs`: seconds of silence while the api itself answers fine.

**Any call that was already inside the editor is suspended.** A cross-process
`Application.Run` that reaches a breakpoint does not return until execution continues, and
the caller sits inside that call for as long as the dialog or the break stands.

So "the editor is stuck" and "the api is blocked" are different states, and a diagnosis that
conflates them will send you looking in the wrong place. The api being responsive proves
nothing about whether the developer can use the editor.

## 2. Four failures, and what each one taught

### The Add Watch hang: a modal opened with no committed way to close it

A reconnaissance script opened Add Watch with `CommandBars.FindControl(1, 1820).Execute()`
and relied on a HELPER PROCESS to fill the fields and press OK, because `Execute` does not
return until the dialog closes. The helper received its arguments through
`Start-Process -ArgumentList`, which joins the list with spaces and quotes nothing, so the
title `"Add Watch"` arrived as two arguments and shifted every argument after it. The helper
never found its window, never pressed anything, and the developer's screen showed an Add
Watch dialog with an empty Expression field above a frozen editor.

The lesson is not "quote your arguments". It is that the script had no answer for "what
closes this dialog if the plan fails", and a plan whose failure mode is an indefinite hang
is not a plan. Dismissal must be guaranteed before the dialog is opened, not attempted
after.

### The Macros dialog: a command aimed at nothing

Pressing Run through the xlide api opened the Macros dialog and left it standing. Run acts
on the CARET, and the host copies the surface's caret into the native pane before every
command (`SyncCaretToPane`). The caret was on line 1, outside any procedure, so the editor
did what it always does there: it asked which macro to run, modally, and waited for a person.

Two things came out of this. A run aimed at a procedure must put the caret inside it first,
which `revealLine` cannot do because it only scrolls - that gap is why the api has a `caret`
route. And a command that is perfectly valid can still raise a modal, so "this request does
not open dialogs" is never a safe assumption.

### Application.Run: never kill a caller that is suspended inside the editor

To reach a breakpoint from outside, a probe called `Application.Run` from a background job.
The call blocked inside the break, which is correct and expected. The probe then force-killed
the job. Excel crashed and restarted itself.

Tearing down a cross-process call that the host is suspended inside destroys the channel from
under it. If a blocked caller must be abandoned, the break has to be released first (reset,
or continue), and only then may the caller be cleaned up. Better still, do not hold such a
call at all: drive the run in process, through the product's own command path.

### The guard that cancelled the developer's dialog

The first attempt at automatic dismissal compared the standing dialogs against "whatever was
standing when the door last looked". A dialog the DEVELOPER opened between two api requests
was absent from that snapshot, so the guard took it for its own and cancelled it underneath
them.

An automatic dismissal that can touch a dialog a person opened is worse than the hang it
prevents. Attribution has to be positive evidence that this code raised the dialog, not the
absence of evidence that something else did.

## 3. The rules

1. **Decide the exit before the entrance.** Anything that opens a modal programmatically owns
   its dismissal, with a timeout, before the first click.
2. **Never hide a modal you cannot dismiss.** The ghost-palette trick (layer at alpha zero,
   park off screen) works on dialogs too, and that is exactly what makes it dangerous: a
   hidden modal leaves the developer with a frozen editor and no window to close. This is the
   heart of decision 11.
3. **Press only safe buttons automatically.** Cancel, then Close, then No. Never OK, Yes,
   Save, Delete, or Run. A dialog nobody read must not be agreed with, and every safe button
   means "as you were".
4. **Type, do not plant.** The Add Watch dialog rejects text set with `WM_SETTEXT`
   ("Empty watch expression") and accepts the same characters sent as `WM_CHAR`. Assume any
   VBA dialog validates against state that only real input maintains.
5. **Attribute positively.** Only dismiss what you can show you raised. The working rule here
   is a short watch after each request (250ms, 750ms, 1750ms) that claims only dialogs
   appearing in that window; anything that appears while the door is idle belongs to the
   developer.
6. **Diagnose with window enumeration, not with the host thread.** Enumeration, class and
   caption reads, and a posted click all work from any thread. Anything that needs the host
   thread cannot report on the thing holding the host thread.
7. **Post, do not send.** `PostMessage` for a button click: a dialog that raises another
   dialog would hold a `SendMessage` caller inside the nested loop.
8. **Release before cleanup.** Never kill a process that is suspended inside a call into the
   editor.

## 4. What the product carries now

`Diagnostics/DialogWatch.cs` (Debug builds only) enumerates the dialogs standing in the
process, with their buttons, and dismisses one by button caption with the ampersand stripped
so `&Cancel` answers to `Cancel`. Nothing in it touches the host thread.

The xlide api exposes that as `GET dialogs` and `POST dismiss`, and every request that
touches the session first sweeps dialogs the door itself raised, using the safe-button rule.
A request that means to open a dialog passes `keep=1`. `stats` and `dialogs` both report
`heartbeatAgeMs`, which is how a stuck editor is told from a busy one.
[xlide-api.md](xlide-api.md) has the routes.

## 5. Traps worth keeping

- `Start-Process -ArgumentList` joins arguments with spaces and quotes nothing. Any argument
  containing a space silently becomes two. Keep harness arguments single tokens.
- `user32` exports `PostMessageA` and `PostMessageW`, never a bare `PostMessage`. A
  `LibraryImport` spelled without the suffix throws `EntryPointNotFoundException` on first
  call - which is how a latent bug in the pane-close fallback was found, by the dialog watch
  calling the same import.
- The watch dialogs share one template: Expression edit `4853`, Procedure combo `4856`,
  Module combo `4857`, radios `4850`/`4851`/`4852`, OK `1`, Cancel `2`, and Edit Watch's
  Delete `4859`. The full map is in
  [watch-window-investigation.md](watch-window-investigation.md).
- A modal pumps, so a probe that waits on a sleep will usually appear to work and will fail
  under load. Wait on a stated condition instead: the dialog list, or the heartbeat.
- Excel's own dialogs (Macros here) live in the Excel process alongside the VBE's, and both
  answer to the same enumeration. Match on caption before pressing anything.
