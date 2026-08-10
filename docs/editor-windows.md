# Editor window structure

Measured on Excel 365 x64 (16.0.20228.20124), VBA 7.1, Windows 11, by
`tools/harness/Get-EditorWindowTree.ps1`. None of this is documented by the vendor, and the editor
object model exposes no way to ask which window shows which module, so it is established by
observation and re-measurable at any time by running that script.

## What matters

A code pane is a window of class `VbaWindow`, captioned with the component name and a kind suffix.

```text
wndclass_desked_gsk   "Microsoft Visual Basic for Applications - [ProbeModule (Code)]"
  MDIClient
  VbaWindow           "ProbeModule (Code)"      1392 x 916
  VbaWindow           "Immediate"               1380 x 79
  PROJECT             "Project - VBAProject"
  wndclass_pbrs       "Properties - ProbeModule"
  VBSlider                                       splitters between docked panes
```

Three consequences for the editor surface.

The class alone does not identify a code pane. The Immediate window is also a `VbaWindow`. The
caption distinguishes them: a code pane's caption ends in `(Code)`, and the leading text is the
component name. Captions are localised, so matching on the suffix is a heuristic and must never be
the only evidence. The object model's `CodePanes` collection is the authority for which modules have
panes open; window enumeration supplies the window handle that collection does not expose. The two
are correlated, not substituted.

Panes are MDI children. They move, resize, maximise, and restack as a group when the user rearranges
the editor, and the surface positioned over one has to follow every one of those. That is what the
window event hooks are for: location change, show, hide, destroy, and focus, scoped to the editor's
process.

Docking is real and dynamic. `VBSlider` windows are the splitters between docked panes, so a change
anywhere in the layout can resize a pane without the pane itself being touched by the user.

## Window classes seen

| Class | What it is |
| --- | --- |
| `wndclass_desked_gsk` | The editor's own top-level window |
| `MDIClient` | Container for the document children |
| `VbaWindow` | A code pane, and also the Immediate window |
| `VBMdiChildHack` | Present alongside the document area; purpose not established |
| `PROJECT` | The project explorer pane, containing a `SysTreeView32` |
| `wndclass_pbrs` | The properties pane |
| `VBSlider` | A splitter between docked panes |
| `MsoCommandBar`, `MsoCommandBarDock` | Menus and toolbars, shared with the host |
| `ObtbarWndClass` | The pane-splitter control at the bottom left of a code pane |

## Reproducing this

```powershell
tools\harness\Get-EditorWindowTree.ps1
```

Starts a host, adds a component so a real pane exists, opens it, and prints the tree with each
window's class, visibility, size, position, and caption, followed by the panes the object model
reports. Add `-KeepOpen` to inspect the result by hand.

## Measured host facts (re-measure only if a host build disagrees)

Carried here from the handoff log on 2026-08-09. The window classes below overlap what is
above; the registration mechanism was recorded nowhere else.

Registry: `HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64`, subkey = ProgID. Frame class
`wndclass_desked_gsk`, document area `MDIClient`, and the code panes, Immediate, Locals, Watch and
Object Browser windows all share class `VbaWindow` - a window's class says nothing about what it
is; only the object model knows.

Add-in discovery is **HKCU-only** - vbe7.dll's own strings carry
`HKEY_CURRENT_USER\Software\Microsoft\VBA\VBE\6.0` and `...\Addins64`, and Rubberduck's installer
writes the Addins keys under HKCU even when elevated (HKLM there is only .NET COM classes). There
is no HKLM enumeration to fall back on, and the per-user registration needs no elevation - that
is the entire documented mechanism.

**The dev-loop registration mirage (lesson 17, read it):** the agent tool shell virtualizes
registry writes, so registrations written from it exist only in the sandbox's private layer.
Harness-launched Excel inherits that layer and loads the add-in; Excel the developer launches
reads the real hive, which never got the key. A day went to Click-to-Run/App-V theories before a
regedit screenshot showed the truth. Registration on the dev machine is therefore done by the
developer running `tools\Register-DevShim.ps1` in their own terminal (no admin); the published
shim path is stable across rebuilds so one real registration survives the dev loop. Elevated
child processes escape the sandbox (UAC-relaunched writes land for real) - useful, but not a
substitute. Verify persistence with the developer's regedit, never with in-sandbox reads.

Click-to-Run keeps machine-level VBA values (`Vbe71DllPath`) only inside its overlay
(`HKLM\...\ClickToRun\REGISTRY\MACHINE`); real HKLM has no VBA branch on C2R machines. The
installer writes the per-user registration only and never prompts for elevation; when it already
runs elevated on a C2R machine it also plants the overlay copy, and `--overlay-only` runs that
step deliberately (`tools\Register-InOfficeOverlay.ps1` is the standalone form).
`Register-MachineWide.ps1` (real-HKLM mirror) was retired - the VBE never reads that hive.

`Window.Type`: 0 code, 2 object browser, 3 watch, 4 locals, 5 immediate, 6 project, 7 properties,
10 linked frame, 15 our tool window. `VBProject.Mode`: 1 = break, 2 = design.

Command IDs (measured by enumerating CommandBars): Run 186, Break 189, Reset 228,
ToggleBreakpoint 51, StepInto 188, StepOver 194, StepOut 2559, RunToCursor 1811, QuickWatch 229,
AddWatch 1820, EditWatch 940, CallStack 620, Compile 578, ClearAllBreakpoints 579,
SetNextStatement 1812, ShowNextStatement 1813, Comment 192, Uncomment 2552, Save 3, Undo 128,
Redo 129, Find 141, Replace 313, ObjectBrowser 473, ImmediateWindow 2554, LocalsWindow 2555,
WatchWindow 2556, ProjectExplorer 2557, PropertiesWindow 222, References 942, Options 522,
Macros 930, ProjectProperties 2578, InsertProcedure 559, InsertUserForm 512, InsertModule 3039,
InsertClassModule 2579, Import 524, Export 525.

**Menu item IDs are NOT unique.** 746 is shared by New Project, Close Project, Remove <module>,
Make, four Insert placeholders, Digital Signature, and MSDN on the Web; 830 by every window-list
entry; 761 by every toolbar toggle. The current command set is by-ID and every ID in it is unique
(checked against the full dump), but menu replication MUST execute by path
(`bar.Controls.Item(i).Controls.Item(j).Execute()`), which is verified working live.

`MenuBar.Visible = false` returns E_FAIL - a genuine refusal (the identical call hides the
Standard toolbar). The menu bar cannot be hidden, only covered.

DWM: attribute 20 (19 on older builds) = dark title bar; 34 border colour, 35 caption colour,
36 caption text; values are COLORREF `0x00BBGGRR`; refused harmlessly before Windows 11. The
frame's client area is inset ~11 px and a 1-2 px pale line at that inset is drawn by the frame
itself; only covering it hides it. The menu bar is drawn by Office and no attribute reaches it.

A full menu tree dump (11 menus, ~90 items, captions + IDs + nesting + enabled state) was taken
2026-08-01; regenerate with a CommandBars walk when needed.

Pane lifecycle facts, measured while making tabs behave:
- With maximised panes the editor keeps a WINDOW only for the ACTIVE pane and destroys the
  others. The window map says one pane however many are open; only `CodePanes` knows the open
  set. Never feed the tab list from windows.
- `CodePanes` keeps a corpse entry for a just-closed pane: `Count` still counts it and
  `Item(n).CodeModule` throws. Read the collection per item, tolerantly, always.
- `CodePane.Show` displays but does not activate a pane that is already open behind another; no
  activation means no window event. Set `ActiveCodePane` directly - and that is an OBJECT
  assignment, which must go out as `DISPATCH_PROPERTYPUTREF`; an ordinary put is refused.
- Activating a native pane MOVES KEYBOARD FOCUS to it. Return focus to the surface
  (`controller.MoveFocus`) or every keyboard shortcut works exactly once.
- Ctrl+W and Ctrl+PageDown/PageUp are BROWSER accelerators (close window, switch browser tabs);
  unclaimed they never reach the page. Claim them at the accelerator hook like F1.
- The page has a `trace` message that writes into the shim log (`page: ...`), which is how a
  page-side silence is told apart from a transport one. `globalThis.xlideBridge` is reachable
  from devtools for the same reason.
- `tools`-side probes may use synthetic input against the harness (never in the product), BUT
  the user has judged SendKeys-style probing unreliable and it misfired once (it executed
  File -> Import through the menu). The standing verification loop is: the user drives, and the
  log answers - both halves log every evaluation, its outcome, and each change in what the
  hidden window holds, with non-printables escaped.

Immediate window facts, measured while making capture behave:
- The UIA TextPattern reading of the (hidden) Immediate window ends with a trailing empty line
  and a NUL, U+0000. New output is inserted BEFORE that tail, so the raw text never simply
  grows. Diff on readings trimmed of ALL trailing whitespace-or-control characters, by class,
  never by a list; a raw comparison first replayed the whole buffer per evaluation and then
  swallowed every print.
- A reading of empty is a project-reset artefact, not a cleared window (it is hidden; nobody
  can clear it): ignore it, never adopt it as the baseline. A reading that does not continue
  the baseline means trimming; adopt it silently.
- An unhandled run-time error in code started by Application.Run drops the editor into break
  mode INSIDE the calling frame and shows its dialog. The evaluator's scratch procedure carries
  its own On Error handler and returns the language's message as a Chr$(1)-marked value; the
  scratch module (ImmediateEvaluator.ScratchModule) is filtered out of the pane picker, the tab
  list, the explorer and the debugger's module switch, because it surfacing was one flash per
  evaluation.
- Evaluation must check VBProject.Mode FRESH, not the cached break flag: evaluating during an
  unnoticed break adds a module to a stopped project and fails incomprehensibly.
- A successful statement prints nothing (native parity); ?expr prints its value; errors print
  the editor's own message.
