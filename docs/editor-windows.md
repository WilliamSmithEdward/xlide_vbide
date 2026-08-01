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
