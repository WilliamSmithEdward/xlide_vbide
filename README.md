# xlide

[![Latest release](https://img.shields.io/github/v/release/WilliamSmithEdward/xlide_vbide)](https://github.com/WilliamSmithEdward/xlide_vbide/releases/latest)
[![MIT license](https://img.shields.io/github/license/WilliamSmithEdward/ROneCOne)](LICENSE)
[![Microsoft 365 Excel on Windows x64](https://img.shields.io/badge/Excel-Microsoft_365_Windows_x64-217346)](README.md)

The Visual Basic Editor has looked the same since 1998. One module on screen at a time, a find
dialog that covers the code you are searching, no completion worth the name, and no idea that
anything is wrong until you press F5 and it stops. xlide replaces that surface with Monaco, the
editor behind VS Code, running inside the VBE itself. Your workbooks stay where they are, your
macros run the way they always did, and F5, F8, and break mode behave exactly as before.

It installs in one double click, for your account only, with no administrator rights and nothing to
install first.

## What it is

xlide is a native add-in that Excel's editor loads through its own extensibility model. It draws
over the VBE's document area and puts a modern editing surface there: every open module live at
once, editors side by side, diagnostics as you type. The native editor keeps running underneath as
the text of record, the compile target, and the debugger, so nothing about how your code compiles or
runs changes.

The analysis comes from a VBA analyzer validated against the real compiler over a corpus of
thousands of modules, and it runs in its own process. That matters because the VBE is single
threaded and owns the thread you type on: a project large enough to take seconds to analyse cannot
stall your typing if the analysis is not happening there. The add-in is compiled ahead of time to
native code, so Excel never loads a .NET runtime on its account.

## What you get

- Monaco editing over every code pane, with VBA syntax highlighting, folding, multi-cursor, and
  bookmarks that stay with the module.
- Diagnostics as you type, with severity filters and a Problems panel that navigates to the line.
- Completion, hover, and signature help from the analyzer, plus typing that follows VBA's own
  conventions: auto-casing, block closers, smart Enter and Tab, and auto-indent.
- Every open module holds a live editor. Switching tabs takes a median of 2.5ms and keeps your undo
  history, scroll position, and squiggles.
- Editor groups. Split right or down with Ctrl+\, drag a tab to a group's edge, and work in two
  modules at once.
- Six tool panes that dock where you put them: Explorer, Properties, Problems, Immediate, Locals,
  and Watch. Drag one by its title and a five-zone compass appears over the region under the
  pointer. The arrangement persists.
- Search as one floating widget, scoped to the module, the workbook, or every open workbook. Find
  All lists every match with a preview, and Replace All applies as a single edit that one undo
  reverts.
- Locals and Watch track the debugger through every step, with breakpoints, Run to Cursor, Set Next
  Statement, and the rest of the debug commands on the toolbar.
- An Object Browser as a floating themed window, covering your open projects and every referenced
  type library, read directly from the type libraries themselves. Members of your own modules jump
  to their definition on a double click.
- A project explorer rooted at each open workbook, with project-qualified addressing, so two
  workbooks can both hold a Module1 and xlide knows which one you mean.
- An Immediate panel that mirrors the native window live, and a close-confirm on a dirty tab that
  reverts everywhere when you choose not to save.

## Installing

Download `xlide-setup.exe` from [Releases](https://github.com/WilliamSmithEdward/xlide_vbide/releases)
and run it.

It installs to `%LOCALAPPDATA%\Programs\xlide` for the current user. It asks for no administrator
rights, changes nothing outside your own profile, and needs no runtime, framework, or tool to be
present first. Everything it needs is inside the one executable. Windows will warn before running
it, because it carries no code signature yet.

Close Excel first. If it is open, the installer says so and offers to wait while you close it, or to
force close it for you.

Then start Excel and press Alt+F11.

## Uninstalling

Settings, then Apps, then Installed apps. Find xlide and choose Uninstall. You can also run
`xlide-setup.exe --uninstall` from `%LOCALAPPDATA%\Programs\xlide`, where a copy of the installer is
kept so that removing xlide never depends on still having the download.

Removal takes out the program files, the per-user registration, and xlide's own logs and cache.
Your VBA is untouched throughout: it lives in your workbooks and xlide never writes to them.

One thing to expect afterwards. xlide hides the VBE's own tool windows while it is covering the
screen, and the editor remembers the window layout it was last left with. So the first time you open
the VBE after removing xlide, it will be empty. Use the View menu to bring back the ones you want:
Project Explorer, Properties Window, Immediate Window, Locals Window, and Watch Window.

---

# For developers

Everything below this line is about working on xlide. None of it is needed to use it. If you
installed from the release, you are done: start Excel and press Alt+F11.

## Building from source

You need the .NET 10 SDK, the C++ build tools that ahead-of-time compilation links with, and Node
for the language engine. Excel is needed only for the integration check.

```powershell
tools\dev.ps1            # build, test, register, and verify inside a real editor
tools\verify.ps1         # the whole local gate, about twenty seconds
installer\build.ps1      # produce the installer
```

The development loop finishes in about a second and a half. `tools\page.ps1` rebuilds and reloads
the editor surface in about a second without restarting Excel.

## Repository layout

| Path | Purpose |
| --- | --- |
| `src/Xlide.Vbe.Shim` | The native add-in: editor integration, tool windows, browser surface |
| `src/Xlide.Vbe.Core` | Host-independent logic for the add-in, with no COM or Win32 |
| `engine/` | The language engine sidecar |
| `ui/` | The editor surface rendered in the browser control |
| `installer/` | The single-file installer |
| `tools/` | Development scripts and the integration harness |
| `tests/` | Unit tests. None of them need Excel |
| `docs/` | Architecture, decisions, and findings |

## Documentation

[docs/status.md](docs/status.md) is the current snapshot: what is proven and how.
[docs/architecture.md](docs/architecture.md) covers the design, and
[docs/decisions.md](docs/decisions.md) records the choices that would be expensive to reverse along
with the reasoning behind each. [docs/lessons.md](docs/lessons.md),
[docs/ui-lessons.md](docs/ui-lessons.md), and [docs/editor-windows.md](docs/editor-windows.md) hold
behaviour of the host that is documented nowhere else and was established by measurement.
The newest handover is written for someone starting cold: the highest-dated `docs/handoff-*.md`. They are dated because each supersedes the last, so the most recent one wins and the others are history.

## About the code

This is a clean-room implementation built on Microsoft's documented interfaces: the editor
extensibility model, the forms designer object model, Win32, and published binary format
specifications. The analyzer is the author's own prior work, shared with the
[XLIDE editor extension](https://github.com/WilliamSmithEdward/xlide_vscode) so that both products
agree on what VBA means.

## License

[MIT](LICENSE).
