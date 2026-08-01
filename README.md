# xlide_vbide

A modern development experience for VBA, inside the editor you already use.

xlide_vbide is an add-in for the Visual Basic Editor. It aims at a modern code editing surface, live
diagnostics with quick fixes, real completion and navigation, refactoring, and test tooling, all
rendered inside the editor itself rather than in a separate application. It installs from a single
executable, needs no other runtime or tool on the machine, and installs for the current user without
administrator rights.

## Status

Early, and honest about it. The foundation is built and verified against a real host; the features
that make it worth installing are not there yet. See [docs/status.md](docs/status.md) for what is
proven and [docs/architecture.md](docs/architecture.md) for the design.

Working today:

- The add-in loads into the editor as a native in-process server, with no managed runtime loaded
  into the host.
- The editor docks a tool window on our own control, and a browser surface renders inside it.
- Code panes are located and followed as the editor is rearranged.
- A language engine answers real analysis out of process, over a named pipe.
- The whole product installs and uninstalls cleanly from one executable.

Not built yet: the editor surface over code panes, the UserForm designer, debugging integration, and
the panels that surface any of the analysis to the user.

## How it works

The add-in is a COM in-process server that the editor loads through its documented extensibility
model, registered per user. It is compiled ahead of time to a native library, so no runtime is
deployed with it or loaded into the host process.

Inside the editor it hosts web-based UI in native docked tool windows, and it talks to a language
engine running in its own process. Keeping analysis outside the host is deliberate: the editor is
single threaded and owns the thread the user types on, and a separate process cannot block it, grow
its memory, or crash it.

The analyzer is the one from the
[XLIDE editor extension](https://github.com/WilliamSmithEdward/xlide_vscode), validated against the
real VBA compiler and shared between both products. It is being ported to C# layer by layer, gated
by differential testing against that implementation, to remove the language runtime the current
engine has to ship.

## Building

Needs the .NET 10 SDK, the C++ build tools (for the native linker that ahead-of-time compilation
uses), and Node for the current engine. Excel is needed only for the integration check.

```powershell
tools\dev.ps1            # build, test, register, and verify inside a real editor
tools\dev.ps1 -Unregister
installer\build.ps1      # produce the installer
```

The development loop completes in about one and a half seconds, and a repeat check that reuses the
open host takes about a quarter of a second.

## Repository layout

| Path | Purpose |
| --- | --- |
| `src/Xlide.Vbe.Shim` | The native add-in: editor integration, tool windows, browser surface |
| `src/Xlide.Vbe.Core` | Host-independent logic for the add-in, with no COM or Win32 |
| `src/Xlide.Vba.Analysis` | The analyzer being ported to C# |
| `engine/` | The language engine sidecar in its current form |
| `ui/` | Web UI rendered in the browser surface |
| `installer/` | The single-file installer |
| `tools/` | Development scripts and the integration harness |
| `tests/` | Unit tests. None of them need Excel |
| `docs/` | Architecture, decisions, and findings |

## Picking this up

[docs/handoff.md](docs/handoff.md) is written for someone starting cold: machine setup, how to run
everything, the behaviour of the host that cost real time to discover, what is open, and what to do
next in order.

## Notes on the design

[docs/decisions.md](docs/decisions.md) records the choices that would be expensive to reverse, with
the reasoning behind each. [docs/lessons.md](docs/lessons.md) and
[docs/editor-windows.md](docs/editor-windows.md) record behaviour of the host that is not documented
anywhere and was established by measurement.

This is a clean-room implementation built on Microsoft's documented interfaces: the editor
extensibility model, the forms designer object model, Win32, and published binary format
specifications. Its analyzer is the author's own prior work.

## License

[MIT](LICENSE).
