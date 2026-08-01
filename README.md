# xlide_vbide

A modern development experience for VBA, directly inside the Excel Visual Basic Editor.

xlide_vbide is a VBE add-in that upgrades the editor you already use: a modern code editing
surface, live diagnostics with quick fixes, rich completion and navigation, refactoring, and
test tooling, all rendered inside the VBE itself. It installs from a single executable, requires
no other runtimes or tools on the machine, and installs for the current user without
administrator rights.

## Status

Early development. See [docs/status.md](docs/status.md) for the current build state and
[docs/architecture.md](docs/architecture.md) for the system design.

## How it works, in one paragraph

The add-in is a COM in-process server that the VBE loads through its documented extensibility
model (`IDTExtensibility2`, per-user registration under
`HKCU\Software\Microsoft\VBA\VBE\6.0\Addins64`). Inside the VBE it hosts modern web-based UI
(WebView2) in native docked tool windows, overlays a Monaco-based editor surface on the code
panes, and talks to an out-of-process language engine over a named pipe. The engine is the
XLIDE VBA analyzer: an error-tolerant lexer/parser, project-wide symbol index, type inference,
and a diagnostics rule set validated against the real VBE compiler, shared with the
[XLIDE VS Code extension](https://github.com/WilliamSmithEdward/xlide_vscode).

## Repository layout

| Path | Purpose |
| --- | --- |
| `src/` | .NET add-in: VBE integration, tool windows, editor surface host |
| `engine/` | Language engine sidecar: server wrapper and build of the XLIDE analyzer |
| `ui/` | Web UI bundles rendered in WebView2 (editor surface, panels) |
| `installer/` | The single-file installer |
| `tools/` | Development scripts and the Excel integration harness |
| `tests/` | Unit and integration tests |
| `docs/` | Architecture, decisions, developer guide |

## License

[MIT](LICENSE).
