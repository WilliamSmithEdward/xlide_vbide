# Build status

Updated 2026-08-01, second revision.

## What is proven

The add-in loads inside the Excel Visual Basic Editor, the editor sites our control in a docked
tool window, a WebView2 browser surface renders inside that window, and the whole product installs
and uninstalls from a single 3.92 MB executable.

Verified installer round trip, all observed on this machine: install exits 0 and lands four files
under `%LOCALAPPDATA%\Programs\xlide`; registration points at the installed copy; the installed
copy connects and renders in a real editor; uninstall removes the install folder, the data folder,
and every registry key, verified by enumeration afterwards.

Verified load path, 1.35 seconds end to end with every phase measured: start host 0.25s, attach
0.68s, open editor 0.15s, add-in load 0.01s, browser surface 0.23s, confirm connection 0.03s. The
check passes only when three independent conditions are all observed: the add-in's own log shows a
connection, the editor reports the add-in connected, and the browser surface finished navigating.

Known cost of shipping unsigned during development: Windows Security warns that the installer's
publisher cannot be confirmed. Signing is a release requirement, recorded in decisions.md.

## The original proof

Verified end to end by `tools/harness/Invoke-VbeLoadCheck.ps1` against Excel 365 x64
(16.0.20228.20124, VBA 7.1) on Windows 11:

```text
Editor add-in list: Xlide.VbeAddIn connect=True

[info] log opened, pid 24592, x64, host C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE
[info] DllGetClassObject for {588903f2-4cde-4607-828a-6870a1f3fdc1}
[info] CreateInstance for {00000000-0000-0000-c000-000000000046} returned 0x00000000
[info] OnConnection, mode Startup
[info] editor version 7.01
[info] main window caption 'Microsoft Visual Basic for Applications'
[info] OnStartupComplete
[info] 1 project(s) loaded
[info]   project 'VBAProject' with 2 component(s)
```

That output establishes several things at once: the per-user registration is found, a native
library with no runtime behind it is loaded into the host, the class is activated, interface
negotiation succeeds, the extensibility contract is honoured, and the editor object model is
readable from inside the host process.

## Component state

| Component | State |
| --- | --- |
| Native COM server, ahead-of-time compiled | Working. 2.12 MB, exports `DllGetClassObject` and `DllCanUnloadNow`, no runtime deployed or loaded |
| Add-in lifetime contract | Working. Connect, startup, shutdown, and disconnect handled, with teardown ordered before host shutdown |
| Editor object model access | Working, through a late-binding wrapper with single-ownership release |
| Registration | Working, per user, no administrator rights. One source of truth, unit tested |
| Docked tool window | Working. The editor sites our control through the documented tool window mechanism |
| Browser surface | Working. WebView2 hosted in the tool window, navigated, observed on screen |
| Content root and message bridge | Working. Serves a folder under a virtual host name and exchanges messages with the page, falling back to the built-in document when no bundle is present |
| Code pane tracking | Working. Panes are identified by component, located, and followed as the editor is rearranged |
| Installer | Working. One 3.92 MB executable, install and clean uninstall verified by round trip |
| Development loop | Working. Full check in 1.35s; instance reuse brings a repeat check to 0.24s |
| Language engine sidecar | In progress |
| Editor surface, designer, debugging | Designed, not started |

## How to run it

```powershell
tools\dev.ps1
```

Builds, runs the unit tests, publishes the native shim, registers it for the current user, then
starts Excel, opens the editor, and prints the add-in's log. Add `-KeepOpen` to keep Excel running.
`tools\dev.ps1 -Unregister` removes the registration.

## Machine requirements for development

- .NET 10 SDK.
- The C++ build tools, for the native linker used by ahead-of-time publishing. The linker is located
  through `vswhere.exe`, which must be on `PATH`; `tools/dev.ps1` adds it.
- Excel, for the integration check only. Unit tests do not need it.

## Next

1. Site the tool window host in the editor and bring up the browser surface inside it.
2. Connect the engine and show real diagnostics for the active module.
3. Position the editor surface over a code pane and converge it with the code module.
