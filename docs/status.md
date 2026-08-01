# Build status

Updated 2026-08-01.

## What is proven

The add-in loads and runs inside the Excel Visual Basic Editor.

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
| Native COM server, ahead-of-time compiled | Working. 1.91 MB, exports `DllGetClassObject` and `DllCanUnloadNow`, no runtime deployed or loaded |
| Add-in lifetime contract | Working. Connect, startup, shutdown, and disconnect handled, with teardown ordered before host shutdown |
| Editor object model access | Working, through a late-binding wrapper with single-ownership release |
| Registration | Working, per user, no administrator rights. One source of truth, unit tested |
| Development loop | Working. `tools/dev.ps1` builds, tests, publishes, registers, and verifies in a real editor |
| Docked tool window and browser surface | In progress |
| Language engine sidecar | In progress |
| Installer | In progress |
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
