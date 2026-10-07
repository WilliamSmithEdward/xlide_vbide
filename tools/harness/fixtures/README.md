# Locked VBA project fixture

`LockedProjectFixture.xlsm` is a checked-in workbook for manually testing [issue #66](https://github.com/WilliamSmithEdward/xlide_vbide/issues/66). Its VBA project is locked for viewing with the **test-only password `Test66`**. The worksheet also shows the password and reset instruction. Use an XLIDE build containing the issue #66 fix (main from `e5bf24c` onward).

Open the workbook in Excel, then open XLIDE. Or start a fresh Excel harness process from the repository root:

```powershell
tools\harness\Start-Excel.ps1 -Fresh -Workbook tools\harness\fixtures\LockedProjectFixture.xlsm
```

XLIDE Explorer should show `LockedProjectFixture.xlsm` with a lock and no child modules. Activate that row; XLIDE should remain visible while the native `VBAProject Password` dialog opens. Enter `Test66`. XLIDE should then show `ThisWorkbook`, `Sheet1`, `Runner`, and `Helper`. Canceling the password prompt should leave the locked row in XLIDE.

**Close the workbook without saving after an unlock.** Reopening the checked-in file then restores the locked state for the next test. If it was saved while unlocked, restore this fixture from Git before testing again.
