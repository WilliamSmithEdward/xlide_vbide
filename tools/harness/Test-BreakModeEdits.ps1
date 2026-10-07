# Focused live regression for issue #71. Run with one disposable Debug Excel session
# holding scratch.xlsm and trusted access to its VBA project.
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$excel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($excel.Count -ne 1) { throw 'Run against exactly one disposable Excel session.' }
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$book = $app.ActiveWorkbook
if ($book.Name -ne 'scratch.xlsm') { throw "Expected scratch.xlsm, found $($book.Name)." }
$components = $book.VBProject.VBComponents
$name = 'Issue71Probe'

try {
    foreach ($existing in @($components)) {
        if ($existing.Name -eq $name) { $components.Remove($existing) }
    }
    $module = $components.Add(1)
    $module.Name = $name
    $module.CodeModule.AddFromString(@'
Public Sub DebugStopTest()
    Debug.Print 1
    Stop
    Debug.Print 2
End Sub
'@)
    $app.OnTime([datetime]::Now.AddSeconds(3), 'Issue71Probe.DebugStopTest')
    & node (Join-Path $PSScriptRoot 'break-mode-edits.mjs') $excel[0].Id safe
    if ($LASTEXITCODE -ne 0) { throw 'The live statement edit check failed.' }
    $app.OnTime([datetime]::Now.AddSeconds(3), 'Issue71Probe.DebugStopTest')
    & node (Join-Path $PSScriptRoot 'break-mode-edits.mjs') $excel[0].Id reset
    if ($LASTEXITCODE -ne 0) { throw 'The live reset-required edit check failed.' }
}
finally {
    $reset = $app.VBE.CommandBars.FindControl(1, 228)
    if ($reset -and $reset.Enabled) { $reset.Execute() }
    foreach ($existing in @($components)) {
        if ($existing.Name -eq $name) { $components.Remove($existing) }
    }
}
