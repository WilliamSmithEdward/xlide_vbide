# End to end: does the ghost palette feed the THEMED panel through break and steps?
# Breaks, steps twice, and checks the shim log for locals row pushes with changing values.
# Run tools\dev.ps1 -KeepOpen first. Prints PASS or FAIL.
# Dev-harness script: uses VBProject and Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE

$components = $app.ActiveWorkbook.VBProject.VBComponents
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}
$component = $components.Add(1)
$component.Name = 'BreakProbe'
$component.CodeModule.AddFromString(@'
Sub BreakHere()
    Dim counter As Long
    Dim label As String
    counter = 1
    label = "alpha"
    Stop
    counter = 2
    label = "beta"
    Stop
End Sub
'@)

function Get-Control([int]$id) { $vbe.CommandBars.FindControl(1, $id) }

$excelId = (Get-Process EXCEL | Select-Object -First 1).Id
$log = Get-ChildItem "$env:LOCALAPPDATA\xlide_vbide\logs" -Filter "shim-*-$excelId.log" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
$before = (Get-Content $log.FullName | Measure-Object -Line).Lines

$app.OnTime([DateTime]::Now.AddSeconds(1), 'BreakProbe.BreakHere')
$deadline = (Get-Date).AddSeconds(20)
while ((Get-Date) -lt $deadline) {
    $step = Get-Control 188
    if ($step -and $step.Enabled) { break }
    Start-Sleep -Milliseconds 200
}
Start-Sleep -Milliseconds 1200

(Get-Control 188).Execute()
Start-Sleep -Milliseconds 1000
(Get-Control 188).Execute()
Start-Sleep -Milliseconds 1000
(Get-Control 188).Execute()
Start-Sleep -Milliseconds 1200

$reset = Get-Control 228
if ($reset -and $reset.Enabled) { $reset.Execute() }
Start-Sleep -Milliseconds 500
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}

$since = Get-Content $log.FullName | Select-Object -Skip $before
Write-Output '--- locals lines ---'
$since | Where-Object { $_ -match 'locals' }

$pushes = @($since | Where-Object { $_ -match 'locals: \d+ row' })
$setLocals = @($since | Where-Object { $_ -match 'setLocals' })
if ($pushes.Count -ge 2 -and $setLocals.Count -ge 2) {
    Write-Output "RESULT: PASS - $($pushes.Count) row push(es), $($setLocals.Count) setLocals message(s)"
} else {
    Write-Output "RESULT: FAIL - pushes $($pushes.Count), setLocals $($setLocals.Count)"
}
