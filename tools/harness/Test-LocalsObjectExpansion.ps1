# Focused live regression for issue #70. Requires a single Debug dev Excel session with the
# disposable scratch.xlsm workbook and trusted VBProject access.
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$excel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($excel.Count -ne 1) { throw 'Run against exactly one disposable Excel session.' }
Import-Module (Join-Path $PSScriptRoot 'XlideApi.psm1') -Force
$api = Get-XlideApi -ProcessId $excel[0].Id
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$book = $app.ActiveWorkbook
if ($book.Name -ne 'scratch.xlsm') { throw "Expected disposable scratch.xlsm, found $($book.Name)." }
$vbe = $app.VBE
$components = $book.VBProject.VBComponents
$moduleName = 'Issue70Probe'
$className = 'Issue70Class'
$interfaceName = 'iIssue70'

try {
    $interface = $components.Add(2)
    $interface.Name = $interfaceName
    $interface.CodeModule.AddFromString(@'
Option Explicit
Public Property Get Status() As String
End Property
'@)
    $class = $components.Add(2)
    $class.Name = $className
    $class.CodeModule.AddFromString(@'
Option Explicit
Implements iIssue70
Public State As String
Private currentCount As Long
Public Property Get Count() As Long
    Count = currentCount
End Property
Public Property Let Count(ByVal nextCount As Long)
    currentCount = nextCount
End Property
Private Property Get iIssue70_Status() As String
    iIssue70_Status = State
End Property
'@)
    $module = $components.Add(1)
    $module.Name = $moduleName
    $module.CodeModule.AddFromString(@'
Option Explicit
Public Sub InspectObject()
    Dim sample As Collection
    Dim inner As Collection
    Dim model As Issue70Class
    Dim iface As iIssue70
    Set sample = New Collection
    Set inner = New Collection
    inner.Add "nested value"
    sample.Add inner, "alpha"
    sample.Add "second", "beta"
    Set model = New Issue70Class
    model.State = "ready"
    model.Count = 7
    Set iface = model
    Stop
    sample.Add "after stop"
End Sub
'@)
    $app.OnTime([datetime]::Now.AddSeconds(3), 'Issue70Probe.InspectObject')
    $deadline = [datetime]::Now.AddSeconds(30)
    while ([datetime]::Now -lt $deadline) {
        if ((Invoke-RestMethod "$($api.Base)/state" -TimeoutSec 2).debugMode -eq 'break') { break }
        Start-Sleep -Milliseconds 200
    }
    if ((Invoke-RestMethod "$($api.Base)/state" -TimeoutSec 2).debugMode -ne 'break') {
        throw 'The disposable object procedure did not stop in the debugger.'
    }
    & node (Join-Path $PSScriptRoot 'locals-object-expansion.mjs') $excel[0].Id
    if ($LASTEXITCODE -ne 0) { throw 'The live Locals panel check failed.' }
}
finally {
    $reset = $vbe.CommandBars.FindControl(1, 228)
    if ($reset -and $reset.Enabled) { $reset.Execute() }
    foreach ($existing in @($components)) {
        if ($existing.Name -in @($moduleName, $className, $interfaceName)) { $components.Remove($existing) }
    }
}
