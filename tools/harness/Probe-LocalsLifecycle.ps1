# Does the Locals window keep tracking the debugger across hide/show cycles?
# One fresh instance, three breaks: visible (first ever), hidden, re-shown. What each read
# returns decides the Locals panel design: read-hidden like Immediate, or keep-visible-covered.
# Dev-harness script: uses VBProject and Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Set-Location F:\GitHub\xlide\xlide_vbide
& tools\harness\Get-EditorScreenshot.ps1 -KeepOpen | Out-Null

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE
$components = $app.ActiveWorkbook.VBProject.VBComponents
$component = $components.Add(1)
$component.Name = 'BreakProbe'
$component.CodeModule.AddFromString(@'
Sub BreakOne()
    Dim label As String
    label = "one"
    Stop
End Sub
Sub BreakTwo()
    Dim label As String
    label = "two"
    Stop
End Sub
Sub BreakThree()
    Dim label As String
    label = "three"
    Stop
End Sub
'@)

function Get-Control([int]$id) { $vbe.CommandBars.FindControl(1, $id) }

function Set-Locals([bool]$visible) {
    foreach ($window in $vbe.Windows) {
        if ($window.Type -eq 4) { $window.Visible = $visible }
    }
}

$script:handle = [IntPtr]::Zero
function Find-LocalsHandle {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty, 'wndclass_desked_gsk')
    $frame = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    foreach ($element in $frame.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)) {
        $c = $element.Current
        if ($c.ClassName -eq 'VbaWindow' -and $c.Name -like 'Locals*') {
            $script:handle = [IntPtr]$c.NativeWindowHandle
            return
        }
    }
}

function Read-Locals([string]$title) {
    if ($script:handle -eq [IntPtr]::Zero) { Write-Output "  $title -- NO HANDLE"; return }
    $window = [System.Windows.Automation.AutomationElement]::FromHandle($script:handle)
    $lines = @()
    foreach ($child in $window.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)) {
        $c = $child.Current
        $kind = $c.ControlType.ProgrammaticName
        if ($kind -eq 'ControlType.ListItem' -or $kind -eq 'ControlType.Edit') {
            $lines += "[$kind] $($c.Name)"
        }
    }
    Write-Output "=== $title"
    $lines | ForEach-Object { "  $_" }
}

function Break-And([string]$procedure) {
    $app.OnTime([DateTime]::Now.AddSeconds(1), "BreakProbe.$procedure")
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        $step = Get-Control 188
        if ($step -and $step.Enabled) { break }
        Start-Sleep -Milliseconds 200
    }
    Start-Sleep -Milliseconds 700
}

function Reset-Debugger {
    $reset = Get-Control 228
    if ($reset -and $reset.Enabled) { $reset.Execute() }
    Start-Sleep -Milliseconds 700
}

# 1: first ever show, then break. The banked success path.
Set-Locals $true
Start-Sleep -Milliseconds 500
Find-LocalsHandle
Write-Output "locals hwnd: $script:handle"
Break-And 'BreakOne'
Read-Locals '1: visible first time, expect label=one'
Reset-Debugger

# 2: hidden before the break. Does it track unseen?
Set-Locals $false
Start-Sleep -Milliseconds 500
Break-And 'BreakTwo'
Read-Locals '2: hidden during break, expect label=two if it tracks hidden'
Reset-Debugger

# 3: shown again before the break. Does re-show reattach?
Set-Locals $true
Start-Sleep -Milliseconds 500
Break-And 'BreakThree'
Read-Locals '3: re-shown, expect label=three if re-show reattaches'
Reset-Debugger

# 4: still shown, but toggle DURING the break: hide+show while broken.
Break-And 'BreakOne'
Set-Locals $false
Start-Sleep -Milliseconds 300
Set-Locals $true
Start-Sleep -Milliseconds 700
Read-Locals '4: toggled while broken, expect label=one if mid-break reshow refreshes'
Reset-Debugger

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-Process xlide-engine -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Write-Output 'done'
