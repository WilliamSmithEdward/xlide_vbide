<#
.SYNOPSIS
    A workbook whose macros are DISABLED, and whether the editor says so instead of going quiet.

.DESCRIPTION
    Excel puts itself into design mode for a workbook it will not run, and it will not come out:
    the Design Mode toggle stays pressed however many times it is pressed, every Reset control
    stays greyed, and the button answers "executed" each time. Nothing the product offers clears
    it - the way out is closing the workbook and opening it again with macros enabled.

    That is issue #9's shape: every debug command refused, no dialog standing, no form showing,
    Excel's own surface interactive, and only restarting Excel clears it. It is measured here so
    the product goes on saying WHICH design mode it is in, because the advice differs: one the
    developer pressed, which the toggle undoes, and one Excel is enforcing, which it will not.

    THE PAIR IS THE MEASUREMENT. The same one-module workbook is opened both ways. A macro that
    will not run means nothing on its own - it reads identically for a project that will not
    compile, which is how three earlier attempts at this measured nothing at all.

.EXAMPLE
    tools\harness\Test-MacrosDisabled.ps1
#>
[CmdletBinding()]
param(
    # Where the probe workbook is built. Nothing here touches a fixture.
    [string] $Folder = (Join-Path $env:TEMP 'xlide-macros-disabled'),

    [int] $TimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
$failures = New-Object System.Collections.Generic.List[string]

function Check([string] $what, [bool] $ok, [string] $detail) {
    if ($ok) { Write-Host "  ok   $what  -- $detail" }
    else { Write-Host "  FAIL $what  -- $detail"; $failures.Add($what) }
}

# One module, one Function returning a number, nothing that can fail to compile. A copy of a real
# fixture will not do: they carry deliberate compile errors, and VBA runs NOTHING while any module
# in the project has one - so "the macro would not run" would mean two things at once.
$source = @'
Option Explicit

Public Function Answer() As Long
    Answer = 42
End Function
'@

New-Item -ItemType Directory -Force $Folder | Out-Null
$book = Join-Path $Folder 'MacroProbe.xlsm'
Remove-Item $book -ErrorAction SilentlyContinue

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

Write-Host 'building the probe workbook'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $made = $excel.Workbooks.Add()
    $project = $made.VBProject
    if ($null -eq $project) {
        # NOT A FAILURE OF THE PRODUCT. Writing a module needs "Trust access to the VBA project
        # object model", which is the developer's machine setting and moves; nothing this product
        # SHIPS requires it. Said out loud rather than reported as a defect.
        Write-Host '  SKIPPED: "Trust access to the VBA project object model" is off, so this'
        Write-Host '           script cannot build its probe workbook. Turn it on to run this test.'
        $excel.Quit()
        exit 0
    }

    $module = $project.VBComponents.Add(1)
    $module.Name = 'Probe'
    $module.CodeModule.AddFromString($source)
    $made.SaveAs($book, 52)
    $made.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Start-Sleep -Seconds 2

# msoAutomationSecurityLow = 1, msoAutomationSecurityForceDisable = 3. Reached this way rather
# than through the mark of the web, because the mark sends the file to Protected View first and
# the only way out of that is a click on the banner.
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false

function Try-Run([string] $name) {
    try { return @{ Ran = $true; Detail = "returned $($excel.Run("'$name'!Probe.Answer"))" } }
    catch { return @{ Ran = $false; Detail = $_.Exception.Message.Split([char]10)[0].Trim() } }
}

try {
    Write-Host ''
    Write-Host 'the control: the same file with macros ENABLED'
    $excel.AutomationSecurity = 1
    $opened = $excel.Workbooks.Open($book)
    Start-Sleep -Seconds 2
    $control = Try-Run $opened.Name
    Check 'the probe macro runs when macros are enabled' $control.Ran $control.Detail
    $opened.Close($false)
    Start-Sleep -Seconds 1

    Write-Host ''
    Write-Host 'and now with macros DISABLED'
    $excel.AutomationSecurity = 3
    $opened = $excel.Workbooks.Open($book)
    Start-Sleep -Seconds 2
    $blocked = Try-Run $opened.Name
    Check 'and it does not when they are disabled' (-not $blocked.Ran) $blocked.Detail
    Check 'the VBA project is still loaded either way' ([bool] $opened.HasVBProject) 'HasVBProject'

    $excel.CommandBars.ExecuteMso('VisualBasic')
    $hostPid = (Get-Process EXCEL | Where-Object { $_.MainWindowTitle } | Select-Object -First 1).Id

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $door = $null
    while ($null -eq $door -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        $answer = & node (Join-Path $PSScriptRoot 'xlide-api.mjs') --pid $hostPid doctor 2>$null | Out-String
        if ($answer -match '"healthy"') { $door = $true }
    }
    Check 'the editor opens on a workbook it cannot run' ([bool] $door) "pid $hostPid"
    if (-not $door) { throw 'the door never came up' }

    $probe = Join-Path $PSScriptRoot 'macros-disabled-probe.mjs'
    $answer = & node $probe --pid $hostPid 2>&1 | Out-String
    Write-Host $answer.TrimEnd()
    if ($LASTEXITCODE -ne 0) { $failures.Add('the api probe reported failures') }
} finally {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}

Write-Host ''
if ($failures.Count -gt 0) {
    Write-Host "FAIL $($failures.Count): $($failures -join '; ')"
    exit 1
}
Write-Host 'PASS'
exit 0
