# Closing a HIDDEN pane must still remove its tab from the strip.
#
# Guards the 2026-08-04 fix (the tab X's third and real mechanism): the pane tracker only
# holds the pane windows it can match - the active one, in practice - so a hidden pane's
# close changes nothing in its picture, Changed never fires, and the strip kept a dead tab.
# Any window destroy now arms a moment of polls that re-read the object model's open list
# and republish; this probe closes the hidden pane exactly as the host's close path does
# and expects a setModules WITHOUT the closed module within a second. Prints PASS or FAIL.
#
# Run tools\dev.ps1 -KeepOpen first, then this.
# Dev-harness script: uses Application.VBE per decision 10's harness exception.
#
# NOT IN THE GATE, and on a trust-off machine it cannot be: the host-originated close of a
# HIDDEN pane exists only as a window operation on a CodePane reached through the VBA object
# model, which "Trust access to the VBA project object model" gates. The gate's machine runs
# with that trust OFF (so does everything api-driven, by design). The way this suite joins a
# gate is a pane?action=closeNative route on the debug api - audit item A12 - which would make
# the host-originated close drivable with no trust at all.
$ErrorActionPreference = 'Continue'

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE
if ($null -eq $vbe) {
    Write-Output ('RESULT: FAIL - Application.VBE answered null, which is what "Trust access to '
        + 'the VBA project object model" being OFF looks like. This probe needs it ON; nothing '
        + 'api-driven does.')
    exit 1
}

# Both modules open as panes, CleanModule active (BrokenModule's pane hidden behind it).
foreach ($name in @('BrokenModule', 'CleanModule')) {
    foreach ($project in $vbe.VBProjects) {
        foreach ($component in $project.VBComponents) {
            if ($component.Name -eq $name) { $component.CodeModule.CodePane.Show() }
        }
    }
}
Start-Sleep -Milliseconds 1200

$excelId = (Get-Process EXCEL | Select-Object -First 1).Id
$log = Get-ChildItem "$env:LOCALAPPDATA\xlide_vbide\logs" -Filter "shim-*-$excelId.log" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
$before = (Get-Content $log.FullName | Measure-Object -Line).Lines
Write-Output "log at line $before; closing the hidden BrokenModule pane now"

# Close the HIDDEN pane exactly as the host's close path does: through its window.
foreach ($pane in $vbe.CodePanes) {
    $module = $pane.CodeModule
    if ($module.Parent.Name -eq 'BrokenModule') {
        $pane.Window.Close()
        Write-Output 'closed BrokenModule pane window'
    }
}

Start-Sleep -Seconds 4

$since = Get-Content $log.FullName | Select-Object -Skip $before
$since | Where-Object { $_ -match 'code panes|setModules|follow:|abandoned|closeModule|closed' }

$republished = $since | Where-Object { $_ -match 'setModules' -and $_ -notmatch 'BrokenModule' }
if ($republished) {
    Write-Output 'RESULT: PASS - the strip was republished without the closed module'
} else {
    Write-Output 'RESULT: FAIL - no republish; the strip would keep showing a dead tab'
}
