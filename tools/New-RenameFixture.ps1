<#
.SYNOPSIS
    Builds a workbook that exercises every case rename has to get right.

.DESCRIPTION
    Rename is the one feature that CHANGES code rather than describing it, so the cases that
    matter are the ones where it must do nothing: a name inside a string, a module that merely
    starts with the same word, a procedure of the same name in another module. A fixture that
    only contains things which should change proves nothing.

    Every module below is labelled in its own comments with what to try and what must NOT move.
    The checklist is in the workbook, not only here, so it is still there next time.

    Built through the debug api, so "Trust access to the VBA project object model" does NOT have
    to be on. That setting gates VBComponents.Add through Workbook.VBProject; the add-in is
    already past it, because the host hands it the VBE. A Debug build must be registered and
    loading, which is what makes the door exist.

    The result deliberately does NOT compile: Helpers and Rival each declare a public
    Recalculate, and Consumer calls it bare. That collision is the point of the fixture, so never
    compile it as part of a wider experiment - every Run against it will fail for a reason that
    has nothing to do with what you asked.

.EXAMPLE
    tools\New-RenameFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-RenameFixture.ps1 -Path C:\temp\rename.xlsm -Quiet
    Builds it somewhere specific and closes Excel afterwards.
#>
[CmdletBinding()]
param(
    # Where to save it. Defaults beside the repo's other build output.
    [string] $Path,

    # Build and close, rather than leaving it open to work in.
    [switch] $Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Path) {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $fixtures = Join-Path $repoRoot 'artifacts\fixtures'
    if (-not (Test-Path $fixtures)) { New-Item -ItemType Directory -Force $fixtures | Out-Null }
    $Path = Join-Path $fixtures 'RenameFixture.xlsm'
}

if (Test-Path $Path) { Remove-Item $Path -Force }

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

$modules['Helpers'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' A STANDARD module. Rename it from the explorer, or from the qualifier in Consumer.
' Everything that reaches into it by name must follow.

Public Sub Recalculate(ByVal label As String)
    Debug.Print label
End Sub

Public Sub Nearby()
    ' A bare call INSIDE the declaring module. Renaming Recalculate must move this:
    ' module-local scope settles which one is meant.
    Recalculate "same module"
End Sub
'@ }

$modules['HelpersExtra'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' Its name STARTS WITH Helpers. Renaming Helpers must not touch this module or any
' mention of it, which is the difference between matching a name and matching a prefix.

Public Sub Thing()
End Sub
'@ }

$modules['Rival'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' Declares its OWN Recalculate. Renaming Helpers.Recalculate must leave every line of
' this module alone: same name, different procedure.

Public Sub Recalculate(ByVal label As String)
    Debug.Print "rival " & label
End Sub

Public Sub AlsoNearby()
    ' Bare, and inside the module that declares it, so it belongs to Rival's Recalculate.
    Recalculate "rival's own"
End Sub
'@ }

$modules['IShape'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' An INTERFACE. Renaming it has to reach three different shapes at once:
'   Implements IShape        in RoundShape
'   Private Sub IShape_Draw  in RoundShape  <- the prefix is the contract, not a convention
'   Dim s As IShape          in Consumer
' Leave the prefix behind and RoundShape stops implementing anything, and VBA says so.

Public Sub Draw()
End Sub
'@ }

$modules['RoundShape'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

Implements IShape

Private Sub IShape_Draw()
    Debug.Print "round shape"
End Sub
'@ }

$modules['Widget'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' An ORDINARY class. Renaming it must follow As Widget and New Widget in Consumer.

Public Sub Spin()
End Sub
'@ }

$modules['Consumer'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' The module that names everything else. LEAVE THIS ONE CLOSED for the test that matters:
' a rename must reach a module with no tab open, because that is the module a developer
' will never notice was missed.

Public Sub Drive()
    ' A qualifier. Right-click Helpers here and rename it: the module rename should run
    ' from code, not only from the explorer.
    Helpers.Recalculate "qualified"

    ' Bare, from a module that declares neither. Renaming Helpers.Recalculate must leave
    ' this alone, because two modules declare that name and nothing can prove which.
    Recalculate "ambiguous"

    ' A different module whose name merely begins the same way.
    HelpersExtra.Thing

    Dim w As Widget
    Set w = New Widget
    w.Spin

    Dim s As IShape
    Set s = New RoundShape
    s.Draw

    ' A module name inside a STRING is data. Nothing here may change.
    Debug.Print "Helpers and Widget and IShape are just words here"

    ' And in a comment: Helpers, Widget, IShape.
End Sub
'@ }

$modules['Watcher'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' LEAVE THIS ONE OPEN. A rename has to update the tab you are looking at as well as the
' modules you are not, and the two go through different paths to get there.

Public Sub Watch()
    Helpers.Recalculate "watching"
    Dim w As Widget
End Sub
'@ }

$sheetCode = @'
Option Explicit

' A DOCUMENT module. Renaming one of these is the case xlide has NOT been tested on: a
' sheet carries a code name and a sheet name and they are not the same thing. Try it and
' see, but expect to find something rather than expect it to work.

Public Sub Refresh()
    Debug.Print "sheet"
End Sub
'@

# ---------------------------------------------------------------- building it

# Built through the DOOR, not through Workbook.VBProject.
#
# The old shape needed "Trust access to the VBA project object model" turned on, because
# VBComponents.Add is exactly what that setting gates. The add-in is already past that gate - the
# host hands it the VBE at OnConnection - so the components go in through the debug api and the
# setting can stay OFF, which is where it belongs (2026-08-07).
#
# Three phases, because each needs a different thing:
#   1. An empty macro workbook. Only Excel can make one, and an automation-created Excel is fine
#      for it - no add-in is needed to save a blank file.
#   2. The same workbook opened as an ORDINARY process with the editor up, which is what loads
#      the add-in and therefore what opens the door.
#   3. The components, through the door.

$harness = Join-Path $PSScriptRoot 'harness'
$client = Join-Path $harness 'xlide-api.mjs'

Write-Host '1. Making an empty macro workbook.'
$maker = New-Object -ComObject Excel.Application
$maker.Visible = $false
$maker.DisplayAlerts = $false
try {
    $blank = $maker.Workbooks.Add()
    # 52 is xlOpenXMLWorkbookMacroEnabled: a workbook that cannot hold macros is no fixture.
    $blank.SaveAs($Path, 52)
    $blank.Close($false)
}
finally {
    try { $maker.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($maker) | Out-Null
}

Write-Host '2. Opening it with the editor, which is what loads the add-in.'
& (Join-Path $harness 'Start-Excel.ps1') -Workbook $Path -Fresh | Write-Host

# The module texts go to node as JSON, so that quoting, CRLFs and VBA's own doubled quotes cross
# once rather than being escaped through two shells.
$plan = @{
    modules = @(
        foreach ($name in $modules.Keys) {
            @{ name = $name; kind = $modules[$name].Kind; code = $modules[$name].Code }
        }
    )
    sheetCode = $sheetCode
    openAtEnd = 'Watcher'
}

# Written WITHOUT a byte-order mark. In PowerShell 5.1 `-Encoding utf8` means "UTF-8 with a BOM",
# and JSON.parse refuses one - naming a character that does not appear to be in the file.
$planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-fixture-$PID.json"
[System.IO.File]::WriteAllText(
    $planPath,
    ($plan | ConvertTo-Json -Depth 5),
    (New-Object System.Text.UTF8Encoding $false))

Write-Host '3. Writing the components through the debug api.'
try {
    & node (Join-Path $harness 'build-fixture.mjs') $planPath | Write-Host
    if ($LASTEXITCODE -ne 0) { throw 'the fixture could not be built through the api' }
}
finally {
    Remove-Item $planPath -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Helpers        standard module, renamed from the explorer or from code'
Write-Host '  HelpersExtra   shares a prefix and must not move'
Write-Host '  Rival          declares its own Recalculate and must not move'
Write-Host '  IShape         interface: Implements, the IShape_Draw prefix, and As IShape'
Write-Host '  RoundShape     implements IShape'
Write-Host '  Widget         ordinary class: As Widget, New Widget'
Write-Host '  Consumer       names everything - LEFT CLOSED on purpose'
Write-Host '  Watcher        names some of it - LEFT OPEN on purpose'
Write-Host '  Sheet1         document module, the untested case'
Write-Host ''
Write-Host 'After any rename, check Consumer - the module with no tab is the one that'
Write-Host 'silently gets missed, and it is where every string and comment must be intact.'
Write-Host ''
Write-Host 'NOTE: this fixture deliberately does NOT compile. Helpers and Rival each declare a'
Write-Host 'public Recalculate and Consumer calls it bare, which is the collision it exists for.'
Write-Host 'Never compile it as part of a wider experiment.'
Write-Host ''

if ($Quiet) {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
