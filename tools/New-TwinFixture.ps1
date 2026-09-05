<#
.SYNOPSIS
    Builds the SECOND workbook: the one whose module names collide with RenameFixture's.

.DESCRIPTION
    A module name is not an identity. Two open workbooks can each hold a Helpers, each with its
    own Recalculate, and every defect in this product's handling of that fact has been found by
    eye rather than by a test. The reason is that the interesting state needs two workbooks open
    at once, and the harness could only reach it if somebody had already built this file.

    This fixture is the twin of RenameFixture.xlsm and is meant to be opened ALONGSIDE it:

        tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\RenameFixture.xlsm,artifacts\fixtures\TwinFixture.xlsm
        node tools\harness\surface-walk.mjs --steps 80 --seed 424242

    The walk reports a `collision` count, which is how many of its steps reached a state holding
    two same-named modules. A run reporting zero has passed every label check vacuously, and that
    line is how the `pane` route's dropped project argument was found.

    The names below deliberately match RenameFixture's, and the BODIES deliberately do not. Same
    name and same text is a fixture that cannot tell a mix-up from a success: every check would
    pass while reading the wrong workbook's module. So each one says which workbook it is from,
    in text a comparison will notice.

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. That setting gates VBComponents.Add through Workbook.VBProject; the add-in is
    already past it, because the host hands it the VBE. A Debug build must be registered and
    loading, which is what makes the door exist.

    Unlike RenameFixture, this one COMPILES. It has to be safe to Run and Compile in a session
    where both are open, because the twin is scenery for tests about the other one.

.EXAMPLE
    tools\New-TwinFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-TwinFixture.ps1 -Path C:\temp\twin.xlsm -Quiet
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
    $Path = Join-Path $fixtures 'TwinFixture.xlsm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

$modules['Helpers'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' THE COLLIDING MODULE. RenameFixture.xlsm holds a Helpers too, with its own Recalculate.
' Nothing that acts on "Helpers" may act on both, and nothing that reads "Helpers" may
' answer from whichever workbook enumerated first.
'
' The body says which workbook it came from ON PURPOSE. Identical twins prove nothing: a
' route reading the wrong one would return text that matched, and every check would pass.

Public Sub Recalculate(ByVal label As String)
    Debug.Print "twin " & label
End Sub

Public Sub OnlyInTheTwin()
    ' Declared here and nowhere in RenameFixture, so "does this module come from the twin"
    ' can be answered by a name lookup as well as by comparing text.
    Recalculate "only in the twin"
End Sub
'@ }

$modules['Consumer'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' A SECOND collision, and the one that matters for navigation: go to definition from here
' must land in this workbook's Helpers, not in the other workbook's.

Public Sub Drive()
    Helpers.Recalculate "qualified, in the twin"
End Sub
'@ }

$modules['Widget'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' A colliding CLASS, so the collision is not only a standard-module case.

Public Sub Spin()
    Debug.Print "twin widget"
End Sub
'@ }

$modules['TwinOnly'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' Belongs to no other fixture. A tree, tab strip or explorer reading the wrong workbook is
' missing this name entirely, which is a cheaper thing to assert on than comparing text.

Public Sub Marker()
    Debug.Print "twin only"
End Sub
'@ }

$sheetCode = @'
Option Explicit

' A DOCUMENT module in the second workbook. Two workbooks both hold a Sheet1, which is the
' collision case that needs no VBA at all to create.

Public Sub Refresh()
    Debug.Print "twin sheet"
End Sub
'@

# ---------------------------------------------------------------- building it

# Built through the DOOR, not through Workbook.VBProject - the same three phases
# New-RenameFixture.ps1 uses, and for the same reason: VBComponents.Add is exactly what the
# trust setting gates, and the add-in is already past that gate.
#
#   1. An empty macro workbook. Only Excel can make one, and an automation-created Excel is
#      fine for it - no add-in is needed to save a blank file.
#   2. The same workbook opened as an ORDINARY process with the editor up, which is what loads
#      the add-in and therefore what opens the door.
#   3. The components, through the door.

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Helpers'

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Helpers    collides with RenameFixture.xlsm''s Helpers, different body'
Write-Host '  Consumer   collides, and qualifies Helpers so navigation has a wrong answer to give'
Write-Host '  Widget     a colliding CLASS, so the case is not only standard modules'
Write-Host '  TwinOnly   belongs to no other fixture, so the workbook can be identified by name'
Write-Host '  Sheet1     a second document module, the collision that needs no VBA to create'
Write-Host ''
Write-Host 'Open it ALONGSIDE RenameFixture.xlsm - alone it proves nothing:'
Write-Host '  tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\RenameFixture.xlsm,artifacts\fixtures\TwinFixture.xlsm'
Write-Host ''
Write-Host 'This one DOES compile, unlike RenameFixture, so it is safe to Run in a shared session.'
Write-Host ''

if ($Quiet) {
    # ONLY THE BUILDER'S OWN SESSION, which Invoke-FixtureLaunch named in XLIDE_PID. This used
    # to stop every Excel on the machine, the owner's open workbooks included (2026-09-05).
    if ($env:XLIDE_PID) { Stop-Process -Id ([int] $env:XLIDE_PID) -Force -ErrorAction SilentlyContinue }
}
