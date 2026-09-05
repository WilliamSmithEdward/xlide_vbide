<#
.SYNOPSIS
    Builds the SECOND workbook for the folder layout: the one whose names collide with
    FolderFixture's (#23).

.DESCRIPTION
    A module name is not an identity, and neither is a folder name. Two open workbooks can each
    hold a Helpers, and each can have a folder called Shared; a tree that pooled either by name
    would draw one workbook's module under the other's folder. Every defect of that class in this
    product has been found with two workbooks open, so this fixture exists to be opened beside
    FolderFixture.xlsm and nowhere else:

      Helpers     the same name as FolderFixture's Helpers, in a DIFFERENT folder (Accounts.Ledger)
      Ledger      the same name as FolderFixture's Ledger, with NO annotation: at the root here
      TwinOnly    in a folder called Shared, which FolderFixture also has: two folders, one name
      Posting     a class in a folder only this workbook has
      Sheet1      a document module with no annotation, where FolderFixture's is in Accounts

        tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\FolderFixture.xlsm,artifacts\fixtures\FolderTwinFixture.xlsm
        node tools\harness\folders.mjs

    The bodies say which workbook they came from ON PURPOSE. Identical twins prove nothing: a
    route reading the wrong one would return text that matched, and every check would pass.

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading. Everything here compiles.

.EXAMPLE
    tools\New-FolderTwinFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-FolderTwinFixture.ps1 -Quiet
    Builds it and closes Excel afterwards.
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
    $Path = Join-Path $fixtures 'FolderTwinFixture.xlsm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

$modules['Helpers'] = @{ Kind = $StandardModule; Code = @"
'@Folder("Accounts.Ledger")
Option Explicit

' THE COLLIDING MODULE. FolderFixture.xlsm holds a Helpers in Shared; this one is in
' Accounts.Ledger. A tree filing by name alone would put one of them in the wrong folder.

Public Sub Recalculate(ByVal label As String)
    Debug.Print "twin " & label
End Sub
"@ }

$modules['Ledger'] = @{ Kind = $StandardModule; Code = @"
Option Explicit

' NO ANNOTATION, where FolderFixture's Ledger is in Accounts.Ledger: the same name at the root
' of one workbook and inside a folder of the other.

Public Sub Post(ByVal amount As Currency)
    Debug.Print "twin ledger " & amount
End Sub
"@ }

$modules['TwinOnly'] = @{ Kind = $StandardModule; Code = @"
'@Folder("Shared")
Option Explicit

' A FOLDER CALLED SHARED, which FolderFixture also has. Two workbooks, two folders, one name:
' this module must never be drawn under the other workbook's Shared, nor theirs under this.

Public Sub Marker()
    Debug.Print "twin only"
End Sub
"@ }

$modules['Posting'] = @{ Kind = $ClassModule; Code = @"
'@Folder("Twin.Only")
Option Explicit

' A folder only this workbook has, so the two workbooks' folder lists differ by construction.

Public Function Kind() As String
    Kind = "twin"
End Function
"@ }

$sheetCode = @"
Option Explicit

' A DOCUMENT module with no annotation, where FolderFixture's Sheet1 is in Accounts.

Public Sub Refresh()
    Debug.Print "twin sheet"
End Sub
"@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Helpers'

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Accounts'
Write-Host '    Ledger                      Helpers'
Write-Host '  Shared                        TwinOnly'
Write-Host '  Twin'
Write-Host '    Only                        Posting'
Write-Host '  (root)                        Sheet1, ThisWorkbook, Ledger'
Write-Host ''
Write-Host 'Open it ALONGSIDE FolderFixture.xlsm - alone it proves nothing:'
Write-Host '  tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\FolderFixture.xlsm,artifacts\fixtures\FolderTwinFixture.xlsm'
Write-Host ''

if ($Quiet) {
    # ONLY THE BUILDER'S OWN SESSION, which Invoke-FixtureLaunch named in XLIDE_PID. This used
    # to stop every Excel on the machine, the owner's open workbooks included (2026-09-05).
    if ($env:XLIDE_PID) { Stop-Process -Id ([int] $env:XLIDE_PID) -Force -ErrorAction SilentlyContinue }
}
