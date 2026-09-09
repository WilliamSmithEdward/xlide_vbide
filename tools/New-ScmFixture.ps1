<#
.SYNOPSIS
    Builds ScmFixture.xlsm: the workbook source control is exercised against.

.DESCRIPTION
    The Source Control pane puts a git repository behind the folder a project's modules are
    exported to. Proving it means committing, branching, blaming, restoring and importing over a
    real workbook, and asking questions no other fixture can answer without ambiguity:

      - did the commit take the ticked module and leave the other one?
      - does blame carry the committed file's lines onto the editor's lines of THIS module?
      - did the checkout's import put back the text this module was built with?

    Every one of those compares a module's text against a commit's, so the bodies here are
    written to make a wrong answer visible rather than plausible. Each module says its own name in
    a string, so a blame or a comparison that has fetched the wrong module's text reads as the
    wrong module rather than as text that happens to look similar. And the three modules are
    deliberately DIFFERENT LENGTHS, so a line count taken from the wrong one cannot coincide with
    the right one - which matters most to blame, whose whole answer is line numbers.

    A SEPARATE FIXTURE, not a borrowed one. The suite commits, checks a branch out over the
    workbook, saves it, and imports into it - every one of which rewrites modules and the file on
    disk. A fixture another suite reads would be either useless for that or a landmine for the
    other suite. The repository itself is NOT in the workbook: the suite initialises one in a
    temporary folder of its own and deletes it at the end, so this file carries nothing but VBA.

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading, which is what makes the door exist.

    It COMPILES, so a session holding it is safe to Run and Compile.

.EXAMPLE
    tools\New-ScmFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-ScmFixture.ps1 -Quiet
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
    $Path = Join-Path $fixtures 'ScmFixture.xlsm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

# THE ONE THE SUITE EDITS, BLAMES AND RESTORES. Short on purpose: blame answers one row per line
# and every one of them is checked, and a restore is a whole-module comparison in the reply.
$modules['Ledger'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' THE MODULE THE SOURCE CONTROL SUITE EDITS, BLAMES AND RESTORES. Its body is expected to
' move; the suite writes to it, commits it, blames it and restores it from a commit.
'
' It names itself in the string below so a blame or a comparison that has fetched some other
' module's text reads as the wrong module rather than as text that merely looks similar.

Public Function Balance(ByVal opening As Long, ByVal movement As Long) As Long
    Balance = opening + movement
    Debug.Print "Ledger.Balance"
End Function

Public Sub Settle(ByVal amount As Long)
    Debug.Print "Ledger.Settle " & amount
End Sub
'@ }

# COMMITTED IN A ROUND OF ITS OWN and carried across a branch. Longer than Ledger, so a line
# count taken from the wrong module is a different number rather than the same one by luck.
$modules['Reports'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' A SECOND STANDARD MODULE. The suite edits it beside Ledger and commits only Ledger, which
' is how "the commit took the ticked row and left the other" is a real question; then it
' carries this module's uncommitted edit across a checkout, the way git carries a change that
' does not collide.
'
' Deliberately longer than Ledger, so a line count taken from the wrong module is a different
' number rather than the same one by luck.

Public Function Summarise(ByVal opening As Long, ByVal closing As Long) As String
    Summarise = "Reports.Summarise: " & opening & " to " & closing
End Function

Public Function Heading(ByVal title As String) As String
    Heading = "Reports: " & title
End Function

Public Sub Announce()
    Debug.Print "Reports.Announce"
End Sub

Public Function Share(ByVal part As Long, ByVal whole As Long) As Double
    If whole = 0 Then
        Share = 0
    Else
        Share = part / whole
    End If
End Function
'@ }

# A CLASS, so the rows, the files and the folder are exercised over a .cls with the header a
# class file carries - the header the live rows deliberately do not compare. The suite edits
# its FILE in the folder and imports it back.
$modules['Account'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' A class module, so source control is exercised over more than one component kind. The suite
' edits this module's file in the export folder, watches the pane notice, and imports it.

Private mNumber As String

Public Property Get Number() As String
    Number = mNumber
End Property

Public Property Let Number(ByVal value As String)
    mNumber = value
    Debug.Print "Account.Number"
End Property
'@ }

$sheetCode = @'
Option Explicit

' A DOCUMENT module. It cannot be created or deleted like the others, so it is the row a
' checkout can never add or remove - and it exports as a .cls that carries the marks making it
' a document, which a repository has to round-trip untouched.

Public Sub Refresh()
    Debug.Print "Sheet1.Refresh"
End Sub
'@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Ledger'

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Ledger    the module the suite edits, blames and restores; its body moves'
Write-Host '  Reports   longer, committed in a round of its own, carried across a branch'
Write-Host '  Account   a CLASS, edited in the folder by the suite and imported back'
Write-Host '  Sheet1    a document module: the kind a checkout can never add or remove'
Write-Host ''
Write-Host 'The repository is not in the workbook: scm.mjs initialises one in a temporary folder.'
Write-Host ''

if ($Quiet) {
    # ONLY THE BUILDER'S OWN SESSION, which Invoke-FixtureLaunch named in XLIDE_PID. This used
    # to stop every Excel on the machine, the owner's open workbooks included (2026-09-05).
    if ($env:XLIDE_PID) { Stop-Process -Id ([int] $env:XLIDE_PID) -Force -ErrorAction SilentlyContinue }
}
