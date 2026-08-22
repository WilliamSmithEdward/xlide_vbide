<#
.SYNOPSIS
    Builds ChangeFixture.xlsm: the workbook the change log is exercised against.

.DESCRIPTION
    The change log records what happened to a project's module code, by whom, in rounds. Reading
    it back means asking questions no other fixture can answer without ambiguity:

      - did the round that touched Ledger touch ONLY Ledger?
      - is the text it held BEFORE the round the text this fixture was built with?
      - are two rounds by two different hands two rounds, or did they fold into one?

    Every one of those compares a module's text against what it used to be, so the bodies here are
    written to make a wrong answer visible rather than plausible. Each module says its own name in
    a string, so a log entry that has fetched the wrong module's text reads as the wrong module
    rather than as text that happens to look similar. And the three modules are deliberately
    DIFFERENT LENGTHS, so a line count taken from the wrong one cannot coincide with the right one.

    A SEPARATE FIXTURE, not a borrowed one. The suites for the change log edit modules and then
    assert on what the log says was edited - which means they must be free to change text without
    any other suite's expectations resting on it. DebugFixture is read by nine suites; a change
    log run against it would be either useless or a landmine for the other eight.

    Built through the debug api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading, which is what makes the door exist.

    It COMPILES, so a session holding it is safe to Run and Compile.

.EXAMPLE
    tools\New-ChangeFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-ChangeFixture.ps1 -Quiet
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
    $Path = Join-Path $fixtures 'ChangeFixture.xlsm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

# THE ONE THE SUITES EDIT. Short on purpose: a suite rewrites it repeatedly, and every rewrite
# is a whole-module comparison in the log's own reply.
$modules['Ledger'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' THE MODULE THE CHANGE LOG SUITES EDIT. Its body is expected to move; nothing else in this
' workbook should, which is how "the round touched only Ledger" is a real question.
'
' It names itself in the string below so a log entry that has fetched some other module's text
' reads as the wrong module rather than as text that merely looks similar.

Public Function Balance(ByVal opening As Long, ByVal movement As Long) As Long
    Balance = opening + movement
    Debug.Print "Ledger.Balance"
End Function
'@ }

# NEVER EDITED, and different in length from Ledger so a count read off the wrong module cannot
# coincidentally match. A round that claims to have touched this one is wrong.
$modules['Untouched'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' NOT EDITED BY ANY SUITE. It exists so a change log that over-reports has something to
' over-report ABOUT: a round listing this module is a round that recorded a write nobody made.
'
' Deliberately longer than Ledger, so a line count taken from the wrong module is a different
' number rather than the same one by luck.

Public Sub Announce()
    Debug.Print "Untouched.Announce"
End Sub

Public Function Describe(ByVal what As String) As String
    Describe = "Untouched: " & what
End Function

Public Function Doubled(ByVal value As Long) As Long
    Doubled = value * 2
End Function
'@ }

# A CLASS, so the log is not only exercised over standard modules.
$modules['Ticket'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' A class module, so the log is exercised over more than one component kind. Edited by the
' suites that check a second module joining a round.

Private mReference As String

Public Property Get Reference() As String
    Reference = mReference
End Property

Public Property Let Reference(ByVal value As String)
    mReference = value
    Debug.Print "Ticket.Reference"
End Property
'@ }

$sheetCode = @'
Option Explicit

' A DOCUMENT module. It cannot be created or deleted like the others, which makes it the case
' the log has to record differently from a standard module - and the case a rollback could never
' put back, which is one of several reasons this product logs rather than reverts.

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
Write-Host '  Ledger     the module the suites edit; its body is expected to move'
Write-Host '  Untouched  edited by nobody, and a different length, so over-reporting is visible'
Write-Host '  Ticket     a CLASS, so the log is not exercised over standard modules alone'
Write-Host '  Sheet1     a document module: the kind that cannot be added or removed'
Write-Host ''
Write-Host 'The change log itself lives beside the shim log, per workbook:'
Write-Host "  $env:LOCALAPPDATA\xlide_vbide\changes\"
Write-Host ''

if ($Quiet) {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
