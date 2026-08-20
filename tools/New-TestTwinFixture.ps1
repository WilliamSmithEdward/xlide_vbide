<#
.SYNOPSIS
    Builds TestTwinFixture.xlsm: the SECOND file of tests, to be opened beside TestFixture.xlsm.

.DESCRIPTION
    The runner discovers and runs tests in EVERY open file, and every defect in that fact needs
    two files open at once to reach. This is the second file.

    Its module names deliberately match TestFixture's - Invoice, InvoiceTests - because a module
    name is not an identity across files, and a runner that keyed anything on the name alone
    would run this file's Adds when asked for the other's. The BODIES deliberately differ, and
    the outcomes differ too: the twin's deliberate red fails with its own message, so a check
    reading the wrong file's result cannot pass by accident.

    Its Ledger module has no counterpart in TestFixture, which is what makes a file scope
    visibly different from a module scope: scoping to this file shows tests the other file has
    no name for.

    XlideAssert is installed here as well, because the generated runner calls it INSIDE the file
    it is running - so a session with two files of tests needs two copies, and this fixture is
    also how that per-file gate gets exercised.

    Built through the debug api, so "Trust access to the VBA project object model" does NOT have
    to be on.

.EXAMPLE
    tools\New-TestTwinFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\TestFixture.xlsm,artifacts\fixtures\TestTwinFixture.xlsm
    Opens the pair, which is the state the Tests pane's file scope is for.
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\TestTwinFixture.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

$StandardModule = 1

$modules = [ordered]@{}

# SAME NAME as TestFixture's, DIFFERENT ARITHMETIC. A check that reads this file's TotalOf
# expecting the other file's answer fails loudly rather than passing by coincidence.
$modules['Invoice'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' The twin's product under test. Same name as TestFixture's Invoice, deliberately: a module
' name is only unique inside its own file, and this pair is how that gets tested.

Public Function TotalOf(ByVal net As Double, ByVal taxRate As Double) As Double
    ' The twin charges a flat handling fee on top, so its answers differ from its namesake's.
    TotalOf = net * (1 + taxRate) + 5
End Function

Public Function Handling() As Double
    Handling = 5
End Function
'@ }

$modules['InvoiceTests'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' @xlide-test tags="math,twin"
Public Sub TotalOf_AddsTaxAndHandling()
    XlideAssert.AreEqual 113, Invoice.TotalOf(100, 0.08)
End Sub

' @xlide-test tags="twin"
Public Sub Handling_IsFive()
    XlideAssert.AreEqual 5, Invoice.Handling()
End Sub

' RED ON PURPOSE, and red with ITS OWN WORDS: a check that read the other file's failure
' message would see "rounding drifted" instead of this, which is how the pair proves a result
' was filed against the file it came from.
' @xlide-test tags="twin"
Public Sub Handling_KnownWrong()
    XlideAssert.AreEqual 7, Invoice.Handling(), "the twin's handling fee moved"
End Sub
'@ }

# NO COUNTERPART in TestFixture: scoping to this file shows a module the other file has no
# name for, which is what makes a file scope visibly more than a module scope.
$modules['Ledger'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

Public Function Balance(ByVal opening As Double, ByVal movement As Double) As Double
    Balance = opening + movement
End Function
'@ }

$modules['LedgerTests'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' @xlide-test tags="ledger,twin"
Public Sub Balance_AddsMovement()
    XlideAssert.AreEqual 150, Ledger.Balance(100, 50)
End Sub

' @xlide-test-xfail reason="defect 40: a withdrawal past zero is not refused yet"
Public Sub Balance_RefusesOverdraft()
    XlideAssert.AreEqual 0, Ledger.Balance(10, -50)
End Sub
'@ }

$sheetCode = @'
Option Explicit

' Empty on purpose. The twin touches no cells either, so its tests run forever.
'@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'InvoiceTests'

# The support module goes in here too: the generated runner calls XlideAssert inside the file
# it runs, so each file of tests carries its own copy.
Write-Host '4. Installing XlideAssert through the tests route, and saving.'
$harness = Join-Path $PSScriptRoot 'harness'
& node (Join-Path $harness 'xlide-api.mjs') tests install | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'XlideAssert could not be installed through the api' }
& node (Join-Path $harness 'xlide-api.mjs') command save | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'the workbook could not be saved with XlideAssert in it' }

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  Invoice        SAME NAME as TestFixture''s, different arithmetic: it charges handling'
    Write-Host '  InvoiceTests   three tests, one red with its own message'
    Write-Host '  Ledger         no counterpart in TestFixture, so a file scope shows more than a module one'
    Write-Host '  LedgerTests    a green and an expected failure'
    Write-Host '  XlideAssert    installed and saved, because the runner calls it inside this file'
    Write-Host ''
    Write-Host '  Open it BESIDE TestFixture.xlsm:'
    Write-Host '    tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\TestFixture.xlsm,artifacts\fixtures\TestTwinFixture.xlsm'
    Write-Host ''
}
