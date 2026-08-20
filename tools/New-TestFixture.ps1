<#
.SYNOPSIS
    Builds TestFixture.xlsm: a workbook whose tests exercise the test runner.

.DESCRIPTION
    A small product module (Invoice) and a test module (InvoiceTests) written the way a
    developer writes them: mostly green, one deliberately red so the pane's failure row and
    the Failed rerun have something real to show, plus a skip, an expected failure, and an
    expected-error - one of every outcome the runner can answer.

    XlideAssert is installed through the runner's own install action after the modules land,
    then saved with the workbook, so the fixture opens ready to run: press the beaker on the
    bar, then Run All.

    Nothing here touches the worksheet, so the tests can run forever without a rebuild.

.EXAMPLE
    tools\New-TestFixture.ps1
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\TestFixture.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

$StandardModule = 1

$modules = [ordered]@{}

$modules['Invoice'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' THE PRODUCT UNDER TEST: small enough to hold in the head, real enough to get wrong.

Public Function TotalOf(ByVal net As Double, ByVal taxRate As Double) As Double
    TotalOf = net * (1 + taxRate)
End Function

Public Function DiscountFor(ByVal quantity As Long) As Double
    If quantity >= 100 Then
        DiscountFor = 0.1
    ElseIf quantity >= 10 Then
        DiscountFor = 0.05
    Else
        DiscountFor = 0
    End If
End Function

Public Function LabelFor(ByVal net As Double) As String
    LabelFor = "Total: " & Format$(net, "0.00")
End Function

Public Sub RequirePositive(ByVal amount As Double)
    If amount <= 0 Then
        Err.Raise 5, "Invoice", "amount must be positive"
    End If
End Sub
'@ }

$modules['InvoiceTests'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' The runner's authoring surface, one test per behavior. A directive comment immediately
' above a zero-argument public Sub makes it a test; the assertions are XlideAssert's.

' @xlide-test tags="math,smoke"
Public Sub TotalOf_AddsTax()
    XlideAssert.AreEqual 108, Invoice.TotalOf(100, 0.08)
End Sub

' @xlide-test tags="math"
Public Sub TotalOf_ZeroRateIsIdentity()
    XlideAssert.AreEqual 250, Invoice.TotalOf(250, 0)
End Sub

' @xlide-test tags="discount"
Public Sub DiscountFor_TenPercentAtAHundred()
    XlideAssert.AreEqual 0.1, Invoice.DiscountFor(100)
End Sub

' @xlide-test tags="discount"
Public Sub DiscountFor_NothingUnderTen()
    XlideAssert.AreEqual 0, Invoice.DiscountFor(9)
End Sub

' @xlide-test tags="text"
Public Sub LabelFor_FormatsTwoPlaces()
    XlideAssert.StartsWith Invoice.LabelFor(12.5), "Total: "
    XlideAssert.Contains Invoice.LabelFor(12.5), "12.50"
End Sub

' @xlide-test expected-error=5
Public Sub RequirePositive_RefusesZero()
    Invoice.RequirePositive 0
End Sub

' @xlide-test
Public Sub RequirePositive_ThrowsByHelper()
    XlideAssert.Throws 5, "CallWithZero"
End Sub

' The Throws target: a plain public Sub, not itself a test.
Public Sub CallWithZero()
    Invoice.RequirePositive 0
End Sub

' RED ON PURPOSE, so the pane's failure row, the message under it, and the Failed rerun have
' something true to show. The output line proves assertions are not fatal.
' @xlide-test tags="smoke" owner=william
Public Sub Rounding_KnownWrong()
    XlideAssert.WriteLine "the rate table says 8.25 here"
    XlideAssert.AreEqual 108.25, Invoice.TotalOf(100, 0.08), "rounding drifted"
End Sub

' @xlide-test-skip reason="needs the ledger workbook open"
Public Sub Ledger_CrossBooks()
    XlideAssert.Fail "never runs without the ledger"
End Sub

' @xlide-test-xfail reason="defect 12: discount boundary is off by one"
Public Sub DiscountFor_BoundaryAtTen()
    XlideAssert.AreEqual 0.05, Invoice.DiscountFor(9)
End Sub
'@ }

$sheetCode = @'
Option Explicit

' Empty on purpose. The test fixture touches no cells, so its tests run forever.
'@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'InvoiceTests'

# XlideAssert through the runner's own install action - the same module a press of the pane's
# Install button writes - then saved, so the fixture opens ready to run.
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
    Write-Host '  Invoice        the product under test: totals, discounts, labels, a guard that raises'
    Write-Host '  InvoiceTests   eleven directives, one of every outcome: green, a deliberate red with'
    Write-Host '                 output, expected-error, Assert.Throws, a skip, and an xfail'
    Write-Host '  XlideAssert    installed and saved, so the fixture opens ready to run'
    Write-Host ''
    Write-Host '  Press the beaker on the bar, then Run All. Rounding_KnownWrong stays red on purpose.'
    Write-Host ''
}
