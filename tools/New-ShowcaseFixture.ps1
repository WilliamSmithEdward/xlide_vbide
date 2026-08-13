<#
.SYNOPSIS
    Builds QuarterlyReport.xlsm: the workbook the README's pictures are taken in.

.DESCRIPTION
    Not a test fixture. The other workbooks are shaped for suites - modules named for what they
    exercise, code written to be stepped through - and a picture of DebugFixture.xlsm is a
    picture of scaffolding. This one is shaped to be LOOKED at: a workbook name that could be
    anyone's, modules named for work rather than for tests, and code that does something a
    reader recognises.

    It COMPILES, deliberately, because the debugger scene presses Run. The module of deliberate
    findings the diagnostics scene needs (Validation) is NOT here: tools\tour\capture-tour.mjs
    adds it through the api mid-session, after the debugger scene is taken, and never saves -
    so this workbook stays runnable however many times the tour is reshot.

    Regenerate when the tour's content changes, not per release. The pictures are committed;
    this is how they stay reproducible.

.EXAMPLE
    tools\New-ShowcaseFixture.ps1
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\QuarterlyReport.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

$StandardModule = 1
$ClassModule = 2

# No `Dim x As New <class>` anywhere here: the analyzer wrongly flags member access on one as
# object-variable-not-set (xlide_vscode#16), and the hero shot's problems panel must read zero.
# Explicit Set ... = New is what the code means anyway.
$modules = [ordered]@{}

$modules['Reporting'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' Builds the quarter's summary: a total per invoice, then the grand total.

Public Sub BuildQuarterlySummary()
    Dim invoices As Collection
    Dim inv As Invoice
    Dim grandTotal As Double
    Dim counted As Long

    Set invoices = LoadInvoices()

    For Each inv In invoices
        grandTotal = grandTotal + inv.Total
        counted = counted + 1
    Next inv

    Debug.Print counted & " invoices, " & Format$(grandTotal, "#,##0.00") & " total"
End Sub

Public Function LoadInvoices() As Collection
    Dim loaded As Collection
    Set loaded = New Collection

    loaded.Add MakeInvoice("ACME-1042", 1250#, 40#)
    loaded.Add MakeInvoice("ACME-1043", 987.5, 0#)
    loaded.Add MakeInvoice("NORTH-2117", 2402.25, 65#)

    Set LoadInvoices = loaded
End Function

Private Function MakeInvoice(ByVal number As String, ByVal subtotal As Double, ByVal freight As Double) As Invoice
    Dim made As Invoice
    Set made = New Invoice
    made.Number = number
    made.Subtotal = subtotal
    made.AddFreight freight
    Set MakeInvoice = made
End Function
'@ }

$modules['Invoice'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' One invoice of the quarter. Totals are derived, never stored.

Private mNumber As String
Private mSubtotal As Double
Private mFreight As Double

Public Property Get Number() As String
    Number = mNumber
End Property

Public Property Let Number(ByVal value As String)
    mNumber = value
End Property

Public Property Get Subtotal() As Double
    Subtotal = mSubtotal
End Property

Public Property Let Subtotal(ByVal value As Double)
    mSubtotal = value
End Property

Public Sub AddFreight(ByVal amount As Double)
    mFreight = mFreight + amount
End Sub

Public Property Get Tax() As Double
    Tax = mSubtotal * 0.0825
End Property

Public Property Get Total() As Double
    Total = mSubtotal + mFreight + Tax
End Property
'@ }

$modules['Formatting'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' Presentation of the Summary sheet. Nothing here computes; it only dresses.

Public Sub PolishSummary(ByVal target As Worksheet)
    With target.Range("A1").CurrentRegion
        .Columns.AutoFit
        .Rows(1).Font.Bold = True
    End With

    target.Columns("D:F").NumberFormat = "#,##0.00"
    target.Range("A1").Value = "Quarter at a glance"
End Sub

Public Function SheetNamed(ByVal wanted As String) As Worksheet
    Dim candidate As Worksheet
    For Each candidate In ThisWorkbook.Worksheets
        If candidate.Name = wanted Then
            Set SheetNamed = candidate
            Exit Function
        End If
    Next candidate
End Function
'@ }

$sheetCode = @'
Option Explicit

' Nothing here. The sheet exists because a workbook has one.
'@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Reporting'

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  Reporting   the summary build: a Collection, a For Each, and a call per invoice'
    Write-Host '  Invoice     a class with properties and derived totals, for the dot menu'
    Write-Host '  Formatting  Excel object model use that reads as real work'
    Write-Host ''
    Write-Host '  It COMPILES. The findings module the diagnostics scene needs is added live by'
    Write-Host '  tools\tour\capture-tour.mjs and never saved.'
    Write-Host ''
}
