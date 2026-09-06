<#
.SYNOPSIS
    Builds the Access database the Access-specific behaviour is tested against.

.DESCRIPTION
    THE GATE HAD NO ACCESS IN IT AT ALL, which is why the test runner could never have run a test
    there and nobody knew (2026-09-06). Every other fixture is an .xlsm, so every suite ran against
    Excel and the three places the product asks "which host is this?" were only ever answered one
    way. This is the other answer.

    Access is not Excel with a different icon. Its VBA has no UserForms, its `Application.Run`
    takes a bare procedure name where Excel's wants the file and Word's the module, its
    application object is reached off the OMain frame rather than a document window, and it keeps
    one database per process rather than a workbook collection. The database here is deliberately
    small - what it exists for is to be OPENED IN ACCESS, so that the suites can ask those
    questions of a real host.

      Pricing        plain arithmetic, the thing under test: a discount table and an extended price
      PricingTests   four '@xlide-test procedures, three that pass and one that fails on purpose,
                     so a run has both outcomes to report
      Helpers        one function with no arguments, for the Immediate window to evaluate

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading. XlideAssert is installed through the
    tests route, the same way the pane's own button installs it.

.EXAMPLE
    tools\New-AccessFixture.ps1
    Builds it and leaves Access open on it.

.EXAMPLE
    tools\New-AccessFixture.ps1 -Quiet
    Builds it and closes the session it opened.
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

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'harness\AccessAttach.psm1') -Force
if (-not $Path) {
    $fixtures = Join-Path $repoRoot 'artifacts\fixtures'
    if (-not (Test-Path $fixtures)) { New-Item -ItemType Directory -Force $fixtures | Out-Null }
    $Path = Join-Path $fixtures 'AccessFixture.accdb'
}

# Component types, as VBComponents.Add takes them. Access has no UserForm kind at all.
$StandardModule = 1

# DOUBLE-QUOTED here-strings throughout: a module line that starts with '@ at column 0 closes a
# single-quoted one, and every test below opens with the directive (lessons, 2026-09-05).
$modules = [ordered]@{}

$modules['Pricing'] = @{ Kind = $StandardModule; Code = @"
Option Compare Database
Option Explicit

' The arithmetic under test. Nothing here touches Access itself, on purpose: what the Access
' fixture is for is the HOST, and a test that also depended on a table would confuse the two.

' Quantity discounts: 5% from five units, 10% from ten.
Public Function DiscountRate(ByVal Qty As Long) As Double
    If Qty >= 10 Then
        DiscountRate = 0.1
    ElseIf Qty >= 5 Then
        DiscountRate = 0.05
    Else
        DiscountRate = 0
    End If
End Function

Public Function Extended(ByVal Qty As Long, ByVal Price As Currency, ByVal Express As Boolean) As Currency
    Extended = Qty * Price * (1 - DiscountRate(Qty))
    If Express Then Extended = Extended + 4.5
End Function
"@ }

$modules['PricingTests'] = @{ Kind = $StandardModule; Code = @"
Option Compare Database
Option Explicit

' THREE THAT PASS AND ONE THAT FAILS, because a runner that can only report success proves half
' of itself. The failing one is failing on purpose and says so in its own name.

'@xlide-test
Public Sub NoDiscountUnderFive()
    XlideAssert.IsTrue Pricing.DiscountRate(4) = 0
End Sub

'@xlide-test
Public Sub FivePercentFromFive()
    XlideAssert.IsTrue Pricing.DiscountRate(5) = 0.05
End Sub

'@xlide-test
Public Sub ExpressAddsItsFee()
    XlideAssert.IsTrue Pricing.Extended(1, 2, True) = 6.5
End Sub

'@xlide-test
Public Sub FailsOnPurpose()
    XlideAssert.IsTrue Pricing.DiscountRate(1) = 99
End Sub
"@ }

$modules['Helpers'] = @{ Kind = $StandardModule; Code = @"
Option Compare Database
Option Explicit

' For the Immediate window, which is the other thing that could not reach the host in Access:
' one function with no arguments and an answer that is obvious when it is wrong.
Public Function Greeting() As String
    Greeting = "hello from access"
End Function
"@ }

# ---------------------------------------------------------------- building it

Write-Host '1. Making an empty database.'
New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null

# A session left open on the LAST build of this fixture still holds the file, and a database
# cannot be replaced underneath the host reading it. Only a session showing this fixture's own
# name is closed - the -Fresh guard in the launcher below is what protects everything else.
foreach ($running in @(Get-Process MSACCESS -ErrorAction SilentlyContinue)) {
    $where = Get-AccessDatabasePath -ProcessId $running.Id
    if ($null -ne $where -and $where -eq $Path) {
        Write-Host "  closing the session still holding it (pid $($running.Id))"
        Stop-Process -Id $running.Id -Force
        Start-Sleep -Seconds 3
    }
}

if (Test-Path $Path) { Remove-Item $Path -Force }
$lock = [System.IO.Path]::ChangeExtension($Path, '.laccdb')
if (Test-Path $lock) { Remove-Item $lock -Force -ErrorAction SilentlyContinue }

# Made through automation, which is FINE here: the maker never needs the add-in, only a database
# on disk for the real launch below to open as an ordinary process.
$maker = New-Object -ComObject Access.Application
try {
    $maker.NewCurrentDatabase($Path)
    $maker.CloseCurrentDatabase()
}
finally {
    try { $maker.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($maker) | Out-Null
}
Start-Sleep -Seconds 1

Write-Host '2. Opening it with the editor, which is what loads the add-in.'
& (Join-Path $PSScriptRoot 'harness\Start-Access.ps1') -Database $Path -Fresh | Write-Host
$env:XLIDE_PID = (Get-Process MSACCESS | Select-Object -First 1).Id

$plan = @{
    modules = @(
        foreach ($name in $modules.Keys) {
            @{ name = $name; kind = $modules[$name].Kind; code = $modules[$name].Code }
        }
    )
    openAtEnd = 'Pricing'
}

# Written WITHOUT a byte-order mark: PowerShell 5.1's `-Encoding utf8` means "UTF-8 with a BOM",
# and JSON.parse refuses one.
$planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-access-fixture-$PID.json"
[System.IO.File]::WriteAllText(
    $planPath,
    ($plan | ConvertTo-Json -Depth 8),
    (New-Object System.Text.UTF8Encoding $false))

Write-Host '3. Writing the components through the xlide api.'
$blankStamp = (Get-Item $Path).LastWriteTimeUtc
try {
    & node (Join-Path $PSScriptRoot 'harness\build-fixture.mjs') $planPath | Write-Host
    if ($LASTEXITCODE -ne 0) { throw 'the fixture could not be built through the api' }
}
finally {
    Remove-Item $planPath -ErrorAction SilentlyContinue
}

Write-Host '4. Installing XlideAssert, the way the Tests pane installs it.'
& node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID tests install | Write-Host

# EACH MODULE SAVED BY NAME, which is Access's own rule and nobody else's. A module CREATED
# through the editor's object model lives in the VBA project and is not written into the database
# by saving the database: close the session without saving the module itself and the module is
# simply gone, which is how the first build of this fixture came back with an empty file. Excel
# and Word carry a new module along with the document; Access wants to be told about each one.
$saver = Get-AccessApplication -ProcessId ([int] $env:XLIDE_PID) -TimeoutSeconds 30
if ($null -eq $saver) { throw 'Could not reach Access to save the modules it was given.' }
$acModule = 5
foreach ($name in @($modules.Keys) + @('XlideAssert')) {
    $saver.DoCmd.Save($acModule, $name)
    Write-Host "  saved $name into the database"
}

& node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID command save | Write-Host

# THE SAVE IS NOT OVER WHEN THE SAVE COMMAND RETURNS: the api answers when the host has been told,
# and the host writes the file on its own message loop afterwards. Waited for here, where the path
# is known, so a -Quiet close cannot beat the write to disk (lessons, 2026-09-05).
$deadline = (Get-Date).AddSeconds(30)
while ((Get-Item $Path).LastWriteTimeUtc -le $blankStamp) {
    if ((Get-Date) -gt $deadline) {
        throw "the editor's save did not reach $Path within 30 seconds"
    }
    Start-Sleep -Milliseconds 200
}

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Pricing        the arithmetic under test'
Write-Host '  PricingTests   four tests: three pass, one fails on purpose'
Write-Host '  Helpers        one no-argument function, for the Immediate window'
Write-Host ''
Write-Host '  tools\harness\Start-Access.ps1 -Database artifacts\fixtures\AccessFixture.accdb -Fresh'
Write-Host '  node tools\harness\access.mjs'
Write-Host ''

if ($Quiet) {
    # ONLY THE SESSION THIS BUILT, named in XLIDE_PID by the launch above.
    if ($env:XLIDE_PID) { Stop-Process -Id ([int] $env:XLIDE_PID) -Force -ErrorAction SilentlyContinue }
}
