<#
.SYNOPSIS
    Builds the Word document the Word-specific behaviour is tested against.

.DESCRIPTION
    THE THIRD HOST, AND THE LAST WITHOUT A FIXTURE. Access got its database on 2026-09-06, when
    the test runner turned out never to have run there. Word had a launcher from 2026-08-19 and
    nothing for it to open, so every Word branch in the product - the `_WwG` document pane the
    application is reached through, `Application.Run` spelled `Module.Proc`, the Documents
    collection the save and the dirty check walk - was written by symmetry and proven by nobody.
    This is the document those branches are proven on (2026-09-10).

    Word is not Excel with a different icon either, in two ways the suite has to know:

      NORMAL IS ALWAYS OPEN. Word's global template is a VBA project of its own, listed in the
      editor beside every document, so a Word session is two projects by construction. That is
      what makes "the active project" a real question here, and why every write this builder
      makes names the document: a fixture module landing in the developer's own Normal.dotm
      would follow them to every document they open afterwards.

      A DOCUMENT CARRIES ITS MODULES. Unlike Access, a module added through the object model is
      written into the .docm when the document is saved, so the editor's Save is the whole save.

      Pricing        plain arithmetic, the thing under test: a discount table and an extended price
      PricingTests   four '@xlide-test procedures, three that pass and one that fails on purpose,
                     so a run has both outcomes to report
      Helpers        one function with no arguments, for the Immediate window to evaluate

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading. XlideAssert is installed through the
    tests route, scoped to the document by name: unscoped, the install goes into every open file
    that lacks it, and in Word that includes Normal.

.EXAMPLE
    tools\New-WordFixture.ps1
    Builds it and leaves Word open on it.

.EXAMPLE
    tools\New-WordFixture.ps1 -Quiet
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
Import-Module (Join-Path $PSScriptRoot 'harness\WordAttach.psm1') -Force
if (-not $Path) {
    $fixtures = Join-Path $repoRoot 'artifacts\fixtures'
    if (-not (Test-Path $fixtures)) { New-Item -ItemType Directory -Force $fixtures | Out-Null }
    $Path = Join-Path $fixtures 'WordFixture.docm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1

# DOUBLE-QUOTED here-strings throughout: a module line that starts with '@ at column 0 closes a
# single-quoted one, and every test below opens with the directive (lessons, 2026-09-05).
$modules = [ordered]@{}

$modules['Pricing'] = @{ Kind = $StandardModule; Code = @"
Option Explicit

' The arithmetic under test. Nothing here touches Word itself, on purpose: what the Word fixture
' is for is the HOST, and a test that also depended on a paragraph would confuse the two.

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
Option Explicit

' For the Immediate window, which runs in the ACTIVE project - and with Normal open beside the
' document, that is the question this answers: one function with no arguments and an answer that
' is obvious when it is wrong.
Public Function Greeting() As String
    Greeting = "hello from word"
End Function
"@ }

# Word's owner file beside a document - `~$` and the name less its first letter or two - stays
# behind when the session holding it is killed, and the next open then reads as locked by another
# user. Stale once nothing holds the document, which the census above guarantees.
function Remove-OwnerFileOf([string] $document) {
    $leaf = Split-Path -Leaf $document
    Get-ChildItem (Split-Path -Parent $document) -Filter '~$*.docm' -Force -ErrorAction SilentlyContinue |
        Where-Object { $leaf.EndsWith($_.Name.Substring(2), [StringComparison]::OrdinalIgnoreCase) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

# THE SAVE IS NOT OVER WHEN THE SAVE COMMAND RETURNS: the api answers when the host has been told,
# and the host writes the file on its own message loop afterwards. Waited for here, where the path
# is known, so a -Quiet close cannot beat the write to disk (lessons, 2026-09-05).
function Wait-WrittenAfter([datetime] $stamp, [string] $what) {
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Item $Path).LastWriteTimeUtc -le $stamp) {
        if ((Get-Date) -gt $deadline) { throw "$what did not reach $Path within 30 seconds" }
        Start-Sleep -Milliseconds 200
    }
}

# ---------------------------------------------------------------- building it

Write-Host '1. Making an empty macro-enabled document.'
New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null

# A session left open on the LAST build of this fixture still holds the file, and a document
# cannot be replaced underneath the host reading it. Only a session holding this fixture's own
# path is closed - the -Fresh guard in the launcher below is what protects everything else.
foreach ($running in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    $held = Get-WordDocumentPaths -ProcessId $running.Id
    if ($null -ne $held -and $held -contains $Path) {
        Write-Host "  closing the session still holding it (pid $($running.Id))"
        Stop-Process -Id $running.Id -Force
        Start-Sleep -Seconds 3
    }
}

if (Test-Path $Path) { Remove-Item $Path -Force }
Remove-OwnerFileOf $Path

# Made through automation, which is FINE here: the maker never needs the add-in, only a
# macro-enabled file on disk for the real launch below to open as an ordinary process.
$maker = New-Object -ComObject Word.Application
$maker.DisplayAlerts = 0
$blank = $null
try {
    $blank = $maker.Documents.Add()
    $blank.SaveAs2($Path, 13)  # wdFormatXMLDocumentMacroEnabled
    $blank.Close($false)
}
finally {
    if ($null -ne $blank) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($blank) | Out-Null }
    try { $maker.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($maker) | Out-Null
}
Start-Sleep -Seconds 1
$blankStamp = (Get-Item $Path).LastWriteTimeUtc

Write-Host '2. Opening it with the editor, which is what loads the add-in.'
# THE PID THE LAUNCHER PRINTS, never the first WINWORD in the process list: that may be the
# developer's own Word, and the launcher then gives the fixture a process of its own.
$said = @(& (Join-Path $PSScriptRoot 'harness\Start-Word.ps1') -Document $Path -Fresh *>&1 | ForEach-Object { "$_" })
$said | Write-Host
$started = $said | Select-String 'pid=(\d+)' | Select-Object -Last 1
if (-not $started) { throw 'Start-Word.ps1 did not print the pid it started.' }
$env:XLIDE_PID = $started.Matches[0].Groups[1].Value
$document = Split-Path -Leaf $Path

$plan = @{
    # NAMED, because Word holds two projects: unnamed, the components go into whichever project
    # the editor calls active, and after a fresh launch that can be Normal.
    project = $document
    modules = @(
        foreach ($name in $modules.Keys) {
            @{ name = $name; kind = $modules[$name].Kind; code = $modules[$name].Code }
        }
    )
    openAtEnd = 'Pricing'
}

# Written WITHOUT a byte-order mark: PowerShell 5.1's `-Encoding utf8` means "UTF-8 with a BOM",
# and JSON.parse refuses one.
$planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-word-fixture-$PID.json"
[System.IO.File]::WriteAllText(
    $planPath,
    ($plan | ConvertTo-Json -Depth 8),
    (New-Object System.Text.UTF8Encoding $false))

Write-Host '3. Writing the components through the xlide api, into the document by name.'
try {
    & node (Join-Path $PSScriptRoot 'harness\build-fixture.mjs') $planPath | Write-Host
    if ($LASTEXITCODE -ne 0) { throw 'the fixture could not be built through the api' }
}
finally {
    Remove-Item $planPath -ErrorAction SilentlyContinue
}
Wait-WrittenAfter $blankStamp "the editor's save of the modules"

Write-Host '4. Installing XlideAssert into the document, the way the Tests pane installs it with the document chosen.'
# SCOPED TO THE DOCUMENT. Unscoped, the install goes into every open file that lacks it, and in
# Word that is Normal too - the positional shape is the client's own: action, module, test, file.
$builtStamp = (Get-Item $Path).LastWriteTimeUtc
& node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID tests install '' '' $document | Write-Host
& node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID command save | Write-Host
Wait-WrittenAfter $builtStamp "the editor's save after the install"

# READ BACK, not assumed: every module of the fixture's reads from the DOCUMENT by name, and
# none of them reads from any other project the editor holds - Normal, in every Word session.
# The client's `module` verb answers a module's text or fails naming what it could not find,
# and the failure is the answer wanted for the other projects.
$expected = @($modules.Keys) + @('XlideAssert')
foreach ($name in $expected) {
    & node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID module $name $document 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "$name did not make it into $document" }
}

$projects = & node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID projects | Out-String | ConvertFrom-Json
foreach ($other in @($projects.projects | Where-Object { $_.project -ne $document })) {
    foreach ($name in $expected) {
        & node (Join-Path $PSScriptRoot 'harness\xlide-api.mjs') --pid $env:XLIDE_PID module $name $other.project 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { throw "$name landed in $($other.project), which is not the fixture" }
    }
    Write-Host "  $($other.project) is open beside it and holds nothing of the fixture's"
}

# The absence checks above end on a native failure by design, and a caller reading the last exit
# code would take the fixture for unbuilt. This script's verdict is a throw or the file.
$global:LASTEXITCODE = 0

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Pricing        the arithmetic under test'
Write-Host '  PricingTests   four tests: three pass, one fails on purpose'
Write-Host '  Helpers        one no-argument function, for the Immediate window'
Write-Host ''
Write-Host '  tools\harness\Start-Word.ps1 -Document artifacts\fixtures\WordFixture.docm -Fresh'
Write-Host '  node tools\harness\word.mjs'
Write-Host ''

if ($Quiet) {
    # ONLY THE SESSION THIS BUILT, named in XLIDE_PID by the launch above.
    if ($env:XLIDE_PID) { Stop-Process -Id ([int] $env:XLIDE_PID) -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    Remove-OwnerFileOf $Path
}
