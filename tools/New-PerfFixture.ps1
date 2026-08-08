<#
.SYNOPSIS
    Builds PerfFixture.xlsm: modules of graduated size, for measuring how the editor scales.

.DESCRIPTION
    Every performance figure taken so far came from a module of a few dozen lines, which answers
    the wrong question. What a developer wants to know is not "is hover fast" but "is hover still
    fast in the module I actually work in" — and the honest way to answer that is a curve rather
    than a number.

    So this holds the same shape of code at four sizes:

      Small     ~100 lines
      Medium    ~1,000 lines
      Large     ~4,000 lines
      Huge      ~10,000 lines

    Each is the same generated pattern, so a difference between them is size and nothing else.
    Every one compiles, and every procedure is reachable, so the analyzer has real work to do
    rather than a wall of comments to skip.

    It is deliberately NOT part of the gate: it takes a while to build and longer to measure, and
    a timing that runs on every commit becomes a timing nobody reads.

.EXAMPLE
    tools\New-PerfFixture.ps1
    node tools\harness\perf-scaling.mjs
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\PerfFixture.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null
if (Test-Path $Path) { Remove-Item $Path -Force }

$StandardModule = 1

# One generated procedure, repeated. Each calls the one before it, so the analyzer has a real
# call graph to resolve rather than a list of unrelated stubs.
function New-Body {
    param([int] $Count, [string] $Prefix)

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.AppendLine('Option Explicit')
    [void]$builder.AppendLine('')
    [void]$builder.AppendLine("' Generated. Every procedure calls the one before it, so the analyzer has a real call")
    [void]$builder.AppendLine("' graph to resolve rather than a list of unrelated stubs.")
    [void]$builder.AppendLine('')

    for ($n = 0; $n -lt $Count; $n++) {
        $previous = if ($n -eq 0) { $null } else { "$Prefix$($n - 1)" }
        [void]$builder.AppendLine("Public Function $Prefix$n(ByVal seed As Long) As Long")
        [void]$builder.AppendLine('    Dim total As Long')
        [void]$builder.AppendLine("    total = seed * $($n + 1)")
        if ($previous) {
            [void]$builder.AppendLine('    If total > 100 Then')
            [void]$builder.AppendLine("        total = total - $previous(seed)")
            [void]$builder.AppendLine('    End If')
        }
        [void]$builder.AppendLine("    $Prefix$n = total")
        [void]$builder.AppendLine('End Function')
        [void]$builder.AppendLine('')
    }

    return $builder.ToString().TrimEnd() + "`r`n"
}

$modules = [ordered]@{}
$modules['Small']  = @{ Kind = $StandardModule; Code = (New-Body -Count 12   -Prefix 'S') }
$modules['Medium'] = @{ Kind = $StandardModule; Code = (New-Body -Count 125  -Prefix 'M') }
$modules['Large']  = @{ Kind = $StandardModule; Code = (New-Body -Count 500  -Prefix 'L') }
$modules['Huge']   = @{ Kind = $StandardModule; Code = (New-Body -Count 1250 -Prefix 'H') }

$sheetCode = @'
Option Explicit

' Empty. This fixture is about module size and nothing else.
'@

# ---------------------------------------------------------------- building it

$harness = Join-Path $PSScriptRoot 'harness'

Write-Host '1. Making an empty macro workbook.'
$maker = New-Object -ComObject Excel.Application
$maker.Visible = $false
$maker.DisplayAlerts = $false
try {
    $blank = $maker.Workbooks.Add()
    $blank.SaveAs($Path, 52)
    $blank.Close($false)
}
finally {
    try { $maker.Quit() } catch { }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($maker) | Out-Null
}

Write-Host '2. Opening it with the editor, which is what loads the add-in.'
& (Join-Path $harness 'Start-Excel.ps1') -Workbook $Path -Fresh | Write-Host

$plan = @{
    modules = @(
        foreach ($name in $modules.Keys) {
            @{ name = $name; kind = $modules[$name].Kind; code = $modules[$name].Code }
        }
    )
    sheetCode = $sheetCode
    openAtEnd = 'Small'
}

$planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-perf-fixture-$PID.json"
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

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  The same generated shape at four sizes, so a difference between them is size'
    Write-Host '  and nothing else. Measure with: node tools\harness\perf-scaling.mjs'
    Write-Host ''
}
