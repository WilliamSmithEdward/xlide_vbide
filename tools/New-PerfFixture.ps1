<#
.SYNOPSIS
    Builds PerfFixture.xlsm: modules of graduated size, for measuring how the editor scales.

.DESCRIPTION
    Every performance figure taken so far came from a module of a few dozen lines, which answers
    the wrong question. What a developer wants to know is not "is hover fast" but "is hover still
    fast in the module I actually work in" - and the honest way to answer that is a curve rather
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

# AT THE CEILING. VBA will not hold more than 65,534 lines in one module, and 7,200 procedures of
# this shape come to 64,805 - close enough to the wall to find anything that scales with module
# size, with room to type in it. Measured 2026-08-08 by writing progressively larger modules and
# reading them back: 65,000 lines was accepted, and took the EDITOR 17.4 seconds to take (its own
# parse, not this product's write, which is two COM calls at any size).
#
# It exists because the four sizes above all fit comfortably and so agreed with each other about
# what was fast. Anything quadratic in module size is invisible at 11,000 lines and obvious here.
$modules['Massive'] = @{ Kind = $StandardModule; Code = (New-Body -Count 7200 -Prefix 'V') }

$sheetCode = @'
Option Explicit

' Empty. This fixture is about module size and nothing else.
'@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Small'

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  The same generated shape at five sizes, so a difference between them is size'
    Write-Host '  and nothing else. Massive sits at VBA''s per-module line ceiling, which is where'
    Write-Host '  anything quadratic in module size stops hiding. Measure with:'
    Write-Host '    node tools\harness\perf-scaling.mjs'
    Write-Host ''
}
