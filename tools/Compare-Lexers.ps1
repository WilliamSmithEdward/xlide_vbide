<#
.SYNOPSIS
    Compares the ported lexer with the reference implementation over a corpus.

.DESCRIPTION
    This is the gate for the lexer half of the port. Reading ported code and judging it correct
    scales badly and misses exactly the cases nobody thought to check. Running both implementations
    over thousands of real modules and diffing the token streams does not.

    Any difference is a defect in the port until proven otherwise. Where the port is deliberately
    different, that difference belongs here as a documented exclusion, not in a code comment nobody
    reads.

.EXAMPLE
    tools\Compare-Lexers.ps1
    Compares over every VBA source file found in the default corpus.

.EXAMPLE
    tools\Compare-Lexers.ps1 -Path 'F:\some\folder' -ShowFirst 5
#>
[CmdletBinding()]
param(
    # Folders to search for VBA sources. Defaults to the corpora in the neighbouring checkout.
    [string[]] $Path,

    # Stop after this many differing files.
    [int] $ShowFirst = 3,

    # Rebuild both dumpers first.
    [switch] $Rebuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent $here
$reference = Join-Path (Split-Path -Parent $repoRoot) 'xlide_vscode'

$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path (Join-Path $localDotnet 'dotnet.exe')) {
    $env:PATH = "$localDotnet;$env:PATH"

    # A built program launches through a small native host that locates the runtime through the
    # registry or this variable, not through PATH. Without it, a runtime installed under the user's
    # profile is invisible and the program reports that .NET is not installed at all.
    $env:DOTNET_ROOT = $localDotnet
}

$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

if (-not $Path) {
    $Path = @(
        (Join-Path $reference 'syntax_corpus'),
        (Join-Path $reference 'excel_test_workbook'),
        (Join-Path $repoRoot 'tests')
    ) | Where-Object { Test-Path $_ }
}

if (-not $Path) {
    throw 'No corpus found. Pass -Path with a folder containing .bas, .cls, or .frm files.'
}

$nodeDump = Join-Path $repoRoot 'engine\dist\lexdump.cjs'

if ($Rebuild -or -not (Test-Path $nodeDump)) {
    Write-Host '==> Building the reference dumper' -ForegroundColor Cyan
    Push-Location (Join-Path $repoRoot 'engine')
    try { node build.mjs | Out-Null } finally { Pop-Location }
}

Write-Host '==> Building the ported dumper' -ForegroundColor Cyan
dotnet build (Join-Path $repoRoot 'tools\Xlide.Lex.Dump\Xlide.Lex.Dump.csproj') -c Release --nologo | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'The ported dumper did not build.' }

$portedDump = Join-Path $repoRoot 'artifacts\bin\Xlide.Lex.Dump\release\xlide-lexdump.exe'
if (-not (Test-Path $portedDump)) { throw "The ported dumper was not produced at $portedDump." }

$files = @()
foreach ($root in $Path) {
    $files += Get-ChildItem $root -Recurse -File -Include *.bas, *.cls, *.frm, *.vba -ErrorAction SilentlyContinue
}

if ($files.Count -eq 0) {
    throw "No VBA sources found under: $($Path -join ', ')"
}

Write-Host "==> Comparing $($files.Count) file(s)" -ForegroundColor Cyan

$same = 0
$differing = New-Object System.Collections.Generic.List[string]
$shown = 0

foreach ($file in $files) {
    $expected = & node $nodeDump $file.FullName 2>$null
    $actual = & $portedDump $file.FullName 2>$null

    if (($expected -join "`n") -eq ($actual -join "`n")) {
        $same++
        continue
    }

    $differing.Add($file.FullName)

    if ($shown -lt $ShowFirst) {
        $shown++
        Write-Host ''
        Write-Host "--- $($file.FullName)" -ForegroundColor Yellow

        # Report the first place they diverge. A single early difference usually explains every
        # later one, so showing the whole diff buries the cause.
        $limit = [Math]::Min($expected.Count, $actual.Count)
        for ($i = 0; $i -lt $limit; $i++) {
            if ($expected[$i] -ne $actual[$i]) {
                Write-Host "  at token $i" -ForegroundColor Yellow
                Write-Host "    reference: $($expected[$i])"
                Write-Host "    ported:    $($actual[$i])"
                break
            }
        }

        if ($expected.Count -ne $actual.Count) {
            Write-Host "  token count: reference $($expected.Count), ported $($actual.Count)" -ForegroundColor Yellow
        }
    }
}

Write-Host ''
Write-Host "identical: $same of $($files.Count)" -ForegroundColor $(if ($differing.Count -eq 0) { 'Green' } else { 'Yellow' })

if ($differing.Count -gt 0) {
    Write-Host "differing: $($differing.Count)" -ForegroundColor Red
    exit 1
}

Write-Host 'The ported lexer agrees with the reference on every file in the corpus.' -ForegroundColor Green
exit 0
