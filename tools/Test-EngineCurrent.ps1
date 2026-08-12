<#
.SYNOPSIS
    Which engine sources are newer than the packaged executable. Empty means it is current.

.DESCRIPTION
    The add-in launches engine\dist\xlide-engine.exe. `npm run build` writes only the bundle, so an
    engine change can be built, tested, committed and released while the thing that actually runs
    is hours old and refuses every new method as unknown.

    THE ANALYZER COUNTS AS ENGINE SOURCE. It lives in the neighbouring xlide_vscode checkout and is
    bundled INTO this executable, so a pull over there changes what the add-in runs without
    touching a single file in this repository. Watching engine\src alone would call a stale
    executable current, which is the one answer this must never give.

    This is the shared half of two callers that had drifted apart. tools\verify.ps1 asks and then
    packages when the answer is not empty; tools\release.ps1 asks and REFUSES, because a release is
    the wrong moment to silently rebuild the thing being shipped. Before this existed, release.ps1
    did not ask at all when it was run with -SkipGate, and installer\build.ps1 copies whatever .exe
    is sitting in engine\dist - so the one path that skipped the gate was also the one path that
    could ship a stale engine, silently.

.OUTPUTS
    The FileInfo objects that are newer than the executable. Nothing when it is current.
#>
[CmdletBinding()]
param(
    # The repository root. Defaults to the parent of this script's folder.
    [string] $RepoRoot = (Split-Path -Parent $PSScriptRoot)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$engineRoot = Join-Path $RepoRoot 'engine'
$exe = Join-Path $engineRoot 'dist\xlide-engine.exe'
if (-not (Test-Path $exe)) {
    throw 'engine\dist\xlide-engine.exe has never been packaged; run npm run package in engine\'
}

$builtAt = (Get-Item $exe).LastWriteTimeUtc

$watched = @(Join-Path $engineRoot 'src')
$analyzer = Join-Path (Split-Path -Parent $RepoRoot) 'xlide_vscode\src'
if (Test-Path $analyzer) { $watched += $analyzer }

# Emitted so a caller can say which coverage it got. An absent analyzer checkout is not a failure
# here - it is a fact the caller has to be able to report, because "current" means less without it.
$script:AnalyzerWasFound = $watched.Count -gt 1

@(Get-ChildItem $watched -Recurse -File -Include *.ts, *.mjs, *.js |
    Where-Object { $_.LastWriteTimeUtc -gt $builtAt })
