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

    tools\verify.ps1 asks and then packages when the answer is not empty. It once had a second
    caller, tools\release.ps1, which asked and refused, because installer\build.ps1 copies whatever
    .exe is sitting in engine\dist. Releases are now built in CI (.github\workflows\publish.yml),
    which packages the engine fresh from the pinned analyzer every time, so no release can carry a
    stale one.

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

# THE SOURCES THE ENGINE WAS ACTUALLY BUILT FROM, which is the pin when one was used. Comparing a
# pinned build against the working tree next door asks the wrong question and answers it wrongly:
# the tree is edited all day, so it always looks newer, and the release it blocks is one whose
# engine cannot contain those edits at all. XLIDE_ANALYZER_ROOT is set by whoever ran the build
# (see tools\Pin-Analyzer.ps1) and is read the same way by engine\build.mjs.
$analyzer = if ($env:XLIDE_ANALYZER_ROOT) {
    $env:XLIDE_ANALYZER_ROOT
} else {
    Join-Path (Split-Path -Parent $RepoRoot) 'xlide_vscode\src'
}
if (Test-Path $analyzer) { $watched += $analyzer }

# Emitted so a caller can say which coverage it got. An absent analyzer checkout is not a failure
# here - it is a fact the caller has to be able to report, because "current" means less without it.
$script:AnalyzerWasFound = $watched.Count -gt 1

@(Get-ChildItem $watched -Recurse -File -Include *.ts, *.mjs, *.js |
    Where-Object { $_.LastWriteTimeUtc -gt $builtAt })
