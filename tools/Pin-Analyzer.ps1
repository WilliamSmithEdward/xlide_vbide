<#
.SYNOPSIS
    Makes a read-only copy of the analyzer checkout at one commit, to build the engine from.

.DESCRIPTION
    The engine bundles the editor extension's analyzer from the neighbouring checkout, which is
    a WORKING TREE someone is usually working in. Three releases have now been blocked by that:
    the currency guard compares file times, so any edit in progress makes the packaged engine
    look stale, and repackaging to satisfy the guard would ship whatever is half-written that
    minute. Committing does not help - git does not touch a file's modification time.

    So a release builds from a PIN instead: a clone of that repository at a chosen commit, made
    once, never edited, and left alone afterwards. Cloning reads the source repository and writes
    nothing to it, and it copies committed objects only, so work in progress there is neither
    included nor disturbed.

    Prints the root to build with. Point the engine build at it:

        $env:XLIDE_ANALYZER_ROOT = (tools\Pin-Analyzer.ps1 -Ref v9.0.0)
        npm run package --prefix engine

    The same variable is read by the engine-currency checks in verify.ps1 and release.ps1, so a
    pinned build is compared against the pinned sources rather than against the working tree.

.PARAMETER Ref
    The commit, tag or branch to pin. A release pins the analyzer's release commit.

.PARAMETER Source
    The analyzer checkout to clone from. Defaults to the sibling the engine normally builds
    against.

.PARAMETER Force
    Rebuild the pin even if it is already there.

.EXAMPLE
    tools\Pin-Analyzer.ps1 -Ref e887745
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $Ref,

    [string] $Source,

    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Source) { $Source = Join-Path (Split-Path -Parent $repoRoot) 'xlide_vscode' }

if (-not (Test-Path (Join-Path $Source '.git'))) {
    throw "No git repository at $Source, so there is nothing to pin. Pass -Source."
}

# RESOLVED IN THE SOURCE, so the pin folder is named by commit whatever spelling was asked for,
# and two names for one commit do not make two clones.
$commit = (git -C $Source rev-parse --verify "$Ref^{commit}" 2>$null)
if ($LASTEXITCODE -ne 0 -or -not $commit) {
    throw "$Source has no commit for '$Ref'."
}
$commit = $commit.Trim()
$short = $commit.Substring(0, 7)

# Under artifacts, which is not in source control: a pin is a build input, reproducible from the
# ref at any time, and far too large for history.
$pinRoot = Join-Path $repoRoot "artifacts\analyzer-pin\$short"
$src = Join-Path $pinRoot 'src'

if ((Test-Path $src) -and -not $Force) {
    Write-Host "Pin already at $pinRoot ($short)" -ForegroundColor DarkGray
    return $src
}

if (Test-Path $pinRoot) { Remove-Item $pinRoot -Recurse -Force }
New-Item -ItemType Directory -Path (Split-Path -Parent $pinRoot) -Force | Out-Null

Write-Host "==> Pinning $Source at $short" -ForegroundColor Cyan

# THE OBJECT STORE IS COPIED, NOT CLONED. `git clone` reads the source's `.git` as a path of its
# own and refuses it as "dubious ownership" whenever the two checkouts were made by different
# accounts, which is the normal state on this machine - and the remedy git suggests is an entry
# in the GLOBAL config, a shared setting changed to do one build. A copy needs no trust: the
# pin's `.git` is then this user's own. It carries committed objects only, so work in progress
# in the source is neither included nor disturbed, and the checkout below is pristine.
New-Item -ItemType Directory -Path $pinRoot -Force | Out-Null
Copy-Item (Join-Path $Source '.git') (Join-Path $pinRoot '.git') -Recurse -Force

git -C $pinRoot config core.bare false
git -C $pinRoot checkout --quiet --force $commit
if ($LASTEXITCODE -ne 0) { throw "Could not check out $commit in the pin." }

$dirty = @(git -C $pinRoot status --porcelain)
if ($dirty.Count -gt 0) {
    throw "The pin at $pinRoot is not pristine after checkout ($($dirty.Count) path(s) differ)."
}

if (-not (Test-Path (Join-Path $src 'analysisWorkerLogic.ts'))) {
    throw "The pin at $pinRoot has no src\analysisWorkerLogic.ts; is $Source the analyzer repository?"
}

# No null-coalescing here: this is a tool a 5.1 console may run.
$described = (git -C $pinRoot describe --tags --always $commit 2>$null)
if (-not $described) { $described = $short }
Write-Host "Pinned $described at $pinRoot"
Write-Host "  build with: `$env:XLIDE_ANALYZER_ROOT = '$src'"

return $src
