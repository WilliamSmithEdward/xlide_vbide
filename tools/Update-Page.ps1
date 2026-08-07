<#
.SYNOPSIS
    Builds the editor page and puts it in front of a running editor, without restarting the host.

.DESCRIPTION
    A page change does not need a republish. The shim serves the bundle over loopback from a
    folder on disk, reading each file as it is asked for, so replacing that folder and reloading
    the page is the whole update. What forces a restart is the SHIM: a host holds an add-in
    library open for its lifetime, and nothing can replace a file Excel is holding.

    Knowing which of the two you changed is worth about a minute per iteration, and a whole
    session's worth of them was spent closing Excel for changes that never touched the shim
    (2026-08-07).

.EXAMPLE
    tools\Update-Page.ps1
    Build, copy, reload every live editor, and report the build each is now running.

.EXAMPLE
    tools\Update-Page.ps1 -NoBuild
    Copy whatever was built last and reload. For when the gate has just built it.
#>
[CmdletBinding()]
param(
    # Use the bundle already in ui\editor\dist.
    [switch] $NoBuild,

    # Copy the files but leave the page as it is.
    [switch] $NoReload,

    # Which published shim to update beside.
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Debug'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$pageRoot = Join-Path $repoRoot 'ui\editor'
$built = Join-Path $pageRoot 'dist'
$publish = Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64"
$served = Join-Path $publish 'ui\editor\dist'

if (-not $NoBuild) {
    Push-Location $pageRoot
    try { & npm run build | Write-Host } finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw 'the page did not build' }
}

if (-not (Test-Path $built)) { throw "No bundle at $built." }
if (-not (Test-Path $publish)) { throw "No published shim at $publish. Publish once before using this." }

New-Item -ItemType Directory -Force -Path $served | Out-Null
Copy-Item -Path (Join-Path $built '*') -Destination $served -Recurse -Force
Write-Host "Copied the bundle into $served." -ForegroundColor Green

if ($NoReload) { exit 0 }

# Reloading is the api's own business: it reloads and WAITS for the page to say it is ready,
# which is what makes the build stamp it reports the new page's rather than the old one's.
& node (Join-Path $repoRoot 'tools\harness\reload-page.mjs') | Write-Host
if ($LASTEXITCODE -ne 0) { throw 'the page was copied but a live editor would not reload' }
