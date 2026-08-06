<#
.SYNOPSIS
    Builds the production installer and attaches it to a release.

.DESCRIPTION
    A tag on its own gives someone the source. What they want is the thing that installs, so every
    release carries xlide-setup.exe.

    The installer is built here rather than in CI because it contains the language engine, and the
    engine is built from the neighbouring editor-extension checkout that CI does not have. That is
    the dependency decision 3a removes; until it does, releasing is a local step.

    The gate runs first and the installer refuses to build without an engine or a page, so an
    incomplete build fails here instead of being uploaded and discovered by whoever downloads it.

.EXAMPLE
    tools\release.ps1 -Tag v0.2.1

.EXAMPLE
    tools\release.ps1 -Tag v0.2.1 -SkipGate
    Skips the verification gate, for when it has already been run in this session.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $Tag,

    [switch] $SkipGate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not (git -C $repoRoot tag --list $Tag)) {
    throw "No tag $Tag. Create and push it first, then run this."
}

$head = (git -C $repoRoot rev-parse HEAD).Trim()
$tagged = (git -C $repoRoot rev-list -n 1 $Tag).Trim()
if ($head -ne $tagged) {
    throw "HEAD is $($head.Substring(0,7)) but $Tag is $($tagged.Substring(0,7)). Check out the tag before building what it will carry."
}

if (-not $SkipGate) {
    Write-Host '==> Gate' -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'verify.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'The gate failed; nothing was uploaded.' }
}

Write-Host '==> Installer' -ForegroundColor Cyan
& (Join-Path $repoRoot 'installer\build.ps1')
if ($LASTEXITCODE -ne 0) { throw 'The installer build failed; nothing was uploaded.' }

$installer = Join-Path $repoRoot 'artifacts\xlide-setup.exe'
if (-not (Test-Path $installer)) {
    throw "No installer at $installer. An editor-only build produces a differently named file and is not releasable."
}

$size = (Get-Item $installer).Length / 1MB
Write-Host ''
Write-Host ("==> Attaching xlide-setup.exe ({0:N1} MB) to {1}" -f $size, $Tag) -ForegroundColor Cyan

# --clobber so a re-run replaces the asset rather than failing, which is what you want when a
# release is corrected without moving the tag.
gh release upload $Tag $installer --clobber
if ($LASTEXITCODE -ne 0) { throw "Uploading to $Tag failed." }

Write-Host ''
Write-Host "Attached to $Tag." -ForegroundColor Green
gh release view $Tag --json assets --jq '.assets[] | "  \(.name)  \(.size / 1048576 | floor) MB"'
