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

# Checked before anything is built, because the alternative is finding out after the gate and a
# compressed installer that there was nothing to attach it to.
gh release view $Tag --json tagName 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "No release for $Tag. Create it first: gh release create $Tag --title ... --notes ..."
}

if (-not $SkipGate) {
    Write-Host '==> Gate' -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'verify.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'The gate failed; nothing was uploaded.' }
}

# ASKED WHETHER THE GATE RAN OR NOT.
#
# installer\build.ps1 copies whatever .exe is sitting in engine\dist, and the gate's currency step
# was the only thing anywhere that checked it was the right one. So -SkipGate, the flag whose whole
# purpose is a faster re-release, was also the only path that could ship an engine older than the
# analyzer it was supposedly built from - silently, with a green run and a correct-looking
# installer. That is precisely the failure the release checklist exists for, and it was reachable
# by a flag.
#
# It REFUSES rather than packaging, which is the difference between this and the gate. A release is
# the wrong moment to rebuild the thing being shipped: the decision about what goes out was made
# when the tag was cut, and an installer that quietly contains a newer engine than the one the gate
# measured is not the artifact anybody approved.
Write-Host '==> Engine currency' -ForegroundColor Cyan
$stale = @(& (Join-Path $PSScriptRoot 'Test-EngineCurrent.ps1') -RepoRoot $repoRoot)
if ($stale.Count -gt 0) {
    $names = ($stale | Select-Object -First 4 | ForEach-Object { $_.Name }) -join ', '
    throw ("$($stale.Count) engine source(s) are newer than engine\dist\xlide-engine.exe ($names). " +
        'The installer copies that executable as it stands, so this release would ship a stale ' +
        'engine. Close Excel, run `npm run package --prefix engine`, and start again.')
}

$analyzerCheckout = Join-Path (Split-Path -Parent $repoRoot) 'xlide_vscode\src'
if (Test-Path $analyzerCheckout) {
    Write-Host '  packaged after every engine source, analyzer included'
} else {
    Write-Warning ('the analyzer checkout was not found, so the engine was only checked against ' +
        'engine\src. The analyzer is bundled INTO the executable, so this is a weaker answer than ' +
        'it looks.')
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
# Tab-separated, and formatted here. PowerShell re-parses arguments on their way to a native
# command, so a jq expression containing double quotes arrives as eight arguments instead of one.
$assets = gh release view $Tag --json assets --jq '.assets[] | [.name, (.size / 1048576 | floor | tostring)] | @tsv'
foreach ($asset in $assets) {
    $fields = $asset -split "`t"
    Write-Host ("  {0}  {1} MB" -f $fields[0], $fields[1])
}
