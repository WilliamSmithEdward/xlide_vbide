<#
.SYNOPSIS
    Builds xlide-setup.exe, a single self-contained installer.

.DESCRIPTION
    Publishes the product, stages it as the installer's payload, then compiles the installer ahead
    of time so the result is one native executable that depends on nothing already being present on
    the machine.
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',

    # Skip rebuilding the product and use whatever was published last.
    [switch] $NoBuildProduct
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$setupProject = Join-Path $PSScriptRoot 'Xlide.Setup\Xlide.Setup.csproj'
$payloadDir = Join-Path $PSScriptRoot 'Xlide.Setup\payload'
$shimPublish = Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64"
$enginePublish = Join-Path $repoRoot 'engine\dist'
$uiSource = Join-Path $repoRoot 'ui'

$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path (Join-Path $localDotnet 'dotnet.exe')) { $env:PATH = "$localDotnet;$env:PATH" }

$vsInstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer"
if (Test-Path $vsInstaller) { $env:PATH = "$vsInstaller;$env:PATH" }

$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

if (-not $NoBuildProduct) {
    Write-Host '==> Publishing the shim' -ForegroundColor Cyan
    dotnet publish (Join-Path $repoRoot 'src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj') -c $Configuration -r win-x64
    if ($LASTEXITCODE -ne 0) { throw 'Publishing the shim failed.' }
}

Write-Host '==> Staging the payload' -ForegroundColor Cyan
if (Test-Path $payloadDir) { Remove-Item $payloadDir -Recurse -Force }
New-Item -ItemType Directory -Path $payloadDir | Out-Null

# Only runtime files belong in the payload. Debug symbols would more than triple its size.
$shipped = 0
foreach ($file in (Get-ChildItem $shimPublish -File -ErrorAction SilentlyContinue)) {
    if ($file.Extension -in @('.pdb', '.lib', '.exp')) { continue }
    Copy-Item $file.FullName (Join-Path $payloadDir $file.Name)
    Write-Host ("    {0} ({1:N0} KB)" -f $file.Name, ($file.Length / 1KB))
    $shipped++
}

if ($shipped -eq 0) { throw "Nothing to package: no published files under $shimPublish." }

if (Test-Path $enginePublish) {
    $engine = Get-ChildItem $enginePublish -File -Filter '*.exe' -ErrorAction SilentlyContinue
    foreach ($file in $engine) {
        Copy-Item $file.FullName (Join-Path $payloadDir $file.Name)
        Write-Host ("    {0} ({1:N0} KB)" -f $file.Name, ($file.Length / 1KB))
    }
}
else {
    Write-Host '    (no engine build yet)' -ForegroundColor DarkGray
}

if (Test-Path $uiSource) {
    $uiTarget = Join-Path $payloadDir 'ui'
    Copy-Item $uiSource $uiTarget -Recurse -Force
    $uiCount = @(Get-ChildItem $uiTarget -Recurse -File).Count
    Write-Host "    ui ($uiCount file(s))"
}
else {
    Write-Host '    (no ui assets yet)' -ForegroundColor DarkGray
}

Write-Host '==> Building the installer' -ForegroundColor Cyan
dotnet publish $setupProject -c $Configuration -r win-x64
if ($LASTEXITCODE -ne 0) { throw 'Building the installer failed.' }

$setupExe = Join-Path $repoRoot "artifacts\publish\Xlide.Setup\$($Configuration.ToLowerInvariant())_win-x64\xlide-setup.exe"
if (-not (Test-Path $setupExe)) { throw "The installer was not produced at $setupExe." }

$output = Join-Path $repoRoot 'artifacts\xlide-setup.exe'
Copy-Item $setupExe $output -Force

$size = (Get-Item $output).Length / 1MB
Write-Host ''
Write-Host ("Installer: {0} ({1:N2} MB)" -f $output, $size) -ForegroundColor Green
Write-Host 'Install with a double click, or "xlide-setup.exe --silent". Remove with "--uninstall".'
