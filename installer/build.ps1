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
    [switch] $NoBuildProduct,

    # Package deliberately without the language engine. The result installs an editor with no
    # analysis -- no diagnostics, no completion, no hover -- and is named so that it cannot be
    # mistaken for the product. Without this switch a missing engine is an error, because an
    # installer that quietly omits half of what it is for is worse than one that refuses to build.
    [switch] $WithoutEngine
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$setupProject = Join-Path $PSScriptRoot 'Xlide.Setup\Xlide.Setup.csproj'
$payloadDir = Join-Path $PSScriptRoot 'Xlide.Setup\payload'
$shimPublish = Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64"
$enginePublish = Join-Path $repoRoot 'engine\dist'
# Only the built page ships, not the folder it is built in. WebViewPaths.EditorContentRelativePath
# is what the shim looks for beside itself, and it is this and nothing else: staging the whole ui
# tree put ui\editor\node_modules -- monaco's sources, esbuild's binaries, typescript -- inside the
# installer, which is 146 MB of build tooling that no user runs.
$uiSource = Join-Path $repoRoot 'ui\editor\dist'
$uiRelative = 'ui\editor\dist'

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

$engine = @()
if (Test-Path $enginePublish) {
    $engine = @(Get-ChildItem $enginePublish -File -Filter '*.exe' -ErrorAction SilentlyContinue)
}

# The engine is not built by this script and is not in source control -- it embeds a language
# runtime and is far too large for history -- so it is entirely possible to arrive here without one
# and produce an installer that looks complete and analyses nothing. Say so and stop.
if ($engine.Count -eq 0 -and -not $WithoutEngine) {
    throw @"
No language engine at $enginePublish, so this installer would carry no analysis at all.
Build it first:
    cd engine; npm run package
Or pass -WithoutEngine to package an editor-only build on purpose.
"@
}

foreach ($file in $engine) {
    Copy-Item $file.FullName (Join-Path $payloadDir $file.Name)
    Write-Host ("    {0} ({1:N0} KB)" -f $file.Name, ($file.Length / 1KB))
}

if ($engine.Count -eq 0) {
    Write-Host '    (no engine: this build installs an editor without analysis)' -ForegroundColor Yellow
}

# The page is not optional: without it the tool window falls back to the native pane, which is not
# the product either. Same reasoning as the engine, so the same refusal.
if (-not (Test-Path (Join-Path $uiSource 'editor.js'))) {
    throw @"
No built editor page at $uiSource, so this installer would show the native pane instead.
Build it first:
    cd ui\editor; npm run build
"@
}

$uiTarget = Join-Path $payloadDir $uiRelative
New-Item -ItemType Directory -Path $uiTarget -Force | Out-Null
Copy-Item (Join-Path $uiSource '*') $uiTarget -Recurse -Force
$uiBytes = (Get-ChildItem $uiTarget -Recurse -File | Measure-Object -Property Length -Sum).Sum
Write-Host ("    {0} ({1} file(s), {2:N1} MB)" -f $uiRelative, @(Get-ChildItem $uiTarget -Recurse -File).Count, ($uiBytes / 1MB))

Write-Host '==> Building the installer' -ForegroundColor Cyan
dotnet publish $setupProject -c $Configuration -r win-x64
if ($LASTEXITCODE -ne 0) { throw 'Building the installer failed.' }

$setupExe = Join-Path $repoRoot "artifacts\publish\Xlide.Setup\$($Configuration.ToLowerInvariant())_win-x64\xlide-setup.exe"
if (-not (Test-Path $setupExe)) { throw "The installer was not produced at $setupExe." }

# An editor-only build carries the difference in its name, so it cannot be handed to someone, or
# uploaded to a release, as though it were the product.
$outputName = if ($engine.Count -eq 0) { 'xlide-setup-editor-only.exe' } else { 'xlide-setup.exe' }
$output = Join-Path $repoRoot "artifacts\$outputName"
Copy-Item $setupExe $output -Force

$size = (Get-Item $output).Length / 1MB
Write-Host ''
Write-Host ("Installer: {0} ({1:N2} MB)" -f $output, $size) -ForegroundColor Green
Write-Host 'Install with a double click, or "xlide-setup.exe --silent". Remove with "--uninstall".'
