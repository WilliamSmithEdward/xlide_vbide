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
$shimPublish32 = Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x86"
$enginePublish = Join-Path $repoRoot 'engine\dist'
# Only the built page ships, not the folder it is built in. WebViewPaths.EditorContentRelativePath
# is what the shim looks for beside itself, and it is this and nothing else: staging the whole ui
# tree put ui\editor\node_modules -- monaco's sources, esbuild's binaries, typescript -- inside the
# installer, which is 146 MB of build tooling that no user runs.
$uiSource = Join-Path $repoRoot 'ui\editor\dist'
$uiRelative = 'ui\editor\dist'

$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path (Join-Path $localDotnet 'dotnet.exe')) {
    $env:PATH = "$localDotnet;$env:PATH"

    # A built program launches through a small native host that locates the runtime through the
    # registry or this variable, not through PATH. Without it the payload packer reports that .NET
    # is not installed, on a machine whose only runtime lives under the user's profile.
    $env:DOTNET_ROOT = $localDotnet
}

$vsInstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer"
if (Test-Path $vsInstaller) { $env:PATH = "$vsInstaller;$env:PATH" }

$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

<#
.SYNOPSIS
    The machine a Windows library was built for, read from its own header: 0x8664 is 64-bit,
    0x014C is 32-bit.
#>
function Get-LibraryMachine([string] $Path) {
    # Resolved here: .NET opens a relative path against the process's folder, not PowerShell's.
    $Path = (Resolve-Path -LiteralPath $Path).Path
    $stream = [IO.File]::OpenRead($Path)
    try {
        $header = New-Object byte[] 4096
        $read = $stream.Read($header, 0, $header.Length)
    } finally {
        $stream.Dispose()
    }

    $at = [BitConverter]::ToInt32($header, 0x3C)
    if ($at -le 0 -or ($at + 6) -gt $read -or $header[$at] -ne 0x50 -or $header[$at + 1] -ne 0x45) {
        throw "$Path is not a Windows library."
    }

    return [BitConverter]::ToUInt16($header, $at + 4)
}

if (-not $NoBuildProduct) {
    # One of each width. A host loads only a library of its own, and the installer registers
    # both, so an installer with one missing would leave one kind of Office without xlide.
    foreach ($runtime in 'win-x64', 'win-x86') {
        Write-Host "==> Publishing the shim ($runtime)" -ForegroundColor Cyan
        dotnet publish (Join-Path $repoRoot 'src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj') -c $Configuration -r $runtime
        if ($LASTEXITCODE -ne 0) { throw "Publishing the shim for $runtime failed." }
    }
}

Write-Host '==> Staging the payload' -ForegroundColor Cyan
if (Test-Path $payloadDir) { Remove-Item $payloadDir -Recurse -Force }
New-Item -ItemType Directory -Path $payloadDir | Out-Null

# Only runtime files belong in the payload. Debug symbols would more than triple its size.
$shipped = 0
foreach ($file in (Get-ChildItem $shimPublish -File -ErrorAction SilentlyContinue)) {
    if ($file.Extension -in @('.pdb', '.lib', '.exp')) { continue }

    # A COPY SET ASIDE BY HAND IS NOT A RUNTIME FILE. When a running host holds the shim or the
    # engine, the way past it is to rename the locked file aside and publish a fresh one - after
    # which a full-size copy sits in the publish folder wearing a name nothing looks for. This
    # loop shipped whatever it found, so such a copy would ride into the installer in silence.
    if ($file.Name -match '\.held-\d+\.') {
        Write-Host ("    skipped a set-aside copy: {0}" -f $file.Name) -ForegroundColor Yellow
        continue
    }
    Copy-Item $file.FullName (Join-Path $payloadDir $file.Name)
    Write-Host ("    {0} ({1:N0} KB)" -f $file.Name, ($file.Length / 1KB))
    $shipped++
}

if ($shipped -eq 0) { throw "Nothing to package: no published files under $shimPublish." }

# THE 32-BIT SHIM, for a 32-bit Office. Two files and nothing else: the library, and the browser
# loader, which has to be the width of the library that loads it. The page, the engine and the
# icon are staged once, with the 64-bit shim, and the 32-bit one looks one folder up for them.
# The folder is InstallLayout.X86FolderName, and the installer registers the library it finds
# there, so its name here and there has to be the same word.
$x86Dir = Join-Path $payloadDir 'x86'
New-Item -ItemType Directory -Path $x86Dir | Out-Null
foreach ($name in 'Xlide.Vbe.Shim.dll', 'WebView2Loader.dll') {
    $source = Join-Path $shimPublish32 $name
    if (-not (Test-Path $source)) {
        throw "No 32-bit $name at $source, so this installer would leave a 32-bit Office without xlide. Run without -NoBuildProduct."
    }

    Copy-Item $source (Join-Path $x86Dir $name)
    Write-Host ("    x86\{0} ({1:N0} KB)" -f $name, ((Get-Item $source).Length / 1KB))
}

# EACH LIBRARY IS THE WIDTH ITS FOLDER SAYS. The two publish folders differ by one word and the
# four files by none, so a mix-up is a copy that succeeds and an add-in that never loads: Windows
# refuses a library of the other width and the editor says nothing at all. The width is read out
# of each file, since nothing else here would notice.
foreach ($expected in @(
        @{ Folder = $payloadDir; Machine = 0x8664; Says = '64-bit' },
        @{ Folder = $x86Dir; Machine = 0x014C; Says = '32-bit' })) {
    foreach ($name in 'Xlide.Vbe.Shim.dll', 'WebView2Loader.dll') {
        $file = Join-Path $expected.Folder $name
        $machine = Get-LibraryMachine $file
        if ($machine -ne $expected.Machine) {
            throw ("{0} should be a {1} library and its header says machine 0x{2:X4}." -f $file, $expected.Says, $machine)
        }
    }
}
Write-Host '    each library is the width its folder says'

# THE ENGINE IS TAKEN BY NAME, not by extension. `*.exe` ships whatever executable happens to be
# in the folder, and on 2026-09-13 that was `xlide-engine.held-130938.exe` - a 98 MB copy set
# aside days earlier when a running Excel held the real one. It went into v0.16.2's installer at
# 23.6 MB compressed, to be installed beside the engine on every machine that took that build,
# and nothing in the packaging said a word about it. A second executable here is somebody's
# leftover; it is named out loud and left behind.
$engineName = 'xlide-engine.exe'
$engine = @()
if (Test-Path $enginePublish) {
    $engine = @(Get-ChildItem $enginePublish -File -Filter $engineName -ErrorAction SilentlyContinue)

    $strays = @(Get-ChildItem $enginePublish -File -Filter '*.exe' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ne $engineName })
    if ($strays.Count -gt 0) {
        Write-Host ("    left behind, not the engine: {0}" -f (($strays | ForEach-Object { $_.Name }) -join ', ')) -ForegroundColor Yellow
    }
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

Write-Host '==> Compressing the payload' -ForegroundColor Cyan
# Built and invoked rather than "dotnet run", which forwards its own arguments to the program and
# makes the program's argument list depend on the SDK's parsing rather than on what was written here.
dotnet build (Join-Path $repoRoot 'tools\Xlide.Payload.Pack\Xlide.Payload.Pack.csproj') -c Release --nologo -v q
if ($LASTEXITCODE -ne 0) { throw 'Building the payload packer failed.' }

$packer = Join-Path $repoRoot 'artifacts\bin\Xlide.Payload.Pack\release\xlide-payload-pack.exe'
if (-not (Test-Path $packer)) { throw "The payload packer was not produced at $packer." }

& $packer $payloadDir (Join-Path $repoRoot 'artifacts\payload-pack-cache')
if ($LASTEXITCODE -ne 0) { throw 'Compressing the payload failed.' }

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
