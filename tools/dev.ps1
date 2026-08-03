<#
.SYNOPSIS
    Build, register, and run the add-in inside a real editor in one command.

.EXAMPLE
    tools\dev.ps1
    Build, register, and launch. The usual loop.

.EXAMPLE
    tools\dev.ps1 -KeepOpen
    Leave Excel running to work in the editor by hand.

.EXAMPLE
    tools\dev.ps1 -Unregister
    Remove the registration and leave the machine clean.
#>
[CmdletBinding()]
param(
    # Skip the build and use whatever was published last.
    [switch] $NoBuild,

    # Register and build, but do not launch Excel.
    [switch] $NoRun,

    # Leave Excel running after the check.
    [switch] $KeepOpen,

    # Reuse the Excel instance a previous -KeepOpen run left open. Only valid when the shim itself
    # has not changed, because a host holds an add-in library open for its lifetime.
    [switch] $Reuse,

    # Remove the registration and exit.
    [switch] $Unregister,

    # Build configuration.
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$shimProject = Join-Path $repoRoot 'src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj'
$registerProject = Join-Path $repoRoot 'tools\Xlide.Vbe.Register\Xlide.Vbe.Register.csproj'
$publishDir = Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64"
$shimPath = Join-Path $publishDir 'Xlide.Vbe.Shim.dll'

# The SDK and the native linker are not necessarily on the machine PATH.
$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path (Join-Path $localDotnet 'dotnet.exe')) {
    $env:PATH = "$localDotnet;$env:PATH"
}

$vsInstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer"
if (Test-Path $vsInstaller) {
    # Ahead-of-time publishing locates the native linker through vswhere.
    $env:PATH = "$vsInstaller;$env:PATH"
}

$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'

function Invoke-Step {
    param([string] $Name, [scriptblock] $Body)

    Write-Host ''
    Write-Host "==> $Name" -ForegroundColor Cyan
    & $Body
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
        throw "$Name failed with exit code $LASTEXITCODE."
    }
}

if ($Unregister) {
    Invoke-Step 'Unregister' {
        dotnet run --project $registerProject -c Release -- --remove
    }

    Write-Host ''
    Write-Host 'The add-in is no longer registered.' -ForegroundColor Green
    exit 0
}

if (-not $NoBuild) {
    Invoke-Step 'Test' {
        # Debug deliberately, whatever the publish configuration: Smart App Control (observed
        # ON 2026-08-02) blocks freshly built unsigned RELEASE test assemblies from running
        # (0x800711C7, surfacing as xUnit's "did not return valid JSON"), while Debug ones and
        # the NativeAOT shim itself load fine. The gate is about correctness, not codegen.
        dotnet test (Join-Path $repoRoot 'tests\Xlide.Vbe.Core.Tests') -c Debug --nologo
    }

    Invoke-Step 'Publish the shim (ahead-of-time, native)' {
        dotnet publish $shimProject -c $Configuration -r win-x64
    }
}

if (-not (Test-Path $shimPath)) {
    throw "No published shim at $shimPath. Run without -NoBuild."
}

$shimInfo = Get-Item $shimPath
Write-Host ''
Write-Host ("Shim: {0} ({1:N2} MB, built {2:HH:mm:ss})" -f $shimInfo.FullName, ($shimInfo.Length / 1MB), $shimInfo.LastWriteTime)

Invoke-Step 'Register for the current user' {
    dotnet run --project $registerProject -c Release -- --apply --shim $shimPath
}

if ($NoRun) {
    Write-Host ''
    Write-Host 'Registered. Start Excel and open the editor to load the add-in.' -ForegroundColor Green
    exit 0
}

Invoke-Step 'Load into the editor' {
    $checkArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'harness\Invoke-VbeLoadCheck.ps1'))
    if ($KeepOpen) { $checkArgs += '-KeepOpen' }
    if ($Reuse) { $checkArgs += '-Reuse' }

    & powershell @checkArgs
}
