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

# What the shim actually spawns in a build tree. Not the bundle beside it: see the build step.
$enginePath = Join-Path $repoRoot 'engine\dist\xlide-engine.exe'

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

<#
.SYNOPSIS
    Waits until the published shim can be replaced, naming whoever is holding it.

.DESCRIPTION
    A host holds the add-in library open for its entire life, and an Excel that has just
    been asked to close keeps its grip for seconds afterwards. Publishing into that window
    fails partway through and looks like a compiler problem, which is how a stale DLL gets
    tested three times in a row while the fix sits undeployed.

    Writability is tested the only way that is not a guess: by opening the file for exclusive
    writing. Excel processes are named while waiting, because "something has it" is not a
    thing anyone can act on.
#>
function Wait-ForShimUnlocked {
    param([string] $Path, [int] $TimeoutSeconds = 30)

    if (-not (Test-Path $Path)) {
        return
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $announced = $false

    while ((Get-Date) -lt $deadline) {
        try {
            $stream = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
            $stream.Close()
            if ($announced) { Write-Host 'The shim is free; publishing.' -ForegroundColor Green }
            return
        } catch {
            if (-not $announced) {
                $holders = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
                $who = if ($holders.Count -gt 0) { "Excel $($holders -join ', ')" } else { 'another process' }
                Write-Host "Waiting for $who to release the shim..." -ForegroundColor Yellow
                $announced = $true
            }

            Start-Sleep -Milliseconds 500
        }
    }

    $stillThere = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $named = if ($stillThere.Count -gt 0) { " Excel $($stillThere -join ', ') is running." } else { '' }
    throw "The shim at $Path is still locked after $TimeoutSeconds seconds.$named Close the host and run again; publishing over a locked file leaves a stale build behind."
}

if ($Unregister) {
    Invoke-Step 'Unregister' {
        dotnet run --project $registerProject -c Debug -- --remove
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

    Invoke-Step 'Build the engine (bundle, then executable)' {
        # PACKAGED, not just bundled. `node build.mjs` writes engine.cjs; the shim runs
        # xlide-engine.exe, which only `--package` refreshes. Building without it leaves the
        # executable at whatever it was and every measurement afterwards describes the OLD
        # engine while the source says otherwise (2026-08-08: a 14x improvement measured as no
        # improvement at all, twice, before the timestamps were read). The staleness assertion
        # below is what makes that impossible to miss again.
        Push-Location (Join-Path $repoRoot 'engine')
        try {
            # esbuild and postject write their progress to STDERR even when they succeed, and
            # under $ErrorActionPreference = 'Stop' Windows PowerShell turns a native command's
            # stderr into a terminating error. So the preference is lifted for the call and the
            # EXIT CODE is what decides, which is the only thing that actually knows.
            $wasPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try {
                node build.mjs --package
            } finally {
                $ErrorActionPreference = $wasPreference
            }

            if ($LASTEXITCODE -ne 0) { throw "The engine build failed ($LASTEXITCODE)." }
        } finally {
            Pop-Location
        }

        # The analyzer counts as engine source: it lives in the neighbouring checkout and is
        # bundled INTO this executable, so a pull over there changes what the add-in runs without
        # touching a file in this repository. The gate watches both; so does this.
        $engineSources = @(Join-Path $repoRoot 'engine\src')
        $analyzerSources = Join-Path (Split-Path -Parent $repoRoot) 'xlide_vscode\src'
        if (Test-Path $analyzerSources) { $engineSources += $analyzerSources }

        $newestEngineSource = Get-ChildItem $engineSources -Recurse -Include *.ts |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $engineExe = Get-Item $enginePath -ErrorAction SilentlyContinue
        if (-not $engineExe) { throw "No engine executable at $enginePath after the build." }
        if ($newestEngineSource -and $newestEngineSource.LastWriteTime -gt $engineExe.LastWriteTime) {
            throw ("The packaged engine is OLDER than $($newestEngineSource.Name) " +
                "($($engineExe.LastWriteTime.ToString('HH:mm:ss')) against " +
                "$($newestEngineSource.LastWriteTime.ToString('HH:mm:ss'))). The package did not " +
                'take; do not test this build.')
        }

        Write-Host ("Engine: {0} ({1:N2} MB, built {2:HH:mm:ss})" -f
            $engineExe.FullName, ($engineExe.Length / 1MB), $engineExe.LastWriteTime)
    }

    Invoke-Step 'Publish the shim (ahead-of-time, native)' {
        # A host holds the shim open for its whole life, and Excel takes seconds to let go
        # after it is asked to close. Publishing into a locked file fails in the middle of
        # its own output, and the failure reads like a build error rather than what it is -
        # or worse, an earlier step's output scrolls past and the stale DLL is what gets
        # tested (2026-08-06: three rounds of "why is my fix not in the log", the answer
        # being that it was never deployed). So the lock is waited out, by name, first.
        Wait-ForShimUnlocked -Path $shimPath -TimeoutSeconds 30
        dotnet publish $shimProject -c $Configuration -r win-x64
    }
}

if (-not (Test-Path $shimPath)) {
    throw "No published shim at $shimPath. Run without -NoBuild."
}

$shimInfo = Get-Item $shimPath
Write-Host ''
Write-Host ("Shim: {0} ({1:N2} MB, built {2:HH:mm:ss})" -f $shimInfo.FullName, ($shimInfo.Length / 1MB), $shimInfo.LastWriteTime)

# A shim older than the newest source it should contain is a stale deploy wearing a fresh
# report, and every minute spent reading a log for a fix that was never deployed is spent
# because nothing said so here (2026-08-06).
if (-not $NoBuild) {
    $newestSource = Get-ChildItem (Join-Path $repoRoot 'src') -Recurse -Include *.cs |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newestSource -and $newestSource.LastWriteTime -gt $shimInfo.LastWriteTime) {
        throw ("The published shim is OLDER than $($newestSource.Name) " +
            "($($shimInfo.LastWriteTime.ToString('HH:mm:ss')) against " +
            "$($newestSource.LastWriteTime.ToString('HH:mm:ss'))). The publish did not take; " +
            'do not test this build.')
    }
}

Invoke-Step 'Register for the current user' {
    # Debug, for the same reason the test gate is (see the note above it): Smart App Control
    # blocks a FRESHLY BUILT unsigned RELEASE managed assembly from loading (0x800711C7), and
    # this tool is managed. It ran on a stale Release build until its sources changed, then
    # every dev loop died at registration with the shim already published (2026-08-05).
    dotnet run --project $registerProject -c Debug -- --apply --shim $shimPath
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
