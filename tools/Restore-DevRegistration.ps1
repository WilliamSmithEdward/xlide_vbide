<#
.SYNOPSIS
    Puts the development build back in charge after the production installer has been tested.

.DESCRIPTION
    Installing the product repoints the add-in registration at %LOCALAPPDATA%\Programs\xlide, so
    the editor loads the installed copy and the dev loop keeps deploying to a path nothing reads.
    The symptom is a build that reports success and changes nothing on screen.

    This removes the installed product if it is there and hands registration back to the published
    dev shim. The order matters: the uninstaller strips the registration unconditionally, so
    registering first and uninstalling second would undo the repair.

    Which build it restores is inferred rather than assumed. The dev loop publishes Release by
    default but is routinely run as -Configuration Debug, since the debug api only exists there, so
    a script that hardcoded one of them would quietly move you to the other. It keeps whichever is
    registered now, falls back to the most recently published, and -Configuration overrides both.

    The data folder is kept by default. The installed product and the dev build share it, so a
    plain uninstall would take the dev logs and the debug api's discovery files with it.

    RUN THIS FROM YOUR OWN TERMINAL, not from an agent or sandboxed shell. Sandboxed environments
    can virtualize registry writes, leaving a registration only the sandbox can see while Excel
    launched normally sees nothing. That mirage cost this project a full diagnostic day.

.EXAMPLE
    tools\Restore-DevRegistration.ps1

.EXAMPLE
    tools\Restore-DevRegistration.ps1 -Configuration Debug

.EXAMPLE
    tools\Restore-DevRegistration.ps1 -LeaveInstalled
    Re-registers the dev shim but leaves the installed product's files in place.
#>
[CmdletBinding()]
param(
    # Which published build to register. Inferred when omitted.
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration,

    # Leave the installed product's files alone and only take the registration back.
    [switch] $LeaveInstalled,

    # Let the uninstaller remove the shared data folder as well: logs, the WebView2 cache, and the
    # debug api's discovery files.
    [switch] $DiscardData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$clsid = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'
$progId = 'Xlide.VbeAddIn'
$inproc = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
$addin = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"
$installFolder = Join-Path $env:LOCALAPPDATA 'Programs\xlide'
$uninstaller = Join-Path $installFolder 'xlide-setup.exe'
$registerProject = Join-Path $repoRoot 'tools\Xlide.Vbe.Register\Xlide.Vbe.Register.csproj'

function Get-ShimPath([string] $configuration) {
    Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($configuration.ToLowerInvariant())_win-x64\Xlide.Vbe.Shim.dll"
}

# A built program finds its runtime through DOTNET_ROOT or the registry, not through PATH, so a
# runtime living under the user's profile is invisible without this.
$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path (Join-Path $localDotnet 'dotnet.exe')) {
    $env:PATH = "$localDotnet;$env:PATH"
    $env:DOTNET_ROOT = $localDotnet
}

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if ($excel) {
    throw "Excel is running (process $($excel.Id)). Close it and run this again."
}

if (-not $Configuration) {
    # What is registered now is the best evidence of what you were working on.
    $current = if (Test-Path $inproc) { (Get-ItemProperty $inproc).'(default)' } else { $null }

    foreach ($candidate in 'Debug', 'Release') {
        if ($current -and $current -eq (Get-ShimPath $candidate)) { $Configuration = $candidate; break }
    }

    if (-not $Configuration) {
        $published = 'Debug', 'Release' |
            ForEach-Object { [pscustomobject]@{ Name = $_; Path = (Get-ShimPath $_) } } |
            Where-Object { Test-Path $_.Path } |
            Sort-Object { (Get-Item $_.Path).LastWriteTime } -Descending

        if (-not $published) {
            throw 'No published dev shim in either configuration. Run tools\dev.ps1 -NoRun first, then this.'
        }

        $Configuration = $published[0].Name
        Write-Host "==> Nothing of ours is registered; taking the most recent publish ($Configuration)" -ForegroundColor DarkGray
    } else {
        Write-Host "==> Keeping the configuration already registered ($Configuration)" -ForegroundColor DarkGray
    }
}

$devShim = Get-ShimPath $Configuration
if (-not (Test-Path $devShim)) {
    throw "No published dev shim at $devShim. Run tools\dev.ps1 -NoRun -Configuration $Configuration first."
}

if ($LeaveInstalled) {
    Write-Host '==> Leaving the installed product in place' -ForegroundColor DarkGray
} elseif (Test-Path $uninstaller) {
    Write-Host '==> Removing the installed product' -ForegroundColor Cyan

    $arguments = @('--uninstall', '--silent')
    if (-not $DiscardData) { $arguments += '--keep-data' }

    & $uninstaller @arguments
    if ($LASTEXITCODE -ne 0) { throw "The uninstaller exited with $LASTEXITCODE. Nothing was changed." }

    # It relaunches itself from a temporary copy so it can delete its own folder, so the work
    # finishes after the process we started has already returned.
    $deadline = (Get-Date).AddSeconds(30)
    while ((Test-Path $installFolder) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }

    if (Test-Path $installFolder) {
        Write-Host "    $installFolder is still present; taking the registration back regardless." -ForegroundColor Yellow
    }
} elseif (Test-Path $installFolder) {
    Write-Host "==> $installFolder exists but carries no uninstaller; leaving it" -ForegroundColor Yellow
} else {
    Write-Host '==> Nothing installed' -ForegroundColor DarkGray
}

Write-Host "==> Registering the $Configuration shim" -ForegroundColor Cyan
# The same tool the dev loop uses, built Debug for the same reason it is there: Smart App Control
# blocks a freshly built unsigned Release managed assembly from loading, and this tool is managed.
dotnet run --project $registerProject -c Debug -- --apply --shim $devShim
if ($LASTEXITCODE -ne 0) { throw 'Registration failed.' }

Write-Host ''
Write-Host '==> Verifying' -ForegroundColor Cyan

$registered = if (Test-Path $inproc) { (Get-ItemProperty $inproc).'(default)' } else { $null }
$loadBehavior = if (Test-Path $addin) { (Get-ItemProperty $addin).LoadBehavior } else { $null }

Write-Host "    server        $registered"
Write-Host "    LoadBehavior  $loadBehavior"

if ($registered -ne $devShim) {
    throw "The registered server is '$registered', not the dev shim at $devShim."
}

# The editor rewrites this to 0 when a load fails, and a 0 here is why an add-in that worked
# yesterday is absent today with nothing in the log to say so.
if ($loadBehavior -ne 3) {
    throw "LoadBehavior is $loadBehavior, not 3. The editor turns it off after a failed load; set it back and check the shim log."
}

Write-Host ''
Write-Host 'The dev build is registered. Start Excel and press Alt+F11.' -ForegroundColor Green
