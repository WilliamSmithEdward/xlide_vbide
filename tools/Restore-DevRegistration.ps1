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

    The data folder is kept by default. The installed product and the dev build share it, so a
    plain uninstall would take the dev logs and the debug api's discovery files with it.

    RUN THIS FROM YOUR OWN TERMINAL, not from an agent or sandboxed shell. Sandboxed environments
    can virtualize registry writes, leaving a registration only the sandbox can see while Excel
    launched normally sees nothing. That mirage cost this project a full diagnostic day.

.EXAMPLE
    tools\Restore-DevRegistration.ps1

.EXAMPLE
    tools\Restore-DevRegistration.ps1 -LeaveInstalled
    Re-registers the dev shim but leaves the installed product's files in place.

.EXAMPLE
    tools\Restore-DevRegistration.ps1 -DiscardData
    Also removes %LOCALAPPDATA%\xlide_vbide, which is shared with the dev build.
#>
[CmdletBinding()]
param(
    # Leave the installed product's files alone and only take the registration back.
    [switch] $LeaveInstalled,

    # Let the uninstaller remove the shared data folder as well: logs, the WebView2 cache, and the
    # debug api's discovery files.
    [switch] $DiscardData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$progId = 'Xlide.VbeAddIn'
$clsid = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'
$devShim = Join-Path $repoRoot 'artifacts\publish\Xlide.Vbe.Shim\release_win-x64\Xlide.Vbe.Shim.dll'
$installFolder = Join-Path $env:LOCALAPPDATA 'Programs\xlide'
$uninstaller = Join-Path $installFolder 'xlide-setup.exe'

# Both the installer and the registration refuse to touch anything while the editor holds the
# add-in open, and a half-applied registration is worse than none.
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if ($excel) {
    throw "Excel is running (process $($excel.Id)). Close it and run this again."
}

if (-not (Test-Path $devShim)) {
    throw "No published dev shim at $devShim. Run tools\dev.ps1 -NoRun first, then this."
}

if ($LeaveInstalled) {
    Write-Host '==> Leaving the installed product in place' -ForegroundColor DarkGray
} elseif (Test-Path $uninstaller) {
    Write-Host '==> Removing the installed product' -ForegroundColor Cyan

    $arguments = @('--uninstall', '--silent')
    if (-not $DiscardData) { $arguments += '--keep-data' }

    & $uninstaller @arguments
    if ($LASTEXITCODE -ne 0) { throw "The uninstaller exited with $LASTEXITCODE; nothing was changed." }

    # It relaunches itself from a temporary copy so it can delete its own folder, so the work
    # finishes after the process we started has already returned.
    $deadline = (Get-Date).AddSeconds(30)
    while ((Test-Path $installFolder) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }

    if (Test-Path $installFolder) {
        Write-Host "    $installFolder is still present; the registration is taken back regardless." -ForegroundColor Yellow
    }
} elseif (Test-Path $installFolder) {
    Write-Host "==> $installFolder exists but carries no uninstaller; leaving it" -ForegroundColor Yellow
} else {
    Write-Host '==> Nothing installed' -ForegroundColor DarkGray
}

Write-Host '==> Registering the dev shim' -ForegroundColor Cyan
# Delegated rather than copied. One description of the registry layout is the whole point: a second
# copy is a second thing to keep true.
& (Join-Path $PSScriptRoot 'Register-DevShim.ps1')

Write-Host ''
Write-Host '==> Verifying' -ForegroundColor Cyan

$inproc = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"
$addin = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"

$registered = (Get-ItemProperty -Path $inproc -ErrorAction SilentlyContinue).'(default)'
$loadBehavior = (Get-ItemProperty -Path $addin -ErrorAction SilentlyContinue).LoadBehavior

Write-Host "    server        $registered"
Write-Host "    LoadBehavior  $loadBehavior"

if ($registered -ne $devShim) {
    throw "The registered server is $registered, not the dev shim at $devShim."
}

# The editor rewrites this to 0 when a load fails, and a 0 here is why an add-in that was working
# yesterday is absent today with nothing in the log to say so.
if ($loadBehavior -ne 3) {
    throw "LoadBehavior is $loadBehavior, not 3. The editor turns it off after a failed load; set it back and check the shim log."
}

Write-Host ''
Write-Host 'The dev build is registered. Start Excel and press Alt+F11.' -ForegroundColor Green
