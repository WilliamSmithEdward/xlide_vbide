<#
.SYNOPSIS
    Registers a development build of the add-in for the current user.

.DESCRIPTION
    Assumes nothing exists. Every key and value is created, including the VBA\VBE\6.0\Addins64
    branch, which is absent on a machine where no add-in has ever registered per user.

    It writes through the same registry API the installer uses rather than the PowerShell registry
    provider, and it reads every value back afterwards. The old version reported success whatever
    happened, so a write that went nowhere still printed "Registered for the current user" and the
    keys were simply absent in regedit.

    Which build it registers is inferred: the dev loop publishes Release by default but is routinely
    run as -Configuration Debug, since the debug api only exists there. It keeps whichever is
    already registered, otherwise takes the most recent publish, and -Configuration overrides both.

    RUN THIS FROM YOUR OWN TERMINAL, not from an agent or sandboxed shell. Sandboxed environments
    can virtualize registry writes, leaving a registration that only the sandbox and its children
    can see while Excel launched normally sees nothing. That mirage cost this project a full
    diagnostic day, and the verification below is what would now catch it.

.EXAMPLE
    tools\Register-DevShim.ps1

.EXAMPLE
    tools\Register-DevShim.ps1 -Configuration Debug
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration,

    # Register a specific file instead of a published dev build.
    [string] $Shim
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$progId = 'Xlide.VbeAddIn'
$clsid = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'

$classesPath = 'Software\Classes'
$inprocPath = "$classesPath\CLSID\$clsid\InprocServer32"
$addinPath = "Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"

function Get-ShimPath([string] $configuration) {
    Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($configuration.ToLowerInvariant())_win-x64\Xlide.Vbe.Shim.dll"
}

if (-not $Shim) {
    if (-not $Configuration) {
        $current = [Microsoft.Win32.Registry]::GetValue("HKEY_CURRENT_USER\$inprocPath", '', $null)

        foreach ($candidate in 'Debug', 'Release') {
            if ($current -and $current -eq (Get-ShimPath $candidate)) { $Configuration = $candidate; break }
        }

        if (-not $Configuration) {
            # Debug first, and not by timestamp. The gate publishes Release on every run, so the
            # most recently written shim is almost always Release even on a machine that has been
            # developing against Debug all day. Debug is also the one carrying the debug api, which
            # is what dev testing is for.
            $published = 'Debug', 'Release' |
                ForEach-Object { [pscustomobject]@{ Name = $_; Path = (Get-ShimPath $_) } } |
                Where-Object { Test-Path $_.Path }

            if (-not $published) {
                throw 'No published shim in either configuration. Run tools\dev.ps1 -NoRun first, then this.'
            }

            $Configuration = $published[0].Name
        }
    }

    $Shim = Get-ShimPath $Configuration
}

$Shim = [System.IO.Path]::GetFullPath($Shim)
if (-not (Test-Path $Shim)) { throw "No shim at $Shim. Run tools\dev.ps1 -NoRun first." }

# CreateSubKey creates every missing level, so a machine with no VBA\VBE branch at all is fine.
function Set-Key([string] $path, [hashtable] $values) {
    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($path, $true)
    if (-not $key) { throw "Could not create HKCU\$path." }

    try {
        foreach ($name in $values.Keys) {
            $value = $values[$name]
            if ($value -is [int]) {
                $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::DWord)
            } else {
                $key.SetValue($name, $value, [Microsoft.Win32.RegistryValueKind]::String)
            }
        }
    } finally {
        $key.Dispose()
    }
}

Write-Host "Registering $Shim"

Set-Key $classesPath @{}
Set-Key "$classesPath\$progId" @{ '' = 'xlide' }
Set-Key "$classesPath\$progId\CLSID" @{ '' = $clsid }
Set-Key "$classesPath\CLSID\$clsid" @{ '' = 'xlide' }
Set-Key $inprocPath @{ '' = $Shim; 'ThreadingModel' = 'Apartment' }
Set-Key $addinPath @{
    'Description'      = 'Modern VBA development inside the Visual Basic Editor.'
    'FriendlyName'     = 'xlide'
    'LoadBehavior'     = 3
    'CommandLineSafe'  = 0
}

# Read back rather than trust the writes. A virtualized hive accepts every write above and shows
# nothing to Excel, and the whole point of this script is knowing which of those two happened.
$checks = @(
    @{ Path = $inprocPath; Name = ''; Expected = $Shim }
    @{ Path = $inprocPath; Name = 'ThreadingModel'; Expected = 'Apartment' }
    @{ Path = "$classesPath\$progId\CLSID"; Name = ''; Expected = $clsid }
    @{ Path = $addinPath; Name = 'LoadBehavior'; Expected = 3 }
    @{ Path = $addinPath; Name = 'FriendlyName'; Expected = 'xlide' }
)

$failed = @()
foreach ($check in $checks) {
    $actual = [Microsoft.Win32.Registry]::GetValue("HKEY_CURRENT_USER\$($check.Path)", $check.Name, $null)
    $label = if ($check.Name) { "$($check.Path)\$($check.Name)" } else { "$($check.Path)\(default)" }

    if ($actual -ne $check.Expected) {
        $failed += "  $label is '$actual', expected '$($check.Expected)'"
    }
}

if ($failed.Count) {
    Write-Host ''
    Write-Host 'The registration did not take:' -ForegroundColor Red
    $failed | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    throw 'Registry writes did not survive a read back. If this ran inside an agent or sandboxed shell, run it from your own terminal instead.'
}

Write-Host ''
Write-Host 'Registered for the current user, verified by reading it back.' -ForegroundColor Green
Write-Host ''

# Printed as read, not as intended. This output is the evidence: regedit caches a branch it has
# already drawn and will show the old, empty Addins64 until it is refreshed with F5.
foreach ($check in $checks) {
    $actual = [Microsoft.Win32.Registry]::GetValue("HKEY_CURRENT_USER\$($check.Path)", $check.Name, $null)
    $label = if ($check.Name) { $check.Name } else { '(default)' }
    Write-Host ("  HKCU\{0}" -f $check.Path)
    Write-Host ("      {0,-16} {1}" -f $label, $actual)
}

Write-Host ''
Write-Host 'If regedit still shows an empty Addins64, press F5 in it: it caches what it has drawn.'
Write-Host 'Start Excel and press Alt+F11.'
