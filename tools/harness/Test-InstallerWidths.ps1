<#
.SYNOPSIS
    Installs xlide for real, checks that 64-bit and 32-bit processes are each handed their own
    library, and uninstalls it again. For a machine that is thrown away afterwards.

.DESCRIPTION
    The installer puts a library of each width in one folder and registers each with the Office
    of that width. Whether that works is a fact about a registry and a file system, so it is
    checked on a real one:

      1. run the installer, silently;
      2. see that the files and the registry keys of both widths are there;
      3. ask COM for the add-in by its ProgID from a 64-bit and from a 32-bit process, the way
         the editor asks, and see which file each was given (Test-ShimLoad.ps1 -Registered);
      4. run the uninstall, and see that none of it is left.

    THIS REALLY INSTALLS. It replaces whatever xlide the machine has, and it rewrites the add-in
    registration, which on a developer's machine is the registration the dev loop depends on.
    So it runs only where GITHUB_ACTIONS says the machine is a CI runner, unless -Force says
    this machine is one to do it to anyway.

    No Office is needed and none is started. That an Office of each width then loads what it is
    given is not something this shows.

.EXAMPLE
    tools\harness\Test-InstallerWidths.ps1 -Installer artifacts\xlide-setup.exe
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $Installer,

    # Run on a machine that is not a CI runner. It will replace the installation and the
    # registration this machine has.
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($env:GITHUB_ACTIONS -ne 'true' -and -not $Force) {
    throw ('This installs xlide for real and rewrites its registration, so it is for a CI runner. ' +
        'Pass -Force to run it on this machine anyway.')
}

$setup = (Resolve-Path $Installer).Path
$install = Join-Path $env:LOCALAPPDATA 'Programs\xlide'
$class = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'
$progId = 'Xlide.VbeAddIn'

# What an installation is made of, for both widths: the files, the key each editor lists its
# add-ins under, and the class as a process of each width resolves it.
$parts = [ordered] @{
    'the 64-bit library'       = Join-Path $install 'Xlide.Vbe.Shim.dll'
    'the 64-bit loader'        = Join-Path $install 'WebView2Loader.dll'
    'the 32-bit library'       = Join-Path $install 'x86\Xlide.Vbe.Shim.dll'
    'the 32-bit loader'        = Join-Path $install 'x86\WebView2Loader.dll'
    'the editor page'          = Join-Path $install 'ui\editor\dist\index.html'
    'the 64-bit add-in key'    = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"
    'the 32-bit add-in key'    = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\$progId"
    'the class, 64-bit view'   = "HKCU:\Software\Classes\CLSID\$class"
    'the class, 32-bit view'   = "HKCU:\Software\Classes\WOW6432Node\CLSID\$class"
    'the ProgID'               = "HKCU:\Software\Classes\$progId"
}

function Invoke-Setup([string[]] $Arguments) {
    # Through Start-Process, so the exit code is the installer's own whatever it wrote.
    $run = Start-Process -FilePath $setup -ArgumentList $Arguments -NoNewWindow -PassThru -Wait
    return $run.ExitCode
}

Write-Host "==> Installing from $setup" -ForegroundColor Cyan
$code = Invoke-Setup @('--silent')
if ($code -ne 0) { throw "The installer answered $code." }

$missing = @($parts.Keys | Where-Object { -not (Test-Path -LiteralPath $parts[$_]) })
if ($missing.Count -gt 0) { throw "Not there after the install: $($missing -join ', ')." }
Write-Host "    all $($parts.Count) parts of both widths are there"

foreach ($view in 'CLSID', 'WOW6432Node\CLSID') {
    $server = (Get-Item -LiteralPath "HKCU:\Software\Classes\$view\$class\InprocServer32").GetValue('')
    Write-Host "    $view resolves to $server"
}

Write-Host '==> Asking COM for the add-in from a process of each width' -ForegroundColor Cyan
& (Join-Path $PSScriptRoot 'Test-ShimLoad.ps1') -Registered -Configuration Release -InstallFolder $install
$asked = $LASTEXITCODE

# Uninstalled whatever the check said, so a failed run does not leave the runner installed for
# the steps after it.
Write-Host '==> Uninstalling' -ForegroundColor Cyan
$code = Invoke-Setup @('--uninstall', '--silent')
if ($code -ne 0) { throw "The uninstall answered $code." }

$left = @($parts.Keys | Where-Object { Test-Path -LiteralPath $parts[$_] })
if (Test-Path -LiteralPath $install) { $left += 'the installation folder' }
if ($left.Count -gt 0) { throw "Left behind after the uninstall: $($left -join ', ')." }
Write-Host '    nothing of either width is left'

if ($asked -ne 0) { throw 'A process was not handed the library installed for its width. See above.' }

Write-Host ''
Write-Host 'RESULT: passed' -ForegroundColor Green
exit 0
