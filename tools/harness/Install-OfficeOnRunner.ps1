<#
.SYNOPSIS
    Installs Microsoft 365 Excel of a chosen width on a CI runner, unattended, and leaves it
    ready to be driven.

.DESCRIPTION
    32-bit and 64-bit Office cannot be installed side by side, so the only machine that can
    hold a 32-bit Office beside this product's development is one made for the purpose and
    thrown away: a hosted runner. This puts Excel on one, in the width asked for.

    It does what the Office Deployment Tool does for anyone: the tool is downloaded from
    Microsoft at a pinned address, held to a pinned SHA-256 and to Microsoft's signature, and
    run with a configuration naming the width, the channel, the language and Excel alone.
    Then Excel is opened once through automation and closed, which is what gets its first
    start out of the way, and the two settings the harness needs are set for the user.

    THIS ACCEPTS MICROSOFT'S LICENSE TERMS FOR OFFICE on the machine it runs on: the
    configuration says AcceptEULA, because an unattended install cannot be asked. The Office
    it leaves is not activated. Both are why this runs only where GITHUB_ACTIONS says the
    machine is a CI runner, unless -Force says to do it here anyway.

    The method is the one DecimalTurn/setup-vba uses for 64-bit Office on the same runner
    image, with the Deployment Tool called directly instead of through a Chocolatey package.

    MOVING THE PIN. Nothing follows this download automatically. When Microsoft retires the
    address, take the new one from the Deployment Tool's download page, download it once,
    check its signature, and record its SHA-256 here.

.EXAMPLE
    tools\harness\Install-OfficeOnRunner.ps1 -Width 32
#>
[CmdletBinding()]
param(
    [ValidateSet(32, 64)]
    [int] $Width = 32,

    # Run on a machine that is not a CI runner. It installs Office there and accepts its terms.
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($env:GITHUB_ACTIONS -ne 'true' -and -not $Force) {
    throw ('This installs Microsoft Office and accepts its license terms, so it is for a CI runner. ' +
        'Pass -Force to run it on this machine anyway.')
}

. (Join-Path $PSScriptRoot 'LibraryMachine.ps1')

# The Office Deployment Tool, as Microsoft publishes it. The digest is the one the
# office365proplus Chocolatey package records for the same file.
$toolUrl = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe'
$toolSha256 = '0f9e6df376853154e05d81e2550183ed621ec97fc3f0c290666683a057086b92'

$work = Join-Path $(if ($env:RUNNER_TEMP) { $env:RUNNER_TEMP } else { $env:TEMP }) 'xlide-office'
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Path $work | Out-Null

Write-Host '==> Downloading the Office Deployment Tool' -ForegroundColor Cyan
$tool = Join-Path $work 'officedeploymenttool.exe'
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -Uri $toolUrl -OutFile $tool -UseBasicParsing

$digest = (Get-FileHash $tool -Algorithm SHA256).Hash.ToLowerInvariant()
if ($digest -ne $toolSha256) {
    throw "The Deployment Tool's SHA-256 is $digest and $toolSha256 is pinned. It is not run."
}

$signature = Get-AuthenticodeSignature $tool
if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
    throw "The Deployment Tool is not validly signed by Microsoft ($($signature.Status)). It is not run."
}
Write-Host "    the pinned file, signed by $($signature.SignerCertificate.Subject -replace ',.*$', '')"

$extracted = Join-Path $work 'tool'
$run = Start-Process -FilePath $tool -ArgumentList "/quiet /norestart /extract:`"$extracted`"" -Wait -PassThru
if ($run.ExitCode -ne 0) { throw "The Deployment Tool did not unpack (exit $($run.ExitCode))." }

$setup = Join-Path $extracted 'setup.exe'
if (-not (Test-Path $setup)) { throw "The Deployment Tool unpacked without a setup.exe at $setup." }

# Excel alone, in English, because the harness reads the editor's own words in places.
$excluded = 'Access', 'Bing', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Word'
$configuration = Join-Path $work 'configuration.xml'
$lines = @(
    '<Configuration>'
    "  <Add OfficeClientEdition=`"$Width`" Channel=`"MonthlyEnterprise`">"
    '    <Product ID="O365ProPlusRetail">'
    '      <Language ID="en-us" />'
)
$lines += $excluded | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }
$lines += @(
    '    </Product>'
    '  </Add>'
    '  <Display Level="None" AcceptEULA="TRUE" />'
    '  <Updates Enabled="FALSE" />'
    '</Configuration>'
)
[IO.File]::WriteAllLines($configuration, $lines)
Get-Content $configuration | ForEach-Object { Write-Host "    $_" }

Write-Host "==> Installing $Width-bit Excel" -ForegroundColor Cyan
$started = Get-Date
$run = Start-Process -FilePath $setup -ArgumentList "/configure `"$configuration`"" -Wait -PassThru
if ($run.ExitCode -ne 0) {
    throw "Office setup answered $($run.ExitCode). 30183 is a download from Microsoft's servers that failed."
}
Write-Host ("    installed in {0:N0} seconds" -f ((Get-Date) - $started).TotalSeconds)

$excel = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $excel) { throw 'Setup finished and there is no EXCEL.EXE where Click-to-Run puts one.' }

$installed = Get-LibraryWidth $excel
if ($installed -ne $Width) { throw "$excel is a $installed-bit Excel and $Width-bit was asked for." }
Write-Host "    $excel is $installed-bit, version $((Get-Item $excel).VersionInfo.ProductVersion)"

Write-Host '==> Starting Excel once, so its first start is behind it' -ForegroundColor Cyan
$application = New-Object -ComObject Excel.Application
try { $application.Quit() } catch { Write-Host "    Excel did not quit when asked: $($_.Exception.Message)" }
[Runtime.InteropServices.Marshal]::ReleaseComObject($application) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

foreach ($name in 'EXCEL', 'OfficeClickToRun', 'OfficeC2RClient') {
    Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_ | Stop-Process -Force -ErrorAction Stop; Write-Host "    stopped $($_.Name) $($_.Id)" }
        catch { Write-Host "    could not stop $($_.Name) $($_.Id): $($_.Exception.Message)" }
    }
}
Start-Sleep -Seconds 5

# Macros enabled, because the suites run code; and access to the VBA project, which the
# product never needs and two harness probes do (docs/decisions.md, decision 10).
Write-Host '==> Settings the harness needs' -ForegroundColor Cyan
$security = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
New-Item -Path $security -Force | Out-Null
Set-ItemProperty -Path $security -Name 'VBAWarnings' -Value 1 -Type DWord
Set-ItemProperty -Path $security -Name 'AccessVBOM' -Value 1 -Type DWord
Write-Host "    VBAWarnings and AccessVBOM are 1 under $security"

Write-Host ''
Write-Host "RESULT: PASS - $Width-bit Excel is installed" -ForegroundColor Green
exit 0
