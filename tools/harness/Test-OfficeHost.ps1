<#
.SYNOPSIS
    Loads the registered Debug shim into the Excel this machine has, checks that the editor came
    up in the width expected, and with -Suites drives it.

.DESCRIPTION
    Test-ShimLoad.ps1 says the 32-bit shim loads and runs its own code in a 32-bit process.
    This is the other question, which only an Office can answer: does the editor come up in it.

    It registers nothing and installs nothing. It expects a Debug shim of the Excel's own width
    to be registered already (xlide-register --apply --bitness x86|x64 --shim <path>), starts a
    separate Excel on the scratch workbook, opens the editor through the ribbon, and checks:

      - the add-in loaded and its door answers healthy: engine up, page ready, readers up;
      - Excel is a process of the width asked for, and so is the shim it loaded;
      - the shim it loaded is the one named, when one is named.

    With -Suites it then builds the debug fixture through the api and runs the suites that
    cover what depends on the width of a pointer: the api's own probe, the Object Browser's
    window, the debugger with the Locals and Immediate readers, and the leak count.

    On a development machine this is the 64-bit Office and says nothing new. It is written for
    the CI job that installs an Excel of each width (.github/workflows/office.yml), where the
    64-bit run is the control and the 32-bit run is the news.

.EXAMPLE
    tools\harness\Test-OfficeHost.ps1 -Width 32 -Suites
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet(32, 64)]
    [int] $Width,

    # The library the session must report it loaded. Not checked when empty.
    [string] $Shim = '',

    # Also build the debug fixture and run the covering suites against it.
    [switch] $Suites,

    # Seconds to wait for Excel's window.
    [int] $TimeoutSeconds = 180
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $PSScriptRoot 'LibraryMachine.ps1')

$passed = 0
$failures = New-Object System.Collections.Generic.List[string]
function Check([string] $what, [bool] $ok, [string] $detail = '') {
    $tail = if ($detail) { "  -- $detail" } else { '' }
    if ($ok) {
        $script:passed++
        Write-Host "  ok   $what$tail"
    } else {
        $script:failures.Add($what)
        Write-Host "  FAIL $what$tail" -ForegroundColor Red
    }
}

function Stop-Mine([int] $id) {
    if ($id -le 0) { return }
    $running = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($running -and $running.ProcessName -eq 'EXCEL') {
        Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
        $running.WaitForExit(15000) | Out-Null
    }
}

function Start-Session([string[]] $launcherArguments) {
    $said = @(& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Start-Excel.ps1') `
        @launcherArguments -TimeoutSeconds $TimeoutSeconds 2>&1 | ForEach-Object { "$_" })
    $said | ForEach-Object { Write-Host "    $_" }
    $named = $said | Select-String 'pid=(\d+)' | Select-Object -Last 1
    return [pscustomobject] @{
        Id      = if ($named) { [int] $named.Matches[0].Groups[1].Value } else { 0 }
        Healthy = [bool] ($said | Select-String 'door is healthy')
    }
}

Write-Host "==> Excel, $Width-bit, on the scratch workbook" -ForegroundColor Cyan
$session = Start-Session @('-Separate')
Check 'Excel starts and the editor opens' ($session.Id -gt 0) "process $($session.Id)"
Check 'the add-in loaded and its door answers healthy' $session.Healthy

if ($session.Id -gt 0) {
    try {
        $answer = & node (Join-Path $PSScriptRoot 'xlide-api.mjs') --pid $session.Id doctor 2>$null | Out-String
        $doctor = $null
        try { $doctor = $answer | ConvertFrom-Json } catch { Write-Host "    the doctor's answer was not JSON: $answer" }

        if ($null -ne $doctor) {
            Check 'the engine is up' ([bool] $doctor.engineUp)
            Check 'the editor page is ready' ([bool] $doctor.surfaceReady)
            Check 'the Locals and Immediate readers are up' ([bool] $doctor.ghostReadersUp)
            Check 'the doctor has no findings' (@($doctor.findings).Count -eq 0) (@($doctor.findings) -join '; ')

            $loaded = "$($doctor.shimPath)"
            Check "the shim it loaded is $Width-bit" ((Test-Path $loaded) -and (Get-LibraryWidth $loaded) -eq $Width) $loaded
            if ($Shim) {
                $expected = (Resolve-Path $Shim).Path
                Check 'and it is the one named' ($loaded -ieq $expected) "$loaded against $expected"
            }
        } else {
            Check 'the doctor answers' $false
        }

        $excel = Get-Process -Id $session.Id -ErrorAction SilentlyContinue
        $excelWidth = if ($excel -and $excel.Path) { Get-LibraryWidth $excel.Path } else { 0 }
        Check "Excel is a $Width-bit process" ($excelWidth -eq $Width) "$(if ($excel) { $excel.Path }) is $excelWidth-bit"

        # What the shim wrote about itself from inside, which no outside reading can get wrong.
        $logs = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'
        $log = Get-ChildItem $logs -Filter "shim-*-$($session.Id).log" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime | Select-Object -Last 1
        $opened = if ($log) { Get-Content $log.FullName -Encoding UTF8 -TotalCount 1 } else { '' }
        $says = if ($Width -eq 32) { ', x86,' } else { ', x64,' }
        Check "the shim's own log says it is running $Width-bit" ("$opened".Contains($says)) "$opened"
    }
    finally {
        Stop-Mine $session.Id
    }
}

if ($Suites -and $failures.Count -eq 0) {
    Write-Host ''
    Write-Host '==> The debug fixture, built through the api' -ForegroundColor Cyan
    $built = @(& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repoRoot 'tools\New-DebugFixture.ps1') -Quiet 2>&1 |
        ForEach-Object { "$_" })
    $built | ForEach-Object { Write-Host "    $_" }
    $named = $built | Select-String 'pid=(\d+)' | Select-Object -Last 1
    $fixtureId = if ($named) { [int] $named.Matches[0].Groups[1].Value } else { 0 }
    $fixture = Join-Path $repoRoot 'artifacts\fixtures\DebugFixture.xlsm'
    Check 'the debug fixture is built and its session is up' ($fixtureId -gt 0 -and (Test-Path $fixture)) "process $fixtureId"

    if ($fixtureId -gt 0) {
        $env:XLIDE_PID = "$fixtureId"
        try {
            Write-Host ''
            Write-Host '==> The suites that depend on the width of a pointer' -ForegroundColor Cyan
            foreach ($suite in 'Test-DebugApi.ps1', 'objbrowser-live-probe.mjs', 'debugger-features.mjs',
                    'locals-hidden-rows.mjs', 'immediate-panel.mjs', 'com-leak.mjs') {
                $file = Join-Path $PSScriptRoot $suite
                $answer = if ($suite.EndsWith('.mjs')) { @(node $file 2>&1 | ForEach-Object { "$_" }) }
                    else { @(powershell -NoProfile -ExecutionPolicy Bypass -File $file 2>&1 | ForEach-Object { "$_" }) }
                $exit = $LASTEXITCODE

                # Two spellings of a verdict are in use, and a suite gives one of them.
                $result = $answer | Select-String 'RESULT: (PASS|FAIL)' | Select-Object -Last 1
                $counted = $answer | Select-String '(\d+) passed, (\d+) failed' | Select-Object -Last 1
                $ok = if ($result) { "$result" -match 'RESULT: PASS' }
                    elseif ($counted) { "$counted" -match ', 0 failed' -and $exit -eq 0 }
                    else { $false }
                $said = if ($result) { "$result".Trim() } elseif ($counted) { "$counted".Trim() + ", exit $exit" } else { "no verdict, exit $exit" }

                Check $suite $ok $said
                if (-not $ok) {
                    $answer | Select-String '\sFAIL\s' | Select-Object -First 12 | ForEach-Object { Write-Host "      $($_.Line.Trim())" }
                    $answer | Select-Object -Last 12 | ForEach-Object { Write-Host "      > $_" }
                }

                if (-not (Get-Process -Id $fixtureId -ErrorAction SilentlyContinue)) {
                    Check 'Excel is still standing' $false "it went during $suite"
                    break
                }
            }
        }
        finally {
            Remove-Item Env:XLIDE_PID -ErrorAction SilentlyContinue
            Stop-Mine $fixtureId
        }
    }
}

Write-Host ''
if ($failures.Count -eq 0) {
    Write-Host "RESULT: PASS - $passed checks, in $Width-bit Excel" -ForegroundColor Green
    exit 0
}

Write-Host "RESULT: FAIL - $($failures.Count) of $($passed + $failures.Count): $($failures -join '; ')" -ForegroundColor Red
exit 1
