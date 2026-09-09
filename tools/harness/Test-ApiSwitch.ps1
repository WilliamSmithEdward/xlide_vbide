<#
.SYNOPSIS
    Proves the api door actually shuts, and that the choice survives a restart.

.DESCRIPTION
    THE ONE CHECK THAT CANNOT BE WRITTEN IN THE HARNESS. Every other suite drives the editor
    through the api; this one turns the api OFF, which severs exactly that connection. So it is
    written here, outside the process, where the questions are socket questions:

      - is the discovery file gone?
      - is anything still listening on the port?
      - does the port refuse a connection?

    and then, across a restart:

      - does a remembered `false` keep the door shut even in a build that leans open?

    IT RESTARTS EXCEL, TWICE, and puts the setting back the way it found it. That is the cost of
    asking honestly: the door's whole point is that nothing reaches in once it is closed, so the
    only way back in is the way a developer would take - the card, or the settings file.

.EXAMPLE
    tools\harness\Test-ApiSwitch.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$dataFolder = Join-Path $env:LOCALAPPDATA 'xlide_vbide'
$settingsPath = Join-Path $dataFolder 'settings.json'

$passed = 0
$failed = 0

function Check([string] $name, [bool] $ok, [string] $detail = '') {
    if ($ok) {
        $script:passed++
        Write-Host "ok   $name" -ForegroundColor Green
    } else {
        $script:failed++
        Write-Host "FAIL $name$(if ($detail) { " - $detail" })" -ForegroundColor Red
    }
}

$ownedPidFile = Join-Path $dataFolder 'harness-excel.pid'

function Get-RecordedPid {
    if (-not (Test-Path $ownedPidFile)) { return 0 }
    $recorded = 0
    if ([int]::TryParse((Get-Content $ownedPidFile -Raw).Trim(), [ref] $recorded)) { return $recorded }
    return 0
}

function Get-DoorFile {
    # THE SESSION UNDER TEST: the Excel XLIDE_PID names (the gate's), else the one a restart here
    # recorded, else the newest door whose process is alive. A door file outlives an Excel
    # stopped by id, and the last file by name is anybody's - another automation's, when its
    # tests opened the editor (2026-09-08).
    $files = @(Get-ChildItem (Join-Path $dataFolder 'xlide-api-*.json') -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending)
    $preferred = @()
    if ($env:XLIDE_PID) { $preferred += [int] $env:XLIDE_PID }
    $recorded = Get-RecordedPid
    if ($recorded -gt 0) { $preferred += $recorded }
    foreach ($id in $preferred) {
        $match = $files | Where-Object { $_.BaseName -eq "xlide-api-$id" } | Select-Object -First 1
        if ($match -and (Get-Process -Id $id -ErrorAction SilentlyContinue)) { return $match }
    }
    foreach ($file in $files) {
        if ($file.BaseName -match 'xlide-api-(\d+)$' -and (Get-Process -Id ([int] $Matches[1]) -ErrorAction SilentlyContinue)) {
            return $file
        }
    }
    return $null
}

function Get-Door {
    $file = Get-DoorFile
    if (-not $file) { return $null }
    $door = Get-Content $file.FullName -Raw | ConvertFrom-Json
    $door | Add-Member -NotePropertyName 'hostPid' -NotePropertyValue ([int] ($file.BaseName -replace '^xlide-api-', '')) -Force
    $door
}

function Test-Listening([int] $port) {
    @(Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue).Count -gt 0
}

function Set-ApiSetting([bool] $on) {
    $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
    $settings | Add-Member -NotePropertyName 'api.enabled' -NotePropertyValue $on -Force
    ($settings | ConvertTo-Json) | Out-File $settingsPath -Encoding utf8
}

function Restart-Editor([int] $hostPid) {
    # ONLY THE SESSION UNDER TEST is stopped, by id, and the Excel a restart here recorded when
    # it started one. Every Excel by name was the old way, and it ended other automations'
    # Excels in the middle of whatever statement they were running (#24).
    $targets = @()
    if ($hostPid -gt 0) { $targets += $hostPid }
    $recorded = Get-RecordedPid
    if ($recorded -gt 0) { $targets += $recorded }
    foreach ($id in @($targets | Select-Object -Unique)) {
        $standing = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -ne $standing -and $standing.ProcessName -eq 'EXCEL') {
            Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
            $standing.WaitForExit(10000) | Out-Null
        }
    }
    Start-Sleep -Seconds 1
    # -AllowExistingExcel: whatever else is running is left running, and the check starts its
    # own process beside it.
    & powershell -NoProfile -ExecutionPolicy Bypass `
        -File (Join-Path $PSScriptRoot 'Invoke-VbeLoadCheck.ps1') -KeepOpen -AllowExistingExcel | Out-Null
    Start-Sleep -Seconds 3
}

# ---- the door is up to begin with -----------------------------------------------------------

$door = Get-Door
if (-not $door) {
    throw 'no session is advertising an api door; open the editor with the api on first'
}

$port = [int] $door.port
Check 'the door is listening before anything is asked of it' (Test-Listening $port)

# ---- shut it, through its own card ------------------------------------------------------------
#
# Fired and not awaited: the request rides the door it is closing, so the reply never arrives.
# That is not a failure, it is the behaviour under test.

Push-Location $repoRoot
try {
    $script = @"
import { open, waitFor } from "file:///$($repoRoot -replace '\\', '/')/tools/harness/xlide-api.mjs";
const api = await open();
await api.act("toolbar", { command: "openAgent" });
await waitFor("the card", async () => (await api.ui()).agent !== null, { budgetMs: 15000 });
await waitFor("its state", async () => (await api.ui()).agent?.busy === false, { budgetMs: 15000 });
api.act("agentCard", { press: "toggle" }).catch(() => {});
setTimeout(() => process.exit(0), 2500);
"@
    $shutScript = Join-Path $env:TEMP 'xlide-shut-the-door.mjs'
    $script | Out-File $shutScript -Encoding utf8
    node $shutScript | Out-Null
} finally {
    Pop-Location
}

Start-Sleep -Seconds 2

Check 'the discovery file is gone' (-not (Get-Door))
Check 'nothing is listening on the port any more' (-not (Test-Listening $port))

$refused = $false
try {
    Invoke-RestMethod "http://127.0.0.1:$port/$($door.token)/state" -TimeoutSec 3 | Out-Null
} catch {
    $refused = $true
}
Check 'and the port refuses a connection' $refused

$saved = Get-Content $settingsPath -Raw | ConvertFrom-Json
Check 'the choice was written down' ($saved.'api.enabled' -eq $false)

# ---- and it stays shut across a restart -------------------------------------------------------
#
# THE POINT OF THE WHOLE CHANGE. A dev build leans OPEN, so a door that stays shut here is a
# remembered `false` outranking the build's own lean - which is the same mechanism that keeps a
# shipped build shut when nobody has said anything at all.

Restart-Editor $door.hostPid
Check 'a remembered no keeps the door shut, even in a build that leans open' (-not (Get-Door))

# ---- put it back --------------------------------------------------------------------------

Set-ApiSetting $true
Restart-Editor 0

$back = Get-Door
Check 'and a remembered yes opens it again' ($null -ne $back)
if ($back) {
    Check 'on a fresh port and a fresh token' ($back.port -ne $port -or $back.token -ne $door.token)
}

Write-Host ''
Write-Host "$passed passed, $failed failed" -ForegroundColor $(if ($failed) { 'Red' } else { 'Green' })
exit $(if ($failed) { 1 } else { 0 })
