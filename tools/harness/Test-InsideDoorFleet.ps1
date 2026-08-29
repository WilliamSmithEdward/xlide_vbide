<#
.SYNOPSIS
    Two Excel PROCESSES, one inside-door name, and what happens to it when the holder dies.

.DESCRIPTION
    The inside door (`GetObject(, "Xlide.Api")`) is registered in the running object table by
    every session, and GetObject binds ONE of them - the first, then whoever survives. That is
    fine while sessions close politely. A session that is force-killed is the question this
    probe exists to answer, and it could not be asked until Start-Excel learned `-Separate`:
    one process is one add-in load and one registration, so the fleet behaviour needs two.

    Measured first on 2026-08-29, one process only: a force-killed SOLO session leaves the name
    clean - GetActiveObject answers MK_E_UNAVAILABLE, exactly as if nothing had ever registered,
    so Windows removes a dead process's entry. That result is what narrowed #11 to the
    multi-registrant case: whether GetObject binds the DEAD first registrant (an RPC failure
    against a process that is gone) before it will fall back to the live second one.

    WHAT IT DOES: starts A, starts B in its own process, asks the name who holds it, force-kills
    the holder, then asks the name repeatedly for a minute. Every answer is reported - the pid it
    binds, or the HRESULT it refuses with - so the reading is data rather than a verdict.

    It leaves both sessions closed.

.EXAMPLE
    tools\harness\Test-InsideDoorFleet.ps1
#>
[CmdletBinding()]
param(
    # Seconds to keep asking the name after the holder is killed.
    [int] $WatchSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$fixture = Join-Path $repoRoot 'artifacts\fixtures\DebugFixture.xlsm'
$twin = Join-Path $repoRoot 'artifacts\fixtures\TwinFixture.xlsm'
foreach ($one in @($fixture, $twin)) {
    if (-not (Test-Path $one)) { throw "No workbook at $one - build the fixtures first." }
}

# What the name answers RIGHT NOW: the pid behind it, or the failure it refuses with. Never
# throws, because the refusal is the measurement as much as the answer is.
function Read-InsideDoor {
    try {
        $door = [Runtime.InteropServices.Marshal]::GetActiveObject('Xlide.Api')
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match '0x800401E3|MK_E_UNAVAILABLE') { return 'unavailable (no registration)' }
        return "bind failed: $message"
    }

    try {
        $agent = $door.Request('agent') | ConvertFrom-Json
        return "pid $($agent.pid)"
    }
    catch {
        $message = $_.Exception.Message
        if ($message -match '0x800706BE|RPC') { return 'CORPSE (bound a dead process: RPC failed)' }
        return "bound, but Request failed: $message"
    }
    finally {
        if ($null -ne $door) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($door) }
    }
}

$started = @()
try {
    Write-Host '== session A (first registrant: it should hold the name) =='
    $a = & (Join-Path $PSScriptRoot 'Start-Excel.ps1') -Workbook $fixture -Fresh
    $a | Select-Object -Last 1 | Out-Host
    $pidA = (Get-Process EXCEL | Sort-Object StartTime | Select-Object -Last 1).Id
    $started += $pidA
    Write-Host "   A is pid $pidA"

    Write-Host ''
    Write-Host '== session B (its OWN process, so a second registration of the one name) =='
    $b = & (Join-Path $PSScriptRoot 'Start-Excel.ps1') -Workbook $twin -Separate
    $b | Select-Object -Last 1 | Out-Host
    $pidB = (Get-Process EXCEL | Where-Object { $_.Id -ne $pidA } | Sort-Object StartTime | Select-Object -Last 1).Id
    $started += $pidB
    Write-Host "   B is pid $pidB"

    if ($pidA -eq $pidB -or -not $pidB) {
        throw 'Excel did not start a second process; /x was refused and there is nothing to measure.'
    }

    Write-Host ''
    Write-Host '== who holds the name, with both alive =='
    $holder = Read-InsideDoor
    Write-Host "   the name answers: $holder"

    $holderPid = if ($holder -match 'pid (\d+)') { [int]$Matches[1] } else { 0 }
    if (-not $holderPid) { throw "The name did not answer with a pid ($holder); nothing to kill." }
    $survivor = if ($holderPid -eq $pidA) { $pidB } else { $pidA }

    Write-Host ''
    Write-Host "== force-killing the HOLDER (pid $holderPid); pid $survivor stays up =="
    Stop-Process -Id $holderPid -Force
    $started = @($started | Where-Object { $_ -ne $holderPid })

    Write-Host ''
    Write-Host "== asking the name every 2s for ${WatchSeconds}s =="
    $deadline = (Get-Date).AddSeconds($WatchSeconds)
    $seen = @{}
    $firstSurvivorAt = $null
    $began = Get-Date
    while ((Get-Date) -lt $deadline) {
        $answer = Read-InsideDoor
        $at = [int]((Get-Date) - $began).TotalSeconds
        if (-not $seen.ContainsKey($answer)) {
            $seen[$answer] = $at
            Write-Host ("   +{0,3}s  {1}" -f $at, $answer)
        }
        if ($null -eq $firstSurvivorAt -and $answer -eq "pid $survivor") { $firstSurvivorAt = $at }
        Start-Sleep -Seconds 2
    }

    Write-Host ''
    Write-Host '== what that means =='
    if ($null -ne $firstSurvivorAt) {
        Write-Host "   the survivor (pid $survivor) answers the name, first seen +${firstSurvivorAt}s." -ForegroundColor Green
        if ($firstSurvivorAt -eq 0) {
            Write-Host '   No corpse window at all: the failover is immediate.' -ForegroundColor Green
        }
        else {
            Write-Host "   For ${firstSurvivorAt}s before that, the name did NOT reach the live session." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "   The survivor (pid $survivor) NEVER answered the name in ${WatchSeconds}s." -ForegroundColor Red
        Write-Host '   That is #11: a live, registered session unreachable by name because the' -ForegroundColor Red
        Write-Host '   holder died. Every distinct answer seen is listed above.' -ForegroundColor Red
    }
}
finally {
    foreach ($one in $started) { Stop-Process -Id $one -Force -ErrorAction SilentlyContinue }
    Write-Host ''
    Write-Host 'Both sessions closed.'
}
