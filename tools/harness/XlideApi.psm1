# The one PowerShell answer to "which Excel am I driving?".
#
# Nine scripts rebuilt the discovery path by hand, and six of them took whatever
# `Get-Process EXCEL | Select-Object -First 1` returned - so with two Excels open, a probe
# could assert against the wrong session and a test could write into the wrong workbook. The
# node client learned this lesson first (xlide-api.mjs discover()/open()): enumerate every
# discovery file, prove each candidate with a /state call because a discovery file outlives a
# killed Excel, and THROW WITH THE LIST rather than guess when the choice is ambiguous. This
# module is that lesson carried to the PowerShell side, in the same two shapes: Find-XlideApi
# is discover() and answers every live session; Get-XlideApi is open() and answers exactly
# one or refuses.
#
# The route shapes are untouched; this is only how the base URL is found. Callers that print
# a suite verdict catch the throw and turn it into their own "RESULT: FAIL - ..." line, so
# the gate's parser keeps seeing what it has always seen.

Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Every live xlide session, newest first. Never throws; none answers an empty list.

.DESCRIPTION
    Enumerates %LOCALAPPDATA%\xlide_vbide\xlide-api-*.json and proves each candidate alive by
    asking /state - a discovery file outlives a killed Excel, so answering is the only proof
    of life. Each answer carries:

        Pid, Port, Token, DevtoolsPort, StartedAt   - the discovery file's fields
        Base                                        - "http://127.0.0.1:<port>/<token>"
        State                                       - the /state reply that proved liveness
#>
function Find-XlideApi {
    [CmdletBinding()]
    param(
        [int] $ProcessId
    )

    $directory = Join-Path $env:LOCALAPPDATA 'xlide_vbide'
    if (-not (Test-Path $directory)) {
        return @()
    }

    $live = @()
    foreach ($file in Get-ChildItem $directory -Filter 'xlide-api-*.json' -ErrorAction SilentlyContinue) {
        try {
            $entry = Get-Content $file.FullName -Raw | ConvertFrom-Json
        } catch {
            # A file being written as it is read, or a corpse mid-sweep.
            continue
        }

        if ($ProcessId -and $entry.pid -ne $ProcessId) {
            continue
        }

        try {
            $base = "http://127.0.0.1:$($entry.port)/$($entry.token)"
            $state = Invoke-RestMethod "$base/state" -TimeoutSec 2
            $live += [pscustomobject] @{
                Pid          = [int] $entry.pid
                Port         = [int] $entry.port
                Token        = [string] $entry.token
                DevtoolsPort = [int] $entry.devtoolsPort
                StartedAt    = [string] $entry.startedAt
                Schema       = [int] $entry.api
                Base         = $base
                State        = $state
            }
        } catch {
            # A session that will not answer is not a session.
        }
    }

    return @($live | Sort-Object StartedAt -Descending)
}

<#
.SYNOPSIS
    The one live xlide session, or the one hosted by -ProcessId. Refuses to guess.

.PARAMETER ProcessId
    Only accept the session hosted by this Excel process id.

.PARAMETER TimeoutSeconds
    Keep looking this long before giving up. The door appears a moment after the editor
    opens, so a script that just launched Excel waits here instead of sleeping and hoping.
#>
function Get-XlideApi {
    [CmdletBinding()]
    param(
        [int] $ProcessId,
        [int] $TimeoutSeconds = 0
    )

    # XLIDE_PID fills in when the caller named nothing, the same rule the node client follows:
    # the gate sets it to the Excel it launched, so another Office host of the owner's beside a
    # run (an Access database, 2026-09-05) cannot make a bare call ambiguous.
    if (-not $ProcessId -and $env:XLIDE_PID) {
        $ProcessId = [int] $env:XLIDE_PID
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    do {
        $live = @(Find-XlideApi -ProcessId $ProcessId)

        if ($live.Count -eq 1) {
            return $live[0]
        }

        if ($live.Count -gt 1) {
            # Only a -ProcessId caller said which one they meant; an ambiguous bare call is
            # refused with the list, because guessing which Excel to drive is how a test
            # writes into the wrong workbook.
            $named = ($live | ForEach-Object { "pid $($_.Pid) started $($_.StartedAt)" }) -join '; '
            throw "several xlide sessions are live ($named); say which with -ProcessId"
        }

        if ($TimeoutSeconds -gt 0 -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 200
        }
    } while ($TimeoutSeconds -gt 0 -and (Get-Date) -lt $deadline)

    $who = if ($ProcessId) { "for Excel $ProcessId" } else { 'in any Excel' }
    throw "no live xlide xlide api $who (Release build, Excel not running, or the editor was never opened)"
}

Export-ModuleMember -Function Find-XlideApi, Get-XlideApi
