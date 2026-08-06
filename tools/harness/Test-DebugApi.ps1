# Standing PASS/FAIL probe for the shim's debug api: every route, and the break-mode round
# trip that is the whole point of it (push code, set a breakpoint, run, read live locals).
# Run tools\dev.ps1 -KeepOpen first, or any harness session; Debug builds only.
# Dev-harness script: uses Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

$checks = [ordered] @{}
function Check([string] $name, [scriptblock] $test) {
    try {
        $result = & $test
        $checks[$name] = if ($result) { 'PASS' } else { 'FAIL' }
    } catch {
        $checks[$name] = "FAIL ($($_.Exception.Message))"
    }
}

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excel) { Write-Output 'RESULT: FAIL - no Excel is running'; return }

$discoveryPath = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
if (-not (Test-Path $discoveryPath)) {
    Write-Output "RESULT: FAIL - no discovery file for Excel $($excel.Id) (Release build, or the editor was never opened)"
    return
}

$d = Get-Content $discoveryPath | ConvertFrom-Json
$api = "http://127.0.0.1:$($d.port)/$($d.token)"
Write-Output "api: $api (pid $($d.pid), devtools $($d.devtoolsPort))"

Check 'discovery carries a schema, port, token, and devtools port' {
    $d.api -ge 1 -and $d.port -gt 0 -and $d.token.Length -ge 8 -and $d.devtoolsPort -gt 0
}

Check 'the token gates: a wrong one is refused' {
    try {
        Invoke-RestMethod "http://127.0.0.1:$($d.port)/wrongtoken/state" -TimeoutSec 5 | Out-Null
        $false
    } catch {
        $true
    }
}

Check 'state answers with a live surface' {
    $s = Invoke-RestMethod "$api/state" -TimeoutSec 8
    $s.surfaceReady -and $s.configuration -eq 'debug'
}

Check 'windows lists the editor own windows' {
    (Invoke-RestMethod "$api/windows" -TimeoutSec 8).windows.Count -gt 0
}

Check 'stats reports counters' {
    $t = Invoke-RestMethod "$api/stats" -TimeoutSec 8
    $t.uptimeSeconds -ge 0 -and $t.handleCount -gt 0 -and $t.logLines -gt 0
}

Check 'log slices and filters server side' {
    $l = Invoke-RestMethod "$api/log?match=surface&max=5" -TimeoutSec 8
    $l.next -gt 0 -and $l.lines.Count -le 5
}

Check 'messages tap holds page traffic' {
    (Invoke-RestMethod "$api/messages?last=5" -TimeoutSec 8).messages.Count -ge 0
}

Check 'problems answers the analyzer findings' {
    $null -ne (Invoke-RestMethod "$api/problems" -TimeoutSec 8).findings
}

Check 'capture renders the frame to a bitmap' {
    $bmp = Join-Path $env:TEMP 'xlide-api-probe.bmp'
    Invoke-WebRequest "$api/capture" -OutFile $bmp -TimeoutSec 20
    (Get-Item $bmp).Length -gt 10000
}

Check 'module reads through the session reader' {
    (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text.Length -gt 0
}

# The round trip: write a runner, break inside it, read locals, reset. This is the shape
# every debugger-milestone regression will take.
$runner = @'
Option Explicit

Public Function Total(ByVal values As Variant) As Double
    Dim i As Long
    Dim sum As Double

    For i = LBound(values) To UBound(values)
        sum = sum + values(i)
    Next i

    Total = sum
End Function

Public Sub RunTotal()
    Dim answer As Double
    answer = Total(Array(1, 2, 3))
End Sub
'@

Check 'module writes through the session writer' {
    Invoke-RestMethod "$api/module?name=CleanModule" -Method Post -Body $runner -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1500
    (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text -match 'RunTotal'
}

Check 'breakpoint is idempotent with state=on' {
    $first = (Invoke-RestMethod "$api/breakpoint?module=CleanModule&line=8&state=on" -Method Post -TimeoutSec 8).ran
    $again = (Invoke-RestMethod "$api/breakpoint?module=CleanModule&line=8&state=on" -Method Post -TimeoutSec 8).ran
    $first -and $again
}

Check 'caret lands inside the procedure to be run' {
    # Run acts on the caret, and the host copies the SURFACE's caret into the native pane
    # before every command, so scrolling is not aiming: with the caret on line 1 the editor
    # opens its Macros dialog and waits for a person (2026-08-06). Line 16 is inside RunTotal.
    (Invoke-RestMethod "$api/caret?module=CleanModule&line=16" -Method Post -TimeoutSec 8).ran
}

$reachedBreak = $false
Check 'a run reaches the breakpoint and state says break' {
    # Through the product's own Run, in process. Never Application.Run from a killable job:
    # that call BLOCKS inside the break, and killing its caller while Excel is suspended in
    # it took Excel down (2026-08-06).
    try {
        Invoke-RestMethod "$api/command?name=run" -Method Post -TimeoutSec 8 | Out-Null
    } catch {
        # A run that reaches a breakpoint does not let the host thread answer until it does;
        # the timeout is expected, and the state poll below is the real assertion.
    }

    $deadline = (Get-Date).AddSeconds(25)
    do {
        Start-Sleep -Milliseconds 400
        $s = Invoke-RestMethod "$api/state" -TimeoutSec 5
    } while ($s.debugMode -ne 'break' -and (Get-Date) -lt $deadline)
    $script:reachedBreak = $s.debugMode -eq 'break'
    $script:reachedBreak
}

Check 'locals answers the stopped scope' {
    if (-not $script:reachedBreak) { throw 'never reached break' }
    Start-Sleep -Milliseconds 1500
    $rows = (Invoke-RestMethod "$api/locals" -TimeoutSec 8).rows
    ($rows | Where-Object { $_.expression -eq 'sum' }) -and ($rows | Where-Object { $_.expression -eq 'i' })
}

Check 'watches answers stopped with no watches set' {
    if (-not $script:reachedBreak) { throw 'never reached break' }
    (Invoke-RestMethod "$api/watches" -TimeoutSec 8).stopped
}

Check 'command by name resets out of the break' {
    Invoke-RestMethod "$api/command?name=reset" -Method Post -TimeoutSec 8 | Out-Null
    Start-Sleep -Milliseconds 1200
    (Invoke-RestMethod "$api/state" -TimeoutSec 8).debugMode -ne 'break'
}

Check 'breakpoint clears with state=off' {
    (Invoke-RestMethod "$api/breakpoint?module=CleanModule&line=8&state=off" -Method Post -TimeoutSec 8).ran -eq $false
}

Check 'an unknown route is named, not crashed on' {
    (Invoke-RestMethod "$api/nonesuch" -TimeoutSec 8).error -match 'unknown route'
}

Check 'dialogs answers when nothing is standing' {
    $null -ne (Invoke-RestMethod "$api/dialogs" -TimeoutSec 8).dialogs
}

Check 'eval runs script in the live page' {
    $r = Invoke-RestMethod "$api/eval" -Method Post -Body 'document.querySelectorAll(".panel-tab").length' -TimeoutSec 15
    $r.answered -and [int]$r.result -ge 4
}

Check 'stats carries a fresh host heartbeat' {
    (Invoke-RestMethod "$api/stats" -TimeoutSec 8).heartbeatAgeMs -lt 5000
}

# The hardening, tested the only way that counts: raise the modal that hung this api's own
# development. Run with the caret outside a procedure opens the editor's Macros dialog, which
# owns the editor until somebody answers it - and note it does NOT block the door, because a
# VBA modal pumps messages, so no timeout ever fires and only an explicit sweep can help.
$script:modalDetail = $null
Check 'a modal this door raised is seen, then cleared by the next request' {
    Invoke-RestMethod "$api/caret?module=CleanModule&line=1" -Method Post -TimeoutSec 8 | Out-Null
    Start-Sleep -Milliseconds 500
    Invoke-RestMethod "$api/command?name=run" -Method Post -TimeoutSec 20 | Out-Null
    Start-Sleep -Milliseconds 1200

    # Seen: the dialog and its buttons, read without the host thread the dialog is holding.
    $seen = Invoke-RestMethod "$api/dialogs" -TimeoutSec 8
    $sawIt = $seen.dialogs.Count -ge 1 -and $seen.dialogs[0].caption -eq 'Macros'

    # Cleared: the guard waits until the editor is genuinely wedged - three seconds of a
    # stopped poll - before it touches anything, so a dialog that is merely passing through
    # is never swept. Past that, the next request that needs the editor clears it.
    Start-Sleep -Seconds 4
    Invoke-RestMethod "$api/state" -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1500
    $after = Invoke-RestMethod "$api/dialogs" -TimeoutSec 8
    $cleared = $after.dialogs.Count -eq 0

    # Recovered: the poll is ticking again, which is what a person would call unstuck.
    Start-Sleep -Milliseconds 1500
    $beating = (Invoke-RestMethod "$api/stats" -TimeoutSec 8).heartbeatAgeMs -lt 5000

    # Out of the pipeline on purpose: anything written here joins the return value and
    # makes the check pass on its own, which is how this check first went green wrongly.
    $script:modalDetail = "saw '$($seen.dialogs[0].caption)' buttons [$($seen.dialogs[0].buttons -join ', ')], cleared $cleared, beating $beating"
    [bool] ($sawIt -and $cleared -and $beating)
}

Write-Output ''
foreach ($name in $checks.Keys) { "  {0,-52} {1}" -f $name, $checks[$name] }
if ($script:modalDetail) { Write-Output "  modal guard: $script:modalDetail" }
$failed = @($checks.Values | Where-Object { $_ -ne 'PASS' })
Write-Output ''
if ($failed.Count -eq 0) {
    Write-Output "RESULT: PASS - $($checks.Count) checks"
} else {
    Write-Output "RESULT: FAIL - $($failed.Count) of $($checks.Count) checks failed"
}
