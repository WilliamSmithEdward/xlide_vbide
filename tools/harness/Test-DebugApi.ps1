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

Import-Module (Join-Path $PSScriptRoot 'XlideApi.psm1') -Force
try {
    $found = Get-XlideApi
} catch {
    Write-Output "RESULT: FAIL - $($_.Exception.Message)"
    return
}

$d = [pscustomobject] @{ pid = $found.Pid; port = $found.Port; token = $found.Token; devtoolsPort = $found.DevtoolsPort; api = $found.Schema }
$api = $found.Base
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

# IT BRINGS ITS OWN MODULE.
#
# Every module check below was written against a module named CleanModule assumed to be in
# whatever workbook is open. No fixture has one and the scratch workbook is empty, so six checks
# failed structurally against every fixture in the repo, reading as a product failure in the
# gate's live half. A probe that depends on a fixture it does not create is a probe that will one
# day test nothing (2026-08-08).
$existing = (Invoke-RestMethod "$api/project" -TimeoutSec 8).components | Where-Object { $_.name -eq 'CleanModule' }
if (-not $existing) {
    Invoke-RestMethod "$api/component?action=add&kind=module&name=CleanModule" -Method Post -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1500
}

# READ FROM A MODULE THE FIXTURE GUARANTEES, not the one this probe just made. A freshly added
# CleanModule is EMPTY, so asking it for non-empty text is a check that can only pass against a
# fixture already poisoned by an earlier run's residue - which is exactly what it did: it went
# green all week against a saved-over DebugFixture.xlsm and went red the first time the gate ran
# against a pristine one (2026-08-12). Runner ships in the fixture with a body, from
# New-DebugFixture.ps1, so an empty answer here means the reader, not the weather.
Check 'module reads through the session reader' {
    (Invoke-RestMethod "$api/module?name=Runner" -TimeoutSec 8).text.Length -gt 0
}

# A KIND ASKED FOR BY NAME IS THE KIND YOU GET.
#
# `kind` was parsed as an int and nothing else, so `kind=class` failed to parse, fell through
# to the default, and added a STANDARD module while still answering ok. The caller had no way
# to tell: it surfaced much later as the analyzer objecting that a Friend member was in a
# module that cannot hold one, which is a true finding about a fixture that was built wrong.
# `kind=module` above was equally unparsed and only looked right because 1 is the default
# (2026-08-09).
Check 'a component kind given as a word is honoured, not defaulted' {
    Invoke-RestMethod "$api/component?action=add&kind=class&name=KindProbe" -Method Post -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1200
    $made = (Invoke-RestMethod "$api/project" -TimeoutSec 8).components | Where-Object { $_.name -eq 'KindProbe' }
    Invoke-RestMethod "$api/component?action=remove&name=KindProbe" -Method Post -TimeoutSec 10 | Out-Null
    $made.kind -eq 'class'
}

Check 'a kind it cannot read is refused rather than guessed at' {
    $answer = Invoke-RestMethod "$api/component?action=add&kind=widget&name=NeverMade" -Method Post -TimeoutSec 10
    $gone = -not ((Invoke-RestMethod "$api/project" -TimeoutSec 8).components | Where-Object { $_.name -eq 'NeverMade' })
    $answer.error -and $gone
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

# A WRITE THAT DID NOT HAPPEN MUST NOT ANSWER LIKE ONE THAT DID.
#
# This route replied ok to everything. Writing to a module that is not there, writing text the
# editor refuses - both came back the same as a write that landed, because the writer logged its
# failures and returned nothing. The fixture builder already worked around it by reading the line
# count back afterwards, which was the right instinct about the wrong thing to have to do.
#
# The module that is not there is the cheap half of it. The expensive half - a write the editor
# refuses partway, which used to cost the module its previous text - is in write-rollback.mjs,
# which is not run here because provoking it leaves the editor needing a restart (2026-08-09).
# A LINE THE EDITOR CANNOT HOLD IS NOT WRITTEN AT ALL.
#
# It does not refuse one, it BREAKS one, at 1,023 characters and with no continuation character, so
# a statement is cut in half and a string literal left unterminated. A 2,018-character Debug.Print
# became a 1,023-character fragment followed by a 995-character one, and nothing said a word: the
# module held code that could not compile and the surface held the line the developer wrote
# (2026-08-09). 1,022 is the widest that comes back identical, and the boundary is where the
# off-by-one would hide, so both sides of it are checked.
Check 'a line the editor would break is refused, and the module keeps what it had' {
    $before = (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text
    $long = "Option Explicit`r`nPublic Sub Wide()`r`n    Debug.Print `"$('a' * 1200)`"`r`nEnd Sub"

    $reported = $false
    try {
        $answer = Invoke-RestMethod "$api/module?name=CleanModule" -Method Post -Body $long -TimeoutSec 10
        $reported = [bool] $answer.error
    } catch { $reported = $true }

    Start-Sleep -Milliseconds 800
    $after = (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text
    $reported -and ($after -eq $before)
}

Check 'a line of exactly 1,022 characters still writes' {
    $wide = "Option Explicit`r`n' $('a' * 1020)"
    Invoke-RestMethod "$api/module?name=CleanModule" -Method Post -Body $wide -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1200
    $back = (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text

    # Put the module back the way the checks below expect to find it.
    Invoke-RestMethod "$api/module?name=CleanModule" -Method Post -Body $runner -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1200

    ($back -split "`r`n" | Measure-Object -Maximum -Property Length).Maximum -eq 1022
}

# The same lie, told by the route next door. Opening a module that is not there did nothing and
# answered exactly like opening one that is: `ran: true`. A caller could not tell an open that
# landed from an open that found nothing, which for the bridge this door is meant to become is the
# difference between a working editor and a silent one (2026-08-09).
Check 'opening a module that is not there is reported, not reported ok' {
    $shown = (Invoke-RestMethod "$api/state" -TimeoutSec 8).module

    $reported = $false
    try {
        $answer = Invoke-RestMethod "$api/pane?action=open&module=NoSuchModuleHere" -Method Post -TimeoutSec 10
        $reported = [bool] $answer.error
    } catch { $reported = $true }

    # And it did not move the developer somewhere else on the way to failing.
    $reported -and ((Invoke-RestMethod "$api/state" -TimeoutSec 8).module -eq $shown)
}

Check 'a write to a module that is not there is reported, not reported ok' {
    try {
        $answer = Invoke-RestMethod "$api/module?name=NoSuchModuleHere" -Method Post `
            -Body 'Sub A()
End Sub' -TimeoutSec 10
        # An error reply carries `error`; a success carries `ran`. Answering ok is the defect.
        [bool] $answer.error
    } catch {
        # A non-200 is a report too, and a clearer one.
        $true
    }
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
    #
    # ASKED, THEN CONFIRMED. `ran` says the command was carried out, not that the caret ended up
    # where it was sent - and on a freshly started session it sometimes does not, because the
    # module was written moments earlier and the pane is still catching up. The next check then
    # runs with the caret on line 1, the editor raises its Macros dialog and waits for a person,
    # no breakpoint is ever reached, and three checks fail naming none of that. It read as a
    # product defect twice before the modal guard's own log gave it away (2026-08-09).
    $landed = $false
    $tries = 0
    while (-not $landed -and $tries -lt 10) {
        Invoke-RestMethod "$api/caret?module=CleanModule&line=16" -Method Post -TimeoutSec 8 | Out-Null
        Start-Sleep -Milliseconds 300
        $native = Invoke-RestMethod "$api/native" -TimeoutSec 8
        $landed = $native.activeModule -eq 'CleanModule' -and $native.caretLine -ge 14 -and $native.caretLine -le 17
        $tries++
    }

    if (-not $landed) {
        $where = Invoke-RestMethod "$api/native" -TimeoutSec 8
        Write-Output ("     the caret is on {0}:{1}, not inside RunTotal; a run from there opens the Macros dialog" -f $where.activeModule, $where.caretLine)
    }

    $landed
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
    # Matches on the ROUTE NAME being echoed, not a fixed phrase. The message used to say
    # "unknown route", which was often a lie: seventeen real routes are guarded on their
    # arguments and fall through to the same default, so a route that exists was told it does
    # not. The wording changed with that fix and this assertion did not.
    $answer = (Invoke-RestMethod "$api/nonesuch" -TimeoutSec 8).error
    ($null -ne $answer) -and ($answer -match 'nonesuch')
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

Check 'doctor finds a healthy session' {
    $d = Invoke-RestMethod "$api/doctor" -TimeoutSec 10
    $d.engineUp -and $d.ghostReadersUp -and $d.surfaceReady -and $d.pageBuildStamp -ne '(none reported)'
}

Check 'perf answers with raw samples' {
    $p = Invoke-RestMethod "$api/perf" -TimeoutSec 8
    $p.marshalMs.Count -gt 0
}

Check 'log waits for an event instead of sleeping for it' {
    # The point is the TIMING: it must return when the line arrives, not when the wait
    # expires, which is what makes it usable as an await instead of a guess.
    #
    # It watches for the line the COMMAND ITSELF writes, not for a publish that follows.
    # Matching on "modules: publish" made the check depend on modules being open, which they
    # are early in a fresh session and are not after two suites have run against it - so it
    # passed all day and then failed on a run where nothing had changed but the order
    # (2026-08-06). A probe should assert what it is testing, not the weather.
    # showNextStatement, NOT save. Any command writes the "command: " line this check waits on,
    # and in design mode this one is disabled, so pressing it does nothing at all. The old
    # choice was save, which SAVED THE WORKBOOK - with this probe's CleanModule and runner text
    # in it - so every gate run quietly rewrote DebugFixture.xlsm on disk, and later checks
    # inherited residue the fixture generator never wrote (2026-08-12).
    $since = (Invoke-RestMethod "$api/log?max=1" -TimeoutSec 8).next
    $job = Start-Job -ArgumentList $api {
        param($api)
        Start-Sleep 2
        Invoke-RestMethod "$api/command?name=showNextStatement" -Method Post -TimeoutSec 10
    }

    $clock = [Diagnostics.Stopwatch]::StartNew()
    $answer = Invoke-RestMethod "$api/log?since=$since&match=command: &waitMs=15000" -TimeoutSec 30
    $clock.Stop()
    Remove-Job $job -Force -ErrorAction SilentlyContinue

    # Arrived, and arrived because it was written rather than because the wait gave up.
    $answer.lines.Count -ge 1 -and $clock.Elapsed.TotalSeconds -lt 12
}

Check 'eval awaits a promise instead of answering an empty object' {
    # The browser's own ExecuteScript hands back {} for a promise, which reads as a page
    # fault; every async probe written against this door looked broken until the shape was
    # recognised (2026-08-06). Both halves are checked: a promise resolves, and a plain
    # value still comes straight back.
    $async = (Invoke-RestMethod "$api/eval" -Method Post -TimeoutSec 25 `
        -Body '(async () => { await new Promise(r => setTimeout(r, 400)); return { awaited: true }; })()').result
    $plain = (Invoke-RestMethod "$api/eval" -Method Post -TimeoutSec 15 -Body '2 + 3').result
    $async -match '"awaited":true' -and $plain.Trim() -eq '5'
}

Check 'await waits for a condition and times out honestly' {
    # A predicate that is already true costs nothing; one that arrives is waited for; one
    # that never comes gives up at its budget rather than at the transport's.
    $now = Invoke-RestMethod "$api/await" -Method Post -Body '!!window.xlideBridge' -TimeoutSec 30

    Invoke-RestMethod "$api/eval" -Method Post -TimeoutSec 15 `
        -Body 'window.__apiProbeFlag = false; setTimeout(() => { window.__apiProbeFlag = true; }, 700); "armed"' | Out-Null
    $arrives = Invoke-RestMethod "$api/await" -Method Post -Body 'window.__apiProbeFlag === true' -TimeoutSec 40

    $clock = [Diagnostics.Stopwatch]::StartNew()
    $never = Invoke-RestMethod "$api/await?waitMs=2000" -Method Post -Body 'window.__neverEver === 42' -TimeoutSec 40
    $clock.Stop()

    $now.met -and $now.elapsedMs -lt 200 `
        -and $arrives.met -and $arrives.elapsedMs -ge 500 `
        -and (-not $never.met) -and $clock.Elapsed.TotalSeconds -lt 8
}

Check 'layout answers with the whole visible arrangement' {
    $l = Invoke-RestMethod "$api/layout" -TimeoutSec 20
    $bottom = $l.sections | Where-Object { $_.side -eq 'bottom' }
    $l.editorGroups.Count -ge 1 -and $l.editorArea.w -gt 0 -and $bottom.standing -and $l.documents.Count -ge 1
}

Check 'inspect names the rule that set a property' {
    # The diagnosis this route exists for: a structural class of ours losing an argument to
    # a rule in the bundled stylesheet, which read as a flex bug in our own code and cost an
    # hour (2026-08-06).
    $i = Invoke-RestMethod "$api/inspect?selector=.dock-tabs&styles=display,align-items&rules=1&max=3" -TimeoutSec 20
    $first = $i.elements[0]
    $i.matched -ge 1 -and $first.w -gt 0 -and $first.styles.display -eq 'flex' `
        -and ($first.rules -join ' ') -match 'dock-tabs'
}

Check 'console keeps what the page said to itself' {
    # Only UNCAUGHT errors reach the shim log; a handled warning is invisible without a
    # DevTools client, which is exactly the live-test situation.
    Invoke-RestMethod "$api/eval" -Method Post -TimeoutSec 15 `
        -Body 'console.warn("probe: a handled warning"); "said"' | Out-Null
    Start-Sleep -Milliseconds 400
    $c = Invoke-RestMethod "$api/console?last=20" -TimeoutSec 15
    $c.installed -and ($c.lines -join "`n") -match 'probe: a handled warning'
}

Check 'capture crops to an element, and lands on it' {
    # A whole frame is a picture in which a 54-pixel drop zone cannot be seen. The crop's
    # size is checked against what the page says the element measures, which is also what
    # catches the origin being wrong: the first cut used the surface's PARENT window and
    # landed tens of pixels high, on the toolbar instead of the pane header (2026-08-06).
    $box = (Invoke-RestMethod "$api/inspect?selector=%23toolbar" -TimeoutSec 20).elements[0]
    $shot = Join-Path $env:TEMP 'xlide-crop-probe.bmp'
    Invoke-WebRequest "$api/capture?selector=%23toolbar&pad=0" -OutFile $shot -TimeoutSec 30

    $bytes = [IO.File]::ReadAllBytes($shot)
    $isBitmap = $bytes[0] -eq 0x42 -and $bytes[1] -eq 0x4D
    $width = [BitConverter]::ToInt32($bytes, 18)
    $height = [BitConverter]::ToInt32($bytes, 22)
    Remove-Item $shot -Force -ErrorAction SilentlyContinue

    # Within a pixel: the page rounds, and so does the crop.
    $isBitmap -and ([Math]::Abs($width - $box.w) -le 2) -and ([Math]::Abs($height - $box.h) -le 2)
}

Check 'a crop of nothing is refused, not guessed at' {
    $shot = Join-Path $env:TEMP 'xlide-crop-miss.txt'
    Invoke-WebRequest "$api/capture?selector=.no-such-element-anywhere" -OutFile $shot -TimeoutSec 20
    $answer = Get-Content $shot -Raw
    Remove-Item $shot -Force -ErrorAction SilentlyContinue
    $answer -match 'nothing matches'
}

Check 'bench times the surface and answers a shape' {
    $b = Invoke-RestMethod "$api/bench?what=type&n=15" -TimeoutSec 60
    $b.runs -eq 15 -and $b.samplesMs.Count -eq 15 `
        -and $b.minMs -le $b.medianMs -and $b.medianMs -le $b.maxMs -and $b.maxMs -lt 2000
}

Check 'an unknown benchmark is named, not guessed at' {
    (Invoke-RestMethod "$api/bench?what=nonesuch" -TimeoutSec 15).error -match 'unknown benchmark'
}

Check 'assert reports what it saw, not just a verdict' {
    $held = Invoke-RestMethod "$api/assert?that=surfaceReady&timeoutMs=3000" -Method Post -TimeoutSec 20
    $missed = Invoke-RestMethod "$api/assert?that=shownModule&value=NoSuchModule&timeoutMs=1000" -Method Post -TimeoutSec 20
    $held.held -and (-not $missed.held) -and $missed.saw -and $missed.saw -ne '(nothing)'
}

Check 'journal captures a whole moment in one request' {
    $j = Invoke-RestMethod "$api/journal?lines=40" -TimeoutSec 25
    $j.pid -gt 0 -and $j.log.Count -gt 0 -and $j.state -match 'shownModule' -and $null -ne $j.dialogs
}

Check 'history hands the session back as a script' {
    $h = Invoke-RestMethod "$api/history" -TimeoutSec 10
    $h.requests.Count -gt 0 -and $h.script -match 'Invoke-RestMethod'
}

Check 'keep=1 protects a dialog the caller meant to open' {
    # References is a real native modal. The guard must leave it alone, because opening it was
    # the point - and the first cut checked for it synchronously, which is too early to see a
    # dialog that has not appeared yet, so it protected nothing (2026-08-06).
    try { Invoke-RestMethod "$api/command?name=references&keep=1" -Method Post -TimeoutSec 12 | Out-Null } catch { }
    Start-Sleep -Seconds 3
    $opened = @((Invoke-RestMethod "$api/dialogs" -TimeoutSec 8).dialogs).Count -ge 1

    1..4 | ForEach-Object { try { Invoke-RestMethod "$api/state" -TimeoutSec 8 | Out-Null } catch { } }
    Start-Sleep -Milliseconds 1500
    $survived = @((Invoke-RestMethod "$api/dialogs" -TimeoutSec 8).dialogs).Count -ge 1

    Invoke-RestMethod "$api/dismiss?button=Cancel" -Method Post -TimeoutSec 10 | Out-Null
    Start-Sleep -Milliseconds 1500
    $closed = @((Invoke-RestMethod "$api/dialogs" -TimeoutSec 8).dialogs).Count -eq 0

    $opened -and $survived -and $closed
}

Check 'eval reaches the palette page as well as the editor' {
    Invoke-RestMethod "$api/command?name=objectBrowser" -Method Post -TimeoutSec 10 | Out-Null
    Start-Sleep -Seconds 2
    $answer = Invoke-RestMethod "$api/eval?surface=palette" -Method Post -Body 'document.title' -TimeoutSec 15

    # LEFT AS FOUND, which a second objectBrowser press does not do: that command means
    # "summon", not "toggle", so the old second press left the palette standing in front for
    # the rest of the session - and objbrowser-live-probe, later in this same session, rightly
    # failed on a palette already showing before its summons (2026-08-12). The palette route is
    # the door's own way to put it away - the same hide its close box performs, state intact -
    # and it briefly went through a raw window message instead, whose FindWindow half silently
    # found nothing. Read back through the door, so a hide that stops working is THIS check
    # going red rather than a mystery in whichever probe meets the leftovers next.
    $putAway = $false
    $hid = Invoke-RestMethod "$api/palette?action=hide" -Method Post -TimeoutSec 10
    if ($hid.did) {
        $deadline = (Get-Date).AddSeconds(5)
        do {
            Start-Sleep -Milliseconds 250
            $putAway = -not (Invoke-RestMethod "$api/state" -TimeoutSec 8).paletteVisible
        } while (-not $putAway -and (Get-Date) -lt $deadline)
    }

    $answer.answered -and $answer.result -match 'Object Browser' -and $putAway
}

Check 'a page exception reaches the shim log' {
    Invoke-RestMethod "$api/eval" -Method Post -TimeoutSec 15 `
        -Body 'setTimeout(() => { throw new Error("probe: a deliberate page fault"); }, 0); "thrown"' | Out-Null
    Start-Sleep -Milliseconds 1500
    @((Invoke-RestMethod "$api/log?match=deliberate page fault&max=5" -TimeoutSec 8).lines).Count -ge 1
}

# The hardening, tested the only way that counts: raise the modal that hung this api's own
# development. Run with the caret outside a procedure opens the editor's Macros dialog, which
# owns the editor until somebody answers it - and note it does NOT block the door, because a
# VBA modal pumps messages, so no timeout ever fires and only an explicit sweep can help.
$script:modalDetail = $null
Check 'a modal this door raised is seen, and the door sweeps it' {
    # THE HARDENING MOVED AND THIS CHECK DID NOT. It used to raise the Macros dialog, see it
    # standing, and expect the NEXT request to sweep it. The door clears a dialog its own request
    # raised within THAT request now, so there was nothing left to see and the check read the
    # empty list as a failure to raise one (2026-08-08).
    #
    # Asked in two halves, because the two claims need opposite conditions. `keep=1` holds the
    # dialog long enough to prove it was raised and to read its buttons WITHOUT the host thread
    # it is holding, which is the property that matters: a door that cannot see a modal cannot
    # report why everything else timed out. Then the same gesture without `keep`, which must come
    # back with nothing standing and the poll beating.
    Invoke-RestMethod "$api/caret?module=CleanModule&line=1" -Method Post -TimeoutSec 8 | Out-Null
    Start-Sleep -Milliseconds 500
    Invoke-RestMethod "$api/command?name=run&keep=1" -Method Post -TimeoutSec 20 | Out-Null
    Start-Sleep -Milliseconds 1200

    $seen = Invoke-RestMethod "$api/dialogs" -TimeoutSec 8
    $sawIt = $seen.dialogs.Count -ge 1 -and $seen.dialogs[0].caption -eq 'Macros'
    $buttons = if ($sawIt) { $seen.dialogs[0].buttons -join ', ' } else { '' }

    # Kept on purpose, so it has to be answered here. Leaving it standing is what made an earlier
    # rewrite of this check cascade into every run check after it.
    Invoke-RestMethod "$api/dismiss?button=Cancel" -Method Post -TimeoutSec 15 | Out-Null
    Start-Sleep -Milliseconds 2000

    Invoke-RestMethod "$api/caret?module=CleanModule&line=1" -Method Post -TimeoutSec 8 | Out-Null
    Start-Sleep -Milliseconds 500
    Invoke-RestMethod "$api/command?name=run" -Method Post -TimeoutSec 20 | Out-Null
    Start-Sleep -Milliseconds 2000

    $after = Invoke-RestMethod "$api/dialogs" -TimeoutSec 8
    $cleared = $after.dialogs.Count -eq 0

    # Recovered: the poll is ticking again, which is what a person would call unstuck.
    Start-Sleep -Milliseconds 1500
    $beating = (Invoke-RestMethod "$api/stats" -TimeoutSec 8).heartbeatAgeMs -lt 5000

    # Out of the pipeline on purpose: anything written here joins the return value and
    # makes the check pass on its own, which is how this check first went green wrongly.
    $script:modalDetail = "saw '$(if ($sawIt) { 'Macros' } else { 'nothing' })' buttons [$buttons], swept $cleared, beating $beating"
    [bool] ($sawIt -and $cleared -and $beating)
}

# Last, deliberately: this one RESTARTS the page. Every check above depends on host state
# built up in order, and a reload mid-suite left the modal check watching an editor whose
# caret had been re-established underneath it (2026-08-06).
Check 'reload brings the page back and names the bundle it runs' {
    $r = Invoke-RestMethod "$api/reload" -Method Post -TimeoutSec 45
    $r.ready -and $r.elapsedMs -lt 25000 -and $r.pageBuildStamp -ne '(none reported)' -and (-not $r.stale)
}

Check 'layout reset puts a rearranged workspace back' {
    # Also this suite's own cleanup: it leaves the arrangement the way it found it.
    Invoke-RestMethod "$api/eval" -Method Post -TimeoutSec 15 -Body @'
(() => {
  localStorage.setItem("xlide.docks.v1", JSON.stringify({
    sides: { left: null, right: { kind: "group", tabs: ["problems"], active: "problems" }, top: null,
             bottom: { kind: "group", tabs: ["immediate","locals","watch","explorer","properties"], active: "immediate" } },
    sizes: {}, closed: []
  }));
  return "rearranged";
})()
'@ | Out-Null

    $reset = Invoke-RestMethod "$api/layout?reset=1" -Method Post -TimeoutSec 45
    Start-Sleep -Milliseconds 800

    $l = Invoke-RestMethod "$api/layout" -TimeoutSec 20
    $left = $l.sections | Where-Object { $_.side -eq 'left' }
    $bottom = $l.sections | Where-Object { $_.side -eq 'bottom' }
    $panes = ($bottom.groups | ForEach-Object { $_.tabs.pane }) -join ','

    $reset.ran -and $left.standing -and $bottom.standing -and $panes -match 'problems'
}

# --- The agent front door, the knowledge routes, and the inside door ---

Check 'discovery advertises product, host, and a ready agent URL' {
    $disc = Get-Content (Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($d.pid).json") -Raw | ConvertFrom-Json
    $disc.product -eq 'xlide_vbide' -and $disc.host.Length -gt 0 -and $disc.agent -match '/agent$'
}

Check 'agent introduces the product and names the host' {
    $a = Invoke-RestMethod "$api/agent" -TimeoutSec 8
    $a.product -eq 'xlide_vbide' -and $a.host.Length -gt 0 -and $a.pid -eq $d.pid -and $a.startHere.Count -ge 3
}

Check 'every breadcrumb the agent hands out resolves' {
    $a = Invoke-RestMethod "$api/agent" -TimeoutSec 8
    $bad = @()
    foreach ($p in $a.next.PSObject.Properties) {
        try { Invoke-WebRequest $p.Value -UseBasicParsing -TimeoutSec 10 | Out-Null } catch { $bad += $p.Name }
    }
    $bad.Count -eq 0
}

Check 'every route the table calls safe answers as itself, not as unknown' {
    # The drift guard: a row whose route the switch no longer serves answers the default arm's
    # "no route" refusal, and this walk catches it. Only bareGetIsSafe rows are called bare.
    $r = Invoke-RestMethod "$api/agent/routes" -TimeoutSec 8
    if ($r.count -lt 50) { return $false }
    $bad = @()
    foreach ($row in ($r.routes | Where-Object { $_.bareGetIsSafe })) {
        try {
            $answer = Invoke-WebRequest "$api/$($row.name)" -UseBasicParsing -TimeoutSec 20
            if ($answer.Content -match 'no route ') { $bad += $row.name }
        } catch { $bad += $row.name }
    }
    if ($bad.Count -gt 0) { Write-Output "    drifted: $($bad -join ', ')" }
    $bad.Count -eq 0
}

Check 'agent/route details one route by name' {
    (Invoke-RestMethod "$api/agent/route?name=module" -TimeoutSec 8).name -eq 'module'
}

Check 'agent/examples hands runnable recipes' {
    (Invoke-RestMethod "$api/agent/examples" -TimeoutSec 8).examples.Count -ge 4
}

Check 'model teaches the host object model' {
    $m = Invoke-RestMethod "$api/model" -TimeoutSec 12
    $m.known -and $m.typeCount -gt 100 -and $m.host -eq 'excel'
}

Check 'model expands a type through its alias' {
    $w = Invoke-RestMethod "$api/model?type=Worksheet" -TimeoutSec 12
    $w.type.name -eq 'Excel.Worksheet' -and $w.type.memberCount -ge 20
}

Check 'analyzer teaches the rule catalogue with its authorities' {
    $rules = Invoke-RestMethod "$api/analyzer" -TimeoutSec 12
    $rules.ruleCount -ge 20 -and ($rules.rules | Where-Object { $_.code -eq 'unterminated-string' }).spec -match 'MS-VBAL'
}

Check 'the inside door answers from the running object table' {
    $door = [Runtime.InteropServices.Marshal]::GetActiveObject('Xlide.Api')
    ($door.Request('agent') | ConvertFrom-Json).pid -eq $d.pid
}

Check 'Guide answers the front door from inside' {
    $door = [Runtime.InteropServices.Marshal]::GetActiveObject('Xlide.Api')
    ($door.Guide() | ConvertFrom-Json).product -eq 'xlide_vbide'
}

Check 'the inside door serves the full surface: a module lifecycle' {
    $door = [Runtime.InteropServices.Marshal]::GetActiveObject('Xlide.Api')
    $null = $door.Request('component?action=add&name=DoorSuiteProof&kind=module')
    $null = $door.Request('module?name=DoorSuiteProof', "Sub DoorSuiteProofRuns()`r`nEnd Sub`r`n")
    $written = ($door.Request('module?name=DoorSuiteProof') | ConvertFrom-Json).text -match 'DoorSuiteProofRuns'
    $removed = ($door.Request('component?action=remove&name=DoorSuiteProof') | ConvertFrom-Json).ok
    $written -and $removed
}

Check 'the inside door refuses a page route toward the http door' {
    $door = [Runtime.InteropServices.Marshal]::GetActiveObject('Xlide.Api')
    ($door.Request('ui') | ConvertFrom-Json).error -match 'HTTP door'
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
