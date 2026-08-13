# Standing PASS/FAIL probe: discarding a tab's changes takes that tab's problems with it.
#
# Guards the 2026-08-06 regression. Closing a tab and declining to save reverts the module, and
# the analyzer answers for the reverted text -- but a live answer is only accepted while its
# module is still the one on screen, and a discarded module is closing. Nothing replaced the
# findings computed from the text that had just been thrown away, so the Problems panel went on
# reporting errors in code that no longer existed.
#
# Written as a live probe rather than a unit test because that is where the bug lived: every
# piece worked, and the defect was in what happens between them when a module goes away
# mid-flight. A seam test in Test-CloseConfirm.ps1 guards that the call is still there.
#
# HONEST LIMIT, and do not mistake this probe for more than it is. Run against a build with the
# fix disabled, it still passed all five checks: something else clears the findings within the
# time it waits. So it pins the flow -- type, get problems, close, decline, tab goes, panel ends
# empty -- but it is NOT proof that the regression cannot come back, because it never went red
# for it. Making it discriminate means finding what the slower path is and asserting the clear
# happens before that path could have run. See docs/testing.md.
#
# Run tools\dev.ps1 -KeepOpen -Configuration Debug first. Debug builds only.
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

$api = $found.Base
Write-Output "api: $api (pid $($found.Pid))"

function Page([string] $script) {
    $r = Invoke-RestMethod "$api/eval" -Method Post -Body $script -TimeoutSec 15
    if (-not $r.answered) { throw 'the page did not answer' }
    $out = $r.result
    if ($out -is [string] -and $out.StartsWith('"')) { $out = $out | ConvertFrom-Json }
    $out
}

function WaitFor([string] $what, [scriptblock] $condition, [int] $seconds = 20) {
    $deadline = (Get-Date).AddSeconds($seconds)
    do {
        if (& $condition) { return $true }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $deadline)
    throw "timed out waiting for $what"
}

function ProblemsFor([string] $module) {
    $r = Invoke-RestMethod "$api/problems?module=$module" -TimeoutSec 10

    # The leading comma is load-bearing. A function returns its output through the pipeline, which
    # enumerates it, so a one-element array arrives at the caller as the bare element -- and .Count
    # on a PSCustomObject is $null, not 1, so "exactly one finding" read as "no findings" and every
    # wait timed out while the api was answering correctly the whole time.
    return , @($r.findings | Where-Object { $_.module -eq $module })
}

# A module to work in, chosen rather than assumed. The session's shown module is empty whenever
# the surface has tabs but no active one, which is how a fresh -KeepOpen session starts, so the
# tab strip is the list that actually says what is open.
#
# It must also start with no problems of its own: the probe asserts that the findings it caused
# are gone, and a module whose saved text is already broken can never satisfy that.
# Returned as one delimited string rather than JSON: the eval route answers with the JSON
# encoding of the page's result, Page already decodes that once, and decoding an array a second
# time yields one element that is the array.
$open = @((Page "[...document.querySelectorAll('.tab')].map(t => t.textContent.trim()).join('\n')") -split "`n" | Where-Object { $_ })
if ($open.Count -eq 0) { Write-Output 'RESULT: FAIL - no tabs are open to work with'; return }

$module = $open | Where-Object { (ProblemsFor $_).Count -eq 0 } | Select-Object -First 1
if (-not $module) {
    Write-Output "RESULT: FAIL - every open module already has problems ($($open -join ', ')); nothing to prove a clear against"
    return
}

Write-Output "module: $module (of $($open -join ', '))"

# Shown, so the live analysis this probe is about actually runs for it.
Page "(() => { const t = [...document.querySelectorAll('.tab')].find(e => e.textContent.trim() === '$module'); t?.click(); return !!t; })()" | Out-Null

$original = (Invoke-RestMethod "$api/module?name=$module" -TimeoutSec 10).text
if ($null -eq $original) { Write-Output 'RESULT: FAIL - the module text could not be read'; return }

try {
    # Typed into the editor rather than written through the api. A host write carries a new
    # baseline with it, so the module would not be unsaved and the close would never ask -- and
    # the whole bug lives on the far side of that question. Typing is also what the person who
    # reported it did.
    #
    # A type mismatch rather than a syntax error: the analyzer reports this one as a finding,
    # where an unfinished statement is a parse failure it declines to diagnose, and a probe that
    # waits for a finding nobody promised fails for a reason that has nothing to do with the bug.
    $typed = 'Sub XlideDiscardProbe()\n    Dim n As Long\n    n = "oops"\nEnd Sub\n'
    Check 'the typed text produces problems' {
        Page "(() => { const ed = globalThis.xlideBridge?.workspace?.activeEditor?.(); if (!ed) return 'no editor'; const m = ed.getModel(); const e = m.getFullModelRange().getEndPosition(); ed.executeEdits('xlide-probe', [{ range: { startLineNumber: e.lineNumber, startColumn: e.column, endLineNumber: e.lineNumber, endColumn: e.column }, text: '\n$typed' }]); return 'typed'; })()" | Out-Null
        WaitFor 'findings to appear' { (ProblemsFor $module).Count -gt 0 }
    }

    Check 'the tab shows as unsaved' {
        WaitFor 'the dirty dot' {
            (Page "[...document.querySelectorAll('.tab')].some(t => t.textContent.trim() === '$module' && t.classList.contains('dirty'))") -eq $true
        }
    }

    # Closed through the page's own X, not by calling the host: the bug was in the sequence a real
    # close runs. The X is armed at pointerdown and fired at pointerup and never listens for a
    # click, because a press can survive the element being rebuilt underneath it, so a synthetic
    # click does nothing at all.
    Check 'the close question is asked' {
        Page "(() => { const t = [...document.querySelectorAll('.tab')].find(e => e.textContent.trim() === '$module'); const x = t?.querySelector('.tab-close'); if (!x) return 'no close box'; const o = { bubbles: true, cancelable: true, composed: true, button: 0, pointerId: 1, isPrimary: true }; x.dispatchEvent(new PointerEvent('pointerdown', o)); x.dispatchEvent(new PointerEvent('pointerup', o)); return 'pressed'; })()" | Out-Null
        WaitFor 'the close-confirm modal' {
            (Page "!!document.getElementById('close-confirm-backdrop')") -eq $true
        }
    }

    Check "Don't Save closes the tab" {
        Page "(() => { const b = [...document.querySelectorAll('#close-confirm-backdrop button')].find(x => /don.t save/i.test(x.textContent)); b?.click(); return !!b; })()" | Out-Null
        WaitFor 'the tab to go' {
            (Page "![...document.querySelectorAll('.tab')].some(e => e.textContent.trim() === '$module')") -eq $true
        }
    }

    # The regression itself.
    Check 'the discarded text takes its problems with it' {
        WaitFor 'findings to clear' { (ProblemsFor $module).Count -eq 0 }
    }
} finally {
    # The module goes back however this ended, so a failed run does not leave broken VBA behind.
    try {
        Invoke-RestMethod "$api/module?name=$module" -Method Post -Body $original -TimeoutSec 15 | Out-Null
        Write-Output "restored: $module"
    } catch {
        Write-Output "WARNING: $module could not be restored - $($_.Exception.Message)"
    }
}

Write-Output ''
foreach ($name in $checks.Keys) { Write-Output ("  {0,-45} {1}" -f $name, $checks[$name]) }
$failed = @($checks.Values | Where-Object { $_ -ne 'PASS' }).Count
Write-Output ''
Write-Output $(if ($failed -eq 0) { "RESULT: PASS - $($checks.Count) checks" } else { "RESULT: FAIL - $failed of $($checks.Count) checks" })
