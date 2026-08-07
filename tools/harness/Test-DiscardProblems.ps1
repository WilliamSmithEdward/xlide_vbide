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
# mid-flight. A seam test in Test-CloseConfirm.ps1 guards that the call is still there; this
# guards that it still does anything.
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

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excel) { Write-Output 'RESULT: FAIL - no Excel is running'; return }

$discoveryPath = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
if (-not (Test-Path $discoveryPath)) {
    Write-Output "RESULT: FAIL - no discovery file for Excel $($excel.Id)"
    return
}

$d = Get-Content $discoveryPath | ConvertFrom-Json
$api = "http://127.0.0.1:$($d.port)/$($d.token)"
Write-Output "api: $api (pid $($d.pid))"

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
    @($r.findings | Where-Object { $_.module -eq $module })
}

# The module the session is showing. The fixture opens its own, so the probe works with what is
# there rather than insisting on a name.
$state = Invoke-RestMethod "$api/state" -TimeoutSec 10
$module = $state.module
if (-not $module) { Write-Output 'RESULT: FAIL - no module is open to work with'; return }
Write-Output "module: $module"

$original = (Invoke-RestMethod "$api/module?name=$module" -TimeoutSec 10).text
if ($null -eq $original) { Write-Output 'RESULT: FAIL - the module text could not be read'; return }

try {
    # Something the analyzer will certainly object to, appended so the rest of the module still
    # parses and the findings that appear are the ones this probe made.
    $broken = $original + "`r`nSub XlideDiscardProbe()`r`n    If`r`nEnd Sub`r`n"
    Invoke-RestMethod "$api/module?name=$module" -Method Post -Body $broken -TimeoutSec 15 | Out-Null

    Check 'the broken text produces problems' {
        WaitFor 'findings to appear' { (ProblemsFor $module).Count -gt 0 }
    }

    # Close the tab and answer Don't Save, through the page's own path rather than by calling the
    # host directly: the bug was in the sequence a real close runs, and a shortcut past the page
    # would not have reproduced it.
    Check 'the close question is asked' {
        Page "(() => { const t = [...document.querySelectorAll('.tab')].find(e => e.textContent.includes('$module')); t?.querySelector('.tab-close, .tab-dirty')?.click(); return !!t; })()" | Out-Null
        WaitFor 'the close-confirm modal' {
            (Page "!!document.getElementById('close-confirm-backdrop')") -eq $true
        }
    }

    Check "Don't Save closes the tab" {
        Page "(() => { const b = [...document.querySelectorAll('#close-confirm-backdrop button')].find(x => /don.t save/i.test(x.textContent)); b?.click(); return !!b; })()" | Out-Null
        WaitFor 'the tab to go' {
            (Page "![...document.querySelectorAll('.tab')].some(e => e.textContent.includes('$module'))") -eq $true
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
