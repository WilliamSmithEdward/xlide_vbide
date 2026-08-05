# Closing a dirty tab must ask before it closes, and each answer must do what it says.
#
# Guards the 2026-08-05 close-confirm feature, developer-confirmed live: a tab whose module
# text differs from the workbook's saved text gets a Save / Don't Save / Cancel question;
# Save saves the workbook and closes, Don't Save reverts the module to its saved baseline
# and closes, Cancel keeps the tab. Two legs, PASS/FAIL:
#
#   1. Seams — the contract's load-bearing pieces exist in the sources, in the BUILT page
#      bundle, and in the PUBLISHED bundle (the stale-deploy tripwire: a rebuilt page that
#      never reached the publish tree has bitten before).
#   2. Behaviour — close-confirm-page-probe.mjs drives the built page headless through the
#      whole flow: ask, Escape, dedupe, queue, Don't Save, Cancel, middle-click, Save.
#
# Needs no Excel and touches none. The host half (baseline compare, workbook save, module
# revert) lives in AddInSession and speaks in the shim log during live tests:
#   close: <module> differs from <workbook>'s saved text; asking
#   close: saved <workbook> / close: <module> reverted to <workbook>'s saved text
$ErrorActionPreference = 'Continue'

$here = $PSScriptRoot
$repo = Split-Path -Parent (Split-Path -Parent $here)
$failures = 0

function Test-Seam {
    param([string] $Label, [string] $Path, [string[]] $Patterns)

    if (-not (Test-Path $Path)) {
        Write-Output "seam: FAIL - $Label - missing file $Path"
        $script:failures += 1
        return
    }

    foreach ($pattern in $Patterns) {
        if (-not (Select-String -Path $Path -Pattern $pattern -Quiet)) {
            Write-Output "seam: FAIL - $Label - no match for '$pattern'"
            $script:failures += 1
            return
        }
    }

    Write-Output "seam: ok - $Label"
}

Test-Seam 'page shell asks and dedupes' (Join-Path $repo 'ui\editor\src\shell.ts') @(
    'confirmClose\(', 'askedCloseConfirm', 'close-confirm-backdrop')
Test-Seam 'bridge carries the question and the answer' (Join-Path $repo 'ui\editor\src\bridge.ts') @(
    'type: "confirmClose"', 'action\?: string')
Test-Seam 'surface message is registered' (Join-Path $repo 'src\Xlide.Vbe.Shim\Editor\EditorMessages.cs') @(
    'ConfirmCloseMessage', 'JsonSerializable\(typeof\(ConfirmCloseMessage\)\)')
Test-Seam 'surface can ask and can drop edits' (Join-Path $repo 'src\Xlide.Vbe.Shim\Editor\EditorSurface.cs') @(
    'public void ConfirmClose', 'public void DiscardEdits')
Test-Seam 'session gates, saves, and reverts' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession.cs') @(
    'OnModuleCloseRequested', 'case "save"', 'case "discard"', 'SaveWorkbookOf', 'ModuleDiffersFromSaved')
Test-Seam 'Ctrl\+W goes through the same gate' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession.cs') @(
    'OnModuleCloseRequested\(shown')
Test-Seam 'built bundle carries the modal' (Join-Path $repo 'ui\editor\dist\editor.js') @(
    'close-confirm-backdrop', 'confirmClose')

$published = Join-Path $repo 'artifacts\publish\Xlide.Vbe.Shim\release_win-x64\ui\editor\dist\editor.js'
if (Test-Path $published) {
    Test-Seam 'PUBLISHED bundle carries the modal (stale deploy)' $published @('close-confirm-backdrop')
} else {
    Write-Output 'seam: skip - no publish tree on this machine; the stale-deploy tripwire has nothing to check'
}

Write-Output 'page: driving the built bundle headless (Edge + DevTools protocol)...'
$verdictText = & node (Join-Path $here 'close-confirm-page-probe.mjs') 2>$null | Select-Object -Last 1

if (-not $verdictText) {
    Write-Output 'page: FAIL - the probe printed no verdict'
    $failures += 1
} else {
    try {
        $verdict = $verdictText | ConvertFrom-Json
        foreach ($check in $verdict.checks) {
            if ($check.ok) {
                Write-Output "page: ok - $($check.name)"
            } else {
                $detail = if ($check.detail) { " ($($check.detail))" } else { '' }
                Write-Output "page: FAIL - $($check.name)$detail"
                $failures += 1
            }
        }
    } catch {
        Write-Output "page: FAIL - unreadable verdict: $verdictText"
        $failures += 1
    }
}

if ($failures -eq 0) {
    Write-Output 'RESULT: PASS - the close question, all three answers, the queue, and the deploy are as pinned'
} else {
    Write-Output "RESULT: FAIL - $failures check(s) down; the close-confirm behaviour has drifted"
}
