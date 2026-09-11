# The Object Browser is a floating xlide palette, and everything the developer pinned about
# it on 2026-08-05 must hold: it opens as our own themed top-level window (never the native
# one), wears the editor's icon, hides with the editor and stays away until summoned, and
# its page carries the scope search, the whole-group pull, and the stacked details pane.
#
# Legs, PASS/FAIL:
#   1. Seams - the contract's load-bearing pieces exist in the sources, in the BUILT page
#      bundle, and in the PUBLISHED bundle (the stale-deploy tripwire).
#   2. Page behaviour - objbrowser-page-probe.mjs drives the built page headless: boot,
#      Group/Object/All scopes, the whole-group pull, details rows, splitter keyboard.
#
# Neither leg needs Excel. The LIVE behaviour is the gate's (verify.ps1 -Live), driven through
# the api with the trust setting off: objbrowser-live-probe.mjs clicks the summons, reads real
# libraries and members, double-clicks a member to navigate, and reads the icon off the
# palette's window (state.paletteIcon); window-routes.mjs closes and reshows the editor and
# holds the palette to following it down, staying away, and coming back as the same palette
# on the next summons. A third leg here launched an Excel of its own, attached with
# GetActiveObject, and hid the editor through Application.VBE, which the trust setting gates.
# It rotted outside the gate and was retired once its checks ran inside it (2026-09-10).
$ErrorActionPreference = 'Continue'

$here = $PSScriptRoot
$repo = Split-Path -Parent (Split-Path -Parent $here)
$failures = 0

# The session is a partial class across AddInSession.cs and AddInSession.DebugApi.cs, so a
# seam is looked for across `AddInSession*.cs` rather than in whichever file it was in on the
# day the check was written. Select-String and Test-Path both take the wildcard (2026-08-09).
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

Test-Seam 'the palette window exists and behaves' (Join-Path $repo 'src\Xlide.Vbe.Shim\Editor\BrowserPalette.cs') @(
    'XlidePalette', 'view=objbrowser', 'public void Hide', 'AdoptOwnerIcon', 'Reveal\(\)', 'Win32.SwHide')
Test-Seam 'the command is intercepted before the native execute' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'command == VbeCommands.Command.ObjectBrowser', 'OpenBrowserPalette\(\);')
Test-Seam 'the session answers libraries, types, members, and hides with the frame' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'BrowseLibraries', 'BrowseTypes', 'BrowseMembers', 'ScanModuleMembers', '_browserPalette\?\.Hide\(\)')
Test-Seam 'members carry a line and libraries a kind' (Join-Path $repo 'src\Xlide.Vbe.Shim\Editor\EditorMessages.cs') @(
    'record ObMemberRow', '"line"', 'record ObLibraryRow', '"kind"')
Test-Seam 'the page boots the palette view' (Join-Path $repo 'ui\editor\src\main.ts') @(
    'view.*objbrowser', 'bootObjectBrowserPage')
Test-Seam 'the page carries scopes, the group pull, and the details pane' (Join-Path $repo 'ui\editor\src\objectbrowser.ts') @(
    'objbrowser-scope', 'pullWhole', 'objbrowser-splitter', 'objbrowser-detail-signature')
Test-Seam 'built bundle carries the palette page' (Join-Path $repo 'ui\editor\dist\editor.js') @(
    'objbrowser-scope', 'Pick a type on the left')
# THE DOOR SHIPS IN EVERY BUILD NOW, shut in Release unless the agent card opens it - the phrase
# the gate reads out of the binary. It was gated to Debug when these seams were written, and they
# went on asking for an `#if DEBUG` that no longer exists (found 2026-09-10).
Test-Seam 'the api door is in the server, in every build' (Join-Path $repo 'src\Xlide.Vbe.Shim\Diagnostics\ApiServer.cs') @(
    'ApiReply', '\\"api\\":')
Test-Seam 'and a Release build keeps it shut unless told otherwise' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'keeps the door shut unless told otherwise')
Test-Seam 'the dev build asks for the DevTools protocol' (Join-Path $repo 'src\Xlide.Vbe.Shim\WebView\WebView2Surface.cs') @(
    'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS', 'DevToolsPort', 'MessageTap')
Test-Seam 'the api carries the log, messages, capture, breakpoint, and immediate routes' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'case "log"', 'case "messages"', 'case "capture"', 'case "breakpoint"', 'case "immediate"')
Test-Seam 'the api carries the locals, watches, problems, module, and stats routes' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'case "locals"', 'case "watches"', 'case "problems"', 'case "module"', 'case "stats"')
Test-Seam 'the perf counters exist' (Join-Path $repo 'src\Xlide.Vbe.Shim\Diagnostics\PerfCounters.cs') @(
    'PlacementFull', 'Marshal', 'RaiseToAtLeast')

$published = Join-Path $repo 'artifacts\publish\Xlide.Vbe.Shim\debug_win-x64\ui\editor\dist\editor.js'
if (Test-Path $published) {
    Test-Seam 'PUBLISHED bundle carries the palette page (stale deploy)' $published @('objbrowser-scope')
} else {
    Write-Output 'seam: skip - no publish tree on this machine; the stale-deploy tripwire has nothing to check'
}

function Invoke-NodeProbe {
    param([string] $Leg, [string] $Script)

    $verdictText = & node (Join-Path $script:here $Script) 2>$null | Select-Object -Last 1

    if (-not $verdictText) {
        Write-Output "${Leg}: FAIL - the probe printed no verdict"
        $script:failures += 1
        return
    }

    try {
        $verdict = $verdictText | ConvertFrom-Json
        foreach ($check in $verdict.checks) {
            if ($check.ok) {
                Write-Output "${Leg}: ok - $($check.name)"
            } else {
                $detail = if ($check.detail) { " ($($check.detail))" } else { '' }
                Write-Output "${Leg}: FAIL - $($check.name)$detail"
                $script:failures += 1
            }
        }
    } catch {
        Write-Output "${Leg}: FAIL - unreadable verdict: $verdictText"
        $script:failures += 1
    }
}

Write-Output 'page: driving the built palette page headless (Edge + DevTools protocol)...'
Invoke-NodeProbe 'page' 'objbrowser-page-probe.mjs'

if ($failures -eq 0) {
    Write-Output 'RESULT: PASS - the floating Object Browser''s seams, scopes, group pull, and details pane are as pinned'
} else {
    Write-Output "RESULT: FAIL - $failures check(s) down; the Object Browser behaviour has drifted"
}
