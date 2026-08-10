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
#   3. Live behaviour - its own Excel, driven SEMANTICALLY through the dev build's two
#      doors: the DevTools protocol clicks real elements on the live pages, and the shim's
#      debug api answers with the native truth. This is what pins the double-click
#      navigate leg that posted mouse messages never could. Window lifecycle (hide with
#      the editor, stay away, re-present) runs through the same doors; only the icon check
#      still asks Win32, because an icon is not a page's business.
#
# The live leg launches and kills its own Excel; run it with no Excel you care about open.
# It needs a DEBUG publish: the doors do not exist in Release, by design.
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
Test-Seam 'the dev doors are gated to Debug' (Join-Path $repo 'src\Xlide.Vbe.Shim\Diagnostics\DebugServer.cs') @(
    '^#if DEBUG', 'DebugReply', '\\"api\\":')
Test-Seam 'the dev build asks for the DevTools protocol' (Join-Path $repo 'src\Xlide.Vbe.Shim\WebView\WebView2Surface.cs') @(
    'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS', 'DevToolsPort', 'MessageTap')
Test-Seam 'the api carries the log, messages, capture, breakpoint, and immediate routes' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'case "log"', 'case "messages"', 'case "capture"', 'case "breakpoint"', 'case "immediate"')
Test-Seam 'the api carries the locals, watches, problems, module, and stats routes' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession*.cs') @(
    'case "locals"', 'case "watches"', 'case "problems"', 'case "module"', 'case "stats"')
Test-Seam 'the perf counters exist and are Debug-gated' (Join-Path $repo 'src\Xlide.Vbe.Shim\Diagnostics\PerfCounters.cs') @(
    '^#if DEBUG', 'PlacementFull', 'Marshal', 'RaiseToAtLeast')

$published = Join-Path $repo 'artifacts\publish\Xlide.Vbe.Shim\debug_win-x64\ui\editor\dist\editor.js'
if (Test-Path $published) {
    Test-Seam 'PUBLISHED bundle carries the palette page (stale deploy)' $published @('objbrowser-scope')
} else {
    Write-Output 'seam: skip - no publish tree on this machine; the stale-deploy tripwire has nothing to check'
}

function Invoke-NodeProbe {
    param([string] $Leg, [string] $Script, [string[]] $Arguments = @())

    $verdictText = & node (Join-Path $script:here $Script) @Arguments 2>$null | Select-Object -Last 1

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

# --- live leg -------------------------------------------------------------------------

Add-Type -Namespace XlideObTest -Name Native -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll", EntryPoint = "GetClassLongPtrW")] public static extern IntPtr GetClassLongPtr(IntPtr h, int index);

delegate bool EnumProc(IntPtr h, IntPtr l);

public static IntPtr TopLevel(int processId, string className)
{
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() != className) { return true; }
        found = h;
        return false;
    }, IntPtr.Zero);
    return found;
}
'@

function Test-Live {
    param([string] $Label, [bool] $Ok, [string] $Detail = '')

    if ($Ok) {
        Write-Output "live: ok - $Label"
    } else {
        $said = if ($Detail) { " ($Detail)" } else { '' }
        Write-Output "live: FAIL - $Label$said"
        $script:failures += 1
    }
}

Write-Output 'live: launching Excel with the VBE...'

$excelPath = "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE"
$scratch = Join-Path $here 'fixtures\scratch.xlsm'
$process = $null
$excel = $null

try {
    $process = Start-Process -FilePath $excelPath -ArgumentList "`"$scratch`"" -PassThru
    $deadline = (Get-Date).AddSeconds(45)

    while ($null -eq $excel -and (Get-Date) -lt $deadline) {
        try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { Start-Sleep -Milliseconds 300 }
    }
    if ($null -eq $excel) { throw 'Excel never answered on COM.' }

    # This one test genuinely needs the VBA project object model: it HIDES and re-shows the
    # editor, and there is no ungated equivalent of that - Excel's ribbon command opens the
    # editor but does not close it. Everything else in the harness works with the setting off,
    # so the failure is named rather than left to arrive as a null-reference three lines later.
    if ($null -eq $excel.VBE) {
        throw 'Application.VBE is null: this test needs "Trust access to the VBA project ' +
              'object model" enabled (Trust Center > Macro Settings). It is the only one that does.'
    }

    $excel.VBE.MainWindow.Visible = $true

    # The dev door announces itself in a per-process discovery file; its presence IS the
    # session being up, and its contents are how everything below asks and acts.
    $discoveryPath = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($process.Id).json"
    while (-not (Test-Path $discoveryPath) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 200 }
    if (-not (Test-Path $discoveryPath)) { throw 'no debug api discovery file; is this a Debug publish?' }

    $discovery = Get-Content $discoveryPath -Raw | ConvertFrom-Json
    $api = "http://127.0.0.1:$($discovery.port)/$($discovery.token)"

    $state = $null
    while ((Get-Date) -lt $deadline) {
        try { $state = Invoke-RestMethod "$api/state"; if ($state.surfaceReady) { break } } catch { }
        Start-Sleep -Milliseconds 250
    }
    Test-Live 'the debug api answers and the surface is ready' ($null -ne $state -and $state.surfaceReady)

    # The whole in-page story - summon by real click, real libraries, members from real
    # code, and the double-click navigate - runs in the node probe over the two doors.
    Invoke-NodeProbe 'live' 'objbrowser-live-probe.mjs' @('--api', $api, '--cdp', "$($discovery.devtoolsPort)")

    # The icon is a window property, not a page's; Win32 answers for it.
    $palette = [XlideObTest.Native]::TopLevel($process.Id, 'XlidePalette')
    $icon = if ($palette -ne [IntPtr]::Zero) { [XlideObTest.Native]::SendMessageW($palette, 0x007F, [IntPtr] 0, [IntPtr] 0) } else { [IntPtr]::Zero }
    if ($icon -eq [IntPtr]::Zero -and $palette -ne [IntPtr]::Zero) { $icon = [XlideObTest.Native]::GetClassLongPtr($palette, -34) }
    Test-Live 'the palette wears an icon' ($icon -ne [IntPtr]::Zero)

    # Lifecycle through the doors: hide with the editor, stay away, return on a summons.
    $excel.VBE.MainWindow.Visible = $false
    Start-Sleep -Milliseconds 800
    $state = Invoke-RestMethod "$api/state"
    Test-Live 'the palette hides when the editor closes' (-not $state.paletteVisible)

    $excel.VBE.MainWindow.Visible = $true
    Start-Sleep -Seconds 2
    $state = Invoke-RestMethod "$api/state"
    Test-Live 'the palette stays away until summoned' (-not $state.paletteVisible)

    $null = Invoke-RestMethod -Method Post "$api/command?name=objectBrowser"
    Start-Sleep -Milliseconds 800
    $state = Invoke-RestMethod "$api/state"
    Test-Live 'a summons by name re-presents the same palette' ($state.paletteOpen -and $state.paletteVisible)
}
catch {
    Write-Output "live: FAIL - $($_.Exception.Message)"
    $script:failures += 1
}
finally {
    if ($excel) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    if ($process) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Get-Process xlide-engine -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

if ($failures -eq 0) {
    Write-Output 'RESULT: PASS - the floating Object Browser, its scopes, its details pane, its navigate, and its lifecycle are as pinned'
} else {
    Write-Output "RESULT: FAIL - $failures check(s) down; the Object Browser behaviour has drifted"
}
