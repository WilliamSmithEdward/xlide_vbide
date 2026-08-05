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
#   3. Live behaviour - its own Excel: no palette at startup, the toolbar click summons a
#      visible XlidePalette wearing an icon, the typelib catalog answers, the native type-2
#      window never shows, the palette hides with the editor and stays hidden, and a second
#      summons re-presents the same window.
#
# The live leg launches and kills its own Excel; run it with no Excel you care about open.
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

Test-Seam 'the palette window exists and behaves' (Join-Path $repo 'src\Xlide.Vbe.Shim\Editor\BrowserPalette.cs') @(
    'XlidePalette', 'view=objbrowser', 'public void Hide', 'AdoptOwnerIcon', 'Reveal\(\)', 'Win32.SwHide')
Test-Seam 'the command is intercepted before the native execute' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession.cs') @(
    'command == VbeCommands.Command.ObjectBrowser', 'OpenBrowserPalette\(\);')
Test-Seam 'the session answers libraries, types, members, and hides with the frame' (Join-Path $repo 'src\Xlide.Vbe.Shim\AddIn\AddInSession.cs') @(
    'BrowseLibraries', 'BrowseTypes', 'BrowseMembers', 'ScanModuleMembers', '_browserPalette\?\.Hide\(\)')
Test-Seam 'members carry a line and libraries a kind' (Join-Path $repo 'src\Xlide.Vbe.Shim\Editor\EditorMessages.cs') @(
    'record ObMemberRow', '"line"', 'record ObLibraryRow', '"kind"')
Test-Seam 'the page boots the palette view' (Join-Path $repo 'ui\editor\src\main.ts') @(
    'view.*objbrowser', 'bootObjectBrowserPage')
Test-Seam 'the page carries scopes, the group pull, and the details pane' (Join-Path $repo 'ui\editor\src\objectbrowser.ts') @(
    'objbrowser-scope', 'pullWhole', 'objbrowser-splitter', 'objbrowser-detail-signature')
Test-Seam 'built bundle carries the palette page' (Join-Path $repo 'ui\editor\dist\editor.js') @(
    'objbrowser-scope', 'Pick a type on the left')

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

# --- live leg -------------------------------------------------------------------------

Add-Type -Namespace XlideObTest -Name Native -MemberDefinition @'
[DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr context);
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int hgt, uint flags);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll")] static extern IntPtr PostMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll", EntryPoint = "GetClassLongPtrW")] public static extern IntPtr GetClassLongPtr(IntPtr h, int index);

delegate bool EnumProc(IntPtr h, IntPtr l);

public static void UseRealPixels() { SetProcessDpiAwarenessContext(new IntPtr(-4)); }

public static IntPtr TopLevel(int processId, string className, bool visibleOnly)
{
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        if (visibleOnly && !IsWindowVisible(h)) { return true; }
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() != className) { return true; }
        found = h;
        return false;
    }, IntPtr.Zero);
    return found;
}

public static IntPtr ChildOfClass(IntPtr parent, string className)
{
    IntPtr found = IntPtr.Zero;
    EnumChildWindows(parent, (h, l) =>
    {
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() == className) { found = h; return false; }
        return true;
    }, IntPtr.Zero);
    return found;
}

public static void Click(IntPtr window, int x, int y)
{
    IntPtr point = (IntPtr)((y << 16) | (x & 0xFFFF));
    PostMessageW(window, 0x0200, IntPtr.Zero, point);
    PostMessageW(window, 0x0201, (IntPtr)1, point);
    PostMessageW(window, 0x0202, IntPtr.Zero, point);
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
[XlideObTest.Native]::UseRealPixels()

$logRoot = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'
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

    $excel.VBE.MainWindow.Visible = $true

    $frame = [IntPtr]::Zero
    while ($frame -eq [IntPtr]::Zero -and (Get-Date) -lt $deadline) {
        $frame = [XlideObTest.Native]::TopLevel($process.Id, 'wndclass_desked_gsk', $true)
        if ($frame -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 200 }
    }
    if ($frame -eq [IntPtr]::Zero) { throw 'the editor window never appeared.' }

    # The surface must be up before the toolbar exists to click.
    $logPattern = "shim-*-$($process.Id).log"
    $ready = $false
    while (-not $ready -and (Get-Date) -lt $deadline) {
        $log = Get-ChildItem $logRoot -Filter $logPattern -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($log -and (Select-String -Path $log.FullName -Pattern 'editor surface: ready' -Quiet)) { $ready = $true }
        else { Start-Sleep -Milliseconds 200 }
    }
    Test-Live 'the editor surface reports ready' $ready

    Test-Live 'no palette exists at startup' `
        ([XlideObTest.Native]::TopLevel($process.Id, 'XlidePalette', $false) -eq [IntPtr]::Zero)

    # Summon: the toolbar button at its measured spot in a 1500-wide frame.
    [void] [XlideObTest.Native]::SetWindowPos($frame, [IntPtr]::Zero, 60, 40, 1500, 900, 0x0004)
    [void] [XlideObTest.Native]::SetForegroundWindow($frame)
    Start-Sleep -Milliseconds 700
    $browser = [XlideObTest.Native]::ChildOfClass($frame, 'Chrome_RenderWidgetHostHWND')
    Test-Live 'the editor hosts a browser child' ($browser -ne [IntPtr]::Zero)
    [XlideObTest.Native]::Click($browser, 757, 39)

    $palette = [IntPtr]::Zero
    $summonDeadline = (Get-Date).AddSeconds(12)
    while ($palette -eq [IntPtr]::Zero -and (Get-Date) -lt $summonDeadline) {
        $palette = [XlideObTest.Native]::TopLevel($process.Id, 'XlidePalette', $true)
        if ($palette -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 200 }
    }
    Test-Live 'the toolbar click summons a visible palette' ($palette -ne [IntPtr]::Zero)
    if ($palette -eq [IntPtr]::Zero) { throw 'no palette; the remaining live checks have no subject.' }

    Start-Sleep -Seconds 3

    # The icon: WM_GETICON small, class small as the fallback - some icon must be there.
    $icon = [XlideObTest.Native]::SendMessageW($palette, 0x007F, [IntPtr] 0, [IntPtr] 0)
    if ($icon -eq [IntPtr]::Zero) { $icon = [XlideObTest.Native]::GetClassLongPtr($palette, -34) }
    Test-Live 'the palette wears an icon' ($icon -ne [IntPtr]::Zero)

    # The catalog answered: the log names the libraries and the project modules.
    $log = Get-ChildItem $logRoot -Filter $logPattern | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $libraries = Select-String -Path $log.FullName -Pattern 'object browser: (\d+) librarie' | Select-Object -Last 1
    $count = if ($libraries) { [int] $libraries.Matches[0].Groups[1].Value } else { 0 }
    Test-Live 'the catalog lists the project and the referenced libraries' ($count -ge 2) "counted $count"
    Test-Live 'the project answers with its modules' `
        (Select-String -Path $log.FullName -Pattern '-> \d+ module\(s\)' -Quiet)

    # The native Browser never shows: every type-2 window in the editor stays invisible.
    $nativeShown = $false
    $windows = $excel.VBE.Windows
    for ($i = 1; $i -le $windows.Count; $i++) {
        $window = $windows.Item($i)
        if ($window.Type -eq 2 -and $window.Visible) { $nativeShown = $true }
    }
    Test-Live 'the native Object Browser window stays hidden' (-not $nativeShown)

    # Hide with the editor; stay hidden on its return.
    $excel.VBE.MainWindow.Visible = $false
    Start-Sleep -Milliseconds 800
    Test-Live 'the palette hides when the editor closes' (-not [XlideObTest.Native]::IsWindowVisible($palette))

    $excel.VBE.MainWindow.Visible = $true
    Start-Sleep -Seconds 2
    Test-Live 'the palette stays away until summoned' (-not [XlideObTest.Native]::IsWindowVisible($palette))

    # A second summons re-presents the SAME window, state intact.
    $frame = [XlideObTest.Native]::TopLevel($process.Id, 'wndclass_desked_gsk', $true)
    [void] [XlideObTest.Native]::SetWindowPos($frame, [IntPtr]::Zero, 60, 40, 1500, 900, 0x0004)
    [void] [XlideObTest.Native]::SetForegroundWindow($frame)
    Start-Sleep -Milliseconds 700
    $browser = [XlideObTest.Native]::ChildOfClass($frame, 'Chrome_RenderWidgetHostHWND')
    [XlideObTest.Native]::Click($browser, 757, 39)
    $reshown = $false
    $summonDeadline = (Get-Date).AddSeconds(8)
    while (-not $reshown -and (Get-Date) -lt $summonDeadline) {
        $reshown = [XlideObTest.Native]::IsWindowVisible($palette)
        if (-not $reshown) { Start-Sleep -Milliseconds 200 }
    }
    Test-Live 'a second summons re-presents the same window' $reshown
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
    Write-Output 'RESULT: PASS - the floating Object Browser, its scopes, its details pane, and its lifecycle are as pinned'
} else {
    Write-Output "RESULT: FAIL - $failures check(s) down; the Object Browser behaviour has drifted"
}
