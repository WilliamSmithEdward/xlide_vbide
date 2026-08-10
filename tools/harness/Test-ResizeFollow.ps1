# Does the overlay follow a frame resize, with panes open and with none visible -
# and does the BROWSER follow the overlay?
#
# Guards the 2026-08-04 fix: placement used to follow only PANE events, so the empty
# workspace resized in silence. Also guards the 2026-08-05 fix: the browser's size rode
# only on WM_SIZE delivery, and a raced resize left the page laid out for a width the
# window no longer had - its minimap and scrollbar fell off the right edge. Placement now
# asserts the browser bounds every pass, so the Chromium child must match the overlay
# after every resize here. Run tools\dev.ps1 -KeepOpen first, then this. Expected:
# MATCH on every line (measured DPI-aware; without UseRealPixels a scaling artifact reads
# as a one-pixel mismatch that is not real).
# Resizes the VBE frame twice per state and compares frame client, overlay, and browser.
$ErrorActionPreference = 'Continue'

Add-Type -Namespace Xlide -Name Rsz -MemberDefinition @'
[DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr context);
public static void UseRealPixels() { SetProcessDpiAwarenessContext(new IntPtr(-4)); }
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int w, int hgt, uint flags);
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public delegate bool EnumProc(IntPtr h, IntPtr l);

public static IntPtr FindChild(IntPtr root, string className)
{
    IntPtr found = IntPtr.Zero;
    EnumChildWindows(root, (h, l) =>
    {
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() == className) { found = h; return false; }
        return true;
    }, IntPtr.Zero);
    return found;
}

public static IntPtr FrameOf(int processId)
{
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() == "wndclass_desked_gsk") { found = h; return false; }
        return true;
    }, IntPtr.Zero);
    return found;
}
'@

[Xlide.Rsz]::UseRealPixels()
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE
$processId = (Get-Process EXCEL | Select-Object -First 1).Id
$frame = [Xlide.Rsz]::FrameOf($processId)
$overlay = [Xlide.Rsz]::FindChild($frame, 'XlideEditorOverlay')
Write-Output "frame $frame overlay $overlay"

function Report([string]$title) {
    $client = New-Object Xlide.Rsz+RECT
    [void] [Xlide.Rsz]::GetClientRect($frame, [ref] $client)
    $o = New-Object Xlide.Rsz+RECT
    [void] [Xlide.Rsz]::GetWindowRect($overlay, [ref] $o)
    $f = New-Object Xlide.Rsz+RECT
    [void] [Xlide.Rsz]::GetWindowRect($frame, [ref] $f)
    $ow = $o.Right - $o.Left; $oh = $o.Bottom - $o.Top
    $cw = $client.Right - $client.Left; $ch = $client.Bottom - $client.Top

    # The Chromium child is the page's real viewport; a stale one is the clipped-minimap bug
    # even while the overlay itself matches. Re-found per report: the browser can rebuild it.
    $browser = [Xlide.Rsz]::FindChild($overlay, 'Chrome_WidgetWin_0')
    $bw = 0; $bh = 0
    if ($browser -ne [IntPtr]::Zero) {
        $b = New-Object Xlide.Rsz+RECT
        [void] [Xlide.Rsz]::GetWindowRect($browser, [ref] $b)
        $bw = $b.Right - $b.Left; $bh = $b.Bottom - $b.Top
    }

    $match = if ($ow -eq $cw -and $oh -eq $ch -and $bw -eq $cw -and $bh -eq $ch) { 'MATCH' } else { 'MISMATCH' }
    Write-Output "$title -- client ${cw}x${ch} overlay ${ow}x${oh} browser ${bw}x${bh} -> $match"
}

function Resize([int]$w, [int]$h) {
    [void] [Xlide.Rsz]::SetWindowPos($frame, [IntPtr]::Zero, 0, 0, $w, $h, 0x0016) # NOZORDER|NOACTIVATE|NOMOVE
    Start-Sleep -Milliseconds 1200
}

Write-Output '--- state: panes open (harness scratch shows CleanModule)'
Report 'before'
Resize 900 700
Report 'after 900x700'
Resize 1400 900
Report 'after 1400x900'

Write-Output '--- state: all panes closed'
foreach ($window in $vbe.Windows) {
    if ($window.Type -eq 0) { try { $window.Close() } catch {} }
}
Start-Sleep -Milliseconds 1500
Report 'closed, before'
Resize 1000 800
Report 'closed, after 1000x800'
Resize 1500 1000
Report 'closed, after 1500x1000'
Write-Output 'done'
