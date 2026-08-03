# Does the surface stay put when a native tool window opens, with a hole cut where it sits?
#
# Guards the fix for "opening non-xlide windows reverts the toolbar": the surface must never
# retreat to the document area. It keeps the frame's whole client and punches window-region
# holes where native tool windows sit, and the holes must heal when they close.
#
# Run tools\dev.ps1 -KeepOpen first, then this. It attaches to the running harness Excel,
# shows Locals (pane-class, moves are heard by the tracker) and the Object Browser (its own
# class, tracked only by the cutout poll), captures the frame after each change, and closes
# both. Expected: our menu bar and toolbar visible in every capture; each native window fully
# visible inside its hole; the healed capture identical to a plain boot. The shim log tells
# the same story in rectangles: "surface: N native hole(s) cut" then "whole again".
# Dev-harness script: reads Application.VBE, per decision 10's harness exception.
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -Namespace Xlide -Name Cut -MemberDefinition @'
[DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr context);
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
delegate bool EnumProc(IntPtr h, IntPtr l);

public static void UseRealPixels() { SetProcessDpiAwarenessContext(new IntPtr(-4)); }
public static bool Render(IntPtr window, IntPtr dc) { return PrintWindow(window, dc, 0x00000002); }

public static IntPtr TopLevel(int processId, string className)
{
    IntPtr found = IntPtr.Zero;
    long best = 0;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId || !IsWindowVisible(h)) { return true; }
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() != className) { return true; }
        RECT r;
        if (!GetWindowRect(h, out r)) { return true; }
        long area = (long)(r.Right - r.Left) * (r.Bottom - r.Top);
        if (area > best) { best = area; found = h; }
        return true;
    }, IntPtr.Zero);
    return found;
}
'@

[Xlide.Cut]::UseRealPixels()

$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$outDir = Join-Path (Split-Path -Parent (Split-Path -Parent $here)) 'artifacts'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Capture([int]$processId, [string]$name) {
    $frame = [Xlide.Cut]::TopLevel($processId, 'wndclass_desked_gsk')
    if ($frame -eq [IntPtr]::Zero) { throw 'No editor frame.' }

    $rect = New-Object Xlide.Cut+RECT
    [void] [Xlide.Cut]::GetWindowRect($frame, [ref] $rect)
    $bitmap = New-Object System.Drawing.Bitmap ($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $dc = $graphics.GetHdc()
            try { if (-not [Xlide.Cut]::Render($frame, $dc)) { throw 'Render refused.' } }
            finally { $graphics.ReleaseHdc($dc) }
        } finally { $graphics.Dispose() }
        $path = Join-Path $outDir "$name.png"
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Output "captured: $path"
    } finally { $bitmap.Dispose() }
}

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$processId = 0
foreach ($process in Get-Process EXCEL) { $processId = $process.Id }
$vbe = $app.VBE

function Set-ToolWindow([int]$type, [bool]$visible) {
    foreach ($window in $vbe.Windows) {
        if ($window.Type -eq $type) {
            $window.Visible = $visible
            Write-Output "window type ${type}: '$($window.Caption)' visible=$visible"
        }
    }
}

# Locals: pane-class, docked by default.
Set-ToolWindow 4 $true
Start-Sleep -Milliseconds 1500
Capture $processId 'cutout-locals'

# Object Browser on top of it: its own window class, the caption-only fallback.
Set-ToolWindow 2 $true
Start-Sleep -Milliseconds 1500
Capture $processId 'cutout-browser'

# Close both; the holes must heal without any pane event.
Set-ToolWindow 2 $false
Set-ToolWindow 4 $false
Start-Sleep -Milliseconds 1500
Capture $processId 'cutout-healed'

Write-Output "done"
