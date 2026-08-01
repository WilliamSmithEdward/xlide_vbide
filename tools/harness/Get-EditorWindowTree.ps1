<#
.SYNOPSIS
    Reports the window structure the editor builds around a code pane.

.DESCRIPTION
    The editor exposes no way to ask which window shows which module, and its window classes are
    not documented anywhere. Positioning anything over a code pane means knowing that structure, so
    it is measured here rather than assumed.

    Opens a host, adds a module so a real code pane exists, then prints every window in the editor
    process with its class, rectangle, and depth.
#>
[CmdletBinding()]
param(
    [switch] $KeepOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -Namespace Xlide -Name Tree -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
delegate bool EnumProc(IntPtr h, IntPtr l);

static string Describe(IntPtr h, int depth)
{
    var cls = new System.Text.StringBuilder(128);
    GetClassNameW(h, cls, 128);
    var text = new System.Text.StringBuilder(160);
    GetWindowTextW(h, text, 160);

    RECT r;
    GetWindowRect(h, out r);

    return new string(' ', depth * 2)
        + cls.ToString().PadRight(28 - System.Math.Min(depth * 2, 20))
        + " " + (IsWindowVisible(h) ? "vis" : "   ")
        + " " + (r.Right - r.Left).ToString().PadLeft(5) + "x" + (r.Bottom - r.Top).ToString().PadLeft(5)
        + "  at " + r.Left.ToString().PadLeft(5) + "," + r.Top.ToString().PadLeft(5)
        + (text.Length > 0 ? "  \"" + text.ToString() + "\"" : "");
}

public static System.Collections.Generic.List<string> Dump(int processId)
{
    var lines = new System.Collections.Generic.List<string>();

    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }

        var cls = new System.Text.StringBuilder(128);
        GetClassNameW(h, cls, 128);

        // Only the editor's own top-level window and its descendants are interesting. The host has
        // many unrelated top-level windows.
        if (cls.ToString() != "wndclass_desked_gsk") { return true; }

        lines.Add(Describe(h, 0));
        WalkChildren(h, 1, lines);
        return true;
    }, IntPtr.Zero);

    return lines;
}

static void WalkChildren(IntPtr parent, int depth, System.Collections.Generic.List<string> lines)
{
    EnumChildWindows(parent, (child, l) =>
    {
        lines.Add(Describe(child, depth));
        return true;
    }, IntPtr.Zero);
}

public static object WorkbookWindowOf(int processId)
{
    IntPtr sheet = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        EnumChildWindows(h, (child, l2) =>
        {
            var name = new System.Text.StringBuilder(128);
            GetClassNameW(child, name, 128);
            if (name.ToString() == "EXCEL7") { sheet = child; return false; }
            return true;
        }, IntPtr.Zero);
        return sheet == IntPtr.Zero;
    }, IntPtr.Zero);

    if (sheet == IntPtr.Zero) { return null; }

    var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
    object window;
    return AccessibleObjectFromWindow(sheet, 0xFFFFFFF0u, ref dispatch, out window) == 0 ? window : null;
}
'@

$excelPath = "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE"
$scratch = Join-Path $PSScriptRoot 'fixtures\scratch.xlsx'
if (-not (Test-Path $scratch)) { & (Join-Path $PSScriptRoot 'New-ScratchWorkbook.ps1') | Out-Null }

$process = Start-Process -FilePath $excelPath -ArgumentList $scratch -PassThru
Write-Host "Excel is process $($process.Id)."

try {
    $deadline = (Get-Date).AddSeconds(60)
    $window = $null
    while ($null -eq $window -and (Get-Date) -lt $deadline) {
        $window = [Xlide.Tree]::WorkbookWindowOf($process.Id)
        if ($null -eq $window) { Start-Sleep -Milliseconds 25 }
    }

    if ($null -eq $window) { throw 'Could not reach the host through its window.' }

    $excel = $window.Application
    $excel.DisplayAlerts = $false

    # A code pane only exists once a component is open in one, so make one.
    $component = $excel.ActiveWorkbook.VBProject.VBComponents.Add(1)
    $component.Name = 'ProbeModule'
    $component.CodeModule.AddFromString("Option Explicit`r`n`r`nSub Probe()`r`n    Dim n As Long`r`n    n = 1`r`nEnd Sub`r`n")

    $excel.VBE.MainWindow.Visible = $true
    $component.CodeModule.CodePane.Show()

    Start-Sleep -Milliseconds 1200

    Write-Host ''
    Write-Host 'Editor window tree:' -ForegroundColor Cyan
    foreach ($line in [Xlide.Tree]::Dump($process.Id)) { Write-Host $line }

    Write-Host ''
    Write-Host 'Code panes reported by the object model:' -ForegroundColor Cyan
    foreach ($pane in $excel.VBE.CodePanes) {
        Write-Host ("  {0}  top={1} visible lines={2}" -f $pane.CodeModule.Parent.Name, $pane.TopLine, $pane.CountOfVisibleLines)
    }
}
finally {
    if (-not $KeepOpen) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "Leaving Excel $($process.Id) running."
    }
}
