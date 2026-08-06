<#
.SYNOPSIS
    Opens the editor with the add-in loaded and captures it to an image.

.DESCRIPTION
    Looking at the thing is often the only way to know it is right, so capturing it has to be
    reliable rather than occasional.

    Three rules keep this from hanging, each learned by hitting it:

    The host is launched as a process and attached to through its own window. Attaching through the
    running object table costs ten to forty seconds because the host publishes itself there long
    after it is usable.

    Every wait is a deadline against an observed condition, never a sleep. The script waits for the
    editor window to exist and for the add-in to report that it has finished, and gives up rather
    than blocking forever.

    The window is asked to render itself, rather than the screen being read. Reading the screen
    cannot block, which is tempting, but it reads whatever is actually in front, and a background
    process is not allowed to reliably put a window there. That failure is silent and produces a
    perfect capture of the wrong application. Rendering the window works regardless of what is on
    top of it, and captures the browser surface, which an ordinary render request does not.

    Rendering can block on a window that is not pumping messages, so it runs on its own thread with
    a deadline and the script reports a timeout rather than hanging on one.

.EXAMPLE
    tools\harness\Get-EditorScreenshot.ps1
    Opens the editor with sample modules, captures it, and closes it.

.EXAMPLE
    tools\harness\Get-EditorScreenshot.ps1 -KeepOpen -OutFile shot.png
#>
[CmdletBinding()]
param(
    [string] $OutFile,

    # Leave the host running afterwards.
    [switch] $KeepOpen,

    # Skip adding sample modules; capture whatever the workbook already has.
    [switch] $NoSampleCode,

    # Upper bound on waiting for the editor and the add-in.
    [int] $TimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$repoRoot = Split-Path -Parent (Split-Path -Parent $here)

if (-not $OutFile) { $OutFile = Join-Path $repoRoot 'artifacts\editor.png' }
New-Item -ItemType Directory -Force -Path (Split-Path $OutFile) | Out-Null

$logRoot = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'

Add-Type -AssemblyName System.Drawing

Add-Type -Namespace Xlide -Name Shot -MemberDefinition @'
[DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr context);
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
[DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
[DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] static extern bool IsHungAppWindow(IntPtr h);
[DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
delegate bool EnumProc(IntPtr h, IntPtr l);

/// <summary>
/// Measures windows in real pixels rather than in the scaled coordinates Windows invents for a
/// process that has not said it understands scaling.
///
/// Without this the editor measures two thirds of its real size on a scaled display, the bitmap is
/// made that size, and the window renders into it at full size: the right and bottom of every
/// capture are silently cut off. It reads as missing user interface rather than as a cropped image,
/// and it hid a panel that was working.
/// </summary>
public static void UseRealPixels()
{
    // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2. Fails harmlessly if already set.
    SetProcessDpiAwarenessContext(new IntPtr(-4));
}

/// Reports whether a window will answer a render request at all.
public static bool IsResponsive(IntPtr window) { return !IsHungAppWindow(window); }

/// <summary>
/// Renders a window into a device context.
///
/// PW_RENDERFULLCONTENT is what makes a composited surface appear; without it a window hosting a
/// browser renders as a hole where the browser should be.
/// </summary>
public static bool Render(IntPtr window, IntPtr dc) { return PrintWindow(window, dc, 0x00000002); }

/// <summary>
/// Finds the largest visible top-level window of a class in a process.
///
/// Largest rather than first. A process can own several windows of one class, and the editor owns
/// small ones of its frame class that are not the frame: taking the first match produced a capture
/// 237 pixels wide of something that was not the editor.
/// </summary>
public static IntPtr TopLevel(int processId, string className)
{
    IntPtr found = IntPtr.Zero;
    long best = 0;

    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        if (!IsWindowVisible(h)) { return true; }

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

/// <summary>
/// Finds a window by its caption.
///
/// A tool window that has not been docked is its own top-level window, so rendering the editor
/// frame does not include it. Capturing it needs its own handle.
/// </summary>
public static IntPtr ByCaption(int processId, string caption)
{
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }

        var text = new System.Text.StringBuilder(256);
        GetWindowTextW(h, text, 256);
        if (text.ToString() == caption) { found = h; return false; }
        return true;
    }, IntPtr.Zero);

    return found;
}

/// The workbook window, which answers with the host's object model without the running object table.
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

[Xlide.Shot]::UseRealPixels()

# Terminating the host is normal here, and the host treats it as a crash. Clearing what it offers
# on the next start keeps a second run from failing for a reason unrelated to what is being looked at.
foreach ($version in @('16.0', '15.0')) {
    $resiliency = "HKCU:\Software\Microsoft\Office\$version\Excel\Resiliency"
    if (Test-Path $resiliency) {
        foreach ($child in @('DocumentRecovery', 'DisabledItems', 'StartupItems', 'CrashingAddinList')) {
            $path = Join-Path $resiliency $child
            if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

$addInKey = 'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Xlide.VbeAddIn'
if ((Test-Path $addInKey) -and (Get-ItemProperty $addInKey -Name LoadBehavior).LoadBehavior -ne 3) {
    Set-ItemProperty $addInKey -Name LoadBehavior -Value 3 -Type DWord
    Write-Host 'Re-enabled the add-in, which the editor had disabled after a failed load.'
}

$excelPath = "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE"
$scratch = Join-Path $here 'fixtures\scratch.xlsm'
if (-not (Test-Path $scratch)) { & (Join-Path $here 'New-ScratchWorkbook.ps1') | Out-Null }

$process = Start-Process -FilePath $excelPath -ArgumentList $scratch -PassThru
$excel = $null

try {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)

    $window = $null
    while ($null -eq $window -and (Get-Date) -lt $deadline) {
        $window = [Xlide.Shot]::WorkbookWindowOf($process.Id)

        # Excel sometimes hands the launch off to a fresh process of its own (measured
        # 2026-08-05 evening: the launched pid exits within seconds and a respawn opens the
        # workbook; the harness then waited on the corpse, threw, and never opened the VBE).
        # The session is whichever live Excel answers for the workbook, so the handoff is
        # followed: everything downstream keys off $process.Id - the shim log pattern, the
        # editor window, the cleanup - and moves with it.
        if ($null -eq $window -and $process.HasExited) {
            foreach ($candidate in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
                $window = [Xlide.Shot]::WorkbookWindowOf($candidate.Id)
                if ($null -ne $window) {
                    Write-Host "Excel $($process.Id) handed off to $($candidate.Id); following."
                    $process = $candidate
                    break
                }
            }
        }

        if ($null -eq $window) { Start-Sleep -Milliseconds 25 }
    }

    if ($null -eq $window) { throw "Could not reach Excel $($process.Id) through its window." }

    $excel = $window.Application
    $excel.DisplayAlerts = $false

    if (-not $NoSampleCode) {
        $components = $excel.ActiveWorkbook.VBProject.VBComponents

        # Removed first rather than added blindly. The workbook is macro-enabled, so a module can
        # outlive a run, and renaming a new component to a name already in use throws.
        foreach ($name in @('BrokenModule', 'CleanModule')) {
            foreach ($existing in @($components)) {
                if ($existing.Name -eq $name) { $components.Remove($existing) }
            }
        }

        $broken = $components.Add(1)
        $broken.Name = 'BrokenModule'
        $broken.CodeModule.AddFromString("Option Explicit`r`n`r`n' A defect the analyzer can prove.`r`nSub Broken()`r`n    Dim n As Long`r`n    n = ""oops""`r`nEnd Sub`r`n")

        $clean = $components.Add(1)
        $clean.Name = 'CleanModule'
        $clean.CodeModule.AddFromString("Option Explicit`r`n`r`nPublic Function Total(ByVal values As Variant) As Double`r`n    Dim i As Long`r`n    Dim sum As Double`r`n`r`n    For i = LBound(values) To UBound(values)`r`n        sum = sum + values(i)`r`n    Next i`r`n`r`n    Total = sum`r`nEnd Function`r`n")

        $broken.CodeModule.CodePane.Show()
    }

    $excel.VBE.MainWindow.Visible = $true

    # Wait for everything the capture is meant to show, not just the first thing that finishes.
    #
    # Waiting on analysis alone once produced a screenshot taken a full second before the editing
    # surface had laid out, and the resulting image was read as proof that the surface was broken.
    # It was not: it had not loaded yet. Every condition below is a separate line in the log, so a
    # timeout says which part was missing rather than that something was.
    $logPattern = "shim-*-$($process.Id).log"
    $required = [ordered] @{
        'analysis' = 'analysis: \d+ finding'
        'surface'  = 'editor surface: ready'
    }

    $missing = @($required.Keys)
    while ((Get-Date) -lt $deadline -and $missing.Count -gt 0) {
        $log = Get-ChildItem $logRoot -Filter $logPattern -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if ($log) {
            $text = Get-Content $log.FullName -Raw -ErrorAction SilentlyContinue
            if ($text) {
                $missing = @($required.Keys | Where-Object { $text -notmatch $required[$_] })
            }
        }

        if ($missing.Count -gt 0) { Start-Sleep -Milliseconds 50 }
    }

    if ($missing.Count -gt 0) {
        Write-Warning "Captured before these reported ready: $($missing -join ', '). The image below shows that state, not a finished one."
    }

    # Maximise so the capture is not clipped by whatever else is on screen, and bring it forward so
    # reading the screen reads this window.
    $editor = $null
    while ($null -eq $editor -or $editor -eq [IntPtr]::Zero) {
        $editor = [Xlide.Shot]::TopLevel($process.Id, 'wndclass_desked_gsk')
        if ($editor -eq [IntPtr]::Zero) {
            if ((Get-Date) -ge $deadline) { throw 'The editor window never appeared.' }
            Start-Sleep -Milliseconds 50
        }
    }

    # Maximised so the capture is not a sliver, and brought forward as a courtesy. The capture does
    # not depend on either succeeding: it renders the window rather than reading the screen.
    [void] [Xlide.Shot]::ShowWindow($editor, 3)
    [void] [Xlide.Shot]::SetForegroundWindow($editor)

    # The surfaces have said they are ready, which means they have their content, not that a frame
    # has been presented. There is no event for that, so this is the one honest fixed settle, and it
    # is short because the conditions above now cover everything it used to be standing in for.
    Start-Sleep -Milliseconds 300

    # A window that is not answering messages will not answer a render request either, and asking
    # anyway is how a harness hangs. Checked rather than risked.
    if (-not [Xlide.Shot]::IsResponsive($editor)) { throw 'The editor is not responding, so it cannot be captured.' }

    $rect = New-Object Xlide.Shot+RECT
    [void] [Xlide.Shot]::GetWindowRect($editor, [ref] $rect)

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -le 0 -or $height -le 0) { throw 'The editor window has no size.' }

    $bitmap = New-Object System.Drawing.Bitmap $width, $height
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $rendered = $false
        try {
            $dc = $graphics.GetHdc()
            try { $rendered = [Xlide.Shot]::Render($editor, $dc) }
            finally { $graphics.ReleaseHdc($dc) }
        }
        finally {
            $graphics.Dispose()
        }

        if (-not $rendered) { throw 'The editor refused to render itself.' }

        $bitmap.Save($OutFile, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }

    Write-Host ("Captured {0} ({1} x {2})" -f $OutFile, $width, $height) -ForegroundColor Green

    exit 0
}
finally {
    if ($excel) {
        try { $excel.DisplayAlerts = $false } catch { }
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    if (-not $KeepOpen) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Get-Process xlide-engine -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "Leaving Excel $($process.Id) running."
    }
}
