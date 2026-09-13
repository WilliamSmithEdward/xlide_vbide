# Does closing the VBE window leave Excel standing, three times in a row?
#
# Guards the 2026-08-04 crash (lesson 27): the frame-hide event drove object-model calls
# into the editor's own close handling, and Excel died with VBE7/ntdll/shim faulting by
# turn. Runs against whatever editor session is open. Expected: excel alive=True on all
# six lines, frame visibility alternating False/True - and enforced, not just printed.
#
# Closes via WM_SYSCOMMAND SC_CLOSE, the same message the developer's click on the X
# sends; that window message IS the subject, so it cannot go through the api. The reopen
# is Excel executing its own ribbon button (ExecuteMso), which needs no VBA project
# trust - the same trick Start-Excel.ps1 uses, for the same reason.
#
# THE GATE RUNS THIS LAST, after everything else in the live half, because what it guards
# is Excel dying: a regression here takes the session with it by design, and nothing may
# be scheduled behind it.
$ErrorActionPreference = 'Continue'

Add-Type -Namespace Xlide -Name Close -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object window);
public delegate bool EnumProc(IntPtr h, IntPtr l);

// OBJID_NATIVEOM: a worksheet window hands back the workbook window, and its Application is the
// one belonging to THIS process - which the running object table cannot promise.
const uint NativeObjectModel = 0xFFFFFFF0u;

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
    return AccessibleObjectFromWindow(sheet, NativeObjectModel, ref dispatch, out window) == 0 ? window : null;
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

# The instance under test is the one the caller named through XLIDE_PID when it did; the first
# EXCEL in the process list may be an automation Excel of somebody else's (2026-09-08).
$excelProcess = if ($env:XLIDE_PID) { Get-Process -Id ([int]$env:XLIDE_PID) -ErrorAction SilentlyContinue }
if (-not $excelProcess) { $excelProcess = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1 }
if (-not $excelProcess) { Write-Output 'RESULT: FAIL - no Excel is running'; exit 1 }
$processId = $excelProcess.Id

# THROUGH ITS WINDOW, not the running object table. The ROT hands back whichever Excel registered
# itself, which is not necessarily the one XLIDE_PID names: with a second Excel running the add-in
# - a recovery copy, another session's fixture - this test closed the aimed editor and pressed the
# ribbon button on the OTHER instance, so the frame never came back and every cycle broke
# (2026-09-13). Start-Excel.ps1 attaches this way for the same reason.
$window = [Xlide.Close]::WorkbookWindowOf($processId)
if (-not $window) {
    Write-Output "RESULT: FAIL - Excel $processId has no worksheet window to reach its object model through"; exit 1
}

$app = $window.Application
Write-Output "excel $processId"

$broken = 0
for ($cycle = 1; $cycle -le 3; $cycle++) {
    $frame = [Xlide.Close]::FrameOf($processId)
    if ($frame -eq [IntPtr]::Zero) {
        Write-Output "cycle ${cycle}: NO FRAME"
        $broken += 1
        break
    }

    # SC_CLOSE, the same as the developer clicking the X.
    [void] [Xlide.Close]::SendMessage($frame, 0x0112, [IntPtr]0xF060, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 1500

    $alive = $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)
    $visible = [Xlide.Close]::IsWindowVisible($frame)
    Write-Output "cycle ${cycle}: closed -> excel alive=$alive, frame visible=$visible"
    if (-not $alive) { $broken += 1; break }
    if ($visible) { $broken += 1 }

    # The reopen: Excel pressing its own Developer > Visual Basic button.
    $app.CommandBars.ExecuteMso('VisualBasic')
    Start-Sleep -Milliseconds 1500
    $alive = $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)
    $visible = [Xlide.Close]::IsWindowVisible($frame)
    Write-Output "cycle ${cycle}: reopened -> excel alive=$alive, frame visible=$visible"
    if (-not $alive) { $broken += 1; break }
    if (-not $visible) { $broken += 1 }
}

# RELEASED BEFORE THE VERDICT. A wrapper finalised after its host has gone is what DCOM answers
# by starting a fresh Excel, and this test exists because hosts die here (lesson 27).
foreach ($held in $app, $window) {
    if ($held) { [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($held) }
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

if ($broken -eq 0) {
    Write-Output 'RESULT: PASS - three close and reopen cycles, Excel standing after each'
    exit 0
}
Write-Output "RESULT: FAIL - $broken expectation(s) broke across the cycles; a dead host here is lesson 27 returned"
exit 1
