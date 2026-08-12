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
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
public delegate bool EnumProc(IntPtr h, IntPtr l);

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

$excelProcess = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excelProcess) { Write-Output 'RESULT: FAIL - no Excel is running'; exit 1 }
$processId = $excelProcess.Id

try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    Write-Output "RESULT: FAIL - Excel $processId is not reachable through the running object table"; exit 1
}
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

if ($broken -eq 0) {
    Write-Output 'RESULT: PASS - three close and reopen cycles, Excel standing after each'
    exit 0
}
Write-Output "RESULT: FAIL - $broken expectation(s) broke across the cycles; a dead host here is lesson 27 returned"
exit 1
