# Does closing the VBE window leave Excel standing, three times in a row?
#
# Guards the 2026-08-04 crash (lesson 27): the frame-hide event drove object-model calls
# into the editor's own close handling, and Excel died with VBE7/ntdll/shim faulting by
# turn. Run tools\dev.ps1 -KeepOpen first, then this. Expected: excel alive=True on all
# six lines, frame visibility alternating False/True.
# Closes via WM_SYSCOMMAND SC_CLOSE, reopens via the object model, and checks each cycle.
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

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE
$processId = (Get-Process EXCEL | Select-Object -First 1).Id
Write-Output "excel $processId"

for ($cycle = 1; $cycle -le 3; $cycle++) {
    $frame = [Xlide.Close]::FrameOf($processId)
    if ($frame -eq [IntPtr]::Zero) { Write-Output "cycle ${cycle}: NO FRAME"; break }

    # SC_CLOSE, the same as the developer clicking the X.
    [void] [Xlide.Close]::SendMessage($frame, 0x0112, [IntPtr]0xF060, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 1500

    $alive = $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)
    $visible = [Xlide.Close]::IsWindowVisible($frame)
    Write-Output "cycle ${cycle}: closed -> excel alive=$alive, frame visible=$visible"
    if (-not $alive) { break }

    $vbe.MainWindow.Visible = $true
    Start-Sleep -Milliseconds 1500
    $alive = $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)
    $visible = [Xlide.Close]::IsWindowVisible($frame)
    Write-Output "cycle ${cycle}: reopened -> excel alive=$alive, frame visible=$visible"
    if (-not $alive) { break }
}

Write-Output 'done'
