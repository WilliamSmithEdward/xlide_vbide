# Does an INVISIBLE Locals palette still feed? Layered windows render into their own
# surface regardless of occlusion, so alpha-0 may keep the editor's feed alive while the
# window cannot be seen - which would let the THEMED panel be the only visible Locals.
# Tests alpha 1, alpha 0, and off-screen, reading through break and step each time.
# Dev-harness script: uses VBProject and Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -Namespace Xlide -Name Ghost -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern long GetWindowLongPtrW(IntPtr h, int i);
[DllImport("user32.dll")] public static extern long SetWindowLongPtrW(IntPtr h, int i, long v);
[DllImport("user32.dll")] public static extern bool SetLayeredWindowAttributes(IntPtr h, uint key, byte alpha, uint flags);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int w, int hgt, uint f);
public delegate bool EnumProc(IntPtr h, IntPtr l);

public static IntPtr FindByTitle(int processId, string title)
{
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        var text = new System.Text.StringBuilder(128);
        GetWindowTextW(h, text, 128);
        if (text.ToString() == title) { found = h; return false; }
        return true;
    }, IntPtr.Zero);
    return found;
}
'@

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE
$processId = (Get-Process EXCEL | Select-Object -First 1).Id

$components = $app.ActiveWorkbook.VBProject.VBComponents
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}
$component = $components.Add(1)
$component.Name = 'BreakProbe'
$component.CodeModule.AddFromString(@'
Sub BreakHere()
    Dim counter As Long
    counter = 1
    Stop
    counter = 2
    counter = 3
    counter = 4
    Stop
End Sub
'@)

function Get-Control([int]$id) { $vbe.CommandBars.FindControl(1, $id) }

$locals = $null
foreach ($window in $vbe.Windows) { if ($window.Type -eq 4) { $locals = $window } }
$locals.Visible = $true
Start-Sleep -Milliseconds 400

# Already floating from the earlier probe; float it if a fresh session redocked it.
try { $locals.LinkedWindowFrame.LinkedWindows.Remove($locals) } catch { }
try { $locals.Left = 300; $locals.Top = 300; $locals.Width = 240; $locals.Height = 150 } catch { }
Start-Sleep -Milliseconds 400

$palette = [Xlide.Ghost]::FindByTitle($processId, 'Locals')
Write-Output "palette hwnd: $palette"
if ($palette -eq [IntPtr]::Zero) { Write-Output 'NO PALETTE'; exit 1 }

# GWL_EXSTYLE = -20; WS_EX_LAYERED = 0x80000, WS_EX_TRANSPARENT = 0x20, WS_EX_NOACTIVATE = 0x8000000.
$exStyle = [Xlide.Ghost]::GetWindowLongPtrW($palette, -20)
[void] [Xlide.Ghost]::SetWindowLongPtrW($palette, -20, $exStyle -bor 0x80000 -bor 0x20 -bor 0x8000000)
[void] [Xlide.Ghost]::SetLayeredWindowAttributes($palette, 0, 1, 2)  # LWA_ALPHA, alpha 1: a ghost
Write-Output 'ghosted at alpha 1, click-through'

function Read-Counter {
    param([IntPtr]$handle)
    $element = [System.Windows.Automation.AutomationElement]::FromHandle($handle)
    foreach ($child in $element.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)) {
        $c = $child.Current
        if ($c.ControlType.ProgrammaticName -eq 'ControlType.ListItem' -and $c.Name -like '*counter*') {
            return $c.Name.Trim()
        }
    }
    return 'no counter row'
}

$app.OnTime([DateTime]::Now.AddSeconds(1), 'BreakProbe.BreakHere')
$deadline = (Get-Date).AddSeconds(20)
while ((Get-Date) -lt $deadline) {
    $step = Get-Control 188
    if ($step -and $step.Enabled) { break }
    Start-Sleep -Milliseconds 200
}
Start-Sleep -Milliseconds 800
Write-Output "alpha 1, entry (expect 1): $(Read-Counter $palette)"

(Get-Control 188).Execute()
Start-Sleep -Milliseconds 900
(Get-Control 188).Execute()
Start-Sleep -Milliseconds 900
Write-Output "alpha 1, after two steps (expect 2): $(Read-Counter $palette)"

# Fully invisible.
[void] [Xlide.Ghost]::SetLayeredWindowAttributes($palette, 0, 0, 2)
(Get-Control 188).Execute()
Start-Sleep -Milliseconds 900
Write-Output "alpha 0, after third step (expect 3): $(Read-Counter $palette)"

# Off virtual screen entirely, still alpha 0.
[void] [Xlide.Ghost]::SetWindowPos($palette, [IntPtr]::Zero, -20000, -20000, 0, 0, 0x0015)
(Get-Control 188).Execute()
Start-Sleep -Milliseconds 900
Write-Output "off-screen, after fourth step (expect 4): $(Read-Counter $palette)"

$reset = Get-Control 228
if ($reset -and $reset.Enabled) { $reset.Execute() }
Start-Sleep -Milliseconds 400

# Restore: on-screen, opaque, styles back, hidden.
[void] [Xlide.Ghost]::SetWindowPos($palette, [IntPtr]::Zero, 300, 300, 0, 0, 0x0015)
[void] [Xlide.Ghost]::SetLayeredWindowAttributes($palette, 0, 255, 2)
[void] [Xlide.Ghost]::SetWindowLongPtrW($palette, -20, $exStyle)
$locals.Visible = $false
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}
Write-Output 'restored and hidden; done'
