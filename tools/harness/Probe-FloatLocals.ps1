# Can the Locals window be FLOATED through the object model, sized small, and does it still
# feed at break and on steps while tiny? Floating means genuinely on-screen, which is the one
# state the editor reliably feeds (lesson 25) - if a tiny floating Locals tracks steps, the
# themed panel gets its data source and the palette becomes a caption-stripped captive.
# Dev-harness script: uses VBProject and Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -Namespace Xlide -Name Float -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public delegate bool EnumProc(IntPtr h, IntPtr l);

public static string DescribeTopLevels(int processId)
{
    var text = new System.Text.StringBuilder();
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId || !IsWindowVisible(h)) { return true; }
        var cls = new System.Text.StringBuilder(128);
        GetClassNameW(h, cls, 128);
        var title = new System.Text.StringBuilder(128);
        GetWindowTextW(h, title, 128);
        RECT r;
        GetWindowRect(h, out r);
        text.AppendLine(string.Format("{0} '{1}' {2},{3} {4}x{5} hwnd={6:X}",
            cls, title, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, h.ToInt64()));
        return true;
    }, IntPtr.Zero);
    return text.ToString();
}

public static IntPtr FindTopLevelByTitle(int processId, string title)
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

# A module that breaks with locals worth watching.
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
    Stop
End Sub
'@)

function Get-Control([int]$id) { $vbe.CommandBars.FindControl(1, $id) }

$locals = $null
foreach ($window in $vbe.Windows) {
    if ($window.Type -eq 4) { $locals = $window }
}
if (-not $locals) { Write-Output 'NO LOCALS WINDOW'; exit 1 }

$locals.Visible = $true
Start-Sleep -Milliseconds 500

# 1: what is it docked in, and can LinkedWindows.Remove undock it?
try {
    $frame = $locals.LinkedWindowFrame
    Write-Output "linked frame: type $($frame.Type), caption '$($frame.Caption)'"
    try {
        $frame.LinkedWindows.Remove($locals)
        Write-Output 'undocked via LinkedWindows.Remove'
    } catch {
        Write-Output "LinkedWindows.Remove refused: $($_.Exception.Message)"
    }
} catch {
    Write-Output "no linked frame: $($_.Exception.Message)"
}
Start-Sleep -Milliseconds 500

# 2: can its geometry be set now?
try {
    $host = $locals.LinkedWindowFrame
    Write-Output "frame after undock: type $($host.Type), caption '$($host.Caption)'"
    $host.Left = 200; $host.Top = 200; $host.Width = 260; $host.Height = 160
    Write-Output "geometry set on the frame: $($host.Left),$($host.Top) $($host.Width)x$($host.Height)"
} catch {
    Write-Output "frame geometry refused: $($_.Exception.Message); trying the window itself"
    try {
        $locals.Left = 200; $locals.Top = 200; $locals.Width = 260; $locals.Height = 160
        Write-Output "geometry set on the window: $($locals.Left),$($locals.Top) $($locals.Width)x$($locals.Height)"
    } catch {
        Write-Output "window geometry refused too: $($_.Exception.Message)"
    }
}
Start-Sleep -Milliseconds 500

Write-Output '--- top-level windows now ---'
Write-Output ([Xlide.Float]::DescribeTopLevels($processId))

# 3: does the tiny floating window feed at break and on steps?
function Read-Counter([IntPtr]$handle) {
    if ($handle -eq [IntPtr]::Zero) { return 'no hwnd' }
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

$palette = [Xlide.Float]::FindTopLevelByTitle($processId, 'Locals')
Write-Output "palette hwnd: $palette"

$app.OnTime([DateTime]::Now.AddSeconds(1), 'BreakProbe.BreakHere')
$deadline = (Get-Date).AddSeconds(20)
while ((Get-Date) -lt $deadline) {
    $step = Get-Control 188
    if ($step -and $step.Enabled) { break }
    Start-Sleep -Milliseconds 200
}
Start-Sleep -Milliseconds 800
if ($palette -eq [IntPtr]::Zero) { $palette = [Xlide.Float]::FindTopLevelByTitle($processId, 'Locals') }
Write-Output "entry (expect counter=1): $(Read-Counter $palette)"

(Get-Control 188).Execute()
Start-Sleep -Milliseconds 900
Write-Output "after step (expect 1, line counter=2 pending): $(Read-Counter $palette)"

(Get-Control 188).Execute()
Start-Sleep -Milliseconds 900
Write-Output "after second step (expect 2): $(Read-Counter $palette)"

$reset = Get-Control 228
if ($reset -and $reset.Enabled) { $reset.Execute() }
Start-Sleep -Milliseconds 500
$locals.Visible = $false
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}
Write-Output 'done'
