# End to end: does a REAL watch reach the themed Watch panel through break and steps?
# There is no object-model API for watches, so this drives the native Add Watch dialog
# (command 1820) with a helper process: the Execute call blocks on the modal, the helper
# fills the Expression edit (control 4853) and presses OK (control 1). Then it breaks and
# reads ONLY the shim log - never the ghost's accessibility tree from outside, which can
# reset the project mid-break when the in-process reader is alive (lesson 33).
# A second "watch: 1 row(s)" push after stepping IS the value-tracking proof: pushes only
# happen when the reading's content changes. The watch dies with the Excel session.
# Run tools\dev.ps1 -KeepOpen first. Prints PASS or FAIL.
# Dev-harness script: uses VBProject and Application.VBE per decision 10's harness exception.
$ErrorActionPreference = 'Continue'

$watcher = @'
$ErrorActionPreference = "Continue"
Add-Type -Namespace X -Name W -MemberDefinition @"
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr h, int id);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, string l);
[DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
public delegate bool EnumProc(IntPtr h, IntPtr l);
public static IntPtr FindDialog(int pid, string title) {
  IntPtr found = IntPtr.Zero;
  EnumWindows((h, l) => {
    int o; GetWindowThreadProcessId(h, out o);
    if (o != pid || !IsWindowVisible(h)) { return true; }
    var s = new System.Text.StringBuilder(64); GetWindowTextW(h, s, 64);
    if (s.ToString() == title) { found = h; return false; }
    return true;
  }, IntPtr.Zero);
  return found;
}
"@
$excelId = [int]$args[0]
$expression = $args[1]
$deadline = (Get-Date).AddSeconds(15)
$dialog = [IntPtr]::Zero
while ((Get-Date) -lt $deadline -and $dialog -eq [IntPtr]::Zero) {
  $dialog = [X.W]::FindDialog($excelId, "Add Watch")
  Start-Sleep -Milliseconds 150
}
if ($dialog -eq [IntPtr]::Zero) { exit 1 }
$edit = [X.W]::GetDlgItem($dialog, 4853)
$ok = [X.W]::GetDlgItem($dialog, 1)
if ($edit -eq [IntPtr]::Zero -or $ok -eq [IntPtr]::Zero) { exit 2 }
# Typed, not set: the dialog answers "Empty watch expression" to text planted by
# WM_SETTEXT (measured 2026-08-05) - it only believes text that arrived as keystrokes.
foreach ($ch in $expression.ToCharArray()) {
  [void][X.W]::SendMessageW($edit, 0x0102, [IntPtr][int]$ch, [IntPtr]::Zero)  # WM_CHAR
  Start-Sleep -Milliseconds 20
}
Start-Sleep -Milliseconds 150
[void][X.W]::SendMessageW($ok, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
exit 0
'@

$app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$vbe = $app.VBE
$excelId = (Get-Process EXCEL | Select-Object -First 1).Id

$log = Get-ChildItem "$env:LOCALAPPDATA\xlide_vbide\logs" -Filter "shim-*-$excelId.log" |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1

$components = $app.ActiveWorkbook.VBProject.VBComponents
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}
$component = $components.Add(1)
$component.Name = 'BreakProbe'
$component.CodeModule.AddFromString(@'
Sub BreakHere()
    Dim counter As Long
    Dim label As String
    counter = 1
    label = "alpha"
    Stop
    counter = 2
    label = "beta"
    Stop
End Sub
'@)

# The dialog seeds its Module and Procedure context from the active pane's selection, so
# the caret goes inside BreakHere before the dialog opens.
$pane = $component.CodeModule.CodePane
$pane.Show()
$pane.SetSelection(4, 1, 4, 1)

$watcherFile = Join-Path $env:TEMP 'xlide-watchpanel-watcher.ps1'
Set-Content $watcherFile $watcher
$helper = Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', $watcherFile, $excelId, 'counter') -PassThru -WindowStyle Hidden

function Get-Control([int]$id) { $vbe.CommandBars.FindControl(1, $id) }
(Get-Control 1820).Execute()
$helper.WaitForExit()
if ($helper.ExitCode -ne 0) {
    Write-Output "RESULT: FAIL - the Add Watch dialog could not be driven (watcher exit $($helper.ExitCode))"
    foreach ($existing in @($components)) {
        if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
    }
    return
}

$before = (Get-Content $log.FullName | Measure-Object -Line).Lines

$app.OnTime([DateTime]::Now.AddSeconds(1), 'BreakProbe.BreakHere')

# Quiet until the break is well established; the only client of the ghost is the shim.
Start-Sleep -Seconds 6

# Two steps move counter to 2: the watch's value changes, which must produce a second push
# (pushes only happen when the reading's content changes). External-command breaks poll at
# one second, so the settles stay comfortably above the cadence.
(Get-Control 188).Execute()
Start-Sleep -Milliseconds 1500
(Get-Control 188).Execute()
Start-Sleep -Seconds 3

$reset = Get-Control 228
if ($reset -and $reset.Enabled) { $reset.Execute() }

# The exit clear rides the next poll after the reset; settle past it before reading.
Start-Sleep -Seconds 2
foreach ($existing in @($components)) {
    if ($existing.Name -eq 'BreakProbe') { $components.Remove($existing) }
}

$since = Get-Content $log.FullName | Select-Object -Skip $before
Write-Output '--- watch lines ---'
$since | Where-Object { $_ -match 'watch' }

$pushes = @($since | Where-Object { $_ -match 'watch: [1-9]\d* row' })
$cleared = @($since | Where-Object { $_ -match '"type":"setWatches","stopped":false' })
$verdict = @()
if ($pushes.Count -lt 1) { $verdict += 'no non-empty watch push during the break' }
elseif ($pushes.Count -lt 2) { $verdict += 'no second push followed the value-changing steps' }
if ($cleared.Count -lt 1) { $verdict += 'no clear at break exit' }
if ($verdict.Count -eq 0) {
    Write-Output "RESULT: PASS - $($pushes.Count) non-empty push(es), value tracked the steps, cleared at exit"
} else {
    Write-Output "RESULT: FAIL - $($verdict -join '; ') (pushes $($pushes.Count), clears $($cleared.Count))"
}
