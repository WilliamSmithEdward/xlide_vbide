# Isolated two-process live pane docking regression. Only closes its own Excel processes.
param([switch] $RestoreDefaults)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$previousRoot = $env:XLIDE_TEST_PREFERENCES_ROOT
$env:XLIDE_TEST_PREFERENCES_ROOT = Join-Path $repoRoot ('artifacts\pane-split-' + [guid]::NewGuid().ToString('N'))
$owned = @{}
Add-Type -Namespace XlidePaneSplitHarness -Name Windows -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr data);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr window, out int process);
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr window, System.Text.StringBuilder name, int count);
[DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr window, int x, int y, int width, int height, bool repaint);
public delegate bool EnumProc(IntPtr window, IntPtr data);
public static void Size(int pid) {
    EnumWindows((window,data) => { int owner; GetWindowThreadProcessId(window,out owner);
        var name=new System.Text.StringBuilder(100); GetClassNameW(window,name,100);
        if(owner==pid && name.ToString()=="wndclass_desked_gsk") {MoveWindow(window,60,50,1100,850,true); return false;} return true;
    }, IntPtr.Zero);
}
'@
try {
    foreach ($phase in @('drag', 'restore')) {
        $output = & (Join-Path $PSScriptRoot 'Start-Excel.ps1') -Workbook 'artifacts\fixtures\DebugFixture.xlsm' -Separate *>&1
        $line = $output | ForEach-Object { "$_" } | Select-String '^pid=(\d+)$' | Select-Object -Last 1
        if (!$line) { $output | Out-Host; throw 'Excel did not launch' }
        $hostPid = [int]$line.Matches[0].Groups[1].Value
        $owned[$hostPid] = (Get-Process -Id $hostPid).StartTime
        [XlidePaneSplitHarness.Windows]::Size($hostPid)
        $probe = if ($RestoreDefaults) { 'restore-pane-defaults.mjs' } else { 'editor-pane-split.mjs' }
        & node (Join-Path $PSScriptRoot $probe) $hostPid $phase
        if ($LASTEXITCODE) { throw "Pane split check failed: $phase" }
        $process = Get-Process -Id $hostPid -ErrorAction SilentlyContinue
        if ($process -and $process.StartTime -eq $owned[$hostPid]) { Stop-Process -Id $hostPid -Force }
        Wait-Process -Id $hostPid -Timeout 10 -ErrorAction SilentlyContinue
        $owned.Remove($hostPid)
    }
    if ($RestoreDefaults) { Write-Host 'RESULT: PASS - pane defaults confirmation and restart restoration.' }
    else { Write-Host 'RESULT: PASS - editor-adjacent pane split and restart restoration.' }
} finally {
    foreach ($hostPid in @($owned.Keys)) {
        $process = Get-Process -Id $hostPid -ErrorAction SilentlyContinue
        if ($process -and $process.StartTime -eq $owned[$hostPid]) { Stop-Process -Id $hostPid -Force }
    }
    $env:XLIDE_TEST_PREFERENCES_ROOT = $previousRoot
}
