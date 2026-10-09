# Focused cross-process/cross-host persistence regression for issue #76.
# Run against a published Debug shim. Owns its Office processes and restores user preferences.
[CmdletBinding()]
param([switch] $SkipPowerPoint)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$preferenceRoot = if ($env:XLIDE_TEST_PREFERENCES_ROOT) { $env:XLIDE_TEST_PREFERENCES_ROOT } else { $env:LOCALAPPDATA }
$store = Join-Path $preferenceRoot 'xlide_vbide'
$expected = Join-Path $repoRoot 'artifacts\issue76-expected.json'
$owned = [Collections.Generic.Dictionary[int,datetime]]::new()
if (Get-Process EXCEL,WINWORD,MSACCESS,POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close Office hosts before this preference-isolated harness; it restores the user files afterward.'
}
$backups = @{}
foreach ($name in @('settings.json', 'ui-state.json')) {
    $path = Join-Path $store $name
    $backups[$name] = if (Test-Path -LiteralPath $path) { [IO.File]::ReadAllBytes($path) } else { $null }
}

Add-Type -Namespace XlidePreferencesHarness -Name Windows -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr data);
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr window, out int process);
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr window, System.Text.StringBuilder name, int count);
[DllImport("user32.dll")] public static extern bool GetWindowPlacement(IntPtr window, ref Placement placement);
[DllImport("user32.dll")] public static extern bool SetWindowPlacement(IntPtr window, ref Placement placement);
[DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr window, uint message, IntPtr w, IntPtr l);
[DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr context);
[StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
[StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
[StructLayout(LayoutKind.Sequential)] public struct Placement { public int Length,Flags,Show; public Point Min,Max; public Rect Normal; }
public delegate bool EnumProc(IntPtr window, IntPtr data);
public static IntPtr Find(int pid, string wanted) {
    IntPtr found=IntPtr.Zero;
    EnumWindows((window,data) => { int owner; GetWindowThreadProcessId(window,out owner);
        var name=new System.Text.StringBuilder(100); GetClassNameW(window,name,100);
        if(owner==pid && name.ToString()==wanted) {found=window; return false;} return true;
    }, IntPtr.Zero); return found;
}
public static string Geometry(IntPtr window) {
    var p=new Placement {Length=Marshal.SizeOf(typeof(Placement))};
    if(!GetWindowPlacement(window,ref p)) throw new Exception("No window placement");
    return String.Format("{0},{1},{2},{3},{4}",p.Normal.Left,p.Normal.Top,p.Normal.Right,p.Normal.Bottom,p.Show==3);
}
public static void Place(IntPtr window,int left,int top,int width,int height,bool maximized) {
    var p=new Placement {Length=Marshal.SizeOf(typeof(Placement)),Show=maximized?3:1,
        Normal=new Rect {Left=left,Top=top,Right=left+width,Bottom=top+height}};
    if(!SetWindowPlacement(window,ref p)) throw new Exception("Placement refused");
    SendMessageW(window,0x232,IntPtr.Zero,IntPtr.Zero);
}
'@
[void] [XlidePreferencesHarness.Windows]::SetProcessDpiAwarenessContext([IntPtr](-4))

function Launch([string] $HostName) {
    if ($HostName -eq 'powerpoint') {
        if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'PowerPoint is already running; use a disposable host.' }
        $process = Start-Process -FilePath "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE" -WindowStyle Hidden -PassThru
        $owned.Add($process.Id, $process.StartTime)
        $app = $null
        $deadline = (Get-Date).AddSeconds(30)
        while (!$app -and (Get-Date) -lt $deadline) {
            try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
            catch { Start-Sleep -Milliseconds 200 }
        }
        if (!$app) { throw 'PowerPoint did not expose its running application.' }
        $presentation = $app.Presentations.Add(-1)
        $module = $presentation.VBProject.VBComponents.Add(1)
        $module.CodeModule.AddFromString("Public Sub LayoutProbe()`r`nEnd Sub")
        $module.CodeModule.CodePane.Show()
        $app.VBE.MainWindow.Visible = $true
        foreach($com in @($module,$presentation,$app)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($com)}
        return $process.Id
    }
    $output = switch ($HostName) {
        'excel' { & (Join-Path $PSScriptRoot 'Start-Excel.ps1') -Workbook 'artifacts\fixtures\DebugFixture.xlsm' -Separate *>&1 }
        'word' { & (Join-Path $PSScriptRoot 'Start-Word.ps1') -Document 'artifacts\fixtures\WordFixture.docm' *>&1 }
        'access' { & (Join-Path $PSScriptRoot 'Start-Access.ps1') -Database 'artifacts\fixtures\AccessFixture.accdb' *>&1 }
    }
    $pidLine = $output | ForEach-Object { "$_" } | Select-String '^pid=(\d+)$' | Select-Object -Last 1
    if (!$pidLine) { $output | Out-Host; throw "No $HostName process was launched" }
    $hostPid = [int] $pidLine.Matches[0].Groups[1].Value
    $owned.Add($hostPid, (Get-Process -Id $hostPid).StartTime)
    return $hostPid
}
function Close-Owned([int] $HostPid) {
    $process = Get-Process -Id $HostPid -ErrorAction SilentlyContinue
    if ($owned.ContainsKey($HostPid) -and $process -and $process.StartTime -eq $owned[$HostPid]) {
        Stop-Process -Id $HostPid -Force -ErrorAction SilentlyContinue
    }
    $owned.Remove($HostPid) | Out-Null
    Wait-Process -Id $HostPid -Timeout 10 -ErrorAction SilentlyContinue
}
function Page-Probe([int] $HostPid, [string] $Phase) {
    & node (Join-Path $PSScriptRoot 'user-preferences.mjs') $HostPid $Phase $expected
    if ($LASTEXITCODE -ne 0) { throw "Preference UI check failed in process $HostPid" }
}
function Open-Palette([int] $HostPid) {
    & node (Join-Path $PSScriptRoot 'xlide-api.mjs') --pid $HostPid command objectBrowser | Out-Null
    $until = (Get-Date).AddSeconds(10)
    do {
        $window = [XlidePreferencesHarness.Windows]::Find($HostPid, 'XlidePalette')
        if ($window -ne [IntPtr]::Zero) { return $window }
        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -lt $until)
    throw 'Object Browser did not open'
}

try {
    $first = Launch 'excel'
    $frame = [XlidePreferencesHarness.Windows]::Find($first, 'wndclass_desked_gsk')
    [XlidePreferencesHarness.Windows]::Place($frame,60,50,1100,760,$false)
    Page-Probe $first 'seed'
    $frameExpected = [XlidePreferencesHarness.Windows]::Geometry($frame)
    $palette = Open-Palette $first
    [XlidePreferencesHarness.Windows]::Place($palette,180,120,680,500,$false)
    $paletteExpected = [XlidePreferencesHarness.Windows]::Geometry($palette)
    Close-Owned $first

    $hosts = @('excel','word','access')
    if (!$SkipPowerPoint) { $hosts += 'powerpoint' }
    foreach ($hostName in $hosts) {
        $hostPid = Launch $hostName
        Page-Probe $hostPid 'verify'
        $frame = [XlidePreferencesHarness.Windows]::Find($hostPid, 'wndclass_desked_gsk')
        $frameNow = [XlidePreferencesHarness.Windows]::Geometry($frame)
        if ($frameNow -ne $frameExpected) { throw "$hostName editor geometry differs: expected $frameExpected, restored $frameNow" }
        $palette = Open-Palette $hostPid
        $paletteNow = [XlidePreferencesHarness.Windows]::Geometry($palette)
        if ($paletteNow -ne $paletteExpected) { throw "$hostName palette geometry differs: expected $paletteExpected, restored $paletteNow" }
        Write-Host "PASS: $hostName restored editor and Object Browser window bounds."
        Close-Owned $hostPid
    }
    $hostPid = Launch 'excel'
    Page-Probe $hostPid 'verify'
    $frame = [XlidePreferencesHarness.Windows]::Find($hostPid, 'wndclass_desked_gsk')
    [XlidePreferencesHarness.Windows]::Place($frame,60,50,1100,760,$true)
    $maxFrame = [XlidePreferencesHarness.Windows]::Geometry($frame)
    $palette = Open-Palette $hostPid
    [XlidePreferencesHarness.Windows]::Place($palette,180,120,680,500,$true)
    $maxPalette = [XlidePreferencesHarness.Windows]::Geometry($palette)
    Close-Owned $hostPid
    $hostPid = Launch 'word'
    Page-Probe $hostPid 'verify'
    $frame = [XlidePreferencesHarness.Windows]::Find($hostPid, 'wndclass_desked_gsk')
    $palette = Open-Palette $hostPid
    if ([XlidePreferencesHarness.Windows]::Geometry($frame) -ne $maxFrame -or
        [XlidePreferencesHarness.Windows]::Geometry($palette) -ne $maxPalette) {
        throw 'Maximized window state or normal restore bounds were lost across hosts'
    }
    Write-Host 'PASS: maximized editor and Object Browser restore across Excel and Word.'
    Close-Owned $hostPid
    Write-Host 'RESULT: PASS - shared user preferences survive fresh Office processes.'
}
finally {
    foreach ($hostPid in @($owned.Keys)) { Close-Owned $hostPid }
    foreach ($name in $backups.Keys) {
        $path = Join-Path $store $name
        if ($null -ne $backups[$name]) { [IO.File]::WriteAllBytes($path, $backups[$name]) }
        elseif (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
    }
}
