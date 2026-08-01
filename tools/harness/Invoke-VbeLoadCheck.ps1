<#
.SYNOPSIS
    Starts an owned Excel instance and reports whether the add-in loaded and ran inside it.

.DESCRIPTION
    The check is driven by what the add-in itself records from inside the host, not by automating
    the host from outside. That is both faster and better evidence.

    Attaching to Excel from outside costs ten to forty seconds, because Excel does not publish
    itself for automation until long after it is visibly running. None of that wait tells us
    anything about the add-in: the host loads add-ins during its own startup regardless of whether
    anything is attached. Reading what the add-in wrote skips the wait entirely and observes the
    thing being tested rather than a proxy for it.

    Pass -Deep to additionally attach and ask the editor whether it considers the add-in connected.
    That is worth doing before a release and not worth doing on every edit.

    Process safety: Excel launched here is a real child process, started with a document so it
    initialises promptly, recorded by identity, and terminated only by that identity. An instance
    the developer opened is never touched.
#>
[CmdletBinding()]
param(
    # Upper bound on how long to wait for the add-in to load. A deadline, not a delay.
    [int] $TimeoutSeconds = 60,

    # Leave Excel running afterwards so the next run can reuse it.
    [switch] $KeepOpen,

    # Reuse the instance a previous run left open. A repeat check then costs a fraction of a second
    # instead of restarting the host. The reused instance still holds the previous build of the
    # shim, because a host keeps an add-in library open for its lifetime, so use this after changing
    # anything except the shim.
    [switch] $Reuse,

    # Also attach to the host and ask the editor whether the add-in is connected.
    [switch] $Deep,

    # Proceed even when the developer already has Excel open.
    [switch] $AllowExistingExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$logRoot = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'
$ownedPidFile = Join-Path $env:LOCALAPPDATA 'xlide_vbide\harness-excel.pid'
$startedAt = Get-Date

$phaseClock = [System.Diagnostics.Stopwatch]::StartNew()
$phases = New-Object System.Collections.Generic.List[string]

function Complete-Phase {
    param([string] $Name)
    $phases.Add(('{0,-24} {1,7:N2}s' -f $Name, $phaseClock.Elapsed.TotalSeconds))
    $phaseClock.Restart()
}

Add-Type -Namespace Xlide -Name Attach -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

delegate bool EnumProc(IntPtr h, IntPtr l);

// OBJID_NATIVEOM. Asking a worksheet window for its native object model yields the workbook
// window, and its Application, without going through the running object table.
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
'@

function Get-ExcelProcessIds {
    $found = New-Object System.Collections.Generic.List[int]
    foreach ($process in (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue)) {
        $found.Add([int] $process.Id)
    }
    return , $found.ToArray()
}

function Get-RecordedInstanceId {
    if (-not (Test-Path $ownedPidFile)) { return 0 }
    $recorded = 0
    if ([int]::TryParse((Get-Content $ownedPidFile -Raw).Trim(), [ref] $recorded)) { return $recorded }
    return 0
}

function Find-ExcelExecutable {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE"
    )
    foreach ($candidate in $candidates) { if (Test-Path $candidate) { return $candidate } }

    $fromRegistry = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)'
    if ($fromRegistry -and (Test-Path $fromRegistry)) { return $fromRegistry }

    throw 'Could not locate EXCEL.EXE.'
}

function Reset-ExcelResiliency {
    # A harness terminates Excel by design, and Excel treats termination as a crash. On the next
    # start it offers document recovery and disables items it blames, both of which appear before
    # the host is usable and produce a failure unrelated to what is being tested.
    foreach ($version in @('16.0', '15.0')) {
        $resiliency = "HKCU:\Software\Microsoft\Office\$version\Excel\Resiliency"
        if (-not (Test-Path $resiliency)) { continue }
        foreach ($child in @('DocumentRecovery', 'DisabledItems', 'StartupItems', 'CrashingAddinList')) {
            $path = Join-Path $resiliency $child
            if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Reset-AddInLoadBehavior {
    # When an add-in fails to connect, the editor rewrites LoadBehavior to 0 so it will not try
    # again. Every run after the first failure then silently skips the add-in, which looks like a
    # different fault entirely.
    $key = 'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Xlide.VbeAddIn'
    if (-not (Test-Path $key)) {
        Write-Warning 'The add-in is not registered. Run tools\dev.ps1 -NoRun first.'
        return
    }

    $current = (Get-ItemProperty $key -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
    if ($current -ne 3) {
        Write-Host "LoadBehavior was $current (the editor disabled the add-in after a failed connect). Restoring 3."
        Set-ItemProperty $key -Name LoadBehavior -Value 3 -Type DWord
    }
}

$preExisting = Get-ExcelProcessIds
if ($preExisting.Length -gt 0 -and -not $AllowExistingExcel) {
    $recorded = Get-RecordedInstanceId
    $allOurs = $Reuse -and $recorded -gt 0 -and @($preExisting | Where-Object { $_ -ne $recorded }).Length -eq 0
    if (-not $allOurs) {
        Write-Error "Excel is already running (PID $($preExisting -join ', ')). Close it, or pass -AllowExistingExcel."
        exit 2
    }
}

if (-not $Reuse) {
    Reset-ExcelResiliency
    Reset-AddInLoadBehavior
}

$excelProcess = $null
$reused = $false
$excel = $null

try {
    if ($Reuse) {
        $recorded = Get-RecordedInstanceId
        $candidate = if ($recorded -gt 0) { Get-Process -Id $recorded -ErrorAction SilentlyContinue } else { $null }
        if ($candidate -and $candidate.ProcessName -eq 'EXCEL') {
            $excelProcess = $candidate
            $reused = $true
            Write-Host "Reusing the instance from an earlier run (process $recorded)."
        }
    }

    if (-not $reused) {
        # Started with a document so the host initialises promptly, and as an ordinary process so
        # add-ins load. A host created through automation runs in embedding mode and loads none.
        $scratch = Join-Path $PSScriptRoot 'fixtures\scratch.xlsm'
        if (-not (Test-Path $scratch)) { & (Join-Path $PSScriptRoot 'New-ScratchWorkbook.ps1') | Out-Null }

        $excelProcess = Start-Process -FilePath (Find-ExcelExecutable) -ArgumentList $scratch -PassThru
        New-Item -ItemType Directory -Force -Path (Split-Path $ownedPidFile) | Out-Null
        Set-Content -Path $ownedPidFile -Value $excelProcess.Id

        Write-Host "Started Excel as process $($excelProcess.Id)."
    }

    Complete-Phase 'start host'

    # Attach through the host's own window rather than through the running object table.
    #
    # This matters more than it looks. A host publishes itself in the running object table lazily,
    # ten to forty seconds after it is visibly ready, and waiting for that was almost the entire
    # cost of this check. Asking a worksheet window for its native object model answers as soon as
    # the window exists, which is under a second. It is also more precise: it names the instance by
    # window, so there is no possibility of adopting a different one.
    $attachDeadline = (Get-Date).AddSeconds(60)
    $workbookWindow = $null
    while ($null -eq $workbookWindow -and (Get-Date) -lt $attachDeadline) {
        $workbookWindow = [Xlide.Attach]::WorkbookWindowOf($excelProcess.Id)
        if ($null -eq $workbookWindow) { Start-Sleep -Milliseconds 25 }
    }

    if ($null -eq $workbookWindow) {
        throw "Could not reach Excel $($excelProcess.Id) through its window."
    }

    $excel = $workbookWindow.Application
    $excel.DisplayAlerts = $false
    Complete-Phase 'attach'

    # The host loads editor add-ins when the editor starts, not when the host does, so this is the
    # step that actually triggers what is being tested.
    $excel.VBE.MainWindow.Visible = $true
    Complete-Phase 'open editor'

    # Wait for the add-in to say it finished starting. The log is matched on the host's process
    # identity, so a reused instance resolves to the log it wrote when it started.
    $logPattern = "shim-*-$($excelProcess.Id).log"
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $log = $null

    while ((Get-Date) -lt $deadline) {
        $log = Get-ChildItem $logRoot -Filter $logPattern -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if ($log -and (Select-String -Path $log.FullName -Pattern 'OnStartupComplete' -Quiet)) { break }
        if ($excelProcess.HasExited) { throw "Excel $($excelProcess.Id) exited before the add-in loaded." }

        Start-Sleep -Milliseconds 50
    }

    Complete-Phase 'add-in load'

    # The browser surface comes up asynchronously through the host message loop, so it settles
    # shortly after. Its absence is reported rather than waited out.
    $surface = $false
    if ($log) {
        $surfaceDeadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $surfaceDeadline) {
            if (Select-String -Path $log.FullName -Pattern 'webview: navigated' -Quiet) { $surface = $true; break }
            Start-Sleep -Milliseconds 50
        }
    }

    Complete-Phase 'browser surface'

    # Ask the editor itself whether it considers the add-in connected. The log says the add-in ran;
    # this says the host accepted it. Both are needed, because a class can be created and then
    # declined, and the host reports nothing when that happens.
    $connected = $null
    try {
        foreach ($addIn in $excel.VBE.Addins) {
            if ($addIn.ProgId -eq 'Xlide.VbeAddIn') { $connected = [bool] $addIn.Connect }
        }
        Write-Host "Editor reports connected = $connected"
    }
    catch {
        Write-Warning "Could not read the editor's add-in list: $($_.Exception.Message)"
    }

    Complete-Phase 'confirm connection'

    if ($log) {
        Write-Host ''
        Write-Host "--- $($log.Name) ---" -ForegroundColor Green
        Get-Content $log.FullName
        Write-Host '--- end of log ---' -ForegroundColor Green
    }

    Write-Host ''
    Write-Host 'Timings:'
    foreach ($phase in $phases) { Write-Host "  $phase" }
    Write-Host ('  {0,-24} {1,7:N2}s' -f 'total', ((Get-Date) - $startedAt).TotalSeconds)
    Write-Host ''

    # Each condition is checked against something observed. A library can load, and its class can be
    # created, while the add-in is never connected: that is what happens when the host cannot obtain
    # an interface it requires, and it reports nothing. Passing on "a log exists" would call that a
    # success, which is the failure this check exists to catch.
    if (-not $log) {
        Write-Warning "The add-in wrote nothing under $logRoot within $TimeoutSeconds seconds."
        Write-Warning 'The host did not activate the class. Check the registration and the server path.'
        exit 1
    }

    if (-not (Select-String -Path $log.FullName -Pattern 'OnConnection' -Quiet)) {
        Write-Warning 'The class was activated but the add-in was never connected.'
        Write-Warning 'The host obtained the object and then declined it, which it does silently.'
        exit 1
    }

    if ($connected -ne $true) {
        Write-Warning 'The add-in ran but the editor does not report it as connected.'
        exit 1
    }

    if (-not $surface) {
        Write-Warning 'The add-in connected but the browser surface never finished navigating.'
        exit 1
    }

    Write-Host 'RESULT: the add-in is connected, and its surface is running inside the editor.' -ForegroundColor Green
    exit 0
}
finally {
    if ($excel) {
        try { $excel.DisplayAlerts = $false } catch { }
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    if (-not $KeepOpen -and $excelProcess) {
        # Terminating the recorded identity is the only reliable way to close a host that was made
        # visible: it considers itself user-owned and can outlive every automation client.
        Stop-Process -Id $excelProcess.Id -Force -ErrorAction SilentlyContinue
        Remove-Item $ownedPidFile -Force -ErrorAction SilentlyContinue
    }
    elseif ($KeepOpen -and $excelProcess) {
        Write-Host "Leaving Excel $($excelProcess.Id) running. Re-run with -Reuse to check again quickly."
    }
}
