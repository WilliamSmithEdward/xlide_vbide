<#
.SYNOPSIS
    Starts an owned Excel instance, opens the Visual Basic Editor, and reports whether the add-in
    loaded.

.DESCRIPTION
    This is the smallest useful integration check: it proves the editor found the registration,
    activated the class, and called into the shim.

    Two details of how it starts Excel are load bearing.

    First, Excel is started as a real child process rather than through automation. An Excel created
    by automation runs in embedding mode, and in that mode the host does not load add-ins at all. A
    harness that creates Excel the convenient way reports a false failure for every add-in ever
    written.

    Second, having launched it, the script attaches to that instance and verifies by process
    identity that it attached to the one it started. Automation hands back whatever Excel is already
    running, so without that check a stray attach could drive, and later terminate, the developer's
    own session. The script refuses to run when Excel is already open unless told otherwise, and it
    terminates only the process identity it recorded.
#>
[CmdletBinding()]
param(
    # Upper bound on how long to wait for the add-in to finish loading. This is a deadline, not a
    # delay: the check proceeds the moment the add-in reports it is done, which is normally well
    # under a second.
    [int] $TimeoutSeconds = 20,

    # Leave Excel running afterwards, so the next run can reuse it.
    [switch] $KeepOpen,

    # Reuse the instance a previous run left open instead of starting a new one.
    #
    # Starting Excel costs about thirty seconds on a typical machine, against about a quarter of a
    # second for everything this add-in does, so reuse is the difference between a loop that is
    # worth running and one that is not. It only reuses an instance this harness started and
    # recorded, never one the developer opened.
    #
    # A reused instance still has the previous build of the shim loaded, because a host holds an
    # add-in library open for its lifetime. Use this after changing anything except the shim.
    [switch] $Reuse,

    # Proceed even when the developer already has Excel open.
    [switch] $AllowExistingExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$logRoot = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'
$startedAt = Get-Date

# Phase timings. A harness that is slow for an unknown reason gets worked around rather than fixed,
# so it reports where its time goes.
$phaseClock = [System.Diagnostics.Stopwatch]::StartNew()
$phases = New-Object System.Collections.Generic.List[string]

function Complete-Phase {
    param([string] $Name)

    $phases.Add(('{0,-28} {1,7:N2}s' -f $Name, $phaseClock.Elapsed.TotalSeconds))
    $phaseClock.Restart()
}

function Get-ExcelProcessIds {
    $found = New-Object System.Collections.Generic.List[int]
    foreach ($process in (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue)) {
        $found.Add([int] $process.Id)
    }

    return , $found.ToArray()
}

function Find-ExcelExecutable {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) { return $candidate }
    }

    $fromRegistry = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)'
    if ($fromRegistry -and (Test-Path $fromRegistry)) { return $fromRegistry }

    throw 'Could not locate EXCEL.EXE.'
}

function Reset-ExcelResiliency {
    <#
        A harness terminates Excel by design, and Excel treats termination as a crash. On the next
        start it offers document recovery and disables items it blames for the crash. Both are modal
        or pane-level states that appear before the instance is drivable, so a suite that does not
        clear them fails on its second run for reasons unrelated to what it is testing.
    #>
    foreach ($version in @('16.0', '15.0')) {
        $resiliency = "HKCU:\Software\Microsoft\Office\$version\Excel\Resiliency"
        if (-not (Test-Path $resiliency)) { continue }

        foreach ($child in @('DocumentRecovery', 'DisabledItems', 'StartupItems', 'CrashingAddinList')) {
            $path = Join-Path $resiliency $child
            if (Test-Path $path) {
                Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
                Write-Verbose "cleared $path"
            }
        }
    }
}

function Reset-AddInLoadBehavior {
    <#
        When an add-in fails to connect, the editor rewrites its LoadBehavior to 0 so it will not
        try again. That is reasonable for a user and useless for a developer: every run after the
        first failure silently skips the add-in, and the log stays empty for a reason that has
        nothing to do with the change being tested. Restoring it before each run makes failures
        reproducible.
    #>
    $key = 'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Xlide.VbeAddIn'
    if (-not (Test-Path $key)) {
        Write-Warning "The add-in is not registered. Run tools\dev.ps1 -Register first."
        return
    }

    $current = (Get-ItemProperty $key -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior
    if ($current -ne 3) {
        Write-Host "LoadBehavior was $current (the editor disabled the add-in after a failed connect). Restoring 3."
        Set-ItemProperty $key -Name LoadBehavior -Value 3 -Type DWord
    }
}

$ownedPidFile = Join-Path $env:LOCALAPPDATA 'xlide_vbide\harness-excel.pid'

function Get-RecordedInstanceId {
    if (-not (Test-Path $ownedPidFile)) { return 0 }

    $recorded = 0
    if ([int]::TryParse((Get-Content $ownedPidFile -Raw).Trim(), [ref] $recorded)) { return $recorded }
    return 0
}

$preExisting = Get-ExcelProcessIds
if ($preExisting.Length -gt 0 -and -not $AllowExistingExcel) {
    # An instance this harness started and recorded is ours to drive again. Anything else belongs
    # to the developer and is left alone.
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

Add-Type -Namespace Xlide -Name Win -MemberDefinition @'
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int processId);
'@

$excelPath = Find-ExcelExecutable
Write-Host "Excel: $excelPath"

$excel = $null
$excelProcess = $null
$reused = $false

try {
    if ($Reuse -and (Test-Path $ownedPidFile)) {
        $recorded = 0
        if ([int]::TryParse((Get-Content $ownedPidFile -Raw).Trim(), [ref] $recorded)) {
            $candidate = Get-Process -Id $recorded -ErrorAction SilentlyContinue
            if ($candidate -and $candidate.ProcessName -eq 'EXCEL') {
                $excelProcess = $candidate
                $reused = $true
                Write-Host "Reusing the instance from an earlier run (process $recorded)."
            }
        }
    }

    if (-not $reused) {
        # Started with a document, and as an ordinary process so add-ins load normally. Excel with
        # no document takes appreciably longer to publish itself for automation, which is otherwise
        # the largest single cost in this check.
        $scratch = Join-Path $PSScriptRoot 'fixtures\scratch.xlsx'
        if (-not (Test-Path $scratch)) {
            & (Join-Path $PSScriptRoot 'New-ScratchWorkbook.ps1') | Out-Null
        }

        $excelProcess = Start-Process -FilePath $excelPath -ArgumentList $scratch -PassThru
        Write-Host "Started Excel as process $($excelProcess.Id)."

        New-Item -ItemType Directory -Force -Path (Split-Path $ownedPidFile) | Out-Null
        Set-Content -Path $ownedPidFile -Value $excelProcess.Id
    }

    # Wait for the instance to publish itself so it can be driven. Poll quickly: a warm start is
    # ready in a few hundred milliseconds and there is no reason to pay a fixed price for it.
    $deadline = (Get-Date).AddSeconds(60)
    while ($null -eq $excel -and (Get-Date) -lt $deadline) {
        try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
        if ($null -eq $excel) { Start-Sleep -Milliseconds 100 }
    }

    if ($null -eq $excel) {
        throw "Excel $($excelProcess.Id) never became available to automation."
    }

    [int] $attachedPid = 0
    [void] [Xlide.Win]::GetWindowThreadProcessId([System.IntPtr]$excel.Hwnd, [ref] $attachedPid)

    if ($attachedPid -ne $excelProcess.Id) {
        throw "Attached to Excel $attachedPid, which is not the instance this script started ($($excelProcess.Id)). Refusing to drive it."
    }

    Write-Host "Attached to the owned instance ($attachedPid)."
    Complete-Phase 'host start and attach'

    $excel.DisplayAlerts = $false
    if ($excel.Workbooks.Count -eq 0) { [void] $excel.Workbooks.Add() }
    Complete-Phase 'workbook'

    Write-Host 'Opening the Visual Basic Editor...'
    $excel.VBE.MainWindow.Visible = $true
    Complete-Phase 'open editor'

    # Wait for the thing itself rather than for a duration. The add-in writes a line when it has
    # finished its startup work, so that line is the signal. A fixed sleep would be wrong in both
    # directions: too long on every healthy run, and too short whenever the machine is loaded.
    $loadDeadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $log = $null
    while ((Get-Date) -lt $loadDeadline) {
        # Matched on the host's process identity rather than on time. The log belonging to this
        # instance is the right one whether it was written a moment ago or by an earlier run that
        # left the instance open.
        $log = Get-ChildItem $logRoot -Filter "shim-*-$($excelProcess.Id).log" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($log -and (Select-String -Path $log.FullName -Pattern 'OnStartupComplete' -Quiet)) {
            break
        }

        Start-Sleep -Milliseconds 100
    }

    # The browser surface comes up asynchronously through the host message loop, so it settles a
    # little after the add-in reports startup. Give it a short grace period, and stop as soon as it
    # arrives. Its absence is reported rather than waited out.
    if ($log) {
        $surfaceDeadline = (Get-Date).AddSeconds(10)
        while ((Get-Date) -lt $surfaceDeadline) {
            if (Select-String -Path $log.FullName -Pattern 'webview: navigated' -Quiet) { break }
            Start-Sleep -Milliseconds 100
        }
    }

    Complete-Phase 'add-in load'

    Write-Host ''
    Write-Host 'Timings:'
    foreach ($phase in $phases) { Write-Host "  $phase" }
    Write-Host ('  {0,-28} {1,7:N2}s' -f 'total', ((Get-Date) - $startedAt).TotalSeconds)

    # Whether the editor connected the add-in, taken from the editor rather than inferred.
    $connected = $false
    $addInReport = 'the editor reported no add-ins'
    try {
        $descriptions = New-Object System.Collections.Generic.List[string]
        foreach ($addIn in $excel.VBE.Addins) {
            $descriptions.Add("$($addIn.ProgId) connect=$($addIn.Connect)")
            if ($addIn.ProgId -eq 'Xlide.VbeAddIn' -and $addIn.Connect) {
                $connected = $true
            }
        }

        if ($descriptions.Count -gt 0) {
            $addInReport = $descriptions -join '; '
        }
    }
    catch {
        $addInReport = "could not enumerate add-ins: $($_.Exception.Message)"
    }

    Write-Host "Editor add-in list: $addInReport"

    # $log was already resolved against this instance's process identity while waiting for it.
    # Looking it up a second time by timestamp would miss a reused instance, whose add-in loaded
    # before this run started.
    if ($log) {
        Write-Host ''
        Write-Host "--- $($log.Name) ---" -ForegroundColor Green
        Get-Content $log.FullName
        Write-Host '--- end of log ---' -ForegroundColor Green
    }

    Write-Host ''

    # Each condition is checked against something observed, not against the absence of an error.
    # A library can load, and its class can be created, while the add-in is never connected: that
    # is what happens when the host cannot obtain an interface it requires, and it reports nothing.
    # Passing on "a log exists" would call that state a success.
    $ran = $log -and (Select-String -Path $log.FullName -Pattern 'OnConnection' -Quiet)

    if (-not $log) {
        Write-Warning "No log was written under $logRoot after $($startedAt.ToString('HH:mm:ss'))."
        Write-Warning 'The editor did not activate the class. Check the registration and the server path.'
        exit 1
    }

    if (-not $ran) {
        Write-Warning 'The class was activated but the add-in was never connected.'
        Write-Warning 'The host obtained the object and then declined it, which it does silently.'
        exit 1
    }

    if (-not $connected) {
        Write-Warning 'The add-in ran but the editor does not report it as connected.'
        Write-Warning 'Treat this as a failure: the connection did not survive startup.'
        exit 1
    }

    Write-Host 'RESULT: the add-in is connected and running inside the editor.' -ForegroundColor Green
    exit 0
}
finally {
    if (-not $KeepOpen) {
        if ($excel) {
            try { $excel.DisplayAlerts = $false } catch { }
            try { $excel.Quit() } catch { }
            [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        # Quit is a request, not a guarantee. Watch for the exit and terminate the recorded identity
        # the moment it is clear it is not going to happen, rather than waiting a fixed interval.
        if ($excelProcess) {
            $exitDeadline = (Get-Date).AddSeconds(3)
            while ((Get-Date) -lt $exitDeadline) {
                if (-not (Get-Process -Id $excelProcess.Id -ErrorAction SilentlyContinue)) { break }
                Start-Sleep -Milliseconds 100
            }

            if (Get-Process -Id $excelProcess.Id -ErrorAction SilentlyContinue) {
                Write-Host "Excel $($excelProcess.Id) did not exit on request. Terminating the owned instance."
                Stop-Process -Id $excelProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
    elseif ($excelProcess) {
        Write-Host "Leaving Excel $($excelProcess.Id) running as requested."
    }
}
