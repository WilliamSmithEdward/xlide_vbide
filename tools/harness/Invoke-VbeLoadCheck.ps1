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
    # Seconds to leave the editor open before collecting results.
    [int] $DwellSeconds = 6,

    # Leave Excel running afterwards for manual inspection.
    [switch] $KeepOpen,

    # Proceed even when the developer already has Excel open.
    [switch] $AllowExistingExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$logRoot = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'
$startedAt = Get-Date

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

$preExisting = Get-ExcelProcessIds
if ($preExisting.Length -gt 0 -and -not $AllowExistingExcel) {
    Write-Error "Excel is already running (PID $($preExisting -join ', ')). Close it, or pass -AllowExistingExcel."
    exit 2
}

Reset-ExcelResiliency
Reset-AddInLoadBehavior

Add-Type -Namespace Xlide -Name Win -MemberDefinition @'
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int processId);
'@

$excelPath = Find-ExcelExecutable
Write-Host "Excel: $excelPath"

$excel = $null
$excelProcess = $null

try {
    # /e suppresses the start screen. This is an ordinary launch, so add-ins load normally.
    $excelProcess = Start-Process -FilePath $excelPath -ArgumentList '/e' -PassThru
    Write-Host "Started Excel as process $($excelProcess.Id)."

    # Wait for the instance to publish itself so it can be driven.
    $deadline = (Get-Date).AddSeconds(60)
    while ($null -eq $excel -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
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

    $excel.DisplayAlerts = $false
    [void] $excel.Workbooks.Add()

    Write-Host 'Opening the Visual Basic Editor...'
    $excel.VBE.MainWindow.Visible = $true

    Start-Sleep -Seconds $DwellSeconds

    $addInReport = 'the editor reported no add-ins'
    try {
        $descriptions = New-Object System.Collections.Generic.List[string]
        foreach ($addIn in $excel.VBE.Addins) {
            $descriptions.Add("$($addIn.ProgId) connect=$($addIn.Connect)")
        }

        if ($descriptions.Count -gt 0) {
            $addInReport = $descriptions -join '; '
        }
    }
    catch {
        $addInReport = "could not enumerate add-ins: $($_.Exception.Message)"
    }

    Write-Host "Editor add-in list: $addInReport"

    $log = Get-ChildItem $logRoot -Filter 'shim-*.log' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $startedAt } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($log) {
        Write-Host ''
        Write-Host "--- $($log.Name) ---" -ForegroundColor Green
        Get-Content $log.FullName
        Write-Host '--- end of log ---' -ForegroundColor Green
        Write-Host ''
        Write-Host 'RESULT: the add-in was loaded and ran inside the editor.' -ForegroundColor Green
        exit 0
    }

    Write-Host ''
    Write-Warning "No shim log was written under $logRoot after $($startedAt.ToString('HH:mm:ss'))."
    Write-Warning 'The editor did not activate the class. Check the registration and the server path.'
    exit 1
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

        # Quit is a request, not a guarantee. Terminate the recorded identity if it survives.
        if ($excelProcess) {
            Start-Sleep -Milliseconds 2000
            $survivor = Get-Process -Id $excelProcess.Id -ErrorAction SilentlyContinue
            if ($survivor) {
                Write-Host "Excel $($excelProcess.Id) did not exit on request. Terminating the owned instance."
                Stop-Process -Id $excelProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
    elseif ($excelProcess) {
        Write-Host "Leaving Excel $($excelProcess.Id) running as requested."
    }
}
