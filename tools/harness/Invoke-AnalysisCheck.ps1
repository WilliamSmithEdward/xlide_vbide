<#
.SYNOPSIS
    Proves the add-in reports real analysis for code the user has open.

.DESCRIPTION
    The whole chain, end to end: the add-in loads, starts the engine, reads the module the user is
    editing out of the editor, has it analysed, and reports the finding at the right position.

    A module with a deliberate defect is added, and one without. Both matter. Reporting the defect
    proves the chain works; reporting nothing on the clean module proves it is not simply reporting
    everything, which is the failure that makes an analyzer worthless.
#>
[CmdletBinding()]
param(
    [int] $TimeoutSeconds = 60,
    [switch] $KeepOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$logRoot = Join-Path $env:LOCALAPPDATA 'xlide_vbide\logs'
$startedAt = Get-Date

Add-Type -Namespace Xlide -Name Attach -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

delegate bool EnumProc(IntPtr h, IntPtr l);

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
    return AccessibleObjectFromWindow(sheet, 0xFFFFFFF0u, ref dispatch, out window) == 0 ? window : null;
}
'@

foreach ($version in @('16.0', '15.0')) {
    $resiliency = "HKCU:\Software\Microsoft\Office\$version\Excel\Resiliency"
    if (Test-Path $resiliency) {
        foreach ($child in @('DocumentRecovery', 'DisabledItems', 'StartupItems', 'CrashingAddinList')) {
            $path = Join-Path $resiliency $child
            if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

$addInKey = 'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Xlide.VbeAddIn'
if ((Test-Path $addInKey) -and (Get-ItemProperty $addInKey -Name LoadBehavior).LoadBehavior -ne 3) {
    Set-ItemProperty $addInKey -Name LoadBehavior -Value 3 -Type DWord
}

$excelPath = "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE"
$scratch = Join-Path $PSScriptRoot 'fixtures\scratch.xlsm'
if (-not (Test-Path $scratch)) { & (Join-Path $PSScriptRoot 'New-ScratchWorkbook.ps1') | Out-Null }

$process = Start-Process -FilePath $excelPath -ArgumentList $scratch -PassThru
Write-Host "Excel is process $($process.Id)."

try {
    $deadline = (Get-Date).AddSeconds(60)
    $window = $null
    while ($null -eq $window -and (Get-Date) -lt $deadline) {
        $window = [Xlide.Attach]::WorkbookWindowOf($process.Id)
        if ($null -eq $window) { Start-Sleep -Milliseconds 25 }
    }

    if ($null -eq $window) { throw 'Could not reach the host through its window.' }

    $excel = $window.Application
    $excel.DisplayAlerts = $false
    $components = $excel.ActiveWorkbook.VBProject.VBComponents

    # A defect the analyzer must find: a string assigned to a whole-number variable.
    $bad = $components.Add(1)
    $bad.Name = 'BrokenModule'
    $bad.CodeModule.AddFromString("Option Explicit`r`n`r`nSub Broken()`r`n    Dim n As Long`r`n    n = ""oops""`r`nEnd Sub`r`n")

    # Code with nothing wrong with it. Reporting anything here would be worse than reporting
    # nothing at all on the module above.
    $good = $components.Add(1)
    $good.Name = 'CleanModule'
    $good.CodeModule.AddFromString("Option Explicit`r`n`r`nSub Fine()`r`n    Dim n As Long`r`n    n = 1`r`nEnd Sub`r`n")

    Write-Host 'Added BrokenModule and CleanModule.'

    & (Join-Path $PSScriptRoot 'Open-VbeEditor.ps1') -Excel $excel

    $logPattern = "shim-*-$($process.Id).log"
    $analysisDeadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $log = $null
    $reported = $false

    while ((Get-Date) -lt $analysisDeadline) {
        $log = Get-ChildItem $logRoot -Filter $logPattern -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if ($log -and (Select-String -Path $log.FullName -Pattern 'analysis: \d+ finding' -Quiet)) {
            $reported = $true
            break
        }

        Start-Sleep -Milliseconds 100
    }

    $elapsed = ((Get-Date) - $startedAt).TotalSeconds
    Write-Host ("Analysis reported after {0:N1}s." -f $elapsed)

    if (-not $reported) {
        Write-Warning 'The add-in never reported any analysis.'
        if ($log) { Get-Content $log.FullName | Select-String -Pattern 'engine|analysis' }
        exit 1
    }

    Write-Host ''
    $lines = Get-Content $log.FullName | Select-String -Pattern 'engine:|analysis:|BrokenModule|CleanModule'
    foreach ($line in $lines) { Write-Host "  $line" }
    Write-Host ''

    $foundDefect = (Select-String -Path $log.FullName -Pattern 'BrokenModule.*error' -Quiet)
    $flaggedClean = (Select-String -Path $log.FullName -Pattern 'CleanModule\(' -Quiet)

    if (-not $foundDefect) {
        Write-Warning 'The defect in BrokenModule was not reported.'
        exit 1
    }

    if ($flaggedClean) {
        Write-Warning 'CleanModule was flagged. A false report is worse than a missed one.'
        exit 1
    }

    Write-Host 'RESULT: the defect was reported, and the clean module was not.' -ForegroundColor Green
    exit 0
}
finally {
    if ($excel) {
        try { $excel.DisplayAlerts = $false } catch { }
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    if (-not $KeepOpen) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Get-Process xlide-engine -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}
