<#
.SYNOPSIS
    Starts Excel on a workbook and opens the editor, so the debug api has something to answer for.

.DESCRIPTION
    Three things here are not obvious, and each one cost a session to learn.

    STARTED AS AN ORDINARY PROCESS, with a document on the command line. A host created through
    automation (New-Object -ComObject Excel.Application) runs in EMBEDDING MODE and loads no
    add-ins at all, so the thing under test is never there.

    ATTACHED THROUGH ITS WINDOW, not the running object table. A host publishes itself in the ROT
    lazily — ten to forty seconds after it is visibly ready — and none of that wait says anything
    about the add-in. Asking a worksheet window for its native object model answers in under a
    second, and names the instance by window so there is no chance of adopting a different one.

    THE EDITOR IS WHAT LOADS THE ADD-IN, not the host. Excel loads VBE add-ins when the VBE
    starts, so nothing is under test until MainWindow.Visible is set.

.EXAMPLE
    tools\harness\Start-Excel.ps1
    Starts on the scratch workbook.

.EXAMPLE
    tools\harness\Start-Excel.ps1 -Workbook artifacts\fixtures\RenameFixture.xlsm
    Starts on the rename fixture. NOTE: that fixture deliberately does not compile.
#>
[CmdletBinding()]
param(
    # The workbook to open. Relative paths are taken from the repository root.
    [string] $Workbook,

    # Close any Excel already running first. A publish needs this anyway, because a host holds an
    # add-in library open for its lifetime.
    [switch] $Fresh,

    # Seconds to wait for the host's window to appear.
    [int] $TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Add-Type -Namespace XlideHarness -Name Attach -MemberDefinition @'
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

if (-not $Workbook) {
    $Workbook = Join-Path $PSScriptRoot 'fixtures\scratch.xlsm'
    if (-not (Test-Path $Workbook)) { & (Join-Path $PSScriptRoot 'New-ScratchWorkbook.ps1') | Out-Null }
}

if (-not [System.IO.Path]::IsPathRooted($Workbook)) {
    $Workbook = Join-Path $repoRoot $Workbook
}

if (-not (Test-Path $Workbook)) { throw "No workbook at $Workbook." }

if ($Fresh) {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# A harness terminates Excel by design, and Excel reads termination as a crash: on the next start
# it offers document recovery and disables what it blames, both of which stand in front of the
# thing being tested.
foreach ($version in @('16.0', '15.0')) {
    $key = "HKCU:\Software\Microsoft\Office\$version\Excel\Resiliency"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

$process = Start-Process -FilePath (Find-ExcelExecutable) -ArgumentList $Workbook -PassThru
Write-Host "Started Excel as process $($process.Id) on $(Split-Path -Leaf $Workbook)."

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$window = $null
while ($null -eq $window -and (Get-Date) -lt $deadline) {
    $window = [XlideHarness.Attach]::WorkbookWindowOf($process.Id)
    if ($null -eq $window) { Start-Sleep -Milliseconds 50 }
}
if ($null -eq $window) { throw "Could not reach Excel $($process.Id) through its window." }

$excel = $window.Application
$excel.DisplayAlerts = $false

# The editor is opened through Excel's OWN ribbon command, not through $excel.VBE.
#
# That matters more than it looks: `Application.VBE` and `Workbook.VBProject` are exactly what
# "Trust access to the VBA project object model" refuses, and with it off they come back NULL —
# not an exception, a null, so a try/catch reports success and prints nothing. ExecuteMso is
# Excel executing its own Developer > Visual Basic button, and is not gated. So this script, and
# everything the debug api does after it, works with that setting OFF (verified 2026-08-07).
$excel.CommandBars.ExecuteMso('VisualBasic')
Write-Host 'Editor opened (through the ribbon command, which needs no VBA project trust).'

# What is in the project is asked of the DOOR, for the same reason.
$listed = $false
$deadline = (Get-Date).AddSeconds(30)
while (-not $listed -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    try {
        $answer = & node (Join-Path $PSScriptRoot 'xlide-api.mjs') doctor 2>$null | Out-String
        if ($answer -match '"healthy"') {
            $listed = $true
            if ($answer -match '"healthy":\s*true') {
                Write-Host 'The add-in is up and its door is healthy.' -ForegroundColor Green
            }
            else {
                Write-Host 'The add-in is up but the doctor has findings:' -ForegroundColor Yellow
                Write-Host $answer
            }
        }
    }
    catch { }
}

if (-not $listed) {
    Write-Host 'The door did not answer. Debug build? Registered? Try: node tools\harness\xlide-api.mjs doctor' -ForegroundColor Yellow
}

Write-Host ''
Write-Host "pid=$($process.Id)"
