<#
.SYNOPSIS
    Opens the VBE in an Excel process that is ALREADY RUNNING, by process id.

.DESCRIPTION
    `Start-Excel.ps1` launches a host and opens its editor in one motion, which covers the usual
    case and only that case. There was no way to open the editor in an Excel somebody else
    started, or in a second one started alongside the first.

    That gap is why the two-instance bug could not be tested from here. A second Excel could be
    STARTED easily enough, but a VBE add-in loads when the VBE starts, not when Excel does, so an
    instance nobody opened the editor in never loaded xlide at all: the add-in looked absent when
    it had simply never been asked for (2026-08-08).

    Attaches the same way Start-Excel does, through the workbook window rather than the running
    object table, because a host publishes itself in the ROT ten to forty seconds after it is
    visibly ready and none of that wait says anything about the add-in.

.EXAMPLE
    tools\harness\Open-VbeIn.ps1 -ProcessId 28480

.EXAMPLE
    Get-Process EXCEL | ForEach-Object { tools\harness\Open-VbeIn.ps1 -ProcessId $_.Id }
    Every running Excel, which is how the several-instances case is set up.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [int] $ProcessId,

    [int] $TimeoutSeconds = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not ([System.Management.Automation.PSTypeName]'XlideHarness.Attach').Type) {
    Add-Type -Namespace XlideHarness -Name Attach -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc f, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

delegate bool EnumProc(IntPtr h, IntPtr l);

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

    var iid = new Guid("00020400-0000-0000-C000-000000000046");
    object window;
    if (AccessibleObjectFromWindow(sheet, NativeObjectModel, ref iid, out window) != 0) { return null; }
    return window;
}
'@
}

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$window = $null
while ($null -eq $window -and (Get-Date) -lt $deadline) {
    $window = [XlideHarness.Attach]::WorkbookWindowOf($ProcessId)
    if ($null -eq $window) { Start-Sleep -Milliseconds 100 }
}

if ($null -eq $window) {
    throw "Could not reach Excel $ProcessId through its window. Is it running and showing a workbook?"
}

$excel = $window.Application

# Through Excel's OWN ribbon command, never $excel.VBE. `Application.VBE` is one of the two
# properties "Trust access to the VBA project object model" gates, and with it off it comes back
# NULL rather than raising, so a try/catch reports success and opens nothing. ExecuteMso is not
# gated: it is Excel pressing its own Developer > Visual Basic button.
$excel.CommandBars.ExecuteMso('VisualBasic')

Write-Host "Opened the editor in Excel $ProcessId."

# The add-in loads with the VBE, so the door appears a moment after this returns rather than with
# it. Waited for here so a caller can drive immediately instead of sleeping and hoping.
$discovery = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$ProcessId.json"
$doorDeadline = (Get-Date).AddSeconds($TimeoutSeconds)
while (-not (Test-Path $discovery) -and (Get-Date) -lt $doorDeadline) {
    Start-Sleep -Milliseconds 200
}

if (Test-Path $discovery) {
    Write-Host "The add-in is up in $ProcessId; its door is at $discovery."
} else {
    Write-Host "The editor opened but no door appeared for $ProcessId. Debug build? Registered?" -ForegroundColor Yellow
}
