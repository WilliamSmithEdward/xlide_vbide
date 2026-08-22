<#
.SYNOPSIS
    Starts Word on a document and opens the editor, so the xlide api has something to answer for.

.DESCRIPTION
    Start-Excel.ps1's twin, built the day Word testing became routine (2026-08-19); the three
    hard-won rules there hold unchanged here:

    STARTED AS AN ORDINARY PROCESS, with a document on the command line. A host created through
    automation (New-Object -ComObject Word.Application) runs in EMBEDDING MODE and loads no
    add-ins at all, so the thing under test is never there.

    ATTACHED THROUGH ITS WINDOW, not the running object table. A host publishes itself in the ROT
    lazily - ten to forty seconds after it is visibly ready. Asking the document pane (class
    `_WwG`) for its native object model answers in under a second and names the instance by
    window, so there is no chance of adopting a different one.

    THE EDITOR IS WHAT LOADS THE ADD-IN, not the host. ExecuteMso('VisualBasic') is Word running
    its own Developer > Visual Basic button - the same idMso as Excel's - and is not gated by
    "Trust access to the VBA project object model".

.EXAMPLE
    tools\harness\Start-Word.ps1
    Starts on the scratch document (created on first use).

.EXAMPLE
    tools\harness\Start-Word.ps1 -Document artifacts\fixtures\Something.docm -Fresh
#>
[CmdletBinding()]
param(
    # The document or documents to open. Relative paths are taken from the repository root.
    [string[]] $Document,

    # Close any Word already running first. Word only - an Excel session beside it is somebody
    # else's work and stays.
    [switch] $Fresh,

    # Seconds to wait for the host's window to appear.
    [int] $TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Add-Type -Namespace XlideHarness -Name AttachWord -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

delegate bool EnumProc(IntPtr h, IntPtr l);

// OBJID_NATIVEOM. Asking Word's document pane for its native object model yields the document
// Window, and its Application, without going through the running object table.
const uint NativeObjectModel = 0xFFFFFFF0u;

public static object DocumentWindowOf(int processId)
{
    IntPtr pane = IntPtr.Zero;

    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }

        EnumChildWindows(h, (child, l2) =>
        {
            var name = new System.Text.StringBuilder(128);
            GetClassNameW(child, name, 128);
            if (name.ToString() == "_WwG") { pane = child; return false; }
            return true;
        }, IntPtr.Zero);

        return pane == IntPtr.Zero;
    }, IntPtr.Zero);

    if (pane == IntPtr.Zero) { return null; }

    var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
    object window;
    return AccessibleObjectFromWindow(pane, NativeObjectModel, ref dispatch, out window) == 0 ? window : null;
}
'@

function Find-WordExecutable {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\WINWORD.EXE"
    )
    foreach ($candidate in $candidates) { if (Test-Path $candidate) { return $candidate } }

    $fromRegistry = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe' -ErrorAction SilentlyContinue).'(default)'
    if ($fromRegistry -and (Test-Path $fromRegistry)) { return $fromRegistry }

    throw 'Could not locate WINWORD.EXE.'
}

if (-not $Document -or $Document.Count -eq 0) {
    $scratch = Join-Path $PSScriptRoot 'fixtures\scratch.docm'
    if (-not (Test-Path $scratch)) {
        # Made through automation, which is FINE here: the maker never needs the add-in, only a
        # macro-enabled file on disk for the real launch below to open as an ordinary process.
        Write-Host 'Making the scratch document.'
        $maker = New-Object -ComObject Word.Application
        $maker.DisplayAlerts = 0
        try {
            $blank = $maker.Documents.Add()
            $blank.SaveAs2($scratch, 13)  # wdFormatXMLDocumentMacroEnabled
            $blank.Close($false)
        }
        finally {
            try { $maker.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($maker) | Out-Null
        }
    }
    $Document = @($scratch)
}

$Document = @($Document | ForEach-Object {
    $one = $_
    if (-not [System.IO.Path]::IsPathRooted($one)) { $one = Join-Path $repoRoot $one }
    if (-not (Test-Path $one)) { throw "No document at $one." }
    $one
})

if ($Fresh) {
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}

# A harness terminates Word by design, and Word reads termination as a crash: on the next start
# it offers document recovery and disables what it blames, both of which stand in front of the
# thing being tested.
foreach ($version in @('16.0', '15.0')) {
    $key = "HKCU:\Software\Microsoft\Office\$version\Word\Resiliency"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

# Quoted individually: a fixture path with a space in it becomes two arguments otherwise.
$arguments = @($Document | ForEach-Object { '"{0}"' -f $_ })
$process = Start-Process -FilePath (Find-WordExecutable) -ArgumentList $arguments -PassThru
$names = ($Document | ForEach-Object { Split-Path -Leaf $_ }) -join ', '
Write-Host "Started Word as process $($process.Id) on $names."

$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$window = $null
while ($null -eq $window -and (Get-Date) -lt $deadline) {
    $window = [XlideHarness.AttachWord]::DocumentWindowOf($process.Id)
    if ($null -eq $window) { Start-Sleep -Milliseconds 50 }
}
if ($null -eq $window) { throw "Could not reach Word $($process.Id) through its window." }

$word = $window.Application
$word.DisplayAlerts = 0

# The editor is opened through Word's OWN ribbon command, not through $word.VBE, for exactly
# Start-Excel.ps1's reason: Application.VBE is what the trust setting refuses, and ExecuteMso
# is the host pressing its own button.
$word.CommandBars.ExecuteMso('VisualBasic')
Write-Host 'Editor opened (through the ribbon command, which needs no VBA project trust).'

# The doctor, BY PID: another live session - an Excel fixture beside this Word - is a designed
# state, and the bare verb refuses to guess between instances.
$listed = $false
$deadline = (Get-Date).AddSeconds(30)
while (-not $listed -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    try {
        $answer = & node (Join-Path $PSScriptRoot 'xlide-api.mjs') --pid $process.Id doctor 2>$null | Out-String
        if ($answer -match '"healthy"') {
            if ($answer -match '"healthy":\s*true') {
                $listed = $true
                Write-Host 'The add-in is up and its door is healthy.' -ForegroundColor Green
            }
            elseif ((Get-Date) -ge $deadline.AddSeconds(-3)) {
                $listed = $true
                Write-Host 'The add-in is up but the doctor has findings:' -ForegroundColor Yellow
                Write-Host $answer
            }
        }
    }
    catch { }
}

if (-not $listed) {
    Write-Host 'The door did not answer. Debug build? Registered? Try: node tools\harness\xlide-api.mjs instances' -ForegroundColor Yellow
}

Write-Host ''
Write-Host "pid=$($process.Id)"
