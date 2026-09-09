<#
.SYNOPSIS
    Starts Excel on a workbook and opens the editor, so the xlide api has something to answer for.

.DESCRIPTION
    Three things here are not obvious, and each one cost a session to learn.

    STARTED AS AN ORDINARY PROCESS, with a document on the command line. A host created through
    automation (New-Object -ComObject Excel.Application) runs in EMBEDDING MODE and loads no
    add-ins at all, so the thing under test is never there.

    ATTACHED THROUGH ITS WINDOW, not the running object table. A host publishes itself in the ROT
    lazily - ten to forty seconds after it is visibly ready - and none of that wait says anything
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
    # The workbook or workbooks to open. Relative paths are taken from the repository root.
    #
    # SEVERAL is not a convenience. Two workbooks holding a module of the same name is the state
    # three separate defects have lived in -- navigation, tab labels, breakpoints -- and there
    # was no way to set it up from the harness at all, so every one of them was found by hand
    # (2026-08-07). Excel takes them on one command line and puts them in ONE process, which is
    # what makes them one session and one door.
    [string[]] $Workbook,

    # Close the Excels this harness started first - the ones holding a fixture or chaos
    # workbook - each by its own id. A publish needs this anyway, because a host holds an add-in
    # library open for its lifetime. Every other Excel stays, whoever started it: the developer's
    # own workbooks, and another automation's hidden instances, which a sweep by name was ending
    # mid-statement (#24). When one stays, the fixture starts in a process of its own.
    [switch] $Fresh,

    # Close EVERY Excel the census saw, strangers included, each by id. For a machine that is
    # genuinely yours and on which nobody else's automation is running.
    [switch] $Force,

    # Start a SEPARATE Excel process rather than letting Excel reuse the one already running.
    #
    # Excel puts every workbook in one process by default, which is what makes several fixtures
    # one session and one door - usually the point. The opposite state is its own thing: two
    # PROCESSES means two add-in loads, two doors, and two registrations of the one inside-door
    # name, which is where the fleet behaviour lives (#11) and which nothing here could set up.
    # `/x` is Excel's own switch for it.
    [switch] $Separate,

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

if (-not $Workbook -or $Workbook.Count -eq 0) {
    $scratch = Join-Path $PSScriptRoot 'fixtures\scratch.xlsm'
    if (-not (Test-Path $scratch)) { & (Join-Path $PSScriptRoot 'New-ScratchWorkbook.ps1') | Out-Null }
    $Workbook = @($scratch)
}

$Workbook = @($Workbook | ForEach-Object {
    $one = $_
    if (-not [System.IO.Path]::IsPathRooted($one)) { $one = Join-Path $repoRoot $one }
    if (-not (Test-Path $one)) { throw "No workbook at $one." }
    $one
})

if ($Fresh) {
    <#
        -Fresh CLOSES ONLY WHAT THIS HARNESS STARTED. For a long time it closed every Excel on
        the machine, behind a census that refused when a workbook it had not put there was
        open. The refusal kept the developer's own workbooks safe (2026-08-29, one of theirs
        survived a sweep only because they had just reopened it) and nothing else: an Excel
        another automation drives over COM holds no workbook of ours and no window this census
        can read, and a stop by name ended it in the middle of whatever statement was running.
        vbaSQLBridge measured six runs in sixty lost while a fixture Excel was up, none in forty
        after it had gone, with no crash record anywhere, because nothing crashed (#24).

        So the census decides what is OURS - a process holding a workbook under the fixture or
        chaos folders - and only those are stopped, each by its id. The answer comes from the
        WORKBOOKS each process holds rather than its title bar, because a title names only the
        active one. Everything else is left standing and named, and the fixture then starts in
        a process of its own, because Excel would otherwise open it inside whichever instance is
        already up. -Force stops every process the census saw, strangers included.
    #>
    $mine = @(
        (Join-Path $repoRoot 'artifacts\fixtures'),
        (Join-Path $repoRoot 'artifacts\chaos'),
        (Join-Path $PSScriptRoot 'fixtures')
    ) | ForEach-Object {
        if (Test-Path $_) { Get-ChildItem $_ -File | ForEach-Object { $_.FullName } }
    }

    $ours = @()
    $theirs = @()
    foreach ($running in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        $held = $null
        # EVERY WRAPPER THIS READ TAKES IS GIVEN BACK BEFORE THE SWEEP. A wrapper the collector
        # finalises after its Excel has been killed makes DCOM start a fresh hidden Excel to
        # answer the release (2026-09-08, a "Book1" after every kill). The window, its
        # Application, the Workbooks collection and each workbook are released here, in reverse,
        # while the process is still alive.
        $itsWindow = $null
        $itsApp = $null
        $itsBooks = $null
        try {
            $itsWindow = [XlideHarness.Attach]::WorkbookWindowOf($running.Id)
            if ($null -ne $itsWindow) {
                $itsApp = $itsWindow.Application
                $itsBooks = $itsApp.Workbooks
                $held = @()
                foreach ($book in $itsBooks) {
                    $held += $book.FullName
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
                }
            }
        }
        catch { $held = $null }
        finally {
            foreach ($wrapper in @($itsBooks, $itsApp, $itsWindow)) {
                if ($null -ne $wrapper) {
                    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrapper) | Out-Null } catch { }
                }
            }
        }

        if ($null -eq $held) {
            # A hidden automation instance, or one mid-teardown: no window to ask, and not ours
            # to close on a guess. Unreadable is not empty.
            $theirs += "pid $($running.Id) (no workbook window this harness can read)"
        }
        elseif (@($held | Where-Object { $mine -contains $_ }).Count -gt 0) {
            $ours += $running.Id
        }
        else {
            $what = if ($held.Count -gt 0) { $held -join ', ' } else { 'no workbook open' }
            $theirs += "pid $($running.Id) ($what)"
        }
    }

    # The collector runs NOW, with every Excel still alive: a wrapper missed above is released
    # while its server can still answer, rather than after the stop below.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $stopping = @(if ($Force) { Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } } else { $ours })
    foreach ($id in $stopping) {
        Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
    }
    foreach ($id in $stopping) {
        $stopped = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -ne $stopped) { $stopped.WaitForExit(10000) | Out-Null }
    }
    if ($stopping.Count -gt 0) {
        $whose = if ($Force) { 'every Excel' } else { "this harness's Excel" }
        Write-Host "Closed $whose by id: $($stopping -join ', ')."
    }

    if (-not $Force -and $theirs.Count -gt 0) {
        Write-Host ("Left standing, not this harness's to close:" + [Environment]::NewLine + '  ' +
            ($theirs -join ([Environment]::NewLine + '  ')))
        # SEPARATE, or the fixture opens inside theirs: Excel hands a workbook on its command line
        # to an instance already running, and the door would then be in a process this harness
        # never started.
        $Separate = $true
    }
}

# A harness terminates Excel by design, and Excel reads termination as a crash: on the next start
# it offers document recovery and disables what it blames, both of which stand in front of the
# thing being tested.
foreach ($version in @('16.0', '15.0')) {
    $key = "HKCU:\Software\Microsoft\Office\$version\Excel\Resiliency"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

# Quoted individually: a fixture path with a space in it becomes two arguments otherwise, and
# Excel then opens neither and offers to create them.
$arguments = @($Workbook | ForEach-Object { '"{0}"' -f $_ })
if ($Separate) { $arguments = @('/x') + $arguments }
$process = Start-Process -FilePath (Find-ExcelExecutable) -ArgumentList $arguments -PassThru
$names = ($Workbook | ForEach-Object { Split-Path -Leaf $_ }) -join ', '
$apart = if ($Separate) { ', in a process of its own' } else { '' }
Write-Host "Started Excel as process $($process.Id) on $names$apart."

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
# "Trust access to the VBA project object model" refuses, and with it off they come back NULL -
# not an exception, a null, so a try/catch reports success and prints nothing. ExecuteMso is
# Excel executing its own Developer > Visual Basic button, and is not gated. So this script, and
# everything the xlide api does after it, works with that setting OFF (verified 2026-08-07).
$commandBars = $excel.CommandBars
$commandBars.ExecuteMso('VisualBasic')
Write-Host 'Editor opened (through the ribbon command, which needs no VBA project trust).'

# GIVEN BACK NOW, not left for the collector. A fixture generator dot-sources this script and
# stops the very Excel it attached to when it is done; a wrapper finalised after that kill makes
# DCOM start a fresh hidden Excel to answer the release (2026-09-08). Everything after this line
# talks to the door over HTTP and needs none of these.
foreach ($wrapper in @($commandBars, $excel, $window)) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrapper) | Out-Null } catch { }
}
$commandBars = $null
$excel = $null
$window = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# What is in the project is asked of the DOOR, for the same reason.
$listed = $false
$deadline = (Get-Date).AddSeconds(30)
while (-not $listed -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    try {
        # BY PID: with a second live session beside this one - a Word fixture, another Excel -
        # the bare verb refuses to guess between instances, and this loop then read the refusal
        # as "not healthy yet" for its whole budget (caught 2026-08-19, building the Word twin).
        $answer = & node (Join-Path $PSScriptRoot 'xlide-api.mjs') --pid $process.Id doctor 2>$null | Out-String
        if ($answer -match '"healthy"') {
            # Healthy, or out of patience. The engine is started alongside the surface and takes
            # a beat longer to connect, so the FIRST answer here is routinely "up but the engine
            # is not answering yet" - a finding that resolves itself a second later. Reporting it
            # trains the reader to ignore the doctor, which is the one thing it must not do.
            #
            # Invisible until 2026-08-08, because engineUp was hardcoded to whether the service
            # object existed and so was true before the engine had started at all.
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
    Write-Host 'The door did not answer. Debug build? Registered? Try: node tools\harness\xlide-api.mjs doctor' -ForegroundColor Yellow
}

Write-Host ''
Write-Host "pid=$($process.Id)"
