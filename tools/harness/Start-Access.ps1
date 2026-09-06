<#
.SYNOPSIS
    Starts Access on a database and opens the editor, so the xlide api has something to answer for.

.DESCRIPTION
    Start-Excel.ps1's third twin, written the day the test runner turned out never to have worked
    in Access (2026-09-06). The three rules the other two carry hold here unchanged - started as
    an ordinary process, attached through its window rather than the running object table, and the
    editor opened through the host's own ribbon command - and Access adds one of its own:

    ITS FRAME ANSWERS WITH THE APPLICATION ITSELF. Excel's worksheet pane and Word's document pane
    answer OBJID_NATIVEOM with a Window, whose `.Application` is the host. Access answers on its
    top-level `OMain` frame, and what comes back IS the Application - `.Name` on it reads
    "Microsoft Access". There is no document window to walk to either: Access's MDI children are
    forms and reports, so a database with none open would have no window to ask at all.

.EXAMPLE
    tools\harness\Start-Access.ps1 -Database artifacts\fixtures\AccessFixture.accdb -Fresh
#>
[CmdletBinding()]
param(
    # The database to open. A relative path is taken from the repository root.
    [Parameter(Mandatory = $true)]
    [string] $Database,

    # Close any Access already running first. Access only - an Excel session beside it is somebody
    # else's work and stays. It REFUSES when a running Access holds a database this harness did
    # not open; -Force sweeps anyway.
    [switch] $Fresh,

    # Sweep even when a database that is none of this harness's business is open.
    [switch] $Force,

    # Seconds to wait for the host's window to appear.
    [int] $TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $PSScriptRoot 'AccessAttach.psm1') -Force

function Find-AccessExecutable {
    $candidates = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\MSACCESS.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\MSACCESS.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\MSACCESS.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\MSACCESS.EXE"
    )
    foreach ($candidate in $candidates) { if (Test-Path $candidate) { return $candidate } }
    $fromRegistry = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msaccess.exe' -ErrorAction SilentlyContinue).'(default)'
    if ($fromRegistry -and (Test-Path $fromRegistry)) { return $fromRegistry }
    throw 'Could not locate MSACCESS.EXE.'
}

if (-not [System.IO.Path]::IsPathRooted($Database)) { $Database = Join-Path $repoRoot $Database }
if (-not (Test-Path $Database)) { throw "No database at $Database." }

if ($Fresh) {
    # THE SAME GUARD the other two launchers carry: -Fresh closes every Access on the machine, and
    # a database this harness did not open is somebody's actual work. A process whose database
    # cannot be read counts as a stranger, because unreadable is not empty.
    $mine = @(
        (Join-Path $repoRoot 'artifacts\fixtures'),
        (Join-Path $repoRoot 'artifacts\chaos')
    ) | ForEach-Object {
        if (Test-Path $_) { Get-ChildItem $_ -File | ForEach-Object { $_.FullName } }
    }

    $strangers = @()
    foreach ($running in @(Get-Process MSACCESS -ErrorAction SilentlyContinue)) {
        $held = $null
        try {
            $where = Get-AccessDatabasePath -ProcessId $running.Id
            if ($null -ne $where) { $held = @($where) }
        }
        catch { $held = $null }

        if ($null -eq $held) {
            $strangers += "pid $($running.Id) (its database could not be read)"
            continue
        }
        foreach ($one in $held) {
            if ($mine -notcontains $one) { $strangers += "$one (pid $($running.Id))" }
        }
    }

    if ($strangers.Count -gt 0 -and -not $Force) {
        throw ("-Fresh closes every Access, and these are not this harness's to close:" +
            [Environment]::NewLine + '  ' + ($strangers -join ([Environment]::NewLine + '  ')) +
            [Environment]::NewLine +
            'Close them yourself, or pass -Force if the machine is yours to sweep.')
    }

    Get-Process MSACCESS -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    # A killed Access leaves its lock file behind, and the next open then reads as a second user
    # on the database rather than as the only one.
    $lock = [System.IO.Path]::ChangeExtension($Database, '.laccdb')
    if (Test-Path $lock) { Remove-Item $lock -Force -ErrorAction SilentlyContinue }
}

# A harness terminates Access by design, and Access reads termination as a crash: on the next
# start it offers to disable what it blames, which stands in front of the thing being tested.
foreach ($version in @('16.0', '15.0')) {
    $key = "HKCU:\Software\Microsoft\Office\$version\Access\Resiliency"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

$process = Start-Process -FilePath (Find-AccessExecutable) -ArgumentList ('"{0}"' -f $Database) -PassThru
Write-Host "Started Access as process $($process.Id) on $(Split-Path -Leaf $Database)."

$access = Get-AccessApplication -ProcessId $process.Id -TimeoutSeconds $TimeoutSeconds
if ($null -eq $access) { throw "Could not reach Access $($process.Id) through its OMain frame." }

# The editor through Access's OWN ribbon command, not through $access.VBE: that property is one of
# the two "Trust access to the VBA project object model" gates, and with the switch off it comes
# back null rather than raising, so a try/catch would report success and open nothing.
$access.CommandBars.ExecuteMso('VisualBasic')
Write-Host 'Editor opened (through the ribbon command, which needs no VBA project trust).'

# The doctor, BY PID: another live session beside this one is a designed state, and the bare verb
# refuses to guess between instances.
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
