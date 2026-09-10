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
    window, so there is no chance of adopting a different one. That call lives in WordAttach.psm1,
    shared with the fixture generator.

    THE EDITOR IS WHAT LOADS THE ADD-IN, not the host. ExecuteMso('VisualBasic') is Word running
    its own Developer > Visual Basic button - the same idMso as Excel's - and is not gated by
    "Trust access to the VBA project object model".

.EXAMPLE
    tools\harness\Start-Word.ps1
    Starts on the scratch document (created on first use).

.EXAMPLE
    tools\harness\Start-Word.ps1 -Document artifacts\fixtures\WordFixture.docm -Fresh
#>
[CmdletBinding()]
param(
    # The document or documents to open. Relative paths are taken from the repository root.
    [string[]] $Document,

    # Close the Words this harness started first - the ones holding a fixture or chaos document -
    # each by its own id. Word only, and every other Word stays, whoever started it: a stop by
    # name ended other automations' hosts mid-statement (#24). When one stays, the fixture starts
    # in a process of its own.
    [switch] $Fresh,

    # Close EVERY Word the census saw, strangers included, each by id.
    [switch] $Force,

    # Start a SEPARATE Word process rather than letting Word hand the document to the one already
    # running. Word, like Excel, gives a document on its command line to a running instance, and
    # the door would then be in a process this harness never started - the developer's own Word,
    # with their own document open beside the fixture. `/w` is Word's own switch for a new
    # instance, and it takes a document (measured 2026-09-10: the second document opened in a
    # process of its own, titled with its own name, while the first Word stood).
    [switch] $Separate,

    # Seconds to wait for the host's window to appear.
    [int] $TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

Import-Module (Join-Path $PSScriptRoot 'WordAttach.psm1') -Force

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
        $blank = $null
        try {
            $blank = $maker.Documents.Add()
            $blank.SaveAs2($scratch, 13)  # wdFormatXMLDocumentMacroEnabled
            $blank.Close($false)
        }
        finally {
            if ($null -ne $blank) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($blank) | Out-Null }
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
    # THE SAME RULE Start-Excel.ps1 carries, and for the same reason: only a Word holding a
    # document this harness put there is this harness's to close, and it is closed by id. A
    # process whose documents cannot be read is left standing, because unreadable is not empty,
    # and so is one holding somebody's actual work; a stop by name ended other automations'
    # hosts mid-statement (#24). -Force closes every Word the census saw.
    $mine = @(
        (Join-Path $repoRoot 'artifacts\fixtures'),
        (Join-Path $repoRoot 'artifacts\chaos'),
        (Join-Path $PSScriptRoot 'fixtures')
    ) | ForEach-Object {
        if (Test-Path $_) { Get-ChildItem $_ -File | ForEach-Object { $_.FullName } }
    }

    $ours = @()
    $theirs = @()
    foreach ($running in @(Get-Process WINWORD -ErrorAction SilentlyContinue)) {
        # Every wrapper the read takes is given back inside the module, while the process can
        # still answer the release; a wrapper finalised after a kill makes DCOM start a fresh
        # hidden host to answer it (2026-09-08).
        $held = Get-WordDocumentPaths -ProcessId $running.Id

        if ($null -eq $held) {
            $theirs += "pid $($running.Id) (no document window this harness can read)"
        }
        elseif (@($held | Where-Object { $mine -contains $_ }).Count -gt 0) {
            $ours += $running.Id
        }
        else {
            $what = if ($held.Count -gt 0) { $held -join ', ' } else { 'no document open' }
            $theirs += "pid $($running.Id) ($what)"
        }
    }

    # The collector runs NOW, with every Word still alive: a wrapper missed above is released
    # while its server can still answer, rather than after the stop below.
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $stopping = @(if ($Force) { Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } } else { $ours })
    foreach ($id in $stopping) {
        Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
    }
    foreach ($id in $stopping) {
        $stopped = Get-Process -Id $id -ErrorAction SilentlyContinue
        if ($null -ne $stopped) { $stopped.WaitForExit(10000) | Out-Null }
    }
    if ($stopping.Count -gt 0) {
        $whose = if ($Force) { 'every Word' } else { "this harness's Word" }
        Write-Host "Closed $whose by id: $($stopping -join ', ')."
    }

    if (-not $Force -and $theirs.Count -gt 0) {
        Write-Host ("Left standing, not this harness's to close:" + [Environment]::NewLine + '  ' +
            ($theirs -join ([Environment]::NewLine + '  ')))
        # SEPARATE, or the fixture opens inside theirs: Word hands a document on its command line
        # to an instance already running, and the door would then be in a process this harness
        # never started.
        $Separate = $true
    }

    # A killed Word leaves its owner file beside the document - `~$` and the name less its first
    # letter or two - and the next open then reads as locked by another user, which makes the
    # document read-only and every save in a suite a refusal. Stale by construction here: a Word
    # holding one of these documents was this harness's and has just been closed.
    foreach ($one in $Document) {
        $leaf = Split-Path -Leaf $one
        Get-ChildItem (Split-Path -Parent $one) -Filter '~$*.docm' -Force -ErrorAction SilentlyContinue |
            Where-Object { $leaf.EndsWith($_.Name.Substring(2), [StringComparison]::OrdinalIgnoreCase) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
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
if ($Separate) { $arguments = @('/w') + $arguments }
$names = ($Document | ForEach-Object { Split-Path -Leaf $_ }) -join ', '
$apart = if ($Separate) { ', in a process of its own' } else { '' }

$process = Start-Process -FilePath (Find-WordExecutable) -ArgumentList $arguments -PassThru
Write-Host "Started Word as process $($process.Id) on $names$apart."

$window = Get-WordDocumentWindow -ProcessId $process.Id -TimeoutSeconds $TimeoutSeconds
if ($null -eq $window) { throw "Could not reach Word $($process.Id) through its window." }

$word = $window.Application
$word.DisplayAlerts = 0

# The editor is opened through Word's OWN ribbon command, not through $word.VBE, for exactly
# Start-Excel.ps1's reason: Application.VBE is what the trust setting refuses, and ExecuteMso
# is the host pressing its own button.
$commandBars = $word.CommandBars
$commandBars.ExecuteMso('VisualBasic')
Write-Host 'Editor opened (through the ribbon command, which needs no VBA project trust).'

# GIVEN BACK NOW, not left for the collector. The fixture generator runs this script in its own
# process and stops the very Word it attached to when it is done; a wrapper finalised after that
# kill makes DCOM start a fresh hidden Word to answer the release (2026-09-08). Everything after
# this line talks to the door over HTTP and needs none of these.
foreach ($wrapper in @($commandBars, $word, $window)) {
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrapper) | Out-Null } catch { }
}
$commandBars = $null
$word = $null
$window = $null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

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
