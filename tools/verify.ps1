<#
.SYNOPSIS
    Everything that must hold before a commit, in one command.

.DESCRIPTION
    The gate was being assembled by hand each time — page typecheck here, page build there,
    dotnet build, dotnet test, a Release publish, a string check for the debug door — and a
    hand-assembled gate is one someone eventually runs four fifths of. This runs the lot and
    says PASS or FAIL once, naming what failed.

    It does NOT need an editor: everything here is buildable and checkable on any machine.
    The live probes are a separate act, because they need a host — `-Live` runs them too,
    which requires an editor already open (tools\dev.ps1 -KeepOpen).

.EXAMPLE
    tools\verify.ps1
    The pre-commit gate.

.EXAMPLE
    tools\verify.ps1 -Live
    The gate, plus the standing probes against a running editor.

.EXAMPLE
    tools\verify.ps1 -Quick
    Skip the Release publish, which is the slowest step by far.
#>
[CmdletBinding()]
param(
    # Also run the standing probes against an already-open editor.
    [switch] $Live,

    # Skip the ahead-of-time Release publish.
    [switch] $Quick
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$repoRoot = Split-Path -Parent $PSScriptRoot
$pageRoot = Join-Path $repoRoot 'ui\editor'

# The SDK and the native linker are not necessarily on the machine PATH — the same reason
# dev.ps1 does this. A verify that cannot find the compiler is not a verdict about the code.
$localDotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet'
if (Test-Path (Join-Path $localDotnet 'dotnet.exe')) {
    $env:PATH = "$localDotnet;$env:PATH"
}

$vsInstaller = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer"
if (Test-Path $vsInstaller) {
    $env:PATH = "$vsInstaller;$env:PATH"
}

$results = [ordered] @{}

function Step([string] $name, [scriptblock] $work) {
    Write-Host ''
    Write-Host "==> $name" -ForegroundColor Cyan

    $clock = [Diagnostics.Stopwatch]::StartNew()
    try {
        $detail = & $work
        $clock.Stop()
        $results[$name] = [pscustomobject] @{
            Ok = $true
            Seconds = $clock.Elapsed.TotalSeconds
            Detail = $detail
        }
    } catch {
        $clock.Stop()
        $results[$name] = [pscustomobject] @{
            Ok = $false
            Seconds = $clock.Elapsed.TotalSeconds
            Detail = $_.Exception.Message
        }
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Step 'vendored spec' {
    # The page bundles the spec repo's typing helpers from a copy in ui/editor/vendor so that CI can
    # build without a neighbouring checkout. This is the step that notices when the copy and the
    # spec have parted ways -- which only this machine can notice, since CI has no spec repo to
    # compare against.
    Push-Location $pageRoot
    try {
        $out = npm run spec:check 2>&1
        $out | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the vendored spec copy is out of step' }
    } finally { Pop-Location }
    if ($out -match 'not present') { 'manifest only' } else { 'matches the spec repo' }
}

Step 'engine executable is current' {
    # The add-in launches engine\dist\xlide-engine.exe, NOT engine\dist\engine.cjs. `npm run build`
    # writes only the bundle; only `npm run package` rebuilds the executable. So an engine change
    # can be built, tested against the bundle, committed, and published, while the thing that
    # actually runs is hours old and refuses every new method as unknown.
    #
    # That happened on 2026-08-06, and the only reason it surfaced was a live session's log. This
    # step is the cheap version of that log: if any engine source is newer than the executable, the
    # executable is stale, and no amount of green elsewhere means anything.
    $engineRoot = Join-Path $repoRoot 'engine'
    $exe = Join-Path $engineRoot 'dist\xlide-engine.exe'
    if (-not (Test-Path $exe)) { throw 'engine\dist\xlide-engine.exe has never been packaged; run npm run package in engine\' }

    $builtAt = (Get-Item $exe).LastWriteTimeUtc
    $newer = @(Get-ChildItem (Join-Path $engineRoot 'src') -Recurse -File |
        Where-Object { $_.LastWriteTimeUtc -gt $builtAt })

    if ($newer.Count -gt 0) {
        $names = ($newer | Select-Object -First 4 | ForEach-Object { $_.Name }) -join ', '
        throw "engine sources are newer than the packaged executable ($names). Run: npm run package --prefix engine"
    }

    'packaged after every engine source'
}

Step 'page typecheck' {
    Push-Location $pageRoot
    try {
        npm run typecheck 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the page does not typecheck' }
    } finally { Pop-Location }
    'clean'
}

Step 'page build' {
    Push-Location $pageRoot
    try {
        $out = npm run build 2>&1
        $out | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the page bundle would not build' }
        $warnings = ($out | Select-String 'warnings: (\d+)').Matches.Groups[1].Value
        if ($warnings -ne '0') { throw "the bundle built with $warnings warning(s)" }
    } finally { Pop-Location }
    'no warnings'
}

Step 'page tests' {
    # The split tree's arithmetic, then the bundle's structure.
    Push-Location $pageRoot
    try {
        npm test 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the page tests failed' }
    } finally { Pop-Location }
    'tree algebra, bundle structure'
}

Step 'page probes (headless)' {
    foreach ($probe in 'close-confirm-page-probe.mjs', 'objbrowser-page-probe.mjs') {
        $answer = node (Join-Path $repoRoot "tools\harness\$probe") 2>&1 | Select-Object -Last 1
        if ($answer -notmatch '"pass":true') { throw "$probe did not pass" }
    }
    'close-confirm, object browser'
}

# A route table is complete on the day it is written and quietly is not, six routes later. This
# reads the routes out of the shim itself, so the documents and the client cannot fall behind the
# door without the gate saying so (2026-08-07: the reference had all of them, the driving guide
# had twenty, and one route had no client method at all).
Step 'debug api is documented' {
    $answer = node (Join-Path $repoRoot 'tools\harness\audit-routes.mjs') 2>&1
    $answer | Where-Object { $_ -notmatch '^ok ' } | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) { throw 'a debug api route is undocumented or unreachable' }
    ($answer | Select-Object -Last 1) -replace '^ok\s+', ''
}

# Named, not left to whichever directory the gate was started from. Both of these used to be
# bare `dotnet build` / `dotnet test`, which pick the solution out of the CURRENT directory — so
# the gate passed from the repo root and failed with MSB1003 from anywhere else (2026-08-07). A
# check whose answer depends on where you are standing is not a check.
$solution = Join-Path $repoRoot 'xlide_vbide.slnx'

Step 'solution build (Release)' {
    dotnet build $solution -c Release --nologo -v q 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'the solution does not build in Release' }
    'clean'
}

Step 'unit tests' {
    $out = dotnet test $solution -c Release --no-build --nologo -v q 2>&1
    $out | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'unit tests failed' }
    $passed = ($out | Select-String 'Passed:\s+(\d+)' -AllMatches).Matches |
        ForEach-Object { [int] $_.Groups[1].Value } |
        Measure-Object -Sum
    "$($passed.Sum) passed"
}

Step 'Release carries no debug api' {
    # Verified against the binary rather than trusted: the whole door is one #if away from
    # shipping, and a door that ships is a port open on a user's machine.
    $shim = Join-Path $repoRoot 'artifacts\bin\Xlide.Vbe.Shim\release_win-x64\Xlide.Vbe.Shim.dll'
    if (-not (Test-Path $shim)) { throw "no Release assembly at $shim" }

    $bytes = [IO.File]::ReadAllBytes($shim)
    $text = [Text.Encoding]::Unicode.GetString($bytes) + [Text.Encoding]::ASCII.GetString($bytes)
    foreach ($needle in '__xlideConsole', 'unknown benchmark', 'the pending result was lost', 'debug-api-') {
        if ($text.Contains($needle)) { throw "Release carries the debug api: found '$needle'" }
    }
    'no debug strings'
}

if (-not $Quick) {
    Step 'native publish (Release)' {
        dotnet publish (Join-Path $repoRoot 'src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj') `
            -c Release -r win-x64 --nologo -v q 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the ahead-of-time publish failed' }

        $published = Join-Path $repoRoot 'artifacts\publish\Xlide.Vbe.Shim\release_win-x64'
        foreach ($needed in 'ui\editor\dist\index.html', 'ui\editor\dist\editor.js') {
            if (-not (Test-Path (Join-Path $published $needed))) {
                throw "the published shim is missing $needed; it would show the native pane"
            }
        }

        if (Test-Path (Join-Path $published 'hostfxr.dll')) {
            throw 'a runtime was published beside the shim; the build is no longer ahead-of-time'
        }

        '{0:N1} MB' -f ((Get-Item (Join-Path $published 'Xlide.Vbe.Shim.dll')).Length / 1MB)
    }
}

if ($Live) {
    Step 'live probes' {
        $excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $excel) { throw 'no editor is open; start one with tools\dev.ps1 -KeepOpen' }

        # A session that has only just launched is still seeding: the engine is starting, the
        # first analysis pass has not run, and the ghost readers may not be attached. Probes
        # that assert a healthy session then fail on the truth that it is not healthy YET —
        # which is a real answer to the wrong question (2026-08-06, on a release gate).
        $discovery = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
        if (Test-Path $discovery) {
            $api = (Get-Content $discovery -Raw | ConvertFrom-Json)
            $door = "http://127.0.0.1:$($api.port)/$($api.token)"
            $ready = $false
            $deadline = (Get-Date).AddSeconds(30)
            do {
                try {
                    $doctor = Invoke-RestMethod "$door/doctor" -TimeoutSec 5
                    $ready = $doctor.engineUp -and $doctor.surfaceReady -and $doctor.ghostReadersUp
                } catch { $ready = $false }
                if (-not $ready) { Start-Sleep -Milliseconds 500 }
            } while (-not $ready -and (Get-Date) -lt $deadline)

            if (-not $ready) { throw 'the session never became healthy enough to probe' }
        }

        foreach ($probe in 'Test-DebugApi.ps1', 'Test-SplitWorkspace.ps1', 'Test-DiscardProblems.ps1', 'Test-Churn.ps1') {
            $answer = powershell -NoProfile -ExecutionPolicy Bypass `
                -File (Join-Path $repoRoot "tools\harness\$probe") 2>&1
            $answer | Out-Host
            $verdict = $answer | Select-String 'RESULT: (PASS|FAIL)' | Select-Object -Last 1
            if ("$verdict" -notmatch 'PASS') {
                # Name the checks, not just the file. A summary that says only "did not pass"
                # sends the reader back through the scrollback for the one line that matters,
                # and at the end of a run that line has usually scrolled away (2026-08-06).
                $broken = @($answer | Select-String '\s+FAIL' | ForEach-Object { $_.Line.Trim() })
                $named = if ($broken.Count -gt 0) { ': ' + ($broken -join '; ') } else { '' }
                throw "$probe did not pass$named"
            }
        }
        'debug api, split workspace, churn'
    }
}

Write-Host ''
Write-Host '---------------------------------------------------------------' -ForegroundColor DarkGray
foreach ($entry in $results.GetEnumerator()) {
    $mark = if ($entry.Value.Ok) { 'PASS' } else { 'FAIL' }
    $colour = if ($entry.Value.Ok) { 'Green' } else { 'Red' }
    Write-Host ("  {0,-32} {1,-5} {2,6:N1}s  {3}" -f
        $entry.Key, $mark, $entry.Value.Seconds, $entry.Value.Detail) -ForegroundColor $colour
}

$failed = @($results.Values | Where-Object { -not $_.Ok })
Write-Host ''
if ($failed.Count -eq 0) {
    Write-Host "RESULT: PASS - $($results.Count) steps" -ForegroundColor Green
    exit 0
}

Write-Host "RESULT: FAIL - $($failed.Count) of $($results.Count) steps" -ForegroundColor Red
exit 1
