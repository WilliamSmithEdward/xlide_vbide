<#
.SYNOPSIS
    Everything that must hold before a commit, in one command.

.DESCRIPTION
    The gate was being assembled by hand each time â€” page typecheck here, page build there,
    dotnet build, dotnet test, a Release publish, a string check for the debug door â€” and a
    hand-assembled gate is one someone eventually runs four fifths of. This runs the lot and
    says PASS or FAIL once, naming what failed.

    It does NOT need an editor: everything here is buildable and checkable on any machine.
    The live probes are a separate act, because they need a host â€” `-Live` runs them too,
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

# The SDK and the native linker are not necessarily on the machine PATH â€” the same reason
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

    # The ANALYZER counts as engine source. It lives in the neighbouring checkout and is bundled
    # INTO this executable, so a pull over there changes what the add-in runs without touching a
    # single file in this repository. Watching engine\src alone would call that stale executable
    # current, which is the one answer this step exists to never give.
    $watched = @(Join-Path $engineRoot 'src')
    $analyzer = Join-Path (Split-Path -Parent $repoRoot) 'xlide_vscode\src'
    if (Test-Path $analyzer) { $watched += $analyzer }

    $newer = @(Get-ChildItem $watched -Recurse -File -Include *.ts, *.mjs, *.js |
        Where-Object { $_.LastWriteTimeUtc -gt $builtAt })

    if ($newer.Count -gt 0) {
        $names = ($newer | Select-Object -First 4 | ForEach-Object { $_.Name }) -join ', '
        throw "engine sources are newer than the packaged executable ($names). Run: npm run package --prefix engine"
    }

    if ($watched.Count -eq 1) {
        # Said out loud rather than passed over: the check ran, but half of what it is meant to
        # watch was not there to watch.
        return 'packaged after every engine source (the analyzer checkout was not found)'
    }

    'packaged after every engine source, analyzer included'
}

Step 'no variant is read as an object' {
    # THE ONE LEAK THE SWEEP CANNOT SEE.
    #
    # `variant.As<object>()` asks the runtime for a managed wrapper over whatever interface the
    # variant holds. That wrapper is nobody's: it is not taken through ComRuntime, so the live
    # count never sees it, and it is never disposed, so the FINALIZER thread releases it - which
    # for an apartment-threaded editor object is an access violation the runtime cannot throw. It
    # FailFasts, and Excel goes with it.
    #
    # It has killed Excel twice, months apart in code terms: once as the fallback of the
    # variant-to-text conversion (2026-08-07), and once more as the DEFAULT branch of the same
    # switch, which the first fix did not cover because a variant carrying an interface with a
    # flag on it - `VT_BYREF | VT_DISPATCH`, `VT_ARRAY | VT_VARIANT` - matches neither
    # `VT_DISPATCH` nor `VT_UNKNOWN` by name and fell through (2026-08-08).
    #
    # com-leak.mjs cannot catch either one. It measures the wrappers this product TOOK, and these
    # are taken behind its back; it read a balanced 13 through both. So the guard is here, on the
    # SHAPE of the code, because that is the only place this particular defect is visible before
    # it is a crash report.
    $offenders = Get-ChildItem (Join-Path $repoRoot 'src') -Recurse -Include *.cs |
        Select-String -Pattern '\.As<\s*object\s*>\s*\(' |
        Where-Object { $_.Line -notmatch '^\s*(//|\*|/\*)' }

    if ($offenders) {
        $where = ($offenders | ForEach-Object { "$($_.Filename):$($_.LineNumber)" }) -join ', '
        throw ("a variant is being read as an object ($where). That builds a COM wrapper nothing " +
            'owns and the finalizer thread releases it, which ends Excel. Name the variant type ' +
            'and describe it, or take it through ComRuntime.')
    }

    'no path materialises a wrapper the counter cannot see'
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

Step 'engine language matrix' {
    # Non-ASCII through the real engine over the real pipe, for every script the companion
    # product supports. This repo's risk is not the companion's: it never decodes the .xlsm's
    # bytes, so the exposure is OFFSET ARITHMETIC. Every language feature names a position as a
    # number of units into the source, and a byte count or a code-point count anywhere in the
    # chain drifts by the width of the non-ASCII text to its left -- invisible in an English
    # module, and it corrupts an edit rather than failing.
    #
    # A gate step because it cannot be noticed any other way: the fixtures are English.
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        node test/language.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the engine language matrix failed' }
    } finally { Pop-Location }
    '18 scripts through open, diagnose, outline, definition and rename'
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
Step 'debug api is documented and driven' {
    $answer = node (Join-Path $repoRoot 'tools\harness\audit-routes.mjs') 2>&1
    $answer | Where-Object { $_ -notmatch '^ok ' } | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) { throw 'a debug api route is undocumented, unreachable, or driven by nothing' }
    ($answer | Select-Object -Last 1) -replace '^ok\s+', ''
}

# Named, not left to whichever directory the gate was started from. Both of these used to be
# bare `dotnet build` / `dotnet test`, which pick the solution out of the CURRENT directory â€” so
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
    <#
        EACH SUITE GETS THE FIXTURE IT WAS WRITTEN FOR.

        The live half used to run every probe and suite against whatever workbook happened to be
        open, and they do not want the same one: the api probe and the Immediate suite need the
        debug fixture's Runner, and the format and three-copies suites need the rename fixture's
        HelpersExtra. Whichever was open, the others failed - so the live gate could not pass, and
        a gate that cannot pass is a gate that gets skipped rather than read (2026-08-08).

        Opening a fixture costs an Excel restart, which is why this is grouped rather than done
        per suite: two launches, not six.
    #>
    function Use-Fixture([string] $fixture) {
        & (Join-Path $repoRoot 'tools\harness\Start-Excel.ps1') `
            -Workbook (Join-Path $repoRoot (Join-Path 'artifacts\fixtures' $fixture)) -Fresh | Out-Host

        $excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $excel) { throw "Excel did not start on $fixture" }
        return $excel
    }

    Step 'live probes' {
        $excel = Use-Fixture 'DebugFixture.xlsm'

        # A session that has only just launched is still seeding: the engine is starting, the
        # first analysis pass has not run, and the ghost readers may not be attached. Probes
        # that assert a healthy session then fail on the truth that it is not healthy YET â€”
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

    # THE NODE SUITES, which existed and passed and which nothing ran.
    #
    # A suite nobody runs is a suite that rots, and these three cover the defects that cost the
    # most this week: where a squiggle lands after Format Module, whether the workbook, the surface
    # and the analyzer are holding the same text after every operation that touches a module, and
    # whether the Immediate window answers what a line came to. All three were written against
    # bugs that had already shipped.
    #
    # They report "N passed, M failed" rather than a RESULT line, so they are read that way.
    Step 'language and editor suites' {
        $ran = @()
        # analysis-freshness runs against whatever workbook is open: it brings its own two
        # modules, uniquely named per run, and takes them away again. It guards a SILENT failure -
        # a squiggle that should be drawn and is not, because a module's findings were reused
        # after another module changed the signature it calls.
        #
        # It was out of this list while it was flaky. Every cause turned out to be in the suite or
        # next to it, the last two being fixed module names that let one run inherit the previous
        # run's answers, and a helper that dropped the field its own timing bound was built from.
        # Nine checks, four runs clean on the rename fixture and three on the one holding a module
        # at VBA's line ceiling (2026-08-08).
        # Grouped by the fixture each needs, so the gate opens two rather than failing five.
        $plan = [ordered] @{
            # module-sync writes into a temporary folder of its own and takes back every module it
            # adds. Its last section is the one worth having: it applies the same import through
            # the DIALOG and through the route and compares the result byte for byte, which is the
            # only check that would notice the api growing its own idea of what an import means.
            # module-sync runs TWICE, once per planner. The two are meant to decide identically,
            # and the only way to know they still do is to put the same 32 checks through both.
            # It is also the one thing that catches a silent fall back to the built-in planner,
            # which is how a green run against the wrong implementation happened (2026-08-09).
            'DebugFixture.xlsm'  = @('immediate-watch.mjs', 'analysis-freshness.mjs',
                                     'module-sync.mjs xlide', 'module-sync.mjs builtIn')
            'RenameFixture.xlsm' = @('format-positions.mjs', 'three-copies.mjs')
        }

        foreach ($fixture in $plan.Keys) {
          Use-Fixture $fixture | Out-Null
          foreach ($suite in $plan[$fixture]) {
            # A suite may name arguments after its file, which is how module-sync says which
            # planner this run is about.
            $parts = $suite -split ' '
            $answer = node (Join-Path $repoRoot "tools\harness\$($parts[0])") @($parts | Select-Object -Skip 1) 2>&1
            $answer | Out-Host

            $verdict = $answer | Select-String '(\d+) passed, (\d+) failed' | Select-Object -Last 1
            if (-not $verdict) { throw "$suite reported no verdict" }
            if ("$verdict" -notmatch ', 0 failed') {
                # The failing checks by name, for the same reason the probes above name theirs:
                # at the end of a run the one line that matters has usually scrolled away.
                $broken = @($answer | Select-String '^\s*FAIL' | ForEach-Object { $_.Line.Trim() })
                $named = if ($broken.Count -gt 0) { ': ' + ($broken -join '; ') } else { '' }
                throw "$suite did not pass$named"
            }

            $ran += "$suite $("$verdict".Trim())"
          }
        }

        $ran -join '; '
    }

    # NO LEAKS, AS A RELEASE-BLOCKING PROPERTY rather than a quality aspiration (the developer,
    # 2026-08-07). A leaked COM wrapper in this product does not waste memory, it kills the host:
    # the editor's objects are apartment-threaded, and one released by the finalizer thread is an
    # access violation the runtime cannot throw, so it FailFasts Excel. One missing Dispose leaked
    # 441 wrappers per project() call and killed Excel four times in a day, reported as three
    # different faults against three different libraries with nothing connecting them.
    #
    # Every read route, many rounds, live count before and after. See lessons.md entry 36.
    Step 'no leaks' {
        # Whatever the suites left open is fine: the sweep brings its own state and puts it back.
        $excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $excel) { $excel = Use-Fixture 'DebugFixture.xlsm' }

        $answer = node (Join-Path $repoRoot 'tools\harness\com-leak.mjs') 2>&1
        $answer | Out-Host

        $verdict = $answer | Select-String '(\d+) passed, (\d+) failed' | Select-Object -Last 1
        if (-not $verdict) { throw 'the leak sweep reported no verdict' }
        if ("$verdict" -notmatch ', 0 failed') {
            $leaking = @($answer | Select-String '^FAIL' | ForEach-Object { $_.Line.Trim() })
            throw ('wrappers are being leaked: ' + ($leaking -join '; '))
        }

        "$verdict".Trim()
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
