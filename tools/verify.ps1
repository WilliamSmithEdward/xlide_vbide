<#
.SYNOPSIS
    Everything that must hold before a commit, in one command.

.DESCRIPTION
    The gate was being assembled by hand each time - page typecheck here, page build there,
    dotnet build, dotnet test, a Release publish, a string check for the api door - and a
    hand-assembled gate is one someone eventually runs four fifths of. This runs the lot and
    says PASS or FAIL once, naming what failed.

    It does NOT need an editor: everything here is buildable and checkable on any machine.
    The live probes are a separate act, because they need a host - `-Live` runs them too,
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

    # The pre-release tier: everything -Live runs, then the suites that earn depth rather than
    # speed - a third fixture launch, half a minute of deliberate settle sleeps, a randomized
    # walk. Run before a release, not before a commit.
    [switch] $Deep,

    # Skip the ahead-of-time Release publish.
    [switch] $Quick
)

if ($Deep) { $Live = $true }

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

$repoRoot = Split-Path -Parent $PSScriptRoot
$pageRoot = Join-Path $repoRoot 'ui\editor'

# The SDK and the native linker are not necessarily on the machine PATH - the same reason
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
    # Asked through the shared check, so the release path cannot answer this differently - it used
    # to answer it not at all when run with -SkipGate.
    $engineRoot = Join-Path $repoRoot 'engine'
    $exe = Join-Path $engineRoot 'dist\xlide-engine.exe'
    $builtAt = if (Test-Path $exe) { (Get-Item $exe).LastWriteTimeUtc } else { [datetime]::MinValue }

    $newer = @(& (Join-Path $PSScriptRoot 'Test-EngineCurrent.ps1') -RepoRoot $repoRoot)

    $analyzerPath = Join-Path (Split-Path -Parent $repoRoot) 'xlide_vscode\src'
    $covers = if (Test-Path $analyzerPath) { 'analyzer included' } else { 'the analyzer checkout was not found' }

    if ($newer.Count -eq 0) {
        return "packaged after every engine source, $covers"
    }

    <#
        PACKAGED, NOT COMPLAINED ABOUT.

        This threw and told the developer to run `npm run package` themselves, which is a gate
        naming a command it could have run. It fired on four runs where nothing was wrong, and the
        cost of that is not the reading: it is that a step which is usually noise gets answered by
        rerunning it, and once it meant the measurements had been taken against an engine older
        than the analyzer under test.

        Packaging only when something IS newer, rather than on every run: the injection writes a
        90 MB executable and takes the better part of a minute, which is most of this gate's whole
        runtime to spend on the common case where nothing changed.
    #>
    $names = ($newer | Select-Object -First 4 | ForEach-Object { $_.Name }) -join ', '
    Write-Host "  $($newer.Count) engine source(s) newer than the executable ($names); packaging" -ForegroundColor Yellow

    # TYPECHECKED BEFORE IT IS PACKAGED, because this gate ships the result.
    #
    # The page is typechecked here and the solution is compiled here; the engine was the one layer
    # that went into the installer without either. esbuild strips types and does not check them, so
    # a type error in engine\src bundled cleanly, packaged cleanly, and passed every step of this
    # gate - and the language matrix below only exercises the paths it happens to call. CI does run
    # `check-types`, which means the local gate could ship an installer that CI would then refuse.
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        npm run check-types 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the engine does not typecheck; it will not be packaged' }
    }
    finally {
        Pop-Location
    }

    # A HOST HOLDS THE EXECUTABLE IT IS RUNNING, and the injection fails with EBUSY behind whatever
    # is filtering npm's output. Named here rather than left as a copy error two hundred lines up
    # (2026-08-10: two rounds spent diagnosing a fix that had simply never been built).
    $holders = @(Get-Process EXCEL, xlide-engine -ErrorAction SilentlyContinue)
    if ($holders.Count -gt 0) {
        $who = ($holders | ForEach-Object { "$($_.ProcessName) $($_.Id)" }) -join ', '
        throw ("the engine is stale and cannot be packaged while it is running ($who holds it). " +
            'Close the editor and run the gate again.')
    }

    npm run package --prefix (Join-Path $repoRoot 'engine') 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'packaging the engine failed' }

    # Read back rather than assumed: npm can exit zero having written nothing when the injection
    # step is the part that failed.
    $rebuiltAt = (Get-Item $exe).LastWriteTimeUtc
    if ($rebuiltAt -le $builtAt) {
        throw 'packaging reported success and the executable did not change'
    }

    "packaged $($newer.Count) newer source(s) into the executable, $covers"
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

# CI runs this and the local gate did not, so a dispatcher type error passed the whole local
# ladder - bundling does not typecheck - and failed only on the push (2026-08-13). The gate
# covers what CI covers, or the gate's PASS is a claim about a different build.
Step 'engine typecheck' {
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        npm run check-types 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the engine does not typecheck' }
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

Step 'engine host-supplied facts' {
    # THE FACTS A CODE PANE CANNOT CARRY, and the analyzer therefore has to be told: a form's
    # designer-declared controls, and a class module's default-instance flag. The host reads
    # them and the engine passes them through (xlide_vscode#17..#20, #47..#50). The middle link
    # of that chain is this repo's, and this holds it headlessly - the call-colouring acceptance
    # table, the Control base-class merge, and the meType gate - so a break fails a commit
    # rather than a live run.
    #
    # Every one of them is THREE-STATE, and every defect in the family has been the same
    # mistake: reading absent as the negative one, and reporting correct code as broken.
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        node test/forms.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the engine form suite failed' }

        # A form whose designer could not be READ is a form whose control names cannot be
        # judged, and that is not the same as a form with no controls. Showing a UserForm makes
        # the VBE withhold its designer, so the host seeds no list - and the rule treated the
        # unknown as an emptiness and turned the form's own code-behind red (#5, fixed upstream
        # as xlide_vscode#48). Held here because the seed is this repo's middle link.
        node test/form-unvouched.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the unvouched-form suite failed' }

        # And a CLASS module's `VB_PredeclaredId`, which decides whether its own bare name is a
        # value or a type - so whether `Ticket.ChangeTest` compiles or is `Variable not defined`.
        # The attribute never appears in a code pane, so the host reads it out of the saved
        # package (Core/Vba/SavedModules.cs) and it rides the seed. Held here because the wrong
        # answer is not a missed finding: it is a red squiggle under working code, once for every
        # use of every predeclared singleton in the project.
        node test/class-predeclared.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the predeclared-class suite failed' }
    } finally { Pop-Location }
    'the #20 acceptance table, the Control base merge, the unvouched form, the predeclared class'
}

Step 'engine knowledge routes' {
    # WHAT THE LANGUAGE SERVICE KNOWS, SERVED AS DATA: the rule catalogue answering before
    # initialize, a host's object model by type name, and the host registry switching whole -
    # a project opened as word answers WORD's model, and an unmodelled host says known:false
    # rather than wearing Excel's types. The suite existed with 21 green checks and NOTHING
    # collected them - not this gate, not CI - found on the 2026-08-28 hunt. The api's model
    # and analyzer routes answer FROM these routes, so a break here is an agent taught wrong
    # terrain by a door that sounds sure of itself.
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        node test/knowledge.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the knowledge-route suite failed' }
    } finally { Pop-Location }
    'the catalogue before initialize, per-host models, the registry switching whole'
}

Step 'engine inline comment features' {
    # THE TWO INLINE COMMENT SYNTAXES the shared analyzer defines, held through this engine's
    # wrapper: `' @xlide-analysis-disable-*` suppression directives (every scope, plus the
    # directive diagnostics for malformed ones) and `'''` XML doc comments riding hover,
    # completion and signature help. Both could be dropped by this repo's glue without upstream
    # noticing - the worker filter, the `documentation` field, and the suppression quick fix all
    # cross it - and the owner's rule for this library is that everything stays INLINE, with no
    # sidecar metadata files, so the inline route is the whole feature.
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        node test/inline-comments.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the inline-comment suite failed' }

        # And the machine-wide layer beside the inline one: the rule catalog with each rule's
        # LEGAL moves, an override silencing a warning rule, the permitted error-to-warning
        # downgrade, and the illegal move the engine ignores - which is why the shim refuses it
        # in words before it gets there.
        node test/rule-overrides.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'the rule-override suite failed' }
    } finally { Pop-Location }
    'suppression scopes, doc comments, the suppression fix, and the guarded rule overrides'
}

Step 'generated VBA module casing' {
    # VBA cases identifiers project-wide to the latest declaration it sees, so a lowercase
    # parameter in a module this product INSTALLS re-spells the developer's own code - every
    # `.Value` to `.value`, every `Err.Number` to `Err.number` - permanently, because
    # XlideAssert stays in the workbook. It happened (xlide_vscode#38), and the four names in
    # that report were not the whole set: two more were found only by measuring
    # (xlide_vbide#3).
    #
    # So it is measured, against the analyzer's own tables - the four host object models and
    # the VBA runtime, the same oracle the sibling product's test uses. A gate step because
    # nothing else can see it: the module compiles and runs perfectly while it is renaming
    # the project around it.
    Push-Location (Join-Path $repoRoot 'engine')
    try {
        node test/module-casing.mjs 2>&1 | Out-Host
        if ($LASTEXITCODE -ne 0) { throw 'a generated VBA module re-cases names the host models own' }
    } finally { Pop-Location }
    'three generated modules held to 15k canonical names'
}

Step 'page probes (headless)' {
    # close-confirm-page-probe.mjs is not missing: it runs inside the close-confirm step below,
    # which drives the same file as one of its three legs. Listing it here too would launch
    # Edge twice for the same answer.
    $probes = 'objbrowser-page-probe.mjs', 'tree-page-probe.mjs', 'boot-error-page-probe.mjs', 'sole-workbook-page-probe.mjs', 'drop-page-probe.mjs'
    foreach ($probe in $probes) {
        $answer = node (Join-Path $repoRoot "tools\harness\$probe") 2>&1 | Select-Object -Last 1
        if ($answer -notmatch '"pass":true') { throw "$probe did not pass" }
    }
    'object browser, tree rows, boot failure, sole workbook'
}

# THE CLOSE-CONFIRM WRAPPER, which needs no Excel: seam greps across the page, the shim and the
# PUBLISHED bundle (the stale-deploy tripwire), the page's modal flow headless, and the engine's
# live-copy contract over its own pipe. Two of its three legs ran NOWHERE before this step: the
# seams - one of which asserted a shape that had been deliberately removed and failed on every
# commit for weeks, unnoticed, precisely because nothing ran it - and engine-live-probe.mjs, the
# only suite anywhere that sends didChange over the engine pipe. That is the contract that keeps
# the Problems pane following a revert, and the nearest live suite passed with the fix disabled,
# so this is the one discriminating check for it (triaged 2026-08-12).
Step 'close confirm, seams to engine' {
    $answer = powershell -NoProfile -ExecutionPolicy Bypass `
        -File (Join-Path $repoRoot 'tools\harness\Test-CloseConfirm.ps1') 2>&1
    $answer | Where-Object { $_ -match 'FAIL|skip' } | ForEach-Object { Write-Host "  $_" }
    $verdict = $answer | Select-String 'RESULT: (PASS|FAIL)' | Select-Object -Last 1
    if ("$verdict" -notmatch 'PASS') {
        $broken = @($answer | Select-String 'FAIL' | ForEach-Object { $_.Line.Trim() })
        $named = if ($broken.Count -gt 0) { ': ' + ($broken -join '; ') } else { '' }
        throw "Test-CloseConfirm.ps1 did not pass$named"
    }
    'seams, page modal flow, engine live-copy contract'
}

# A route table is complete on the day it is written and quietly is not, six routes later. This
# reads the routes out of the shim itself, so the documents and the client cannot fall behind the
# door without the gate saying so (2026-08-07: the reference had all of them, the driving guide
# had twenty, and one route had no client method at all).
Step 'xlide api is documented and driven' {
    $answer = node (Join-Path $repoRoot 'tools\harness\audit-routes.mjs') 2>&1
    $answer | Where-Object { $_ -notmatch '^ok ' } | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) { throw 'a xlide api route is undocumented, unreachable, or driven by nothing' }
    ($answer | Select-Object -Last 1) -replace '^ok\s+', ''
}

# Named, not left to whichever directory the gate was started from. Both of these used to be
# bare `dotnet build` / `dotnet test`, which pick the solution out of the CURRENT directory - so
# the gate passed from the repo root and failed with MSB1003 from anywhere else (2026-08-07). A
# check whose answer depends on where you are standing is not a check.
$solution = Join-Path $repoRoot 'xlide_vbide.slnx'

Step 'solution build (Release)' {
    dotnet build $solution -c Release --nologo -v q 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'the solution does not build in Release' }
    'clean'
}

Step 'unit tests' {
    # THROUGH THE SIGNED HOST, not the generated exe. `dotnet test` has xUnit v3's adapter
    # launch Xlide.Vbe.Core.Tests.exe, and that launch is what Smart App Control judges - a
    # fresh hash with no reputation on every rebuild. On 2026-08-12 it started blocking the
    # exe in BOTH configurations ("An Application Control policy has blocked this file"),
    # turning this step red on code that had passed hours earlier byte for byte. A v3 test
    # assembly also runs as `dotnet exec <dll>` inside dotnet.exe, which is Microsoft-signed,
    # so no unsigned process is launched and the per-build lottery ends. The machine's policy
    # is the developer's choice and stays; CI keeps plain `dotnet test`, because runners have
    # no such policy and the standard tooling is worth keeping where it works.
    $testsDll = Join-Path $repoRoot 'artifacts\bin\Xlide.Vbe.Core.Tests\release\Xlide.Vbe.Core.Tests.dll'
    if (-not (Test-Path $testsDll)) { throw "no test assembly at $testsDll; the build step above writes it" }

    $out = dotnet exec $testsDll 2>&1
    $out | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'unit tests failed' }

    $summary = $out | Select-String 'Total:\s*(\d+), Errors:\s*(\d+), Failed:\s*(\d+)' | Select-Object -Last 1
    if (-not $summary) { throw 'the test run reported no summary' }
    $total = [int] $summary.Matches[0].Groups[1].Value
    $errors = [int] $summary.Matches[0].Groups[2].Value
    $failed = [int] $summary.Matches[0].Groups[3].Value
    if ($total -eq 0 -or $errors -ne 0 -or $failed -ne 0) {
        throw "unit tests: $total total, $errors errors, $failed failed"
    }

    "$total passed"
}

Step 'Release ships the api shut, and ships no unauthenticated door' {
    # THE ONE THING ABOUT THE API WORTH PROVING MECHANICALLY. The api is in the release build now
    # by design, so the old check - that none of it shipped - would fail for the right reason and
    # tell nobody anything. What matters instead is which way the shipped build LEANS, and that
    # is a `#if` somebody could write the wrong way round without any test noticing.
    #
    # It is readable because the lean is a const and the log line is a ternary over it: the
    # compiler folds it, so exactly one of the two phrases is in the binary. Proving the absent
    # one is absent is what makes this a real check rather than a search that always succeeds.
    $shim = Join-Path $repoRoot 'artifacts\bin\Xlide.Vbe.Shim\release_win-x64\Xlide.Vbe.Shim.dll'
    if (-not (Test-Path $shim)) { throw "no Release assembly at $shim" }

    # BOTH ALIGNMENTS. A .NET string literal lives in the #US heap as UTF-16, and nothing makes it
    # start on an even byte: decoding from offset 0 alone garbles every literal that happens to
    # land on an odd one, which reads as "not found". This scan had one alignment, so a needle that
    # WAS in the binary could report absent and the step would pass having checked nothing. Caught
    # 2026-08-22 by a needle known to be present reading False in both builds.
    $bytes = [IO.File]::ReadAllBytes($shim)
    $text = [Text.Encoding]::Unicode.GetString($bytes) +
        [Text.Encoding]::Unicode.GetString($bytes, 1, $bytes.Length - 1) +
        [Text.Encoding]::ASCII.GetString($bytes)

    $shut = 'this build keeps the door shut unless told otherwise'
    $open = 'this build opens the door unless told otherwise'
    if (-not $text.Contains($shut)) { throw "Release does not say it ships the api shut" }
    if ($text.Contains($open)) { throw "RELEASE SHIPS THE API OPEN: found '$open'" }

    # And the api arrived, rather than the whole thing having been compiled out by accident.
    if (-not $text.Contains('xlide-api-')) { throw 'Release carries no api at all' }

    # The DevTools protocol is NOT part of the api and must never ship: it takes no token at all,
    # and `--remote-allow-origins=*` means anything local that reaches the port drives the page.
    foreach ($needle in '--remote-debugging-port', '--remote-allow-origins') {
        if ($text.Contains($needle)) { throw "Release carries an unauthenticated door: found '$needle'" }
    }
    'shut by default, no devtools'
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
        # A name may join several workbooks with ' + ': they open on one command line into ONE
        # Excel, which is what makes them one session and one door - the state the
        # cross-workbook defect class lives in, and the state no gate session had until
        # 2026-08-12 (two of that week's defects lived exactly there).
        $books = @($fixture -split ' \+ ' | ForEach-Object {
            Join-Path $repoRoot (Join-Path 'artifacts\fixtures' $_.Trim())
        })
        & (Join-Path $repoRoot 'tools\harness\Start-Excel.ps1') -Workbook $books -Fresh | Out-Host

        $excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $excel) { throw "Excel did not start on $fixture" }
        return $excel
    }

    # One fixture group: a fresh launch, then each node suite, each held to the ONE verdict
    # spelling. A suite may name arguments after its file, which is how module-sync says which
    # planner a run is about and the walk names its step count.
    function Invoke-SuiteGroup([string] $fixture, [string[]] $suites) {
        Use-Fixture $fixture | Out-Null
        $ran = @()
        foreach ($suite in $suites) {
            $parts = $suite -split ' '
            $answer = node (Join-Path $repoRoot "tools\harness\$($parts[0])") @($parts | Select-Object -Skip 1) 2>&1
            $exitCode = $LASTEXITCODE
            $answer | Out-Host

            $verdict = $answer | Select-String '(\d+) passed, (\d+) failed' | Select-Object -Last 1
            if (-not $verdict) { throw "$suite reported no verdict" }

            # A GREEN VERDICT AND A NON-ZERO EXIT DISAGREE, and the verdict is the one that lies.
            # A suite that dies after its first check prints what it managed - "1 passed, 0
            # failed" - and every guard below reads that as a pass, because they only ever look
            # at the text. write-rollback did exactly that on 2026-08-29: it threw on its second
            # step, reported one passing check out of four, and the gate went green over a suite
            # that never reached the defect it exists to pin (#15). The exception it died on set
            # the exit code, which nothing was reading.
            if ($exitCode -ne 0 -and "$verdict" -match ', 0 failed') {
                throw ("$suite exited $exitCode while reporting '$("$verdict".Trim())' - it died " +
                    'after its checks rather than failing one, so the verdict is only what it ' +
                    'reached, not what it set out to do')
            }
            # "0 passed, 0 failed" is not a verdict either: it is what a suite prints when it
            # died before its first check, and reading it as green is how a run that checked
            # NOTHING passed a gate (2026-08-12, properties-pane refused at its first add).
            if ("$verdict" -match '^\s*0 passed, 0 failed') { throw "$suite ran zero checks, which is not a pass" }
            if ("$verdict" -notmatch ', 0 failed') {
                # The failing checks by name: at the end of a run the one line that matters has
                # usually scrolled away.
                $broken = @($answer | Select-String '^\s*FAIL' | ForEach-Object { $_.Line.Trim() })
                $named = if ($broken.Count -gt 0) { ': ' + ($broken -join '; ') } else { '' }
                throw "$suite did not pass$named"
            }

            $ran += "$suite $("$verdict".Trim())"
        }
        return $ran
    }

    Step 'live probes' {
        $excel = Use-Fixture 'DebugFixture.xlsm'

        # A session that has only just launched is still seeding: the engine is starting, the
        # first analysis pass has not run, and the ghost readers may not be attached. Probes
        # that assert a healthy session then fail on the truth that it is not healthy YET  -
        # which is a real answer to the wrong question (2026-08-06, on a release gate).
        Import-Module (Join-Path $repoRoot 'tools\harness\XlideApi.psm1') -Force
        $found = $null
        try { $found = Get-XlideApi -ProcessId $excel.Id -TimeoutSeconds 10 } catch { }
        if ($found) {
            $door = $found.Base
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

        # The tail of this list is ORDERED, not alphabetical. objbrowser-live-probe asserts no
        # palette exists before its summons and leaves one open behind it, so it runs after
        # everything that would mind and before Test-ResizeFollow - which runs LAST because it
        # closes every pane and leaves the frame resized, and the next step's fresh relaunch is
        # what puts that back. ResizeFollow is the only thing anywhere that RESIZES the host
        # window, so the twice-shipped placement chain (frame -> overlay -> Chromium child) is
        # exercised by it alone; the palette probe is the only live coverage the Object Browser
        # has - its summons, real-data panes, member navigation, and the native browser staying
        # retired (both triaged 2026-08-12).
        foreach ($probe in 'Test-DebugApi.ps1', 'Test-SplitWorkspace.ps1', 'Test-DiscardProblems.ps1', 'Test-Churn.ps1', 'objbrowser-live-probe.mjs', 'Test-ResizeFollow.ps1') {
            $answer = if ($probe.EndsWith('.mjs')) {
                node (Join-Path $repoRoot "tools\harness\$probe") 2>&1
            } else {
                powershell -NoProfile -ExecutionPolicy Bypass `
                    -File (Join-Path $repoRoot "tools\harness\$probe") 2>&1
            }
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
        'xlide api, split workspace, churn'
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
        # Grouped by the fixture each needs, so the gate opens a handful rather than failing
        # nineteen. AN ARRAY OF GROUPS, not a map keyed by fixture: two groups may open the
        # SAME fixture on purpose - write-rollback's own launch below is the case - and a key
        # cannot say that twice.
        $plan = @(
            # module-sync writes into a temporary folder of its own and takes back every module it
            # adds. Its last section is the one worth having: it applies the same import through
            # the DIALOG and through the route and compares the result byte for byte, which is the
            # only check that would notice the api growing its own idea of what an import means.
            # module-sync runs TWICE, once per planner. The two are meant to decide identically,
            # and the only way to know they still do is to put the same 32 checks through both.
            # It is also the one thing that catches a silent fall back to the built-in planner,
            # which is how a green run against the wrong implementation happened (2026-08-09).
            # import-guard runs FIRST on this fixture, before the sync suites open and close
            # anything: it needs a module on the surface with edits pending, and that precondition
            # stops holding once the surface has been worked over (2026-08-09).
            # THE DEBUGGER SUITES RUN HERE, and until 2026-08-11 they ran nowhere. Both need a
            # project that COMPILES, which is this fixture and no other: debugger-features drives
            # run, break, reset, stepping and breakpoints, and step-into-features covers the one
            # direction nothing else does - the HOST moving on its own and the surface following,
            # which is what a developer meets when Step Into crosses into a module the page was
            # not showing. write-rollback used to run LAST in this group - the only thing that
            # drives the `mark` route, and a refused write that half-lands is the defect that
            # costs a developer their module's previous text. Last-in-group protected the
            # suites after it from the session it spends, and could not protect the probe
            # itself from the eighteen suites before it: their adds and removes wear the same
            # identifier table the refused-body write is about to punch, and on 2026-08-31
            # that composition took VBE7 down a recursion instead of its clean refusal - a
            # 0xc00000fd stack overflow in oleaut32, two sessions in one gate, where the same
            # write against a fresh session answers "Out of memory" and rolls back (#19). It
            # has a one-suite group of its own directly after this one.
            # window-routes drives the three 2026-08-12 routes (palette hide, frame close and
            # show, pane closeNative) and their follow contracts - the palette going down with
            # the frame and staying away on its return, the hidden pane's tab leaving the
            # strip.
            # designer-features builds a whole UserForm through the designer routes, verifies
            # the read against the plan that built it, round-trips the mutations, and removes
            # the component.
            # write-fidelity is here because it needs nothing of the fixture but a project to
            # add a module to. It pins one property the rest of this list cannot see: that a
            # whole-module write puts back exactly the text it was given. The editor appends a
            # line reading `()` to any module holding a `Declare` broken over a continuation
            # when it is filled with AddFromString, which cost the owner a module that would
            # not compile, pasted from a file that was fine (2026-08-21).
            # inline-comments-live and analysis-rules-live both ADD components (the carrier
            # modules their checks write on), so nothing in this group may spend the session's
            # identifier budget. For one day they sat wired after write-rollback, where a full
            # -Live pass fed them a session it had already spent: both died on "no xlide
            # instance is answering" and the gate read it as their failure (2026-08-28, found
            # the first time the full pass ran after they landed).
            @{ Fixture = 'DebugFixture.xlsm'
               Suites  = @('import-guard.mjs', 'immediate-watch.mjs',
                                     'analysis-freshness.mjs', 'menu-bar.mjs',
                                     'write-fidelity.mjs',
                                     'module-sync.mjs xlide', 'module-sync.mjs builtIn',
                                     'debugger-features.mjs', 'step-into-features.mjs',
                                     'properties-pane.mjs', 'window-routes.mjs',
                                     'designer-features.mjs', 'test-runner.mjs',
                                     'pane-scope.mjs',
                                     'inline-comments-live.mjs', 'analysis-rules-live.mjs',
                                     'unsaved-workbook.mjs',
                                     # The api's own switch, read from outside the page: the card
                                     # says whether the door is open, hands out THIS session's
                                     # address, and says what turning it on costs. It presses
                                     # copy and close and never the toggle, which is what makes
                                     # it safe here - turning the door off would sever the
                                     # connection every suite after it drives over, and that half
                                     # is Test-ApiSwitch.ps1's, by hand. Written with care and
                                     # run by nothing until 2026-08-29.
                                     'agent-card.mjs') }
            # write-rollback ALONE, on a launch of its own: the fresh session is the suite's
            # PRECONDITION now, not a courtesy. The refused-body write it provokes meets a clean
            # identifier table - the state it is proven green against in isolation - instead of
            # one eighteen suites have worn, which is where VBE7 recursed to a stack overflow
            # rather than refusing (2026-08-31, #19). One extra launch, about fifteen seconds.
            # The relaunch AFTER it stays what it always was: the restart the suite's own
            # header demands, because the session it leaves cannot add a component.
            @{ Fixture = 'DebugFixture.xlsm'; Suites = @('write-rollback.mjs') }
            # colouring runs here because it declares its own module and needs nothing of the
            # fixture. It pins the one visible feature that had no check at all: a tokenizer
            # rebuilt per project, whose two defects on 2026-08-09 were both found by eye.
            # rename-features runs here for the fixture it was built against, and it is NOT
            # covered by three-copies.mjs beside it, which was the reason given for leaving it
            # out. That suite asks one question - do the workbook, the surface and the analyzer
            # hold the same text - across eight operations, one of which happens to be a rename.
            # A rename that wrongly rewrote Rival's own Recalculate would leave all three copies
            # in perfect agreement and pass. This is the suite that asks whether the rename
            # touched the right things and left the wrong ones alone, which is the whole of what
            # rename has to get right.
            #
            # THE TWIN IS OPEN ON PURPOSE. This whole group runs with TwinFixture beside the
            # rename fixture, because two workbooks holding same-named modules is the state
            # three defect classes have lived in and no gate session ever had - every one of
            # those defects was found by hand. Opening both costs zero extra launches, all six
            # suites were proven green in the double session before the widening (2026-08-12),
            # and rename-boundary.mjs runs last: the one question only two workbooks can ask,
            # whether a rename crosses modules and STOPS at the workbook, byte for byte,
            # through the rename and its undo.
            #
            # sync-scope joins them for the same reason rename-boundary is here: it asks the one
            # question only two workbooks can ask. Import and export used to name no project, and
            # the host filled the gap with the SHOWN project for the plan's identity and the
            # ACTIVE one for its modules - routinely different with two open, so a plan titled one
            # workbook listed the other's modules and applying it would have written the folder
            # over the wrong code (2026-08-21).
            @{ Fixture = 'RenameFixture.xlsm + TwinFixture.xlsm'
               Suites  = @(
                                     'format-positions.mjs', 'three-copies.mjs', 'colouring.mjs',
                                     'settings-bite.mjs', 'rename-features.mjs', 'sync-scope.mjs',
                                     'search-features.mjs', 'rename-boundary.mjs') }
            # TWO FILES OF TESTS, which is the only state the runner's file dimension exists in.
            # The runner reads every open project, XlideAssert lives per file, and both fixtures
            # hold a module called InvoiceTests on purpose - so a result filed by module name
            # alone lands on the wrong file's test, and a gate that read support once for the
            # session refuses the file that has it. None of that is askable with one file open,
            # which is why the pair exists (tools\New-TestTwinFixture.ps1).
            #
            # AND A THIRD FILE WITH NOTHING TO SAY, which costs no extra launch and buys the
            # other half of the file dimension: a project holding neither tests nor XlideAssert.
            # The pane's cache had two writers with two rules about such a file, and all three
            # symptoms - no install offered for the file being worked in, the file vanishing
            # mid-typing, and a full COM walk of the session on every tree publish - were
            # invisible while every open file held tests (2026-08-21).
            #
            # tests-support joins them because it needs files that DISAGREE about XlideAssert, and
            # this is the only session that has them by construction: two test fixtures carrying
            # it and one file that does not. The install chip used to read green over a file that
            # plainly had none - the pane listed a file only while it held tests or was being
            # worked in, nothing republished it when the developer moved, and the session's answer
            # counted only files holding tests, which with none open is vacuously satisfied
            # (2026-08-21).
            @{ Fixture = 'TestFixture.xlsm + TestTwinFixture.xlsm + DebugFixture.xlsm'
               Suites  = @('multi-file.mjs', 'tests-support.mjs') }
            # THE CHANGE LOG GETS ITS OWN WORKBOOK (tools\New-ChangeFixture.ps1), because its
            # suite edits modules and then asserts on what the log says was edited. Any fixture
            # another suite reads would be either useless for that or a landmine for the other
            # suite - DebugFixture alone is read by nine. Its Untouched module is edited by
            # nobody, which is how "the log did not over-report" is a real question.
            @{ Fixture = 'ChangeFixture.xlsm'; Suites = @('change-log.mjs') }
        )

        foreach ($group in $plan) {
          $ran += Invoke-SuiteGroup $group.Fixture $group.Suites
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

    # THE LAST STEP OF THE LIVE HALF, by construction: what this guards is Excel DYING when
    # the developer closes the editor window (lesson 27, shipped once), so a regression takes
    # the session with it and nothing may SHARE a session behind it. Three SC_CLOSE and reopen
    # cycles; ~10 seconds on the session the sweep above leaves open. The -Deep phase after
    # this is safe because every deep group relaunches fresh.
    Step 'closing the editor leaves Excel standing' {
        $excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $excel) { throw 'no Excel to close the editor of; the live half should have left one open' }

        $answer = powershell -NoProfile -ExecutionPolicy Bypass `
            -File (Join-Path $repoRoot 'tools\harness\Test-CloseVbe.ps1') 2>&1
        $answer | Out-Host
        $verdict = $answer | Select-String 'RESULT: (PASS|FAIL)' | Select-Object -Last 1
        if ("$verdict" -notmatch 'PASS') { throw 'Test-CloseVbe.ps1 did not pass' }
        'three close and reopen cycles, Excel standing'
    }
}

if ($Deep) {
    <#
        THE PRE-RELEASE TIER. Everything here earns depth rather than speed, which is why it is
        not in -Live: a third fixture launch for the completion questions, half a minute of
        deliberate settle sleeps in the non-ASCII round trip, and a randomized walk whose whole
        value is the action orders nobody scripts. Run it before a release; a commit does not
        owe it (triaged 2026-08-12, one agent per suite, cost grounded line by line).

        Three triaged suites are NOT here although they are unique cover - Test-ObjectBrowser's
        window-lifecycle trio, Test-GhostLocalsPanel's setLocals push, Test-WatchPanel's
        populated Watch panel - because all three reach through the VBA project object model,
        and this machine runs with that trust OFF, as does everything api-driven by design.
        Their path into a tier is routes: frame visibility, and something that can populate a
        watch. Until then they are hand-run documentation of what the api cannot yet do.
    #>
    Step 'deep: completion menus, signature help, and the shutdown revival' {
        # The dot menu is the feature a developer touches most, and nothing in -Live asks it
        # anything: its receivers (a project class, a UDT, an enum, the host's type libraries)
        # exist only in LanguageFixture. Tolerates the two upstream analyzer defects filed as
        # xlide_vscode#11, and announces the day they start passing.
        #
        # session-lifecycle runs LAST and alone in its session for a reason of its own: it
        # drives a cancelled shutdown - OnBeginShutdown without a process exit - and proves the
        # watchdog revives the session, which is the 2026-08-02 field failure (the add-in came
        # back dead inside a living Excel) and had no test because reaching it meant closing
        # Excel by hand and pressing Cancel. It costs ~30s: a teardown, two 1500ms watchdog
        # ticks, and the revived session reseeding. That is depth, not speed, so it earns the
        # -Deep tier rather than the per-commit gate. It leaves a healthy session behind, but
        # a fresh one on a new port, so nothing may share its session after it - last is right.
        (Invoke-SuiteGroup 'LanguageFixture.xlsm' @('language-features.mjs', 'session-lifecycle.mjs')) -join '; '
    }

    Step 'deep: the reconstruction path, both sides of the old 64k' {
        # Every keystroke's shadow rebuild, held at a small size and a large one - and the
        # pin that the whole-text copy stays dropped: it hid this path below 64,000
        # characters for the product's whole life, which is why nothing had tested it
        # (the audit's C14). Its own session: the suite seeds two ~60KB modules and churns
        # them.
        (Invoke-SuiteGroup 'DebugFixture.xlsm' @('reconstruct-branch.mjs')) -join '; '
    }

    Step 'deep: a workbook Excel will not run' {
        # THE STATE ISSUE #9 IS MADE OF. Excel puts itself into design mode for a workbook whose
        # macros are disabled and will not come out: every Reset control greys, the Design Mode
        # toggle stays pressed however often it is pressed, and nothing the product offers clears
        # it - the way out is closing the workbook and opening it again with macros enabled.
        #
        # Its own Excel, built and torn down by the script, because it has to open a workbook a
        # particular way and no shared session can be in that state. It builds its probe through
        # the VBA project object model, so it SKIPS out loud when that trust is off rather than
        # failing - nothing this product ships requires the switch.
        $answer = powershell -NoProfile -ExecutionPolicy Bypass `
            -File (Join-Path $repoRoot 'tools\harness\Test-MacrosDisabled.ps1') 2>&1
        $answer | Where-Object { $_ -match 'FAIL|SKIPPED|passed' } | ForEach-Object { Write-Host "  $_" }
        if ($answer -match 'SKIPPED') { return 'skipped: VBA project object model trust is off' }
        if (($answer | Select-String '^PASS' | Select-Object -Last 1) -eq $null) {
            throw 'Test-MacrosDisabled.ps1 did not pass'
        }
        'design mode Excel will not leave, and the product names why'
    }

    Step 'deep: non-ASCII round trip, then the randomized walk' {
        # language-live-probe writes 14 scripts through COM, the door, the page and the outline
        # (23s of settles by construction); surface-walk runs last because it churns the whole
        # workspace and ends with resetLayout - the two-workbook session it needs is the same
        # one, opened once. The walk fails on vacuity now: a run that never held both
        # workbooks' colliding tabs is a run that proved nothing.
        (Invoke-SuiteGroup 'RenameFixture.xlsm + TwinFixture.xlsm' @(
            'language-live-probe.mjs', 'surface-walk.mjs --steps 80')) -join '; '
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
