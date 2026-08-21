<#
.SYNOPSIS
    Rebuild the editor page and put it into the RUNNING editor, without restarting anything.

.DESCRIPTION
    The surface is a web page. Changing it needs none of what changing the shim needs - no
    publish, no re-registration, no Excel restart - but the loop was being run by hand:
    build, copy the bundle into the published shim's directory, reload the page, wait, look.
    A dozen times in an afternoon (2026-08-06). This is that loop, once, with the waiting
    done properly through the debug api instead of by sleeping a guess.

    The shim keeps its DLL locked while Excel runs, but the page bundle beside it is just
    files: they can be replaced under a live session, and a reload picks them up.

.EXAMPLE
    tools\page.ps1
    Build, deploy, reload. The usual loop.

.EXAMPLE
    tools\page.ps1 -Reset
    Also put the pane arrangement back to the default, for when a probe left it rearranged.

.EXAMPLE
    tools\page.ps1 -Watch
    Stay running, and do it again whenever a source file changes.

.EXAMPLE
    tools\page.ps1 -NoDeploy
    Build and typecheck only. For when no editor is open.
#>
[CmdletBinding()]
param(
    # Skip the deploy and reload; just build the bundle.
    [switch] $NoDeploy,

    # Reset the pane arrangement to the default as part of the reload.
    [switch] $Reset,

    # Rebuild and redeploy whenever a page source file changes, until Ctrl+C.
    [switch] $Watch,

    # Skip the typecheck, for a fast turn when the types were checked a moment ago.
    [switch] $NoTypecheck,

    # Which published shim to deploy into.
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Debug'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$pageRoot = Join-Path $repoRoot 'ui\editor'
$distRoot = Join-Path $pageRoot 'dist'
$publishRoot = Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-x64\ui\editor\dist"

function Write-Step([string] $text) {
    Write-Host ''
    Write-Host "==> $text" -ForegroundColor Cyan
}

# EVERY live session, because the bundle is shared. The deploy writes one publish directory
# and every session serves its page from it, so reloading one session and not another leaves
# the other running the old page with no sign anywhere - and the old form of this function
# refused outright with two sessions live, which threw AFTER the files had already been
# deployed (hit 2026-08-19, the day an Excel and a Word first ran side by side). Find-XlideApi
# is the module's discover(): every live session, proven by /state, never a guess.
function Get-LiveApis {
    Import-Module (Join-Path $PSScriptRoot 'harness\XlideApi.psm1') -Force
    return @(Find-XlideApi | ForEach-Object {
        [pscustomobject] @{ Pid = $_.Pid; Api = $_.Base }
    })
}

function Invoke-PageBuild {
    if (-not $NoTypecheck) {
        Write-Step 'Typecheck'
        Push-Location $pageRoot
        try {
            npm run typecheck
            if ($LASTEXITCODE -ne 0) { throw 'The page does not typecheck.' }
        } finally { Pop-Location }
    }

    Write-Step 'Build the page'
    Push-Location $pageRoot
    try {
        npm run build
        if ($LASTEXITCODE -ne 0) { throw 'The page bundle would not build.' }
    } finally { Pop-Location }
}

function Invoke-PageDeploy {
    if ($NoDeploy) { return }

    if (-not (Test-Path $publishRoot)) {
        Write-Host "No published shim at $publishRoot; nothing to deploy into." -ForegroundColor Yellow
        return
    }

    Write-Step 'Deploy into the published shim'

    # COPIED, THEN CHECKED, because the copy can fail without failing.
    #
    # `Copy-Item -Force` writes a NON-TERMINATING error when a file is locked and carries on, so
    # a running host - whose WebView holds editor.js open - left the published bundle at its
    # previous build while this printed "Deployed" and the reload below then re-ran the OLD page.
    # Nothing downstream could see it either: the staleness guard compares the running page
    # against the bundle IN THE PUBLISH, and both were equally old, so they agreed. It cost three
    # wrong conclusions in one session, including a check that looked like it could not catch the
    # regression it had been written for (2026-08-21).
    #
    # So the destination is compared against the source afterwards, by size and time, and a file
    # that did not land is named along with whatever is holding it.
    Copy-Item (Join-Path $distRoot '*') $publishRoot -Recurse -Force -ErrorAction SilentlyContinue

    $missed = @()
    foreach ($built in Get-ChildItem $distRoot -Recurse -File) {
        $relative = $built.FullName.Substring($distRoot.Length).TrimStart('\')
        $landed = Join-Path $publishRoot $relative
        $there = Get-Item $landed -ErrorAction SilentlyContinue
        if (-not $there -or $there.Length -ne $built.Length -or $there.LastWriteTimeUtc -lt $built.LastWriteTimeUtc) {
            $missed += $relative
        }
    }

    if ($missed.Count -gt 0) {
        $holding = @(Get-Process EXCEL, WINWORD, POWERPNT -ErrorAction SilentlyContinue |
            ForEach-Object { "$($_.ProcessName) $($_.Id)" })
        $blame = if ($holding.Count -gt 0) {
            " A host is running and holds the page open: $($holding -join ', '). Close it and run this again."
        } else {
            ''
        }

        throw "The page did NOT reach the publish: $($missed -join ', ') did not land.$blame"
    }

    Write-Host "Deployed to $publishRoot"

    # Wrapped at the CALL: PowerShell unwraps a returned array, so zero sessions came back as
    # $null and one came back bare, and .Count on either is a strict-mode stop.
    $sessions = @(Get-LiveApis)
    if ($sessions.Count -eq 0) {
        Write-Host 'No editor is open, so nothing to reload. Start one with tools\dev.ps1 -KeepOpen.' -ForegroundColor Yellow
        return
    }

    foreach ($live in $sessions) {
        Write-Step "Reload the live page (pid $($live.Pid))"

        if ($Reset) {
            # Reset reloads on its own, so this is the reload as well.
            $reset = Invoke-RestMethod "$($live.Api)/layout?reset=1" -Method Post -TimeoutSec 45
            if (-not $reset.ran) { throw "The page did not come back after the layout reset (pid $($live.Pid))." }
            Write-Host 'Arrangement reset to the default.'
        } else {
            $reload = Invoke-RestMethod "$($live.Api)/reload" -Method Post -TimeoutSec 45
            if (-not $reload.ready) { throw "The page did not come back within $($reload.elapsedMs)ms (pid $($live.Pid))." }

            # The stamp is the point: it proves the running page IS the build that just happened,
            # which is the question three rounds of confusion were once spent on.
            Write-Host ("Back in {0}ms running {1}" -f $reload.elapsedMs, $reload.pageBuildStamp)
            if ($reload.stale) {
                throw "The page came back STALE (pid $($live.Pid)): it is running $($reload.pageBuildStamp) while the bundle on disk is $($reload.bundleBuiltUtc)."
            }
        }

        # Anything the page complained about on the way up, which the shim log does not carry.
        try {
            $console = Invoke-RestMethod "$($live.Api)/console?last=40" -TimeoutSec 8
            $noisy = @($console.lines | Where-Object { $_ -match '^(warn|error):' })
            if ($noisy.Count -gt 0) {
                Write-Host ''
                Write-Host "The page said (pid $($live.Pid)):" -ForegroundColor Yellow
                $noisy | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
            }
        } catch {
            # An older shim without the console route; not worth failing a deploy over.
        }
    }
}

Invoke-PageBuild
Invoke-PageDeploy

if (-not $Watch) {
    Write-Host ''
    $ending = if ($NoDeploy) { 'built.' } else { 'built and live.' }
    Write-Host "RESULT: the page is $ending" -ForegroundColor Green
    return
}

Write-Host ''
Write-Host 'Watching for changes. Ctrl+C to stop.' -ForegroundColor Green

$watcher = New-Object IO.FileSystemWatcher (Join-Path $pageRoot 'src'), '*.*'
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true

try {
    while ($true) {
        # A save often lands as several events; take one and let the rest settle before building.
        $change = $watcher.WaitForChanged([IO.WatcherChangeTypes]::All, 1000)
        if ($change.TimedOut) { continue }

        Start-Sleep -Milliseconds 250
        while (-not $watcher.WaitForChanged([IO.WatcherChangeTypes]::All, 150).TimedOut) { }

        Write-Host ''
        Write-Host ("--- {0} changed" -f $change.Name) -ForegroundColor DarkGray
        try {
            Invoke-PageBuild
            Invoke-PageDeploy
        } catch {
            # A watch loop that dies on the first typo is a watch loop nobody uses.
            Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} finally {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
}
