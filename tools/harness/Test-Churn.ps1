# Standing PASS/FAIL probe for LEAKS in the workspace's churn.
#
# The surface creates and destroys real things now: an editor per group, made on every split
# and disposed when a group dissolves; a Monaco model per open module, made when a pane opens
# and disposed when it closes; a listener per model; DOM for every dock group. Before the
# workspace could rearrange, none of that moved. A dispose that misses one of them leaks
# quietly, and the symptom arrives days later as a session that has become slow - by which
# time the cause is unrecoverable.
#
# So this does the churn on purpose, many times, and asserts that the counts come back to
# where they started. Counts, not memory: a model count is exact and a heap is not, and an
# exact number that must return to its starting value is a far better leak detector than a
# megabyte figure nobody can interpret. Memory and handles are REPORTED for context, and
# only fail on growth too large to be noise.
#
# Run tools\dev.ps1 -KeepOpen first; Debug builds only.
$ErrorActionPreference = 'Continue'

$checks = [ordered] @{}
$script:notes = @()

function Check([string] $name, [scriptblock] $test) {
    try {
        $result = & $test
        $checks[$name] = if ($result) { 'PASS' } else { 'FAIL' }
    } catch {
        $checks[$name] = "FAIL ($($_.Exception.Message))"
    }
}

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excel) { Write-Output 'RESULT: FAIL - no Excel is running'; return }

$discovery = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
if (-not (Test-Path $discovery)) { Write-Output "RESULT: FAIL - no discovery file for Excel $($excel.Id)"; return }

$d = Get-Content $discovery -Raw | ConvertFrom-Json
$api = "http://127.0.0.1:$($d.port)/$($d.token)"
Write-Output "api: $api (pid $($d.pid))"

function Page([string] $script) {
    $r = Invoke-RestMethod "$api/eval" -Method Post -Body $script -TimeoutSec 30
    if (-not $r.answered) { throw 'the page did not answer' }
    $out = $r.result
    if ($out -is [string] -and $out.StartsWith('"')) { $out = $out | ConvertFrom-Json }
    $out
}

# Everything the churn could leak, counted in one place. The page's own numbers, because the
# page is where these things live.
function Get-Counts {
    $raw = Page @'
(() => {
  // Monaco is bundled, not global, so the page has to be asked for its own census.
  const census = window.xlideBridge.modelCensus();
  return JSON.stringify({
    models: census.models,
    documents: census.documents,
    editors: window.xlideBridge.workspace.editors().length,
    editorGroups: document.querySelectorAll(".editor-group").length,
    dockGroups: document.querySelectorAll(".dock-group").length,
    domNodes: document.querySelectorAll("*").length,
    overlays: document.querySelectorAll(".drop-overlay, .drag-dim, .drop-compass").length,
    monacoNodes: document.querySelectorAll(".monaco-editor").length
  });
})()
'@
    $raw | ConvertFrom-Json
}

# GDI and USER objects, which is what "a window, a timer, or a device context left behind" means.
# Kernel handle count cannot answer that question here: this probe makes about a hundred HTTP
# requests and the door's sockets swamp the signal.
Add-Type -Namespace Xlide -Name Gui -MemberDefinition @'
[DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr process, uint flags);
'@ -ErrorAction SilentlyContinue

function Get-HostStats {
    $s = Invoke-RestMethod "$api/stats" -TimeoutSec 10
    $handle = (Get-Process -Id $excel.Id).Handle
    [pscustomobject] @{
        ManagedMb = [Math]::Round($s.managedMemoryBytes / 1MB, 1)
        WorkingMb = [Math]::Round($s.workingSetBytes / 1MB, 1)
        Handles = $s.handleCount
        Gdi = [Xlide.Gui]::GetGuiResources($handle, 0)
        User = [Xlide.Gui]::GetGuiResources($handle, 1)
    }
}

# A clean starting point, so the first cycle is not measuring the arrangement settling.
Invoke-RestMethod "$api/layout?reset=1" -Method Post -TimeoutSec 45 | Out-Null
Start-Sleep -Milliseconds 800

$cycles = 12
$before = Get-Counts
$hostBefore = Get-HostStats
Write-Output ("start: {0} models, {1} editors, {2} editor groups, {3} dock groups, {4} DOM nodes" -f
    $before.models, $before.editors, $before.editorGroups, $before.dockGroups, $before.domNodes)

Check "splitting and dissolving $cycles times leaves no editor behind" {
    for ($i = 0; $i -lt $cycles; $i++) {
        # Split, then send the tab back, which dissolves the group it just made.
        Page @'
(() => {
  const ws = window.xlideBridge.workspace;
  ws.splitActive("right");
  return ws.groupCount();
})()
'@ | Out-Null
        Start-Sleep -Milliseconds 250

        Page @'
(() => {
  const ws = window.xlideBridge.workspace;
  const groups = document.querySelectorAll(".editor-group");
  if (groups.length < 2) { return "not split"; }
  // Drag the second group's tab onto the first group's strip, which empties and dissolves it.
  const fire = (el, type, x, y) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, clientX: x, clientY: y, button: 0, buttons: type === "pointerup" ? 0 : 1, bubbles: true, cancelable: true }));
  const tab = groups[1].querySelector(".tab");
  const target = groups[0].querySelector(".group-tabs");
  const from = tab.getBoundingClientRect(), to = target.getBoundingClientRect();
  fire(tab, "pointerdown", from.x + 10, from.y + 8);
  fire(window, "pointermove", from.x + 40, from.y + 8);
  fire(window, "pointermove", to.x + to.width - 6, to.y + to.height / 2);
  fire(window, "pointerup", to.x + to.width - 6, to.y + to.height / 2);
  return window.xlideBridge.workspace.groupCount();
})()
'@ | Out-Null
        Start-Sleep -Milliseconds 250
    }

    $after = Get-Counts
    $script:notes += ("after splits: {0} editors (was {1}), {2} monaco nodes (was {3})" -f
        $after.editors, $before.editors, $after.monacoNodes, $before.monacoNodes)

    # Exact: every editor made by a split must be gone when its group dissolved, and the
    # Monaco DOM with it.
    $after.editors -eq $before.editors -and $after.monacoNodes -le $before.monacoNodes
}

Check "docking and undocking $cycles times leaves no pane group behind" {
    for ($i = 0; $i -lt $cycles; $i++) {
        Page @'
(() => {
  const fire = (el, type, x, y) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, clientX: x, clientY: y, button: 0, buttons: type === "pointerup" ? 0 : 1, bubbles: true, cancelable: true }));
  const dragToPetal = (pane, regionEl, zone) => {
    const tab = [...document.querySelectorAll(".panel-tab")].find(t => t.dataset.panel === pane);
    if (!tab || !regionEl) { return "no-tab"; }
    const from = tab.getBoundingClientRect(), box = regionEl.getBoundingClientRect();
    fire(tab, "pointerdown", from.x + 10, from.y + 8);
    fire(window, "pointermove", from.x + 40, from.y + 8);
    fire(window, "pointermove", box.x + box.width / 2, box.y + box.height / 2);
    const petal = document.querySelector(".drop-petal-" + zone);
    if (!petal || petal.hidden) { fire(window, "pointerup", box.x, box.y); return "no-petal"; }
    const p = petal.getBoundingClientRect();
    fire(window, "pointermove", p.x + p.width / 2, p.y + p.height / 2);
    fire(window, "pointerup", p.x + p.width / 2, p.y + p.height / 2);
    return "dropped";
  };

  // Out to a section of its own on the right, then back into the bottom group's tabs.
  dragToPetal("watch", document.getElementById("editor-area"), "right");
  const bottom = document.querySelector('#dock-bottom .dock-body');
  dragToPetal("watch", bottom, "center");
  return "cycled";
})()
'@ | Out-Null
        Start-Sleep -Milliseconds 300
    }

    $after = Get-Counts
    $script:notes += ("after docking: {0} dock groups (was {1}), {2} DOM nodes (was {3})" -f
        $after.dockGroups, $before.dockGroups, $after.domNodes, $before.domNodes)

    # The right section must be gone again, so the group count returns; the DOM is allowed a
    # little slack for text nodes and the like, but not unbounded growth.
    $after.dockGroups -eq $before.dockGroups -and ($after.domNodes - $before.domNodes) -lt 200
}

Check 'the drag furniture is never left on the page' {
    # A dim and a compass that outlive their drag are both a leak and a surface that looks
    # permanently mid-gesture (2026-08-06).
    (Get-Counts).overlays -eq 0
}

Check 'models and documents match the modules actually open' {
    $after = Get-Counts
    $script:notes += ("models {0}, documents {1}" -f $after.models, $after.documents)

    # One model per open document, and no document alive without a tab to show it.
    #
    # This used to also demand that the count be UNCHANGED, which is not a property the churn
    # has. Models are made lazily, when a document is first SHOWN: four tabs sit at one model
    # until something displays them. The split in this very probe shows a second module, so the
    # count legitimately rises, and the check passed only when the starting state happened to
    # have visited a second document already. It failed on a clean session and passed on a dirty
    # one, which is exactly backwards (2026-08-08).
    #
    # What a leak would look like is a model with no document, or a document with no tab. Both
    # are checked; growth within the open tabs is not a leak, it is laziness working.
    $tabs = [int] (Page '(() => document.querySelectorAll(".tab").length)()')
    $script:notes += ("tabs open: {0}" -f $tabs)

    $after.models -eq $after.documents -and $after.documents -le $tabs
}

Check 'the host did not grow unreasonably through the churn' {
    $hostAfter = Get-HostStats
    $script:notes += ("host: managed {0} -> {1} MB, working {2} -> {3} MB, handles {4} -> {5}, gdi {6} -> {7}, user {8} -> {9}" -f
        $hostBefore.ManagedMb, $hostAfter.ManagedMb,
        $hostBefore.WorkingMb, $hostAfter.WorkingMb,
        $hostBefore.Handles, $hostAfter.Handles,
        $hostBefore.Gdi, $hostAfter.Gdi,
        $hostBefore.User, $hostAfter.User)

    # WINDOWS AND DRAWING OBJECTS, not kernel handles.
    #
    # What this check is for is the unmistakable case: a window, a timer, or a device context left
    # behind by every cycle. It used to ask the kernel handle count, and its own comment admitted
    # the problem, that the door's HTTP sockets are the largest contributor across a run. Loosening
    # the threshold did not fix that, it only postponed it: the check passed when run alone and
    # failed inside the gate, where it runs after forty-five other probes have left sockets waiting
    # to close. It was measuring the harness (2026-08-09).
    #
    # GDI and USER counts are exactly the objects named above and no socket touches them, so the
    # threshold can be tight enough to mean something: a couple per cycle is noise, ten each is a
    # cycle that leaks one.
    #
    # Managed memory and handles are still REPORTED. A heap that has not collected is not a leak,
    # and a number worth seeing is not always a number worth failing on.
    $gdiGrowth = $hostAfter.Gdi - $hostBefore.Gdi
    $userGrowth = $hostAfter.User - $hostBefore.User
    $script:notes += ("gdi grew {0}, user grew {1}, over {2} cycles" -f $gdiGrowth, $userGrowth, $cycles)

    $gdiGrowth -lt ($cycles * 2) -and $userGrowth -lt ($cycles * 2)
}

# Leave the arrangement the way it was found.
Invoke-RestMethod "$api/layout?reset=1" -Method Post -TimeoutSec 45 | Out-Null

Write-Output ''
foreach ($entry in $checks.GetEnumerator()) { "  {0,-58} {1}" -f $entry.Key, $entry.Value }
Write-Output ''
foreach ($note in $script:notes) { "  $note" }

$failed = @($checks.Values | Where-Object { $_ -ne 'PASS' })
Write-Output ''
Write-Output ("RESULT: " + $(if ($failed.Count -eq 0) { "PASS ($($checks.Count) checks)" } else { "FAIL ($($failed.Count) of $($checks.Count))" }))
