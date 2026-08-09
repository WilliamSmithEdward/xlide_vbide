# Standing PASS/FAIL probe for the editor groups and the multi-document protocol (decision
# 12): two modules live at once, a split showing both, edits addressed to the module they
# changed, undo surviving a tab switch, the menu bar without its Window menu, and the panel
# docks surviving a reload. Run tools\dev.ps1 -KeepOpen first; Debug builds only.
#
# Waits POLL for the condition they need rather than sleeping a guess: the fixture opens its
# own modules while the probe runs, and a fixed sleep raced that boot traffic (2026-08-06).
$ErrorActionPreference = 'Continue'

$checks = [ordered] @{}
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

$discoveryPath = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
if (-not (Test-Path $discoveryPath)) {
    Write-Output "RESULT: FAIL - no discovery file for Excel $($excel.Id)"
    return
}

$d = Get-Content $discoveryPath | ConvertFrom-Json
$api = "http://127.0.0.1:$($d.port)/$($d.token)"
Write-Output "api: $api (pid $($d.pid))"

function Page([string] $script) {
    $r = Invoke-RestMethod "$api/eval" -Method Post -Body $script -TimeoutSec 15
    if (-not $r.answered) { throw 'the page did not answer' }

    # The eval route answers with the JSON encoding of the result: a page string arrives
    # wrapped in quotes. Decoded here so checks compare values, not encodings.
    $out = $r.result
    if ($out -is [string] -and $out.StartsWith('"')) { $out = $out | ConvertFrom-Json }
    $out
}

function WaitFor([string] $what, [scriptblock] $condition, [int] $seconds = 15) {
    $deadline = (Get-Date).AddSeconds($seconds)
    do {
        if (& $condition) { return $true }
        Start-Sleep -Milliseconds 400
    } while ((Get-Date) -lt $deadline)
    throw "timed out waiting for $what"
}

# Two documents, so there is something to split against: the fixture's CleanModule, and
# ThisWorkbook, which every workbook has. The activation is asked for and then WAITED for —
# the host opens panes and echoes tabs on its own schedule.
Check 'both modules open, ThisWorkbook active' {
    Page 'window.xlideBridge.activateModule("CleanModule")' | Out-Null
    WaitFor 'CleanModule tab' { (Page 'document.querySelectorAll(".tab[data-module=CleanModule]").length') -ge 1 } | Out-Null
    Page 'window.xlideBridge.activateModule("ThisWorkbook")' | Out-Null
    WaitFor 'ThisWorkbook active' {
        (Page 'document.querySelector(".tab.active") ? document.querySelector(".tab.active").dataset.module : ""') -eq 'ThisWorkbook'
    }
}

Check 'both documents hold live models' {
    [int](Page 'window.xlideBridge.documents.all().length') -ge 2
}

$script:splitStood = $false
Check 'split right makes two groups showing different modules' {
    Page 'window.xlideBridge.workspace.splitActive("right")' | Out-Null
    WaitFor 'two groups with distinct models' {
        $uris = (Page 'JSON.stringify(window.xlideBridge.workspace.editors().map(e => e.getModel() ? e.getModel().uri.toString().toLowerCase() : null))') | ConvertFrom-Json
        $uris.Count -eq 2 -and $uris[0] -and $uris[1] -and $uris[0] -ne $uris[1]
    } | Out-Null
    $script:splitStood = $true
    $true
}

Check 'an edit to a background document writes its own module' {
    if (-not $script:splitStood) { throw 'the split never stood' }
    # ThisWorkbook is host-active; CleanModule is a background document. The edit goes to
    # CleanModule's MODEL — the same channel every keystroke uses — and must land in
    # CleanModule, never in the active module.
    $r = Page @'
(() => {
  const docs = window.xlideBridge.documents;
  const clean = docs.all().find(d => d.module.toLowerCase() === "cleanmodule");
  if (!clean) return "no-doc";
  const model = docs.get(clean.module, clean.project);
  if (!model) return "no-model";
  model.pushEditOperations(null, [{ range: { startLineNumber: 1, startColumn: 1, endLineNumber: 1, endColumn: 1 }, text: "' split probe touched this\n" }], () => null);
  return "edited";
})()
'@
    if ($r -ne 'edited') { throw "edit not applied: $r" }
    WaitFor 'the write to land' {
        (Invoke-RestMethod "$api/module?name=CleanModule" -TimeoutSec 8).text -match 'split probe touched this'
    }
}

Check 'the active module was not the one written' {
    -not ((Invoke-RestMethod "$api/module?name=ThisWorkbook" -TimeoutSec 8).text -match 'split probe touched this')
}

Check 'undo reverts on the same live model' {
    if (-not $script:splitStood) { throw 'the split never stood' }
    $line = Page @'
(() => {
  const docs = window.xlideBridge.documents;
  const clean = docs.all().find(d => d.module.toLowerCase() === "cleanmodule");
  if (!clean) return "no-doc";
  const model = docs.get(clean.module, clean.project);
  if (!model) return "no-model";
  model.undo();
  return model.getLineContent(1);
})()
'@
    $line -notmatch 'split probe'
}

Check 'Backspace deletes in every group, not just one' {
    if (-not $script:splitStood) { throw 'the split never stood' }
    # A standalone editor's addCommand goes into a keybinding service SHARED by every editor,
    # and the when-clause is the only scoping there is. Two groups once meant two identical
    # Backspace rules; the later one won everywhere and deleted in ITS editor, so Backspace
    # looked dead in the group being typed in (2026-08-06). Each rule is scoped to its own
    # editor now, and this pins it: every group must delete its own text.
    # A freshly split group's editor gets its input element a beat after its model, so the
    # keystroke waits for one rather than racing it.
    WaitFor 'every group to have an input element' {
        (Page 'window.xlideBridge.workspace.editors().every(e => !!(e.getDomNode().querySelector(".native-edit-context") || e.getDomNode().querySelector("textarea")))') -eq $true
    } | Out-Null

    # Each editor deletes text this check inserted, so the assertion never depends on what the
    # fixture's first line happens to hold: at the start of an empty line Backspace is
    # correctly a no-op, and reading that as a fault chased a bug that was not there.
    $r = Page @'
(() => {
  const out = [];
  const editors = window.xlideBridge.workspace.editors();
  for (const ed of editors) {
    const model = ed.getModel();
    if (!model) { out.push("no-model"); continue; }
    const el = ed.getDomNode().querySelector(".native-edit-context") || ed.getDomNode().querySelector("textarea");
    if (!el) { out.push("no-input-el"); continue; }

    model.pushEditOperations(null, [{ range: { startLineNumber: 1, startColumn: 1, endLineNumber: 1, endColumn: 1 }, text: "zz" }], () => null);
    ed.focus();
    ed.setPosition({ lineNumber: 1, column: 3 });

    const others = editors.filter(other => other !== ed).map(other => other.getModel().getValue());
    el.dispatchEvent(new KeyboardEvent("keydown", { key: "Backspace", code: "Backspace", keyCode: 8, which: 8, bubbles: true, cancelable: true }));

    const mine = model.getLineContent(1).startsWith("z") && !model.getLineContent(1).startsWith("zz");
    const strayed = editors.filter(other => other !== ed).some((other, i) => other.getModel().getValue() !== others[i]);

    model.undo();
    model.undo();
    out.push(strayed ? "strayed" : mine ? "ok" : "dead");
  }
  return out.join(",");
})()
'@
    if ($r -match 'dead|strayed|no-') { throw "backspace: $r" }
    $r -match 'ok'
}

Check 'the menu bar has no Window menu' {
    $captions = Page 'JSON.stringify([...document.querySelectorAll("#menubar [role=menuitem], #menubar button")].map(b => b.textContent))'
    $captions -notmatch 'Window'
}

Check 'a pane dragged to the right section stays there across a reload' {
    Page @'
(() => {
  localStorage.setItem("xlide.docks.v1", JSON.stringify({
    sides: {
      left: { kind: "split", direction: "column", children: [
        { kind: "group", tabs: ["explorer"], active: "explorer" },
        { kind: "group", tabs: ["properties"], active: "properties" }], sizes: [0.6, 0.4] },
      right: { kind: "group", tabs: ["watch"], active: "watch" },
      top: null,
      bottom: { kind: "group", tabs: ["problems", "immediate", "locals"], active: "problems" }
    },
    // SIZED FROM THE WINDOW, not fixed.
    //
    // These were 260, 300 and 200 flat, which fits the window this was written against and not
    // the one the editor actually opens with. On a 640x409 host the two side sections claimed 560
    // of 640 and the bottom claimed 200 of 409, leaving the editor 70 wide and 128 tall - so
    // "the editor keeps the room the sections do not claim" was correctly reporting that it had
    // not, and the two compass checks below had no room to drop into. Three checks failing for a
    // window size, on a fresh session only, because a warmed one had been resized by hand
    // (2026-08-09).
    //
    // A quarter each way leaves the editor half the window at any size, which is what the checks
    // below are actually about.
    sizes: {
      left: Math.max(120, Math.round(innerWidth * 0.25)),
      right: Math.max(120, Math.round(innerWidth * 0.25)),
      top: 200,
      bottom: Math.max(80, Math.round(innerHeight * 0.25))
    },
    closed: []
  }));
  return "seeded";
})()
'@ | Out-Null
    Page 'location.reload()' | Out-Null
    WaitFor 'the page to come back with the right section standing' {
        try {
            $state = (Page 'JSON.stringify({ up: !!window.xlideBridge, dock: document.getElementById("dock-right") ? document.getElementById("dock-right").hidden : null, watch: document.getElementById("watch") ? document.getElementById("watch").hidden : null })') | ConvertFrom-Json
            $state.up -and $state.dock -eq $false -and $state.watch -eq $false
        } catch { $false }
    } 25

    # AND WAIT FOR THE SIZES, not just for the sections to exist.
    #
    # Standing and SIZED are different moments. The wait above is satisfied as soon as the right
    # section is on screen, which on a freshly started session is before the seeded sizes have been
    # applied — and every geometry check below then measures a layout that is still arriving. It
    # failed the three of them on a fresh session and passed on a warmed one, which read as a
    # product defect for an hour and was this (2026-08-09).
    #
    # Waits for the seeded bottom size specifically, so what follows can assert what it is actually
    # about: given the sections have the room they asked for, the editor keeps the rest.
    WaitFor 'the seeded sizes to be applied' {
        try {
            $applied = (Page 'JSON.stringify({ bottom: Math.round(document.getElementById("dock-bottom").getBoundingClientRect().height), right: Math.round(document.getElementById("dock-right").getBoundingClientRect().width) })') | ConvertFrom-Json
            $applied.bottom -ge 60 -and $applied.right -ge 100
        } catch { $false }
    } 20
}

Check 'the editor keeps the room the sections do not claim' {
    # A flex column sized the editor area to its content — nothing, for a Monaco container —
    # and the bottom section swallowed the workspace (2026-08-06).
    #
    # It was also RIGHT once, 2026-08-09, and the seeded layout above is why: fixed sizes of 260
    # and 300 in a 640-wide window left the editor 70 wide and 128 tall against a 200 bottom, so
    # this check reported exactly what had happened. The sizes are proportional now.
    #
    # Worth the telling because of how it looked: the three geometry checks failed together on
    # every fresh session and passed on a session that had been resized by hand, so it presented
    # as a flake and stayed hidden through a stash, a revert to the previous commit, and a layout
    # reset - all of which changed nothing, because none of them was the cause. What found it was
    # printing the numbers. Verdicts alone cannot tell a real defect from a fixture that does not
    # fit.
    $sizes = (Page 'JSON.stringify({ area: Math.round(document.getElementById("editor-area").getBoundingClientRect().height), bottom: Math.round(document.getElementById("dock-bottom").getBoundingClientRect().height) })') | ConvertFrom-Json
    $sizes.area -gt $sizes.bottom
}

Check 'a pane closes with its X and returns from the Panes menu' {
    $r = Page @'
(() => {
  const group = [...document.querySelectorAll(".dock-group")].find(g => g.querySelector('.panel-tab[data-panel="problems"].active'));
  if (!group) return "no-group";
  const close = group.querySelector(".dock-close");
  if (!close) return "no-close";
  close.click();
  return document.querySelector('.panel-tab[data-panel="problems"]') ? "still-there" : "closed";
})()
'@
    if ($r -ne 'closed') { throw "close failed: $r" }

    $back = Page @'
(() => {
  const menu = () => document.getElementById("panes-menu");
  if (menu()) menu().remove();
  document.querySelector('#toolbar [data-command="openPanes"]').click();
  const item = [...document.querySelectorAll("#panes-menu .panes-menu-item")].find(i => i.dataset.pane === "problems");
  if (!item) return "no-item";
  item.click();
  if (menu()) menu().remove();
  return document.querySelector('.panel-tab[data-panel="problems"]') ? "back" : "still-closed";
})()
'@
    $back -eq 'back'
}

Check 'a drop on a compass zone lands in that zone, and only there' {
    # The compass IS the target: the pointer must come to the zone it means. Guessing an
    # intent from where the pointer happens to be over a wide short region could not tell
    # "left edge" from "just left of centre", and sent drops to the wrong section
    # (developer, 2026-08-06). Synchronous throughout: every handler on the drag path runs
    # inline, and the eval route cannot await a promise.
    $r = Page @'
(() => {
  const fire = (el, type, x, y) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, clientX: x, clientY: y, button: 0, buttons: type === "pointerup" ? 0 : 1, bubbles: true, cancelable: true }));

  const dragToPetal = (pane, zone) => {
    const tab = [...document.querySelectorAll(".panel-tab")].find(t => t.dataset.panel === pane);
    if (!tab) return "no-tab";
    const from = tab.getBoundingClientRect();
    const area = document.getElementById("editor-area").getBoundingClientRect();
    fire(tab, "pointerdown", from.x + 10, from.y + 8);
    fire(window, "pointermove", from.x + 40, from.y + 8);
    fire(window, "pointermove", area.x + area.width / 2, area.y + area.height / 2);

    const petal = document.querySelector(".drop-petal-" + zone);
    if (!petal || petal.hidden) { fire(window, "pointerup", area.x, area.y); return "no-petal"; }
    const p = petal.getBoundingClientRect();
    fire(window, "pointermove", p.x + p.width / 2, p.y + p.height / 2);
    fire(window, "pointerup", p.x + p.width / 2, p.y + p.height / 2);
    return ["left","right","top","bottom"].find(s => [...document.getElementById("dock-" + s).querySelectorAll(".panel-tab")].some(t => t.dataset.panel === pane)) || "nowhere";
  };

  // A release over the editor's middle, off every petal, must drop nothing at all.
  const idle = (() => {
    const tab = [...document.querySelectorAll(".panel-tab")].find(t => t.dataset.panel === "problems");
    const before = ["left","right","top","bottom"].find(s => [...document.getElementById("dock-" + s).querySelectorAll(".panel-tab")].some(t => t.dataset.panel === "problems"));
    const from = tab.getBoundingClientRect();
    const area = document.getElementById("editor-area").getBoundingClientRect();
    fire(tab, "pointerdown", from.x + 10, from.y + 8);
    fire(window, "pointermove", from.x + 40, from.y + 8);
    fire(window, "pointermove", area.x + area.width / 2, area.y + area.height / 2);
    fire(window, "pointerup", area.x + area.width / 2, area.y + area.height / 2);
    const after = ["left","right","top","bottom"].find(s => [...document.getElementById("dock-" + s).querySelectorAll(".panel-tab")].some(t => t.dataset.panel === "problems"));
    return before === after ? "stayed" : "moved";
  })();

  return [dragToPetal("problems", "top"), dragToPetal("problems", "left"), idle].join(",");
})()
'@
    $r -eq 'top,left,stayed'
}

Check 'the preview shows the section it would join, and a drag ends when focus leaves' {
    # An edge whose section already stands is a JOIN, and the preview is that section's own
    # bounds — painting half the editor there described a drop that would not happen. An
    # edge with no section is a NEW one, dashed, because the shape is a proposal. And a drag
    # that loses the window — alt-tab, a screenshot tool, the host stealing focus, which
    # this host does freely — must end, or the dim and compass outlive it (2026-08-06).
    $r = Page @'
(() => {
  const fire = (el, type, x, y) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, clientX: x, clientY: y, button: 0, buttons: type === "pointerup" ? 0 : 1, bubbles: true, cancelable: true }));
  const area = document.getElementById("editor-area").getBoundingClientRect();
  const tab = [...document.querySelectorAll(".panel-tab")].find(t => t.dataset.panel === "properties");
  if (!tab) return "no-tab";
  const from = tab.getBoundingClientRect();

  fire(tab, "pointerdown", from.x + 10, from.y + 8);
  fire(window, "pointermove", from.x + 40, from.y + 8);
  fire(window, "pointermove", area.x + area.width / 2, area.y + area.height / 2);

  const sample = (zone) => {
    const petal = document.querySelector(".drop-petal-" + zone);
    if (!petal || petal.hidden) return zone + ":hidden";
    const p = petal.getBoundingClientRect();
    fire(window, "pointermove", p.x + p.width / 2, p.y + p.height / 2);
    const ov = document.querySelector(".drop-overlay");
    const dock = document.getElementById("dock-" + zone);
    if (!ov) return zone + ":no-preview";
    const kind = ov.className.includes("drop-overlay-join") ? "join" : "new";
    if (dock.hidden) return zone + ":" + kind;
    const a = ov.getBoundingClientRect(), b = dock.getBoundingClientRect();
    const fits = Math.abs(a.y - b.y) < 2 && Math.abs(a.height - b.height) < 2;
    return zone + ":" + kind + (fits ? ":fits" : ":stray");
  };

  const bottom = sample("bottom");
  window.dispatchEvent(new Event("blur"));
  const cleared = !document.querySelector(".drag-dim") && !document.querySelector(".drop-compass") && !document.querySelector(".drop-overlay");
  fire(window, "pointerup", area.x, area.y);
  return bottom + "|" + (cleared ? "cleared" : "stuck");
})()
'@
    $r -eq 'bottom:join:fits|cleared'
}

Check 'a pane tab reorders within its own strip, and the order persists' {
    $r = Page @'
(() => {
  const fire = (el, type, x, y) => el.dispatchEvent(new PointerEvent(type, { pointerId: 1, clientX: x, clientY: y, button: 0, buttons: type === "pointerup" ? 0 : 1, bubbles: true, cancelable: true }));
  const strip = () => {
    const anyTab = document.querySelector('.panel-tab[data-panel="immediate"]');
    return anyTab ? anyTab.closest(".dock-tabs") : null;
  };
  const order = () => [...strip().querySelectorAll(".panel-tab")].map(t => t.dataset.panel);

  const host = strip();
  if (!host) return "no-strip";
  const names = order();
  if (names.length < 2) return "too-few";

  const last = names[names.length - 1];
  const tab = host.querySelector(`.panel-tab[data-panel="${last}"]`);
  const first = host.querySelector(".panel-tab");
  const from = tab.getBoundingClientRect(), to = first.getBoundingClientRect();

  fire(tab, "pointerdown", from.x + 10, from.y + 8);
  fire(window, "pointermove", from.x - 20, from.y + 8);
  fire(window, "pointermove", to.x + 3, to.y + to.height / 2);
  fire(window, "pointerup", to.x + 3, to.y + to.height / 2);

  const after = order();
  const stored = JSON.parse(localStorage.getItem("xlide.docks.v1") || "{}");
  const find = (node) => !node ? null : node.kind === "group" ? (node.tabs.includes(last) ? node : null) : node.children.map(find).find(Boolean) || null;
  const group = ["left","right","top","bottom"].map(s => find(stored.sides ? stored.sides[s] : null)).find(Boolean);
  const persistedFirst = group ? group.tabs[0] : "(none)";
  return [after[0] === last ? "moved" : "stuck", persistedFirst === last ? "saved" : "unsaved"].join(",");
})()
'@
    $r -eq 'moved,saved'
}

Check 'the Panes menu toggles a pane off and back on' {
    $r = Page @'
(() => {
  const menu = () => document.getElementById("panes-menu");
  if (menu()) menu().remove();
  const button = document.querySelector('#toolbar [data-command="openPanes"]');
  if (!button) return "no-button";
  button.click();
  if (!menu()) return "no-menu";

  const itemFor = (name) => [...document.querySelectorAll("#panes-menu .panes-menu-item")].find(i => i.dataset.pane === name);
  const watchOpen = () => !!document.querySelector('.panel-tab[data-panel="watch"]');

  const started = watchOpen();
  itemFor("watch").click();
  const closed = !watchOpen();
  itemFor("watch").click();
  const reopened = watchOpen();
  const explorerLocked = itemFor("explorer").disabled;
  if (menu()) menu().remove();
  return [started, closed, reopened, explorerLocked].join(",");
})()
'@
    $r -eq 'true,true,true,true'
}

Check 'the reload re-opened every document' {
    [int](Page 'window.xlideBridge.documents.all().length') -ge 2
}

Check 'the reload brought the tabs back' {
    # A reloaded page came back with models and NO tabs once (2026-08-06): the tab list only
    # existed in the first boot's held replay. The host re-publishes on every ready now.
    WaitFor 'tabs after the reload' { [int](Page 'document.querySelectorAll(".tab").length') -ge 2 }
}

Check 'the reload brought the explorer back' {
    WaitFor 'explorer rows after the reload' { [int](Page 'document.querySelectorAll("#sidebar-tree [role=treeitem]").length') -ge 1 }
}

Check 'the layout is put back the way it was found' {
    # This probe DRAGS panes around to test the compass, and the arrangement is persistent
    # UI state. Clearing the storage key is not enough: the page still holds the rearranged
    # layout in memory and writes it back on the next render, so the developer opens the
    # editor to find Problems docked on the left with no explanation (2026-08-06). Clearing
    # AND reloading is what actually restores the default.
    Page 'localStorage.removeItem("xlide.docks.v1"); "cleared"' | Out-Null
    Page 'location.reload()' | Out-Null

    WaitFor 'the default arrangement to come back' {
        try {
            $where = (Page @'
(() => {
  if (!window.xlideBridge) return "";
  const side = ["left","right","top","bottom"].find(s => {
    const dock = document.getElementById("dock-" + s);
    return dock && !dock.hidden && dock.querySelector('.panel-tab[data-panel="problems"]');
  });
  return side || "";
})()
'@)
            $where -eq 'bottom'
        } catch { $false }
    } 25
}

$failed = @($checks.GetEnumerator() | Where-Object { $_.Value -ne 'PASS' })
foreach ($entry in $checks.GetEnumerator()) {
    Write-Output ("{0,-58} {1}" -f $entry.Key, $entry.Value)
}
Write-Output ("RESULT: " + $(if ($failed.Count -eq 0) { "PASS ($($checks.Count) checks)" } else { "FAIL ($($failed.Count) of $($checks.Count))" }))
