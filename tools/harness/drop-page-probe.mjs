/*
 * Dragging a tab onto ANOTHER group's strip, with a real pointer.
 *
 * The cross-strip move-at-index existed for a week before anyone could see it: the landing was
 * computed and honoured, but hovering another group's strip showed NOTHING, so every cross-group
 * move went the long way through the compass centre (the developer, 2026-08-12). The insertion
 * bar is the fix, and this probe holds the whole gesture to it: press on a tab, cross to the
 * other strip, see the bar mid-drag with no compass standing, release, and find the tab AT THE
 * INDEX the bar showed. Then the same-strip reorder, whose feedback is the reorder itself.
 *
 * A REAL pointer, through Input.dispatchMouseEvent, because the drag arms at pointerdown with a
 * movement threshold and pointer capture: a synthetic DOM event sequence tests the synthesiser
 * (the standing rule), and CSS-level furniture like the bar only follows input the browser
 * itself delivers.
 */
import { runPageProbe } from "./page-probe.mjs";

const DRIVE = `(async () => {
  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));

  // The demo transport opens its two modules on its own schedule.
  for (let waited = 0; waited < 20000; waited += 200) {
    if (document.querySelectorAll(".tab").length >= 2) { break; }
    await sleep(200);
  }
  check("the demo page opened its two tabs", document.querySelectorAll(".tab").length >= 2);

  const split = window.xlideUi.act("split", { direction: "right" });
  check("the workspace splits right", split.did === true, split.detail);

  const groups = window.xlideUi.state().workspace.groups;
  check("two groups stand, one tab each",
    groups.length === 2 && groups[0].tabs.length === 1 && groups[1].tabs.length === 1,
    JSON.stringify(groups.map((g) => g.tabs.map((t) => t.module))));

  return { pass: checks.every((one) => one.ok), checks };
})()`;

/** Rects the pointer work needs, read live so nothing is guessed. */
const GEOMETRY = `(() => {
  const strips = [...document.querySelectorAll(".group-tabs")];
  const rect = (el) => { const r = el.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height }; };
  return {
    source: rect(strips[0].querySelector(".tab")),
    targetStrip: rect(strips[1]),
    targetTab: rect(strips[1].querySelector(".tab")),
  };
})()`;

await runPageProbe({
  label: "xlide-drop",
  drive: DRIVE,
  after: async ({ verdict, inPage, send, sessionId }) => {
    const mouse = (type, x, y, extra = {}) => send("Input.dispatchMouseEvent",
      { type, x: Math.round(x), y: Math.round(y), button: "left", ...extra }, sessionId);
    const settle = () => new Promise((done) => setTimeout(done, 120));

    const before = await inPage(GEOMETRY);
    const from = { x: before.source.x + before.source.w / 2, y: before.source.y + before.source.h / 2 };
    const to = { x: before.targetTab.x + before.targetTab.w + 30, y: before.targetStrip.y + before.targetStrip.h / 2 };

    // The gesture, one real step at a time: press, clear the 5px threshold, walk the strip row.
    await mouse("mousePressed", from.x, from.y, { buttons: 1, clickCount: 1 });
    await mouse("mouseMoved", from.x + 12, from.y, { buttons: 1 });
    await mouse("mouseMoved", (from.x + to.x) / 2, from.y, { buttons: 1 });
    await mouse("mouseMoved", to.x, to.y, { buttons: 1 });
    await settle();

    // Mid-drag, over the OTHER strip: the bar is showing at the end of that strip, and no
    // compass is standing - the strip is its own target, which is the whole point.
    const midFlight = await inPage(`(() => {
      const bar = document.querySelector(".tab-drop-indicator");
      const strips = [...document.querySelectorAll(".group-tabs")];
      const compass = document.querySelector(".drop-compass");
      return {
        barShowing: !!bar && bar.isConnected,
        barInTargetStrip: !!bar && bar.parentElement === strips[1],
        barIsLast: !!bar && bar.parentElement !== null && bar.parentElement.lastElementChild === bar,
        compassShowing: !!compass && !compass.hidden,
      };
    })()`);

    verdict.checks.push(
      { name: "mid-drag, the insertion bar shows on the target strip", ok: midFlight.barShowing && midFlight.barInTargetStrip, detail: JSON.stringify(midFlight) },
      { name: "at the index the drop would use", ok: midFlight.barIsLast, detail: null },
      { name: "and no compass stands - the strip is its own target", ok: !midFlight.compassShowing, detail: null });

    await mouse("mouseReleased", to.x, to.y, { clickCount: 1 });
    await settle();

    const landed = await inPage(`(() => {
      const s = window.xlideUi.state().workspace;
      return {
        groups: s.groups.map((g) => g.tabs.map((t) => t.module)),
        barGone: !document.querySelector(".tab-drop-indicator"),
      };
    })()`);

    // The emptied source group dissolves, so the landing reads as one group in the dropped
    // order: the moved tab AFTER the one it was dropped beyond.
    verdict.checks.push(
      { name: "the tab landed in the other group, at the dropped index", ok: JSON.stringify(landed.groups) === JSON.stringify([["Module1", "Module2"]]), detail: JSON.stringify(landed.groups) },
      { name: "and the bar went with the drag", ok: landed.barGone, detail: null });

    // The same-strip reorder: drag the second tab to the LEFT of the first. Feedback there is
    // the reorder itself - the element moves as the pointer crosses its neighbour's midpoint.
    const again = await inPage(`(() => {
      const tabs = [...document.querySelectorAll(".group-tabs .tab")];
      const rect = (el) => { const r = el.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height }; };
      return { second: rect(tabs[1]), first: rect(tabs[0]) };
    })()`);

    const grab = { x: again.second.x + again.second.w / 2, y: again.second.y + again.second.h / 2 };
    const drop = { x: again.first.x + 4, y: again.first.y + again.first.h / 2 };

    await mouse("mousePressed", grab.x, grab.y, { buttons: 1, clickCount: 1 });
    await mouse("mouseMoved", grab.x - 12, grab.y, { buttons: 1 });
    await mouse("mouseMoved", drop.x, drop.y, { buttons: 1 });
    await settle();
    await mouse("mouseReleased", drop.x, drop.y, { clickCount: 1 });
    await settle();

    const sorted = await inPage("window.xlideUi.state().workspace.groups.map((g) => g.tabs.map((t) => t.module))");
    verdict.checks.push(
      { name: "a same-strip drag reorders the tabs", ok: JSON.stringify(sorted) === JSON.stringify([["Module2", "Module1"]]), detail: JSON.stringify(sorted) });

    /* ---- the tree as a drag source: a module row dropped on a split zone ---- */

    const moduleRow = await inPage(`(() => {
      const row = document.querySelector('#sidebar-tree [data-component="Module1"]');
      if (!row) { return null; }
      const r = row.getBoundingClientRect();
      const body = document.querySelector(".group-body").getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2, body: { x: body.x + body.width / 2, y: body.y + body.height / 2 } };
    })()`);

    verdict.checks.push({ name: "the tree offers Module1's row to drag", ok: !!moduleRow, detail: null });
    if (!moduleRow) { return; }

    // Press the row, walk to the body centre so the compass stands, read the right petal
    // where it actually is, and land on it: the row becomes a split.
    await mouse("mousePressed", moduleRow.x, moduleRow.y, { buttons: 1, clickCount: 1 });
    await mouse("mouseMoved", moduleRow.x + 12, moduleRow.y, { buttons: 1 });
    await mouse("mouseMoved", moduleRow.body.x, moduleRow.body.y, { buttons: 1 });
    await settle();

    const midTree = await inPage(`(() => {
      const ghost = document.querySelector(".drag-ghost");
      const petal = document.querySelector(".drop-petal-right");
      if (!petal || petal.hidden) { return { ghost: ghost ? ghost.textContent : null, petal: null }; }
      const r = petal.getBoundingClientRect();
      return { ghost: ghost ? ghost.textContent : null, petal: { x: r.x + r.width / 2, y: r.y + r.height / 2 } };
    })()`);

    verdict.checks.push(
      { name: "a ghost chip rides the tree drag, wearing the module's name", ok: midTree.ghost === "Module1", detail: midTree.ghost },
      { name: "the compass stands over the editor body with its right petal offered", ok: !!midTree.petal, detail: null });
    if (!midTree.petal) { await mouse("mouseReleased", moduleRow.body.x, moduleRow.body.y, { clickCount: 1 }); return; }

    await mouse("mouseMoved", midTree.petal.x, midTree.petal.y, { buttons: 1 });
    await settle();
    await mouse("mouseReleased", midTree.petal.x, midTree.petal.y, { clickCount: 1 });
    await settle();

    const afterSplit = await inPage("window.xlideUi.state().workspace.groups.map((g) => ({ tabs: g.tabs.map((t) => t.module), active: g.active }))");
    verdict.checks.push({
      name: "dropping the open module's row on the right petal splits and moves its tab there",
      ok: afterSplit.length === 2
        && JSON.stringify(afterSplit[0].tabs) === JSON.stringify(["Module2"])
        && JSON.stringify(afterSplit[1].tabs) === JSON.stringify(["Module1"]),
      detail: JSON.stringify(afterSplit),
    });

    /* ---- a procedure row dropped on the OTHER strip: placement plus its line ---- */

    const procedure = await inPage(`(() => {
      const rows = [...document.querySelectorAll('#sidebar-tree [data-proc-module="Module1"]')];
      const row = rows.find((one) => (one.textContent ?? "").includes("Describe")) ?? rows[0];
      if (!row) { return null; }
      const r = row.getBoundingClientRect();
      const strip = document.querySelectorAll(".group-tabs")[0].getBoundingClientRect();
      return {
        x: r.x + r.width / 2, y: r.y + r.height / 2,
        line: Number(row.dataset.procLine),
        strip: { x: strip.x + 8, y: strip.y + strip.h / 2 || strip.y + strip.height / 2 },
      };
    })()`);

    verdict.checks.push({ name: "the tree offers a procedure row of Module1 to drag", ok: !!procedure, detail: null });
    if (!procedure) { return; }

    await mouse("mousePressed", procedure.x, procedure.y, { buttons: 1, clickCount: 1 });
    await mouse("mouseMoved", procedure.x + 12, procedure.y, { buttons: 1 });
    await mouse("mouseMoved", procedure.strip.x, procedure.strip.y, { buttons: 1 });
    await settle();
    await mouse("mouseReleased", procedure.strip.x, procedure.strip.y, { clickCount: 1 });
    await settle();
    // The navigate that carries the line is the row click's own path; give it a beat.
    await settle();

    const afterProcedure = await inPage(`(() => {
      const s = window.xlideUi.state();
      return {
        groups: s.workspace.groups.map((g) => g.tabs.map((t) => t.module)),
        focus: { model: s.focus.model, line: s.focus.line },
      };
    })()`);

    verdict.checks.push(
      {
        name: "dropping the procedure row at the strip's front moves the module's tab to index 0",
        ok: JSON.stringify(afterProcedure.groups) === JSON.stringify([["Module1", "Module2"]]),
        detail: JSON.stringify(afterProcedure.groups),
      },
      {
        name: "and the editor stands on the procedure's own line",
        ok: (afterProcedure.focus.model ?? "").toLowerCase().includes("module1")
          && afterProcedure.focus.line === procedure.line,
        detail: JSON.stringify({ focus: afterProcedure.focus, wanted: procedure.line }),
      });

    /* ---- a module with NO tab, dropped at an index: the placement hint's whole reason ---- */

    const closed = await inPage(`(async () => {
      const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
      // Module1 is dirty in the demo, so closing it raises Save / Don't Save; discard answers it.
      window.xlideUi.act("closeActive", {});
      await sleep(150);
      window.xlideUi.act("answerCloseConfirm", { answer: "discard" });
      for (let waited = 0; waited < 5000; waited += 100) {
        const tabs = window.xlideUi.state().workspace.groups.flatMap((g) => g.tabs.map((t) => t.module));
        if (!tabs.includes("Module1")) { return tabs; }
        await sleep(100);
      }
      return null;
    })()`);

    verdict.checks.push({ name: "Module1's tab closes, leaving only its tree row", ok: JSON.stringify(closed) === JSON.stringify(["Module2"]), detail: JSON.stringify(closed) });
    if (!closed) { return; }

    const reopen = await inPage(`(() => {
      const row = document.querySelector('#sidebar-tree [data-component="Module1"]');
      const strips = [...document.querySelectorAll(".group-tabs")];
      const strip = strips[strips.length - 1].getBoundingClientRect();
      const tab = strips[strips.length - 1].querySelector(".tab").getBoundingClientRect();
      const r = row.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2, after: { x: tab.x + tab.width + 30, y: strip.y + strip.height / 2 } };
    })()`);

    await mouse("mousePressed", reopen.x, reopen.y, { buttons: 1, clickCount: 1 });
    await mouse("mouseMoved", reopen.x + 12, reopen.y, { buttons: 1 });
    await mouse("mouseMoved", reopen.after.x, reopen.after.y, { buttons: 1 });
    await settle();
    await mouse("mouseReleased", reopen.after.x, reopen.after.y, { clickCount: 1 });

    // The open is a round trip even in the demo; the placement hint holds the seat meanwhile.
    const placed = await inPage(`(async () => {
      const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
      for (let waited = 0; waited < 8000; waited += 150) {
        const groups = window.xlideUi.state().workspace.groups.map((g) => g.tabs.map((t) => t.module));
        if (JSON.stringify(groups) === JSON.stringify([["Module2", "Module1"]])) { return groups; }
        await sleep(150);
      }
      return window.xlideUi.state().workspace.groups.map((g) => g.tabs.map((t) => t.module));
    })()`);

    verdict.checks.push({
      name: "dragging the closed module's row to the strip opens it AT the dropped index",
      ok: JSON.stringify(placed) === JSON.stringify([["Module2", "Module1"]]),
      detail: JSON.stringify(placed),
    });
  },
});
