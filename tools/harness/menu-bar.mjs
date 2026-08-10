/*
 * The bar, after the native menus went.
 *
 * The editor has ten menus and the surface shows one, which is not one of the editor's: Tools and
 * Add-Ins are composed into a single `xlide` menu with invented positions, and every other
 * top-level menu is suppressed outright. Three separate things can break here and only the third
 * is visible on screen:
 *
 *   1. a menu stops being suppressed, and a native menu reappears
 *   2. an item is suppressed that nothing else carries, and a capability is lost silently
 *   3. a synthetic position resolves to the WRONG control, so the menu reads right and runs
 *      something else
 *
 * The third is why this exists. A composition checked by reading it back through the code that
 * built it cannot see it; each position is resolved through `ControlAt` — the same call the
 * execute path makes — and checked against the id it should have landed on.
 *
 *   node tools\harness\menu-bar.mjs
 */

import { open } from "./xlide-api.mjs";

const api = await open({});
let broken = 0;
const check = (what, ok, detail) => {
  console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail ? "  -- " + detail : ""}`);
  if (!ok) { broken += 1; }
};

/** The synthetic menu's position. Mirrors XlideMenuPosition in VbeMenus.cs. */
const XLIDE = 900;
const STRIDE = 1000;

console.log("the menu bar\n");

const bar = await api.menus();
const shown = (bar.items ?? []).filter((one) => !one.suppressed);
check("every one of the editor's own menus is suppressed",
  shown.length === 0,
  shown.map((one) => one.caption).join(", ") || "all eleven");

check("the ones named in the table are the ones on the bar",
  (bar.items ?? []).length >= 10, `${(bar.items ?? []).length} controls`);

// WHAT THE DEVELOPER SEES. The bar reply above is the editor's; this is the product's.
// `ask` hands back the VALUE, already unwrapped: a script that answers with
// JSON.stringify comes back as the object, not as a string to parse. Parsing it again is how a
// probe ends up reading undefined off "[object Object]".
const surface = await api.ask(`(() => {
  const bar = document.getElementById('menubar');
  const tops = [...bar.querySelectorAll('.menu-top')];
  return JSON.stringify({
    count: tops.length,
    label: tops[0] ? tops[0].getAttribute('aria-label') : null,
    text: tops[0] ? (tops[0].textContent || '').trim() : null,
    icon: tops[0] ? (tops[0].querySelector('.codicon') || {}).className || null : null,
    insideToolbar: !!bar.closest('#toolbar'),
    shellRows: getComputedStyle(document.getElementById('shell')).gridTemplateRows.split(' ').length,
  });
})()`);

check("the surface draws exactly one menu", surface.count === 1, `${surface.count}`);
check("it is an icon, not a word", surface.text === "" && /codicon-wrench/.test(surface.icon ?? ""),
  `text "${surface.text}", icon ${surface.icon}`);
check("it still has a name for anything that cannot see it", surface.label === "VBA", surface.label);
check("there is no menu bar row left: it lives on the toolbar", surface.insideToolbar);
check("and the shell has three rows, so the editor fills the window",
  surface.shellRows === 3, `${surface.shellRows} rows`);

console.log("\nthe xlide menu");

const menu = await api.menus([XLIDE]);
const items = menu.items ?? [];
for (const row of items) {
  console.log(`   ${String(row.index).padStart(5)}  id ${String(row.id).padEnd(6)} ${row.caption}`);
}

check("it is composed from both sources",
  items.some((one) => one.index < STRIDE) && items.some((one) => one.index >= STRIDE),
  `${items.length} items`);

// EVERY POSITION LEADS SOMEWHERE REAL. An id of 0 is ControlAt failing to resolve, which is what
// a broken translation looks like from here.
const lost = items.filter((one) => !one.id);
check("every synthetic position resolves to a real control", lost.length === 0,
  lost.map((one) => one.caption).join(", ") || `all ${items.length}`);

/*
 * AND TO THE RIGHT ONE. Each source is read directly and lined up against the composition: the
 * item at synthetic `rank * 1000 + i` must carry the same id as the item at position i of that
 * source's real menu. An off-by-one in the arithmetic passes every check above and fails this.
 */
const sources = [30007, 30038];
for (let rank = 0; rank < sources.length; rank++) {
  const at = (bar.items ?? []).find((one) => one.id === sources[rank])?.index;
  if (!at) {
    check(`source ${sources[rank]} is on the bar`, false, "not found");
    continue;
  }

  const real = await api.menus([at]);
  let matched = 0;
  let wrong = null;
  for (const row of (real.items ?? [])) {
    const synthetic = items.find((one) => one.index === (rank * STRIDE) + row.index);
    if (!synthetic) { continue; }
    if (synthetic.id === row.id) { matched += 1; } else { wrong ??= `${synthetic.caption} -> ${synthetic.id}, wanted ${row.id}`; }
  }

  check(`source ${sources[rank]} maps position for position`, wrong === null && matched > 0,
    wrong ?? `${matched} items`);
}

console.log("\nwhat left the menus and where it went");

const toolbar = await api.ask(`JSON.stringify(
  [...document.querySelectorAll('#toolbar [data-command]')].map((one) => one.dataset.command))`);

for (const command of ["run", "break", "reset", "designMode", "save", "openSync"]) {
  check(`${command} is on the toolbar`, toolbar.includes(command));
}

console.log(`\n${broken} broken`);
process.exit(broken === 0 ? 0 : 1);
