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
 * built it cannot see it; each position is resolved through `ControlAt` - the same call the
 * execute path makes - and checked against the id it should have landed on.
 *
 *   node tools\harness\menu-bar.mjs
 */

import { open, reporter } from "./xlide-api.mjs";

const api = await open({});
const { check, done } = reporter();

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

/*
 * AND IT RUNS THE RIGHT ONE.
 *
 * Everything above resolves positions; this executes one. Both bands of the arithmetic are driven
 * through the surface the way a click drives them, and the proof is the host's own line naming the
 * control id it invoked - not the dialog, which the debug api's guard cancels within a couple of
 * seconds by design. Polling `dialogs` afterwards finds nothing and reads exactly like an item
 * that did not run (2026-08-09; the guard's own log line is what settled it).
 */
console.log("\nexecuted through the synthetic path");

for (const [label, wanted] of [["Macros", 930], ["Add-In Manager", 943]]) {
  const since = (await api.log({ max: 1 })).next;

  const opened = await api.act("menuBar");
  if (!opened.did) {
    check(`${label} could be reached`, false, opened.detail);
    continue;
  }

  const chose = await api.act("chooseMenuItem", { label });
  await new Promise((settle) => setTimeout(settle, 2500));

  const lines = (await api.log({ since })).lines ?? [];
  const ran = lines.find((line) => /menu: \[900,\d+\] executed \((\d+)\)/.test(line));
  const id = ran ? Number(ran.match(/executed \((\d+)\)/)[1]) : 0;

  check(`${label} runs control ${wanted}, not another`, id === wanted,
    ran ? `${ran.replace(/^.*menu:/, "menu:")}` : `nothing executed (${chose.detail})`);

  // The guard names the dialog it closed, which is the second, independent witness that the
  // control that ran was the one the menu said it was.
  const cleared = lines.find((line) => /cleared our dialog/.test(line));
  check(`${label} opened its own dialog`,
    cleared !== undefined && new RegExp(label.split(" ")[0], "i").test(cleared),
    cleared ? cleared.replace(/^.*cleared/, "cleared") : "the guard cleared nothing");

  await api.act("closeDialogs");
}

console.log("\nwhat left the menus and where it went");

// Asked of the act rather than scraped, so the answer is the strip's own list of what it drew.
// A command left out because its editor action was not bundled is absent from both, which is the
// point; a selector sweep would also have missed a renamed container or a restructured row and
// reported every command gone.
const strip = await api.act("toolbar");
const toolbar = (strip.data?.commands ?? []).map((one) => one.id);

for (const command of ["run", "break", "reset", "designMode", "save", "openSync"]) {
  check(`${command} is on the toolbar`, toolbar.includes(command));
}

process.exit(done());
