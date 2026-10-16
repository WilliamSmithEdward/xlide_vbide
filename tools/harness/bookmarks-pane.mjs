/*
 * Bookmarks, as a developer reaches them (#111): the toolbar's Bookmarks group and the
 * Bookmarks pane.
 *
 * The group is one button on the strip whose four rows open on hover, click or keyboard focus.
 * The pane lists every module's marks with module, line and the line's text; a row goes to its
 * line, Delete takes its mark away, Clear All takes every one, and a closed module's marks stay
 * listed. Marks move with their lines as text is edited, and the pane has to move with them.
 *
 * Mouse and keys go through the page's DevTools protocol (Input.dispatch*), which delivers real,
 * trusted events to the page only - so the hover pause, the hit test and the key handling are
 * the ones a person meets. Nothing goes to the desktop.
 *
 * The file filter needs more than one file open, so this runs in the gate's three-file session
 * (TestFixture, TestTwinFixture and DebugFixture), whose DebugFixture holds Helper and Runner:
 *   node tools\harness\bookmarks-pane.mjs
 */

import { open, reporter, wait } from "./xlide-api.mjs";
import { loopbackUrl } from "./loopback-url.mjs";

const api = await open({});
const { check, done } = reporter();

const targets = await (await fetch(loopbackUrl(api.devtoolsPort, "/json"), { redirect: "error" })).json();
const page = targets.find((row) => row.title === "xlide editor");
if (!page) {
  throw new Error("the editor page's DevTools target is not listed");
}
const socket = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
let nextId = 0;
const waiting = new Map();
socket.onmessage = ({ data }) => {
  const reply = JSON.parse(data);
  const pending = waiting.get(reply.id);
  if (!pending) return;
  waiting.delete(reply.id);
  reply.error ? pending.reject(new Error(reply.error.message)) : pending.resolve(reply.result);
};
const cdp = (method, params) => new Promise((resolve, reject) => {
  const id = ++nextId;
  waiting.set(id, { resolve, reject });
  socket.send(JSON.stringify({ id, method, params }));
});

const js = (expression) => cdp("Runtime.evaluate", { expression, returnByValue: true }).then((r) => r.result.value);
const move = (x, y) => cdp("Input.dispatchMouseEvent", { type: "mouseMoved", x, y, button: "none" });
const click = async (x, y) => {
  await move(x, y);
  await cdp("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 });
  await cdp("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 });
};
const key = async (k, code, keyCode, text) => {
  await cdp("Input.dispatchKeyEvent", { type: "rawKeyDown", key: k, code, windowsVirtualKeyCode: keyCode });
  if (text) {
    await cdp("Input.dispatchKeyEvent", { type: "char", key: k, code, text });
  }
  await cdp("Input.dispatchKeyEvent", { type: "keyUp", key: k, code, windowsVirtualKeyCode: keyCode });
};
// The selector travels as an ARGUMENT, never spliced into the source: callFunctionOn hands the
// page the value itself, so no string of ours is ever read as code there.
const pageGlobal = (await cdp("Runtime.evaluate", { expression: "globalThis" })).result.objectId;
const centre = async (selector) => (await cdp("Runtime.callFunctionOn", {
  objectId: pageGlobal,
  functionDeclaration: `function (selector) {
    const e = document.querySelector(selector);
    if (!e) return null;
    e.scrollIntoView({ block: 'nearest', inline: 'center' });
    const r = e.getBoundingClientRect();
    return { x: r.x + Math.min(40, r.width / 2), y: r.y + r.height / 2 };
  }`,
  arguments: [{ value: selector }],
  returnByValue: true,
})).result.value;
const away = () => move(5, 400);

const pane = async () => (await api.ui()).bookmarksPane.rows;
const paneState = async () => (await api.ui()).bookmarksPane;
const brief = (rows) => rows.map((row) => `${row.module}:${row.line}`).join(", ");

// Every module this suite touches is named with its file: two files can each hold one of a name.
const projects = (await api.ui()).explorer.projects;
const DEBUG = projects.find((one) => one.modules.some((module) => module.name === "Helper"))?.name;
const OTHER = projects.find((one) => one.name !== DEBUG
  && one.modules.some((module) => module.name === "InvoiceTests"))?.name;
const EMPTY = projects.find((one) => one.name !== DEBUG && one.name !== OTHER)?.name;
if (!DEBUG || !OTHER || !EMPTY) {
  throw new Error("this suite needs the three-file session: TestFixture, TestTwinFixture and DebugFixture");
}

const lineOf = async (module, line, project = DEBUG) =>
  (await api.readModule(module, project)).text.split(/\r?\n/)[line - 1];
const markAt = async (module, line, project = DEBUG) => {
  await api.act("activate", { module, project }); await wait(400);
  await api.caret(line, { module, project }); await wait(300);
  await api.act("bookmark", { which: "toggle" }); await wait(300);
};

try {
  console.log("the toolbar's Bookmarks group\n");

  await api.act("bookmark", { which: "clear" });
  await api.act("activate", { module: "Helper", project: DEBUG }); await wait(400);

  const face = await centre('#toolbar [data-group="bookmarks"]');
  check("the strip draws a Bookmarks group button", face !== null);
  await wait(300);
  const at = await centre('#toolbar [data-group="bookmarks"]');
  const flyout = () => js(`(() => { const f = document.querySelector('.toolbar-flyout[data-group="bookmarks"]');
    return f ? [...f.querySelectorAll('.menu-item')].map((r) => r.dataset.command) : null; })()`);

  await away(); await wait(400);
  await move(at.x, at.y); await wait(50); await away(); await wait(400);
  check("a pointer crossing the button does not open the group", (await flyout()) === null);

  await move(at.x, at.y); await wait(400);
  check("resting on it opens its four rows", (await flyout())?.length === 4, JSON.stringify(await flyout()));

  await api.caret(3, { module: "Helper", project: DEBUG }); await wait(300);
  await away(); await wait(500);
  await move(at.x, at.y); await wait(400);
  const toggle = await centre('.toolbar-flyout [data-command="xlide.bookmark.toggle"]');
  await move(at.x, at.y + 14); await move(toggle.x, toggle.y); await wait(100);
  await click(toggle.x, toggle.y); await wait(500);
  let ui = await api.ui();
  check("its Toggle row marks the caret's line", JSON.stringify(ui.bookmarks) === "[3]", JSON.stringify(ui.bookmarks));
  check("choosing a row closes the group", (await flyout()) === null);
  check("and gives the keyboard back to the editor", ui.focus?.hasFocus === true);

  await move(at.x, at.y); await wait(400);
  await away(); await wait(600);
  check("moving away closes it", (await flyout()) === null);

  await js(`document.querySelector('#toolbar [data-group="bookmarks"]').focus()`);
  await key("ArrowDown", "ArrowDown", 40); await wait(100);
  check("Down from the focused button opens it onto its first row",
    await js(`document.activeElement?.dataset?.command === 'xlide.bookmark.toggle'`));
  await key("ArrowUp", "ArrowUp", 38);
  check("Up wraps to the last row", await js(`document.activeElement?.dataset?.command === 'xlide.bookmark.clearAll'`));
  await key("Escape", "Escape", 27); await wait(150);
  check("Escape closes it and stays closed", (await flyout()) === null);
  check("and hands the focus back to the button", await js(`document.activeElement?.dataset?.group === 'bookmarks'`));

  const textBefore = (await api.readModule("Helper", DEBUG)).text;
  await api.caret(1, { module: "Helper", project: DEBUG }); await wait(300);
  await js(`document.querySelector('#toolbar [data-group="bookmarks"]').focus()`);
  await key("ArrowDown", "ArrowDown", 40); await key("ArrowDown", "ArrowDown", 40);
  await key("Enter", "Enter", 13, "\r"); await wait(500);
  check("Enter on Next bookmark jumps from line 1 to the mark", (await api.ui()).focus?.line === 3);
  await wait(1500);
  check("and types nothing into the module", (await api.readModule("Helper", DEBUG)).text === textBefore);

  const listed = await api.act("toolbar");
  const grouped = (listed.data?.commands ?? []).filter((one) => one.group === "bookmarks").map((one) => one.id);
  check("act toolbar lists the grouped commands with their group", grouped.length === 4, JSON.stringify(grouped));
  const pressed = await api.act("toolbar", { command: "xlide.bookmark.clearAll" }); await wait(300);
  check("act toolbar presses a grouped command through its row",
    pressed.did && (await api.ui()).bookmarks.length === 0, pressed.detail);

  console.log("\nthe Bookmarks pane\n");

  check("the pane starts empty", (await pane()).length === 0);
  check("an empty pane says how to make one",
    /Ctrl\+Alt\+K/.test(await js(`document.querySelector('#bookmarks-list .panel-empty')?.textContent ?? ''`)));

  await markAt("Runner", 2);
  await markAt("Helper", 5);
  await markAt("Helper", 3);
  let rows = await pane();
  check("it lists every module's marks, by module then line", brief(rows) === "Helper:3, Helper:5, Runner:2", brief(rows));
  check("each row says what its line says", rows[0]?.text === await lineOf("Helper", 3), JSON.stringify(rows[0]?.text));

  const tab = await centre('.panel-tab[data-panel="bookmarks"]');
  check("the pane has a tab", tab !== null);
  await click(tab.x, tab.y); await wait(300);
  check("the tab wears the count",
    await js(`document.querySelector('.panel-tab[data-panel="bookmarks"] .tab-badge')?.textContent`) === "3");
  const drawn = await js(`[...document.querySelectorAll('#bookmarks-list .bookmark-row .where')].map((w) => w.textContent)`);
  check("with several files open, the rows name file, module and line",
    JSON.stringify(drawn) === JSON.stringify([`${DEBUG} - Helper, line 3`, `${DEBUG} - Helper, line 5`, `${DEBUG} - Runner, line 2`]),
    JSON.stringify(drawn));

  await api.act("activate", { module: "Helper", project: DEBUG }); await wait(300);
  const runnerRow = await centre('#bookmarks-list [data-i="2"]');
  await click(runnerRow.x, runnerRow.y); await wait(800);
  ui = await api.ui();
  check("clicking a row opens its module at its line", ui.statusModule === "Runner" && ui.focus?.line === 2,
    `${ui.statusModule} line ${ui.focus?.line}`);

  await api.caret(1, { module: "Runner", project: DEBUG }); await wait(200);
  await api.act("press", { key: "Enter" }); await wait(600);
  rows = await pane();
  check("a line inserted above a mark moves its row", brief(rows) === "Helper:3, Helper:5, Runner:3", brief(rows));
  await api.act("undo"); await wait(600);
  rows = await pane();
  check("and undo moves it back", brief(rows) === "Helper:3, Helper:5, Runner:2", brief(rows));

  await js(`document.querySelector('#bookmarks-list [data-i="1"]').focus()`);
  await key("Delete", "Delete", 46); await wait(300);
  rows = await pane();
  check("Delete takes the focused row's mark away", brief(rows) === "Helper:3, Runner:2", brief(rows));
  check("and keeps the keyboard in the list", await js(`document.activeElement?.closest('#bookmarks-list') !== null`));
  await api.act("activate", { module: "Helper", project: DEBUG }); await wait(300);
  check("the editor's margin agrees", JSON.stringify((await api.ui()).bookmarks) === "[3]");

  const kept = await lineOf("Helper", 3);
  await api.act("closeActive"); await wait(600);
  rows = await pane();
  check("a closed module's mark stays listed with its text", brief(rows) === "Helper:3, Runner:2" && rows[0]?.text === kept, brief(rows));
  const heldRow = await centre('#bookmarks-list [data-i="0"]');
  await click(heldRow.x, heldRow.y); await wait(1200);
  ui = await api.ui();
  check("its row reopens the module at the line", ui.statusModule === "Helper" && ui.focus?.line === 3,
    `${ui.statusModule} line ${ui.focus?.line}`);
  check("and the reopened module wears the mark again", JSON.stringify(ui.bookmarks) === "[3]");

  let act = await api.act("bookmarksPane", { row: 1, press: "goto" }); await wait(800);
  ui = await api.ui();
  check("act bookmarksPane goto goes to the row's line", act.did && ui.statusModule === "Runner" && ui.focus?.line === 2, act.detail);
  act = await api.act("bookmarksPane", { row: 1, press: "remove" }); await wait(300);
  check("act bookmarksPane remove takes its mark away", act.did && brief(await pane()) === "Helper:3", act.detail);
  act = await api.act("bookmarksPane", { row: 9 });
  check("a row past the end is refused", !act.did, act.detail);

  console.log("\nthe Bookmarks pane's file filter\n");

  await markAt("InvoiceTests", 2, OTHER);
  let state = await paneState();
  check("every open file's marks show by default", state.file === null && state.rows.length === 2,
    JSON.stringify(state.rows.map((row) => `${row.project}/${row.module}:${row.line}`)));
  check("the filter offers every open file", [DEBUG, OTHER, EMPTY].every((name) => state.files.includes(name)),
    JSON.stringify(state.files));
  check("the filter is on screen", await js(`!document.querySelector('#bookmarks-scope-file').hidden`));

  // Picked with the keyboard on the select itself: Down moves a closed select to the next option.
  await js(`document.querySelector('#bookmarks-scope-file').focus()`);
  await key("ArrowDown", "ArrowDown", 40); await wait(300);
  state = await paneState();
  const first = state.files[0];
  check("choosing a file in the select narrows the rows to it",
    state.file === first && state.rows.every((row) => row.project === first),
    `${state.file}: ${brief(state.rows)}`);
  check("and the select says it is narrowing", await js(`document.querySelector('#bookmarks-scope-file').classList.contains('scope-narrowed')`));

  act = await api.act("bookmarksPane", { file: OTHER }); await wait(300);
  state = await paneState();
  check("act bookmarksPane {file} narrows to that file", act.did && state.file === OTHER && brief(state.rows) === "InvoiceTests:2",
    `${act.detail}: ${brief(state.rows)}`);
  const narrowed = await js(`[...document.querySelectorAll('#bookmarks-list .bookmark-row .where')].map((w) => w.textContent)`);
  check("inside one file the rows leave the file out", JSON.stringify(narrowed) === JSON.stringify(["InvoiceTests, line 2"]), JSON.stringify(narrowed));
  check("the tab still counts every file's marks",
    await js(`document.querySelector('.panel-tab[data-panel="bookmarks"] .tab-badge')?.textContent`) === "2");

  check("narrowed, the bar's button clears the file rather than everything",
    await js(`document.querySelector('#bookmarks-clear').textContent.trim()`) === "Clear File");
  await markAt("InvoiceTests", 4, OTHER);
  await api.act("bookmarksPane", { file: OTHER }); await wait(300);
  const clearFile = await centre("#bookmarks-clear");
  await click(clearFile.x, clearFile.y); await wait(300);
  state = await paneState();
  check("and pressing it takes only that file's marks", state.rows.length === 0, brief(state.rows));
  await api.act("bookmarksPane", { file: "all" }); await wait(300);
  check("the other file's marks are still there", brief((await paneState()).rows) === "Helper:3", brief((await paneState()).rows));
  await markAt("InvoiceTests", 2, OTHER);
  check("back on every file, the button clears all again",
    await js(`document.querySelector('#bookmarks-clear').textContent.trim()`) === "Clear All");

  await api.act("bookmarksPane", { file: EMPTY }); await wait(300);
  const emptySaid = await js(`document.querySelector('#bookmarks-list .panel-empty span')?.textContent ?? ''`);
  check("a file with none says so and how many are elsewhere", emptySaid === `No bookmarks in ${EMPTY}. 2 elsewhere.`, emptySaid);
  const showAll = await centre("#bookmarks-list .panel-empty-act");
  await click(showAll.x, showAll.y); await wait(300);
  state = await paneState();
  check("its Show All goes back to every file", state.file === null && state.rows.length === 2, brief(state.rows));
  act = await api.act("bookmarksPane", { file: "NoSuchFile.xlsm" });
  check("a file the filter does not offer is refused", !act.did, act.detail);

  const clear = await centre("#bookmarks-clear");
  await click(clear.x, clear.y); await wait(300);
  check("the bar's Clear All empties the pane", (await pane()).length === 0);
  check("and greys itself", await js(`document.querySelector('#bookmarks-clear').disabled`));
  check("and the tab drops its count", await js(`document.querySelector('.panel-tab[data-panel="bookmarks"] .tab-badge') === null`));

  // A window too narrow for the pane tabs: the strip gets edges, and a count arriving on the
  // last tab while the strip sits at its end stays in view (the first bookmark's, 2026-10-16).
  console.log("\nthe pane tabs in a narrow window\n");
  await cdp("Emulation.setDeviceMetricsOverride", { width: 720, height: 700, deviceScaleFactor: 1, mobile: false });
  await wait(600);
  const strip = () => js(`(() => {
    const tab = document.querySelector('.panel-tab[data-panel="bookmarks"]');
    const strip = tab.closest('.dock-tabs'); const head = strip.parentElement;
    const s = strip.getBoundingClientRect();
    const badge = tab.querySelector('.tab-badge')?.getBoundingClientRect();
    const edge = (end) => { const b = head.querySelector('.tab-edge-' + end); return !!b && !b.hidden; };
    return { overflow: strip.scrollWidth > strip.clientWidth, back: edge('start'), forward: edge('end'),
      right: s.right, badgeRight: badge ? badge.right : null };
  })()`);
  await js(`document.querySelector('.panel-tab[data-panel="bookmarks"]').closest('.dock-tabs').scrollLeft = 0`);
  await wait(300);
  let narrow = await strip();
  check("the tabs overflow, and at the start the strip shows only its forward edge",
    narrow.overflow && narrow.forward && !narrow.back, JSON.stringify(narrow));
  await js(`(() => { const s = document.querySelector('.panel-tab[data-panel="bookmarks"]').closest('.dock-tabs'); s.scrollLeft = s.scrollWidth; })()`);
  await wait(300);
  narrow = await strip();
  check("at the far end the forward edge goes and the back edge comes", !narrow.forward && narrow.back, JSON.stringify(narrow));
  await markAt("Helper", 3); await wait(300);
  narrow = await strip();
  check("the first bookmark's count lands in view at the far end",
    narrow.badgeRight !== null && narrow.badgeRight <= narrow.right + 1, JSON.stringify(narrow));
} finally {
  await cdp("Emulation.clearDeviceMetricsOverride", {}).catch(() => {});
  await api.act("bookmark", { which: "clear" }).catch(() => {});
  await away().catch(() => {});
  socket.close();
}

process.exit(done());
