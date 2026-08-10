// Drives the BUILT editor page (ui/editor/dist) in a headless browser and walks the TREE'S ROW
// CONTROLS end to end against the page's demo transport.
//
// Two things live on a row and both are here, because they are one surface: the plus that adds to
// a workbook, and the menus that a right-click opens. The destructive one gets the most attention
// — the menu offers Remove on a module and not on a document, the box asks before anything goes,
// Cancel and Escape keep the module, and Remove takes it out of the tree AND out of the tab strip
// without raising the unsaved-changes question on the way.
//
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.
//
// Everything goes through `xlideUi.act`, so the api a script would use is what is under test
// here alongside the feature. No dependencies: Node's own http server serves the dist, Edge
// provides the browser, and the DevTools protocol runs over Node's global WebSocket.

import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const dist = resolve(here, "..", "..", "ui", "editor", "dist");

const TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".ttf": "font/ttf",
  ".json": "application/json",
};

const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const act = (name, args) => window.xlideUi.act(name, args ?? {});
  const tabs = () => [...document.querySelectorAll('.tab')].map((tab) => tab.dataset.module);
  const rows = () => [...document.querySelectorAll('[data-component]')]
    .map((one) => one.dataset.component);
  const box = () => document.getElementById('remove-confirm-backdrop');
  const boxText = () => document.getElementById('remove-confirm-title')?.textContent ?? '';

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  // The demo publishes the tree a beat after load; nothing below means anything until it has.
  for (let waited = 0; tabs().length < 2 && waited < 20000; waited += 200) await sleep(200);
  act('expandWorkbook', { workbook: 'Book1.xlsm', open: true });
  for (let waited = 0; !rows().includes('Module1') && waited < 10000; waited += 200) await sleep(200);
  check('the demo tree lists Book1 with Module1 in it', rows().includes('Module1'), rows().join(','));

  // THE PLUS ON EVERY WORKBOOK ROW.
  const plusButtons = [...document.querySelectorAll('.tree-workbook .tree-add')];
  const workbookRows = [...document.querySelectorAll('.tree-workbook')];
  check('every workbook row carries a plus',
    workbookRows.length >= 2 && plusButtons.length === workbookRows.length,
    plusButtons.length + ' of ' + workbookRows.length);

  check('the plus is out of sight until the row is wanted',
    getComputedStyle(plusButtons[0]).opacity === '0',
    getComputedStyle(plusButtons[0]).opacity);

  // Hidden by opacity and NOT by display or visibility, which is what keeps its space in the row
  // (no reflow as the pointer crosses the tree) and keeps it in the tab order.
  check('hiding it did not take it out of the layout',
    plusButtons[0].getBoundingClientRect().width === 24
      && getComputedStyle(plusButtons[0]).visibility === 'visible',
    getComputedStyle(plusButtons[0]).visibility + ', '
      + Math.round(plusButtons[0].getBoundingClientRect().width) + 'px wide');

  // The keyboard's way back to it. A pointer-only reveal would leave this at 0.
  //
  // AFTER THE FADE, not during it. getComputedStyle reports the value the transition has reached,
  // so reading it in the same tick as the focus reports the opacity it is coming FROM — 0, every
  // time, which reads exactly like a rule that does not work.
  plusButtons[0].focus();
  await sleep(200);
  check('focusing it brings it back for a keyboard',
    getComputedStyle(plusButtons[0]).opacity === '1',
    getComputedStyle(plusButtons[0]).opacity);
  plusButtons[0].blur();
  await sleep(200);

  const plus = plusButtons[0];
  const plusBox = plus.getBoundingClientRect();
  check('the plus meets the 24px target', plusBox.width >= 24 && plusBox.height >= 24,
    Math.round(plusBox.width) + 'x' + Math.round(plusBox.height));

  const rowBox = workbookRows[0].getBoundingClientRect();
  check('the plus is right-aligned in its row',
    Math.abs(rowBox.right - plusBox.right) <= 6,
    'row right ' + Math.round(rowBox.right) + ', plus right ' + Math.round(plusBox.right));

  check('the plus is green, not the text colour',
    getComputedStyle(plus.querySelector('.codicon')).color
      !== getComputedStyle(workbookRows[0]).color,
    getComputedStyle(plus.querySelector('.codicon')).color);

  /*
   * ALWAYS VISIBLE, at any pane width. The row used to be content-box, so it was the pane's width
   * PLUS its own padding, and the last ten pixels of it — where a right-aligned control sits —
   * hung off the edge into the tree's own horizontal scroll. The control was there and could not
   * be seen, which is worse than not having it.
   *
   * Checked against the SCROLLPORT rather than against the row, because the row was the thing
   * that was wrong.
   */
  const tree = workbookRows[0].parentElement;
  const visible = () => {
    const port = tree.getBoundingClientRect();
    const button = tree.querySelector('.tree-add').getBoundingClientRect();
    return { inside: button.left >= port.left - 1 && button.right <= port.right + 1,
      detail: 'plus ' + Math.round(button.left) + '-' + Math.round(button.right)
        + ' in pane ' + Math.round(port.left) + '-' + Math.round(port.right) };
  };

  const wasWidth = tree.style.width;
  for (const width of ['420px', '260px', '190px', '150px']) {
    tree.style.width = width;
    const seen = visible();
    // No backticks in here: this whole drive is itself a template literal, and a nested one ends
    // it early. The error lands on a line that looks fine, which is why this note is here.
    check('the plus is on screen in a ' + width + ' pane', seen.inside, seen.detail);
  }
  tree.style.width = wasWidth;

  // A NAME TOO LONG FOR THE PANE MUST NOT PUSH THE PLUS OFF IT. The name is what gives way, and
  // the ellipsis is the browser's, so what is asserted is that the element is narrower than the
  // text it holds and the plus is still inside the row.
  const label = workbookRows[0].querySelector('.tree-workbook-name');
  const wasName = label.textContent;
  label.textContent = 'A'.repeat(200) + 'Workbook_With_A_Very_Long_Name.xlsm';
  const squeezed = plus.getBoundingClientRect();
  const squeezedRow = workbookRows[0].getBoundingClientRect();
  check('a long name truncates rather than pushing the plus out',
    label.scrollWidth > label.clientWidth && squeezed.right <= squeezedRow.right + 1,
    'name ' + label.scrollWidth + 'px in ' + label.clientWidth + 'px, plus right '
      + Math.round(squeezed.right) + ' vs row right ' + Math.round(squeezedRow.right));
  label.textContent = wasName;

  // AND IT OPENS THE RIGHT MENU. Pressed, not act()-ed: this is the one gesture with no api of its
  // own, because the api drives the menus and the button is what a person reaches for.
  plus.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  const plusMenu = [...document.querySelectorAll('.menu-dropdown .menu-item')]
    .map((one) => (one.textContent ?? '').trim());
  check('the plus opens the three kinds and nothing else',
    plusMenu.join(' | ') === 'New Module | New Class Module | New UserForm', plusMenu.join(' | '));
  check('pressing the plus did not also toggle the workbook shut',
    [...document.querySelectorAll('[data-component]')].length > 0,
    rows().join(','));
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));

  // WHAT THE MENU OFFERS, by the class of the thing clicked.
  const moduleMenu = act('treeMenu', { module: 'Module1' });
  check('right-clicking a module opens its menu', moduleMenu.did, moduleMenu.detail);
  check('a module can be removed', moduleMenu.detail.includes('Remove'), moduleMenu.detail);
  act('closeDialogs');
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));

  const sheetMenu = act('treeMenu', { module: 'Sheet1' });
  check('a document module offers no Remove at all',
    sheetMenu.did && !sheetMenu.detail.includes('Remove'), sheetMenu.detail);
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));

  const bookMenu = act('treeMenu', { workbook: 'Book1.xlsm' });
  check('the workbook menu is untouched by all this',
    bookMenu.did && bookMenu.detail.includes('New Module'), bookMenu.detail);
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));

  // NOTHING GOES WITHOUT THE QUESTION.
  act('treeMenu', { module: 'Module1' });
  const chose = act('chooseMenuItem', { label: 'Remove' });
  check('choosing Remove asks first', chose.did && !!box(), chose.detail);
  check('the question names the module', boxText().includes('Module1'), boxText());
  check('Cancel is what has focus', document.activeElement?.textContent === 'Cancel',
    document.activeElement?.textContent ?? 'nothing');
  check('the module is still there while the question is up', rows().includes('Module1'));

  const cancelled = act('answerRemoveConfirm', { answer: 'cancel' });
  await sleep(150);
  check('Cancel keeps the module', cancelled.did && !box() && rows().includes('Module1'),
    rows().join(','));

  act('treeMenu', { module: 'Module1' });
  act('chooseMenuItem', { label: 'Remove' });
  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
  await sleep(150);
  check('Escape keeps the module too', !box() && rows().includes('Module1'), rows().join(','));

  // AND THEN IT GOES, out of both lists.
  const before = tabs().length;
  act('treeMenu', { module: 'Module1' });
  act('chooseMenuItem', { label: 'Remove' });
  const removed = act('answerRemoveConfirm', { answer: 'remove' });
  await sleep(300);

  check('Remove answers that it did', removed.did, removed.detail);
  check('the module leaves the tree', !rows().includes('Module1'), rows().join(','));
  check('its tab goes with it', !tabs().includes('Module1') && tabs().length === before - 1,
    tabs().join(','));
  check('a dirty module does not also raise the save question',
    !document.getElementById('close-confirm-backdrop'));
  check('the rest of the tree is untouched',
    rows().includes('SalesRow') && rows().includes('ThisWorkbook'), rows().join(','));

  const gone = act('treeMenu', { module: 'Module1' });
  check('the tree really republished: there is no row left to right-click', !gone.did, gone.detail);

  return { pass: checks.every((one) => one.ok), checks };
})()`;

function serveDist() {
  const server = createServer(async (request, response) => {
    const asked = request.url === "/" ? "/index.html" : request.url.split("?")[0];
    try {
      const body = await readFile(join(dist, asked));
      response.writeHead(200, { "content-type": TYPES[extname(asked)] ?? "application/octet-stream" });
      response.end(body);
    } catch {
      response.writeHead(404);
      response.end();
    }
  });

  return new Promise((settle) => {
    server.listen(0, "127.0.0.1", () => settle({ server, port: server.address().port }));
  });
}

function findEdge() {
  const candidates = [
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
  ];
  return candidates.find((path) => existsSync(path));
}

function launchEdge(profile) {
  const edge = spawn(findEdge(), [
    "--headless",
    "--remote-debugging-port=0",
    "--remote-allow-origins=*",
    `--user-data-dir=${profile}`,
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-gpu",
    "--disable-extensions",
    "about:blank",
  ]);

  // The browser announces its DevTools endpoint on stderr; port 0 means the announcement
  // is the only place the port exists.
  return new Promise((settle, reject) => {
    let heard = "";
    const listen = (chunk) => {
      heard += chunk.toString();
      const said = heard.match(/DevTools listening on (ws:\/\/\S+)/);
      if (said) {
        settle({ edge, wsUrl: said[1] });
      }
    };
    edge.stderr.on("data", listen);
    edge.on("exit", () => reject(new Error(`the browser exited before announcing DevTools: ${heard}`)));
    setTimeout(() => reject(new Error(`no DevTools announcement in 15s: ${heard}`)), 15000).unref();
  });
}

function connect(wsUrl) {
  const socket = new WebSocket(wsUrl);
  let nextId = 0;
  const waiting = new Map();
  const listening = new Map();

  socket.onmessage = (event) => {
    const message = JSON.parse(event.data);
    if (message.id && waiting.has(message.id)) {
      const { settle, reject } = waiting.get(message.id);
      waiting.delete(message.id);
      if (message.error) {
        reject(new Error(message.error.message));
      } else {
        settle(message.result);
      }
      return;
    }

    const key = `${message.sessionId ?? ""}|${message.method}`;
    if (listening.has(key)) {
      listening.get(key)(message.params);
      listening.delete(key);
    }
  };

  const send = (method, params = {}, sessionId) =>
    new Promise((settle, reject) => {
      const id = ++nextId;
      waiting.set(id, { settle, reject });
      socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    });

  const once = (method, sessionId, timeoutMs = 20000) =>
    new Promise((settle, reject) => {
      listening.set(`${sessionId ?? ""}|${method}`, settle);
      setTimeout(() => reject(new Error(`no ${method} event in ${timeoutMs}ms`)), timeoutMs).unref();
    });

  return new Promise((settle, reject) => {
    socket.onopen = () => settle({ socket, send, once });
    socket.onerror = () => reject(new Error("could not connect to the DevTools socket"));
  });
}

const overall = setTimeout(() => {
  console.error("probe timed out after 90s");
  process.exit(2);
}, 90000);
overall.unref();

const profile = mkdtempSync(join(tmpdir(), "xlide-remove-component-"));
let edge = null;
let server = null;

try {
  if (!findEdge()) {
    throw new Error("Edge is not installed at either standard path");
  }
  if (!existsSync(join(dist, "editor.js"))) {
    throw new Error(`no built page at ${dist}; run node build.mjs in ui\\editor first`);
  }

  const served = await serveDist();
  server = served.server;

  const launched = await launchEdge(profile);
  edge = launched.edge;

  const { socket, send, once } = await connect(launched.wsUrl);

  // Attach to a blank target first and navigate under Page events, so the drive only runs
  // once the document it drives has finished loading.
  const { targetId } = await send("Target.createTarget", { url: "about:blank" });
  const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
  await send("Page.enable", {}, sessionId);
  const loaded = once("Page.loadEventFired", sessionId);
  await send("Page.navigate", { url: `http://127.0.0.1:${served.port}/` }, sessionId);
  await loaded;

  const evaluated = await send(
    "Runtime.evaluate",
    { expression: DRIVE, awaitPromise: true, returnByValue: true },
    sessionId,
  );

  const verdict = evaluated.result?.value;
  if (!verdict || !Array.isArray(verdict.checks)) {
    throw new Error(`the drive returned no verdict: ${JSON.stringify(evaluated)}`);
  }

  /*
   * THE HOVER, driven for real.
   *
   * CSS :hover cannot be provoked from inside the page. Dispatching a mouseover event runs
   * JavaScript listeners and moves no pointer, so the selector never matches and a check written
   * that way passes or fails on something other than what it claims to measure. This moves the
   * actual mouse through the DevTools protocol, which is the only thing the engine treats as a
   * pointer being somewhere.
   */
  const inPage = async (expression) => {
    const answer = await send("Runtime.evaluate", { expression, returnByValue: true }, sessionId);
    return answer.result?.value;
  };

  const where = await inPage(`(() => {
    const row = document.querySelector('.tree-workbook');
    const box = row.getBoundingClientRect();
    // Over the NAME, not over the button: the point of hovering the row is that the whole row
    // arms it, so the measurement has to be taken somewhere the button is not.
    return JSON.stringify({ x: Math.round(box.left + 40), y: Math.round(box.top + box.height / 2) });
  })()`);

  const point = JSON.parse(where);
  const opacityNow = () => inPage(
    "getComputedStyle(document.querySelector('.tree-workbook .tree-add')).opacity");

  await send("Input.dispatchMouseEvent",
    { type: "mouseMoved", x: point.x, y: point.y, buttons: 0 }, sessionId);
  await new Promise((settle) => setTimeout(settle, 300));
  const hovered = await opacityNow();

  await send("Input.dispatchMouseEvent",
    { type: "mouseMoved", x: 5, y: 5, buttons: 0 }, sessionId);
  await new Promise((settle) => setTimeout(settle, 300));
  const left = await opacityNow();

  verdict.checks.push(
    { name: "hovering the workbook NAME shows the plus", ok: hovered === "1", detail: hovered },
    { name: "and it goes again when the pointer leaves", ok: left === "0", detail: left });
  verdict.pass = verdict.checks.every((one) => one.ok);

  console.log(JSON.stringify(verdict));
  process.exitCode = verdict.pass ? 0 : 1;
  socket.close();
} catch (failure) {
  console.log(JSON.stringify({ pass: false, checks: [{ name: "probe ran", ok: false, detail: String(failure.message ?? failure) }] }));
  process.exitCode = 1;
} finally {
  edge?.kill();
  server?.close();
  setTimeout(() => {
    try {
      rmSync(profile, { recursive: true, force: true });
    } catch {
      // The browser can hold its profile a beat past its exit; a leaked temp dir is fine.
    }
    process.exit(process.exitCode ?? 0);
  }, 500);
}
