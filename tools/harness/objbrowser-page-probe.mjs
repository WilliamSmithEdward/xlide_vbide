// Drives the BUILT Object Browser page (ui/editor/dist, ?view=objbrowser) in a headless
// browser against its demo transport and walks the pinned behaviours: the palette page
// boots without the editor shell, the three search scopes act on the right panes, All
// pulls a whole matched group in eagerly, the details pane fills its signature, context,
// and description rows (hiding the empty ones), and the splitter answers the keyboard.
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.
//
// No dependencies: Node's own http server serves the dist, Edge provides the browser, and
// the DevTools protocol runs over Node's global WebSocket. Invoked by Test-ObjectBrowser.ps1.

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
  const rows = (pane) => [...document.querySelectorAll('#objbrowser-' + pane + ' .objbrowser-row')];
  const names = (pane) => rows(pane).map((row) => row.querySelector('.objbrowser-name').textContent);
  const picker = () => document.getElementById('objbrowser-library');
  const detailRow = (part) => document.getElementById('objbrowser-detail-' + part);
  const type = (text) => {
    const box = document.getElementById('objbrowser-search');
    box.value = text;
    box.dispatchEvent(new Event('input', { bubbles: true }));
  };
  const mode = (value) => {
    const pick = document.getElementById('objbrowser-scope');
    pick.value = value;
    pick.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const clickRow = (pane, name) =>
    rows(pane).find((row) => row.querySelector('.objbrowser-name').textContent === name).click();

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  for (let waited = 0; (picker()?.options.length ?? 0) < 3 && waited < 20000; waited += 200) await sleep(200);
  check('the palette page sheds the editor shell', !document.getElementById('shell'));
  check('the demo lists the project and the libraries', picker().options.length === 3,
    [...picker().options].map((option) => option.textContent).join('|'));

  for (let waited = 0; rows('modules').length < 2 && waited < 5000; waited += 100) await sleep(100);
  check('the project scope loads its modules', names('modules').join(',') === 'Module1,ThisWorkbook',
    names('modules').join(','));
  check('the details pane and its splitter exist',
    !!document.getElementById('objbrowser-splitter')
      && !!detailRow('signature') && !!detailRow('context') && !!detailRow('description'));

  // Group: filters the left pane by name and leaves the members pane alone.
  type('module');
  check('Group filters the types pane', names('modules').join(',') === 'Module1', names('modules').join(','));
  check('Group leaves the members pane alone', rows('members').length === 0);

  // All: a group whose own name matches brings its whole membership along, loading it.
  mode('all');
  for (let waited = 0; rows('members').length < 2 && waited < 5000; waited += 100) await sleep(100);
  check('All pulls the whole matched group in', names('members').join(',') === 'Greet,Total',
    names('members').join(','));
  check('spanning members name their group',
    rows('members').every((row) => row.querySelector('.objbrowser-context').textContent === 'Module1'));

  // Object: filters the selected type's members and leaves the list alone.
  mode('object');
  type('');
  clickRow('modules', 'Module1');
  for (let waited = 0; rows('members').length < 2 && waited < 5000; waited += 100) await sleep(100);
  type('gr');
  check('Object filters the selected types members', names('members').join(',') === 'Greet',
    names('members').join(','));
  check('Object leaves the types pane alone', names('modules').length === 2);

  // Details: a project member fills signature and context; the empty description hides.
  clickRow('members', 'Greet');
  check('the signature row carries the declaration',
    detailRow('signature').textContent === 'Public Sub Greet(name As String)',
    detailRow('signature').textContent);
  check('the context row names the module and line',
    detailRow('context').textContent === 'Member of scratch.xlsm.Module1, line 3',
    detailRow('context').textContent);
  check('an empty description row hides', detailRow('description').hidden);

  // Description: a library member that carries one shows it in the third row.
  type('');
  picker().selectedIndex = 1;
  picker().dispatchEvent(new Event('change', { bubbles: true }));
  for (let waited = 0; names('modules').length < 3 && waited < 5000; waited += 100) await sleep(100);
  clickRow('modules', 'Range');
  for (let waited = 0; rows('members').length < 2 && waited < 5000; waited += 100) await sleep(100);
  clickRow('members', 'Address');
  check('a populated description row shows',
    !detailRow('description').hidden && detailRow('description').textContent === 'Returns the address.',
    detailRow('description').textContent);

  // The splitter answers the keyboard.
  const detail = document.getElementById('objbrowser-detail');
  const splitter = document.getElementById('objbrowser-splitter');
  const before = detail.getBoundingClientRect().height;
  splitter.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true, cancelable: true }));
  check('ArrowUp on the splitter grows the details pane', detail.getBoundingClientRect().height > before);

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

const profile = mkdtempSync(join(tmpdir(), "xlide-objbrowser-"));
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

  const { targetId } = await send("Target.createTarget", { url: "about:blank" });
  const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
  await send("Page.enable", {}, sessionId);
  const loaded = once("Page.loadEventFired", sessionId);
  await send("Page.navigate", { url: `http://127.0.0.1:${served.port}/index.html?view=objbrowser` }, sessionId);
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
