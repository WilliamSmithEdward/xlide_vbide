// Drives the BUILT editor page (ui/editor/dist) in a headless browser and walks the
// close-confirm flow end to end against the page's demo transport: a dirty close asks,
// Escape and Cancel keep the tab, Don't Save and Save close it, questions queue one at a
// time, and a repeated ask is deduplicated. Prints a JSON verdict {pass, checks} on stdout
// and exits nonzero when any check fails.
//
// No dependencies: Node's own http server serves the dist, Edge provides the browser, and
// the DevTools protocol runs over Node's global WebSocket. Invoked by Test-CloseConfirm.ps1.

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

// What the page is asked to do, and what must be true after each step. Runs inside the
// browser; the return value is the probe's verdict. The queue step presses a second tab
// while the question is up — a user cannot click through the backdrop, but this is exactly
// the shape a Close Others produces, so the synthetic press stands in for that message.
const DRIVE = `(async () => {
  const sleep = (ms) => new Promise((settle) => setTimeout(settle, ms));
  const tabs = () => [...document.querySelectorAll('.tab')].map((tab) => tab.dataset.module);
  const press = (name) => {
    const tab = [...document.querySelectorAll('.tab')].find((one) => one.dataset.module === name);
    const x = tab.querySelector('.tab-close');
    const opts = { bubbles: true, cancelable: true, button: 0, pointerId: 1 };
    x.dispatchEvent(new PointerEvent('pointerdown', opts));
    x.dispatchEvent(new PointerEvent('pointerup', opts));
  };
  const answer = (label) =>
    [...document.querySelectorAll('#close-confirm-buttons button')].find((b) => b.textContent === label);
  const title = () => document.getElementById('close-confirm-title')?.textContent ?? '';
  const asking = () => !!document.getElementById('close-confirm-backdrop');

  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  for (let waited = 0; tabs().length < 2 && waited < 20000; waited += 200) await sleep(200);
  check('the demo opens two dirty tabs', tabs().length === 2, tabs().join(','));

  press('Module2');
  check('a dirty close asks instead of closing', asking() && tabs().length === 2);
  check('the question names the module', title().includes('Module2'), title());
  check('the answers are Save, Do not Save, Cancel',
    [...document.querySelectorAll('#close-confirm-buttons button')].map((b) => b.textContent).join('|')
      === "Save|Don't Save|Cancel");
  check('Save holds the focus', document.activeElement?.textContent === 'Save');

  document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
  check('Escape cancels and the tab stays', !asking() && tabs().length === 2);

  press('Module2');
  press('Module2');
  press('Module1');
  check('a repeated ask is asked once, the other queued',
    document.querySelectorAll('#close-confirm-backdrop').length === 1 && title().includes('Module2'));

  answer("Don't Save").click();
  check('Do not Save closes the tab', tabs().join(',') === 'Module1', tabs().join(','));
  check('the queued question follows', asking() && title().includes('Module1'), title());

  answer('Cancel').click();
  check('Cancel keeps the tab', !asking() && tabs().join(',') === 'Module1');

  const first = [...document.querySelectorAll('.tab')].find((tab) => tab.dataset.module === 'Module1');
  first.dispatchEvent(new MouseEvent('auxclick', { bubbles: true, cancelable: true, button: 1 }));
  check('a middle-click close asks too', asking());

  answer('Save').click();
  check('Save closes the tab', tabs().length === 0 && !asking());

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

    // An event: settle whoever is waiting for this method on this session.
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

const profile = mkdtempSync(join(tmpdir(), "xlide-close-confirm-"));
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
  // once the document it drives has finished loading — evaluating during the navigation
  // ran in a context the load then destroyed.
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
