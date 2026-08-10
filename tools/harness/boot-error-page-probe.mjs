// Serves the BUILT page with a bundle that throws on load, and checks that the page's own black
// box caught it.
//
// This is the guard for the gap that cost the most on 2026-08-09. The console ring used to be
// installed by the host once the page reported itself ready; a bundle that throws while its
// modules initialise never reaches ready, so the ring never existed and the `console` route
// answered {"installed": false, "lines": []} at the one moment somebody was asking why the screen
// was blank. The cause was found by reading source instead.
//
// boot.js now installs ahead of the bundle and both records and PUSHES the error. Headless is the
// right place to pin it: provoking it against a real Excel means publishing a broken bundle, and a
// gate that has to break the developer's editor to run is a gate nobody runs.
//
// Prints a JSON verdict {pass, checks} on stdout and exits nonzero when any check fails.

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

/** The fault, in the shape it actually took: a const read during its own dead zone. */
const THROWING_BUNDLE =
  'throw new ReferenceError("Cannot access \'BUILTIN_OBJECTS\' before initialization");\n';

const DRIVE = `(() => {
  const checks = [];
  const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

  const ring = window.__xlideConsole;

  check('the ring exists even though the bundle never ran', Array.isArray(ring),
    ring === undefined ? 'undefined: boot.js did not run, or ran after the bundle' : typeof ring);

  const uncaught = (ring || []).filter((one) => one.indexOf('UNCAUGHT') === 0);
  check('it caught the throw', uncaught.length > 0, (ring || []).join(' | ') || 'the ring is empty');

  check('and it names the error, not just that there was one',
    uncaught.some((one) => one.indexOf('ReferenceError') >= 0 && one.indexOf('BUILTIN_OBJECTS') >= 0),
    uncaught.join(' | '));

  check('with a file and a line to go to',
    uncaught.some((one) => /editor\\.js:\\d+:\\d+/.test(one)), uncaught.join(' | '));

  // The bundle is the thing that failed, so nothing it installs can be relied on here. This is
  // what makes the check meaningful: the surface genuinely is not there.
  check('the surface really did not come up', typeof window.xlideUi === 'undefined',
    typeof window.xlideUi);

  return { pass: checks.every((one) => one.ok), checks };
})()`;

function serveBrokenDist() {
  const server = createServer(async (request, response) => {
    const asked = request.url === "/" ? "/index.html" : request.url.split("?")[0];

    // Everything real except the bundle, which is replaced rather than corrupted on disk: the
    // developer's published page is not this probe's to touch.
    if (asked === "/editor.js") {
      response.writeHead(200, { "content-type": "text/javascript" });
      response.end(THROWING_BUNDLE);
      return;
    }

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
  return [
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
  ].find((path) => existsSync(path));
}

function launchEdge(profile) {
  const edge = spawn(findEdge(), [
    "--headless", "--remote-debugging-port=0", "--remote-allow-origins=*",
    `--user-data-dir=${profile}`, "--no-first-run", "--no-default-browser-check",
    "--disable-gpu", "--disable-extensions", "about:blank",
  ]);

  return new Promise((settle, reject) => {
    let heard = "";
    edge.stderr.on("data", (chunk) => {
      heard += chunk.toString();
      const said = heard.match(/DevTools listening on (ws:\/\/\S+)/);
      if (said) { settle({ edge, wsUrl: said[1] }); }
    });
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
      if (message.error) { reject(new Error(message.error.message)); } else { settle(message.result); }
      return;
    }
    const key = `${message.sessionId ?? ""}|${message.method}`;
    if (listening.has(key)) { listening.get(key)(message.params); listening.delete(key); }
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

const overall = setTimeout(() => { console.error("probe timed out after 90s"); process.exit(2); }, 90000);
overall.unref();

const profile = mkdtempSync(join(tmpdir(), "xlide-boot-error-"));
let edge = null;
let server = null;

try {
  if (!findEdge()) { throw new Error("Edge is not installed at either standard path"); }
  if (!existsSync(join(dist, "boot.js"))) {
    throw new Error(`no boot.js at ${dist}; run node build.mjs in ui\\editor first`);
  }

  const served = await serveBrokenDist();
  server = served.server;

  const launched = await launchEdge(profile);
  edge = launched.edge;

  const { socket, send, once } = await connect(launched.wsUrl);
  const { targetId } = await send("Target.createTarget", { url: "about:blank" });
  const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
  await send("Page.enable", {}, sessionId);
  const loaded = once("Page.loadEventFired", sessionId);
  await send("Page.navigate", { url: `http://127.0.0.1:${served.port}/` }, sessionId);
  await loaded;

  const evaluated = await send(
    "Runtime.evaluate", { expression: DRIVE, awaitPromise: true, returnByValue: true }, sessionId);

  const verdict = evaluated.result?.value;
  if (!verdict || !Array.isArray(verdict.checks)) {
    throw new Error(`the drive returned no verdict: ${JSON.stringify(evaluated)}`);
  }

  console.log(JSON.stringify(verdict));
  process.exitCode = verdict.pass ? 0 : 1;
  socket.close();
} catch (failure) {
  console.log(JSON.stringify({
    pass: false,
    checks: [{ name: "probe ran", ok: false, detail: String(failure.message ?? failure) }],
  }));
  process.exitCode = 1;
} finally {
  edge?.kill();
  server?.close();
  setTimeout(() => {
    try { rmSync(profile, { recursive: true, force: true }); } catch { /* the browser can hold it */ }
    process.exit(process.exitCode ?? 0);
  }, 500);
}
