/*
 * Everything a headless page probe needs except the checks.
 *
 * Four probes grew before this existed, and each carried the same hundred and forty lines: serve
 * ui/editor/dist over a loopback port, launch Edge headless, read the DevTools endpoint off its
 * stderr, speak enough of the protocol to attach and navigate, evaluate one expression, then take
 * the browser and the server down and delete the profile. Four copies of that is four places to
 * fix a flake and three of them get missed.
 *
 * WHAT VARIES IS THE THREE THINGS PASSED IN. The checks, which run in the page. What is served,
 * for the one probe that needs the bundle to be broken. And an optional stage after the checks,
 * for the one that needs a real pointer moved, which cannot be done from inside the page.
 *
 *   import { runPageProbe } from "./page-probe.mjs";
 *   await runPageProbe({ label: "xlide-thing", drive: `(() => { ... })()` });
 *
 * The drive is a JavaScript expression, evaluated in the loaded page, answering
 * {pass, checks:[{name, ok, detail}]}. It is a STRING and it is spliced into a template literal
 * on the way here, so it must not contain a backtick: a nested one ends the literal early and the
 * syntax error lands on a line that looks fine.
 */

import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));

/** Where the built page is. Every probe drives the real bundle, never a fixture of one. */
export const dist = resolve(here, "..", "..", "ui", "editor", "dist");

const TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".ttf": "font/ttf",
  ".json": "application/json",
};

function serveDist(override) {
  const server = createServer(async (request, response) => {
    const asked = request.url === "/" ? "/index.html" : request.url.split("?")[0];

    // A probe can replace one file without touching the developer's published page, which is
    // what lets the boot-failure probe serve a bundle that throws.
    const instead = override?.(asked);
    if (instead !== undefined && instead !== null) {
      response.writeHead(200, { "content-type": TYPES[extname(asked)] ?? "text/javascript" });
      response.end(instead);
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
    // THE VIEWPORT IS STATED, NOT INHERITED. These probes assert on geometry - a 24px hit target,
    // a pane that must be on screen, a splitter that must have room to grow - and the window they
    // measured was whatever headless Edge defaults to, which was 800x600 giving a 756x488
    // viewport (measured 2026-08-30). An unstated input to a geometry check is how a green suite
    // turns red on a browser update with nothing in the diff, and how a check comes to pass only
    // because the window was cramped. Both halves of that were live defects this month: a colour
    // read below the fold of a short editor (#17) and a divider that could not travel its asserted
    // 80px in a narrow card.
    "--window-size=1280,900",
    `--user-data-dir=${profile}`, "--no-first-run", "--no-default-browser-check",
    "--disable-gpu", "--disable-extensions", "about:blank",
  ]);

  // The browser announces its DevTools endpoint on stderr; port 0 means the announcement is the
  // only place the port exists.
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

    // An event: settle whoever is waiting for this method on this session.
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

/**
 * Runs one probe end to end and prints its verdict. Sets the exit code; never throws at the
 * caller, because a probe that dies with a stack looks different from one that fails, and the
 * runner reads both the same way.
 *
 * `after` receives the verdict so far and may push more checks onto `verdict.checks`. It gets
 * `inPage` for reading the page and the raw `send`/`sessionId` for anything the page cannot do to
 * itself, a real pointer being the example: CSS :hover cannot be provoked from script, so a check
 * written with a dispatched mouseover measures something other than what it claims.
 */
export async function runPageProbe({ label, drive, serve, after, path = "/", needs = "editor.js" }) {
  const overall = setTimeout(() => {
    console.error("probe timed out after 90s");
    process.exit(2);
  }, 90000);
  overall.unref();

  const profile = mkdtempSync(join(tmpdir(), `${label}-`));
  let edge = null;
  let server = null;

  try {
    if (!findEdge()) { throw new Error("Edge is not installed at either standard path"); }
    if (!existsSync(join(dist, needs))) {
      throw new Error(`no ${needs} at ${dist}; run node build.mjs in ui\\editor first`);
    }

    const served = await serveDist(serve);
    server = served.server;

    const launched = await launchEdge(profile);
    edge = launched.edge;

    const { socket, send, once } = await connect(launched.wsUrl);

    // Attach to a blank target first and navigate under Page events, so the drive only runs once
    // the document it drives has finished loading - evaluating during the navigation ran in a
    // context the load then destroyed.
    const { targetId } = await send("Target.createTarget", { url: "about:blank" });
    const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
    await send("Page.enable", {}, sessionId);
    const loaded = once("Page.loadEventFired", sessionId);
    // `path` because one probe drives the Object Browser, which is the same bundle asked for a
    // different view. Serving ignores the query; only the navigation carries it.
    await send("Page.navigate", { url: `http://127.0.0.1:${served.port}${path}` }, sessionId);
    await loaded;

    const inPage = async (expression) => {
      const answer = await send(
        "Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true }, sessionId);
      return answer.result?.value;
    };

    const verdict = await inPage(drive);
    if (!verdict || !Array.isArray(verdict.checks)) {
      throw new Error(`the drive returned no verdict: ${JSON.stringify(verdict)}`);
    }

    if (after) {
      await after({ verdict, inPage, send, sessionId });
      verdict.pass = verdict.checks.every((one) => one.ok);
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
      try {
        rmSync(profile, { recursive: true, force: true });
      } catch {
        // The browser can hold its profile a beat past its exit; a leaked temp dir is fine.
      }
      process.exit(process.exitCode ?? 0);
    }, 500);
  }
}
