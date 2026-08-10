// Pins the engine contract the no-save close depends on (fixed 2026-08-05): the engine's
// live copy of a module - fed by textDocument/didChange - OUTRANKS its seeded copy in
// diagnosis, a reseed alone cannot heal a stale live copy, and the corrective full-source
// didChange the host now sends after a revert (or a Replace All) is what makes the
// Problems pane follow the text. Walks the BUILT engine (engine/dist/engine.cjs) through
// exactly that story over its own named pipe and prints a JSON verdict {pass, checks};
// exits nonzero when any check fails. Invoked by Test-CloseConfirm.ps1.

import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import net from "node:net";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const engine = resolve(here, "..", "..", "engine", "dist", "engine.cjs");

const CLEAN = [
  "Option Explicit",
  "",
  "Public Sub Probe()",
  "    MsgBox 1 + 1",
  "End Sub",
  "",
].join("\r\n");

// The discarded edit of the live report: an undeclared name under Option Explicit.
const DIRTY = [
  "Option Explicit",
  "",
  "Public Sub Probe()",
  "    h = 1",
  "    MsgBox 1 + 1",
  "End Sub",
  "",
].join("\r\n");

const PROJECT = "Probe.xlsm";
const MODULE = "ProbeModule";

function startEngine(pipeName) {
  const child = spawn(process.execPath, [engine, "--pipe", pipeName]);
  return new Promise((settle, reject) => {
    let heard = "";
    child.stdout.on("data", (chunk) => {
      heard += chunk.toString();
      if (heard.includes("listening ")) {
        settle(child);
      }
    });
    child.on("exit", (code) => reject(new Error(`the engine exited (${code}) before listening: ${heard}`)));
    setTimeout(() => reject(new Error(`no listening line in 15s: ${heard}`)), 15000).unref();
  });
}

function connect(pipePath) {
  return new Promise((settle, reject) => {
    const socket = net.connect(pipePath);
    socket.setNoDelay(true);

    let nextId = 0;
    let buffer = "";
    const waiting = new Map();

    socket.on("data", (chunk) => {
      buffer += chunk.toString("utf8");
      let newline = buffer.indexOf("\n");
      while (newline >= 0) {
        const line = buffer.slice(0, newline).trim();
        buffer = buffer.slice(newline + 1);
        newline = buffer.indexOf("\n");
        if (line.length === 0) {
          continue;
        }

        const message = JSON.parse(line);
        if (message.id !== undefined && waiting.has(message.id)) {
          const { settle: deliver, reject: fail } = waiting.get(message.id);
          waiting.delete(message.id);
          if (message.error) {
            fail(new Error(`${message.error.code}: ${message.error.message}`));
          } else {
            deliver(message.result);
          }
        }
      }
    });

    const call = (method, params) =>
      new Promise((deliver, fail) => {
        const id = ++nextId;
        waiting.set(id, { settle: deliver, reject: fail });
        socket.write(`${JSON.stringify({ jsonrpc: "2.0", id, method, params })}\n`);
      });

    const notify = (method, params) => {
      socket.write(`${JSON.stringify({ jsonrpc: "2.0", method, params })}\n`);
    };

    socket.on("connect", () => settle({ socket, call, notify }));
    socket.on("error", (failure) => reject(failure));
  });
}

const overall = setTimeout(() => {
  console.error("probe timed out after 60s");
  process.exit(2);
}, 60000);
overall.unref();

const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail: detail ?? null });

let child = null;
try {
  if (!existsSync(engine)) {
    throw new Error(`no built engine at ${engine}; run node build.mjs in engine first`);
  }

  const pipeName = `xlide-live-probe-${process.pid}`;
  child = await startEngine(pipeName);
  const { socket, call, notify } = await connect(`\\\\.\\pipe\\${pipeName}`);

  const diagnose = async (generation) => {
    const result = await call("textDocument/diagnostics", {
      documentKey: `probe/${MODULE}`,
      projectId: PROJECT,
      generation,
      moduleName: MODULE,
      moduleType: "standard",
    });
    return result.diagnostics.map((one) => one.code);
  };

  await call("initialize", {});

  const seed = (generation, source) =>
    call("project/open", {
      projectId: PROJECT,
      generation,
      modules: [{ moduleName: MODULE, source, type: "standard" }],
    });

  await seed(1, CLEAN);
  const seededClean = await diagnose(1);
  check("the seeded clean module diagnoses clean", seededClean.length === 0, seededClean.join(","));

  notify("textDocument/didChange", { projectId: PROJECT, moduleName: MODULE, source: DIRTY });
  const liveDirty = await diagnose(1);
  check("the live copy outranks the seed", liveDirty.includes("undeclared-variable"), liveDirty.join(","));

  // What a full pass does after a revert: reseed with the reverted module text. Today's
  // contract is that this alone CANNOT heal the live copy - which is why the host must
  // correct it. If this check ever fails, the engine started dropping live copies on seed,
  // and the shim's correction has become redundant: rethink both together.
  await seed(2, CLEAN);
  const reseeded = await diagnose(2);
  check("a reseed alone does not heal a stale live copy (the trap)",
    reseeded.includes("undeclared-variable"), reseeded.join(","));

  // The fix: the corrective full-source didChange the host sends for every rewrite that
  // bypassed the page - the revert and Replace All.
  notify("textDocument/didChange", { projectId: PROJECT, moduleName: MODULE, source: CLEAN });
  const corrected = await diagnose(2);
  check("the corrective didChange heals it (the fix)", corrected.length === 0, corrected.join(","));

  await call("shutdown", {});
  socket.end();
} catch (failure) {
  check("probe ran", false, String(failure.message ?? failure));
} finally {
  const verdict = { pass: checks.length > 0 && checks.every((one) => one.ok), checks };
  console.log(JSON.stringify(verdict));
  process.exitCode = verdict.pass ? 0 : 1;
  child?.kill();
  setTimeout(() => process.exit(process.exitCode ?? 0), 300).unref();
}
