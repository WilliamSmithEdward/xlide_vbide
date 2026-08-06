/*
 * A client for the shim's debug api: discovery, instance selection, and one method per route.
 *
 * The api exists once per ADD-IN SESSION, which is once per Excel process that has opened the
 * VBE. A developer with several workbooks open usually has ONE process and one api, and the
 * workbooks are told apart by the project argument; a developer running `excel /x` has two
 * processes and two apis, each with its own port, token, and DevTools port. This client
 * covers both: discover() lists every live instance, and open() picks one by workbook, pid,
 * or "the only one".
 *
 * Usage as a module:
 *   import { open, discover } from "./xlide-api.mjs";
 *   const api = await open({ workbook: "scratch.xlsm" });
 *   const { debugMode } = await api.state();
 *
 * Usage from a shell, for a quick look:
 *   node xlide-api.mjs state
 *   node xlide-api.mjs --workbook scratch.xlsm locals
 *   node xlide-api.mjs module CleanModule
 */

import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const DISCOVERY_DIRECTORY = join(process.env.LOCALAPPDATA ?? "", "xlide_vbide");

/** Every instance whose api answers, newest session first. */
export async function discover() {
  let names;
  try {
    names = await readdir(DISCOVERY_DIRECTORY);
  } catch {
    return [];
  }

  const candidates = [];
  for (const name of names) {
    if (!name.startsWith("debug-api-") || !name.endsWith(".json")) {
      continue;
    }

    try {
      const raw = await readFile(join(DISCOVERY_DIRECTORY, name), "utf8");
      candidates.push(JSON.parse(raw));
    } catch {
      // A file being written as it is read, or a corpse mid-sweep.
    }
  }

  // A discovery file outlives a killed Excel, so answering is the only proof of life. The
  // sessions sweep corpses at start-up, but a client should never wait for that.
  const live = await Promise.all(candidates.map(async (entry) => {
    const client = clientFor(entry);
    try {
      const state = await client.state(1500);
      return { ...entry, state, api: client };
    } catch {
      return null;
    }
  }));

  return live
    .filter(Boolean)
    .sort((a, b) => String(b.startedAt ?? "").localeCompare(String(a.startedAt ?? "")));
}

/**
 * One instance: the only live one, or the one matching pid or workbook. Throws with the list
 * when the choice is ambiguous, because guessing which Excel to drive is how a test writes
 * into the wrong workbook.
 */
export async function open({ pid, workbook } = {}) {
  const instances = await discover();
  if (instances.length === 0) {
    throw new Error("no xlide instance is answering; start Excel and open the editor (Debug build only)");
  }

  if (pid !== undefined) {
    const found = instances.find((entry) => entry.pid === Number(pid));
    if (!found) {
      throw new Error(`no instance with pid ${pid}; live pids: ${instances.map((e) => e.pid).join(", ")}`);
    }
    return found.api;
  }

  if (workbook !== undefined) {
    const wanted = String(workbook).toLowerCase();
    const matches = [];
    for (const entry of instances) {
      const windows = await entry.api.windows();
      const holds = windows.windows.some((row) =>
        row.type === 0 && row.caption.toLowerCase().includes(wanted));
      if (holds) {
        matches.push(entry);
      }
    }

    if (matches.length === 0) {
      throw new Error(`no instance holds a workbook matching "${workbook}"`);
    }
    if (matches.length > 1) {
      throw new Error(`several instances hold "${workbook}": pids ${matches.map((e) => e.pid).join(", ")}`);
    }
    return matches[0].api;
  }

  if (instances.length > 1) {
    throw new Error(
      `several instances are live (pids ${instances.map((e) => e.pid).join(", ")}); pass pid or workbook`);
  }

  return instances[0].api;
}

function clientFor(entry) {
  const base = `http://127.0.0.1:${entry.port}/${entry.token}`;

  async function call(route, { method = "GET", body, timeout = 10000, raw = false } = {}) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeout);
    try {
      const response = await fetch(`${base}/${route}`, {
        method,
        body,
        signal: controller.signal,
      });

      if (raw) {
        return Buffer.from(await response.arrayBuffer());
      }

      const answer = await response.json();
      if (answer && typeof answer === "object" && "error" in answer) {
        throw new Error(`${route}: ${answer.error}`);
      }
      return answer;
    } finally {
      clearTimeout(timer);
    }
  }

  const query = (pairs) => {
    const parts = Object.entries(pairs)
      .filter(([, value]) => value !== undefined && value !== null)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`);
    return parts.length > 0 ? `?${parts.join("&")}` : "";
  };

  return {
    pid: entry.pid,
    port: entry.port,
    devtoolsPort: entry.devtoolsPort,
    base,

    state: (timeout) => call("state", { timeout }),
    windows: () => call("windows"),
    stats: () => call("stats"),
    locals: () => call("locals"),
    watches: () => call("watches"),
    problems: (module) => call(`problems${query({ module })}`),
    log: ({ since, match, max } = {}) => call(`log${query({ since, match, max })}`),
    messages: (last) => call(`messages${query({ last })}`),
    capture: (window) => call(`capture${query({ window })}`, { raw: true, timeout: 20000 }),

    command: (name) => call(`command${query({ name })}`, { method: "POST" }),
    placement: () => call("placement", { method: "POST" }),
    immediate: (text) => call(`immediate${query({ text })}`, { method: "POST" }),

    /** state: "on" | "off" | undefined to toggle. Prefer on/off in scripts. */
    breakpoint: (module, line, { project, state } = {}) =>
      call(`breakpoint${query({ module, line, project, state })}`, { method: "POST" }),

    readModule: (name, project) => call(`module${query({ name, project })}`),
    writeModule: (name, text, project) =>
      call(`module${query({ name, project })}`, { method: "POST", body: text }),

    /** Waits for a predicate over state, which is how a harness waits for break mode. */
    async waitFor(predicate, { timeout = 20000, every = 300 } = {}) {
      const deadline = Date.now() + timeout;
      let last;
      while (Date.now() < deadline) {
        last = await this.state();
        if (predicate(last)) {
          return last;
        }
        await new Promise((resolve) => setTimeout(resolve, every));
      }
      throw new Error(`waitFor timed out after ${timeout}ms; last state ${JSON.stringify(last)}`);
    },
  };
}

// A small command line, so a question does not need a script. The guard goes through
// pathToFileURL because a Windows path spells its url with three slashes (file:///F:/...)
// and a hand-built one does not match.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const args = process.argv.slice(2);
  const take = (flag) => {
    const at = args.indexOf(flag);
    if (at < 0) {
      return undefined;
    }
    const value = args[at + 1];
    args.splice(at, 2);
    return value;
  };

  const pid = take("--pid");
  const workbook = take("--workbook");
  const [route, ...rest] = args;

  const api = await open({ pid, workbook });
  const answer = await (async () => {
    switch (route) {
      case undefined:
      case "state": return api.state();
      case "windows": return api.windows();
      case "stats": return api.stats();
      case "locals": return api.locals();
      case "watches": return api.watches();
      case "problems": return api.problems(rest[0]);
      case "log": return api.log({ match: rest[0], max: 200 });
      case "messages": return api.messages(rest[0] ?? 20);
      case "module": return api.readModule(rest[0], rest[1]);
      case "command": return api.command(rest[0]);
      case "immediate": return api.immediate(rest.join(" "));
      case "instances": return (await discover()).map((e) => ({ pid: e.pid, port: e.port, shown: e.state.shownProject }));
      default: throw new Error(`unknown route ${route}`);
    }
  })();

  console.log(JSON.stringify(answer, null, 2));
}
