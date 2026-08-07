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
    log: ({ since, match, max, waitMs } = {}) =>
      call(`log${query({ since, match, max, waitMs })}`, { timeout: (waitMs ?? 0) + 10000 }),
    messages: (last) => call(`messages${query({ last })}`),
    capture: (window) => call(`capture${query({ window })}`, { raw: true, timeout: 20000 }),

    command: (name) => call(`command${query({ name })}`, { method: "POST" }),
    placement: () => call("placement", { method: "POST" }),

    /** Puts the caret where a Run or a Step should act. Scrolling alone will not do it. */
    caret: (line, { module, column, project } = {}) =>
      call(`caret${query({ line, module, column, project })}`, { method: "POST" }),

    /** Native dialogs standing right now. Answers even while the host thread is blocked. */
    dialogs: () => call("dialogs"),

    /**
     * The dialog guard: while on, a NOTICE this door did not raise is cleared as soon as any
     * request notices it, instead of owning the host thread until somebody looks at the screen.
     *
     * Off by default and never turned on by itself, because a dialog a DEVELOPER opened is
     * theirs. A harness run should turn it on: a compile error raised by an experiment stood for
     * six minutes with everything behind it, and nothing in the session could say so, because a
     * VBA modal pumps messages and every other route answers normally while it stands
     * (2026-08-07).
     *
     * `cleared` is the half that matters. A guard that silently swallows an error turns a hang
     * into a mystery, which is the worse trade.
     */
    guard: (on, { forget } = {}) =>
      call(`guard${query({ on: on === undefined ? undefined : String(!!on), forget: forget ? 1 : undefined })}`,
        { method: "POST" }),

    /**
     * Does the project compile, and if not, what does it say?
     *
     * The menu command alone cannot answer: a compile error is a modal, so running it and waiting
     * hangs the thread that raised it. This starts the compile, reads whatever dialog it raises
     * from the outside, and clears it.
     */
    compile: ({ waitMs = 6000 } = {}) =>
      call(`compile${query({ waitMs })}`, { method: "POST", timeout: waitMs + 10000 }),

    /**
     * The documents the surface HOLDS text for — which is not the list of tabs.
     *
     * Text arrives when a module is activated; a tab exists because its pane does. A workspace
     * opened onto eight modules holds one, and two defects came from nothing ever showing the
     * difference (2026-08-07).
     */
    documents: () => call("documents"),

    /**
     * Adds, renames or removes a component — the pieces a fixture is made of.
     *
     * Done from INSIDE, which is the point: reaching in through `Workbook.VBProject` needs "Trust
     * access to the VBA project object model" turned on, and the add-in is already past that gate
     * because the host hands it the VBE. So a fixture can be built with the setting OFF.
     *
     *   await api.component("add", { kind: 1, name: "Helpers" });
     *   await api.writeModule("Helpers", source);
     *   await api.component("remove", { name: "Helpers" });
     *
     * `name` comes back as the component actually ended up named, not as it was asked for: the
     * editor normalises what it dislikes, and refuses some outright — `Circle` belongs to the
     * Excel object library.
     */
    component: (action, { kind, name, newName, project } = {}) =>
      call(`component${query({ action, kind, name, newName, project })}`, { method: "POST" }),

    /**
     * Opens or closes a module's code pane — a TAB, as the strip draws it.
     *
     * `caret` opens one on the way to a line; this is how one goes away. A close goes through the
     * same gate the tab's own X uses, so a module with unwritten edits raises the question;
     * `answer` ("save" or "discard") settles it in advance.
     *
     *   await api.pane("close", { module: "Consumer", answer: "discard" });
     */
    pane: (action, { module, project, answer } = {}) =>
      call(`pane${query({ action, module, project, answer })}`, { method: "POST" }),

    /**
     * The developer's settings — read them, or change ONE without restating the rest.
     *
     * The page's own update takes the whole object, so changing one thing from a harness meant
     * spelling out all seven and getting a default wrong on the way. Named arguments here; the
     * answer is always the settings as they now stand.
     *
     *   await api.settings();                          // read
     *   await api.settings({ treeFollowsEditor: false });
     */
    settings: (changes) =>
      call(`settings${query(changes ?? {})}`, changes ? { method: "POST" } : {}),

    /**
     * What the VBA project actually CONTAINS: every component, its kind, its line count, and
     * whether the editor has a pane open on it.
     *
     * Not the same question as `documents()` (what the surface holds text for) or the tab strip
     * (what has a pane). This is the object model's own answer, read from inside — the question a
     * fixture asks twice, once to build and once to check.
     */
    project: (project) => call(`project${query({ project })}`),

    /**
     * A module's procedures, from the analyzer.
     *
     * For asserting on SHAPE without reading the text back and parsing it a second time, in a
     * second language, with a second set of bugs.
     */
    outline: (module, project) => call(`outline${query({ module, project })}`),

    /**
     * Writes a labelled line into the shim log and answers the offset it landed at.
     *
     * Reading a log for what one step did means finding where that step began, and "scroll up
     * until it looks about right" is how a session ends up reasoning about the wrong three
     * seconds. Mark, act, then `log({ since: mark.at })` — a slice that starts with words you
     * chose is a slice you can be sure is yours.
     *
     *   const mark = await api.mark("renaming Recalculate");
     *   …
     *   const what = await api.log({ since: mark.at });
     */
    mark: (text) => call(`mark${query({ text })}`, { method: "POST" }),

    /** The start-of-session sanity check: right build, everything attached, nothing standing. */
    doctor: () => call("doctor"),

    /** Everything a bug report needs, captured at one moment. */
    journal: (lines) => call(`journal${query({ lines })}`, { timeout: 20000 }),

    /** The requests this door has served, and a script that replays them. */
    history: () => call("history"),

    /**
     * States an expectation and waits for it: stopped, running, surfaceReady, shownModule,
     * noDialogs, localsHas, watchHas, problemFree, responsive. Returns what was seen when it
     * did not hold, which is the half a bare false leaves out.
     */
    assert: (that, { value, timeoutMs = 10000 } = {}) =>
      call(`assert${query({ that, value, timeoutMs })}`, { timeout: timeoutMs + 10000 }),

    /** Recent raw durations for percentile work, rather than a max one outlier owns. */
    perf: () => call("perf"),

    /**
     * Waits for a log line to appear, instead of sleeping and hoping. Returns the matching
     * lines and the offset to continue from, so a test can await an event the way it would
     * await a promise: "the module was written", "the break was entered".
     */
    async waitForLog(match, { since, timeout = 10000, max = 50 } = {}) {
      const from = since ?? (await this.log({ max: 1 })).next;
      return this.log({ since: from, match, max, waitMs: timeout });
    },

    /** Answers a dialog by button caption. Names the button exactly; "Cancel" is usual. */
    dismiss: (button, caption) => call(`dismiss${query({ button, caption })}`, { method: "POST" }),

    /**
     * Runs script in the live page.
     *
     * Prefer `.value` over `.result`. The browser returns a result as JSON, so a script returning
     * a string comes back quoted, and one that builds its answer with JSON.stringify comes back
     * quoted TWICE — and a single parse leaves a string that reads as an object right up until
     * every property of it is undefined, which is a probe reporting false for something that
     * worked (2026-08-07, twice). `value` is already unwrapped.
     */
    eval: (script, surface) =>
      call(`eval${query({ surface })}`, { method: "POST", body: script, timeout: 15000 }),

    /** A page script's answer, already unwrapped. What `eval` is almost always wanted for. */
    async ask(script, surface) {
      const answer = await this.eval(script, surface);
      return answer?.value ?? null;
    },

    /**
     * Waits for a condition IN the page, instead of sleeping a guess.
     *
     * One request, polled inside the page, answering `met` and how long it took — so a probe
     * says what it is waiting for rather than how long it hopes that takes. Every fixed sleep in
     * a probe is a race that has not lost yet: 2500ms was right until a round trip to the host
     * was added to the path it was waiting on, and then it reported the feature broken
     * (2026-08-07).
     *
     *   await api.until("window.xlideBridge.documents.all().length > 1")
     */
    until: (predicate, { waitMs = 10000, surface } = {}) =>
      call(`await${query({ surface, waitMs })}`,
        { method: "POST", body: predicate, timeout: waitMs + 10000 }),

    /**
     * Reloads the page and waits for it to come back, answering with the bundle it is now
     * running. A page change needs no republish and no restart — the bundle is served from a
     * folder on disk — so this plus a copy is the whole page loop. See tools\Update-Page.ps1.
     */
    reload: ({ waitMs = 20000 } = {}) =>
      call(`reload${query({ waitMs })}`, { method: "POST", timeout: waitMs + 10000 }),

    /** The whole visible arrangement: docks, groups, tabs, sizes, open documents. */
    layout: () => call("layout"),

    /** Puts the arrangement back to the default and waits for the page. */
    resetLayout: ({ waitMs = 10000 } = {}) =>
      call(`layout${query({ reset: 1, waitMs })}`, { method: "POST", timeout: waitMs + 10000 }),

    /** What the page said to itself: a ring of console lines, installed at page ready. */
    console: (last) => call(`console${query({ last })}`),

    /** Elements matching a selector: box, classes, hidden, computed styles, and the rules behind them. */
    inspect: (selector, { styles, rules, max } = {}) =>
      call(`inspect${query({ selector, styles, rules: rules ? 1 : undefined, max })}`),

    /**
     * Times a scenario in the page: min, median, p95, max, and the raw samples.
     *
     * `what` is one of tabswitch, layout, type. The raw samples are the point — a median that
     * moved is a fact, and a median that moved because one sample in twenty doubled is a
     * different fact.
     */
    bench: (what, { n } = {}) => call(`bench${query({ what, n })}`, { timeout: 60000 }),

    /**
     * Waits for the editor to be answering again, which is the honest precondition for any
     * assertion after something that might raise a modal. Reports what is in the way.
     */
    async waitUntilResponsive({ timeout = 15000 } = {}) {
      const deadline = Date.now() + timeout;
      while (Date.now() < deadline) {
        try {
          await this.state(2000);
          return true;
        } catch {
          const standing = await this.dialogs().catch(() => null);
          if (standing?.dialogs?.length) {
            throw new Error(
              `blocked by "${standing.dialogs[0].caption}" (buttons: ${standing.dialogs[0].buttons.join(", ")})`);
          }
        }
        await new Promise((resolve) => setTimeout(resolve, 300));
      }
      return false;
    },
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
      case "dialogs": return api.dialogs();
      case "doctor": return api.doctor();
      case "journal": return api.journal(rest[0]);
      case "history": return api.history();
      case "assert": return api.assert(rest[0], { value: rest[1] });
      case "perf": return api.perf();
      case "wait": return api.waitForLog(rest.join(" "));
      case "dismiss": return api.dismiss(rest[0] ?? "Cancel", rest[1]);
      case "eval": return api.eval(rest.join(" "));
      case "immediate": return api.immediate(rest.join(" "));
      case "instances": return (await discover()).map((e) => ({ pid: e.pid, port: e.port, shown: e.state.shownProject }));
      default: throw new Error(`unknown route ${route}`);
    }
  })();

  console.log(JSON.stringify(answer, null, 2));
}
