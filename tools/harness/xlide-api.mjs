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

import { execFile } from "node:child_process";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { promisify } from "node:util";

const DISCOVERY_DIRECTORY = join(process.env.LOCALAPPDATA ?? "", "xlide_vbide");
const run = promisify(execFile);

/**
 * WHY THE HOST DIED, asked of Windows at the moment it stops answering.
 *
 * A probe whose host has gone reports `fetch failed` and `ECONNREFUSED`, which says only that
 * nothing is listening. Whether Excel exited, was closed, or died taking a fault with it, and
 * which library it died in, all live in the Windows event log, and looking there is a step
 * somebody has to think to take.
 *
 * On 2026-08-07 nobody took it for three crashes running. They read as three unrelated
 * instabilities across an afternoon and were nearly filed as "the host is flaky today"; the
 * fourth carried a managed stack naming `ComObject.Finalize`, and that one line explained all
 * four. Hours, for a question the machine could have answered every time.
 *
 * Best effort by design. No event log, no permission, no PowerShell: answer null and let the
 * caller report the plain connection error, because a diagnostic that throws while diagnosing a
 * crash is worse than one that shrugs.
 */
export async function whyDidItDie({ withinMinutes = 5 } = {}) {
  const script = `
    $since = (Get-Date).AddMinutes(-${Number(withinMinutes) || 5})
    $events = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$since} -MaxEvents 40 -ErrorAction SilentlyContinue |
      Where-Object { $_.LevelDisplayName -eq 'Error' -and $_.Message -like '*EXCEL*' }
    if (-not $events) { '[]'; exit }
    $rows = foreach ($e in $events) {
      [pscustomobject]@{
        when     = $e.TimeCreated.ToString('HH:mm:ss')
        provider = $e.ProviderName
        message  = $e.Message
      }
    }
    ConvertTo-Json -InputObject @($rows) -Depth 3 -Compress`;

  try {
    const { stdout } = await run(
      "powershell",
      ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
      { timeout: 20000, maxBuffer: 4 * 1024 * 1024 });

    const rows = JSON.parse(stdout.trim() || "[]");
    if (!Array.isArray(rows) || rows.length === 0) {
      return null;
    }

    return rows.map((row) => {
      const module = /Faulting module name: ([^,]+),/.exec(row.message)?.[1]?.trim() ?? null;
      const code = /Exception code: (0x[0-9a-fA-F]+)/.exec(row.message)?.[1] ?? null;

      // The managed frames, when the runtime got far enough to write them. This is the one that
      // names OUR code rather than whichever library noticed the damage.
      const stack = /Stack:\s*([\s\S]*)$/.exec(row.message)?.[1]
        ?.split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => line.startsWith("at "))
        .slice(0, 8) ?? [];

      return { when: row.when, provider: row.provider, module, code, stack };
    });
  } catch {
    return null;
  }
}

/** Those rows as something worth printing beside a connection error. */
function describeDeath(rows) {
  if (rows === null || rows.length === 0) {
    return "";
  }

  const lines = rows.map((row) => {
    const head = `  ${row.when} ${row.provider}`
      + (row.module ? `: faulted in ${row.module}` : "")
      + (row.code ? ` (${row.code})` : "");
    return row.stack.length > 0 ? `${head}\n${row.stack.map((f) => `      ${f}`).join("\n")}` : head;
  });

  return `\n\nWindows says the host died:\n${lines.join("\n")}`;
}

/** A plain delay. Only ever right for something with no observable condition at all. */
export const wait = (ms) => new Promise((settle) => setTimeout(settle, ms));

/**
 * The reporter every live suite prints through. check(what, ok, detail) logs one line and
 * counts it; done() prints the verdict, lists the failures, and answers the exit code.
 *
 * The verdict line is the ONE spelling the gate parses - "N passed, M failed". Four suites
 * once said "checks, broken" instead, and the gate read the first of them as reporting no
 * verdict at all, so the live half died there on every run since they were wired in
 * (2026-08-12). Nine suites then carried their own copy of this reporter, which is nine
 * places for that lesson to un-learn itself; it lives here now (the audit's B24).
 */
export function reporter() {
  let passed = 0;
  const failures = [];

  const check = (what, ok, detail) => {
    console.log(`  ${ok ? "ok  " : "FAIL"} ${what}${detail ? "  -- " + detail : ""}`);
    if (ok) {
      passed += 1;
    } else {
      failures.push(`${what}${detail ? ` (${detail})` : ""}`);
    }
  };

  const done = () => {
    console.log(`\n${passed} passed, ${failures.length} failed`);
    for (const one of failures) {
      console.log(`  ${one}`);
    }
    return failures.length === 0 ? 0 : 1;
  };

  return { check, done };
}

/**
 * A module created for one suite's run, taken away however the run ends: its pane closed with
 * the edits discarded, then the component removed, both tolerantly - a teardown must not turn
 * a finding into a crash, and half of it must run even when the other half cannot. Three
 * suites carried this pair inline in their finally blocks.
 */
export function scratchModule(api, projectId, name) {
  return {
    async dispose() {
      await api.pane("close", { module: name, project: projectId, answer: "discard" }).catch(() => {});
      await api.component("remove", { name, project: projectId }).catch(() => {});
    },
  };
}

/**
 * POLL FOR THE THING, DO NOT SLEEP A GUESS.
 *
 * Seven suites had grown their own copy of this, identical apart from a default budget, and
 * fifteen had their own `wait` (2026-08-10). Worse than the duplication is what sat beside it:
 * 130 fixed sleeps totalling about 149 seconds, which is most of what a live pass spends. A
 * `wait(3000)` is a guess at how long something takes on the machine that wrote it, and
 * driving-excel.md is blunt about the other half of the cost - every one of them is a race that
 * has not lost yet. Two reported a working feature broken in one afternoon (2026-08-07).
 *
 * `predicate` returns the answer when it is ready and anything falsy while it is not, so the
 * value that satisfied the wait is what comes back:
 *
 *   const tab = await waitFor("the module to open", async () =>
 *     (await api.ui()).workspace.groups[0].tabs.find((t) => t.module === name));
 *
 * A predicate that throws is treated as not-ready-yet, because "the route is not answering
 * because the thing is not there" is the normal shape of waiting for it. The last error is put
 * in the timeout message, or a wait that never comes good says only that it timed out.
 */
export async function waitFor(what, predicate, { budgetMs = 20000, pollMs = 100 } = {}) {
  const deadline = Date.now() + budgetMs;
  let last = null;
  let attempts = 0;
  while (Date.now() < deadline) {
    attempts += 1;
    try {
      const answer = await predicate();
      if (answer) { return answer; }
    } catch (failure) {
      last = failure;
    }
    await wait(pollMs);
  }
  throw new Error(`timed out after ${budgetMs}ms waiting for ${what} (${attempts} attempts)`
    + (last ? `; last error: ${last.message}` : ""));
}

/**
 * WAITS FOR A THING TO STOP MOVING, which is what most "let it catch up" sleeps really wanted.
 *
 * `waitFor` is for something arriving. This is for something SETTLING: an analysis pass can
 * publish a finding and then publish it again a moment later at a different position, so the
 * first answer is not the answer (the developer, 2026-08-08, on exactly that). A fixed
 * `wait(4000)` is the guess that stands in for this, and it is both slower than the settle
 * usually takes and shorter than it sometimes needs.
 *
 * Answers once `read()` has returned the same thing `quiet` times running, compared by JSON. The
 * value is returned, so this reads as an assignment. On timing out it returns the last value
 * rather than throwing: a thing that never stops moving is a finding for the checks below to
 * report, not a reason to abandon the run.
 */
export async function waitUntilStable(read, { quiet = 3, pollMs = 120, budgetMs = 15000 } = {}) {
  const deadline = Date.now() + budgetMs;
  let last = await read();
  let same = 1;

  while (Date.now() < deadline) {
    await wait(pollMs);
    const now = await read();
    same = JSON.stringify(now) === JSON.stringify(last) ? same + 1 : 1;
    last = now;
    if (same >= quiet) { return last; }
  }
  return last;
}

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
    // Ask Windows before blaming the developer for not having started Excel. Most of the time
    // they have, and it died.
    throw new Error(
      "no xlide instance is answering; start Excel and open the editor (Debug build only)"
      + describeDeath(await whyDidItDie()));
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
      }).catch(async (reason) => {
        // A HOST THAT HAS GONE says only `fetch failed` and `ECONNREFUSED`, which is the shape of
        // the failure and not its cause. Windows knows the cause, and asking costs a second at
        // the one moment anybody wants to know. Three crashes on 2026-08-07 were read as
        // unrelated instabilities because nobody thought to ask, and one line from the fourth
        // explained all four.
        throw new Error(`${route}: ${reason?.message ?? reason}` + describeDeath(await whyDidItDie()));
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

    /**
     * The editor's menus with their ids, suppressed ones included.
     *
     *   await api.menus()            the menu bar
     *   await api.menus([8])         the eighth menu's items
     *   await api.menus([900])       the composed xlide menu, and what each position resolves to
     *
     * `suppressed` is the difference between the editor's menu and the product's, which is the
     * only place a suppression can be checked. The ids in VbeMenus were measured by hand once and
     * written into a comment before this existed.
     */
    menus: (path = []) => call(`menus${query({ path: path.join(",") || undefined })}`),

    /**
     * The HOST's own editor, underneath the surface that covers it.
     *
     * Run, Step, Compile and ToggleBreakpoint act on the native ACTIVE CODE PANE and the caret
     * inside it - not on the page. When those disagree, a Run executes where the developer is
     * not looking and a breakpoint lands on the wrong line, and nothing on screen says so.
     *
     * The surface's own belief rides along in the same reply, so the comparison is one call.
     */
    native: ({ text } = {}) => call(`native${query({ text: text ? 1 : undefined })}`),

    /**
     * What the ENGINE is holding for a module, against what the surface holds.
     *
     * Every finding is computed against the engine's copy, and it is maintained incrementally by
     * didChange rather than re-sent whole. When a squiggle is drawn in the wrong place, this is
     * the call that says which side drifted. `text: true` brings both texts back.
     */
    engineSource: (module, { text } = {}) =>
      call(`engine${query({ module, text: text ? 1 : undefined })}`),

    /**
     * EVERY open module's content, host against surface.
     *
     * `inSync()` covers the one on screen; this covers the ones behind it. A background tab
     * holds a copy the developer is not looking at, so a module written from outside while its
     * tab sits behind another goes stale with nothing to notice until it is clicked - and then
     * the developer is the one who notices.
     *
     * A pane the surface holds no text for is reported as `held: false` rather than as a
     * disagreement: not holding a module is a different thing from holding it wrongly.
     */
    async parityAll() {
      const below = await this.native();
      const rows = (below.panes ?? []).map((pane) => ({
        module: pane.module,
        project: pane.project,
        held: pane.surfaceContent !== null,
        agreed: pane.surfaceContent === null || pane.hostContent === pane.surfaceContent,
        hostContent: pane.hostContent,
        surfaceContent: pane.surfaceContent,
      }));

      return {
        agreed: rows.every((one) => one.agreed),
        stale: rows.filter((one) => !one.agreed),
        panes: rows,
      };
    },

    /** True when the native pane, the surface and the page all name the same module. */
    async inSync() {
      const [below, ui] = await Promise.all([this.native(), this.ui()]);
      const page = ui.focus.model ? ui.focus.model.split("/").pop() : null;
      const same = (a, b) => (a ?? "").toLowerCase() === (b ?? "").toLowerCase();

      // THE CONTENT, not a proxy for it. A surface holding an empty document for a module the
      // host has 42 lines of agrees on every name and shows a blank editor, which is how this
      // was found (2026-08-08). Both sides are reduced the same way in the shim - line endings
      // normalised, trailing blanks dropped - so a single changed character registers and the
      // host's CRLF does not.
      //
      // A null native content means there is no pane to compare against, which is not a
      // disagreement; a null surface content against a real pane is.
      const contentAgrees = below.nativeContent === null
        || below.nativeContent === below.surfaceContent;

      return {
        agreed: same(below.activeModule, below.surfaceModule) && same(below.surfaceModule, page)
          && contentAgrees,
        contentAgrees,
        nativeModule: below.activeModule,
        surfaceModule: below.surfaceModule,
        pageModule: page,
        nativeLines: below.nativeLines,
        surfaceLines: below.surfaceLines,
        nativeContent: below.nativeContent,
        surfaceContent: below.surfaceContent,
        nativeCaret: `${below.caretLine}:${below.caretColumn}`,
        pageCaret: `${ui.focus.line}:${ui.focus.column}`,
      };
    },
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
     *
     * CHECK `started` BEFORE `compiled`. There is no positive report to be had - the editor
     * answers a compile with a modal or with silence - so `compiled` means "nothing objected AND
     * the command ran". Without the second half a greyed Compile item read as a clean project,
     * which is the shape of precondition failure that makes everything after it vacuous.
     */
    compile: ({ waitMs = 6000 } = {}) =>
      call(`compile${query({ waitMs })}`, { method: "POST", timeout: waitMs + 10000 }),

    /**
     * The documents the surface HOLDS text for - which is not the list of tabs.
     *
     * Text arrives when a module is activated; a tab exists because its pane does. A workspace
     * opened onto eight modules holds one, and two defects came from nothing ever showing the
     * difference (2026-08-07).
     */
    documents: () => call("documents"),

    /**
     * Adds, renames or removes a component - the pieces a fixture is made of.
     *
     * Done from INSIDE, which is the point: reaching in through `Workbook.VBProject` needs "Trust
     * access to the VBA project object model" turned on, and the add-in is already past that gate
     * because the host hands it the VBE. So a fixture can be built with the setting OFF.
     *
     *   await api.component("add", { kind: 1, name: "Helpers" });
     *   await api.component("add", { kind: "class", name: "Account" });
     *   await api.writeModule("Helpers", source);
     *   await api.component("remove", { name: "Helpers" });
     *
     * `kind` is 1 standard, 2 class, 3 form, or the word for one: "module"/"standard",
     * "class", "form"/"userform". A kind it cannot read is refused rather than defaulted:
     * asking for "class" once handed back a STANDARD module and still said ok, which shows
     * up much later as an analyzer complaining that a Friend member is in the wrong module.
     *
     * `name` comes back as the component actually ended up named, not as it was asked for: the
     * editor normalises what it dislikes, and refuses some outright - `Circle` belongs to the
     * Excel object library.
     */
    component: (action, { kind, name, newName, project } = {}) =>
      call(`component${query({ action, kind, name, newName, project })}`, { method: "POST" }),

    /**
     * A UserForm's design as data: the form's own look, then every control with its identity,
     * geometry (points), container (`parent` names the Frame or Page that holds it), and the
     * first ring of appearance. A property a control does not carry is null, never guessed.
     *
     *   const design = await api.designer("EntryForm");
     *   design.form.insideWidth;               // the canvas area, in points
     *   design.controls.find(c => c.name === "OkButton").caption;
     */
    designer: (module, project) => call(`designer${query({ module, project })}`),

    /**
     * Mutates a form through the MSForms designer object model - the same model the native
     * toolbox calls, which is what keeps a form built here byte-compatible with one built by
     * hand. The three actions are what a form fixture is made of:
     *
     *   await api.designerEdit("add", { module: "EntryForm", type: "commandButton",
     *                                   name: "OkButton", left: 126, top: 200, width: 72, height: 24 });
     *   await api.designerEdit("add", { module: "EntryForm", type: "optionButton",
     *                                   name: "PickA", parent: "Options" });   // into a Frame or a Page
     *   await api.designerEdit("set", { module: "EntryForm", name: "OkButton",
     *                                   property: "Caption", value: "Start", as: "text" });
     *   await api.designerEdit("set", { module: "EntryForm", property: "Caption", value: "Entry" });  // the form
     *   await api.designerEdit("remove", { module: "EntryForm", name: "OkButton" });
     *
     * `type` is a toolbox name (commandButton, label, textBox, comboBox, listBox, checkBox,
     * optionButton, toggleButton, frame, tabStrip, multiPage, scrollBar, spinButton, image) or
     * any full ProgID. `set` answers what the property READS BACK, and one level of dotting
     * reaches an object-valued property's member: `property: "Font.Bold"`. `as` (text, number,
     * flag) overrides the value heuristic - a caption of "123" wants as=text.
     */
    designerEdit: (action, { module, project, type, name, parent, left, top, width, height, property, value, as } = {}) =>
      call(`designer${query({ action, module, project, type, name, parent, left, top, width, height, property, value, as })}`,
        { method: "POST" }),

    /**
     * The form as markup text: the same walk `designer` answers, projected through the form
     * markup language (docs/userform-designer.md, the markup layer). What the markup tab will
     * hold, and what a diff against source control would carry.
     */
    designerMarkup: (module, project) =>
      call(`designer${query({ module, project, format: "markup" })}`).then((answer) => answer.markup),

    /**
     * What a control of a KIND holds untouched: the inventory the markup projection compares
     * against to print only what a developer changed. Measured from a bare instance of the
     * coclass MSForms registers - no form, no workbook, nothing on screen - so it cannot drift
     * from the MSForms this machine actually has. A kind with no ProgID of ours, or one whose
     * coclass will not come up, answers zero properties rather than a guess.
     */
    controlDefaults: (type) => call(`defaults${query({ type })}`),

    /**
     * Applies a markup document to the live form as a NAME-KEYED DIFF: controls only in the
     * markup are added, controls only in the model are removed, matched controls take their
     * header geometry, caption and property lines - and an unspoken property is never
     * touched. A document that does not parse applies NOTHING and the refusal carries the
     * line. Answers {ok, added, removed, set, detail, notes}; `notes` lists the rows the diff
     * deliberately left alone (pages, unknown kinds without a ProgId line).
     */
    applyMarkup: (module, markup, project) =>
      call(`designer${query({ module, project, action: "applyMarkup" })}`, { method: "POST", body: markup }),

    /**
     * Opens or closes a module's code pane - a TAB, as the strip draws it.
     *
     * `caret` opens one on the way to a line; this is how one goes away. A close goes through the
     * same gate the tab's own X uses, so a module with unwritten edits raises the question;
     * `answer` ("save" or "discard") settles it in advance.
     *
     *   await api.pane("close", { module: "Consumer", answer: "discard" });
     *
     * A CLOSE ANSWERS `{closed, detail, awaiting}` and not a bare ok. Three things left the tab
     * where it was while replying success: a save the workbook refused, a discard whose revert
     * write the module refused, and a confirm now standing on screen. That last one is
     * `awaiting: "confirm"`, which is neither a failure nor a close - answer it with
     * `act("answerCloseConfirm", { answer })` rather than polling to find out what happened.
     */
    pane: (action, { module, project, answer, face } = {}) =>
      call(`pane${query({ action, module, project, answer, face })}`, { method: "POST" }),

    /**
     * Puts the Object Browser palette away, the way its own close box does: hidden with its
     * state intact, so the next summons (the objectBrowser command) presents the same page.
     * Answers { did, detail, visible }; did is false when no palette exists yet.
     */
    paletteHide: () => call(`palette${query({ action: "hide" })}`, { method: "POST" }),

    /**
     * The editor window itself: "close" posts the developer's own X click (SC_CLOSE through
     * the pump), so the reply comes back BEFORE the window goes - poll state().frameVisible
     * for the outcome, like every posted effect on this door. "show" is synchronous and its
     * reply is the outcome.
     */
    frame: (action) => call(`frame${query({ action })}`, { method: "POST" }),

    /**
     * The session lifecycle. "cancelledShutdown" runs the real OnBeginShutdown without a
     * process exit, so the session stops and the watchdog revives it - the exact path a
     * developer meets when they cancel Excel's save prompt. This reply arrives BEFORE the
     * teardown (it must, or it would ride the DebugServer that Stop disposes); AFTER it, this
     * client's port goes dead. Reconnect with a fresh discover()/open({pid}) once the revived
     * session has rewritten the discovery file with a new port and startedAt.
     */
    session: (action) => call(`session${query({ action })}`, { method: "POST" }),

    /**
     * The developer's settings - read them, or change ONE without restating the rest.
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
     * (what has a pane). This is the object model's own answer, read from inside - the question a
     * fixture asks twice, once to build and once to check.
     */
    project: (project) => call(`project${query({ project })}`),

    /**
     * EVERY open workbook: its name, its id, how many components it holds, and whether the
     * surface is showing one of them.
     *
     * `project()` answers about ONE, the one named or the active one, so with two workbooks open
     * there was no way to discover the other's name from the host at all. A probe either knew it
     * in advance or read the page's tree, which is the surface's view rather than the object
     * model's. The language suite failed exactly there: it asked `project()`, got whichever
     * workbook happened to be active, and looked for its own fixture's module in the other one.
     *
     *   const mine = (await api.projects()).projects
     *     .find((one) => one.project.toLowerCase().startsWith("language"));
     */
    projects: () => call("projects"),

    /** The open workbook holding a module, by name. Null when no open workbook has one. */
    async projectHolding(moduleName) {
      const wanted = moduleName.toLowerCase();

      for (const row of (await this.projects()).projects ?? []) {
        const inside = await this.project(row.project).catch(() => null);
        if ((inside?.components ?? []).some((one) => (one.name ?? "").toLowerCase() === wanted)) {
          return row;
        }
      }

      return null;
    },

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
     * seconds. Mark, act, then `log({ since: mark.at })` - a slice that starts with words you
     * chose is a slice you can be sure is yours.
     *
     *   const mark = await api.mark("renaming Recalculate");
     *   ...
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

    /**
     * Forces a collection and waits for the finalizers, which makes a leaked COM wrapper fail
     * NOW instead of minutes later in a stack that names nothing.
     *
     * A wrapper nothing disposed is released by the finalizer thread, and for the editor's
     * apartment-threaded objects that is an access violation the runtime cannot throw: it
     * FailFasts and takes Excel with it. Ordinarily that lands whenever a collection happens to
     * run, which is why five crashes across two days were read as unrelated.
     *
     * So: run an operation, call this, and if the call does not come back, THAT operation left a
     * wrapper behind. It is a bisecting tool, not a health check - the count still comes from
     * `stats.comWrappersLive`.
     */
    drainFinalizers: () => call("drainfinalizers", { timeout: 30000 }),

    /** Recent raw durations for percentile work, rather than a max one outlier owns. */
    /**
     * Everything about how fast this session is.
     *
     * `placementMs`/`marshalMs` are the host's raw samples. `engine` is the one that matters and
     * is new: per analyzer method, split into WAIT (queued behind another call) and CALL (the
     * round trip). One request is served at a time, so a diagnostics pass over a big module
     * delays every keystroke's completion behind it, and a combined figure blames completions.
     *
     * `reset: true` forgets the engine figures first, so an experiment measures what it provokes
     * rather than everything since the editor opened.
     */
    perf: ({ reset } = {}) => call(`perf${query({ reset: reset ? 1 : undefined })}`),

    /** perf(), ranked and printable: the analyzer methods this session actually spent time in. */
    async engineCosts({ reset } = {}) {
      const answer = await this.perf({ reset });
      return (answer.engine ?? []).map((row) => ({
        method: row.method,
        calls: row.calls,
        totalMs: row.waitTotalMs + row.callTotalMs,
        waitMs: row.waitTotalMs,
        callMs: row.callTotalMs,
        medianMs: row.medianMs,
        p95Ms: row.p95Ms,
        worstMs: row.waitMaxMs + row.callMaxMs,
        refused: row.refused,
      }));
    },

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
     * The captions of every RUNNING form, or - with action "close" - closes one the way its
     * X would. Off the host thread on purpose: a modally running form holds the host thread
     * inside the Run command, and this is the verb that watches it stand and takes it down.
     */
    userforms: (action, caption) => call(`userform${query({ action, caption })}`,
      action === "close" ? { method: "POST" } : undefined),

    /**
     * Runs script in the live page.
     *
     * Prefer `.value` over `.result`. The browser returns a result as JSON, so a script returning
     * a string comes back quoted, and one that builds its answer with JSON.stringify comes back
     * quoted TWICE - and a single parse leaves a string that reads as an object right up until
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
     * One request, polled inside the page, answering `met` and how long it took - so a probe
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
     * running. A page change needs no republish and no restart - the bundle is served from a
     * folder on disk - so this plus a copy is the whole page loop. See tools\page.ps1.
     */
    reload: ({ waitMs = 20000 } = {}) =>
      call(`reload${query({ waitMs })}`, { method: "POST", timeout: waitMs + 10000 }),

    /** The whole visible arrangement: docks, groups, tabs, sizes, open documents. */
    layout: () => call("layout"),

    /**
     * The surface as the PAGE describes it, decoded: `.value` is the snapshot.
     *
     * Tabs with the labels the strip drew and the MRU order a close falls through, the tree's
     * expansion and unfolded module, which panes and dialogs are up, what has not arrived yet,
     * and where the caret is. Reach for this before writing a DOM script: a scraped row cannot
     * tell "collapsed" from "rendered wrong", and the render being stale is the defect worth
     * catching.
     */
    async ui({ line, column, word } = {}) {
      const answer = await call(`ui${query({ line, column, word })}`, { timeout: 10000 });
      return answer.value ?? answer;
    },

    /**
     * What is AT a word or a position: the colour actually painted and the squiggles covering it.
     *
     * The colour is read off the rendered span rather than derived from a token type and a theme
     * map, because "is this word the wrong colour" is a question about the pixel, and every step
     * between the tokeniser and the pixel can be the wrong one. Only rendered lines have spans,
     * so a position off-screen answers a null colour rather than a guess.
     */
    async at(where) {
      const answer = await this.ui(typeof where === "string" ? { word: where } : where);
      return answer.at;
    },

    /**
     * Drives the surface through the methods a click reaches.
     *
     * `closeActive`, `activate`, `cycleTab`, `split`, `expandWorkbook`, `unfoldModule`,
     * `treeMenu`, `chooseMenuItem`, `answerRemoveConfirm`, `settings`, `sponsors`, `closeDialogs`,
     * `focusEditor`, `search`, `dock`, `bookmark`, `format`, `undo`, `editorAction`, and the
     * language ones: `hover`, `completions`, `signature`, `quickFixes`, `definition`,
     * `references`, `rename`.
     * `act("actions")` answers the live list, which is the one that cannot be out of date. Answers
     * {did, detail}; `did: false` means the page declined, which is an answer and not a failure.
     *
     * REMOVING A COMPONENT takes three of them, and that is the point: `treeMenu` opens the row's
     * menu, `chooseMenuItem` with `Remove` raises the question, `answerRemoveConfirm` answers it.
     * Nothing on this path deletes anything until the last call. The `component` route's
     * `action=remove` is the one that skips the question, and it is a fixture primitive.
     *
     * TWO WAYS TO SEND A KEY, and they are not interchangeable:
     *
     *   act("press", { key: "Enter" })     the EDITOR types it. Enter, Tab, Backspace, Delete,
     *                                      Escape. This is what runs the enter rules, so smart
     *                                      Enter, auto-indent and comment continuation are only
     *                                      reachable through here.
     *   act("key", { code: "KeyW", ctrl }) a synthetic KeyboardEvent at the document, for the
     *                                      chords this product binds there. Monaco does not act
     *                                      on synthesised events, so this cannot type.
     *
     * `press` arrived on 2026-08-09. Until then `type` was the only way in, and it INSERTS a
     * string: Monaco applies its enter rules to a newline typed as one character and not to one
     * arriving inside a longer string, so every Enter behaviour was live-untested.
     *
     * Synthesising events instead is how a working feature gets reported broken: the tab close
     * box arms at pointerdown and fires at pointerup, so `.click()` on it does nothing at all.
     */
    async act(name, args = {}) {
      // `do` is the route's action selector, so an argument of that name would overwrite it and
      // the door would answer "unknown action <the argument's value>". Refused loudly rather
      // than silently: it cost a run of bookmark tests that all failed for this and not for
      // anything about bookmarks (2026-08-08).
      if (Object.hasOwn(args, "do")) {
        throw new Error(`act(${name}): "do" is reserved for the action name; rename that argument`);
      }

      const answer = await call(`act${query({ do: name, ...args })}`, { method: "POST", timeout: 15000 });
      return answer.value ?? answer;
    },

    /**
     * Times what a person WAITS for, across the boundary. `pagecall` is the only scenario:
     * anything whose effect is DELIVERED by the host thread cannot be observed from inside a
     * route body, which is on it. For the cost of reaching the host thread read perf().marshalMs.
     *
     * `bench` times the page's own work and `perf` reports the host's; both have looked healthy
     * while the surface felt slow, because the cost was in the crossing that neither measured.
     * Always read `pagecall` alongside the rest: it is the floor every other figure contains.
     */
    trip: (what, { n } = {}) => call(`trip${query({ what, n })}`, { timeout: 120000 }),

    /**
     * Go To Line, timed end to end: the host moves the caret, the page agrees where it is.
     *
     * Measured HERE and not in the shim, and that is not a stylistic choice. A route body runs
     * on the host thread, and the caret is delivered to the page by a message that same thread
     * pumps - so a route that posts and then waits to see the effect waits forever. Written that
     * way it reported the caret never moving, through four seconds a sample; the identical
     * sequence across requests lands on the first poll (2026-08-07). Anything that observes a
     * POSTED effect belongs on this side of the door.
     *
     * Answers the same shape `trip` does, so the two read together.
     */
    /**
     * A language feature timed INSIDE the page: what the developer actually waits for.
     *
     * `tripFeature` times the same thing from out here and therefore carries the door's promise
     * floor in every sample - which is most of the figure, and hid a whole scaling curve behind
     * a flat line until it was noticed. This runs the provider n times in the page and brings
     * back the distribution, so the door is paid once for the run and appears in none of the
     * numbers.
     *
     * Use this for "is the feature fast". Use `tripFeature` for "is the DOOR fast", which is a
     * question about the harness.
     */
    async timeFeature(what, where, { n = 10 } = {}) {
      const answer = await this.act("timeFeature", { what, n, ...where });
      return answer.data ?? answer;
    },

    async tripFeature(what, where, { n = 10 } = {}) {
      const samples = [];
      let lastDetail = "";

      for (let run = 0; run < n; run++) {
        const began = Date.now();
        const answer = await this.act(what, where);
        samples.push(Date.now() - began);
        lastDetail = answer.detail;
      }

      const ordered = [...samples].sort((a, b) => a - b);
      return {
        what,
        runs: ordered.length,
        minMs: ordered[0],
        medianMs: ordered[ordered.length >> 1],
        p95Ms: ordered[Math.min(ordered.length - 1, Math.floor(ordered.length * 0.95))],
        maxMs: ordered[ordered.length - 1],
        samplesMs: samples,
        detail: lastDetail,
      };
    },

    async tripCaret({ n = 5, lines = [1, 2, 3, 4, 5] } = {}) {
      const before = await this.ui();
      if (!before.focus.host) {
        throw new Error("no group is showing the host-active module, so the caret has nowhere observable to land");
      }

      const samples = [];
      for (let run = 0; run < n; run++) {
        // Never the line it is already on: that sample would measure nothing and pass.
        const line = lines[run % lines.length] === (await this.ui()).focus.host?.line
          ? lines[(run + 1) % lines.length]
          : lines[run % lines.length];

        const began = Date.now();
        await this.caret(line);

        let landed = false;
        while (Date.now() - began < 4000 && !landed) {
          landed = (await this.ui()).focus.host?.line === line;
        }

        if (!landed) {
          throw new Error(`the caret never reached line ${line}`);
        }

        samples.push(Date.now() - began);
      }

      const ordered = [...samples].sort((a, b) => a - b);
      return {
        what: "caret",
        runs: ordered.length,
        minMs: ordered[0],
        medianMs: ordered[ordered.length >> 1],
        p95Ms: ordered[Math.min(ordered.length - 1, Math.floor(ordered.length * 0.95))],
        maxMs: ordered[ordered.length - 1],
        samplesMs: samples,
        detail: "host caret set, to the page agreeing where it is; measured across requests",
      };
    },

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
     * `what` is one of tabswitch, layout, type. The raw samples are the point - a median that
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
    /**
     * The Immediate window: evaluate a line, or read what the window says.
     *
     * ANSWERS THE OUTCOME, not the request. It used to post the line and return `{ran: true}`
     * without waiting, so a caller learned that an evaluation had been asked for and nothing
     * else; what the expression came to went only to the page. That is why this panel had a
     * route and no suite.
     *
     *   await api.immediate("?1+1");       // -> { ran: true, text: "2", failed: false }
     *   await api.immediate("?Nonsense");  // -> { failed: true, text: the error }
     *   await api.immediate();             // -> the whole window as it stands
     *
     * `ran: false` means the evaluation did not finish inside the wait. The line still went in.
     */
    immediate: (text) =>
      call(`immediate${query({ text })}`, { method: "POST", timeout: 20000 }),

    /**
     * Every breakpoint the session is holding, and the mode it is in.
     *
     * There was a way to SET one from the day this door landed and no way to ask what is set,
     * which makes a debugger assertion a matter of remembering what the test did.
     */
    /**
     * What is recorded, per module AND per workbook.
     *
     * Each row carries `project`, because two open workbooks can each hold a module of the same
     * name and the record used to be keyed by the name alone: a breakpoint set in one was
     * reported against the other, and a run that should have stopped did not (fixed 2026-08-08).
     * `breakpointsIn(project)` is the filtered view a two-workbook test wants.
     */
    breakpoints: () => call("breakpoints"),

    /** The rows belonging to one workbook, matched on its file name however it is cased. */
    async breakpointsIn(project) {
      const answer = await this.breakpoints();
      const wanted = String(project ?? "").toLowerCase();
      return (answer.breakpoints ?? []).filter((row) =>
        !wanted || String(row.project ?? "").toLowerCase() === wanted);
    },

    /**
     * Puts the last rename back - every module it touched, and the component's old name.
     *
     * The editor's own undo cannot: a rename edits several modules and the undo stack is per
     * model, so Ctrl+Z reverses one module's share and leaves the rest renamed.
     */
    undoRename: () => call("undoRename", { method: "POST" }),

    /**
     * Types into the editor the way a person does.
     *
     * Through the editor's own keyboard pipeline, so the behaviour that only happens WHILE typing
     * is what gets tested: smart Enter, comment continuation, auto-indent. Setting the text
     * instead goes around every handler that makes typing feel like anything, so a probe that
     * sets text is testing nothing this product does. A `\n` is sent as a real Enter.
     *
     *   await api.type("' a comment\n");     // the next line should continue the apostrophe
     */
    type: (text, { waitMs = 8000 } = {}) =>
      call(`type${query({ waitMs })}`, { method: "POST", body: text, timeout: waitMs + 8000 }),

    /** state: "on" | "off" | undefined to toggle. Prefer on/off in scripts. */
    breakpoint: (module, line, { project, state } = {}) =>
      call(`breakpoint${query({ module, line, project, state })}`, { method: "POST" }),

    /**
     * A module's text. `live: true` reads the SURFACE's copy rather than the workbook's.
     *
     * They differ while the developer has typed and the write-back has not fired, which is the
     * window every typing behaviour lives in: smart Enter, comment continuation and auto-indent
     * all produce text that exists only in the editor until it is written.
     */
    readModule: (name, project, { live } = {}) =>
      call(`module${query({ name, project, live: live ? 1 : undefined })}`),
    /**
     * Writes a module. The client waits longer for a BIG one, because the editor takes longer to
     * take it: 65,000 lines was accepted after 17.4 seconds, all of it the editor's own parse
     * rather than this product's write, which is two COM calls at any size (measured 2026-08-08).
     *
     * The DOOR still gives up at three seconds and answers "the host thread did not answer in
     * time" - that budget is what lets it report a standing dialog instead of hanging, and it is
     * not raised here. The write goes on regardless, so a caller writing a large module waits for
     * the LINE COUNT to arrive rather than for this call to return. `build-fixture.mjs` does.
     *
     * THROWS WHEN THE WRITE DID NOT HAPPEN. A module that is not there, or text the editor
     * refuses, used to answer exactly like a write that landed. Reading the count back is still
     * the right habit for a big write, because that one outlives the door's budget and its
     * complaint arrives too late to be the reply (2026-08-09).
     */
    writeModule: (name, text, project) =>
      call(`module${query({ name, project })}`, {
        method: "POST",
        body: text,
        timeout: Math.max(10000, Math.ceil(text.length / 1000) * 250),
      }),

    /**
     * What an import or an export WOULD do, without doing any of it.
     *
     * The plan is the same object the import/export dialog draws, from the same service, so a row
     * read here is the row shown there. Every row carries its status, whether it is ticked, a
     * side-by-side comparison with the attribute headers hidden, and the same comparison with them
     * left in.
     *
     *   const plan = await api.syncPlan("export", { folder: "C:\\src\\modules" });
     *   plan.items.map(i => `${i.status} ${i.file}`);
     *
     * `folder` may be left out once a project has been synced once: it remembers.
     */
    syncPlan: (direction, { project, folder, mode } = {}) =>
      call(`sync${query({ direction, project, folder, mode })}`, { timeout: 30000 }),

    /**
     * Carries a plan out. This is the Apply button, and it leaves the project in exactly the state
     * the button leaves it in, because the same service does the work either way.
     *
     * `ids` names the rows to carry out, which is what the dialog sends after the developer has
     * ticked and unticked. Leave it out and `select` decides: "checked" (the default) takes the
     * rows the plan itself ticked, "all" takes every row it offered.
     */
    syncApply: (direction, { project, folder, mode, ids, select } = {}) =>
      call(`sync${query({ action: "apply", direction, project, folder, mode, select })}`, {
        method: "POST",
        body: (ids ?? []).join("\n"),
        timeout: 60000,
      }),

    /** The folder and modes a project remembers. Naming any of them writes it. */
    syncSettings: ({ project, folder, exportMode, importMode } = {}) =>
      call(`sync${query({ action: "settings", project, folder, exportMode, importMode })}`),

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

  // BEFORE picking one, because this is the verb for when there are several.
  //
  // `open` refuses to guess when more than one instance is live, which is right - and it ran
  // first for every verb including this one, so the question "which instances are there" could
  // only be answered while the answer was "one". Exactly when it was useless (2026-08-08).
  if (route === "instances") {
    const found = await discover();
    console.log(JSON.stringify(
      found.map((entry) => ({
        pid: entry.pid,
        port: entry.port,
        shown: entry.state?.shownProject ?? null,
        module: entry.state?.shownModule ?? null,
      })),
      null,
      2));
    process.exit(0);
  }

  const api = await open({ pid, workbook });
  const answer = await (async () => {
    switch (route) {
      case undefined:
      case "state": return api.state();
      case "windows": return api.windows();

      // THE ROUTES A WORKBOOK QUESTION NEEDS, which the command line did not carry.
      //
      // The client had them and the CLI did not, so every question about which workbook a thing
      // belongs to had to be asked from a throwaway script - exactly the habit the standing
      // instruction says not to build. `native` and `projects` are the two that answer "which
      // workbook", and `pane` is how a module is opened in a named one (2026-08-08).
      //
      //   xlide-api.mjs native            xlide-api.mjs projects
      //   xlide-api.mjs pane open Sheet1 "VBAProject 01"
      //   xlide-api.mjs pane open EntryForm face design      the FORM's designer tab
      //   xlide-api.mjs outline Runner    xlide-api.mjs caret 12 Runner
      case "native": return api.native({ text: rest[0] === "text" });
      case "projects": return api.projects();
      case "console": return api.console(Number(rest[0] ?? 20));
      // The face is named rather than counted: it arrived after the positionals were spent,
      // and `pane open EntryForm face design` reads better than a fourth silent slot.
      case "pane": return api.pane(rest[0] ?? "open", rest[2] === "face"
        ? { module: rest[1], face: rest[3] }
        : { module: rest[1], project: rest[2], answer: rest[3] });
      // xlide-api.mjs defaults commandButton      what one holds untouched
      case "defaults": return api.controlDefaults(rest[0]);
      case "outline": return api.outline(rest[0], rest[1]);
      case "caret": return api.caret(Number(rest[0] ?? 1), { module: rest[1], project: rest[2] });
      case "breakpoints": return api.breakpoints();
      case "documents": return api.documents();
      case "settings": return api.settings(rest[0] ? JSON.parse(rest[0]) : undefined);
      case "layout": return api.layout();
      case "stats": return api.stats();
      case "locals": return api.locals();
      case "watches": return api.watches();
      case "problems": return api.problems(rest[0]);
      /*
       * THE TAIL, because that is what a question about the last thing that happened wants.
       *
       * This asked for the first 200 lines from the start of the session, so every look at "what
       * just happened" answered with the session's opening moments instead. Reading a timeline
       * of a gesture performed seconds ago meant writing a script to page through offsets, which
       * is the habit the standing instruction says not to build (2026-08-08).
       *
       *   xlide-api.mjs log                  the last 40 lines
       *   xlide-api.mjs log destroy          the last 40 mentioning "destroy"
       *   xlide-api.mjs log destroy 200      the last 200 of them
       */
      case "log": {
        const wanted = Number(rest[1] ?? 40);
        const all = await api.log({ match: rest[0], max: 20000 });
        return { lines: all.lines.slice(-wanted), next: all.next, of: all.lines.length };
      }
      case "messages": return api.messages(rest[0] ?? 20);
      case "module": return api.readModule(rest[0], rest[1]);
      case "command": return api.command(rest[0]);
      case "dialogs": return api.dialogs();
      case "doctor": return api.doctor();
      case "journal": return api.journal(rest[0]);
      case "history": return api.history();
      case "assert": return api.assert(rest[0], { value: rest[1] });
      case "perf": return api.perf({ reset: rest[0] === "reset" });
      case "engine": return api.engineCosts({ reset: rest[0] === "reset" });
      case "wait": return api.waitForLog(rest.join(" "));
      case "dismiss": return api.dismiss(rest[0] ?? "Cancel", rest[1]);
      case "userforms": return api.userforms(rest[0], rest[1]);
      case "palette": return api.paletteHide();
      case "frame": return api.frame(rest[0] ?? "show");
      case "session": return api.session(rest[0] ?? "cancelledShutdown");
      case "eval": return api.eval(rest.join(" "));
      case "ui": return api.ui();
      // node xlide-api.mjs act closeActive / act expandWorkbook workbook TwinFixture.xlsm
      case "act": return api.act(rest[0], Object.fromEntries(
        rest.slice(1).reduce((pairs, value, at, all) =>
          at % 2 === 0 ? [...pairs, [value, all[at + 1]]] : pairs, [])));
      case "trip": return api.trip(rest[0] ?? "pagecall", { n: rest[1] });
      case "immediate": return api.immediate(rest.join(" "));
      // sync export C:\out [mode]     what an export would do
      // sync import C:\in  [mode]     what an import would do
      case "sync": return api.syncPlan(rest[0] ?? "export", { folder: rest[1], mode: rest[2] });
      case "syncApply": return api.syncApply(rest[0] ?? "export", { folder: rest[1], mode: rest[2], select: rest[3] ?? "checked" });
      case "syncSettings": return api.syncSettings({ folder: rest[0], exportMode: rest[1], importMode: rest[2] });
      case "instances": return (await discover()).map((e) => ({ pid: e.pid, port: e.port, shown: e.state.shownProject }));
      // designer EntryForm                        the control tree
      // designer EntryForm markup                 the same form as markup text
      // designer EntryForm add commandButton OkButton     and the other mutations by pairs:
      // designer EntryForm set name OkButton property Caption value Start as text
      case "designer": {
        if (rest[1] === "markup") {
          console.log(await api.designerMarkup(rest[0], rest[2]));
          process.exit(0);
        }

        return rest.length <= 1
        ? api.designer(rest[0])
        : api.designerEdit(rest[1], rest[1] === "add"
          ? { module: rest[0], type: rest[2], name: rest[3], ...Object.fromEntries(
            rest.slice(4).reduce((pairs, value, at, all) =>
              at % 2 === 0 ? [...pairs, [value, all[at + 1]]] : pairs, [])) }
          : { module: rest[0], ...Object.fromEntries(
            rest.slice(2).reduce((pairs, value, at, all) =>
              at % 2 === 0 ? [...pairs, [value, all[at + 1]]] : pairs, [])) });
      }

      default: throw new Error(`unknown route ${route}`);
    }
  })();

  console.log(JSON.stringify(answer, null, 2));
}
