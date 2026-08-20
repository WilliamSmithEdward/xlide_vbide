/*
 * The Tests pane: the project's `@xlide-test` procedures as a tree, run and watched live.
 *
 * The pane is a PROJECTION of one message: the host sends its whole picture (support state,
 * run state, every test with its latest outcome) on every change, and the pane repaints from
 * it. Nothing here diffs, counts or remembers on its own - a panel that keeps a second copy
 * of the runner's state is a panel that drifts from it.
 *
 * Results stream: the host repaints per landing result, so a run fills the tree row by row
 * with the current test wearing the spinner - the live-host advantage over a runner that
 * stages a copy and reports at the end.
 */

import type { SetTestsState, TestRow } from "./bridge.js";

export interface TestsPaneDeps {
  /** Press a runner verb: refresh, install, run, runFailed, runOne or debug. */
  act(action: string, test?: string): void;
  /** Open the test's module at its declaration line. */
  navigate(module: string, line: number): void;
}

interface StatusShape {
  icon: string;
  className: string;
  word: string;
}

const UNKNOWN_STATUS: StatusShape = { icon: "circle-large-outline", className: "tests-none", word: "not run" };

const STATUS_GLYPH: Record<string, StatusShape> = {
  none: UNKNOWN_STATUS,
  running: { icon: "loading codicon-modifier-spin", className: "tests-running", word: "running" },
  passed: { icon: "pass", className: "tests-passed", word: "passed" },
  failed: { icon: "error", className: "tests-failed", word: "failed" },
  error: { icon: "warning", className: "tests-error", word: "errored" },
  skipped: { icon: "circle-slash", className: "tests-skipped", word: "skipped" },
  "skip-marked": { icon: "circle-slash", className: "tests-skipped", word: "marked skip" },
  xfail: { icon: "pass", className: "tests-xfail", word: "expected failure" },
  xpass: { icon: "error", className: "tests-failed", word: "unexpected pass" },
};

export class TestsPane {
  private readonly deps: TestsPaneDeps;
  private readonly list: HTMLElement;
  private readonly summary: HTMLElement;
  private readonly install: HTMLButtonElement;
  private readonly runAll: HTMLButtonElement;
  private readonly runFailed: HTMLButtonElement;
  private state: SetTestsState = { support: "missing", running: false, currentTest: null, rows: [] };

  constructor(root: HTMLElement, deps: TestsPaneDeps) {
    this.deps = deps;
    this.list = root.querySelector("#tests-list") as HTMLElement;
    this.summary = root.querySelector("#tests-summary") as HTMLElement;
    this.install = root.querySelector("#tests-install") as HTMLButtonElement;
    this.runAll = root.querySelector("#tests-run") as HTMLButtonElement;
    this.runFailed = root.querySelector("#tests-run-failed") as HTMLButtonElement;

    this.runAll.addEventListener("click", () => this.deps.act("run"));
    this.runFailed.addEventListener("click", () => this.deps.act("runFailed"));
    (root.querySelector("#tests-refresh") as HTMLButtonElement)
      .addEventListener("click", () => this.deps.act("refresh"));
    this.install.addEventListener("click", () => this.deps.act("install"));
  }

  /** The pane asked to be shown: rediscover, so the tree matches the code as it stands. */
  shown(): void {
    this.deps.act("refresh");
  }

  paint(state: SetTestsState): void {
    this.state = state;
    const counts = { passed: 0, failed: 0, rest: 0 };
    for (const row of state.rows) {
      if (row.status === "passed" || row.status === "xfail") {
        counts.passed++;
      } else if (row.status === "failed" || row.status === "error" || row.status === "xpass") {
        counts.failed++;
      } else {
        counts.rest++;
      }
    }

    this.summary.textContent = state.rows.length === 0
      ? ""
      : `${counts.passed} passed, ${counts.failed} failed of ${state.rows.length}`;

    // The support chip: absent while the assert module is current, an action while it is not.
    if (state.support === "installed") {
      this.install.hidden = true;
    } else {
      this.install.hidden = false;
      this.install.textContent = state.support === "missing" ? "Install XlideAssert" : "Update XlideAssert";
      this.install.title = state.support === "missing"
        ? "Tests call the XlideAssert module; install it into the project to run them."
        : "The project's XlideAssert is older than this product's; update it to run tests.";
    }

    const canRun = state.support === "installed" && !state.running && state.rows.length > 0;
    this.runAll.disabled = !canRun;
    this.runFailed.disabled = !canRun
      || !state.rows.some((row) => row.status === "failed" || row.status === "error" || row.status === "xpass");

    this.list.replaceChildren();
    if (state.rows.length === 0) {
      const empty = document.createElement("div");
      empty.className = "tests-empty";
      empty.textContent = state.support === "missing"
        ? "No tests. Mark a zero-argument Sub in a standard module with ' @xlide-test, and install XlideAssert for it to call."
        : "No tests. Mark a zero-argument Sub in a standard module with ' @xlide-test.";
      this.list.appendChild(empty);
      return;
    }

    let lastModule = "";
    for (const row of state.rows) {
      if (row.module !== lastModule) {
        lastModule = row.module;
        const heading = document.createElement("div");
        heading.className = "tests-module";
        heading.textContent = row.module;
        this.list.appendChild(heading);
      }

      this.list.appendChild(this.renderRow(row));
    }
  }

  private renderRow(row: TestRow): HTMLElement {
    const shape = STATUS_GLYPH[row.status] ?? UNKNOWN_STATUS;
    const line = document.createElement("div");
    line.className = `tests-row ${shape.className}`;
    line.setAttribute("role", "treeitem");

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${shape.icon} tests-glyph`;
    glyph.setAttribute("aria-label", shape.word);

    const name = document.createElement("button");
    name.type = "button";
    name.className = "tests-name";
    name.textContent = row.procedure;
    name.title = `${row.id} - ${shape.word}. Click to open.`;
    name.addEventListener("click", () => this.deps.navigate(row.module, row.line));

    const meta = document.createElement("span");
    meta.className = "tests-meta";
    const chips: string[] = [...row.tags];
    if (row.timeoutMs !== null) {
      chips.push(`timeout:${row.timeoutMs}ms`);
    }

    if (row.expectedError !== null) {
      chips.push(`expects error ${row.expectedError}`);
    }

    meta.textContent = chips.join(" · ");

    const time = document.createElement("span");
    time.className = "tests-time";
    time.textContent = row.durationMs > 0 ? `${Math.round(row.durationMs)} ms` : "";

    const run = document.createElement("button");
    run.type = "button";
    run.className = "tests-act codicon codicon-play";
    run.title = `Run ${row.procedure}`;
    run.disabled = this.state.running || this.state.support !== "installed";
    run.addEventListener("click", () => this.deps.act("runOne", row.id));

    const debug = document.createElement("button");
    debug.type = "button";
    debug.className = "tests-act codicon codicon-debug-alt";
    debug.title = `Debug ${row.procedure}: no trap, so a breakpoint or an error stops in the debugger`;
    debug.disabled = this.state.running || this.state.support !== "installed";
    debug.addEventListener("click", () => this.deps.act("debug", row.id));

    line.append(glyph, name, meta, time, run, debug);

    if (row.message) {
      const holder = document.createElement("div");
      holder.className = "tests-row-block";
      holder.appendChild(line);
      const message = document.createElement("div");
      message.className = "tests-message";
      message.textContent = row.message;
      holder.appendChild(message);
      for (const said of row.output) {
        const output = document.createElement("div");
        output.className = "tests-output";
        output.textContent = said;
        holder.appendChild(output);
      }

      return holder;
    }

    return line;
  }
}
