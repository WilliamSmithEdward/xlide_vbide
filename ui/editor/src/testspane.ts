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
import { ScopeSelect, type ScopeEntry } from "./scopeselect.js";

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
  // An expected failure is EXPECTED: a muted check, never the pass green - a green check over
  // a red message read as a contradiction (the owner, 2026-08-20). Its message mutes too.
  xfail: { icon: "pass", className: "tests-xfail", word: "expected failure" },
  xpass: { icon: "error", className: "tests-failed", word: "unexpected pass" },
};

/** The filter groups, the Problems pane's own idea: pressed shows, pressed-out hides. */
type TestGroup = "passed" | "failed" | "xfail" | "skipped" | "notRun";

const GROUP_ORDER: TestGroup[] = ["passed", "failed", "xfail", "skipped", "notRun"];

const GROUP_SHAPE: Record<TestGroup, { icon: string; many: string }> = {
  passed: { icon: "pass", many: "Passed" },
  failed: { icon: "error", many: "Failed" },
  xfail: { icon: "pass", many: "XFail" },
  skipped: { icon: "circle-slash", many: "Skipped" },
  notRun: { icon: "circle-large-outline", many: "Not Run" },
};

function groupOf(status: string): TestGroup | null {
  switch (status) {
    case "passed":
      return "passed";
    case "failed":
    case "error":
    case "xpass":
      return "failed";
    case "xfail":
      return "xfail";
    case "skipped":
    case "skip-marked":
      return "skipped";
    case "none":
      return "notRun";
    default:
      // A running test is never filtered away: hiding the row in flight is hiding the run.
      return null;
  }
}

/**
 * When the last run finished, said the way a person would: the clock alone for a run from
 * today, the date in front of it for an older one. Seconds are in because reruns land inside
 * the same minute all day, and the whole point of the readout is telling one from the next.
 */
function describeRun(iso: string | null): string {
  if (iso === null) {
    return "";
  }

  const at = new Date(iso);
  if (Number.isNaN(at.getTime())) {
    return "";
  }

  const clock = at.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit", second: "2-digit" });
  const now = new Date();
  const sameDay = at.getFullYear() === now.getFullYear()
    && at.getMonth() === now.getMonth()
    && at.getDate() === now.getDate();
  return sameDay
    ? `Ran ${clock}`
    : `Ran ${at.toLocaleDateString(undefined, { month: "short", day: "numeric" })}, ${clock}`;
}

export class TestsPane {
  private readonly deps: TestsPaneDeps;
  private readonly list: HTMLElement;
  private readonly summary: HTMLElement;
  private readonly ran: HTMLElement;
  private readonly install: HTMLButtonElement;
  private readonly runAll: HTMLButtonElement;
  private readonly runAllLabel: HTMLElement;
  private readonly runFailed: HTMLButtonElement;
  private readonly scope: ScopeSelect;
  private readonly filterButtons = new Map<TestGroup, HTMLButtonElement>();
  private readonly filters: Record<TestGroup, boolean> = {
    passed: true, failed: true, xfail: true, skipped: true, notRun: true,
  };

  /** The active tab's module, for the Current Module scope. Pushed in by the shell. */
  private activeModule: string | null = null;

  private state: SetTestsState = {
    support: "missing", running: false, currentTest: null, ranAt: null, rows: [],
  };

  constructor(root: HTMLElement, deps: TestsPaneDeps) {
    this.deps = deps;
    this.list = root.querySelector("#tests-list") as HTMLElement;
    this.summary = root.querySelector("#tests-summary") as HTMLElement;
    this.ran = root.querySelector("#tests-ran") as HTMLElement;
    this.install = root.querySelector("#tests-install") as HTMLButtonElement;
    this.runAll = root.querySelector("#tests-run") as HTMLButtonElement;
    this.runAllLabel = this.runAll.querySelector(".tests-label") as HTMLElement;
    this.runFailed = root.querySelector("#tests-run-failed") as HTMLButtonElement;

    // The module scope sits with the run buttons, because it changes what they do, and left of
    // the outcome filters: which tests first, then which outcomes of them.
    this.scope = new ScopeSelect("tests-scope", "Show tests from", "tests", () => this.paint(this.state));
    (root.querySelector("#tests-scope-seat") as HTMLElement).appendChild(this.scope.element);

    // A run follows the scope. Scoped to a module, Run All runs that module and Failed reruns
    // its failures alone - a button that ran tests the pane is not showing would be a button
    // that disagrees with the list above it.
    this.runAll.addEventListener("click", () => {
      const only = this.scope.scopeName();
      this.deps.act(only ? "runModule" : "run", only ?? undefined);
    });
    this.runFailed.addEventListener("click", () => {
      this.deps.act("runFailed", this.scope.scopeName() ?? undefined);
    });
    (root.querySelector("#tests-refresh") as HTMLButtonElement)
      .addEventListener("click", () => this.deps.act("refresh"));
    this.install.addEventListener("click", () => this.deps.act("install"));

    // The outcome filters, the Problems pane's own gesture: each shows its count always and
    // hides its rows when pressed out.
    const filterRow = root.querySelector("#tests-filters") as HTMLElement;
    for (const group of GROUP_ORDER) {
      const shape = GROUP_SHAPE[group];
      const button = document.createElement("button");
      button.type = "button";
      button.className = `tests-filter tests-filter-${group}`;
      button.setAttribute("aria-pressed", "true");
      const icon = document.createElement("span");
      icon.className = `codicon codicon-${shape.icon}`;
      icon.setAttribute("aria-hidden", "true");
      const count = document.createElement("span");
      count.className = "filter-count";
      count.textContent = `0 ${shape.many}`;
      button.append(icon, count);
      button.title = "Click to toggle; Ctrl+click to show only this";
      // On POINTERDOWN, not click: the pane repaints itself whenever an analysis pass lands,
      // and a repaint between a real press and its release breaks the browser's click
      // pairing - the same churn that broke the canvas double-click. The down always fires
      // on the pressed node, whatever happens to the DOM after it.
      button.addEventListener("pointerdown", (event) => {
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) {
          // SOLO: only this group shows - and soloing the solo brings everything back.
          const alreadySolo = this.filters[group]
            && GROUP_ORDER.every((other) => other === group || !this.filters[other]);
          for (const other of GROUP_ORDER) {
            this.filters[other] = alreadySolo || other === group;
          }
        } else {
          this.filters[group] = !this.filters[group];
        }

        for (const [other, otherButton] of this.filterButtons) {
          otherButton.setAttribute("aria-pressed", String(this.filters[other]));
        }

        button.setAttribute("aria-pressed", String(this.filters[group]));
        this.paint(this.state);
      });
      filterRow.appendChild(button);
      this.filterButtons.set(group, button);
    }
  }

  /** The pane asked to be shown: rediscover, so the tree matches the code as it stands. */
  shown(): void {
    this.deps.act("refresh");
  }

  /** The active tab moved. A Current Module scope follows it; every other scope holds still. */
  setActiveModule(module: string | null): void {
    if (module === this.activeModule) {
      return;
    }

    this.activeModule = module;
    this.paint(this.state);
  }

  paint(state: SetTestsState): void {
    this.state = state;

    // The scope's option list, rebuilt from the discovered tests so it follows the project as
    // modules come and go. Tests are the ACTIVE project's, so a bare module name is unambiguous
    // here in a way it is not in the Problems pane.
    const scopes = new Map<string, ScopeEntry>();
    for (const row of state.rows) {
      const key = row.module.toLowerCase();
      const already = scopes.get(key);
      if (already) {
        already.count++;
      } else {
        scopes.set(key, { key, name: row.module, label: row.module, count: 1 });
      }
    }

    const activeKey = this.activeModule?.toLowerCase() ?? null;
    this.scope.setEntries(
      [...scopes.values()].sort((a, b) => a.label.localeCompare(b.label)),
      activeKey && this.activeModule ? { key: activeKey, name: this.activeModule } : null);

    const shown = state.rows.filter((row) => this.scope.admits(row.module.toLowerCase()));
    const counts: Record<TestGroup, number> = { passed: 0, failed: 0, xfail: 0, skipped: 0, notRun: 0 };
    for (const row of shown) {
      const group = groupOf(row.status);
      if (group !== null) {
        counts[group]++;
      }
    }

    for (const [group, button] of this.filterButtons) {
      const label = button.querySelector(".filter-count");
      if (label) {
        label.textContent = `${counts[group]} ${GROUP_SHAPE[group].many}`;
      }
    }

    this.summary.textContent = shown.length === 0
      ? ""
      : `${counts.passed} passed, ${counts.failed} failed of ${shown.length}`;

    // When the results landed. "7 passed" with no clock beside it cannot tell a result that
    // just arrived from one left over from an hour ago, and a rerun that changes nothing looks
    // identical to a rerun that never happened.
    this.ran.textContent = state.running ? "running" : describeRun(state.ranAt);
    this.ran.hidden = this.ran.textContent.length === 0;
    this.ran.title = state.ranAt !== null && !state.running
      ? `The last run finished ${new Date(state.ranAt).toLocaleString()}`
      : "";

    // The support chip is a STATUS, always visible: yellow while the assert module needs
    // installing or updating - a press does it - and green once it is current, when the
    // press has nothing left to do.
    this.install.hidden = false;
    if (state.support === "installed") {
      this.install.textContent = "XlideAssert Installed";
      this.install.title = "The XlideAssert module in this project matches this product.";
      this.install.classList.remove("tests-install-needed");
      this.install.classList.add("tests-install-ok");
      this.install.disabled = true;
    } else {
      this.install.textContent = state.support === "missing" ? "Install XlideAssert" : "Update XlideAssert";
      this.install.title = state.support === "missing"
        ? "Tests call the XlideAssert module; install it into the project to run them."
        : "The project's XlideAssert is older than this product's; update it to run tests.";
      this.install.classList.remove("tests-install-ok");
      this.install.classList.add("tests-install-needed");
      this.install.disabled = false;
    }

    // The run buttons say what the scope has made them: Run All means all of what is showing.
    const only = this.scope.scopeName();
    const canRun = state.support === "installed" && !state.running && shown.length > 0;
    this.runAll.disabled = !canRun;
    this.runAllLabel.textContent = only ? "Run Module" : "Run All";
    this.runAll.title = only
      ? `Run the ${shown.length} test${shown.length === 1 ? "" : "s"} in ${only}`
      : "Run every test";
    this.runFailed.disabled = !canRun
      || !shown.some((row) => row.status === "failed" || row.status === "error" || row.status === "xpass");
    this.runFailed.title = only ? `Rerun what failed in ${only}` : "Rerun what failed";

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

    // Scoped to a module with nothing in it - or to a tab that holds no tests - the pane says
    // so and offers its way back, rather than reading as a project that has lost its tests.
    if (shown.length === 0) {
      const empty = document.createElement("div");
      empty.className = "tests-empty";
      const said = document.createElement("span");
      said.textContent = (only ? `No tests in ${only}.` : "The active tab holds no tests.")
        + ` ${state.rows.length} elsewhere in the project.`;
      const showAll = document.createElement("button");
      showAll.type = "button";
      showAll.className = "panel-empty-act";
      showAll.textContent = "Show All";
      showAll.addEventListener("click", () => {
        this.scope.reset();
        this.paint(this.state);
      });
      empty.append(said, showAll);
      this.list.appendChild(empty);
      return;
    }

    let lastModule = "";
    let heading: HTMLElement | null = null;
    let shownUnder = 0;
    for (const row of shown) {
      if (row.module !== lastModule) {
        if (heading !== null && shownUnder === 0) {
          heading.remove();
        }

        lastModule = row.module;
        shownUnder = 0;
        heading = document.createElement("div");
        heading.className = "tests-module";
        heading.textContent = row.module;
        this.list.appendChild(heading);
      }

      const group = groupOf(row.status);
      if (group !== null && !this.filters[group]) {
        continue;
      }

      shownUnder++;
      this.list.appendChild(this.renderRow(row));
    }

    // A heading with every child filtered away says nothing; it goes with them.
    if (heading !== null && shownUnder === 0) {
      heading.remove();
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
      holder.className = `tests-row-block ${shape.className}`;
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
