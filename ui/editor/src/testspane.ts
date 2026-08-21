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
  /**
   * Press a runner verb: refresh, install, run, runFile, runModule, runOne, runFailed or
   * debug. `file` scopes it to one open file, which every verb needs once more than one file
   * can hold a module of the same name.
   */
  act(action: string, test?: string, file?: string): void;
  /** Open the test's module at its declaration line, in the file that holds it. */
  navigate(module: string, line: number, file: string): void;
}

/** A module's identity in a session: its name is only unique inside its own file. */
function moduleKey(file: string, module: string): string {
  return `${file.toLowerCase()}\0${module.toLowerCase()}`;
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

  /** The active tab's module and file, for the Current Module scope. Pushed in by the shell. */
  private activeModule: string | null = null;
  private activeFile: string | null = null;

  private state: SetTestsState = {
    support: "missing", running: false, currentTest: null, ranAt: null, files: [], rows: [],
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

    // The scope sits with the run buttons, because it changes what they do, and left of the
    // outcome filters: which tests first, then which outcomes of them.
    this.scope = new ScopeSelect("tests-scope", "Show tests from", "tests", () => this.paint(this.state));
    (root.querySelector("#tests-scope-seat") as HTMLElement).appendChild(this.scope.element);

    // A run follows the scope. Scoped to a module or a file, Run All runs that, and Failed
    // reruns its failures alone - a button that ran tests the pane is not showing would be a
    // button that disagrees with the list above it. The file rides along every verb, because a
    // module name does not say which open file's copy of it to run.
    this.runAll.addEventListener("click", () => this.deps.act(...this.runTarget()));
    this.runFailed.addEventListener("click", () => {
      const [, target, file] = this.runTarget();
      this.deps.act("runFailed", target, file);
    });
    (root.querySelector("#tests-refresh") as HTMLButtonElement)
      .addEventListener("click", () => this.deps.act("refresh"));

    // The chip installs where it is pointing: into one file when the pane is scoped to one,
    // into every file that needs it when the pane is speaking for the whole session.
    this.install.addEventListener("click", () =>
      this.deps.act("install", undefined, this.scope.scopeFile() ?? undefined));

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
    // NOT "refresh": showing a pane is not a reason to re-read the project. On a large one
    // that walk is a third of a second of COM on the host thread (81,795 lines measured
    // 2026-08-20), and it buys nothing the analysis pass has not already told this pane -
    // auto-rediscovery keeps it current whether it is on screen or not. The refresh BUTTON
    // still forces the read, which is what a button labelled refresh is for.
    this.deps.act("show");
  }

  /** The active tab moved. A Current Module scope follows it; every other scope holds still. */
  setActiveModule(module: string | null, file: string | null): void {
    if (module === this.activeModule && file === this.activeFile) {
      return;
    }

    this.activeModule = module;
    this.activeFile = file;
    this.paint(this.state);
  }

  /**
   * What the run buttons should ask for, from where the scope is pointing: the whole session,
   * one file, or one module of one file.
   */
  private runTarget(): [action: string, target: string | undefined, file: string | undefined] {
    const file = this.scope.scopeFile() ?? undefined;
    switch (this.scope.scopeKind()) {
      case "all":
        return ["run", undefined, undefined];
      case "file":
        return ["runFile", undefined, file];
      default:
        return ["runModule", this.scope.scopeName() ?? undefined, file];
    }
  }

  paint(state: SetTestsState): void {
    this.state = state;

    // The scope's option list, rebuilt from the discovered tests so it follows the session as
    // modules and files come and go. A module is keyed by its file as well as its name: two
    // open files may each hold an InvoiceTests, and they are not the same module.
    const scopes = new Map<string, ScopeEntry>();
    for (const row of state.rows) {
      const key = moduleKey(row.file, row.module);
      const already = scopes.get(key);
      if (already) {
        already.count++;
      } else {
        scopes.set(key, { key, name: row.module, file: row.file, count: 1 });
      }
    }

    // Every open file the host told us about, in its order - the developer's own file first -
    // and the count is of tests, so a file with none says so rather than disappearing.
    const files = state.files.map((file) => ({ name: file.file, count: file.tests }));
    const activeKey = this.activeModule && this.activeFile
      ? moduleKey(this.activeFile, this.activeModule)
      : null;
    this.scope.setEntries(
      [...scopes.values()].sort((a, b) => a.file.localeCompare(b.file) || a.name.localeCompare(b.name)),
      activeKey && this.activeModule && this.activeFile
        ? { key: activeKey, name: this.activeModule, file: this.activeFile }
        : null,
      files);

    const shown = state.rows.filter((row) => this.scope.admits(moduleKey(row.file, row.module), row.file));
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

    // The support chip speaks for WHAT THE PANE IS SHOWING. XlideAssert is a module inside a
    // file, so a session with two files open has two answers: scoped to one file, the chip is
    // that file's and installs into it; unscoped, it is the worst standing among the files that
    // hold tests, and installing fixes all of them - which is what a chip speaking for the whole
    // session has to mean.
    const scopedFile = this.scope.scopeFile();
    const standing = scopedFile
      ? state.files.find((file) => file.file.toLowerCase() === scopedFile.toLowerCase())?.support ?? state.support
      : state.support;
    const needing = state.files.filter((file) => file.tests > 0 && file.support !== "installed");
    const wheres = scopedFile ?? (needing.length > 1 ? `${needing.length} files` : needing[0]?.file ?? "this file");

    // TWO FILES CAN NEED DIFFERENT THINGS AT ONCE - one has no XlideAssert, the other has an old
    // one - and one verb cannot be true of both. The button takes the worse of the two, which is
    // what the press does first anyway, and says how many files it is speaking for; the tooltip
    // is where each file gets its own word (the owner, 2026-08-20).
    const spelt = needing
      .map((file) => `${file.file} needs it ${file.support === "missing" ? "installed" : "updated"}`)
      .join("; ");

    this.install.hidden = false;
    if (standing === "installed") {
      this.install.textContent = "XlideAssert Installed";
      this.install.title = scopedFile
        ? `The XlideAssert module in ${scopedFile} matches this product.`
        : "Every open file that holds tests carries an XlideAssert matching this product.";
      this.install.classList.remove("tests-install-needed");
      this.install.classList.add("tests-install-ok");
      this.install.disabled = true;
    } else {
      const verb = standing === "missing" ? "Install" : "Update";
      const many = !scopedFile && needing.length > 1;
      this.install.textContent = `${verb} XlideAssert${many ? ` (${needing.length} files)` : ""}`;
      this.install.title = many
        ? `${spelt}. Press to do both.`
        : standing === "missing"
          ? `Tests call the XlideAssert module; install it into ${wheres} to run them.`
          : `The XlideAssert in ${wheres} is older than this product's; update it to run tests.`;
      this.install.classList.remove("tests-install-ok");
      this.install.classList.add("tests-install-needed");
      this.install.disabled = false;
    }

    // The run buttons say what the scope has made them: Run All means all of what is showing.
    const kind = this.scope.scopeKind();
    const only = this.scope.scopeName();
    const canRun = standing === "installed" && !state.running && shown.length > 0;
    const howMany = `${shown.length} test${shown.length === 1 ? "" : "s"}`;
    this.runAll.disabled = !canRun;
    this.runAllLabel.textContent = kind === "all" ? "Run All" : kind === "file" ? "Run File" : "Run Module";
    this.runAll.title = only ? `Run the ${howMany} in ${only}` : "Run every test in every open file";
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

    // Scoped to a module or file with nothing in it - or to a tab that holds no tests - the
    // pane says so and offers its way back, rather than reading as a session that has lost its
    // tests.
    if (shown.length === 0) {
      const empty = document.createElement("div");
      empty.className = "tests-empty";
      const said = document.createElement("span");
      said.textContent = (only ? `No tests in ${only}.` : "The active tab holds no tests.")
        + ` ${state.rows.length} elsewhere.`;
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

    // The outcome filters decide what is left BEFORE any heading is drawn, so a heading whose
    // every row was filtered away is never drawn in the first place.
    const visible = shown.filter((row) => {
      const group = groupOf(row.status);
      return group === null || this.filters[group];
    });

    if (visible.length === 0) {
      const empty = document.createElement("div");
      empty.className = "tests-empty";
      const said = document.createElement("span");
      said.textContent = `Every one of the ${shown.length} tests here is hidden by the outcome filters.`;
      const showAll = document.createElement("button");
      showAll.type = "button";
      showAll.className = "panel-empty-act";
      showAll.textContent = "Show All Outcomes";
      showAll.addEventListener("click", () => {
        for (const group of GROUP_ORDER) {
          this.filters[group] = true;
          this.filterButtons.get(group)?.setAttribute("aria-pressed", "true");
        }

        this.paint(this.state);
      });
      empty.append(said, showAll);
      this.list.appendChild(empty);
      return;
    }

    // Two headings, and the outer one only when it is telling the developer something: with one
    // file's tests on screen every row would carry the same file name, which is noise.
    const manyFiles = new Set(visible.map((row) => row.file.toLowerCase())).size > 1;
    let lastFile = "";
    let lastModule = "";
    for (const row of visible) {
      if (manyFiles && row.file !== lastFile) {
        const fileHeading = document.createElement("div");
        fileHeading.className = "tests-file";
        fileHeading.textContent = row.file;
        this.list.appendChild(fileHeading);
      }

      if (row.module !== lastModule || row.file !== lastFile) {
        const heading = document.createElement("div");
        heading.className = "tests-module";
        heading.textContent = row.module;
        this.list.appendChild(heading);
      }

      lastFile = row.file;
      lastModule = row.module;
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
    name.title = `${row.file}!${row.id} - ${shape.word}. Click to open.`;
    name.addEventListener("click", () => this.deps.navigate(row.module, row.line, row.file));

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

    // A row's own acts name its file, because its id alone does not: the same test name can be
    // open in two files at once. Both are gated on THAT file's support module.
    const ready = this.state.files
      .find((file) => file.file.toLowerCase() === row.file.toLowerCase())?.support === "installed";

    const run = document.createElement("button");
    run.type = "button";
    run.className = "tests-act codicon codicon-play";
    run.title = `Run ${row.procedure}`;
    run.disabled = this.state.running || !ready;
    run.addEventListener("click", () => this.deps.act("runOne", row.id, row.file));

    const debug = document.createElement("button");
    debug.type = "button";
    debug.className = "tests-act codicon codicon-debug-alt";
    debug.title = `Debug ${row.procedure}: no trap, so a breakpoint or an error stops in the debugger`;
    debug.disabled = this.state.running || !ready;
    debug.addEventListener("click", () => this.deps.act("debug", row.id, row.file));

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
