/*
 * The frame around the editor: a tab strip over the open modules, and a panel below it.
 *
 * These are ours rather than the host's because the host will not let them be the host's. Its own
 * tool windows refuse to be sized in any state: setting a width or a height throws whether the
 * window floats or is docked, and docking one produces a band six pixels high with a negative
 * client area. A panel put in one is invisible. Owning the layout here is the only way to have a
 * panel that can be read, and it is also the only way to have tabs at all, because the host has no
 * concept of them.
 */

import { Explorer, type ExplorerProject } from "./explorer.js";
import { buildToolbar, type ToolbarCommand } from "./toolbar.js";

export type FindingSeverity = "error" | "warning" | "info" | "hint";

export interface ShellFinding {
  module: string;
  code?: string;
  message: string;
  severity: FindingSeverity;
  line: number;
  column: number;
}

export interface ShellHandlers {
  /** The developer picked a module from the tab strip. */
  activateModule(name: string): void;
  /** The developer picked a finding, and wants to be taken to it. */
  navigate(module: string, line: number, column: number): void;
  /** The panel was shown or hidden, so the editor has to re-measure. */
  layoutChanged(): void;
  /** A toolbar command was chosen. */
  command(command: ToolbarCommand): void;
  /** Whether an editor command exists in this build. Buttons for missing ones are not drawn. */
  commandAvailable(command: ToolbarCommand): boolean;
  /** The developer entered a line in the Immediate panel. */
  evaluate(text: string): void;
  /** Which panel is showing, and whether the panel is open at all. */
  panelChanged(name: string, open: boolean): void;
}

const SEVERITY_MARK: Record<FindingSeverity, string> = {
  error: "x",
  warning: "!",
  info: "i",
  hint: "i",
};

/** Severity order for sorting, worst first. */
const SEVERITY_RANK: Record<FindingSeverity, number> = {
  error: 0,
  warning: 1,
  info: 2,
  hint: 3,
};

/** Smallest useful panel: the header and about one finding. */
const MIN_PANEL_HEIGHT = 60;

/** The editor never gets less than this, however far the splitter is dragged. */
const MIN_EDITOR_HEIGHT = 80;

/** Smallest useful project explorer: a component name without truncation. */
const MIN_SIDEBAR_WIDTH = 120;

/** The editor never gets less than this, however far the splitter is dragged. */
const MIN_EDITOR_WIDTH = 240;

/** How far one arrow key moves a splitter. */
const KEYBOARD_STEP = 24;

export class Shell {
  private readonly handlers: ShellHandlers;
  private readonly shell: HTMLElement;
  private readonly tabStrip: HTMLElement;
  private readonly splitter: HTMLElement;
  private readonly panel: HTMLElement;
  private readonly panelList: HTMLElement;
  private readonly panelCount: HTMLElement;
  private readonly panelToggle: HTMLButtonElement;
  private readonly statusPosition: HTMLElement;
  private readonly statusModule: HTMLElement;
  private readonly statusNotice: HTMLElement;
  private noticeTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly sidebarSplitter: HTMLElement;
  private readonly explorer: Explorer;
  private readonly panelTabs: HTMLElement;
  private readonly problemsBody: HTMLElement;
  private readonly immediateBody: HTMLElement;
  private readonly immediateLog: HTMLElement;
  private readonly immediateInput: HTMLInputElement;

  /** Lines entered in the Immediate panel, newest last, walked with the arrow keys. */
  private readonly history: string[] = [];
  private historyIndex = 0;

  private modules: string[] = [];
  private active: string | null = null;
  private findings: ShellFinding[] = [];
  private panelOpen = true;
  private panelHeight = 180;
  private sidebarWidth = 260;
  private shown = "problems";

  constructor(root: HTMLElement, handlers: ShellHandlers) {
    this.handlers = handlers;

    this.shell = root.querySelector("#shell") as HTMLElement;
    this.splitter = root.querySelector("#panel-splitter") as HTMLElement;
    this.tabStrip = root.querySelector("#tabs") as HTMLElement;
    this.statusPosition = root.querySelector("#status-position") as HTMLElement;
    this.statusModule = root.querySelector("#status-module") as HTMLElement;
    this.statusNotice = root.querySelector("#status-notice") as HTMLElement;

    this.sidebarSplitter = root.querySelector("#sidebar-splitter") as HTMLElement;
    this.explorer = new Explorer(
      root.querySelector("#sidebar-tree") as HTMLElement,
      (name) => handlers.activateModule(name));

    buildToolbar(
      root.querySelector("#toolbar") as HTMLElement,
      (command) => handlers.command(command),
      (command) => handlers.commandAvailable(command));
    this.panel = root.querySelector("#panel") as HTMLElement;
    this.panelList = root.querySelector("#panel-list") as HTMLElement;
    this.panelCount = root.querySelector("#panel-count") as HTMLElement;
    this.panelToggle = root.querySelector("#panel-toggle") as HTMLButtonElement;

    this.panelTabs = root.querySelector("#panel-tabs") as HTMLElement;
    this.problemsBody = root.querySelector("#panel-list") as HTMLElement;
    this.immediateBody = root.querySelector("#immediate") as HTMLElement;
    this.immediateLog = root.querySelector("#immediate-log") as HTMLElement;
    this.immediateInput = root.querySelector("#immediate-input") as HTMLInputElement;

    this.installPanelTabs();
    this.installImmediate();

    this.panelToggle.addEventListener("click", () => this.togglePanel());
    this.installSplitter();
    this.installSidebarSplitter();

    // One listener on the strip rather than one per tab: the tabs are rebuilt whenever the set of
    // open modules changes, and per-tab listeners would have to be torn down with them.
    this.tabStrip.addEventListener("click", (event) => {
      const tab = (event.target as HTMLElement).closest("[data-module]") as HTMLElement | null;
      if (tab?.dataset.module) {
        this.handlers.activateModule(tab.dataset.module);
      }
    });

    this.panelList.addEventListener("click", (event) => {
      const row = (event.target as HTMLElement).closest("[data-line]") as HTMLElement | null;
      this.goTo(row);
    });

    this.panelList.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        this.goTo(document.activeElement as HTMLElement | null);
      }
    });
  }

  /** Replaces the tab strip. The active module is highlighted, not just listed. */
  setModules(modules: string[], active: string | null): void {
    this.modules = modules;
    this.active = active;
    this.statusModule.textContent = active ?? "";
    this.explorer.setActive(active);
    this.renderTabs();
  }

  /** Replaces the project explorer's contents. */
  setProjects(projects: ExplorerProject[]): void {
    this.explorer.setProjects(projects);
  }

  /** Appends a line to the Immediate panel's output. */
  appendImmediate(text: string, kind: "echo" | "result" | "failed" = "result"): void {
    const line = document.createElement("div");
    line.className = `immediate-line ${kind}`;
    line.textContent = text;

    this.immediateLog.appendChild(line);

    // Kept at the bottom, because the newest line is the answer to what was just asked.
    this.immediateLog.scrollTop = this.immediateLog.scrollHeight;
  }

  /** Brings the Immediate panel forward, opening the panel if it was collapsed. */
  showImmediate(): void {
    this.selectPanel("immediate");

    if (!this.panelOpen) {
      this.togglePanel();
    }

    this.immediateInput.focus();
  }

  private installPanelTabs(): void {
    this.panelTabs.addEventListener("click", (event) => {
      const tab = (event.target as HTMLElement).closest("[data-panel]") as HTMLElement | null;
      if (tab?.dataset.panel) {
        this.selectPanel(tab.dataset.panel);

        // Picking a panel while it is collapsed means wanting to see it.
        if (!this.panelOpen) {
          this.togglePanel();
        }
      }
    });
  }

  private selectPanel(name: string): void {
    for (const tab of this.panelTabs.querySelectorAll<HTMLElement>("[data-panel]")) {
      const active = tab.dataset.panel === name;
      tab.classList.toggle("active", active);
      tab.setAttribute("aria-selected", String(active));
    }

    this.problemsBody.hidden = name !== "problems";
    this.immediateBody.hidden = name !== "immediate";
    this.shown = name;

    // The host only watches what it has to. Reading the editor's Immediate window costs a call
    // every time, and there is no reason to pay it while nobody is looking at the result.
    this.handlers.panelChanged(name, this.panelOpen);

    if (name === "immediate") {
      this.immediateInput.focus();
    }
  }

  private installImmediate(): void {
    this.immediateInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.submitImmediate();
        return;
      }

      // The last few lines are usually variations on each other, so they are worth walking.
      if (event.key === "ArrowUp" || event.key === "ArrowDown") {
        if (this.history.length === 0) {
          return;
        }

        event.preventDefault();
        this.historyIndex = event.key === "ArrowUp"
          ? Math.max(0, this.historyIndex - 1)
          : Math.min(this.history.length, this.historyIndex + 1);

        this.immediateInput.value = this.history[this.historyIndex] ?? "";
        this.immediateInput.setSelectionRange(this.immediateInput.value.length, this.immediateInput.value.length);
      }
    });
  }

  private submitImmediate(): void {
    const text = this.immediateInput.value.trim();
    if (text.length === 0) {
      return;
    }

    this.history.push(text);
    this.historyIndex = this.history.length;
    this.immediateInput.value = "";

    this.appendImmediate(text, "echo");
    this.handlers.evaluate(text);
  }

  /**
   * Shows a message briefly in the status line.
   *
   * The status line rather than a dialog: this is for actions that were legitimately declined, and
   * a modal reply to a click the developer has already moved on from is worse than the silence it
   * replaces. It clears itself, so nothing has to be dismissed.
   */
  notify(text: string): void {
    this.statusNotice.textContent = text;
    this.statusNotice.classList.add("visible");

    if (this.noticeTimer !== undefined) {
      clearTimeout(this.noticeTimer);
    }

    this.noticeTimer = setTimeout(() => {
      this.statusNotice.classList.remove("visible");
      this.statusNotice.textContent = "";
      this.noticeTimer = undefined;
    }, 5000);
  }

  /** Shows where the caret is. */
  setPosition(line: number, column: number): void {
    this.statusPosition.textContent = `Ln ${line}, Col ${column}`;
  }

  /** Replaces the panel's contents. */
  setFindings(findings: ShellFinding[]): void {
    this.findings = findings;
    this.renderPanel();
  }

  /**
   * Makes the divider between the editor and the panel draggable.
   *
   * Pointer events rather than mouse events, so a pen or a touch screen works the same way, and
   * pointer capture so a fast drag that leaves the four pixel divider keeps resizing instead of
   * stopping. The divider is also focusable and moves with the arrow keys: a control that can only
   * be operated by dragging cannot be operated without a pointing device at all.
   */
  private installSplitter(): void {
    this.splitter.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      event.preventDefault();
      this.splitter.setPointerCapture(event.pointerId);

      const startY = event.clientY;
      const startHeight = this.panelHeight;

      const move = (moved: PointerEvent) => this.setPanelHeight(startHeight - (moved.clientY - startY));
      const end = (ended: PointerEvent) => {
        this.splitter.releasePointerCapture(ended.pointerId);
        this.splitter.removeEventListener("pointermove", move);
        this.splitter.removeEventListener("pointerup", end);
        this.splitter.removeEventListener("pointercancel", end);
      };

      this.splitter.addEventListener("pointermove", move);
      this.splitter.addEventListener("pointerup", end);
      this.splitter.addEventListener("pointercancel", end);
    });

    this.splitter.addEventListener("keydown", (event) => {
      if (event.key === "ArrowUp") {
        this.setPanelHeight(this.panelHeight + KEYBOARD_STEP);
      } else if (event.key === "ArrowDown") {
        this.setPanelHeight(this.panelHeight - KEYBOARD_STEP);
      } else {
        return;
      }

      event.preventDefault();
    });
  }

  /** Makes the divider between the project explorer and the editor draggable. */
  private installSidebarSplitter(): void {
    this.sidebarSplitter.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      event.preventDefault();
      this.sidebarSplitter.setPointerCapture(event.pointerId);

      const startX = event.clientX;
      const startWidth = this.sidebarWidth;

      const move = (moved: PointerEvent) => this.setSidebarWidth(startWidth + (moved.clientX - startX));
      const end = (ended: PointerEvent) => {
        this.sidebarSplitter.releasePointerCapture(ended.pointerId);
        this.sidebarSplitter.removeEventListener("pointermove", move);
        this.sidebarSplitter.removeEventListener("pointerup", end);
        this.sidebarSplitter.removeEventListener("pointercancel", end);
      };

      this.sidebarSplitter.addEventListener("pointermove", move);
      this.sidebarSplitter.addEventListener("pointerup", end);
      this.sidebarSplitter.addEventListener("pointercancel", end);
    });

    this.sidebarSplitter.addEventListener("keydown", (event) => {
      if (event.key === "ArrowLeft") {
        this.setSidebarWidth(this.sidebarWidth - KEYBOARD_STEP);
      } else if (event.key === "ArrowRight") {
        this.setSidebarWidth(this.sidebarWidth + KEYBOARD_STEP);
      } else {
        return;
      }

      event.preventDefault();
    });
  }

  private setSidebarWidth(width: number): void {
    const largest = Math.max(MIN_SIDEBAR_WIDTH, this.shell.clientWidth - MIN_EDITOR_WIDTH);

    this.sidebarWidth = Math.round(Math.min(largest, Math.max(MIN_SIDEBAR_WIDTH, width)));
    this.shell.style.setProperty("--sidebar-width", `${this.sidebarWidth}px`);
    this.handlers.layoutChanged();
  }

  private setPanelHeight(height: number): void {
    // Bounded against the shell rather than a fixed maximum, so the editor cannot be squeezed out
    // of existence on a short window and the panel cannot be dragged past the top of one.
    const available = this.shell.clientHeight - this.tabStrip.offsetHeight - this.splitter.offsetHeight;
    const largest = Math.max(MIN_PANEL_HEIGHT, available - MIN_EDITOR_HEIGHT);

    this.panelHeight = Math.round(Math.min(largest, Math.max(MIN_PANEL_HEIGHT, height)));
    this.shell.style.setProperty("--panel-height", `${this.panelHeight}px`);
    this.handlers.layoutChanged();
  }

  private togglePanel(): void {
    this.panelOpen = !this.panelOpen;
    this.panel.classList.toggle("collapsed", !this.panelOpen);
    this.shell.classList.toggle("panel-collapsed", !this.panelOpen);
    this.panelToggle.setAttribute("aria-expanded", String(this.panelOpen));
    this.handlers.panelChanged(this.shown, this.panelOpen);
    this.handlers.layoutChanged();
  }

  private goTo(row: HTMLElement | null): void {
    if (!row?.dataset.line) {
      return;
    }
    this.handlers.navigate(
      row.dataset.moduleName ?? "",
      Number(row.dataset.line),
      Number(row.dataset.column ?? "1"),
    );
  }

  private renderTabs(): void {
    this.tabStrip.replaceChildren();

    for (const name of this.modules) {
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "tab" + (name === this.active ? " active" : "");
      tab.dataset.module = name;
      tab.textContent = name;
      // The tab strip is a tab list for anything reading the page aloud, not a row of buttons.
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", String(name === this.active));

      const count = this.findings.filter((f) => f.module === name).length;
      if (count > 0) {
        const badge = document.createElement("span");
        badge.className = "tab-badge";
        badge.textContent = String(count);
        // Read out rather than left as a bare number next to a name.
        badge.title = `${count} problem${count === 1 ? "" : "s"} in ${name}`;
        tab.appendChild(badge);
      }

      this.tabStrip.appendChild(tab);
    }
  }

  private renderPanel(): void {
    const errors = this.findings.filter((f) => f.severity === "error").length;
    const warnings = this.findings.filter((f) => f.severity === "warning").length;

    this.panelCount.textContent = this.findings.length === 0
      ? "no problems"
      : `${errors} error${errors === 1 ? "" : "s"}, ${warnings} warning${warnings === 1 ? "" : "s"}`;

    this.panelList.replaceChildren();

    // Worst first, then by where they are, so the order does not depend on which module the
    // engine happened to analyse first.
    const sorted = [...this.findings].sort((a, b) =>
      SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity]
      || a.module.localeCompare(b.module)
      || a.line - b.line
      || a.column - b.column);

    for (const finding of sorted) {
      const row = document.createElement("div");
      row.className = `row ${finding.severity}`;
      row.tabIndex = 0;
      row.dataset.moduleName = finding.module;
      row.dataset.line = String(finding.line);
      row.dataset.column = String(finding.column);

      // The severity is carried by a glyph and by the text, never by colour alone.
      const mark = document.createElement("span");
      mark.className = "mark";
      mark.textContent = SEVERITY_MARK[finding.severity];
      mark.setAttribute("aria-label", finding.severity);

      const body = document.createElement("div");
      body.className = "body";

      const message = document.createElement("div");
      message.className = "message";
      message.textContent = finding.message;

      const where = document.createElement("div");
      where.className = "where";
      where.textContent = `${finding.module} (${finding.line}, ${finding.column})`
        + (finding.code ? `   ${finding.code}` : "");

      body.append(message, where);
      row.append(mark, body);
      this.panelList.appendChild(row);
    }

    // Counts per component change with the findings, so both the strip and the tree are rebuilt.
    const counts = new Map<string, number>();
    for (const finding of this.findings) {
      counts.set(finding.module, (counts.get(finding.module) ?? 0) + 1);
    }

    this.explorer.setProblemCounts(counts);
    this.renderTabs();
  }
}
