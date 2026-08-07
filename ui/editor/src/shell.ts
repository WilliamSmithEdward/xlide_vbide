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

import { showContextMenu, type ContextMenuItem } from "./contextmenu.js";
import { Explorer, problemCountKey, type ExplorerProcedure, type ExplorerProject } from "./explorer.js";
import { Menubar, type MenuItem } from "./menubar.js";
import { PanelDocks, type PanelSeat } from "./paneldocks.js";
import type { PaneVisibilityControl } from "./settingsdialog.js";
import { buildToolbar, type ToolbarCommand } from "./toolbar.js";

export type FindingSeverity = "error" | "warning" | "info" | "hint";

export interface ShellFinding {
  module: string;
  code?: string;
  message: string;
  severity: FindingSeverity;
  line: number;
  column: number;
  /** The workbook the module belongs to, when the host could say. */
  project?: string | null;
}

export interface ShellProperty {
  name: string;
  value: string;
  /** Whether an edit will be attempted. The host can still refuse one, and says so. */
  writable: boolean;
  /** True and False rather than free text. */
  boolean: boolean;
}

export interface ShellHandlers {
  /** The developer picked a module. The tree names the workbook; the tab strip cannot yet. */
  activateModule(name: string, workbook?: string): void;
  /** Whether some group holds a tab for this module, for the explorer's context menu. */
  moduleIsOpen(name: string): boolean;
  /** The developer picked a finding or a procedure, and wants to be taken to it. */
  navigate(module: string, line: number, column: number, selectLine?: boolean, workbook?: string): void;
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
  /** The developer selected a component in the explorer without opening it. */
  selectComponent(name: string): void;
  /** The developer renamed a module, and everything that names it follows. */
  renameModule(name: string, workbook: string | null, newName: string): void;
  /** The developer opened a menu; the host is asked for its items. [] is the bar itself. */
  menuRequest(path: number[]): void;
  /** The developer chose a menu item, named by its position chain. */
  menuExecute(path: number[]): void;
  /** Every menu closed, so focus belongs to the editor again. */
  menuClosed(): void;
  /** The developer edited a property of the shown component. */
  editProperty(component: string, name: string, value: string): void;
  /** The developer closed a module's tab, however they did it; the workbook when known. The
   * action carries their answer to the unsaved-changes question — "save" or "discard" — and
   * is absent on the plain close the host checks. */
  closeModule(name: string, workbook?: string, action?: string): void;
  /** The developer asked for a new component: 1 module, 2 class module, 3 form. */
  insertComponent(kind: number, project?: string): void;
  /** A module's procedures, for its unfolded node in the tree; null when no answer came. */
  requestOutline(module: string, workbook?: string): Promise<ExplorerProcedure[] | null>;
  /** A line for the host's log, from the corners only the log's data cadence explains. */
  trace(text: string): void;
}

const SEVERITY_MARK: Record<FindingSeverity, string> = {
  error: "x",
  warning: "!",
  info: "i",
  hint: "i",
};

/** One tab: the module, and the workbook it belongs to when the host could say. */
interface TabIdentity {
  name: string;
  project: string | null;
}

/** The identity two tabs are the same by. Case-insensitive, the way the host compares names. */
function tabKey(tab: TabIdentity): string {
  return `${(tab.project ?? "").toLowerCase()}\0${tab.name.toLowerCase()}`;
}

/** Severity order for sorting, worst first. */
const SEVERITY_RANK: Record<FindingSeverity, number> = {
  error: 0,
  warning: 1,
  info: 2,
  hint: 3,
};

/** The Error List's groups: errors, warnings, and everything informational. */
type SeverityGroup = "errors" | "warnings" | "messages";

/** Singular and plural for each toggle's label, the way the studio's Error List spells them. */
const FILTER_LABELS: Record<SeverityGroup, [string, string]> = {
  errors: ["Error", "Errors"],
  warnings: ["Warning", "Warnings"],
  messages: ["Message", "Messages"],
};

function severityGroup(severity: FindingSeverity): SeverityGroup {
  switch (severity) {
    case "error":
      return "errors";
    case "warning":
      return "warnings";
    default:
      return "messages";
  }
}

export class Shell {
  private readonly handlers: ShellHandlers;
  private readonly shell: HTMLElement;
  private readonly panelList: HTMLElement;
  private readonly problemsFilters: HTMLElement;

  /** The four sections around the editor, and everything docked in them. */
  private readonly docks: PanelDocks;

  /** Which severity groups the problems list shows. All on until the developer says otherwise. */
  private readonly severityFilters: Record<SeverityGroup, boolean> = {
    errors: true,
    warnings: true,
    messages: true,
  };
  private readonly statusPosition: HTMLElement;
  private readonly statusModule: HTMLElement;
  private readonly statusNotice: HTMLElement;
  private noticeTimer: ReturnType<typeof setTimeout> | undefined;
  private readonly explorer: Explorer;
  private readonly menubar: Menubar;
  private readonly problemsBody: HTMLElement;
  private readonly immediateBody: HTMLElement;
  private readonly immediateLog: HTMLElement;
  private readonly immediateInput: HTMLInputElement;
  private readonly localsBody: HTMLElement;
  private readonly watchBody: HTMLElement;
  private readonly watchTable: HTMLElement;
  private readonly toolbarRoot: HTMLElement;
  private readonly localsContext: HTMLElement;
  private readonly localsTable: HTMLElement;

  /** Lines entered in the Immediate panel, newest last, walked with the arrow keys. */
  private readonly history: string[] = [];
  private historyIndex = 0;

  /** True between submitting a line and its first answer, when focus belongs back at the prompt. */
  private immediateAwaitingResult = false;

  private findings: ShellFinding[] = [];

  /** Closes the host is holding for an answer about unsaved changes, asked one at a time. */
  private readonly closeConfirms: TabIdentity[] = [];

  /** The close currently being asked about, so a repeated Ctrl+W does not ask twice. */
  private askedCloseConfirm: TabIdentity | null = null;

  /** A rename asked for from a menu; focused when the properties for it arrive. */
  private pendingRename: string | null = null;

  private readonly propertiesList: HTMLElement;
  private readonly propertiesObject: HTMLElement;
  private propertiesComponent = "";
  private propertiesKind = "";
  private properties: ShellProperty[] = [];

  constructor(root: HTMLElement, handlers: ShellHandlers) {
    this.handlers = handlers;

    this.shell = root.querySelector("#shell") as HTMLElement;
    this.statusPosition = root.querySelector("#status-position") as HTMLElement;
    this.statusModule = root.querySelector("#status-module") as HTMLElement;
    this.statusNotice = root.querySelector("#status-notice") as HTMLElement;

    this.explorer = new Explorer(root.querySelector("#sidebar-tree") as HTMLElement, {
      select: (name) => handlers.selectComponent(name),
      open: (name, workbook) => handlers.activateModule(name, workbook),
      context: (name, kind, x, y) => this.componentMenu(name, kind, x, y),
      projectContext: (project, x, y) => this.workbookMenu(project, x, y),
      outline: (module, workbook) => handlers.requestOutline(module, workbook),
      openProcedure: (module, line, workbook) => handlers.navigate(module, line, 1, true, workbook),
      trace: (text) => handlers.trace(text),
    });

    this.menubar = new Menubar(root.querySelector("#menubar") as HTMLElement, {
      request: (path) => handlers.menuRequest(path),
      execute: (path) => handlers.menuExecute(path),
      closed: () => handlers.menuClosed(),
    });

    this.propertiesList = root.querySelector("#properties-list") as HTMLElement;
    this.propertiesObject = root.querySelector("#properties-object") as HTMLElement;

    this.toolbarRoot = root.querySelector("#toolbar") as HTMLElement;
    buildToolbar(
      this.toolbarRoot,
      (command) => handlers.command(command),
      (command) => handlers.commandAvailable(command));
    this.panelList = root.querySelector("#panel-list") as HTMLElement;
    this.problemsFilters = root.querySelector("#problems-filters") as HTMLElement;

    // The severity toggles: each shows its count always and filters the list when pressed out,
    // the way the studio's Error List reads. State lives here, pressed-ness on the button.
    this.problemsFilters.addEventListener("click", (event) => {
      const button = (event.target as HTMLElement).closest("[data-severity-filter]") as HTMLElement | null;
      const group = button?.dataset.severityFilter as SeverityGroup | undefined;
      if (!button || !group) {
        return;
      }

      this.severityFilters[group] = !this.severityFilters[group];
      button.setAttribute("aria-pressed", String(this.severityFilters[group]));
      this.renderPanel();
    });

    this.problemsBody = root.querySelector("#panel-list") as HTMLElement;
    this.immediateBody = root.querySelector("#immediate") as HTMLElement;
    this.immediateLog = root.querySelector("#immediate-log") as HTMLElement;
    this.immediateInput = root.querySelector("#immediate-input") as HTMLInputElement;
    this.localsBody = root.querySelector("#locals") as HTMLElement;
    this.watchBody = root.querySelector("#watch") as HTMLElement;
    this.watchTable = root.querySelector("#watch-table") as HTMLElement;
    this.localsContext = root.querySelector("#locals-context") as HTMLElement;
    this.localsTable = root.querySelector("#locals-table") as HTMLElement;

    // Every tool pane is a SEAT the dock system moves around: its title tab and its body
    // elements travel together. Problems brings its filter row, and the explorer and the
    // properties inspector are panes like any other now — dockable to any side, tabbable
    // onto any group (the developer's ask, 2026-08-06).
    const seat = (name: string, title: string, bodies: HTMLElement[], onShown?: () => void): PanelSeat => ({
      name,
      title,
      tab: this.makePanelTab(name, title),
      bodies,
      ...(onShown ? { onShown } : {}),
    });

    const seats: PanelSeat[] = [
      { ...seat("explorer", "Explorer", [root.querySelector("#sidebar-tree") as HTMLElement]), permanent: true },
      seat("properties", "Properties", [
        root.querySelector("#properties-object") as HTMLElement,
        this.propertiesList,
      ]),
      seat("problems", "Problems", [this.problemsFilters, this.problemsBody]),
      seat("immediate", "Immediate", [this.immediateBody], () => this.immediateInput.focus()),
      seat("locals", "Locals", [this.localsBody]),
      seat("watch", "Watch", [this.watchBody]),
    ];

    const dockOf = (side: string) => root.querySelector(`#dock-${side}`) as HTMLElement;
    const splitterOf = (side: string) => root.querySelector(`#dock-${side}-splitter`) as HTMLElement;

    this.docks = new PanelDocks(
      { left: dockOf("left"), right: dockOf("right"), top: dockOf("top"), bottom: dockOf("bottom") },
      { left: splitterOf("left"), right: splitterOf("right"), top: splitterOf("top"), bottom: splitterOf("bottom") },
      root.querySelector("#editor-area") as HTMLElement,
      // Stands in the editor area's place while no module is open, when the area itself is
      // display:none and measures nothing.
      root.querySelector("#empty-view") as HTMLElement | null,
      root.querySelector("#pane-bodies") as HTMLElement,
      seats,
      {
        visibilityChanged: (name, visible) => handlers.panelChanged(name, visible),
        layoutChanged: () => handlers.layoutChanged(),
      });

    this.installWatchActions(root);
    this.installImmediate();
    this.setLocals(false, null, []);

    this.panelList.addEventListener("click", (event) => {
      const row = (event.target as HTMLElement).closest("[data-line]") as HTMLElement | null;
      this.goTo(row);
    });

    this.panelList.addEventListener("contextmenu", (event) => {
      const row = (event.target as HTMLElement).closest("[data-line]") as HTMLElement | null;
      if (!row) {
        return;
      }

      event.preventDefault();
      const message = row.querySelector(".message")?.textContent ?? "";
      showContextMenu(event.clientX, event.clientY, [
        { label: "Go To", run: () => this.goTo(row) },
        { label: "Copy Message", run: () => void navigator.clipboard.writeText(message) },
      ]);
    });

    this.immediateLog.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      showContextMenu(event.clientX, event.clientY, [
        { label: "Clear", run: () => this.immediateLog.replaceChildren() },
      ]);
    });

    // The visible twin of the context menu's Clear, sitting in the entry row. It clears
    // this panel's history only; the hidden native window keeps its text, and the mirror
    // appends only what is new, so nothing cleared ever comes back on its own.
    (root.querySelector("#immediate-clear") as HTMLButtonElement).addEventListener("click", () => {
      this.immediateLog.replaceChildren();
      this.immediateInput.focus();
    });

    this.panelList.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        this.goTo(document.activeElement as HTMLElement | null);
      }
    });
  }

  /** The panes, for the settings dialog's visibility section. */
  /** The tree itself, for the debug api's `ui` route, which reports its state and drives it. */
  explorerTree(): Explorer {
    return this.explorer;
  }

  paneVisibility(): PaneVisibilityControl {
    return {
      list: () => this.docks.paneStates(),
      setOpen: (name, open) => {
        if (open) {
          this.docks.open(name);
        } else {
          this.docks.close(name);
        }
      },
    };
  }

  /** One pane's title tab, the handle it is dragged by. */
  private makePanelTab(name: string, title: string): HTMLButtonElement {
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = "panel-tab";
    tab.dataset.panel = name;
    tab.textContent = title;
    tab.setAttribute("role", "tab");
    tab.setAttribute("aria-selected", "false");
    return tab;
  }

  /**
   * The active module changed, however it changed. The tabs themselves are the workspace's;
   * the shell keeps the status line and the explorer highlight honest.
   */
  setActiveModule(active: string | null, activeProject: string | null): void {
    this.statusModule.textContent = active ?? "";
    this.explorer.setActive(active, activeProject ?? undefined);
  }

  /** Replaces the project explorer's contents. */
  /** The workspace as last published, for the Object Browser. */
  currentProjects(): ExplorerProject[] {
    return this.explorer.snapshot();
  }

  setProjects(projects: ExplorerProject[]): void {
    this.explorer.setProjects(projects);
  }

  /** Takes one menu's items from the host: the bar for an empty path, a dropdown otherwise. */
  setMenu(path: number[], items: MenuItem[]): void {
    this.menubar.setItems(path, items);
  }

  /** Shows or withdraws the menu bar; while it is withdrawn the native bar is the visible one. */
  setMenuBarVisible(visible: boolean): void {
    this.menubar.setVisible(visible);
  }

  /** Asks the host for the top-level menus. Called once the transport is up. */
  requestMenus(): void {
    this.menubar.refresh();
  }

  /**
   * Shows or hides the empty workspace, which is what the editor area becomes when every tab is
   * closed. The rest of the shell stays: the explorer is how a module gets opened again.
   */
  setWorkspaceEmpty(empty: boolean): void {
    this.shell.classList.toggle("empty", empty);
    this.statusModule.textContent = empty ? "" : this.statusModule.textContent;

    // Nothing open means nothing is being worked on, so the tree stops claiming otherwise and
    // folds all the way back: procedures and workbooks both. Opening anything unfolds its
    // workbook again on the way in.
    if (empty) {
      this.explorer.collapseAll();
    }

    this.handlers.layoutChanged();
  }

  /** Replaces the properties panel with the selected component's properties. */
  setProperties(component: string, kind: string, properties: ShellProperty[]): void {
    this.propertiesComponent = component;
    this.propertiesKind = kind;
    this.properties = properties;
    this.renderProperties();

    // A rename that was asked for lands here, when the name field for it actually exists.
    if (this.pendingRename === component) {
      this.pendingRename = null;
      const input = this.propertiesList.querySelector<HTMLInputElement>("input");
      input?.focus();
      input?.select();
    }
  }

  /** Brings the Properties pane forward wherever it is docked, and puts focus in it. */
  revealProperties(): void {
    this.docks.reveal("properties");
    this.propertiesList.querySelector<HTMLInputElement>("input")?.focus();
  }

  private renderProperties(): void {
    this.propertiesList.replaceChildren();

    // The object header, naming the selection and its class the way the editor's own window does.
    this.propertiesObject.replaceChildren();
    const objectName = document.createElement("span");
    objectName.className = "prop-object-name";
    objectName.textContent = this.propertiesComponent;
    const objectKind = document.createElement("span");
    objectKind.className = "prop-object-kind";
    objectKind.textContent = this.propertiesKind;
    this.propertiesObject.append(objectName, objectKind);

    for (const property of this.properties) {
      const row = document.createElement("div");
      row.className = "prop-row";
      row.setAttribute("role", "listitem");

      const name = document.createElement("span");
      name.className = "prop-name";
      name.textContent = property.name;
      name.title = property.name;

      row.appendChild(name);

      if (property.writable && property.boolean) {
        // A boolean offers its two values rather than a text field that trusts spelling.
        const select = document.createElement("select");
        select.className = "prop-value";
        select.setAttribute("aria-label", `${property.name} of ${this.propertiesComponent}`);

        for (const option of ["True", "False"]) {
          const choice = document.createElement("option");
          choice.value = option;
          choice.textContent = option;
          choice.selected = option === property.value;
          select.appendChild(choice);
        }

        select.addEventListener("change", () => {
          property.value = select.value;
          this.handlers.editProperty(this.propertiesComponent, property.name, select.value);
        });

        row.appendChild(select);
      } else if (property.writable) {
        const input = document.createElement("input");
        input.className = "prop-value";
        input.type = "text";
        input.value = property.value;
        input.spellcheck = false;
        input.setAttribute("aria-label", `${property.name} of ${this.propertiesComponent}`);

        // Committed when the developer is done, not per keystroke: Enter commits and stays,
        // leaving commits, Escape puts the truth back.
        input.addEventListener("keydown", (event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            this.commitProperty(property, input);
          } else if (event.key === "Escape") {
            event.preventDefault();
            input.value = property.value;
            input.blur();
          }
        });
        input.addEventListener("blur", () => this.commitProperty(property, input));

        row.appendChild(input);
      } else {
        const value = document.createElement("span");
        value.className = "prop-value readonly";
        value.textContent = property.value;
        value.title = property.value;
        row.appendChild(value);
      }

      this.propertiesList.appendChild(row);
    }
  }

  private commitProperty(property: ShellProperty, input: HTMLInputElement): void {
    if (input.value === property.value) {
      return;
    }

    // Remembered as sent, so blurring after Enter does not send the same edit twice. The host
    // answers with the real state either way, which also corrects a refused edit.
    property.value = input.value;
    this.handlers.editProperty(this.propertiesComponent, property.name, input.value);
  }

  /** Appends a line to the Immediate panel's output. */
  appendImmediate(text: string, kind: "echo" | "result" | "failed" = "result"): void {
    const line = document.createElement("div");
    line.className = `immediate-line ${kind}`;
    line.textContent = text;

    this.immediateLog.appendChild(line);

    // Kept at the bottom, because the newest line is the answer to what was just asked.
    this.immediateLog.scrollTop = this.immediateLog.scrollHeight;

    // The answer to a submitted line puts focus back at the prompt, and only that: output that
    // arrives on its own must not yank focus away from wherever the developer is typing.
    if (kind !== "echo" && this.immediateAwaitingResult) {
      this.immediateAwaitingResult = false;
      this.immediateInput.focus();
    }
  }

  /** Brings the Immediate pane forward wherever it is docked. */
  showImmediate(): void {
    this.docks.reveal("immediate");
    this.immediateInput.focus();
  }

  /** Brings the Locals pane forward wherever it is docked. */
  showLocalsPanel(): void {
    this.docks.reveal("locals");
  }

  /**
   * Replaces the Locals panel content. Stopped false is the idle state. Stopped true with no
   * rows is a break with nothing readable in scope — the panel must not claim "not stopped"
   * while the editor sits at a breakpoint, whatever the reader managed to see.
   */
  setLocals(stopped: boolean, context: string | null, rows: { expression: string; value: string; kind: string }[]): void {
    this.localsContext.textContent = context ?? "";
    this.localsContext.hidden = context === null;
    this.localsTable.replaceChildren();

    if (rows.length === 0) {
      const empty = document.createElement("div");
      empty.className = "locals-empty";
      empty.textContent = stopped
        ? "No variables to show."
        : "Not stopped. Variables appear here in break mode.";
      this.localsTable.appendChild(empty);
      return;
    }

    const header = document.createElement("div");
    header.className = "locals-row locals-header";
    header.setAttribute("role", "row");
    for (const title of ["Expression", "Value", "Type"]) {
      const cell = document.createElement("span");
      cell.setAttribute("role", "columnheader");
      cell.textContent = title;
      header.appendChild(cell);
    }
    this.localsTable.appendChild(header);

    for (const row of rows) {
      const line = document.createElement("div");
      line.className = "locals-row";
      line.setAttribute("role", "row");

      for (const text of [row.expression, row.value, row.kind]) {
        const cell = document.createElement("span");
        cell.setAttribute("role", "cell");
        cell.textContent = text;
        cell.title = text;
        line.appendChild(cell);
      }

      this.localsTable.appendChild(line);
    }
  }

  /**
   * The editor's debug mode — "design", "run", or "break". Controls that only mean
   * something stopped (the Call Stack button) grey with it, honestly, instead of
   * clicking into silence.
   */
  setDebugMode(mode: string): void {
    const stopped = mode === "break";
    for (const button of this.toolbarRoot.querySelectorAll<HTMLButtonElement>("[data-needs-break]")) {
      button.disabled = !stopped;
    }
  }

  /**
   * Replaces the Watch panel content. Not stopped is the idle state; stopped with no rows
   * means no watches are set.
   */
  setWatches(stopped: boolean, rows: { expression: string; value: string; kind: string; context: string }[]): void {
    this.watchTable.replaceChildren();

    if (rows.length === 0) {
      const empty = document.createElement("div");
      empty.className = "locals-empty";
      empty.textContent = stopped
        ? "No watches. Add one from the Debug menu."
        : "Not stopped. Watches report here in break mode.";
      this.watchTable.appendChild(empty);
      return;
    }

    const header = document.createElement("div");
    header.className = "locals-row watch-row locals-header";
    header.setAttribute("role", "row");
    for (const title of ["Expression", "Value", "Type", "Context"]) {
      const cell = document.createElement("span");
      cell.setAttribute("role", "columnheader");
      cell.textContent = title;
      header.appendChild(cell);
    }
    this.watchTable.appendChild(header);

    for (const row of rows) {
      const line = document.createElement("div");
      line.className = "locals-row watch-row";
      line.setAttribute("role", "row");

      for (const text of [row.expression, row.value, row.kind, row.context]) {
        const cell = document.createElement("span");
        cell.setAttribute("role", "cell");
        cell.textContent = text;
        cell.title = text;
        line.appendChild(cell);
      }

      this.watchTable.appendChild(line);
    }
  }

  /**
   * The Watch panel's own triggers. Watches are created, edited, and deleted through the
   * editor's dialogs (decision 11: they are modal, and driving one invisibly can hang the
   * editor with no window the developer can reach), so these buttons summon those dialogs
   * directly. The panel is where a watch is looked at, so it is where the work belongs; the
   * menu route is retired with them.
   */
  private installWatchActions(root: HTMLElement): void {
    const run = (id: string) => this.handlers.command({
      id,
      target: "host",
      icon: "",
      label: "",
    });

    (root.querySelector("#watch-add") as HTMLButtonElement)
      .addEventListener("click", () => run("addWatch"));
    (root.querySelector("#watch-edit") as HTMLButtonElement)
      .addEventListener("click", () => run("editWatch"));
    (root.querySelector("#watch-quick") as HTMLButtonElement)
      .addEventListener("click", () => run("quickWatch"));
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
    this.immediateAwaitingResult = true;
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

  /**
   * The host is holding a tab close because the module has unsaved changes. Ask, and answer
   * the close with the choice: Save writes the workbook and closes, Don't Save puts the
   * module back to its saved text and closes, Cancel leaves everything as it is. Questions
   * queue one at a time — a Close Others across several dirty modules asks about each in
   * turn, the way answering one file at a time reads everywhere else.
   */
  confirmClose(name: string, project: string | null): void {
    // Asking twice about the same tab answers nothing twice: a repeated Ctrl+W while the
    // question is up would otherwise queue the same question behind itself.
    const mine = tabKey({ name, project });
    if ((this.askedCloseConfirm && tabKey(this.askedCloseConfirm) === mine)
      || this.closeConfirms.some((waiting) => tabKey(waiting) === mine)) {
      return;
    }

    this.closeConfirms.push({ name, project });
    this.showNextCloseConfirm();
  }

  private showNextCloseConfirm(): void {
    if (document.getElementById("close-confirm-backdrop")) {
      return;
    }

    const asked = this.closeConfirms.shift();
    if (!asked) {
      return;
    }

    this.askedCloseConfirm = asked;

    const backdrop = document.createElement("div");
    backdrop.id = "close-confirm-backdrop";

    const card = document.createElement("div");
    card.id = "close-confirm-card";
    card.setAttribute("role", "alertdialog");
    card.setAttribute("aria-modal", "true");
    card.setAttribute("aria-label", "Unsaved changes");

    const title = document.createElement("div");
    title.id = "close-confirm-title";
    title.textContent = `Do you want to save the changes you made to ${asked.name}?`;

    const detail = document.createElement("div");
    detail.id = "close-confirm-detail";
    detail.textContent = "Your changes will be lost if you don't save them.";

    const buttons = document.createElement("div");
    buttons.id = "close-confirm-buttons";

    const resolve = (action: "save" | "discard" | null): void => {
      this.askedCloseConfirm = null;
      document.removeEventListener("keydown", onKey, true);
      backdrop.remove();

      if (action) {
        this.handlers.closeModule(asked.name, asked.project ?? undefined, action);
      }

      if (this.closeConfirms.length > 0) {
        this.showNextCloseConfirm();
      } else {
        // The question took focus; the editor is where it belongs afterwards.
        this.handlers.menuClosed();
      }
    };

    const button = (label: string, action: "save" | "discard" | null, primary = false): HTMLButtonElement => {
      const control = document.createElement("button");
      control.type = "button";
      control.className = primary ? "close-confirm-button primary" : "close-confirm-button";
      control.textContent = label;
      control.addEventListener("click", () => resolve(action));
      buttons.appendChild(control);
      return control;
    };

    const save = button("Save", "save", true);
    button("Don't Save", "discard");
    button("Cancel", null);

    const onKey = (event: KeyboardEvent): void => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        resolve(null);
      }
    };

    // A click beside the card is a Cancel: the safe answer for a gesture that was not one
    // of the three, and the same dismissal the settings card gives.
    backdrop.addEventListener("click", (event) => {
      if (event.target === backdrop) {
        resolve(null);
      }
    });
    document.addEventListener("keydown", onKey, true);

    card.append(title, detail, buttons);
    backdrop.appendChild(card);
    document.body.appendChild(backdrop);
    save.focus();
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

  private goTo(row: HTMLElement | null): void {
    if (!row?.dataset.line) {
      return;
    }
    this.handlers.navigate(
      row.dataset.moduleName ?? "",
      Number(row.dataset.line),
      Number(row.dataset.column ?? "1"),
      false,
      row.dataset.project || undefined,
    );
  }

  /**
   * The menu for one explorer item, shaped by what the component is. Options that make no sense
   * for the class are left out rather than disabled; the host's own operations arrive as the
   * classes grow them.
   */
  private componentMenu(name: string, kind: number, x: number, y: number): void {
    // A document or a form is an object with code behind it; a module is only its code.
    const openLabel = kind === 100 || kind === 3 ? "Open Code" : "Open";

    const items: ContextMenuItem[] = [
      { label: openLabel, run: () => this.handlers.activateModule(name) },
      {},
      { label: "Rename…", run: () => this.beginRename(name, null) },
    ];

    if (this.handlers.moduleIsOpen(name)) {
      items.push({}, { label: "Close", run: () => this.handlers.closeModule(name) });
    }

    showContextMenu(x, y, items);
  }

  /**
   * The workbook's menu, grouped the way the companion editor groups it: what can be created in
   * the workbook first, then the dialogs that belong to the project itself.
   */
  private workbookMenu(project: string, x: number, y: number): void {
    showContextMenu(x, y, [
      { label: "New Module", run: () => this.handlers.insertComponent(1, project) },
      { label: "New Class Module", run: () => this.handlers.insertComponent(2, project) },
      { label: "New UserForm", run: () => this.handlers.insertComponent(3, project) },
      {},
      { label: "References...", run: () => this.hostCommand("references") },
      { label: "Project Properties...", run: () => this.hostCommand("projectProperties") },
    ]);
  }

  private hostCommand(id: string): void {
    this.handlers.command({ id, target: "host", icon: "", label: id });
  }

  /**
   * Renames a module, and with it everything in the workbook that names it.
   *
   * This used to open the Properties panel and leave the developer to retype "(Name)" — which
   * renames the component and leaves every `OldName.Something` in every other module pointing at
   * a module that no longer exists. The rename goes through the host now; the panel is still
   * where a name can be edited by hand for anyone who wants only that.
   */
  private beginRename(name: string, workbook: string | null): void {
    const wanted = window.prompt(`Rename ${name} to:`, name);
    if (wanted === null) {
      return;
    }

    const trimmed = wanted.trim();
    if (trimmed.length === 0 || trimmed === name) {
      return;
    }

    this.handlers.renameModule(name, workbook, trimmed);
  }

  private renderPanel(): void {
    // The toggles always carry the full counts — a filtered-out severity still says how many it
    // is hiding, which is what makes toggling it back on an informed act.
    const totals: Record<SeverityGroup, number> = { errors: 0, warnings: 0, messages: 0 };
    for (const finding of this.findings) {
      totals[severityGroup(finding.severity)]++;
    }

    for (const button of this.problemsFilters.querySelectorAll<HTMLElement>("[data-severity-filter]")) {
      const group = button.dataset.severityFilter as SeverityGroup;
      const label = button.querySelector(".filter-count");
      if (label) {
        label.textContent = `${totals[group]} ${FILTER_LABELS[group][totals[group] === 1 ? 0 : 1]}`;
      }
    }

    this.panelList.replaceChildren();

    // Worst first, then by where they are, so the order does not depend on which module the
    // engine happened to analyse first. Severities the developer pressed out stay out of the
    // list; every other surface — badges, squiggles, tree — keeps the full picture.
    const sorted = [...this.findings]
      .filter((finding) => this.severityFilters[severityGroup(finding.severity)])
      .sort((a, b) =>
        SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity]
        || a.module.localeCompare(b.module)
        || a.line - b.line
        || a.column - b.column);

    // A module name findings know from more than one workbook names its workbook in each row,
    // the way the tabs do; a unique name stays bare.
    const findingHomes = new Map<string, Set<string>>();
    for (const finding of this.findings) {
      const lower = finding.module.toLowerCase();
      if (!findingHomes.has(lower)) {
        findingHomes.set(lower, new Set());
      }
      findingHomes.get(lower)!.add((finding.project ?? "").toLowerCase());
    }

    for (const finding of sorted) {
      const collides = (findingHomes.get(finding.module.toLowerCase())?.size ?? 0) > 1;

      const row = document.createElement("div");
      row.className = `row ${finding.severity}`;
      row.tabIndex = 0;
      row.dataset.moduleName = finding.module;
      row.dataset.line = String(finding.line);
      row.dataset.column = String(finding.column);
      row.dataset.project = finding.project ?? "";

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
      where.textContent = `${finding.module}`
        + (collides && finding.project ? ` (${finding.project})` : "")
        + ` (${finding.line}, ${finding.column})`
        + (finding.code ? `   ${finding.code}` : "");

      body.append(message, where);
      row.append(mark, body);
      this.panelList.appendChild(row);
    }

    // Counts per component change with the findings, so the tree follows them. Filed by
    // (workbook, module), because a count belongs to one workbook's module and a shared name
    // must not pool them. The tab badges are the workspace's, fed the same findings.
    const counts = new Map<string, number>();
    for (const finding of this.findings) {
      const key = problemCountKey(finding.project ?? null, finding.module);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }

    this.explorer.setProblemCounts(counts);
  }
}
