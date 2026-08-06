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

/** Whether two workbook names mean the same workbook; null and empty read as "unknown". */
function sameProject(a: string | null | undefined, b: string | null | undefined): boolean {
  return (a ?? "").toLowerCase() === (b ?? "").toLowerCase();
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

/** Smallest useful panel: the header and about one finding. */
const MIN_PANEL_HEIGHT = 60;

/** The editor never gets less than this, however far the splitter is dragged. */
const MIN_EDITOR_HEIGHT = 80;

/** Smallest useful project explorer: a component name without truncation. */
const MIN_SIDEBAR_WIDTH = 120;

/** Smallest useful properties pane: the header and about two rows. */
const MIN_PROPERTIES_HEIGHT = 64;

/** The tree never gets less than this, however far the properties splitter is dragged. */
const MIN_TREE_HEIGHT = 100;

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
  private readonly problemsFilters: HTMLElement;
  private readonly panelToggle: HTMLButtonElement;

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
  private readonly sidebarSplitter: HTMLElement;
  private readonly explorer: Explorer;
  private readonly menubar: Menubar;
  private readonly panelTabs: HTMLElement;
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

  private active: string | null = null;

  /** The workbook the active module belongs to, so a shared name activates as itself. */
  private activeProject: string | null = null;

  private findings: ShellFinding[] = [];

  /**
   * The tabs in the order the developer arranged them. The host reports WHICH modules are open;
   * where each one sits in the strip is the developer's, and survives every host update. A tab
   * is a (module, workbook) pair: two workbooks holding a Module1 are two tabs.
   */
  private tabOrder: TabIdentity[] = [];

  /** True from a drag until the click it would otherwise become has been swallowed. */
  private dragSuppressesClick = false;

  /** The tab whose close box the pointer went down on, held as data a rebuild cannot destroy. */
  private pressedClose: { name: string; project: string | null } | null = null;

  /** Closes the host is holding for an answer about unsaved changes, asked one at a time. */
  private readonly closeConfirms: TabIdentity[] = [];

  /** The close currently being asked about, so a repeated Ctrl+W does not ask twice. */
  private askedCloseConfirm: TabIdentity | null = null;

  /** What the tab strip last rendered, so an echo that changes nothing rebuilds nothing. */
  private lastTabsKey: string | null = null;

  /** A rename asked for from a menu; focused when the properties for it arrive. */
  private pendingRename: string | null = null;
  private panelOpen = true;
  private panelHeight = 180;
  private sidebarWidth = 260;
  private shown = "problems";

  private readonly propertiesSection: HTMLElement;
  private readonly propertiesHead: HTMLButtonElement;
  private readonly propertiesList: HTMLElement;
  private readonly propertiesObject: HTMLElement;
  private readonly propertiesSplitter: HTMLElement;
  private propertiesComponent = "";
  private propertiesKind = "";
  private properties: ShellProperty[] = [];
  private propertiesOpen = true;
  private propertiesHeight = 300;

  constructor(root: HTMLElement, handlers: ShellHandlers) {
    this.handlers = handlers;

    this.shell = root.querySelector("#shell") as HTMLElement;
    this.splitter = root.querySelector("#panel-splitter") as HTMLElement;
    this.tabStrip = root.querySelector("#tabs") as HTMLElement;
    this.statusPosition = root.querySelector("#status-position") as HTMLElement;
    this.statusModule = root.querySelector("#status-module") as HTMLElement;
    this.statusNotice = root.querySelector("#status-notice") as HTMLElement;

    this.sidebarSplitter = root.querySelector("#sidebar-splitter") as HTMLElement;
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

    this.propertiesSection = root.querySelector("#properties") as HTMLElement;
    this.propertiesHead = root.querySelector("#properties-head") as HTMLButtonElement;
    this.propertiesList = root.querySelector("#properties-list") as HTMLElement;
    this.propertiesObject = root.querySelector("#properties-object") as HTMLElement;
    this.propertiesSplitter = root.querySelector("#properties-splitter") as HTMLElement;
    this.propertiesHead.addEventListener("click", () => this.togglePropertiesOpen());
    this.installPropertiesSplitter();

    this.toolbarRoot = root.querySelector("#toolbar") as HTMLElement;
    buildToolbar(
      this.toolbarRoot,
      (command) => handlers.command(command),
      (command) => handlers.commandAvailable(command));
    this.panel = root.querySelector("#panel") as HTMLElement;
    this.panelList = root.querySelector("#panel-list") as HTMLElement;
    this.problemsFilters = root.querySelector("#problems-filters") as HTMLElement;
    this.panelToggle = root.querySelector("#panel-toggle") as HTMLButtonElement;

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

    this.panelTabs = root.querySelector("#panel-tabs") as HTMLElement;
    this.problemsBody = root.querySelector("#panel-list") as HTMLElement;
    this.immediateBody = root.querySelector("#immediate") as HTMLElement;
    this.immediateLog = root.querySelector("#immediate-log") as HTMLElement;
    this.immediateInput = root.querySelector("#immediate-input") as HTMLInputElement;
    this.localsBody = root.querySelector("#locals") as HTMLElement;
    this.watchBody = root.querySelector("#watch") as HTMLElement;
    this.watchTable = root.querySelector("#watch-table") as HTMLElement;
    this.localsContext = root.querySelector("#locals-context") as HTMLElement;
    this.localsTable = root.querySelector("#locals-table") as HTMLElement;

    this.installWatchActions(root);
    this.installPanelTabs();
    this.installImmediate();
    this.setLocals(false, null, []);

    this.panelToggle.addEventListener("click", () => this.togglePanel());
    this.installSplitter();
    this.installSidebarSplitter();

    // One listener on the strip rather than one per tab: the tabs are rebuilt whenever the set of
    // open modules changes, and per-tab listeners would have to be torn down with them.
    this.tabStrip.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      const tab = target.closest("[data-module]") as HTMLElement | null;
      if (!tab?.dataset.module) {
        return;
      }

      // The close box belongs to the pointer path below, and a click on it must not fall
      // through to activation.
      if (target.closest(".tab-close")) {
        return;
      }

      // A drag that ends over a tab is not a click on it.
      if (this.dragSuppressesClick) {
        return;
      }

      this.handlers.activateModule(tab.dataset.module, tab.dataset.project || undefined);
    });

    // The X is armed at pointerdown and fired at pointerup, never on click. A click needs its
    // press and release to land on the same LIVE element, and the press itself can rebuild the
    // strip: pressing an unfocused surface focuses it, focus stirs the host, and a setModules
    // echo mid-press replaces the pressed element — the click never happens, and the X reads
    // as dead until a second try. (The second life of this bug; the first was a stale drag
    // flag.) The press identity is captured as DATA, which no rebuild can destroy; the release
    // only checks that the pointer is still over the same tab's X, so sliding off to cancel
    // still cancels, even though the element under the pointer may be a rebuilt twin.
    this.tabStrip.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      const target = event.target as HTMLElement;
      const tab = target.closest("[data-module]") as HTMLElement | null;
      this.pressedClose = target.closest(".tab-close") && tab?.dataset.module
        ? { name: tab.dataset.module, project: tab.dataset.project || null }
        : null;
    });

    this.tabStrip.addEventListener("pointerup", (event) => {
      const pressed = this.pressedClose;
      this.pressedClose = null;

      if (!pressed || event.button !== 0) {
        return;
      }

      const target = event.target as HTMLElement;
      const tab = target.closest("[data-module]") as HTMLElement | null;
      if (target.closest(".tab-close")
        && tab?.dataset.module === pressed.name
        && (tab.dataset.project || null) === pressed.project) {
        this.handlers.closeModule(pressed.name, pressed.project ?? undefined);
      }
    });

    this.tabStrip.addEventListener("pointercancel", () => {
      this.pressedClose = null;
    });

    // The middle button closes, the way every tabbed editor closes — any tab, focused or not.
    // The mousedown is claimed too, so the browser's middle-click autoscroll cannot swallow
    // the click before it becomes an auxclick.
    this.tabStrip.addEventListener("mousedown", (event) => {
      if (event.button === 1 && (event.target as HTMLElement).closest("[data-module]")) {
        event.preventDefault();
      }
    });

    this.tabStrip.addEventListener("auxclick", (event) => {
      if (event.button !== 1) {
        return;
      }

      const tab = (event.target as HTMLElement).closest("[data-module]") as HTMLElement | null;
      if (tab?.dataset.module) {
        event.preventDefault();
        this.handlers.closeModule(tab.dataset.module, tab.dataset.project || undefined);
      }
    });

    // Ctrl+W closes the active tab from anywhere in the surface. The host's key hook claims it
    // first when it is listening; this is the page's own answer for every moment it is not, so
    // the shortcut never depends on which corner of the surface has focus.
    document.addEventListener("keydown", (event) => {
      if (!event.ctrlKey || event.shiftKey || event.altKey) {
        return;
      }
      if (event.code !== "KeyW" && event.code !== "F4") {
        return;
      }

      event.preventDefault();
      if (this.active) {
        this.handlers.closeModule(this.active, this.activeProject ?? undefined);
      }
    }, { capture: true });

    this.tabStrip.addEventListener("contextmenu", (event) => {
      const tab = (event.target as HTMLElement).closest("[data-module]") as HTMLElement | null;
      if (tab?.dataset.module) {
        event.preventDefault();
        this.tabMenu(tab.dataset.module, tab.dataset.project || null, event.clientX, event.clientY);
      }
    });

    this.installTabDrag();

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

  /** Which workbooks have unsaved changes, by lowercased display name. */
  private dirtyProjects = new Map<string, boolean>();

  /** Replaces the tab strip. The active module is highlighted, not just listed. */
  setModules(
    modules: string[],
    projects: (string | null)[],
    active: string | null,
    activeProject: string | null,
    dirty: boolean[] = [],
  ): void {
    this.active = active;
    this.activeProject = activeProject;

    // Dirty is a workbook fact — the editor saves all of a workbook's modules together — so
    // one flag serves every tab the workbook owns.
    this.dirtyProjects = new Map();
    modules.forEach((_, index) => {
      const key = (projects[index] ?? "").toLowerCase();
      if (dirty[index]) {
        this.dirtyProjects.set(key, true);
      }
    });

    const open = modules.map((name, index) => ({ name, project: projects[index] ?? null }));
    const openKeys = new Set(open.map(tabKey));
    const heldKeys = new Set(this.tabOrder.map(tabKey));

    // The host says which modules are open; the developer says where their tabs sit. Tabs that
    // are still open keep their positions, new ones join at the end, closed ones leave.
    this.tabOrder = [
      ...this.tabOrder.filter((tab) => openKeys.has(tabKey(tab))),
      ...open.filter((tab) => !heldKeys.has(tabKey(tab))),
    ];

    this.statusModule.textContent = active ?? "";
    this.explorer.setActive(active, activeProject ?? undefined);
    this.renderTabs();
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

  /** Opens the properties section and puts focus in it, for the menu route to it. */
  revealProperties(): void {
    if (!this.propertiesOpen) {
      this.togglePropertiesOpen();
    }

    this.propertiesList.querySelector<HTMLInputElement>("input")?.focus();
  }

  /**
   * Makes the divider above the properties pane draggable, the same way as the other two.
   * The pane is docked at the bottom of the sidebar, so dragging up makes it taller.
   */
  private installPropertiesSplitter(): void {
    this.propertiesSplitter.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      event.preventDefault();
      this.propertiesSplitter.setPointerCapture(event.pointerId);

      const startY = event.clientY;
      const startHeight = this.propertiesHeight;

      const move = (moved: PointerEvent) => this.setPropertiesHeight(startHeight - (moved.clientY - startY));
      const end = (ended: PointerEvent) => {
        this.propertiesSplitter.releasePointerCapture(ended.pointerId);
        this.propertiesSplitter.removeEventListener("pointermove", move);
        this.propertiesSplitter.removeEventListener("pointerup", end);
        this.propertiesSplitter.removeEventListener("pointercancel", end);
      };

      this.propertiesSplitter.addEventListener("pointermove", move);
      this.propertiesSplitter.addEventListener("pointerup", end);
      this.propertiesSplitter.addEventListener("pointercancel", end);
    });

    this.propertiesSplitter.addEventListener("keydown", (event) => {
      if (event.key === "ArrowUp") {
        this.setPropertiesHeight(this.propertiesHeight + KEYBOARD_STEP);
      } else if (event.key === "ArrowDown") {
        this.setPropertiesHeight(this.propertiesHeight - KEYBOARD_STEP);
      } else {
        return;
      }

      event.preventDefault();
    });
  }

  private setPropertiesHeight(height: number): void {
    // Bounded against the sidebar, so the tree keeps a useful minimum and the pane cannot be
    // dragged out through the top of it.
    const sidebar = this.propertiesSection.parentElement as HTMLElement;
    const largest = Math.max(MIN_PROPERTIES_HEIGHT, sidebar.clientHeight - MIN_TREE_HEIGHT);

    this.propertiesHeight = Math.round(Math.min(largest, Math.max(MIN_PROPERTIES_HEIGHT, height)));
    sidebar.style.setProperty("--properties-height", `${this.propertiesHeight}px`);
  }

  private togglePropertiesOpen(): void {
    this.propertiesOpen = !this.propertiesOpen;
    this.propertiesList.hidden = !this.propertiesOpen;
    this.propertiesHead.setAttribute("aria-expanded", String(this.propertiesOpen));
    this.propertiesSection.classList.toggle("collapsed", !this.propertiesOpen);

    // A single collapsed row cannot be resized, so the handle goes with the body.
    this.propertiesSplitter.hidden = !this.propertiesOpen || this.properties.length === 0;
  }

  private renderProperties(): void {
    this.propertiesList.replaceChildren();
    this.propertiesSection.hidden = this.properties.length === 0;
    this.propertiesSplitter.hidden = this.properties.length === 0 || !this.propertiesOpen;

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

  /** Brings the Immediate panel forward, opening the panel if it was collapsed. */
  showImmediate(): void {
    this.selectPanel("immediate");

    if (!this.panelOpen) {
      this.togglePanel();
    }

    this.immediateInput.focus();
  }

  /** Brings the Locals panel forward, opening the panel if it was collapsed. */
  showLocalsPanel(): void {
    this.selectPanel("locals");

    if (!this.panelOpen) {
      this.togglePanel();
    }
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
    this.problemsFilters.hidden = name !== "problems";
    this.immediateBody.hidden = name !== "immediate";
    this.localsBody.hidden = name !== "locals";
    this.watchBody.hidden = name !== "watch";
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
      false,
      row.dataset.project || undefined,
    );
  }

  private renderTabs(): void {
    // Rebuilt only when something the render consumes has changed. The host echoes setModules
    // freely — a focus click alone can produce one — and every needless rebuild destroys the
    // elements a pointer might be pressing at that instant. The key covers everything drawn:
    // order, identity, the active pair, and each tab's badge count.
    const renderKey = this.tabOrder
      .map((tab) => tabKey(tab) + "\u0001" + this.findings.filter((f) => f.module === tab.name
        && (f.project == null || sameProject(f.project, tab.project))).length
        + (this.dirtyProjects.get((tab.project ?? "").toLowerCase()) === true ? "\u0004d" : ""))
      .join("\u0002")
      + "\u0003" + (this.active ?? "") + "\u0001" + (this.activeProject ?? "");

    if (renderKey === this.lastTabsKey) {
      return;
    }

    this.lastTabsKey = renderKey;
    this.tabStrip.replaceChildren();

    // A name two workbooks share earns its workbook in the label; a unique name stays bare.
    const nameCounts = new Map<string, number>();
    for (const tab of this.tabOrder) {
      const lower = tab.name.toLowerCase();
      nameCounts.set(lower, (nameCounts.get(lower) ?? 0) + 1);
    }

    for (const { name, project } of this.tabOrder) {
      const isActive = name === this.active && sameProject(project, this.activeProject);
      const collides = (nameCounts.get(name.toLowerCase()) ?? 0) > 1;
      const isDirty = this.dirtyProjects.get((project ?? "").toLowerCase()) === true;

      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "tab" + (isActive ? " active" : "") + (isDirty ? " dirty" : "");
      tab.dataset.module = name;
      tab.dataset.project = project ?? "";
      tab.textContent = collides && project ? `${name} — ${project}` : name;
      // The tab strip is a tab list for anything reading the page aloud, not a row of buttons.
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", String(isActive));

      // The badge counts this tab's own module: a finding that names its workbook counts only
      // on the matching tab, and one that cannot say counts wherever the name appears.
      const count = this.findings.filter((f) => f.module === name
        && (f.project == null || sameProject(f.project, project))).length;
      if (count > 0) {
        const badge = document.createElement("span");
        badge.className = "tab-badge";
        badge.textContent = String(count);
        // Read out rather than left as a bare number next to a name.
        badge.title = `${count} problem${count === 1 ? "" : "s"} in ${name}`;
        tab.appendChild(badge);
      }

      // The unsaved dot sits where the close box sits, the way the studio wears it: the dot
      // while the workbook is dirty, the X the moment the pointer arrives. Both are always in
      // the tree; the stylesheet decides which one shows.
      const unsaved = document.createElement("span");
      unsaved.className = "tab-dirty codicon codicon-circle-filled";
      unsaved.title = "Unsaved changes in this workbook";
      tab.appendChild(unsaved);

      const close = document.createElement("span");
      close.className = "tab-close codicon codicon-close";
      close.title = "Close (Ctrl+W)";
      close.setAttribute("role", "button");
      close.setAttribute("aria-label", `Close ${name}`);
      tab.appendChild(close);

      this.tabStrip.appendChild(tab);
    }
  }

  /**
   * Makes the tabs draggable into a new order.
   *
   * The dragged tab is moved in the DOM as the pointer crosses its neighbours' midpoints, so the
   * feedback is the reorder itself. The element is moved rather than re-rendered, because the
   * pointer capture that keeps the drag alive belongs to the element and would die with it.
   */
  private installTabDrag(): void {
    this.tabStrip.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      const target = event.target as HTMLElement;
      const tab = target.closest(".tab") as HTMLElement | null;
      if (!tab || target.closest(".tab-close")) {
        return;
      }

      const startX = event.clientX;
      const pointerId = event.pointerId;
      let moved = false;

      const move = (during: PointerEvent) => {
        // A few pixels of slack, so a click with a shaky hand is still a click.
        if (!moved && Math.abs(during.clientX - startX) < 5) {
          return;
        }

        if (!moved) {
          moved = true;
          this.dragSuppressesClick = true;
          tab.classList.add("dragging");

          try {
            tab.setPointerCapture(pointerId);
          } catch {
            // A pointer that has already gone cannot be captured; the window listeners still
            // finish the drag.
          }
        }

        const after = this.tabAfter(during.clientX, tab);
        if (after === null) {
          this.tabStrip.appendChild(tab);
        } else if (after !== tab) {
          this.tabStrip.insertBefore(tab, after);
        }
      };

      const end = () => {
        window.removeEventListener("pointermove", move);
        window.removeEventListener("pointerup", end);
        window.removeEventListener("pointercancel", end);
        tab.classList.remove("dragging");

        if (moved) {
          this.tabOrder = [...this.tabStrip.querySelectorAll<HTMLElement>("[data-module]")]
            .filter((t) => !!t.dataset.module)
            .map((t) => ({ name: t.dataset.module!, project: t.dataset.project || null }));

          // Cleared after the click this drag produces has already been ignored.
          setTimeout(() => {
            this.dragSuppressesClick = false;
          }, 0);
        }
      };

      window.addEventListener("pointermove", move);
      window.addEventListener("pointerup", end);
      // The host steals focus mid-press on all sorts of occasions — showing a module, following
      // a navigation — and an interrupted pointer stream ends in pointercancel, never pointerup.
      // Without this the suppression flag stayed raised forever.
      window.addEventListener("pointercancel", end);
    });
  }

  /** The tab the dragged one should sit before at this pointer position, or null for the end. */
  private tabAfter(x: number, dragging: HTMLElement): HTMLElement | null {
    for (const tab of this.tabStrip.querySelectorAll<HTMLElement>(".tab")) {
      if (tab === dragging) {
        continue;
      }

      const box = tab.getBoundingClientRect();
      if (x < box.left + box.width / 2) {
        return tab;
      }
    }

    return null;
  }

  /**
   * Activates the next or previous tab in the developer's order, wrapping at the ends, and says
   * which one it chose. Public because the key that means it never reaches the page: the browser
   * owns Ctrl+PageDown, so the host claims it and asks for this by name.
   */
  cycleTab(delta: number): string | null {
    if (this.tabOrder.length === 0) {
      return null;
    }

    const activeIndex = this.tabOrder.findIndex((tab) =>
      tab.name === this.active && sameProject(tab.project, this.activeProject));
    const current = Math.max(0, activeIndex);
    const next = this.tabOrder[(current + delta + this.tabOrder.length) % this.tabOrder.length];
    if (next && !(next.name === this.active && sameProject(next.project, this.activeProject))) {
      this.handlers.activateModule(next.name, next.project ?? undefined);
      return next.name;
    }

    return null;
  }

  private tabMenu(name: string, project: string | null, x: number, y: number): void {
    const mine = { name, project };
    showContextMenu(x, y, [
      { label: "Close", run: () => this.handlers.closeModule(name, project ?? undefined) },
      {
        label: "Close Others",
        enabled: this.tabOrder.length > 1,
        run: () => this.tabOrder.filter((other) => tabKey(other) !== tabKey(mine))
          .forEach((other) => this.handlers.closeModule(other.name, other.project ?? undefined)),
      },
      {
        label: "Close All",
        run: () => [...this.tabOrder]
          .forEach((other) => this.handlers.closeModule(other.name, other.project ?? undefined)),
      },
    ]);
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
      { label: "Rename", run: () => this.beginRename(name) },
    ];

    if (this.tabOrder.some((tab) => tab.name === name)) {
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

  /** Selects a component and puts focus in its name, once its properties arrive. */
  private beginRename(name: string): void {
    this.pendingRename = name;
    this.handlers.selectComponent(name);

    if (!this.propertiesOpen) {
      this.togglePropertiesOpen();
    }
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
        + (collides && finding.project ? ` — ${finding.project}` : "")
        + ` (${finding.line}, ${finding.column})`
        + (finding.code ? `   ${finding.code}` : "");

      body.append(message, where);
      row.append(mark, body);
      this.panelList.appendChild(row);
    }

    // Counts per component change with the findings, so both the strip and the tree are
    // rebuilt. Filed by (workbook, module), because a count belongs to one workbook's module
    // and a shared name must not pool them.
    const counts = new Map<string, number>();
    for (const finding of this.findings) {
      const key = problemCountKey(finding.project ?? null, finding.module);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }

    this.explorer.setProblemCounts(counts);
    this.renderTabs();
  }
}
