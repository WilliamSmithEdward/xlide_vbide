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
import { Explorer, type ExplorerProject } from "./explorer.js";
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
  /** The developer closed a module's tab, however they did it. */
  closeModule(name: string): void;
  /** The developer asked for a new component: 1 module, 2 class module, 3 form. */
  insertComponent(kind: number): void;
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
  private readonly panelCount: HTMLElement;
  private readonly panelToggle: HTMLButtonElement;
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

  /** Lines entered in the Immediate panel, newest last, walked with the arrow keys. */
  private readonly history: string[] = [];
  private historyIndex = 0;

  private active: string | null = null;
  private findings: ShellFinding[] = [];

  /**
   * The tabs in the order the developer arranged them. The host reports WHICH modules are open;
   * where each one sits in the strip is the developer's, and survives every host update.
   */
  private tabOrder: string[] = [];

  /** True from a drag until the click it would otherwise become has been swallowed. */
  private dragSuppressesClick = false;

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
      open: (name) => handlers.activateModule(name),
      context: (name, kind, x, y) => this.componentMenu(name, kind, x, y),
      projectContext: (x, y) => this.projectMenu(x, y),
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
      // A drag that ends over a tab is not a click on it.
      if (this.dragSuppressesClick) {
        return;
      }

      const target = event.target as HTMLElement;
      const tab = target.closest("[data-module]") as HTMLElement | null;
      if (!tab?.dataset.module) {
        return;
      }

      if (target.closest(".tab-close")) {
        this.handlers.closeModule(tab.dataset.module);
        return;
      }

      this.handlers.activateModule(tab.dataset.module);
    });

    // The middle button closes, the way every tabbed editor closes.
    this.tabStrip.addEventListener("auxclick", (event) => {
      if (event.button !== 1) {
        return;
      }

      const tab = (event.target as HTMLElement).closest("[data-module]") as HTMLElement | null;
      if (tab?.dataset.module) {
        event.preventDefault();
        this.handlers.closeModule(tab.dataset.module);
      }
    });

    this.tabStrip.addEventListener("contextmenu", (event) => {
      const tab = (event.target as HTMLElement).closest("[data-module]") as HTMLElement | null;
      if (tab?.dataset.module) {
        event.preventDefault();
        this.tabMenu(tab.dataset.module, event.clientX, event.clientY);
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

    this.panelList.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        this.goTo(document.activeElement as HTMLElement | null);
      }
    });
  }

  /** Replaces the tab strip. The active module is highlighted, not just listed. */
  setModules(modules: string[], active: string | null): void {
    this.active = active;

    // The host says which modules are open; the developer says where their tabs sit. Tabs that
    // are still open keep their positions, new ones join at the end, closed ones leave.
    this.tabOrder = [
      ...this.tabOrder.filter((name) => modules.includes(name)),
      ...modules.filter((name) => !this.tabOrder.includes(name)),
    ];

    this.statusModule.textContent = active ?? "";
    this.explorer.setActive(active);
    this.renderTabs();
  }

  /** Replaces the project explorer's contents. */
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

    for (const name of this.tabOrder) {
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
        tab.classList.remove("dragging");

        if (moved) {
          this.tabOrder = [...this.tabStrip.querySelectorAll<HTMLElement>("[data-module]")]
            .map((t) => t.dataset.module)
            .filter((name): name is string => !!name);

          // Cleared after the click this drag produces has already been ignored.
          setTimeout(() => {
            this.dragSuppressesClick = false;
          }, 0);
        }
      };

      window.addEventListener("pointermove", move);
      window.addEventListener("pointerup", end);
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

    const current = Math.max(0, this.tabOrder.indexOf(this.active ?? ""));
    const next = this.tabOrder[(current + delta + this.tabOrder.length) % this.tabOrder.length];
    if (next && next !== this.active) {
      this.handlers.activateModule(next);
      return next;
    }

    return null;
  }

  private tabMenu(name: string, x: number, y: number): void {
    showContextMenu(x, y, [
      { label: "Close", run: () => this.handlers.closeModule(name) },
      {
        label: "Close Others",
        enabled: this.tabOrder.length > 1,
        run: () => this.tabOrder.filter((other) => other !== name)
          .forEach((other) => this.handlers.closeModule(other)),
      },
      { label: "Close All", run: () => [...this.tabOrder].forEach((other) => this.handlers.closeModule(other)) },
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

    if (this.tabOrder.includes(name)) {
      items.push({}, { label: "Close", run: () => this.handlers.closeModule(name) });
    }

    showContextMenu(x, y, items);
  }

  private projectMenu(x: number, y: number): void {
    showContextMenu(x, y, [
      { label: "Insert Module", run: () => this.handlers.insertComponent(1) },
      { label: "Insert Class Module", run: () => this.handlers.insertComponent(2) },
      { label: "Insert UserForm", run: () => this.handlers.insertComponent(3) },
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
