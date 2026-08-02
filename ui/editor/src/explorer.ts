/*
 * The project explorer.
 *
 * This replaces the host's own, which is closed rather than covered: the host will hide a tool
 * window on request, and a hidden window cannot be uncovered by anything the host does later. The
 * project itself is untouched, so everything that reads it keeps working.
 *
 * The shape is the companion editor's tree, behaviour included. Each workbook is its own unit,
 * named by its file; its modules sit flat beneath it in document, form, module, class order with
 * their kind spelled out; a module unfolds into its procedures, and clicking one goes to its
 * line. Expansion is an accordion: one module open at a time, following the module being edited,
 * and activating a module in another workbook moves the whole tree's attention there. A workbook
 * the developer collapsed by hand stays collapsed until the attention genuinely moves again.
 */

/** Component kinds, as the host numbers them. */
export const enum ComponentKind {
  StandardModule = 1,
  ClassModule = 2,
  Form = 3,
  ActiveXDesigner = 11,
  Document = 100,
}

export interface ExplorerComponent {
  name: string;
  kind: number;
}

export interface ExplorerProject {
  name: string;
  components: ExplorerComponent[];
}

/** One procedure of a module: the kind as the tree spells it, and its 1-based line. */
export interface ExplorerProcedure {
  name: string;
  kind: string;
  line: number;
}

interface KindMeta {
  /** The kind spelled the way the companion editor spells it beside a module's name. */
  type: string;
  icon: string;
  order: number;
}

const KIND_META: Record<number, KindMeta> = {
  [ComponentKind.Document]: { type: "document", icon: "symbol-namespace", order: 0 },
  [ComponentKind.Form]: { type: "userform", icon: "window", order: 1 },
  [ComponentKind.ActiveXDesigner]: { type: "designer", icon: "window", order: 1 },
  [ComponentKind.StandardModule]: { type: "standard", icon: "symbol-module", order: 2 },
  [ComponentKind.ClassModule]: { type: "class", icon: "symbol-class", order: 3 },
};

const KIND_FALLBACK: KindMeta = { type: "module", icon: "symbol-module", order: 4 };

function kindMeta(kind: number): KindMeta {
  return KIND_META[kind] ?? KIND_FALLBACK;
}

export interface ExplorerHandlers {
  /** Single click: the component becomes the selection, and nothing opens. */
  select(name: string): void;
  /** Double click, or Enter on the focused item: the component's code opens. */
  open(name: string): void;
  /** Right click on a component: the menu for its class, at this position. */
  context(name: string, kind: number, x: number, y: number): void;
  /** Right click on a workbook's row. */
  projectContext(project: string, x: number, y: number): void;
  /** A module's procedures, for its unfolded node. */
  outline(module: string): Promise<ExplorerProcedure[]>;
  /** A procedure was picked: go to its line in its module. */
  openProcedure(module: string, line: number): void;
}

export class Explorer {
  private readonly root: HTMLElement;
  private readonly handlers: ExplorerHandlers;

  private projects: ExplorerProject[] = [];
  private active: string | null = null;
  private selected: string | null = null;
  private problemCounts = new Map<string, number>();

  /** Which workbooks are open. Workbooks start closed; the first one is opened on arrival. */
  private readonly expandedWorkbooks = new Map<string, boolean>();

  /** The one module whose procedures are unfolded: the accordion. */
  private expandedModule: string | null = null;

  /** The workbook the attention is in, so collapsing others happens only when it moves. */
  private attentionWorkbook: string | null = null;

  /** Fetched procedures by module, dropped whenever the project set is republished. */
  private readonly outlines = new Map<string, ExplorerProcedure[]>();

  private firstProjectsSeen = false;

  constructor(root: HTMLElement, handlers: ExplorerHandlers) {
    this.root = root;
    this.handlers = handlers;

    // One listener for the whole tree. The tree is rebuilt whenever the project changes, and
    // per-item listeners would have to be torn down with it.
    this.root.addEventListener("click", (event) => {
      // The chevron toggles and does nothing else, so unfolding is never also an open.
      const toggle = (event.target as HTMLElement).closest("[data-toggle]") as HTMLElement | null;
      if (toggle?.dataset.toggle) {
        this.toggleModule(toggle.dataset.toggle);
        return;
      }

      const procedure = this.procedureAt(event);
      if (procedure) {
        this.handlers.openProcedure(procedure.module, procedure.line);
        return;
      }

      // A single click selects AND opens. The properties panel follows the selection either
      // way, but a click that only selected read as a tree that had stopped working: opening
      // on a single click is the ergonomic this product follows, deliberately unlike the
      // host's own tree, which asks for a double click.
      const name = this.componentAt(event);
      if (name) {
        this.selected = name;
        this.setExpandedModule(name);
        this.render();
        this.handlers.select(name);
        this.handlers.open(name);
        return;
      }

      const workbook = this.workbookAt(event);
      if (workbook) {
        this.expandedWorkbooks.set(workbook, !(this.expandedWorkbooks.get(workbook) ?? false));
        this.render();
      }
    });

    this.root.addEventListener("dblclick", (event) => {
      const name = this.componentAt(event);
      if (name) {
        this.handlers.open(name);
      }
    });

    this.root.addEventListener("keydown", (event) => {
      // A button already turns Enter into a click; this turns it into an open instead, so the
      // keyboard can do everything the mouse can. Workbook and procedure rows keep the click.
      if (event.key === "Enter") {
        const name = this.componentAt(event);
        if (name) {
          event.preventDefault();
          this.handlers.open(name);
        }
      }
    });

    this.root.addEventListener("contextmenu", (event) => {
      const item = (event.target as HTMLElement).closest("[data-component]") as HTMLElement | null;
      if (item?.dataset.component) {
        event.preventDefault();
        this.selected = item.dataset.component;
        this.render();
        this.handlers.select(item.dataset.component);
        this.handlers.context(
          item.dataset.component,
          Number(item.dataset.kind ?? "0"),
          event.clientX,
          event.clientY);
        return;
      }

      const workbook = this.workbookAt(event);
      if (workbook) {
        event.preventDefault();
        this.handlers.projectContext(workbook, event.clientX, event.clientY);
      }
    });
  }

  private componentAt(event: Event): string | null {
    const item = (event.target as HTMLElement).closest("[data-component]") as HTMLElement | null;
    return item?.dataset.component ?? null;
  }

  private workbookAt(event: Event): string | null {
    const row = (event.target as HTMLElement).closest("[data-project]") as HTMLElement | null;
    return row?.dataset.project ?? null;
  }

  private procedureAt(event: Event): { module: string; line: number } | null {
    const row = (event.target as HTMLElement).closest("[data-proc-module]") as HTMLElement | null;
    if (!row?.dataset.procModule) {
      return null;
    }
    return { module: row.dataset.procModule, line: Number(row.dataset.procLine ?? "1") };
  }

  setProjects(projects: ExplorerProject[]): void {
    this.projects = projects;

    // Procedure lists describe a project that just changed shape; ask again when unfolded.
    this.outlines.clear();

    // The first workbook opens itself so modules are visible without a click.
    const first = projects[0];
    if (!this.firstProjectsSeen && first) {
      this.firstProjectsSeen = true;
      if (!this.expandedWorkbooks.has(first.name)) {
        this.expandedWorkbooks.set(first.name, true);
      }
    }

    if (this.expandedModule) {
      void this.fetchOutline(this.expandedModule);
    }

    this.render();
  }

  setActive(name: string | null): void {
    // The tree follows the module being edited — but only when it genuinely changed. The host
    // republishes the module list on all sorts of occasions with the same active module, and
    // following every push would fold whatever the developer just unfolded by hand.
    const changed = name !== this.active;
    this.active = name;
    if (name && changed) {
      this.setExpandedModule(name);
    }
    this.render();
  }

  /** Problem counts by component, so a defect is visible without opening the module. */
  setProblemCounts(counts: Map<string, number>): void {
    this.problemCounts = counts;

    // Analysis just described the project again; an unfolded procedure list follows it, so a
    // procedure renamed or added while typing appears without any clicking.
    if (this.expandedModule) {
      void this.fetchOutline(this.expandedModule);
    }

    this.render();
  }

  /**
   * The accordion: this module's procedures unfold and every other module folds. Its workbook
   * opens; other workbooks close only when the attention genuinely moved between workbooks, so
   * a workbook collapsed by hand stays collapsed while work continues inside another.
   */
  private setExpandedModule(name: string): void {
    const owner = this.projects.find((project) =>
      project.components.some((component) => component.name === name));

    if (this.expandedModule !== name) {
      this.expandedModule = name;
      void this.fetchOutline(name);
    }

    if (owner) {
      if (this.attentionWorkbook !== owner.name) {
        if (this.attentionWorkbook !== null) {
          for (const project of this.projects) {
            this.expandedWorkbooks.set(project.name, project.name === owner.name);
          }
        }
        this.attentionWorkbook = owner.name;
      }
      this.expandedWorkbooks.set(owner.name, true);
    }
  }

  private toggleModule(name: string): void {
    if (this.expandedModule === name) {
      this.expandedModule = null;
    } else {
      this.setExpandedModule(name);
    }
    this.render();
  }

  private async fetchOutline(module: string): Promise<void> {
    const procedures = await this.handlers.outline(module);
    if (this.expandedModule !== module) {
      return;
    }
    this.outlines.set(module, procedures);
    this.render();
  }

  private render(): void {
    this.root.replaceChildren();

    for (const project of this.projects) {
      const isOpen = this.expandedWorkbooks.get(project.name) ?? false;
      this.root.appendChild(this.workbookRow(project.name, isOpen));

      if (!isOpen) {
        continue;
      }

      const members = [...project.components].sort((a, b) => {
        const left = kindMeta(a.kind);
        const right = kindMeta(b.kind);
        if (left.order !== right.order) {
          return left.order - right.order;
        }
        return a.name.localeCompare(b.name);
      });

      for (const component of members) {
        this.root.appendChild(this.item(component));

        if (component.name === this.expandedModule) {
          for (const procedure of this.outlines.get(component.name) ?? []) {
            this.root.appendChild(this.procedureRow(component.name, procedure));
          }
        }
      }
    }
  }

  private workbookRow(name: string, isOpen: boolean): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-workbook";
    button.dataset.project = name;
    button.setAttribute("role", "treeitem");
    button.setAttribute("aria-expanded", String(isOpen));

    const chevron = document.createElement("span");
    chevron.className = `codicon codicon-chevron-${isOpen ? "down" : "right"}`;
    chevron.setAttribute("aria-hidden", "true");

    const icon = document.createElement("span");
    icon.className = "codicon codicon-file-code";
    icon.setAttribute("aria-hidden", "true");

    button.append(chevron, icon, document.createTextNode(name));
    return button;
  }

  private item(component: ExplorerComponent): HTMLElement {
    const meta = kindMeta(component.kind);
    const isUnfolded = component.name === this.expandedModule;

    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item"
      + (component.name === this.active ? " active" : "")
      + (component.name === this.selected ? " selected" : "");
    button.dataset.component = component.name;
    button.dataset.kind = String(component.kind);
    button.setAttribute("role", "treeitem");
    button.setAttribute("aria-selected", String(component.name === this.selected));
    button.setAttribute("aria-expanded", String(isUnfolded));

    const chevron = document.createElement("span");
    chevron.className = `codicon codicon-chevron-${isUnfolded ? "down" : "right"} tree-twisty`;
    chevron.dataset.toggle = component.name;
    chevron.setAttribute("aria-hidden", "true");

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${meta.icon}`;
    glyph.setAttribute("aria-hidden", "true");

    const kind = document.createElement("span");
    kind.className = "tree-kind";
    kind.textContent = meta.type;

    button.append(chevron, glyph, document.createTextNode(component.name), kind);

    const problems = this.problemCounts.get(component.name) ?? 0;
    if (problems > 0) {
      const badge = document.createElement("span");
      badge.className = "tree-badge";
      badge.textContent = String(problems);
      badge.title = `${problems} problem${problems === 1 ? "" : "s"}`;
      button.appendChild(badge);
    }

    return button;
  }

  private procedureRow(module: string, procedure: ExplorerProcedure): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item tree-proc";
    button.dataset.procModule = module;
    button.dataset.procLine = String(procedure.line);
    button.setAttribute("role", "treeitem");

    const glyph = document.createElement("span");
    glyph.className = "codicon codicon-symbol-method";
    glyph.setAttribute("aria-hidden", "true");

    button.append(glyph, document.createTextNode(`${procedure.kind} ${procedure.name}`));
    return button;
  }
}
