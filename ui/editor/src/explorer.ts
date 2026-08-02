/*
 * The project explorer.
 *
 * This replaces the host's own, which is closed rather than covered: the host will hide a tool
 * window on request, and a hidden window cannot be uncovered by anything the host does later. The
 * project itself is untouched, so everything that reads it keeps working.
 *
 * Each workbook is its own unit, the way the companion editor's tree shows them: the workbook's
 * file name at the root, its modules flat beneath it with their kind spelled out beside the name,
 * ordered document, form, module, class. There is no project node and there are no grouping rows;
 * the kind column carries what the groups used to say.
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
}

export class Explorer {
  private readonly root: HTMLElement;
  private readonly handlers: ExplorerHandlers;

  private projects: ExplorerProject[] = [];
  private active: string | null = null;
  private selected: string | null = null;
  private problemCounts = new Map<string, number>();

  /** Which workbooks are open in the tree; one never seen before starts open. */
  private readonly expanded = new Map<string, boolean>();

  constructor(root: HTMLElement, handlers: ExplorerHandlers) {
    this.root = root;
    this.handlers = handlers;

    // One listener for the whole tree. The tree is rebuilt whenever the project changes, and
    // per-item listeners would have to be torn down with it.
    //
    // A single click selects AND opens. The properties panel follows the selection either way,
    // but a click that only selected read as a tree that had stopped working: opening on a
    // single click is the ergonomic this product follows, deliberately unlike the host's own
    // tree, which asks for a double click.
    this.root.addEventListener("click", (event) => {
      const name = this.componentAt(event);
      if (name) {
        this.selected = name;
        this.render();
        this.handlers.select(name);
        this.handlers.open(name);
        return;
      }

      const workbook = this.workbookAt(event);
      if (workbook) {
        this.expanded.set(workbook, !(this.expanded.get(workbook) ?? true));
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
      // keyboard can do everything the mouse can. A workbook row keeps the click, which toggles.
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

  setProjects(projects: ExplorerProject[]): void {
    this.projects = projects;
    this.render();
  }

  setActive(name: string | null): void {
    this.active = name;
    this.render();
  }

  /** Problem counts by component, so a defect is visible without opening the module. */
  setProblemCounts(counts: Map<string, number>): void {
    this.problemCounts = counts;
    this.render();
  }

  private render(): void {
    this.root.replaceChildren();

    for (const project of this.projects) {
      const isOpen = this.expanded.get(project.name) ?? true;
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

    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item"
      + (component.name === this.active ? " active" : "")
      + (component.name === this.selected ? " selected" : "");
    button.dataset.component = component.name;
    button.dataset.kind = String(component.kind);
    button.setAttribute("role", "treeitem");
    button.setAttribute("aria-selected", String(component.name === this.selected));

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${meta.icon}`;
    glyph.setAttribute("aria-hidden", "true");

    const kind = document.createElement("span");
    kind.className = "tree-kind";
    kind.textContent = meta.type;

    button.append(glyph, document.createTextNode(component.name), kind);

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
}
