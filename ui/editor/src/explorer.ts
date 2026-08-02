/*
 * The project explorer.
 *
 * This replaces the host's own, which is closed rather than covered: the host will hide a tool
 * window on request, and a hidden window cannot be uncovered by anything the host does later. The
 * project itself is untouched, so everything that reads it keeps working.
 *
 * Components are grouped the way the host groups them, because that grouping is what a VBA
 * developer already knows. The names of the groups are ours; the membership comes from the kind
 * the host reports for each component.
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

interface Group {
  title: string;
  kinds: number[];
  icon: string;
}

/** Groups in the order they are shown, matching the order the host shows them in. */
const GROUPS: Group[] = [
  { title: "Excel Objects", kinds: [ComponentKind.Document], icon: "file-binary" },
  { title: "Forms", kinds: [ComponentKind.Form, ComponentKind.ActiveXDesigner], icon: "window" },
  { title: "Modules", kinds: [ComponentKind.StandardModule], icon: "file-code" },
  { title: "Class Modules", kinds: [ComponentKind.ClassModule], icon: "symbol-class" },
];

export interface ExplorerHandlers {
  /** Single click: the component becomes the selection, and nothing opens. */
  select(name: string): void;
  /** Double click, or Enter on the focused item: the component's code opens. */
  open(name: string): void;
}

export class Explorer {
  private readonly root: HTMLElement;
  private readonly handlers: ExplorerHandlers;

  private projects: ExplorerProject[] = [];
  private active: string | null = null;
  private selected: string | null = null;
  private problemCounts = new Map<string, number>();

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
      // keyboard can do everything the mouse can.
      if (event.key === "Enter") {
        const name = this.componentAt(event);
        if (name) {
          event.preventDefault();
          this.handlers.open(name);
        }
      }
    });
  }

  private componentAt(event: Event): string | null {
    const item = (event.target as HTMLElement).closest("[data-component]") as HTMLElement | null;
    return item?.dataset.component ?? null;
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
      const header = document.createElement("div");
      header.className = "tree-project";

      const icon = document.createElement("span");
      icon.className = "codicon codicon-folder-library";
      icon.setAttribute("aria-hidden", "true");

      header.append(icon, document.createTextNode(project.name));
      this.root.appendChild(header);

      for (const group of GROUPS) {
        const members = project.components
          .filter((component) => group.kinds.includes(component.kind))
          .sort((a, b) => a.name.localeCompare(b.name));

        if (members.length === 0) {
          continue;
        }

        const title = document.createElement("div");
        title.className = "tree-group";
        title.textContent = group.title;
        this.root.appendChild(title);

        for (const component of members) {
          this.root.appendChild(this.item(component, group.icon));
        }
      }
    }
  }

  private item(component: ExplorerComponent, icon: string): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item"
      + (component.name === this.active ? " active" : "")
      + (component.name === this.selected ? " selected" : "");
    button.dataset.component = component.name;
    button.setAttribute("role", "treeitem");
    button.setAttribute("aria-selected", String(component.name === this.selected));

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${icon}`;
    glyph.setAttribute("aria-hidden", "true");

    button.append(glyph, document.createTextNode(component.name));

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
