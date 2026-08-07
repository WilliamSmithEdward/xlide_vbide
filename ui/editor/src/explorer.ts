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
 * Closing the last tab folds the whole thing back, workbooks included: with nothing open there is
 * nothing for the tree to be following, and it should read as at rest rather than half unpacked.
 *
 * All of that following is one setting — "Explorer follows the editor", on by default. Off, the
 * tree does nothing on its own: it unfolds what is clicked and stays as it was left, which is the
 * right answer for anyone who arranges the tree deliberately and does not want it rearranged
 * underneath them (the developer, 2026-08-07).
 */

import { currentSettings } from "./settings.js";

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
  /** Double click, or Enter on the focused item: the component's code opens. The workbook says
   *  WHICH one when two workbooks share the name. */
  open(name: string, workbook?: string): void;
  /** Right click on a component: the menu for its class, at this position. */
  context(name: string, kind: number, x: number, y: number): void;
  /** Right click on a workbook's row. */
  projectContext(project: string, x: number, y: number): void;
  /** A module's procedures, for its unfolded node; null when no answer came. */
  outline(module: string, workbook?: string): Promise<ExplorerProcedure[] | null>;
  /** A procedure was picked: go to its line in its module, in its workbook. */
  openProcedure(module: string, line: number, workbook?: string): void;
  /** A line for the host's log, for the defects only the log's data cadence explains. */
  trace?(text: string): void;
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

  /** The workbook that module belongs to, so a shared name unfolds the right workbook's list. */
  private expandedModuleWorkbook: string | null = null;

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
        this.toggleModule(toggle.dataset.toggle, toggle.dataset.workbook);
        return;
      }

      const procedure = this.procedureAt(event);
      if (procedure) {
        this.handlers.openProcedure(procedure.module, procedure.line, procedure.workbook);
        return;
      }

      // A single click selects AND opens. The properties panel follows the selection either
      // way, but a click that only selected read as a tree that had stopped working: opening
      // on a single click is the ergonomic this product follows, deliberately unlike the
      // host's own tree, which asks for a double click.
      const name = this.componentAt(event);
      if (name) {
        const workbook = this.workbookOf(event);
        this.selected = name;
        this.setExpandedModule(name, workbook);
        this.render();
        this.handlers.select(name);
        this.handlers.open(name, workbook);
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

  private procedureAt(event: Event): { module: string; line: number; workbook?: string } | null {
    const row = (event.target as HTMLElement).closest("[data-proc-module]") as HTMLElement | null;
    if (!row?.dataset.procModule) {
      return null;
    }
    return {
      module: row.dataset.procModule,
      line: Number(row.dataset.procLine ?? "1"),
      ...(row.dataset.procWorkbook ? { workbook: row.dataset.procWorkbook } : {}),
    };
  }

  /** The workbook of the component row an event landed on, when the row carries one. */
  private workbookOf(event: Event): string | undefined {
    const item = (event.target as HTMLElement).closest("[data-component]") as HTMLElement | null;
    return item?.dataset.workbook || undefined;
  }

  /** The workspace as last published, for whoever else browses it (the Object Browser). */
  snapshot(): ExplorerProject[] {
    return this.projects;
  }

  setProjects(projects: ExplorerProject[]): void {
    // The host republishes the tree on all sorts of occasions, usually unchanged. An identical
    // push must change nothing on screen: acting on it cleared and refetched the unfolded
    // procedures, which blinked them out and back in a loop.
    const key = JSON.stringify(projects);
    if (key === this.projectsKey) {
      return;
    }

    // A push that got past the guard names itself in the host's log, because a push that
    // CHANGES on a timer is a defect upstream and the diff is the only witness.
    if (this.projectsKey !== "") {
      this.handlers.trace?.(`tree: projects push changed, ${diffExcerpt(this.projectsKey, key)}`);
    }

    this.projectsKey = key;
    this.projects = projects;

    // Outlines for modules that no longer exist go. The rest stay as shown until their refresh
    // lands, so an unfolded list is replaced, never blanked.
    const names = new Set(projects.flatMap((project) => project.components.map((component) => component.name)));
    for (const known of [...this.outlines.keys()]) {
      if (!names.has(known)) {
        this.outlines.delete(known);
      }
    }

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

  /** The last projects payload applied, so an identical republish is a no-op. */
  private projectsKey = "";

  setActive(name: string | null, workbook?: string): void {
    // The tree follows the module being edited — but only when it genuinely changed. The host
    // republishes the module list on all sorts of occasions with the same active module;
    // following every push would fold what was just unfolded by hand, and even redrawing on
    // every push wipes and rebuilds a large unfolded list, which reads as flicker.
    if (name === this.active) {
      return;
    }

    this.active = name;
    if (name) {
      this.setExpandedModule(name, workbook);
    }
    this.render();
  }

  /** Problem counts by component, so a defect is visible without opening the module. */
  setProblemCounts(counts: Map<string, number>): void {
    // Identical counts arrive constantly and must change nothing: neither a redraw of a large
    // unfolded list, nor a re-parse of its module.
    const same = counts.size === this.problemCounts.size
      && [...counts].every(([name, count]) => this.problemCounts.get(name) === count);
    if (same) {
      return;
    }

    this.handlers.trace?.(`tree: counts push changed, ${countsDelta(this.problemCounts, counts)}`);
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
  private setExpandedModule(name: string, workbook?: string): void {
    // Following is what the setting governs. A module CLICKED in the tree still unfolds — that
    // goes through toggleModule — so switching this off makes the tree passive, not inert.
    if (!currentSettings().treeFollowsEditor && !this.unfoldingByHand) {
      return;
    }

    const owner = workbook
      ? this.projects.find((project) => project.name.toLowerCase() === workbook.toLowerCase())
      : this.projects.find((project) =>
        project.components.some((component) => component.name === name));

    const ownerName = owner?.name ?? null;
    if (this.expandedModule !== name || this.expandedModuleWorkbook !== ownerName) {
      this.expandedModule = name;
      this.expandedModuleWorkbook = ownerName;
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

  /**
   * Folds the tree all the way back: every module's procedures, and every workbook.
   *
   * Called when the last tab closes. The unfolded module is the accordion's memory of what was
   * being worked on, and once nothing is open there is no such thing, so procedures hanging open
   * under an empty editor point at work that is no longer there.
   *
   * The workbooks go too. That is a reversal of the first version, which kept them on the grounds
   * that an expanded workbook is a hand-made choice: with nothing open at all, the tree should be
   * back where it starts rather than half unpacked (the developer, 2026-08-07). Opening anything
   * expands its workbook again on the way in.
   *
   * Nothing here happens when the tree has been told not to follow the editor.
   */
  collapseAll(): void {
    if (!currentSettings().treeFollowsEditor) {
      return;
    }

    const wasUnfolded = this.expandedModule !== null
      || [...this.expandedWorkbooks.values()].some((open) => open);

    if (!wasUnfolded) {
      return;
    }

    this.expandedModule = null;
    this.expandedModuleWorkbook = null;
    this.expandedWorkbooks.clear();

    // Forgotten too, or the next activation would count as "the attention has not moved" and
    // leave the workbook it lands in closed.
    this.attentionWorkbook = null;

    this.render();
  }

  /** True while a click is being served, so setExpandedModule knows this was asked for. */
  private unfoldingByHand = false;

  private toggleModule(name: string, workbook?: string): void {
    if (this.expandedModule === name
      && (!workbook || this.expandedModuleWorkbook === workbook)) {
      this.expandedModule = null;
      this.expandedModuleWorkbook = null;
    } else {
      this.unfoldingByHand = true;
      try {
        this.setExpandedModule(name, workbook);
      } finally {
        this.unfoldingByHand = false;
      }
    }
    this.render();
  }

  /** Modules with an outline request in flight, so pushes queue one trailing refresh at most. */
  private readonly outlineFetching = new Set<string>();
  private readonly outlineRefetch = new Set<string>();

  private async fetchOutline(module: string): Promise<void> {
    // One request per module at a time. The pushes that trigger refreshes arrive faster than a
    // busy host answers, and concurrent requests came home out of order: a late empty answer
    // landing after a full one is exactly how an unfolded list vanished mid-show.
    if (this.outlineFetching.has(module)) {
      this.outlineRefetch.add(module);
      return;
    }

    this.outlineFetching.add(module);
    try {
      const procedures = await this.handlers.outline(module, this.expandedModuleWorkbook ?? undefined);
      this.applyOutline(module, procedures);
    } finally {
      this.outlineFetching.delete(module);
      if (this.outlineRefetch.delete(module) && this.expandedModule === module) {
        void this.fetchOutline(module);
      }
    }
  }

  private applyOutline(module: string, procedures: ExplorerProcedure[] | null): void {
    // No answer is not an answer: a timeout or a host failure says nothing about the module,
    // and what is already unfolded stays. Only a real answer may replace it — including a real
    // empty one, which is a module whose procedures were genuinely deleted.
    if (procedures === null || this.expandedModule !== module) {
      return;
    }

    // An unchanged list redraws nothing: refreshes arrive with every analysis pass, and most
    // of them confirm what is already on screen.
    const known = this.outlines.get(module);
    if (known && known.length === procedures.length
      && known.every((procedure, i) => procedure.name === procedures[i]?.name
        && procedure.kind === procedures[i]?.kind
        && procedure.line === procedures[i]?.line)) {
      return;
    }

    this.outlines.set(module, procedures);
    this.render();
  }

  private render(): void {
    // Rebuilding the rows resets the scroll; a redraw must not read as a jump.
    const scrollTop = this.root.scrollTop;
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
        this.root.appendChild(this.item(component, project.name));

        // The unfolded module is one (name, workbook) pair, so a shared name unfolds only in
        // the workbook whose row was opened.
        if (component.name === this.expandedModule
          && (this.expandedModuleWorkbook === null || this.expandedModuleWorkbook === project.name)) {
          for (const procedure of this.outlines.get(component.name) ?? []) {
            this.root.appendChild(this.procedureRow(component.name, procedure, project.name));
          }
        }
      }
    }

    this.root.scrollTop = scrollTop;
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

  private item(component: ExplorerComponent, workbook: string): HTMLElement {
    const meta = kindMeta(component.kind);
    const isUnfolded = component.name === this.expandedModule
      && (this.expandedModuleWorkbook === null || this.expandedModuleWorkbook === workbook);

    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item"
      + (component.name === this.active ? " active" : "")
      + (component.name === this.selected ? " selected" : "");
    button.dataset.component = component.name;
    button.dataset.workbook = workbook;
    button.dataset.kind = String(component.kind);
    button.setAttribute("role", "treeitem");
    button.setAttribute("aria-selected", String(component.name === this.selected));
    button.setAttribute("aria-expanded", String(isUnfolded));

    const chevron = document.createElement("span");
    chevron.className = `codicon codicon-chevron-${isUnfolded ? "down" : "right"} tree-twisty`;
    chevron.dataset.toggle = component.name;
    chevron.dataset.workbook = workbook;
    chevron.setAttribute("aria-hidden", "true");

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${meta.icon}`;
    glyph.setAttribute("aria-hidden", "true");

    const kind = document.createElement("span");
    kind.className = "tree-kind";
    kind.textContent = meta.type;

    button.append(chevron, glyph, document.createTextNode(component.name), kind);

    const problems = this.problemCounts.get(problemCountKey(workbook, component.name)) ?? 0;
    if (problems > 0) {
      const badge = document.createElement("span");
      badge.className = "tree-badge";
      badge.textContent = String(problems);
      badge.title = `${problems} problem${problems === 1 ? "" : "s"}`;
      button.appendChild(badge);
    }

    return button;
  }

  private procedureRow(module: string, procedure: ExplorerProcedure, workbook: string): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item tree-proc";
    button.dataset.procModule = module;
    button.dataset.procLine = String(procedure.line);
    button.dataset.procWorkbook = workbook;
    button.setAttribute("role", "treeitem");

    const glyph = document.createElement("span");
    glyph.className = "codicon codicon-symbol-method";
    glyph.setAttribute("aria-hidden", "true");

    button.append(glyph, document.createTextNode(`${procedure.kind} ${procedure.name}`));
    return button;
  }
}

/**
 * The key a problem count is filed under: the workbook and the module, lowercased, because a
 * count belongs to one workbook's module and a shared name must not pool them. The shell files
 * counts by this and the tree looks them up by it.
 */
export function problemCountKey(workbook: string | null | undefined, name: string): string {
  return `${(workbook ?? "").toLowerCase()}\0${name.toLowerCase()}`;
}

/** Where two serialized payloads part ways, with enough of each side to name the field. */
function diffExcerpt(before: string, after: string): string {
  let at = 0;
  const limit = Math.min(before.length, after.length);
  while (at < limit && before[at] === after[at]) {
    at += 1;
  }

  const window = (text: string): string =>
    text.slice(Math.max(0, at - 24), at + 48).replaceAll('"', "'");
  return `${before.length} -> ${after.length} chars, first diff at ${at}:`
    + ` was "${window(before)}" now "${window(after)}"`;
}

/** The entries that differ between two count maps, capped so the log line stays a line. */
function countsDelta(before: Map<string, number>, after: Map<string, number>): string {
  const names = new Set([...before.keys(), ...after.keys()]);
  const changes: string[] = [];
  for (const name of names) {
    const was = before.get(name);
    const now = after.get(name);
    if (was !== now) {
      changes.push(`${name} ${was ?? "-"} -> ${now ?? "-"}`);
    }
  }

  const shown = changes.slice(0, 4).join(", ");
  return changes.length > 4 ? `${shown}, +${changes.length - 4} more` : shown;
}
