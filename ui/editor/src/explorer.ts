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
 * All of that following is one setting - "Explorer follows the editor", on by default. Off, the
 * tree does nothing on its own: it unfolds what is clicked and stays as it was left, which is the
 * right answer for anyone who arranges the tree deliberately and does not want it rearranged
 * underneath them (the developer, 2026-08-07).
 *
 * TWO LAYOUTS, tabbed at the top of the pane (#23). "Tree" is the flat list above. "Folders"
 * groups a workbook's modules by the '@Folder("Parent.Child") comment at the top of each one -
 * the Rubberduck convention, so a project organised there is organised here without editing a
 * line. A folder row unfolds and folds like a workbook; a module row is the same row it is in the
 * flat list, one level deeper per folder, and the accordion, the selection and the following all
 * work the same way. Modules with no annotation sit at the workbook's root. Which layout shows is
 * a setting, so it survives the session; the tabs are the handle on it.
 *
 * And the CURRENT PROCEDURE: the row of the procedure the caret is in, under the unfolded module,
 * wears a mark, so where the developer is reads off the tree as well as off the status bar.
 */

import { currentSettings, onSettingsApplied, type EditorSettings } from "./settings.js";
import { allFolders, ancestorPaths, buildFolderTree, folderKey, type FolderNode } from "./foldertree.js";

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
  /** The folder its '@Folder annotation names, dotted, or nothing: the host read it. */
  folder?: string | null;
}

/** The two layouts the pane offers. */
export type ExplorerView = EditorSettings["explorerView"];

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

/** The tree as the xlide api reports it. */
export interface ExplorerSnapshot {
  selected: string | null;
  active: string | null;
  attentionWorkbook: string | null;
  unfolded: { module: string; workbook: string | null } | null;
  /** Whether the "Explorer follows the editor" setting is on, since it gates every automatic move. */
  follows: boolean;
  /** Which layout is showing. */
  view: ExplorerView;
  /** The procedure row wearing the caret's mark, or null when no row does. */
  currentProcedure: { module: string; workbook: string | null; name: string; line: number } | null;
  workbooks: {
    name: string;
    expanded: boolean;
    /** Every folder the workbook's annotations make, parents first, whether the folder view is showing or not. */
    folders: { path: string; expanded: boolean; modules: number }[];
    modules: {
      name: string;
      kind: string;
      /** The folder the module's annotation names, or null at the root. */
      folder: string | null;
      problems: number;
      unfolded: boolean;
      procedures: ExplorerProcedure[];
    }[];
  }[];
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
  /** Right click on a component: the menu for its class, at this position. The WORKBOOK is part
   *  of it, because every action the menu offers acts on a module, and a module is a name in a
   *  workbook. Without it the menu's Open, Rename and Close all resolved by bare name and would
   *  act on whichever workbook answered first (2026-08-08). */
  context(name: string, kind: number, x: number, y: number, workbook?: string): void;
  /** Right click on a workbook's row. */
  projectContext(project: string, x: number, y: number): void;
  /** The developer pressed a workbook's plus: what can be added to it, under the button. */
  projectAdd(project: string, x: number, y: number): void;
  /** A module's procedures, for its unfolded node; null when no answer came. */
  outline(module: string, workbook?: string): Promise<ExplorerProcedure[] | null>;
  /** A procedure was picked: go to its line in its module, in its workbook. */
  openProcedure(module: string, line: number, workbook?: string): void;

  /** A form's designer was picked in the tree: open its design face, the way the tab's own menu
   * does. */
  openDesigner(module: string, workbook?: string): void;
  /**
   * A row is being dragged toward the editor: a module, or a procedure carrying its line.
   * `became` fires if the press turns into a real drag, so the click it would otherwise be
   * can be swallowed - a drag that ALSO clicked would unfold the accordion mid-gesture. The
   * flag it sets is reset on the next pointerdown, not on a timer, so a drag ending outside
   * the tree cannot leave a click eaten later.
   */
  dragRow(payload: { module: string; workbook?: string; line?: number; member?: string }, start: PointerEvent, became: () => void): void;
  /**
   * The developer pressed the other layout's tab. The layout is a SETTING, so this asks the host
   * to change it the way the dialog does; the tree redraws when the echo lands, and not before.
   */
  changeView(view: ExplorerView): void;
  /** A line for the host's log, for the defects only the log's data cadence explains. */
  trace?(text: string): void;
}

export class Explorer {
  private readonly root: HTMLElement;
  private readonly handlers: ExplorerHandlers;

  private projects: ExplorerProject[] = [];
  private active: string | null = null;
  private selected: string | null = null;

  /*
   * WHICH WORKBOOK the active and selected rows belong to.
   *
   * A row used to be matched by NAME alone, and every workbook has a ThisWorkbook and a Sheet1.
   * With two workbooks open, clicking one lit BOTH rows: the tree said the module was open in
   * two places at once, and the properties panel and the twisty followed whichever the render
   * reached first (reported 2026-08-08, with two unsaved workbooks side by side).
   *
   * A module is a name IN a workbook, so the pair is what identifies it.
   */
  private activeWorkbook: string | null = null;
  private selectedWorkbook: string | null = null;
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

  /** True from a row drag until the click it would otherwise become has been swallowed. */
  private dragConsumedClick = false;

  /**
   * Folders the developer folded, by workbook and path. Folders start OPEN: the view exists to
   * show the structure, and a fresh tree hiding every module behind a click would show none of
   * it. A folded folder stays folded until the developer opens it, or the module being edited
   * turns out to be inside it and the tree is following.
   */
  private readonly collapsedFolders = new Set<string>();

  /** The procedure row wearing the caret's mark. See setCaret. */
  private currentProcedure: { module: string; workbook: string | null; name: string; line: number } | null = null;

  /** The tab strip above the tree, when the shell gave it one. */
  private readonly views: HTMLElement | null;

  /** The layout last drawn, so a settings echo that changed something else redraws nothing. */
  private drawnView: ExplorerView = currentSettings().explorerView;

  constructor(root: HTMLElement, handlers: ExplorerHandlers, views: HTMLElement | null = null) {
    this.root = root;
    this.handlers = handlers;
    this.views = views;

    // The two layout tabs. They ASK for the layout rather than switching to it: the layout is a
    // setting the host owns, and the tree redraws on the echo, the same road the dialog takes.
    if (views) {
      views.replaceChildren();
      for (const [view, title] of [["tree", "Tree"], ["folders", "Folders"]] as const) {
        const tab = document.createElement("button");
        tab.type = "button";
        tab.className = "explorer-view";
        tab.dataset.view = view;
        tab.textContent = title;
        tab.setAttribute("role", "tab");
        tab.addEventListener("click", () => {
          if (currentSettings().explorerView !== view) {
            this.handlers.changeView(view);
          }
        });
        views.appendChild(tab);
      }
      this.paintViewTabs();
    }

    onSettingsApplied((settings) => {
      if (settings.explorerView !== this.drawnView) {
        this.paintViewTabs();
        this.render();
      }
    });

    // A row can be DRAGGED to the editor: a module row opens where it lands, a procedure row
    // opens and goes to its line. The press only becomes a drag past the movement threshold,
    // so ordinary clicks are untouched; when it does, the click it would still produce is
    // swallowed below, or the gesture would also unfold the accordion on its way out.
    this.root.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      // EVERY fresh press starts with a clean slate. This is the whole click-suppression: a
      // drag sets the flag, the click that drag produces consumes it - and if the drag ended
      // OUTSIDE the tree there is no such click, so the flag would linger and eat the next
      // honest press. Clearing it here, deterministically, at the start of the very gesture
      // that would be eaten, is what a timer could only race ("clicking the workbook row
      // sometimes takes two clicks", 2026-08-12).
      this.dragConsumedClick = false;

      const target = event.target as HTMLElement;
      if (target.closest("[data-toggle]") || target.closest("[data-add-project]")) {
        return;
      }

      const procedure = this.procedureAt(event);
      // The ghost chip wears the member's name as the row spells it; the row is in hand.
      const spelled = (target.closest("[data-proc-module]")?.textContent ?? "").trim();
      const component = procedure ? null : this.componentAt(event);
      const workbook = this.workbookOf(event);
      const payload = procedure
        ? { ...procedure, ...(spelled ? { member: spelled } : {}) }
        : component
          ? { module: component, ...(workbook ? { workbook } : {}) }
          : null;

      if (payload) {
        // The click this press will produce, if the press becomes a drag, is swallowed below.
        // The flag is RESET by the next pointerdown (above), not by a timer, so a drag that
        // ends outside the tree cannot leave it armed for the next click.
        this.handlers.dragRow(payload, event, () => {
          this.dragConsumedClick = true;
        });
      }
    });

    // One listener for the whole tree. The tree is rebuilt whenever the project changes, and
    // per-item listeners would have to be torn down with it.
    this.root.addEventListener("click", (event) => {
      // A click that was really the tail of a drag has already had its effect at the drop.
      if (this.dragConsumedClick) {
        this.dragConsumedClick = false;
        return;
      }

      // The chevron toggles and does nothing else, so unfolding is never also an open.
      const toggle = (event.target as HTMLElement).closest("[data-toggle]") as HTMLElement | null;
      if (toggle?.dataset.toggle) {
        this.toggleModule(toggle.dataset.toggle, toggle.dataset.workbook);
        return;
      }

      // Before the row itself: the plus sits INSIDE the workbook row, so without this the same
      // click would also toggle the workbook open or shut underneath the menu it just opened.
      const add = (event.target as HTMLElement).closest("[data-add-project]") as HTMLElement | null;
      if (add?.dataset.addProject) {
        const box = add.getBoundingClientRect();
        // Under the button rather than at the pointer, so the menu hangs off the control that
        // opened it however it was opened - including from the keyboard, where there is no pointer
        // and a click reports 0,0.
        this.handlers.projectAdd(add.dataset.addProject, Math.round(box.left), Math.round(box.bottom));
        return;
      }

      // Before the procedure test, because a designer row is not one and both are children of
      // the same component.
      const designer = (event.target as HTMLElement).closest("[data-designer-module]") as HTMLElement | null;
      if (designer?.dataset.designerModule) {
        this.handlers.openDesigner(designer.dataset.designerModule, designer.dataset.designerWorkbook);
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
        this.selectedWorkbook = workbook ?? null;
        this.setExpandedModule(name, workbook);
        this.render();
        this.handlers.select(name);
        this.handlers.open(name, workbook);
        return;
      }

      // A folder row folds and unfolds, and nothing else: there is nothing to open.
      const folder = this.folderAt(event);
      if (folder) {
        this.toggleFolder(folder.workbook, folder.path);
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
      // keyboard can do everything the mouse can. Procedure rows keep the click.
      if (event.key === "Enter") {
        const name = this.componentAt(event);
        if (name) {
          event.preventDefault();
          this.handlers.open(name);
          return;
        }
      }

      // The workbook row stopped being a button when it grew one, and with the element went the
      // keyboard behaviour the element carried. Enter and Space toggle it, as they did.
      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }

      const row = event.target as HTMLElement;
      if (row.classList?.contains("tree-workbook") && row.dataset.project) {
        event.preventDefault();
        const workbook = row.dataset.project;
        this.expandedWorkbooks.set(workbook, !(this.expandedWorkbooks.get(workbook) ?? false));
        this.render();

        // The rebuild threw away the element the keyboard was on, and focus went to the body with
        // it: the next Tab would start from the top of the page rather than from the tree.
        (this.root.querySelector(`[data-project="${CSS.escape(workbook)}"]`) as HTMLElement | null)?.focus();
        return;
      }

      // A folder row is a treeitem the same way, and keeps the keyboard the same way.
      if (row.classList?.contains("tree-folder") && row.dataset.folder !== undefined && row.dataset.folderWorkbook) {
        event.preventDefault();
        const { folder, folderWorkbook } = row.dataset;
        this.toggleFolder(folderWorkbook, folder);
        (this.root.querySelector(
          `[data-folder-workbook="${CSS.escape(folderWorkbook)}"][data-folder="${CSS.escape(folder)}"]`) as HTMLElement | null)?.focus();
      }
    });

    this.root.addEventListener("contextmenu", (event) => {
      const item = (event.target as HTMLElement).closest("[data-component]") as HTMLElement | null;
      if (item?.dataset.component) {
        event.preventDefault();
        this.selected = item.dataset.component;
        this.selectedWorkbook = item.dataset.workbook ?? null;
        this.render();
        this.handlers.select(item.dataset.component);
        this.handlers.context(
          item.dataset.component,
          Number(item.dataset.kind ?? "0"),
          event.clientX,
          event.clientY,
          item.dataset.workbook);
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

  private folderAt(event: Event): { workbook: string; path: string } | null {
    const row = (event.target as HTMLElement).closest("[data-folder]") as HTMLElement | null;
    if (!row || row.dataset.folder === undefined || !row.dataset.folderWorkbook) {
      return null;
    }
    return { workbook: row.dataset.folderWorkbook, path: row.dataset.folder };
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
    // The tree follows the module being edited - but only when it genuinely changed. The host
    // republishes the module list on all sorts of occasions with the same active module;
    // following every push would fold what was just unfolded by hand, and even redrawing on
    // every push wipes and rebuilds a large unfolded list, which reads as flicker.
    // BOTH halves, or switching between two workbooks' ThisWorkbook is "no change" and the
    // tree never follows the tab.
    if (name === this.active && (workbook ?? null) === this.activeWorkbook) {
      return;
    }

    this.active = name;
    this.activeWorkbook = workbook ?? null;

    /*
     * THE SELECTION COMES WITH IT, or a closed module keeps its highlight forever.
     *
     * `selected` was set by a click and never by anything else, so it outlived whatever it
     * pointed at. Close the tab of the module you last clicked and its row keeps the grey while
     * a different row goes blue: a highlight on something that is not open, not active, and not
     * being looked at (reported 2026-08-08).
     *
     * Moving it here is not a workaround for that, it is what the two states already mean. A
     * click in this tree SELECTS AND OPENS in one gesture, so they are the same thing every time
     * a developer makes them happen; they diverge only on a right-click, which selects without
     * opening, and the next activation is entitled to take the selection back.
     *
     * The properties panel follows the selection, so this also means it describes the module on
     * screen rather than the last one clicked, which is what every editor does.
     */
    if (name) {
      this.selected = name;
      this.selectedWorkbook = workbook ?? null;
      this.setExpandedModule(name, workbook);
    }

    this.render();
  }

  /**
   * Puts the selection back on the module that is actually open.
   *
   * A right-click marks the row the menu is about, which it must: a menu with no visible target
   * is a menu about nothing. But the mark is part of the GESTURE, and the gesture ends when the
   * menu does - left alone it becomes a grey row pointing at a module nobody is looking at, which
   * is the same defect as a selection outliving its tab (2026-08-08).
   *
   * Restored to the active module rather than cleared to nothing, because the selection is what
   * the properties panel describes: clearing it would leave the panel with no subject, and the
   * module on screen is the honest answer to "what am I looking at".
   */
  restoreSelectionToActive(): void {
    if (this.selected === this.active && this.selectedWorkbook === this.activeWorkbook) {
      return;
    }

    this.selected = this.active;
    this.selectedWorkbook = this.activeWorkbook;
    this.render();

    if (this.active) {
      this.handlers.select(this.active);
    }
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
    // Following is what the setting governs. A module CLICKED in the tree still unfolds - that
    // goes through toggleModule - so switching this off makes the tree passive, not inert.
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
      // The folders above the module open on the way to it, or the accordion would unfold a row
      // the developer cannot see. Only those: a folder folded by hand elsewhere stays folded.
      const folder = owner.components.find((component) => component.name === name)?.folder;
      for (const path of ancestorPaths(folder)) {
        this.collapsedFolders.delete(this.folderStateKey(owner.name, path));
      }
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
   * The workbooks go too, UNLESS THERE IS ONLY ONE. Folding them back was a reversal of the first
   * version, which kept them on the grounds that an expanded workbook is a hand-made choice: with
   * nothing open at all, the tree should be back where it starts rather than half unpacked (the
   * developer, 2026-08-07). Opening anything expands its workbook again on the way in.
   *
   * With a single workbook that leaves a tree of one closed row, which is every module in the
   * project hidden behind a click that has one possible answer. Folding several workbooks is real
   * tidying and folding the only one is just an empty tree, so the count decides (the developer,
   * 2026-08-10). Its procedures still fold: those are the accordion's memory of what was being
   * worked on, and with nothing open there is no such thing.
   *
   * Nothing here happens when the tree has been told not to follow the editor.
   */
  collapseAll(): void {
    if (!currentSettings().treeFollowsEditor) {
      return;
    }

    const soleWorkbook = this.projects.length <= 1;

    const wasUnfolded = this.expandedModule !== null
      || (!soleWorkbook && [...this.expandedWorkbooks.values()].some((open) => open));

    if (!wasUnfolded) {
      return;
    }

    this.expandedModule = null;
    this.expandedModuleWorkbook = null;

    if (!soleWorkbook) {
      this.expandedWorkbooks.clear();

      // Forgotten with them, or the next activation would count as "the attention has not moved"
      // and leave the workbook it lands in closed. Kept when the sole workbook stays open, since
      // there is no closed workbook for it to strand.
      this.attentionWorkbook = null;
    }

    this.render();
  }

  /**
   * The tree as data, for the xlide api's `ui` route.
   *
   * Reported from the fields the render reads rather than from the rendered rows, so a probe
   * asks what the tree BELIEVES. Scraping `.tree-item` was how this was measured before, and it
   * cannot see the difference between "collapsed" and "rendered wrong".
   */
  treeState(): ExplorerSnapshot {
    return {
      selected: this.selected,
      active: this.active,
      attentionWorkbook: this.attentionWorkbook,
      unfolded: this.expandedModule === null ? null : {
        module: this.expandedModule,
        workbook: this.unfoldedWorkbook(),
      },
      follows: currentSettings().treeFollowsEditor,
      view: currentSettings().explorerView,
      currentProcedure: this.currentProcedure === null ? null : { ...this.currentProcedure },
      workbooks: this.projects.map((project) => ({
        name: project.name,
        expanded: this.expandedWorkbooks.get(project.name) ?? false,
        folders: allFolders(this.folderTreeOf(project)).map((folder) => ({
          path: folder.path,
          expanded: this.isFolderExpanded(project.name, folder.path),
          modules: countModules(folder),
        })),
        modules: project.components.map((component) => ({
          name: component.name,
          kind: kindMeta(component.kind).type,
          folder: component.folder ?? null,
          problems: this.problemCounts.get(component.name) ?? 0,
          unfolded: this.isUnfolded(component.name, project.name),
          procedures: (this.outlines.get(component.name) ?? []).map((one) => ({ ...one })),
        })),
      })),
    };
  }

  /** Opens or shuts a folder the way its row does. False when no such folder is drawn. */
  setFolderExpanded(workbook: string, path: string, open: boolean): boolean {
    const project = this.projects.find((one) => one.name === workbook);
    if (!project || !allFolders(this.folderTreeOf(project)).some((folder) => folderKey(folder.path) === folderKey(path))) {
      return false;
    }

    const key = this.folderStateKey(workbook, path);
    if (open) {
      this.collapsedFolders.delete(key);
    } else {
      this.collapsedFolders.add(key);
    }
    this.render();
    return true;
  }

  /**
   * Where the caret is, so the tree can mark the procedure it is in.
   *
   * The NAME is decided by the page's own scan of the text, the same one the status bar shows;
   * the line only picks between rows of one name, which is what a property's Get and Let are.
   * Null clears the mark - the caret left every procedure, or the editor is empty.
   */
  setCaret(module: string | null, workbook: string | null, line: number, procedure: string | null): void {
    const next = module !== null && procedure !== null ? { module, workbook, name: procedure, line } : null;
    const same = (this.currentProcedure === null && next === null)
      || (this.currentProcedure !== null && next !== null
        && this.currentProcedure.module === next.module
        && this.currentProcedure.workbook === next.workbook
        && this.currentProcedure.name === next.name
        && this.markedRow(this.currentProcedure) === this.markedRow(next));
    this.currentProcedure = next;
    if (!same) {
      this.render();
    }
  }

  /** The line of the outline row the mark sits on, resolved the same way the render resolves it. */
  private markedRow(current: { module: string; workbook: string | null; name: string; line: number }): number | null {
    const rows = (this.outlines.get(current.module) ?? []).filter((one) => one.name === current.name);
    if (rows.length === 0) {
      return null;
    }
    // The last row at or above the caret, or the first of the name when the caret is above them
    // all (the leading comments of the first leg).
    const below = rows.filter((one) => one.line <= current.line);
    return (below[below.length - 1] ?? rows[0])!.line;
  }

  private folderStateKey(workbook: string, path: string): string {
    return `${workbook.toLowerCase()}\0${folderKey(path)}`;
  }

  private isFolderExpanded(workbook: string, path: string): boolean {
    return !this.collapsedFolders.has(this.folderStateKey(workbook, path));
  }

  private toggleFolder(workbook: string, path: string): void {
    this.setFolderExpanded(workbook, path, !this.isFolderExpanded(workbook, path));
  }

  /** The workbook's modules arranged by their annotations, in the flat tree's own order within a folder. */
  private folderTreeOf(project: ExplorerProject): FolderNode<ExplorerComponent> {
    return buildFolderTree(project.components, componentOrder);
  }

  private paintViewTabs(): void {
    const view = currentSettings().explorerView;
    this.drawnView = view;
    for (const tab of this.views?.querySelectorAll<HTMLElement>("[data-view]") ?? []) {
      const selected = tab.dataset.view === view;
      tab.classList.toggle("selected", selected);
      tab.setAttribute("aria-selected", String(selected));
    }
  }

  /** Opens or shuts a workbook the way its row does, for a script that would otherwise click. */
  setWorkbookExpanded(workbook: string, open: boolean): boolean {
    if (!this.projects.some((project) => project.name === workbook)) {
      return false;
    }

    this.expandedWorkbooks.set(workbook, open);
    this.render();
    return true;
  }

  /** Unfolds a module's procedures the way its chevron does. */
  unfold(module: string, workbook?: string): void {
    this.toggleModule(module, workbook);
  }

  /**
   * Which workbook the unfolded module belongs to, resolved rather than trusted.
   *
   * `expandedModuleWorkbook` is null whenever the owner could not be worked out when the module
   * was unfolded - the projects had not arrived yet, or the caller had no workbook to give. The
   * render and the snapshot both used to read that null as "every workbook", so with two
   * workbooks holding a `Helpers` the accordion unfolded BOTH of them, which is not an accordion.
   * Found by a randomised walk on 2026-08-07; the fourth defect in this codebase of the form "a
   * name is not an identity across workbooks".
   *
   * A null now means "the first workbook that has one", so exactly one row can ever be unfolded.
   */
  private unfoldedWorkbook(): string | null {
    if (this.expandedModule === null) {
      return null;
    }

    if (this.expandedModuleWorkbook !== null) {
      return this.expandedModuleWorkbook;
    }

    const name = this.expandedModule;
    return this.projects.find((project) =>
      project.components.some((component) => component.name.toLowerCase() === name.toLowerCase()))?.name ?? null;
  }

  /** Whether this workbook's copy of the module is the one unfolded. */
  private isUnfolded(module: string, workbook: string): boolean {
    return this.expandedModule === module && this.unfoldedWorkbook() === workbook;
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
    // and what is already unfolded stays. Only a real answer may replace it - including a real
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

    const view = currentSettings().explorerView;
    this.drawnView = view;

    for (const project of this.projects) {
      const isOpen = this.expandedWorkbooks.get(project.name) ?? false;
      this.root.appendChild(this.workbookRow(project.name, isOpen));

      if (!isOpen) {
        continue;
      }

      if (view === "folders") {
        this.renderFolder(this.folderTreeOf(project), project.name, 0);
        continue;
      }

      for (const component of [...project.components].sort(componentOrder)) {
        this.renderModule(component, project.name, 0);
      }
    }

    this.root.scrollTop = scrollTop;
  }

  /** A folder's contents: its folders first, then its modules, each one level deeper. */
  private renderFolder(node: FolderNode<ExplorerComponent>, workbook: string, depth: number): void {
    for (const folder of node.folders) {
      const isOpen = this.isFolderExpanded(workbook, folder.path);
      this.root.appendChild(this.folderRow(folder, workbook, depth, isOpen));
      if (isOpen) {
        this.renderFolder(folder, workbook, depth + 1);
      }
    }

    for (const component of node.modules) {
      this.renderModule(component, workbook, depth);
    }
  }

  /** A module's row, and the rows under it when it is the unfolded one. */
  private renderModule(component: ExplorerComponent, workbook: string, depth: number): void {
    this.root.appendChild(this.item(component, workbook, depth));

    // The unfolded module is one (name, workbook) pair, so a shared name unfolds only in
    // the workbook whose row was opened.
    if (component.name !== this.expandedModule
      || (this.expandedModuleWorkbook !== null && this.expandedModuleWorkbook !== workbook)) {
      return;
    }

    // A FORM'S DESIGNER, ALWAYS FIRST. It is the thing a developer opens a form for, and
    // until now the only ways in were the tab's own menu and a keyboard command - neither
    // of which announces itself (the owner, 2026-08-18). Its siblings are the handlers,
    // so it sits above them: the design comes before the code that answers it, and a fixed
    // position means the row never moves as procedures are added and renamed.
    if (component.kind === ComponentKind.Form) {
      this.root.appendChild(this.designerRow(component.name, workbook, depth));
    }

    // The caret's mark, on one row: the one markedRow resolves for the module the caret is in.
    const current = this.currentProcedure;
    const marked = current !== null && current.module === component.name
      && (current.workbook === null || current.workbook === workbook)
      ? this.markedRow(current)
      : null;

    for (const procedure of this.outlines.get(component.name) ?? []) {
      const isCurrent = marked !== null && procedure.line === marked && procedure.name === current!.name;
      this.root.appendChild(this.procedureRow(component.name, procedure, workbook, depth, isCurrent));
    }
  }

  /**
   * A folder's row: a twisty, a folder glyph, the name, and - while it is folded - how many
   * modules are inside, so a shut folder says what it hides. A treeitem that takes the keyboard
   * the way the workbook row does. The tooltip is the annotation that makes the folder, which is
   * what a developer types into a module to put it here.
   */
  private folderRow(folder: FolderNode<ExplorerComponent>, workbook: string, depth: number, isOpen: boolean): HTMLElement {
    const row = document.createElement("div");
    row.className = "tree-folder";
    row.dataset.folder = folder.path;
    row.dataset.folderWorkbook = workbook;
    row.tabIndex = 0;
    row.title = `'@Folder("${folder.path}")`;
    row.setAttribute("role", "treeitem");
    row.setAttribute("aria-expanded", String(isOpen));
    indent(row, depth);

    const chevron = document.createElement("span");
    chevron.className = `codicon codicon-chevron-${isOpen ? "down" : "right"} tree-twisty`;
    chevron.setAttribute("aria-hidden", "true");

    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${isOpen ? "folder-opened" : "folder"}`;
    glyph.setAttribute("aria-hidden", "true");

    const label = document.createElement("span");
    label.className = "tree-folder-name";
    label.textContent = folder.name;

    row.append(chevron, glyph, label);

    if (!isOpen) {
      const count = countModules(folder);
      const badge = document.createElement("span");
      badge.className = "tree-kind";
      badge.textContent = `${count} module${count === 1 ? "" : "s"}`;
      row.appendChild(badge);
    }

    return row;
  }

  /**
   * A workbook's row: the twisty, the name, and the button that adds to it.
   *
   * NOT A BUTTON ANY MORE, and it cannot be one. The row carries its own control now, and a button
   * inside a button is not valid HTML: the browser un-nests it and which one a click lands on stops
   * being predictable. So the row is a treeitem that takes the keyboard for itself - Enter and
   * Space toggle it in the tree's own keydown, which is what the button element used to give free.
   *
   * The name is the part that gives way. A long workbook name truncates rather than pushing the
   * plus off the edge of a narrow pane, because a control that is only sometimes reachable is worse
   * than a name that is only sometimes complete, and the full name is on the row's tooltip.
   */
  private workbookRow(name: string, isOpen: boolean): HTMLElement {
    const row = document.createElement("div");
    row.className = "tree-workbook";
    row.dataset.project = name;
    row.tabIndex = 0;
    row.title = name;
    row.setAttribute("role", "treeitem");
    row.setAttribute("aria-expanded", String(isOpen));

    const chevron = document.createElement("span");
    chevron.className = `codicon codicon-chevron-${isOpen ? "down" : "right"}`;
    chevron.setAttribute("aria-hidden", "true");

    const icon = document.createElement("span");
    icon.className = "codicon codicon-file-code";
    icon.setAttribute("aria-hidden", "true");

    const label = document.createElement("span");
    label.className = "tree-workbook-name";
    label.textContent = name;

    const add = document.createElement("button");
    add.type = "button";
    add.className = "tree-add";
    add.dataset.addProject = name;
    // The icon is decorative and the plus is not a word, so the name of the workbook has to be in
    // the label or a screen reader hears "button" beside nine identical ones.
    add.setAttribute("aria-label", `Add to ${name}`);
    add.title = `Add to ${name}`;

    const plus = document.createElement("span");
    plus.className = "codicon codicon-add";
    plus.setAttribute("aria-hidden", "true");
    add.appendChild(plus);

    row.append(chevron, icon, label, add);
    return row;
  }

  private item(component: ExplorerComponent, workbook: string, depth = 0): HTMLElement {
    const meta = kindMeta(component.kind);
    const isUnfolded = this.isUnfolded(component.name, workbook);

    const button = document.createElement("button");
    button.type = "button";
    indent(button, depth);
    // Name AND workbook. See activeWorkbook.
    const isActive = component.name === this.active && workbook === this.activeWorkbook;
    const isSelected = component.name === this.selected && workbook === this.selectedWorkbook;

    button.className = "tree-item"
      + (isActive ? " active" : "")
      + (isSelected ? " selected" : "");
    button.dataset.component = component.name;
    button.dataset.workbook = workbook;
    button.dataset.kind = String(component.kind);
    button.setAttribute("role", "treeitem");
    button.setAttribute("aria-selected", String(isSelected));
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

  /**
   * The designer, as a child of its form. Carries its own data attribute rather than reusing the
   * procedure's, because it is not a line in a file: clicking it opens the design FACE of the
   * form's document, where a procedure row navigates to a line of its code.
   */
  private designerRow(module: string, workbook: string, depth = 0): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item tree-proc tree-designer";
    indent(button, depth);
    button.dataset.designerModule = module;
    button.dataset.designerWorkbook = workbook;
    button.title = `Open the designer for ${module}`;
    button.setAttribute("role", "treeitem");

    const glyph = document.createElement("span");
    glyph.className = "codicon codicon-symbol-color";
    glyph.setAttribute("aria-hidden", "true");

    button.append(glyph, document.createTextNode("Designer"));
    return button;
  }

  private procedureRow(module: string, procedure: ExplorerProcedure, workbook: string, depth = 0, isCurrent = false): HTMLElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "tree-item tree-proc" + (isCurrent ? " current" : "");
    button.dataset.procModule = module;
    button.dataset.procLine = String(procedure.line);
    button.dataset.procWorkbook = workbook;
    button.setAttribute("role", "treeitem");
    if (isCurrent) {
      // The mark is a colour, and a colour alone says nothing to a reader without one.
      button.setAttribute("aria-current", "location");
    }
    indent(button, depth);

    const glyph = document.createElement("span");
    glyph.className = "codicon codicon-symbol-method";
    glyph.setAttribute("aria-hidden", "true");

    button.append(glyph, document.createTextNode(`${procedure.kind} ${procedure.name}`));
    return button;
  }
}

/** The flat tree's order: documents, forms, modules, classes, and by name within a kind. */
function componentOrder(a: ExplorerComponent, b: ExplorerComponent): number {
  const left = kindMeta(a.kind);
  const right = kindMeta(b.kind);
  if (left.order !== right.order) {
    return left.order - right.order;
  }
  return a.name.localeCompare(b.name);
}

/** How many modules a folder holds, its subfolders included. */
function countModules(folder: FolderNode<ExplorerComponent>): number {
  return folder.modules.length + folder.folders.reduce((sum, child) => sum + countModules(child), 0);
}

/**
 * One level of indent per folder, as a variable the stylesheet reads. Only set past the root, so
 * the flat tree's rows carry exactly the attributes they always did.
 */
function indent(row: HTMLElement, depth: number): void {
  if (depth > 0) {
    row.style.setProperty("--tree-depth", String(depth));
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
