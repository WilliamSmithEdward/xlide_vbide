/*
 * The editor area: one or more groups, each a tab strip over a Monaco editor, arranged in a
 * split tree - the ergonomics of the studio's editor grid, built in this codebase's own idiom
 * (decision 13).
 *
 * Ownership is split three ways. The HOST owns which documents are open and which is active:
 * its pane list arrives through setOpen, and every activation is asked for, never assumed.
 * The DEVELOPER owns geography: which group a tab sits in, the order within the strip, and
 * where the splitters rest, all of which survive every host echo. The WORKSPACE owns the
 * mapping between the two: new documents join the active group, closed ones leave wherever
 * they sit, a group whose last tab leaves dissolves, and the host-active document is always
 * the active tab of some group.
 *
 * A document lives in exactly ONE group. Splitting moves the tab rather than duplicating it,
 * because the host's open list is a set - one pane per module - and a model shown twice would
 * need two view states for one identity. (The native editor's own Window > Split is a
 * different feature, and it went with the Window menu.)
 *
 * The pointer-handling here inherits the tab strip's scar tissue wholesale: press identities
 * are captured as data because a host echo can rebuild the pressed element mid-gesture;
 * drags suppress the click they would otherwise become; pointercancel is handled everywhere,
 * because the host steals focus mid-press on all sorts of occasions.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import { installEdgeScroll, type EdgeScroll } from "./edgescroll.js";
import { showContextMenu } from "./contextmenu.js";
import { docKeyOf, type DocumentId, type DocumentStore } from "./documents.js";
import { resizeAt } from "./docktree.js";
import { ALL_ZONES, DragCompass, EDGE_ZONES, zoneRect, type DropZone } from "./dragcompass.js";

export interface WorkspaceHandlers {
  /** Creates and wires a Monaco editor for a new group. The workspace owns its layout only. */
  createEditor(container: HTMLElement): monaco.editor.IStandaloneCodeEditor;
  /** The developer picked a tab or focused a group; the host is asked to activate. */
  activate(id: DocumentId): void;
  /** The developer closed a tab, however they did it; action carries a confirm answer. */
  close(id: DocumentId, action?: string): void;
  /** The active (document, editor) pair changed on this side; the frame follows. */
  activeChanged(id: DocumentId | null, editor: monaco.editor.IStandaloneCodeEditor): void;
  /** Group geometry changed, so anything measuring the page re-measures. */
  layoutChanged(): void;
}

/** How far one arrow key moves a splitter. */
const KEYBOARD_STEP = 24;

/** No group may be squeezed below this many pixels on its split axis. */
const MIN_GROUP_SIZE = 120;

type LayoutNode = SplitNode | LeafNode;

interface SplitNode {
  kind: "split";
  direction: "row" | "column";
  children: LayoutNode[];
  /** Fractions summing to 1, one per child. */
  sizes: number[];
  element: HTMLElement;
}

interface LeafNode {
  kind: "leaf";
  group: EditorGroup;
}

/** One tab as a group holds it. */
interface Tab {
  id: DocumentId;
  dirty: boolean;
}

/** One tab as the debug api reports it: what it holds, and what the strip draws for it. */
export interface TabSnapshot {
  module: string;
  project: string | null;
  /** The label as rendered, workbook and all when the name collides. */
  label: string;
  active: boolean;
  dirty: boolean;
  problems: number;
}

export interface GroupSnapshot {
  number: number;
  active: boolean;
  pending: DocumentId | null;
  /** Document keys, most recently shown first: what a close falls back through. */
  recent: string[];
  tabs: TabSnapshot[];
}

export interface WorkspaceSnapshot {
  groups: GroupSnapshot[];
  active: DocumentId | null;
  empty: boolean;
}

let nextGroupNumber = 1;

class EditorGroup {
  readonly root: HTMLElement;
  readonly strip: HTMLElement;
  private readonly tabbar: HTMLElement;
  private readonly scroller: EdgeScroll;
  readonly body: HTMLElement;
  readonly editor: monaco.editor.IStandaloneCodeEditor;

  tabs: Tab[] = [];
  active: DocumentId | null = null;

  /**
   * The documents this group has shown, most recent first - what it falls back to when it loses
   * its active tab. Position in the strip is where a tab was dropped and says nothing about what
   * anyone was reading.
   */
  private shownHere: string[] = [];

  /** A document this group means to show as soon as the page has its text. */
  pending: DocumentId | null = null;

  /** Each document's scroll and caret in THIS group, restored when its tab returns. */
  private readonly viewStates = new Map<string, monaco.editor.ICodeEditorViewState | null>();

  /** What the strip last rendered, so an echo that changes nothing rebuilds nothing. */
  private lastTabsKey: string | null = null;

  constructor(private readonly workspace: Workspace) {
    this.root = document.createElement("div");
    this.root.className = "editor-group";
    this.root.dataset.group = String(nextGroupNumber++);

    this.strip = document.createElement("div");
    this.strip.className = "group-tabs";
    this.strip.setAttribute("role", "tablist");
    this.strip.setAttribute("aria-label", "Open modules");

    // The strip is wrapped so its edges have somewhere to sit that does not scroll with it.
    // The strip itself keeps its class and its element identity, because the tab drag handlers
    // and every selector that names it are attached to exactly this node.
    this.tabbar = document.createElement("div");
    this.tabbar.className = "group-tabbar";
    this.tabbar.append(this.strip);

    this.body = document.createElement("div");
    this.body.className = "group-body";

    this.root.append(this.tabbar, this.body);
    this.scroller = installEdgeScroll(this.strip, "tab-edge");
    this.editor = workspace.handlers.createEditor(this.body);

    // Focus anywhere inside the group makes it the active group; the editor's own focus is
    // the usual way in, the strip the other.
    this.editor.onDidFocusEditorWidget(() => this.workspace.groupFocused(this));
    this.root.addEventListener("focusin", () => this.workspace.groupFocused(this));
    this.root.addEventListener("pointerdown", () => this.workspace.groupFocused(this), true);

    this.workspace.installStrip(this);
  }

  key(tab: DocumentId): string {
    return docKeyOf(tab.module, tab.project);
  }

  holds(id: DocumentId): boolean {
    const key = docKeyOf(id.module, id.project);
    return this.tabs.some((tab) => this.key(tab.id) === key);
  }

  /** Shows a document this group holds: model on the editor, view states swapped. */
  show(id: DocumentId): void {
    // A group shows what it HOLDS. Asked for anything else it would draw a module with no tab in
    // its own strip and report it as active - which a group waiting on a document whose tab closed
    // meanwhile did (2026-08-07).
    if (!this.holds(id)) {
      return;
    }

    const model = this.workspace.documents.get(id.module, id.project);
    if (!model) {
      return;
    }

    const outgoing = this.editor.getModel();
    if (outgoing === model) {
      this.setActive(id);
      return;
    }

    if (outgoing) {
      const previous = this.active;
      if (previous) {
        this.viewStates.set(docKeyOf(previous.module, previous.project), this.editor.saveViewState());
      }
    }

    this.editor.setModel(model);
    const held = this.viewStates.get(docKeyOf(id.module, id.project));
    if (held) {
      this.editor.restoreViewState(held);
    }

    this.setActive(id);
  }

  private setActive(id: DocumentId): void {
    this.active = id;
    this.pending = null;

    const key = this.key(id);
    this.shownHere = [key, ...this.shownHere.filter((held) => held !== key)];

    this.renderTabs();
  }

  /**
   * Shows something after the active tab left, and says whether it managed to.
   *
   * The most recently shown survivor, not the departed tab's neighbour in the strip: what a group
   * should fall back to is what was last being read in it (the developer, 2026-08-07).
   *
   * A tab that has never been shown HERE may have no model at all - the host publishes a module's
   * text when it is activated, not when its pane opens, so a workspace opened onto eight modules
   * holds text for the one that was looked at. A group whose every survivor is untouched cannot
   * show anything by itself, and returns false rather than leaving a blank pane unexplained.
   */
  /**
   * The surviving tab this group showed most recently, whether or not the page holds its text.
   *
   * promote() answers the narrower question, "what can I show right now", and returns nothing
   * when the page has no text for any survivor. This one answers "what SHOULD be shown", which is
   * what a caller needs when it is willing to go and fetch it.
   */
  mostRecentlyShown(): DocumentId | undefined {
    for (const key of this.shownHere) {
      const found = this.tabs.find((tab) => this.key(tab.id) === key);
      if (found) {
        return found.id;
      }
    }

    return undefined;
  }

  /** The MRU stack itself, for the debug api: what a close will fall back through, in order. */
  shownOrder(): string[] {
    return [...this.shownHere];
  }

  /**
   * Forgets every tab and everything remembered ABOUT those tabs.
   *
   * The per-tab state has to go with the tabs. Emptying `tabs` alone left the MRU stack and the
   * saved view states holding documents that were no longer open: the stack then decided which
   * tab a close falls back to, using entries from before the workspace was emptied, which is the
   * blank-view defect of 2026-08-07 arriving through a second door. The view states are a leak
   * besides - one monaco view state per document ever shown, for the life of the page.
   */
  forget(): void {
    this.tabs = [];
    this.active = null;
    this.pending = null;
    this.shownHere = [];
    this.viewStates.clear();
    this.editor.setModel(null);
  }

  promote(): boolean {
    const survivors = this.tabs.map((tab) => tab.id);
    const remembered = this.shownHere
      .map((key) => survivors.find((id) => this.key(id) === key))
      .filter((id): id is DocumentId => id !== undefined);

    for (const id of [...remembered, ...survivors]) {
      if (this.workspace.documents.get(id.module, id.project)) {
        this.show(id);
        return true;
      }
    }

    return false;
  }

  /** Drops a document from this group; true when it was the active one. */
  remove(id: DocumentId): boolean {
    const key = docKeyOf(id.module, id.project);
    const wasActive = this.active !== null && docKeyOf(this.active.module, this.active.project) === key;

    this.tabs = this.tabs.filter((tab) => this.key(tab.id) !== key);
    this.viewStates.delete(key);
    this.shownHere = this.shownHere.filter((held) => held !== key);

    // A group awaiting text for the tab that just left would wait forever: show() refuses a
    // document the group no longer holds, so nothing would ever clear `pending`, and the
    // fallback below skips any group that has one. Latent rather than seen - the host re-opens
    // every live document at ready, so the page rarely lacks a survivor's text - but a stuck
    // `pending` disables the fallback permanently, which is too sharp an edge to leave.
    if (this.pending && this.key(this.pending) === key) {
      this.pending = null;
    }

    if (wasActive) {
      this.active = null;
      if (this.editor.getModel()) {
        this.editor.setModel(null);
      }
    }

    this.renderTabs();
    return wasActive;
  }

  /**
   * Rebuilds the strip when anything it draws has changed. The render key covers it all.
   *
   * Including whether each name COLLIDES, which is not a fact about this group. A twin tab
   * opening in another group changes what this group's label should say while changing nothing
   * about its own tabs, so a key built from its own tabs alone held it back: the newly opened
   * tab showed its workbook and the one already there stayed bare, which reads as arbitrary
   * (the developer, 2026-08-07). Both names are ambiguous, so both are qualified.
   */
  renderTabs(): void {
    const counts = this.workspace.openNameCounts;
    const renderKey = this.tabs
      .map((tab) => this.key(tab.id)
        + "" + this.workspace.problemCountFor(tab.id)
        + (tab.dirty ? "d" : "")
        + ((counts.get(tab.id.module.toLowerCase()) ?? 0) > 1 ? "+" : ""))
      .join("")
      + "" + (this.active ? this.key(this.active) : "");

    if (renderKey === this.lastTabsKey) {
      return;
    }

    this.lastTabsKey = renderKey;
    this.strip.replaceChildren();

    // A name two workbooks share earns its workbook in the label; a unique name stays bare.
    // Counted across EVERY group: the collision is real wherever the twin tab sits.
    const nameCounts = this.workspace.openNameCounts;

    for (const { id, dirty } of this.tabs) {
      const isActive = this.active !== null && this.key(this.active) === this.key(id);
      const collides = (nameCounts.get(id.module.toLowerCase()) ?? 0) > 1;
      const count = this.workspace.problemCountFor(id);

      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "tab" + (isActive ? " active" : "") + (dirty ? " dirty" : "");
      tab.dataset.module = id.module;
      tab.dataset.project = id.project ?? "";
      // The workbook is added to the LABEL only when two open tabs share a module name, which is
      // exactly when the bare name is ambiguous. Putting it on every tab costs strip width, and
      // the strip runs out: it grew scroll arrows for that reason.
      //
      // Bracketed rather than dashed, because the workbook qualifies the name rather than
      // standing beside it as a second thing of equal weight.
      tab.textContent = collides && id.project ? `${id.module} (${id.project})` : id.module;

      // The tooltip always carries it, collision or not. Otherwise a bare tab offers no way at
      // all to find out which workbook it belongs to, and the answer to "which one is this?"
      // should not depend on some other tab happening to share its name (the developer,
      // 2026-08-07).
      tab.title = id.project ? `${id.module} (${id.project})` : id.module;
      tab.setAttribute("role", "tab");
      tab.setAttribute("aria-selected", String(isActive));
      tab.draggable = false;

      if (count > 0) {
        const badge = document.createElement("span");
        badge.className = "tab-badge";
        badge.textContent = String(count);
        badge.title = `${count} problem${count === 1 ? "" : "s"} in ${id.module}`;
        tab.appendChild(badge);
      }

      // The unsaved dot sits where the close box sits: the dot while the workbook is dirty,
      // the X the moment the pointer arrives. Both are always in the tree; the stylesheet
      // decides which one shows.
      const unsaved = document.createElement("span");
      unsaved.className = "tab-dirty codicon codicon-circle-filled";
      unsaved.title = "Unsaved changes in this workbook";
      tab.appendChild(unsaved);

      const close = document.createElement("span");
      close.className = "tab-close codicon codicon-close";
      close.title = "Close (Ctrl+W)";
      close.setAttribute("role", "button");
      close.setAttribute("aria-label", `Close ${id.module}`);
      tab.appendChild(close);

      this.strip.appendChild(tab);
    }
  }

  dispose(): void {
    this.scroller.dispose();
    this.editor.setModel(null);
    this.editor.dispose();
    this.root.remove();
  }
}

export class Workspace {
  readonly handlers: WorkspaceHandlers;
  readonly documents: DocumentStore;

  private readonly area: HTMLElement;
  private readonly emptyView: HTMLElement;

  private layout: LayoutNode;
  private groups: EditorGroup[] = [];
  private activeGroup: EditorGroup;

  /** The findings, for the tab badges: a finding that names its workbook counts only on the
   * matching tab, and one that cannot say counts wherever the name appears. */
  private findings: { module: string; project?: string | null }[] = [];

  /** How many open tabs share each lowercased module name, across every group. */
  openNameCounts = new Map<string, number>();

  /** True from a drag until the click it would otherwise become has been swallowed. */
  private dragSuppressesClick = false;

  /** The tab whose close box the pointer went down on, held as data a rebuild cannot destroy. */
  private pressedClose: DocumentId | null = null;

  /** The last active pair announced, so echoes do not restate it. */
  private lastAnnounced = "";

  constructor(area: HTMLElement, emptyView: HTMLElement, documents: DocumentStore, handlers: WorkspaceHandlers) {
    this.area = area;
    this.emptyView = emptyView;
    this.documents = documents;
    this.handlers = handlers;

    const first = new EditorGroup(this);
    this.groups = [first];
    this.activeGroup = first;
    this.layout = { kind: "leaf", group: first };
    this.render();

    // Empty until the host says otherwise: the first publish decides.
    this.setEmpty(true);
    this.announceActive();
  }

  /** The editor of the active group - what "the editor" means for every editor-wide feature. */
  activeEditor(): monaco.editor.IStandaloneCodeEditor {
    return this.activeGroup.editor;
  }

  /** The active group's active document, or null when the workspace is empty. */
  activeDocument(): DocumentId | null {
    return this.activeGroup.active;
  }

  /** The editor currently showing a model, or null when no group has it up. */
  editorShowing(model: monaco.editor.ITextModel): monaco.editor.IStandaloneCodeEditor | null {
    for (const group of this.groups) {
      if (group.editor.getModel() === model) {
        return group.editor;
      }
    }
    return null;
  }

  /** Every editor, for feature wiring that must reach them all. */
  editors(): monaco.editor.IStandaloneCodeEditor[] {
    return this.groups.map((group) => group.editor);
  }

  groupCount(): number {
    return this.groups.length;
  }

  /**
   * What the strip would say, as data. For the debug api's `ui` route.
   *
   * Reported from the tabs themselves rather than read back out of the DOM, because a probe
   * that scrapes the strip is measuring the render and calling it the state, and a stale
   * render is precisely the defect worth catching. The label is computed the same way
   * renderTabs computes it, so a disagreement between this and the screen IS the bug.
   */
  snapshot(): WorkspaceSnapshot {
    return {
      groups: this.groups.map((group) => ({
        number: Number(group.root.dataset.group ?? "0"),
        active: group === this.activeGroup,
        pending: group.pending ? { ...group.pending } : null,
        recent: group.shownOrder(),
        tabs: group.tabs.map(({ id, dirty }) => ({
          module: id.module,
          project: id.project,
          label: (this.openNameCounts.get(id.module.toLowerCase()) ?? 0) > 1 && id.project
            ? `${id.module} (${id.project})`
            : id.module,
          active: group.active !== null && group.key(group.active) === group.key(id),
          dirty,
          problems: this.problemCountFor(id),
        })),
      })),
      active: this.activeDocument(),
      empty: this.groups.every((group) => group.tabs.length === 0),
    };
  }

  /**
   * Adopts the host's open list. Membership is the host's; geography is the developer's:
   * tabs still open keep their group and position, new ones join the active group, closed
   * ones leave wherever they sit, and a group emptied by the diff dissolves.
   */
  setOpen(open: DocumentId[], dirty: boolean[], active: DocumentId | null): void {
    const openKeys = new Map(open.map((id, index) => [docKeyOf(id.module, id.project), index] as const));

    this.openNameCounts = new Map();
    for (const id of open) {
      const lower = id.module.toLowerCase();
      this.openNameCounts.set(lower, (this.openNameCounts.get(lower) ?? 0) + 1);
    }

    // Remove closed tabs, refresh dirty flags on the survivors.
    for (const group of [...this.groups]) {
      for (const tab of [...group.tabs]) {
        const index = openKeys.get(docKeyOf(tab.id.module, tab.id.project));
        if (index === undefined) {
          group.remove(tab.id);
        } else {
          tab.dirty = dirty[index] ?? false;
        }
      }
    }

    // Add new tabs to the active group, in the host's order among themselves - unless a drop
    // already chose where one of them goes, in which case the placement hint says the group
    // and the index, and is spent on the tab it named.
    const held = new Set(this.groups.flatMap((group) => group.tabs.map((tab) => docKeyOf(tab.id.module, tab.id.project))));
    open.forEach((id, index) => {
      if (!held.has(docKeyOf(id.module, id.project))) {
        const hint = this.placement !== null
          && this.placement.key === docKeyOf(id.module, id.project)
          && this.groups.includes(this.placement.group)
          ? this.placement
          : null;

        if (hint) {
          hint.group.tabs.splice(Math.min(hint.index, hint.group.tabs.length), 0, { id, dirty: dirty[index] ?? false });
          this.placement = null;
        } else {
          this.activeGroup.tabs.push({ id, dirty: dirty[index] ?? false });
        }
      }
    });

    // Groups emptied by the diff dissolve; the last one stays as the empty workspace.
    this.dissolveEmptyGroups();

    // A group whose active tab closed goes back to what it was showing before it: only the
    // HOST-active document gets a reveal below, and a background group must not sit blank either.
    //
    // promote() can only choose among documents the page HAS TEXT for, and the page gets a
    // module's text when it is activated, not when its tab appears. So a group whose every
    // survivor is untouched had nothing to promote to and sat blank with a full tab strip above
    // it (the developer, 2026-08-07: closing a tab went to a blank view). When that happens the
    // host is asked for the best candidate and the group shows it when the text arrives, which is
    // the same fallback moving a tab out of a group already used.
    for (const group of this.groups) {
      if (group.active || group.tabs.length === 0 || group.promote()) {
        continue;
      }

      const next = group.mostRecentlyShown() ?? group.tabs[0]?.id;
      if (next && !group.pending) {
        group.pending = next;
        this.handlers.activate(next);
      }
    }

    if (active) {
      this.reveal(active);
    }

    /*
     * THE HOST IS TOLD WHAT THE PAGE IS SHOWING, when the host itself named nothing.
     *
     * Closing the active module's pane leaves the host with no active module: it sends the
     * remaining tabs with `active: null`. The page promotes one of them to show - the strip and
     * the editor look completely normal - and the host still believes nothing is active.
     *
     * That is not cosmetic. Every language provider answers only for the HOST-active module and
     * returns nothing otherwise, so hover, completions, signature help and quick fixes all go
     * silent on a tab that looks and behaves like any other. Switching tabs cures it, because
     * that finally tells the host something, which is exactly how it was reported (the
     * developer, 2026-08-08: "hover is not working on the Watch sub keyword").
     *
     * One statement, and it converges: the host activates what it is told, republishes with that
     * active, and this branch does not run again.
     */
    const shown = this.activeGroup.active ?? this.groups.find((group) => group.active)?.active;
    if (!active && shown && open.length > 0) {
      this.handlers.activate(shown);
    }

    for (const group of this.groups) {
      group.renderTabs();
    }

    this.setEmpty(open.length === 0);
  }

  /** Re-renders every strip, for badge changes. */
  setFindings(findings: { module: string; project?: string | null }[]): void {
    this.findings = findings;
    for (const group of this.groups) {
      group.renderTabs();
    }
  }

  /** How many findings a tab's badge carries. */
  problemCountFor(id: DocumentId): number {
    return this.findings.filter((finding) =>
      finding.module.toLowerCase() === id.module.toLowerCase()
      && (finding.project == null
        || (finding.project ?? "").toLowerCase() === (id.project ?? "").toLowerCase())).length;
  }

  /**
   * Makes the host-active document the visible tab of its group. The host's word outranks
   * every page-local choice: the debugger steps into a module and the native editor
   * activates its pane, and the surface must be looking at the same code.
   */
  reveal(id: DocumentId): void {
    const owner = this.groups.find((group) => group.holds(id));
    if (!owner) {
      return;
    }

    owner.show(id);

    if (this.activeGroup !== owner) {
      this.activeGroup = owner;
      this.markActiveGroup();
    }

    this.announceActive();
  }

  /**
   * A document's text arrived. Any group that asked for it shows it.
   *
   * A group that loses its active tab can only fall back to a document the page already holds,
   * and the page holds a module's text once it has been activated - not merely because its pane
   * is open. So a group with nothing to fall back to asks the host and waits here, rather than
   * calling show() on a document that does not exist yet and silently staying blank.
   */
  documentOpened(module: string, project: string | null): void {
    const key = docKeyOf(module, project);
    for (const group of this.groups) {
      if (group.pending && docKeyOf(group.pending.module, group.pending.project) === key) {
        group.show(group.pending);
      }
    }
  }

  /** Empties everything: the host said every pane is closed. */
  clear(): void {
    for (const group of [...this.groups]) {
      group.forget();
      group.renderTabs();
    }
    this.dissolveEmptyGroups();
    this.setEmpty(true);
  }

  private setEmpty(empty: boolean): void {
    this.emptyView.hidden = !empty;
    this.area.classList.toggle("empty", empty);
  }

  /**
   * Whether the empty view is up. A different question from having no tabs: the two disagreeing
   * is a defect, and one that only shows as a blank editor with a tab strip above it.
   */
  emptyViewShown(): boolean {
    return !this.emptyView.hidden;
  }

  /** The group under focus changed. */
  groupFocused(group: EditorGroup): void {
    if (this.activeGroup === group) {
      return;
    }

    this.activeGroup = group;
    this.markActiveGroup();

    // Focusing a group activates its document, so the native active pane - the compile and
    // run target - follows the developer's eyes. Ordered BEFORE any engine request the new
    // focus produces, which is what keeps offset-only requests honest (decision 12).
    if (group.active) {
      this.handlers.activate(group.active);
    }

    this.announceActive();
  }

  private markActiveGroup(): void {
    for (const group of this.groups) {
      group.root.classList.toggle("active-group", group === this.activeGroup && this.groups.length > 1);
    }
  }

  private announceActive(): void {
    const id = this.activeGroup.active;
    const key = `${id ? docKeyOf(id.module, id.project) : ""}${this.groups.indexOf(this.activeGroup)}`;
    if (key === this.lastAnnounced) {
      return;
    }

    this.lastAnnounced = key;
    this.handlers.activeChanged(id, this.activeGroup.editor);
  }

  /** Activates the next or previous tab of the active group, wrapping, and names the choice. */
  cycleTab(delta: number): string | null {
    const group = this.activeGroup;
    if (group.tabs.length === 0) {
      return null;
    }

    const activeKey = group.active ? docKeyOf(group.active.module, group.active.project) : null;
    const at = Math.max(0, group.tabs.findIndex((tab) => docKeyOf(tab.id.module, tab.id.project) === activeKey));
    const next = group.tabs[(at + delta + group.tabs.length) % group.tabs.length];
    if (!next || docKeyOf(next.id.module, next.id.project) === activeKey) {
      return null;
    }

    this.selectTab(next.id);
    return next.id.module;
  }

  /** The active group's active tab is asked to close. */
  closeActive(): void {
    const active = this.activeGroup.active;
    if (active) {
      this.handlers.close(active);
    }
  }

  /**
   * The developer picked a tab: shown NOW, page-locally - the felt win of live models - and
   * the host is asked to activate the native pane behind it. The host's echo confirms, and a
   * refusal reconciles back on the next publish.
   */
  /**
   * Picks a tab exactly as clicking it does, for anything driving the surface from outside.
   *
   * `reveal` is NOT this. Reveal shows a document the page already holds and tells the host
   * nothing, so against a tab whose text has never been fetched it silently does nothing at all:
   * the strip lists every open pane, but the page holds a document only for modules that have
   * been activated. Driving the surface through reveal therefore reported success and moved
   * nothing, which is how it was found (2026-08-07).
   */
  pickTab(id: DocumentId): void {
    this.selectTab(id);
  }

  private selectTab(id: DocumentId, group?: EditorGroup): void {
    const owner = group ?? this.groups.find((candidate) => candidate.holds(id));
    if (!owner) {
      return;
    }

    owner.show(id);
    if (this.activeGroup !== owner) {
      this.activeGroup = owner;
      this.markActiveGroup();
    }

    this.handlers.activate(id);
    this.announceActive();
    owner.editor.focus();
  }

  /** Splits the active group's active tab into a new group beside or below it. */
  splitActive(direction: "right" | "down"): void {
    const from = this.activeGroup;
    if (from.active) {
      // moveTab refuses the degenerate case - splitting a group by its own only tab.
      this.moveTab(from.active, from, { split: { of: from, direction: direction === "right" ? "right" : "bottom" } });
    }
  }

  /* ------------------------------------------------------------------ layout tree */

  private render(): void {
    this.area.replaceChildren(this.buildDom(this.layout));
    this.markActiveGroup();

    // Explicit, after every (re)attach: an editor created in a detached container reports a
    // few pixels and its automatic layout does not recover on its own - one measure against
    // the real container gives the observer truth to track from.
    for (const group of this.groups) {
      group.editor.layout();
    }

    this.handlers.layoutChanged();
  }

  private buildDom(node: LayoutNode): HTMLElement {
    if (node.kind === "leaf") {
      return node.group.root;
    }

    const container = document.createElement("div");
    // split-row, not row: the Problems panel's generic .row rule would reach a bare class.
    container.className = `group-split split-${node.direction}`;
    node.element = container;

    node.children.forEach((child, index) => {
      if (index > 0) {
        container.appendChild(this.buildSplitter(node, index));
      }

      const cell = document.createElement("div");
      cell.className = "group-cell";
      cell.style.flex = `${node.sizes[index] ?? 1} 1 0`;
      cell.appendChild(this.buildDom(child));
      container.appendChild(cell);
    });

    return container;
  }

  /** The draggable divider before child `index` of a split. Keyboard-operable, like every splitter here. */
  private buildSplitter(node: SplitNode, index: number): HTMLElement {
    const splitter = document.createElement("div");
    splitter.className = `group-splitter split-${node.direction}`;
    splitter.setAttribute("role", "separator");
    splitter.setAttribute("aria-orientation", node.direction === "row" ? "vertical" : "horizontal");
    splitter.tabIndex = 0;

    const apply = (deltaPixels: number): void => {
      const container = node.element;
      const total = node.direction === "row" ? container.clientWidth : container.clientHeight;
      if (total <= 0) {
        return;
      }

      // The tree's own arithmetic, which is the tested copy; see the same call in paneldocks.
      node.sizes = resizeAt(node.sizes, index, deltaPixels / total, MIN_GROUP_SIZE / total);

      const cells = [...container.children].filter((child) => child.classList.contains("group-cell")) as HTMLElement[];
      cells.forEach((cell, cellIndex) => {
        cell.style.flex = `${node.sizes[cellIndex] ?? 1} 1 0`;
      });
      this.handlers.layoutChanged();
    };

    splitter.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      event.preventDefault();
      splitter.setPointerCapture(event.pointerId);

      let last = node.direction === "row" ? event.clientX : event.clientY;
      const move = (moved: PointerEvent): void => {
        const position = node.direction === "row" ? moved.clientX : moved.clientY;
        apply(position - last);
        last = position;
      };
      const end = (ended: PointerEvent): void => {
        splitter.releasePointerCapture(ended.pointerId);
        splitter.removeEventListener("pointermove", move);
        splitter.removeEventListener("pointerup", end);
        splitter.removeEventListener("pointercancel", end);
      };

      splitter.addEventListener("pointermove", move);
      splitter.addEventListener("pointerup", end);
      splitter.addEventListener("pointercancel", end);
    });

    splitter.addEventListener("keydown", (event) => {
      const step = node.direction === "row"
        ? event.key === "ArrowLeft" ? -KEYBOARD_STEP : event.key === "ArrowRight" ? KEYBOARD_STEP : 0
        : event.key === "ArrowUp" ? -KEYBOARD_STEP : event.key === "ArrowDown" ? KEYBOARD_STEP : 0;
      if (step !== 0) {
        event.preventDefault();
        apply(step);
      }
    });

    return splitter;
  }

  /** Replaces a leaf with a split holding it and a new group; returns the new group. */
  private splitLeaf(of: EditorGroup, zone: Exclude<DropZone, "center">): EditorGroup {
    const fresh = new EditorGroup(this);
    this.groups.push(fresh);

    const direction: "row" | "column" = zone === "left" || zone === "right" ? "row" : "column";
    const first = zone === "left" || zone === "top";

    const replace = (node: LayoutNode): LayoutNode => {
      if (node.kind === "leaf") {
        if (node.group !== of) {
          return node;
        }

        const leaf: LayoutNode = { kind: "leaf", group: fresh };
        return {
          kind: "split",
          direction,
          children: first ? [leaf, node] : [node, leaf],
          sizes: [0.5, 0.5],
          element: document.createElement("div"),
        };
      }

      // A split in the same direction absorbs the new group as a sibling, the way the
      // studio's grid does, instead of nesting a two-deep tree of the same axis.
      const children = node.children.map(replace);
      const flattened: LayoutNode[] = [];
      const sizes: number[] = [];
      children.forEach((child, index) => {
        if (child.kind === "split" && child.direction === node.direction) {
          for (let inner = 0; inner < child.children.length; inner++) {
            flattened.push(child.children[inner]!);
            sizes.push((node.sizes[index] ?? 1) * (child.sizes[inner] ?? 1));
          }
        } else {
          flattened.push(child);
          sizes.push(node.sizes[index] ?? 1);
        }
      });

      return { ...node, children: flattened, sizes };
    };

    this.layout = replace(this.layout);
    this.render();
    return fresh;
  }

  /** Removes dissolved groups from the tree and merges single-child splits away. */
  private dissolveEmptyGroups(): void {
    // A group standing empty FOR a placement is not empty, it is early: a tree drop on a split
    // zone carves the group before the host round trip delivers the tab, and any publish in
    // between would dissolve it and strand the hint. Same for a group whose pending document
    // is still on its way.
    const spared = (group: EditorGroup): boolean =>
      group === this.placement?.group || group.pending !== null;

    const empty = this.groups.filter((group) => group.tabs.length === 0 && !spared(group));
    if (empty.length === 0 || this.groups.length === 1) {
      return;
    }

    const keep = this.groups.filter((group) => group.tabs.length > 0 || spared(group));
    const survivors = keep.length > 0 ? keep : [this.groups[0]!];

    const prune = (node: LayoutNode): LayoutNode | null => {
      if (node.kind === "leaf") {
        return survivors.includes(node.group) ? node : null;
      }

      const children: LayoutNode[] = [];
      const sizes: number[] = [];
      node.children.forEach((child, index) => {
        const kept = prune(child);
        if (kept) {
          children.push(kept);
          sizes.push(node.sizes[index] ?? 1);
        }
      });

      if (children.length === 0) {
        return null;
      }
      if (children.length === 1) {
        return children[0]!;
      }

      const total = sizes.reduce((sum, size) => sum + size, 0) || 1;
      return { ...node, children, sizes: sizes.map((size) => size / total) };
    };

    this.layout = prune(this.layout) ?? { kind: "leaf", group: survivors[0]! };

    for (const group of this.groups) {
      if (!survivors.includes(group)) {
        group.dispose();
      }
    }
    this.groups = survivors;

    if (!this.groups.includes(this.activeGroup)) {
      this.activeGroup = this.groups[0]!;
    }

    this.render();
    this.announceActive();
  }

  /* ------------------------------------------------------------------ tab moving */

  /**
   * Moves a tab: into another group (at an index), or into a fresh split of a group. The
   * host is not consulted - geography is the developer's - but the active document may
   * change, and that IS the host's, so it is asked.
   */
  private moveTab(
    id: DocumentId,
    from: EditorGroup,
    destination: { group?: EditorGroup; index?: number; split?: { of: EditorGroup; direction: Exclude<DropZone, "center"> } },
  ): void {
    const key = docKeyOf(id.module, id.project);
    const tab = from.tabs.find((candidate) => docKeyOf(candidate.id.module, candidate.id.project) === key);
    if (!tab) {
      return;
    }

    let target: EditorGroup;
    if (destination.split) {
      // Splitting a group by its own only tab would dissolve the source and leave the same
      // picture one splitter wider; refused as a no-op.
      if (destination.split.of === from && from.tabs.length < 2) {
        return;
      }
      target = this.splitLeaf(destination.split.of, destination.split.direction);
    } else if (destination.group) {
      target = destination.group;
    } else {
      return;
    }

    if (target === from && destination.index === undefined) {
      return;
    }

    const leavingIndex = from.tabs.findIndex((candidate) =>
      docKeyOf(candidate.id.module, candidate.id.project) === key);
    const wasActive = from.remove(id);

    // The source group does not sit blank: it goes back to what it was last showing, the way
    // every tabbed editor promotes on close.
    if (wasActive && from.tabs.length > 0 && from !== target && !from.promote()) {
      // Nothing it still holds has text on this page, so there is nothing to promote TO. The
      // host is asked for the departed tab's neighbour and the group shows it when it arrives;
      // the moved tab is activated after this, at the end, so that is what ends up active.
      const neighbour = from.tabs[Math.min(Math.max(leavingIndex, 0), from.tabs.length - 1)];
      if (neighbour) {
        from.pending = neighbour.id;
        this.handlers.activate(neighbour.id);
      }
    }

    const at = destination.index === undefined
      ? target.tabs.length
      : Math.max(0, Math.min(destination.index, target.tabs.length));
    target.tabs.splice(at, 0, tab);

    target.show(id);
    this.activeGroup = target;
    this.markActiveGroup();
    this.dissolveEmptyGroups();
    for (const group of this.groups) {
      group.renderTabs();
    }

    this.handlers.activate(id);
    this.announceActive();
    target.editor.focus();
  }

  /* ------------------------------------------------------------------ drop targeting */

  /**
   * The one bar that says where a strip drop would land. Owned here rather than per drag,
   * because at most one drag exists at a time and the bar must outlive no drag.
   */
  private dropIndicator: HTMLElement | null = null;

  /**
   * Where the NEXT arriving tab of one named document goes, one shot. A tree row dropped on
   * the workspace names a module that may have no tab yet: opening is a host round trip, and
   * when the publish finally carries the new tab, this is how it lands in the group and at
   * the index the drop chose instead of being appended to whatever group is active.
   */
  private placement: { key: string; group: EditorGroup; index: number } | null = null;

  /**
   * Where a pointer mid-drag would land, with the furniture that says so.
   *
   * Over a strip, the landing is an insertion index. The SOURCE strip shows it by live
   * reorder - the dragged tab element itself moves as the pointer crosses its neighbours'
   * midpoints, so the feedback is the reorder. Every other strip shows an INSERTION BAR at
   * the index, and the bar is what makes another group's strip read as a target at all: the
   * cross-strip move worked for a week before anyone could see it, so every cross-group move
   * went the long way through the compass centre instead (the developer, 2026-08-12).
   *
   * Over a group body, the compass names the outcome: centre joins that group's tabs, an
   * edge splits it. A group is offered only the zones it can honour - `zonesFor` is the
   * caller's answer, because what is honourable depends on what is being dragged.
   */
  private landingAt(
    during: PointerEvent,
    compass: DragCompass,
    source: { strip: HTMLElement; tab: HTMLElement; group: EditorGroup } | null,
    zonesFor: (candidate: EditorGroup) => DropZone[],
  ): { strip: EditorGroup; index: number } | { group: EditorGroup; zone: DropZone } | null {
    for (const candidate of this.groups) {
      const bounds = candidate.strip.getBoundingClientRect();
      if (during.clientY >= bounds.top && during.clientY <= bounds.bottom
        && during.clientX >= bounds.left && during.clientX <= bounds.right) {
        compass.clear();

        if (source && candidate === source.group) {
          this.clearDropIndicator();
          const after = this.tabAfter(source.strip, during.clientX, source.tab);
          if (after === null) {
            source.strip.appendChild(source.tab);
          } else if (after !== source.tab) {
            source.strip.insertBefore(source.tab, after);
          }
          return { strip: candidate, index: -1 };
        }

        const after = this.tabAfter(candidate.strip, during.clientX, null);
        const index = after === null
          ? candidate.tabs.length
          : [...candidate.strip.querySelectorAll<HTMLElement>(".tab")].indexOf(after);
        this.showDropIndicator(candidate.strip, after);
        return { strip: candidate, index };
      }
    }

    this.clearDropIndicator();

    for (const candidate of this.groups) {
      const bounds = candidate.body.getBoundingClientRect();
      if (during.clientX >= bounds.left && during.clientX <= bounds.right
        && during.clientY >= bounds.top && during.clientY <= bounds.bottom) {
        const zone = compass.over(bounds, during.clientX, during.clientY, zonesFor(candidate));
        compass.preview(zone ? zoneRect(bounds, zone) : null, zone === "center" ? "join" : "new");
        return zone ? { group: candidate, zone } : null;
      }
    }

    compass.clear();
    return null;
  }

  private showDropIndicator(strip: HTMLElement, before: HTMLElement | null): void {
    if (!this.dropIndicator) {
      this.dropIndicator = document.createElement("div");
      this.dropIndicator.className = "tab-drop-indicator";
    }

    if (before) {
      strip.insertBefore(this.dropIndicator, before);
    } else {
      strip.appendChild(this.dropIndicator);
    }
  }

  private clearDropIndicator(): void {
    this.dropIndicator?.remove();
  }

  /**
   * A drag that BRINGS a document, from outside the strips: the tree's module and procedure
   * rows. Same targets, same furniture, same landings as a tab drag - a strip takes it at an
   * index, a body's compass joins or splits - but the thing dragged is a name, not a tab, so
   * a ghost chip follows the pointer where a tab drag moves the tab itself.
   *
   * On release the document's tab is MOVED there when one exists, and otherwise the landing
   * becomes a one-shot placement for the tab the host is about to deliver; either way `open`
   * runs after - the same open the row's own click performs, which for a procedure row is the
   * navigate that carries its line. The gesture and the click end in the same state, which is
   * the property everything on this surface is held to.
   */
  beginDocumentDrag(
    id: DocumentId,
    label: string,
    start: PointerEvent,
    hooks: { open: () => void; became?: () => void },
  ): void {
    if (start.button !== 0) {
      return;
    }

    const startX = start.clientX;
    const startY = start.clientY;
    let moved = false;
    let drop: { strip: EditorGroup; index: number } | { group: EditorGroup; zone: DropZone } | null = null;
    let ghost: HTMLElement | null = null;
    const compass = new DragCompass();

    const move = (during: PointerEvent): void => {
      if (!moved && Math.abs(during.clientX - startX) < 5 && Math.abs(during.clientY - startY) < 5) {
        return;
      }

      if (!moved) {
        moved = true;
        hooks.became?.();
        ghost = document.createElement("div");
        ghost.className = "drag-ghost";
        ghost.textContent = label;
        document.body.appendChild(ghost);
        compass.begin(() => {
          drop = null;
          end();
        });
      }

      if (ghost) {
        ghost.style.left = `${during.clientX + 12}px`;
        ghost.style.top = `${during.clientY + 12}px`;
      }

      // Every zone is honourable for a document that arrives from outside, except the one
      // no-op: the sole tab of a group re-joining its own group.
      const holder = this.groups.find((group) => group.holds(id));
      drop = this.landingAt(during, compass, null, (candidate) =>
        holder === candidate && candidate.tabs.length === 1 ? ["center"] : ALL_ZONES);
    };

    const end = (): void => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", end);
      window.removeEventListener("pointercancel", end);
      compass.end();
      this.clearDropIndicator();
      ghost?.remove();
      ghost = null;

      if (!moved) {
        return;
      }

      const landing = drop;
      drop = null;
      if (!landing) {
        return;
      }

      let target: EditorGroup;
      let index: number;
      if ("strip" in landing) {
        target = landing.strip;
        index = landing.index === -1 ? target.tabs.length : landing.index;
      } else if (landing.zone === "center") {
        target = landing.group;
        index = landing.group.tabs.length;
      } else {
        target = this.splitLeaf(landing.group, landing.zone);
        index = 0;
      }

      const holder = this.groups.find((group) => group.holds(id));
      if (holder) {
        this.moveTab(id, holder, { group: target, index });
      } else {
        this.placement = { key: docKeyOf(id.module, id.project), group: target, index };
        target.pending = id;
        this.activeGroup = target;
        this.markActiveGroup();
      }

      hooks.open();
    };

    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", end);
    window.addEventListener("pointercancel", end);
  }

  /* ------------------------------------------------------------------ strip wiring */

  /** Installs every strip behaviour on a group. Called once per group, from its constructor. */
  installStrip(group: EditorGroup): void {
    const strip = group.strip;

    const tabIdOf = (element: HTMLElement | null): DocumentId | null => {
      const tab = element?.closest("[data-module]") as HTMLElement | null;
      if (!tab?.dataset.module) {
        return null;
      }
      return { module: tab.dataset.module, project: tab.dataset.project || null };
    };

    strip.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      const id = tabIdOf(target);
      if (!id || target.closest(".tab-close") || this.dragSuppressesClick) {
        return;
      }

      this.selectTab(id, group);
    });

    // The X is armed at pointerdown and fired at pointerup, never on click: the press can
    // rebuild the strip (focus stirs the host, a setModules echo replaces the pressed
    // element), so the press identity is captured as DATA, which no rebuild can destroy.
    strip.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      const target = event.target as HTMLElement;
      this.pressedClose = target.closest(".tab-close") ? tabIdOf(target) : null;
    });

    strip.addEventListener("pointerup", (event) => {
      const pressed = this.pressedClose;
      this.pressedClose = null;

      if (!pressed || event.button !== 0) {
        return;
      }

      const target = event.target as HTMLElement;
      const under = tabIdOf(target);
      if (target.closest(".tab-close")
        && under
        && docKeyOf(under.module, under.project) === docKeyOf(pressed.module, pressed.project)) {
        this.handlers.close(pressed);
      }
    });

    strip.addEventListener("pointercancel", () => {
      this.pressedClose = null;
    });

    // The middle button closes, the way every tabbed editor closes - any tab, focused or
    // not. The mousedown is claimed too, so the browser's middle-click autoscroll cannot
    // swallow the click before it becomes an auxclick.
    strip.addEventListener("mousedown", (event) => {
      if (event.button === 1 && (event.target as HTMLElement).closest("[data-module]")) {
        event.preventDefault();
      }
    });

    strip.addEventListener("auxclick", (event) => {
      if (event.button !== 1) {
        return;
      }

      const id = tabIdOf(event.target as HTMLElement);
      if (id) {
        event.preventDefault();
        this.handlers.close(id);
      }
    });

    strip.addEventListener("contextmenu", (event) => {
      const id = tabIdOf(event.target as HTMLElement);
      if (!id) {
        return;
      }

      event.preventDefault();
      const others = this.groups.flatMap((g) => g.tabs).filter((tab) =>
        docKeyOf(tab.id.module, tab.id.project) !== docKeyOf(id.module, id.project));

      showContextMenu(event.clientX, event.clientY, [
        { label: "Close", run: () => this.handlers.close(id) },
        {
          label: "Close Others",
          enabled: others.length > 0,
          run: () => others.forEach((tab) => this.handlers.close(tab.id)),
        },
        {
          label: "Close All",
          run: () => this.groups.flatMap((g) => [...g.tabs]).forEach((tab) => this.handlers.close(tab.id)),
        },
        // Splitting is not offered here. It is a placement, and every other way of placing a tab
        // is direct: drag it where it should go, or press Ctrl+\. A menu entry that splits
        // somewhere the developer cannot see before choosing is a worse version of the gesture
        // that already exists (developer, 2026-08-06).
      ]);
    });

    this.installTabDrag(group);
  }

  /**
   * Makes the strip's tabs draggable: within the strip to reorder, onto another strip to
   * move, onto a group's body to move there or split it - the drop zone overlay says which.
   *
   * Reordering moves the element in the DOM as the pointer crosses its neighbours'
   * midpoints, so the feedback is the reorder itself; the element is moved rather than
   * re-rendered because the pointer capture that keeps the drag alive belongs to it.
   */
  private installTabDrag(group: EditorGroup): void {
    const strip = group.strip;

    strip.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      const target = event.target as HTMLElement;
      const tab = target.closest(".tab") as HTMLElement | null;
      if (!tab || target.closest(".tab-close") || !tab.dataset.module) {
        return;
      }

      const id: DocumentId = { module: tab.dataset.module, project: tab.dataset.project || null };
      const startX = event.clientX;
      const startY = event.clientY;
      const pointerId = event.pointerId;
      let moved = false;
      let drop: { group: EditorGroup; zone: DropZone } | { strip: EditorGroup; index: number } | null = null;

      // The same compass the tool panes use: a code editor tab arranges against the other
      // code editor tabs with the identical gesture (developer, 2026-08-06). The editor is
      // a contained unit, so the compass only ever appears over the editor's own groups -
      // a module tab never leaves the editor area, and a tool pane never enters it.
      const compass = new DragCompass();

      const move = (during: PointerEvent): void => {
        if (!moved && Math.abs(during.clientX - startX) < 5 && Math.abs(during.clientY - startY) < 5) {
          return;
        }

        if (!moved) {
          moved = true;
          this.dragSuppressesClick = true;
          tab.classList.add("dragging");
          compass.begin(() => {
            drop = null;
            end();
          });

          try {
            tab.setPointerCapture(pointerId);
          } catch {
            // A pointer that has already gone cannot be captured; the window listeners
            // still finish the drag.
          }
        }

        // A group is offered only what it can honour. Over the tab's OWN group, centre is
        // where it already is, and a split is impossible when it is the only tab - the tab
        // would leave the group and dissolve it, which is the same picture one splitter
        // wider. Showing a zone that does nothing is a promise the drop cannot keep, and it
        // reads as a bug from outside (developer, 2026-08-06).
        drop = this.landingAt(during, compass, { strip, tab, group }, (candidate) =>
          candidate !== group ? ALL_ZONES
            : group.tabs.length > 1 ? EDGE_ZONES
            : []);
      };

      const end = (): void => {
        window.removeEventListener("pointermove", move);
        window.removeEventListener("pointerup", end);
        window.removeEventListener("pointercancel", end);
        tab.classList.remove("dragging");
        compass.end();
        this.clearDropIndicator();

        if (!moved) {
          return;
        }

        const landing = drop;
        drop = null;

        if (landing && "strip" in landing) {
          if (landing.strip === group) {
            // The DOM already shows the new order; adopt it as the truth.
            group.tabs = [...strip.querySelectorAll<HTMLElement>("[data-module]")]
              .filter((element) => !!element.dataset.module)
              .map((element) => {
                const key = docKeyOf(element.dataset.module!, element.dataset.project || null);
                return group.tabs.find((t) => docKeyOf(t.id.module, t.id.project) === key)!;
              })
              .filter(Boolean);
          } else {
            this.moveTab(id, group, { group: landing.strip, index: landing.index });
          }
        } else if (landing && "group" in landing) {
          if (landing.zone === "center") {
            if (landing.group !== group) {
              this.moveTab(id, group, { group: landing.group });
            }
          } else {
            this.moveTab(id, group, { split: { of: landing.group, direction: landing.zone } });
          }
        } else {
          // A drag that landed nowhere puts the strip back the way the tabs say it is.
          group.renderTabs();
        }

        // Cleared after the click this drag produces has already been ignored.
        setTimeout(() => {
          this.dragSuppressesClick = false;
        }, 0);
      };

      window.addEventListener("pointermove", move);
      window.addEventListener("pointerup", end);
      // The host steals focus mid-press on all sorts of occasions, and an interrupted
      // pointer stream ends in pointercancel, never pointerup.
      window.addEventListener("pointercancel", end);
    });
  }

  /** The tab the dragged one should sit before at this pointer position, or null for the end. */
  private tabAfter(strip: HTMLElement, x: number, dragging: HTMLElement | null): HTMLElement | null {
    for (const tab of strip.querySelectorAll<HTMLElement>(".tab")) {
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
}
