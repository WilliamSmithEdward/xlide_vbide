/*
 * The tool-pane docking: four sections around the editor - left, right, top, bottom - each
 * holding a split tree of PANEL GROUPS, each group a tab strip over one visible pane. The
 * studio's docking, in this codebase's own idiom (decision 13), and the same shape the
 * editor groups use for code: a pane is dragged by its title tab, a zone preview names the
 * landing - the middle of a group joins its tabs, an edge splits the section beside it, an
 * edge of the editor area begins a new section on that side.
 *
 * Every pane is a SEAT: its tab button and its body elements move together, wherever the
 * developer puts them. WHICH panes exist is the shell's; WHERE they sit is the developer's,
 * and survives the session in localStorage like the splitter positions do.
 *
 * A section exists while it has panes: the last tab leaving a section removes it and the
 * editor takes the room back. There is no collapse chevron - a pane the developer does not
 * want open goes where they put things they are not using, which is any other section or
 * none (the chevron went 2026-08-06, by request).
 */

import {
  firstGroup as firstGroupOf,
  indexAtMidpoints,
  prune,
  resizeAt,
  splitBeside,
  type TreeGroup,
  type TreeNode,
  type TreeSplit,
} from "./docktree.js";
import { ALL_ZONES, DragCompass, EDGE_ZONES, zoneRect, type DropZone } from "./dragcompass.js";
import { beginLiveDrag, endLiveDrag } from "./livedrag.js";

export type DockSide = "left" | "right" | "top" | "bottom";

/** One pane as the dock system holds it: strip button and body elements, moved as one. */
export interface PanelSeat {
  name: string;
  title: string;
  tab: HTMLButtonElement;
  bodies: HTMLElement[];
  /** Called when the pane becomes the shown tab of its group. */
  onShown?: () => void;
  /**
   * True for a pane that may never be closed. The explorer is the only route back to a
   * module once every tab is shut, so closing it could strand the developer.
   */
  permanent?: boolean;
}

export interface PanelDockHandlers {
  /** A pane's visibility changed: it became (or stopped being) some group's shown tab. */
  visibilityChanged(name: string, visible: boolean): void;
  /** Geometry changed, so anything measuring the page re-measures. */
  layoutChanged(): void;
  /** The set of open panes changed, so anything listing them (settings) re-reads. */
  openPanesChanged?(): void;
}

// The tree itself, and the arithmetic on it, live in docktree.ts - pure, and unit tested,
// because pruning and collapsing are where a docking layout's bugs are and a drag is a slow
// way to find them. Here a tab is a pane's name.
type Node = TreeNode<string>;
type Group = TreeGroup<string>;

interface StoredLayout {
  sides?: Partial<Record<DockSide, Node | null>>;
  sizes?: Partial<Record<DockSide, number>>;
  closed?: string[];
}

const STORAGE_KEY = "xlide.docks.v1";

/** The arrangement a first run gets: explorer over properties on the left, the four tool
 * panes tabbed along the bottom - today's layout, said in the new vocabulary. */
function defaultLayout(): Record<DockSide, Node | null> {
  return {
    left: {
      kind: "split",
      direction: "column",
      children: [
        { kind: "group", tabs: ["explorer"], active: "explorer" },
        { kind: "group", tabs: ["properties"], active: "properties" },
      ],
      sizes: [0.62, 0.38],
    },
    right: null,
    top: null,
    bottom: { kind: "group", tabs: ["problems", "immediate", "locals", "watch"], active: "problems" },
  };
}

const DEFAULT_SIZES: Record<DockSide, number> = { left: 260, right: 340, top: 200, bottom: 200 };

/** No section or pane may be squeezed below this on its axis. */
const MIN_PANE = 90;

const KEYBOARD_STEP = 24;

type DropTarget =
  | { kind: "group"; side: DockSide; path: number[]; zone: DropZone }
  /** Onto a strip, at a position in it: reorder within a group, or join one at a place. */
  | { kind: "strip"; side: DockSide; path: number[]; index: number }
  | { kind: "side"; side: DockSide };

export class PanelDocks {
  private readonly seats = new Map<string, PanelSeat>();
  private layout: Record<DockSide, Node | null>;
  private sizes: Record<DockSide, number> = { ...DEFAULT_SIZES };

  /**
   * Panes the developer has closed. A closed pane keeps nothing - where it sat is forgotten
   * - and comes back through the settings dialog, docked at the bottom where a returning
   * pane is findable. Closing is a choice about clutter, not a way to lose a pane: every
   * closed one is listed, checked back on, and returns.
   */
  private closed = new Set<string>();

  private readonly dockElements: Record<DockSide, HTMLElement>;
  private readonly dockSplitters: Record<DockSide, HTMLElement>;

  /** Where dragged panes read the editor area's edge bands from. */
  private readonly editorArea: HTMLElement;
  private readonly emptyView: HTMLElement | null;

  /** Where a closed pane's elements wait, out of the tree. */
  private readonly parking: HTMLElement;

  private readonly handlers: PanelDockHandlers;

  /** Rendered strip/body hosts per group, rebuilt with the layout, indexed by side+path. */
  private groupHosts: { side: DockSide; path: number[]; group: Group; strip: HTMLElement; body: HTMLElement }[] = [];

  /** What each pane's visibility was last told, so echoes are not restated. */
  private readonly toldVisibility = new Map<string, boolean>();

  /** True from a pane drag until the click it would become has been swallowed. */
  private dragArmed = false;

  /**
   * The region a pane docks against when it is dragged over the editor.
   *
   * With no module open, #editor-area is display:none and measures nothing, and the empty view
   * stands in the same place. Measuring only the editor area meant that dragging a pane to the
   * editor's edge did nothing at all until a module had been opened -- no compass, no preview,
   * no drop -- which reads as the drag being broken rather than as the editor being empty
   * (developer, 2026-08-06: "it only happens when an editor tab has not yet been loaded").
   */
  private editorRegion(): DOMRect {
    const area = this.editorArea.getBoundingClientRect();
    if (area.width > 0 && area.height > 0) {
      return area;
    }

    const empty = this.emptyView?.getBoundingClientRect();
    return empty && empty.width > 0 && empty.height > 0 ? empty : area;
  }

  constructor(
    docks: Record<DockSide, HTMLElement>,
    splitters: Record<DockSide, HTMLElement>,
    editorArea: HTMLElement,
    emptyView: HTMLElement | null,
    parking: HTMLElement,
    seats: PanelSeat[],
    handlers: PanelDockHandlers,
  ) {
    this.dockElements = docks;
    this.dockSplitters = splitters;
    this.editorArea = editorArea;
    this.emptyView = emptyView;
    this.parking = parking;
    this.handlers = handlers;

    for (const seat of seats) {
      this.seats.set(seat.name, seat);
      seat.tab.dataset.panel = seat.name;
    }

    this.layout = this.load() ?? defaultLayout();
    this.pruneUnknown();
    for (const side of ["left", "right", "top", "bottom"] as DockSide[]) {
      this.installDockSplitter(side);
    }
    this.render();
  }

  /* ------------------------------------------------------------------ public surface */

  /** Brings a pane forward: its group shows it, wherever the developer left it. A pane
   * that was closed comes back first, so every route to a pane finds one. */
  reveal(name: string): void {
    if (this.closed.has(name)) {
      this.open(name);
      return;
    }

    const found = this.findPane(name);
    if (!found) {
      return;
    }

    found.group.active = name;
    this.render();
    this.seats.get(name)?.onShown?.();
  }

  /** Whether a pane is the shown tab of its group right now. */
  isVisible(name: string): boolean {
    const found = this.findPane(name);
    return found !== null && found.group.active === name;
  }

  /** Every pane, with whether it is open and whether it may be closed at all. */
  paneStates(): { name: string; title: string; open: boolean; permanent: boolean }[] {
    return [...this.seats.values()].map((seat) => ({
      name: seat.name,
      title: seat.title,
      open: !this.closed.has(seat.name),
      permanent: seat.permanent === true,
    }));
  }

  /** Closes a pane: it leaves its group, and the group dissolves if it was the last tab. */
  close(name: string): void {
    const seat = this.seats.get(name);
    if (!seat || seat.permanent || this.closed.has(name)) {
      return;
    }

    const found = this.findPane(name);
    if (found) {
      this.removePane(name, found.group);
    }

    this.closed.add(name);

    // Parked out of the tree, so a stale reference cannot draw it.
    for (const element of seat.bodies) {
      element.hidden = true;
      this.parking.appendChild(element);
    }
    seat.tab.remove();

    this.render();
    this.handlers.openPanesChanged?.();
  }

  /** Re-opens a closed pane, at the bottom, where a returning pane is easy to find. */
  open(name: string): void {
    if (!this.closed.delete(name)) {
      return;
    }

    const bottom = this.layout.bottom;
    if (!bottom) {
      this.layout.bottom = { kind: "group", tabs: [name], active: name };
    } else {
      const group = firstGroupOf(bottom);
      if (group) {
        group.tabs.push(name);
        group.active = name;
      }
    }

    this.render();
    this.seats.get(name)?.onShown?.();
    this.handlers.openPanesChanged?.();
  }

  /* ------------------------------------------------------------------ layout bookkeeping */

  private findPane(name: string): { side: DockSide; group: Group } | null {
    for (const side of ["left", "right", "top", "bottom"] as DockSide[]) {
      const group = this.findInNode(this.layout[side], name);
      if (group) {
        return { side, group };
      }
    }
    return null;
  }

  private findInNode(node: Node | null, name: string): Group | null {
    if (!node) {
      return null;
    }
    if (node.kind === "group") {
      return node.tabs.includes(name) ? node : null;
    }
    for (const child of node.children) {
      const found = this.findInNode(child, name);
      if (found) {
        return found;
      }
    }
    return null;
  }

  /** Drops seats that no longer exist and normalises what the storage said. */
  private pruneUnknown(): void {
    // A pane the storage says is closed is not placed; a permanent one can never be.
    for (const name of [...this.closed]) {
      if (!this.seats.has(name) || this.seats.get(name)?.permanent) {
        this.closed.delete(name);
      }
    }

    const known = new Set([...this.seats.keys()].filter((name) => !this.closed.has(name)));
    const placed = new Set<string>();

    const prune = (node: Node | null): Node | null => {
      if (!node) {
        return null;
      }
      if (node.kind === "group") {
        const tabs = node.tabs.filter((tab) => known.has(tab) && !placed.has(tab));
        tabs.forEach((tab) => placed.add(tab));
        if (tabs.length === 0) {
          return null;
        }
        return { kind: "group", tabs, active: tabs.includes(node.active) ? node.active : tabs[0]! };
      }

      const children: Node[] = [];
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
      return { kind: "split", direction: node.direction, children, sizes: sizes.map((s) => s / total) };
    };

    for (const side of ["left", "right", "top", "bottom"] as DockSide[]) {
      this.layout[side] = prune(this.layout[side]);
    }

    // A pane the stored layout forgot still exists somewhere: the bottom, joining its kin.
    for (const name of known) {
      if (!placed.has(name)) {
        const bottom = this.layout.bottom;
        if (bottom && bottom.kind === "group") {
          bottom.tabs.push(name);
        } else if (bottom && bottom.kind === "split") {
          const first = firstGroupOf(bottom);
          first?.tabs.push(name);
        } else {
          this.layout.bottom = { kind: "group", tabs: [name], active: name };
        }
      }
    }
  }

  private load(): Record<DockSide, Node | null> | null {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return null;
      }
      const stored = JSON.parse(raw) as StoredLayout;
      const sides = stored.sides ?? {};
      this.closed = new Set(stored.closed ?? []);
      for (const side of ["left", "right", "top", "bottom"] as DockSide[]) {
        const size = stored.sizes?.[side];
        if (typeof size === "number" && Number.isFinite(size)) {
          this.sizes[side] = Math.max(MIN_PANE, Math.min(900, Math.round(size)));
        }
      }
      return {
        left: sides.left ?? null,
        right: sides.right ?? null,
        top: sides.top ?? null,
        bottom: sides.bottom ?? null,
      };
    } catch {
      return null;
    }
  }

  private persist(): void {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        sides: this.layout,
        sizes: this.sizes,
        closed: [...this.closed],
      }));
    } catch (error) {
      // Storage can be full or off. The arrangement simply will not survive the reload, which is
      // survivable -- but silently, and "my layout keeps resetting" is then a mystery with nothing
      // behind it. Said once: this runs on every drag, and a broken store would otherwise fill the
      // console with the same line.
      if (!this.warnedAboutStorage) {
        this.warnedAboutStorage = true;
        console.warn("[xlide] the pane arrangement could not be saved; it will not survive a reload", error);
      }
    }
  }

  private warnedAboutStorage = false;

  /* ------------------------------------------------------------------ rendering */

  private render(): void {
    this.groupHosts = [];

    for (const side of ["left", "right", "top", "bottom"] as DockSide[]) {
      const dock = this.dockElements[side];
      const node = this.layout[side];
      const alive = node !== null;

      dock.hidden = !alive;
      this.dockSplitters[side].hidden = !alive;
      dock.style.setProperty("--dock-size", `${this.sizes[side]}px`);

      dock.replaceChildren();
      if (node) {
        dock.appendChild(this.buildNode(side, node, []));
      }
    }

    this.handlers.layoutChanged();
    this.emitVisibility();
    this.persist();
  }

  private buildNode(side: DockSide, node: Node, path: number[]): HTMLElement {
    if (node.kind === "group") {
      return this.buildGroup(side, node, path);
    }

    const container = document.createElement("div");
    container.className = `dock-split split-${node.direction}`;

    node.children.forEach((child, index) => {
      if (index > 0) {
        container.appendChild(this.buildPaneSplitter(node, index, container));
      }
      const cell = document.createElement("div");
      cell.className = "dock-cell";
      cell.style.flex = `${node.sizes[index] ?? 1} 1 0`;
      cell.appendChild(this.buildNode(side, child, [...path, index]));
      container.appendChild(cell);
    });

    return container;
  }

  private buildGroup(side: DockSide, group: Group, path: number[]): HTMLElement {
    const root = document.createElement("div");
    root.className = "dock-group";

    const head = document.createElement("div");
    head.className = "dock-head";

    const strip = document.createElement("div");
    strip.className = "dock-tabs";
    strip.setAttribute("role", "tablist");

    const body = document.createElement("div");
    body.className = "dock-body";

    for (const name of group.tabs) {
      const seat = this.seats.get(name);
      if (!seat) {
        continue;
      }

      strip.appendChild(seat.tab);
      const isActive = group.active === name;
      seat.tab.classList.toggle("active", isActive);
      seat.tab.setAttribute("aria-selected", String(isActive));

      for (const element of seat.bodies) {
        body.appendChild(element);
        element.hidden = !isActive;
      }
    }

    head.appendChild(strip);

    // The close box acts on the SHOWN pane, at the group's top right, the way every docked
    // pane closes. A permanent pane has none: the explorer is the only route back to a
    // module, so it may not be closed at all.
    const shown = this.seats.get(group.active);
    if (shown && !shown.permanent) {
      const close = document.createElement("button");
      close.type = "button";
      close.className = "dock-close codicon codicon-close";
      close.title = `Close ${shown.title}`;
      close.setAttribute("aria-label", `Close ${shown.title}`);
      close.addEventListener("click", () => this.close(shown.name));
      head.appendChild(close);
    }

    root.append(head, body);
    this.installStrip(side, group, path, strip, body);
    this.groupHosts.push({ side, path, group, strip, body });
    return root;
  }

  /* ------------------------------------------------------------------ interaction */

  private installStrip(side: DockSide, group: Group, path: number[], strip: HTMLElement, body: HTMLElement): void {
    strip.addEventListener("click", (event) => {
      const tab = (event.target as HTMLElement).closest(".panel-tab") as HTMLElement | null;
      const name = tab?.dataset.panel;
      if (!name || this.dragArmed) {
        return;
      }

      group.active = name;
      this.render();
      this.seats.get(name)?.onShown?.();
    });

    strip.addEventListener("pointerdown", (event) => this.armDrag(event, group));
    void side; void path; void body;
  }

  /**
   * The pane drag: by the title tab, with a zone preview naming every landing. Window-level
   * listeners so a fast drag survives leaving the strip; pointercancel handled because the
   * host steals focus mid-gesture; the click a drag would become is swallowed.
   */
  private armDrag(event: PointerEvent, from: Group): void {
    if (event.button !== 0) {
      return;
    }

    const tab = (event.target as HTMLElement).closest(".panel-tab") as HTMLElement | null;
    const name = tab?.dataset.panel;
    if (!tab || !name) {
      return;
    }

    const startX = event.clientX;
    const startY = event.clientY;
    let moved = false;
    let target: DropTarget | null = null;

    const compass = new DragCompass();

    const move = (during: PointerEvent): void => {
      if (!moved && Math.abs(during.clientX - startX) < 5 && Math.abs(during.clientY - startY) < 5) {
        return;
      }

      if (!moved) {
        moved = true;
        this.dragArmed = true;
        tab.classList.add("dragging");
        compass.begin(() => {
          target = null;
          end(new PointerEvent("pointercancel"));
        });
        try {
          tab.setPointerCapture(during.pointerId);
        } catch {
          // A pointer already gone cannot be captured; the window listeners still finish.
        }
      }

      target = null;

      const inside = (box: DOMRect): boolean =>
        during.clientX >= box.left && during.clientX <= box.right
        && during.clientY >= box.top && during.clientY <= box.bottom;

      /*
       * THE SMALLEST BODY UNDER THE POINTER WINS, not the first one found.
       *
       * A section's body can extend UNDER a neighbouring section: with a pane docked right, the
       * right body measured 364..704 by 80..1239 and the bottom body 265..520 by 1064..1239, so
       * they overlap in the bottom-right corner and a point there is inside both. `inside` is
       * arithmetic on boxes with no notion of stacking or clipping, so a linear scan answered
       * with whichever came first in render order.
       *
       * That was not a cosmetic mix-up. When the wrong group won and it happened to be the
       * dragged pane's OWN group, the allowed zones came out empty - centre is where it already
       * is, and a lone tab cannot split - so no petal appeared at all and the drop silently did
       * nothing. A pane dragged out to the right could not be dragged back: the churn probe read
       * that as a leaked dock group, because a group it expected to dissolve never did
       * (2026-08-08).
       *
       * Smallest area is the right tie-break: the sections that overlap do so because one
       * extends beneath the other, and the one actually drawn at that point is the smaller.
       */
      const overlapping = this.groupHosts
        .map((host) => ({ host, box: host.body.getBoundingClientRect() }))
        .filter((candidate) => inside(candidate.box))
        .sort((left, right) =>
          (left.box.width * left.box.height) - (right.box.width * right.box.height));

      const hovered = overlapping[0];

      for (const host of this.groupHosts) {
        if (inside(host.strip.getBoundingClientRect())) {
          const index = this.tabIndexAt(host.strip, during.clientX, name);
          target = { kind: "strip", side: host.side, path: host.path, index };

          // The strip shows the landing itself, by moving the tab as the pointer crosses
          // its neighbours' midpoints - the same feedback the editor's tabs give, and
          // clearer than any overlay could be about where a tab will sit.
          const before = [...host.strip.querySelectorAll<HTMLElement>(".panel-tab")]
            .filter((one) => one !== tab)[index] ?? null;
          if (before) {
            host.strip.insertBefore(tab, before);
          } else {
            host.strip.appendChild(tab);
          }

          compass.clear();
          return;
        }

        // Only the group the pointer is actually over, which is the smallest body containing
        // it rather than the first one that happens to.
        if (hovered && hovered.host === host) {
          const box = hovered.box;

          // Only what the group can honour: over the pane's OWN group, centre is where it
          // already is, and a split is impossible when it is the only tab there - a zone
          // that does nothing is a promise the drop cannot keep.
          const allowed = host.group !== from ? ALL_ZONES
            : from.tabs.length > 1 ? EDGE_ZONES
            : [];

          const zone = compass.over(box, during.clientX, during.clientY, allowed);
          target = zone ? { kind: "group", side: host.side, path: host.path, zone } : null;
          // The centre joins this group's tabs; an edge carves a new group beside it.
          compass.preview(zone ? zoneRect(box, zone) : null, zone === "center" ? "join" : "new");
          return;
        }
      }

      // Over the editor itself: the compass's edges begin a section on that side. Measured
      // against the EDITOR AREA, not the workspace, which includes the sections themselves.
      // No centre here - a tool pane never joins the editor's own tabs.
      const area = this.editorRegion();
      if (area.width > 0 && area.height > 0 && inside(area)) {
        const zone = compass.over(area, during.clientX, during.clientY, EDGE_ZONES);
        target = zone && zone !== "center" ? { kind: "side", side: zone } : null;

        if (!zone || zone === "center") {
          compass.preview(null);
        } else if (this.layout[zone]) {
          // That section already stands: the pane joins it, so the preview is the section
          // itself rather than a half of the editor the drop would not touch.
          compass.preview(this.dockElements[zone].getBoundingClientRect(), "join");
        } else {
          compass.preview(zoneRect(area, zone), "new");
        }

        return;
      }

      compass.clear();
    };

    const end = (ended: PointerEvent): void => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", end);
      window.removeEventListener("pointercancel", end);
      tab.classList.remove("dragging");
      compass.end();

      if (!moved) {
        return;
      }

      setTimeout(() => {
        this.dragArmed = false;
      }, 0);

      if (ended.type === "pointercancel" || !target) {
        return;
      }

      this.dropPane(name, from, target);
    };

    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", end);
    window.addEventListener("pointercancel", end);
  }

  /**
   * Where in a strip a pointer at this x means, counting the tabs the dragged one is not:
   * before the first tab whose midpoint the pointer has not yet crossed, or at the end.
   */
  private tabIndexAt(strip: HTMLElement, x: number, dragging: string): number {
    const midpoints = [...strip.querySelectorAll<HTMLElement>(".panel-tab")]
      .filter((tab) => tab.dataset.panel !== dragging)
      .map((tab) => {
        const box = tab.getBoundingClientRect();
        return box.left + box.width / 2;
      });

    return indexAtMidpoints(midpoints, x);
  }

  /** Executes a drop: the pane leaves its group and lands where the preview said. */
  /**
   * Moves a pane to a dock side exactly as dropping it there does.
   *
   * The docking gestures had no way of being driven at all: `layout()` could report the
   * arrangement and `resetLayout()` could put it back, but nothing could MOVE anything, so the
   * whole surface was reachable only by hand with a mouse. Synthesising the drag is not an
   * option either, and this repo has the scars: the drop arms on pointerdown and completes on
   * pointerup against a compass that tracks the pointer, so a synthetic event sequence tests the
   * synthesiser rather than the product.
   *
   * This is the method the real drop calls, with the target a drop on that side would build.
   * Answers false when there is no such pane, rather than pretending.
   */
  movePaneTo(name: string, side: DockSide): boolean {
    const home = this.findPane(name);
    if (!home) {
      return false;
    }

    this.dropPane(name, home.group, { kind: "side", side });
    return true;
  }

  private dropPane(name: string, from: Group, target: DropTarget): void {
    // Dropping a lone pane onto its own group's compass is the identity, whatever the zone.
    // A strip drop is NOT: that is a reorder, and a single tab reorders to the same place.
    if (target.kind === "group") {
      const onto = this.nodeAt(target.side, target.path);
      if (onto === from && from.tabs.length === 1) {
        return;
      }
    }

    this.removePane(name, from);

    if (target.kind === "side") {
      const existing = this.layout[target.side];
      if (!existing) {
        this.layout[target.side] = { kind: "group", tabs: [name], active: name };
      } else {
        const group = firstGroupOf(existing);
        if (group) {
          group.tabs.push(name);
          group.active = name;
        }
      }
    } else {
      const node = this.nodeAt(target.side, target.path);
      if (!node || node.kind !== "group") {
        // The tree changed under the drag (the removal pruned it); land on the side.
        this.dropPane(name, { kind: "group", tabs: [name], active: name }, { kind: "side", side: target.side });
        return;
      }

      if (target.kind === "strip") {
        // A place in the strip, whether that is a reorder within this group or a join at a
        // chosen position. The index counted the tabs the dragged one is not, so it is
        // already an index into the list this splice sees.
        const at = Math.max(0, Math.min(target.index, node.tabs.length));
        node.tabs.splice(at, 0, name);
        node.active = name;
      } else if (target.zone === "center") {
        if (!node.tabs.includes(name)) {
          node.tabs.push(name);
        }
        node.active = name;
      } else {
        this.splitGroup(target.side, node, target.zone, { kind: "group", tabs: [name], active: name });
      }
    }

    this.render();
    this.seats.get(name)?.onShown?.();
  }

  /** The node a path names within a side's tree. */
  private nodeAt(side: DockSide, path: number[]): Node | null {
    let node = this.layout[side];
    for (const index of path) {
      if (!node || node.kind !== "split") {
        return null;
      }
      node = node.children[index] ?? null;
    }
    return node;
  }

  /** Takes a pane out of a group, dissolving the group and its empty ancestors. */
  private removePane(name: string, from: Group): void {
    from.tabs = from.tabs.filter((tab) => tab !== name);
    if (from.active === name) {
      from.active = from.tabs[0] ?? "";
    }

    for (const side of ["left", "right", "top", "bottom"] as DockSide[]) {
      this.layout[side] = prune(this.layout[side]);
    }
  }

  /** Splits a group within its section, putting the newcomer beside it where the zone said. */
  private splitGroup(side: DockSide, of: Group, zone: "left" | "right" | "top" | "bottom", newcomer: Group): void {
    this.layout[side] = splitBeside(this.layout[side], of, zone, newcomer);
  }

  /* ------------------------------------------------------------------ splitters */

  /** The divider between a section and the editor: drags the whole section's size. */
  private installDockSplitter(side: DockSide): void {
    const splitter = this.dockSplitters[side];
    const horizontal = side === "top" || side === "bottom";

    // Geometry only: automaticLayout's observer tracks the editors during a drag, and the one
    // settling layout runs at the drag's end (see livedrag.ts). The keyboard path below lays
    // out immediately - a discrete step deserves its frame.
    const apply = (delta: number): void => {
      const grow = side === "left" || side === "top" ? delta : -delta;
      this.sizes[side] = Math.max(MIN_PANE, Math.min(900, this.sizes[side] + grow));
      this.dockElements[side].style.setProperty("--dock-size", `${this.sizes[side]}px`);
    };

    splitter.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      event.preventDefault();
      splitter.setPointerCapture(event.pointerId);
      beginLiveDrag();
      let last = horizontal ? event.clientY : event.clientX;

      const move = (moved: PointerEvent): void => {
        const position = horizontal ? moved.clientY : moved.clientX;
        apply(position - last);
        last = position;
      };
      const end = (ended: PointerEvent): void => {
        splitter.releasePointerCapture(ended.pointerId);
        splitter.removeEventListener("pointermove", move);
        splitter.removeEventListener("pointerup", end);
        splitter.removeEventListener("pointercancel", end);
        endLiveDrag(() => this.handlers.layoutChanged());
        this.persist();
      };

      splitter.addEventListener("pointermove", move);
      splitter.addEventListener("pointerup", end);
      splitter.addEventListener("pointercancel", end);
    });

    splitter.addEventListener("keydown", (event) => {
      const step = horizontal
        ? event.key === "ArrowUp" ? -KEYBOARD_STEP : event.key === "ArrowDown" ? KEYBOARD_STEP : 0
        : event.key === "ArrowLeft" ? -KEYBOARD_STEP : event.key === "ArrowRight" ? KEYBOARD_STEP : 0;
      if (step !== 0) {
        event.preventDefault();
        apply(step);
        this.handlers.layoutChanged();
        this.persist();
      }
    });
  }

  /** The divider between two panes inside a section. */
  private buildPaneSplitter(node: TreeSplit<string>, index: number, container: HTMLElement): HTMLElement {
    const splitter = document.createElement("div");
    splitter.className = `group-splitter split-${node.direction}`;
    splitter.setAttribute("role", "separator");
    splitter.setAttribute("aria-orientation", node.direction === "row" ? "vertical" : "horizontal");
    splitter.tabIndex = 0;

    const apply = (deltaPixels: number): void => {
      const total = node.direction === "row" ? container.clientWidth : container.clientHeight;
      if (total <= 0) {
        return;
      }

      // The tree's own arithmetic, which is the tested copy. This used to be written out here
      // and again in the editor's splitter, so the two implementations of one rule could drift
      // while the tested one sat unused (2026-08-09).
      node.sizes = resizeAt(node.sizes, index, deltaPixels / total, MIN_PANE / total);

      const cells = [...container.children].filter((child) => child.classList.contains("dock-cell")) as HTMLElement[];
      cells.forEach((cell, cellIndex) => {
        cell.style.flex = `${node.sizes[cellIndex] ?? 1} 1 0`;
      });
    };

    splitter.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }

      event.preventDefault();
      splitter.setPointerCapture(event.pointerId);
      beginLiveDrag();
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
        endLiveDrag(() => this.handlers.layoutChanged());
        this.persist();
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
        this.handlers.layoutChanged();
        this.persist();
      }
    });

    return splitter;
  }

  /* ------------------------------------------------------------------ host visibility */

  /** Tells the host which panes can actually be seen, one message per change. */
  private emitVisibility(): void {
    const visible = new Set<string>();
    for (const host of this.groupHosts) {
      if (host.group.active) {
        visible.add(host.group.active);
      }
    }

    for (const name of this.seats.keys()) {
      const now = visible.has(name);
      if (this.toldVisibility.get(name) !== now) {
        this.toldVisibility.set(name, now);
        this.handlers.visibilityChanged(name, now);
      }
    }
  }
}
