/*
 * The drag furniture shared by both drags - tool panes between sections, and code editor
 * tabs between editor groups. One implementation, so a tab and a pane behave identically
 * under the hand: the page dims to say a drag is a mode, a compass of five zones appears
 * over whatever region the pointer is on, the zone under the pointer lights, and a preview
 * shows the shape a release would produce.
 *
 * The compass IS the target (developer, 2026-08-06). The pointer must come to the zone it
 * means, rather than the code guessing an intent from where the pointer happens to be: over
 * a wide short region, "near the left edge" and "just left of centre" are the same few
 * pixels apart, and a guess there is a coin toss the developer has to undo.
 *
 * The petals are hit-tested by their own geometry rather than by pointer events, because
 * the dragged element holds the pointer capture and nothing else receives any.
 */

/**
 * How far above and below a strip's own edges a drag still counts as over it, for editor
 * tabs and pane tabs alike.
 *
 * Reordering is a horizontal gesture, and a hand making one drifts vertically: with a
 * strip's exact rectangle as the target, a couple of pixels of drift mid-drag dropped the
 * pointer into the compass and the reorder turned into a split offer (the developer,
 * 2026-08-12). The band roughly quadruples a strip's catch height - it began at 24 and was
 * doubled the same day, by feel. It exists only inside the drags' landing tests, so it
 * costs nothing outside a drag; strips are tested before bodies, so inside the band the
 * insertion wins over the compass zones the band overlaps - that trade, a slice of the
 * neighbouring split zones for a reorder that survives drift, is the point.
 */
export const STRIP_DRAG_REACH = 48;

/** The five places a drop can land on a region: its tabs, or one of its four edges. */
export type DropZone = "center" | "left" | "right" | "top" | "bottom";

export const ALL_ZONES: DropZone[] = ["center", "left", "right", "top", "bottom"];
export const EDGE_ZONES: DropZone[] = ["left", "right", "top", "bottom"];

/** The rectangle a zone would occupy inside a region: the half it splits, or the whole. */
export function zoneRect(bounds: DOMRect, zone: DropZone): DOMRect {
  const half = { w: bounds.width / 2, h: bounds.height / 2 };
  switch (zone) {
    case "left": return new DOMRect(bounds.x, bounds.y, half.w, bounds.height);
    case "right": return new DOMRect(bounds.x + half.w, bounds.y, half.w, bounds.height);
    case "top": return new DOMRect(bounds.x, bounds.y, bounds.width, half.h);
    case "bottom": return new DOMRect(bounds.x, bounds.y + half.h, bounds.width, half.h);
    default: return bounds;
  }
}

export class DragCompass {
  private readonly dim = document.createElement("div");
  private readonly compass = document.createElement("div");
  private readonly petals = new Map<DropZone, HTMLElement>();
  private overlay: HTMLElement | null = null;

  /** Called when something outside the pointer stream ends the drag. Set by the owner. */
  private abandon: (() => void) | null = null;

  constructor() {
    this.dim.className = "drag-dim";
    this.compass.className = "drop-compass";
    this.compass.hidden = true;

    for (const zone of ALL_ZONES) {
      const petal = document.createElement("div");
      petal.className = `drop-petal drop-petal-${zone}`;
      this.petals.set(zone, petal);
      this.compass.appendChild(petal);
    }
  }

  /**
   * The drag became real: dim the page and stand the compass by.
   *
   * `abandon` is called when the drag ends without a release the pointer stream could
   * report - the window losing focus (alt-tab, or the host stealing it, which this host
   * does freely), the document being hidden, or Escape. Without it the dim and the compass
   * outlived the drag and the surface was left looking permanently mid-gesture
   * (developer, 2026-08-06).
   */
  begin(abandon: () => void): void {
    document.body.appendChild(this.dim);
    document.body.appendChild(this.compass);

    // The cursor says the mode too. Without this it stays whatever it was over -- an I-beam over
    // the editor, an arrow over a pane -- which reads as ordinary pointing while the page is dim
    // and a compass is following the pointer around.
    document.documentElement.classList.add("drag-live");

    this.abandon = abandon;
    window.addEventListener("blur", this.onAbandon);
    document.addEventListener("visibilitychange", this.onVisibility);
    document.addEventListener("keydown", this.onKey, true);
  }

  private readonly onAbandon = (): void => {
    this.abandon?.();
  };

  private readonly onVisibility = (): void => {
    if (document.hidden) {
      this.abandon?.();
    }
  };

  private readonly onKey = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      this.abandon?.();
    }
  };

  /**
   * The pointer is over a region: places the compass on it, shows only the zones that region
   * allows, and answers the zone under the pointer - null when the pointer is on the region
   * but not on any petal, which means a release would drop nothing.
   */
  over(bounds: DOMRect, x: number, y: number, allowed: DropZone[]): DropZone | null {
    this.compass.hidden = false;
    this.compass.style.left = `${bounds.x + bounds.width / 2}px`;
    this.compass.style.top = `${bounds.y + bounds.height / 2}px`;

    let found: DropZone | null = null;
    for (const [zone, petal] of this.petals) {
      petal.hidden = !allowed.includes(zone);
      if (petal.hidden) {
        continue;
      }

      const box = petal.getBoundingClientRect();
      if (x >= box.left && x <= box.right && y >= box.top && y <= box.bottom) {
        found = zone;
      }
    }

    for (const [zone, petal] of this.petals) {
      petal.classList.toggle("lit", zone === found);
    }

    return found;
  }

  /**
   * Shows (or clears) the shape a release would produce.
   *
   * The kind says what KIND of change it is, because the two read differently and should
   * look different (developer, 2026-08-06): "join" lands the pane in something that is
   * already there, and shows that thing's own bounds; "new" carves fresh space out of what
   * the preview covers, and wears a dashed edge to say the space does not exist yet.
   */
  preview(rect: DOMRect | null, kind: "join" | "new" = "join"): void {
    if (!rect) {
      this.overlay?.remove();
      this.overlay = null;
      return;
    }

    if (!this.overlay) {
      this.overlay = document.createElement("div");
      document.body.appendChild(this.overlay);
    }

    this.overlay.className = `drop-overlay drop-overlay-${kind}`;
    this.overlay.style.left = `${rect.x}px`;
    this.overlay.style.top = `${rect.y}px`;
    this.overlay.style.width = `${rect.width}px`;
    this.overlay.style.height = `${rect.height}px`;
  }

  /** The pointer left every region: no compass, no preview, nothing to drop on. */
  clear(): void {
    this.compass.hidden = true;
    this.preview(null);
  }

  /** The drag is over, however it ended. Safe to call twice. */
  end(): void {
    this.abandon = null;
    window.removeEventListener("blur", this.onAbandon);
    document.removeEventListener("visibilitychange", this.onVisibility);
    document.removeEventListener("keydown", this.onKey, true);

    document.documentElement.classList.remove("drag-live");
    this.dim.remove();
    this.compass.remove();
    this.preview(null);
  }
}
