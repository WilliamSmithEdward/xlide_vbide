/*
 * One context menu at a time, drawn with the same dropdown styling as the menu bar so every menu
 * in the product is one thing.
 *
 * Items are curated per object class by the caller, which is the point: a context menu is a claim
 * about what makes sense for the thing that was clicked, and an option that does not apply to the
 * class is left out rather than shown disabled.
 */

export interface ContextMenuItem {
  /** Absent on a separator. */
  label?: string;
  /** Defaults to true. Disabled items are shown when the action exists but cannot run now. */
  enabled?: boolean;
  run?: () => void;
}

let openMenu: HTMLElement | null = null;
let highlighted = -1;
let closeListeners: (() => void) | null = null;

/** What to run when the menu goes, whichever way it goes. See showContextMenu. */
let dismissed: (() => void) | null = null;

/** Closes the open context menu, if any. */
export function closeContextMenu(): void {
  openMenu?.remove();
  openMenu = null;
  highlighted = -1;
  closeListeners?.();
  closeListeners = null;

  // Last, and cleared before it runs: the callback may open another menu, and it must not be
  // this one's dismissal that the next one inherits.
  const wasDismissed = dismissed;
  dismissed = null;
  wasDismissed?.();
}

/**
 * Shows a context menu at a screen position, replacing any menu already open.
 *
 * `onClosed` runs however the menu goes: an item chosen, Escape, a click elsewhere, or another
 * menu replacing it. A caller that marked something to show WHICH thing the menu is about needs
 * that mark taken back, and there is exactly one place that knows the menu is gone.
 */
export function showContextMenu(
  x: number,
  y: number,
  items: ContextMenuItem[],
  onClosed?: () => void,
): void {
  closeContextMenu();

  const rows = items.filter((item) => item.label !== undefined);
  if (rows.length === 0) {
    // Nothing to show, so nothing was ever marked: the caller's undo still has to run, or a
    // right-click that produced no menu would leave its highlight behind for good.
    onClosed?.();
    return;
  }

  dismissed = onClosed ?? null;

  const menu = document.createElement("div");
  menu.className = "menu-dropdown";
  menu.setAttribute("role", "menu");

  let first = true;
  for (const item of items) {
    if (item.label === undefined) {
      if (!first) {
        const divider = document.createElement("div");
        divider.className = "menu-separator";
        divider.setAttribute("role", "separator");
        menu.appendChild(divider);
      }
      continue;
    }

    first = false;
    const enabled = item.enabled !== false;

    const row = document.createElement("div");
    row.className = "menu-item" + (enabled ? "" : " disabled");
    row.dataset.i = String(rows.indexOf(item));
    row.setAttribute("role", "menuitem");
    if (!enabled) {
      row.setAttribute("aria-disabled", "true");
    }

    const caption = document.createElement("span");
    caption.className = "menu-caption";
    caption.textContent = item.label;
    row.appendChild(caption);

    row.addEventListener("pointerenter", () => setHighlight(menu, Number(row.dataset.i)));
    row.addEventListener("pointerup", () => {
      if (enabled) {
        closeContextMenu();
        item.run?.();
      }
    });

    menu.appendChild(row);
  }

  document.body.appendChild(menu);

  // Clamped after it has a size, so it never opens half off screen.
  menu.style.left = `${Math.max(4, Math.min(x, window.innerWidth - menu.offsetWidth - 4))}px`;
  menu.style.top = `${Math.max(4, Math.min(y, window.innerHeight - menu.offsetHeight - 4))}px`;

  openMenu = menu;

  const onPointerDown = (event: PointerEvent) => {
    if (!(event.target as HTMLElement).closest(".menu-dropdown")) {
      closeContextMenu();
    }
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (!openMenu) {
      return;
    }

    if (event.key === "Escape") {
      closeContextMenu();
    } else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      const delta = event.key === "ArrowDown" ? 1 : -1;
      const next = highlighted < 0
        ? delta > 0 ? 0 : rows.length - 1
        : (highlighted + delta + rows.length) % rows.length;
      setHighlight(openMenu, next);
    } else if (event.key === "Enter" || event.key === " ") {
      const item = rows[highlighted];
      if (item && item.enabled !== false) {
        closeContextMenu();
        item.run?.();
      }
    } else {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
  };

  document.addEventListener("pointerdown", onPointerDown, true);
  document.addEventListener("keydown", onKeyDown, true);
  closeListeners = () => {
    document.removeEventListener("pointerdown", onPointerDown, true);
    document.removeEventListener("keydown", onKeyDown, true);
  };
}

function setHighlight(menu: HTMLElement, i: number): void {
  highlighted = i;
  for (const row of menu.querySelectorAll<HTMLElement>(".menu-item")) {
    row.classList.toggle("highlight", row.dataset.i === String(i));
  }
}
