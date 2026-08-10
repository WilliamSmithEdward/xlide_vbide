/*
 * The menu bar over the editor, drawn by the surface but backed by the host's own menus.
 *
 * Nothing here knows what any menu contains. The items are asked for from the host at the moment a
 * menu opens and rendered exactly as given, so the contents, the language they are written in, and
 * the menus that change while the editor runs (the Window list, enablement, checkmarks) are always
 * what the native menu would have shown. A chosen item is named back to the host by its position
 * chain, never by identifier, because identifiers repeat across the menus and positions do not.
 */

export interface MenuItem {
  /** The item's real position in the host's control collection, one-based. */
  index: number;
  /** Caption as the host holds it: '&' marks the accelerator letter, '&&' is a literal '&'. */
  caption: string;
  enabled: boolean;
  /** Draws a divider above this item. */
  separator: boolean;
  /** Opens a submenu rather than doing something. */
  popup: boolean;
  checked: boolean;
  shortcut?: string | null;
  /**
   * A codicon name to draw instead of the caption. Only the entries this product composes carry
   * one; everything mirrored from the editor is a caption. The caption is still sent when this is
   * set, and becomes the button's accessible name.
   */
  icon?: string | null;
}

export interface MenubarHandlers {
  /** Ask the host for the items of the menu at this path; [] is the bar itself. */
  request(path: number[]): void;
  /** Run the item at this path. */
  execute(path: number[]): void;
  /** Every menu closed, so focus belongs to the editor again. */
  closed(): void;
}

/** One open dropdown: the path it shows, what it hangs from, and what it holds. */
interface Level {
  path: number[];
  anchor: HTMLElement;
  dropdown: HTMLElement;
  items: MenuItem[];
  highlighted: number;
  /** Keyboard opened it, so the first item is highlighted the moment the items arrive. */
  wantsFirstHighlight: boolean;
}

/** How long the pointer can rest on a neighbouring item before the open submenu follows it. */
const SUBMENU_HOVER_DELAY = 180;

/** The accelerator letter of a caption, or null when it has none. */
export function accelOf(caption: string): string | null {
  for (let i = 0; i < caption.length - 1; i++) {
    if (caption[i] === "&") {
      if (caption[i + 1] === "&") {
        i++;
        continue;
      }
      return caption[i + 1] ?? null;
    }
  }
  return null;
}

/** Renders a caption into an element, underlining the accelerator letter. */
function renderCaption(target: HTMLElement, caption: string): void {
  let plain = "";
  for (let i = 0; i < caption.length; i++) {
    const ch = caption[i];
    if (ch === "&") {
      if (caption[i + 1] === "&") {
        plain += "&";
        i++;
        continue;
      }
      const accel = caption[i + 1];
      if (accel !== undefined) {
        if (plain) {
          target.appendChild(document.createTextNode(plain));
          plain = "";
        }
        const underline = document.createElement("span");
        underline.className = "menu-underline";
        underline.textContent = accel;
        target.appendChild(underline);
        i++;
      }
      continue;
    }
    plain += ch;
  }
  if (plain) {
    target.appendChild(document.createTextNode(plain));
  }
}

export class Menubar {
  private readonly root: HTMLElement;
  private readonly handlers: MenubarHandlers;

  private top: MenuItem[] = [];
  private levels: Level[] = [];
  private hoverTimer: ReturnType<typeof setTimeout> | undefined;
  private visible = true;

  constructor(root: HTMLElement, handlers: MenubarHandlers) {
    this.root = root;
    this.handlers = handlers;

    // Capture phase, because while a menu is open it owns the keyboard and the pointer the way a
    // native menu does, and the editor underneath must not also see what closes it.
    document.addEventListener("pointerdown", (event) => this.onPointerDownAnywhere(event), true);
    document.addEventListener("keydown", (event) => this.onKeyDown(event), true);
  }

  /** True while any menu is dropped down. */
  get isOpen(): boolean {
    return this.levels.length > 0;
  }

  /** Asks the host for the top-level menus. */
  refresh(): void {
    this.handlers.request([]);
  }

  /** Shows or withdraws the whole bar. Withdrawn, it also lets go of the keyboard. */
  setVisible(visible: boolean): void {
    if (this.visible === visible) {
      return;
    }
    this.visible = visible;
    this.root.hidden = !visible;
    if (!visible) {
      this.closeAll(false);
    }
  }

  /** Takes one menu's items from the host: the bar for an empty path, a dropdown otherwise. */
  setItems(path: number[], items: MenuItem[]): void {
    if (path.length === 0) {
      this.top = items;
      this.renderBar();
      return;
    }

    // A reply for a menu that is no longer open is stale by definition.
    const level = this.levels.find((l) => samePath(l.path, path));
    if (!level) {
      return;
    }

    level.items = items;
    this.renderDropdown(level);
  }

  private renderBar(): void {
    this.root.replaceChildren();

    for (const item of this.top) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "menu-top";
      button.dataset.index = String(item.index);
      button.setAttribute("role", "menuitem");
      button.setAttribute("aria-haspopup", "true");

      if (item.icon) {
        // The glyph carries no meaning to anything that cannot see it, so the caption becomes the
        // name and the icon is marked decorative - otherwise this is a button called "button".
        button.classList.add("menu-top-icon");
        button.setAttribute("aria-label", item.caption);
        button.title = item.caption;
        const glyph = document.createElement("span");
        glyph.className = `codicon codicon-${item.icon}`;
        glyph.setAttribute("aria-hidden", "true");
        button.replaceChildren(glyph);
      } else {
        renderCaption(button, item.caption);
      }

      // pointerdown rather than click, so the menu opens on press the way native menus do, and
      // the default is prevented so focus stays wherever the developer was working.
      button.addEventListener("pointerdown", (event) => {
        event.preventDefault();
        if (this.levels[0]?.anchor === button) {
          this.closeAll(true);
        } else {
          this.openTop(item, button, false);
        }
      });

      // Once a menu is open, sliding along the bar moves between menus without another press.
      button.addEventListener("pointerenter", () => {
        const first = this.levels[0];
        if (first && first.anchor !== button) {
          this.openTop(item, button, false);
        }
      });

      this.root.appendChild(button);
    }
  }

  private openTop(item: MenuItem, anchor: HTMLElement, fromKeyboard: boolean): void {
    this.clearFrom(0);
    anchor.classList.add("open");
    this.pushLevel([item.index], anchor, fromKeyboard);
  }

  private openTopAt(topIndex: number, fromKeyboard: boolean): void {
    const item = this.top[topIndex];
    const anchor = this.root.children[topIndex] as HTMLElement | undefined;
    if (item && anchor) {
      this.openTop(item, anchor, fromKeyboard);
    }
  }

  private pushLevel(path: number[], anchor: HTMLElement, fromKeyboard: boolean): void {
    const dropdown = document.createElement("div");
    dropdown.className = "menu-dropdown";
    dropdown.setAttribute("role", "menu");
    document.body.appendChild(dropdown);

    const level: Level = {
      path,
      anchor,
      dropdown,
      items: [],
      highlighted: -1,
      wantsFirstHighlight: fromKeyboard,
    };

    this.levels.push(level);
    this.position(level);
    this.handlers.request(path);
  }

  private position(level: Level): void {
    const rect = level.anchor.getBoundingClientRect();
    const depth = this.levels.indexOf(level);

    if (depth <= 0) {
      level.dropdown.style.left = `${rect.left}px`;
      level.dropdown.style.top = `${rect.bottom}px`;
    } else {
      // Hung off the item's right edge, pulled in slightly so the two read as one surface.
      level.dropdown.style.left = `${rect.right - 2}px`;
      level.dropdown.style.top = `${rect.top - 4}px`;
    }

    // Clamped after it has a size, which is only known once it is in the document.
    const box = level.dropdown.getBoundingClientRect();
    if (box.right > window.innerWidth) {
      const flipped = depth <= 0 ? window.innerWidth - box.width - 4 : rect.left - box.width + 2;
      level.dropdown.style.left = `${Math.max(4, flipped)}px`;
    }
    if (box.bottom > window.innerHeight) {
      level.dropdown.style.top = `${Math.max(4, window.innerHeight - box.height - 4)}px`;
    }
  }

  private renderDropdown(level: Level): void {
    level.dropdown.replaceChildren();

    if (level.items.length === 0) {
      const empty = document.createElement("div");
      empty.className = "menu-empty";
      empty.textContent = "(empty)";
      level.dropdown.appendChild(empty);
      this.position(level);
      return;
    }

    level.items.forEach((item, i) => {
      if (item.separator && i > 0) {
        const divider = document.createElement("div");
        divider.className = "menu-separator";
        divider.setAttribute("role", "separator");
        level.dropdown.appendChild(divider);
      }

      const row = document.createElement("div");
      row.className = "menu-item" + (item.enabled ? "" : " disabled");
      row.dataset.i = String(i);
      row.setAttribute("role", "menuitem");
      if (!item.enabled) {
        row.setAttribute("aria-disabled", "true");
      }
      if (item.popup) {
        row.setAttribute("aria-haspopup", "true");
      }

      const check = document.createElement("span");
      check.className = "menu-check codicon codicon-check";
      check.setAttribute("aria-hidden", String(!item.checked));
      if (!item.checked) {
        check.classList.add("blank");
      }

      const caption = document.createElement("span");
      caption.className = "menu-caption";
      renderCaption(caption, item.caption);

      row.append(check, caption);

      if (item.shortcut) {
        const shortcut = document.createElement("span");
        shortcut.className = "menu-shortcut";
        shortcut.textContent = item.shortcut;
        row.appendChild(shortcut);
      }

      if (item.popup) {
        const chevron = document.createElement("span");
        chevron.className = "menu-sub codicon codicon-chevron-right";
        chevron.setAttribute("aria-hidden", "true");
        row.appendChild(chevron);
      }

      row.addEventListener("pointerenter", () => this.onItemHover(level, i, row));
      row.addEventListener("pointerup", () => this.onItemActivate(level, i, row));

      level.dropdown.appendChild(row);
    });

    if (level.wantsFirstHighlight) {
      level.wantsFirstHighlight = false;
      this.setHighlight(level, 0);
    }

    this.position(level);
  }

  private onItemHover(level: Level, i: number, row: HTMLElement): void {
    const item = level.items[i];
    if (!item) {
      return;
    }

    this.setHighlight(level, i);

    if (this.hoverTimer !== undefined) {
      clearTimeout(this.hoverTimer);
    }

    // Both opening a submenu and closing a neighbour's wait the same moment, so the pointer can
    // cut the corner across other items on its way into an open submenu without losing it.
    const depth = this.levels.indexOf(level);

    this.hoverTimer = setTimeout(() => {
      if (item.popup && item.enabled) {
        this.openSub(level, i, row, false);
      } else {
        this.clearFrom(depth + 1);
      }
    }, SUBMENU_HOVER_DELAY);
  }

  private onItemActivate(level: Level, i: number, row: HTMLElement): void {
    const item = level.items[i];
    if (!item || !item.enabled) {
      return;
    }

    if (item.popup) {
      if (this.hoverTimer !== undefined) {
        clearTimeout(this.hoverTimer);
      }
      this.openSub(level, i, row, false);
      return;
    }

    this.handlers.execute([...level.path, item.index]);
    this.closeAll(true);
  }

  private openSub(level: Level, i: number, row: HTMLElement, fromKeyboard: boolean): void {
    const depth = this.levels.indexOf(level);
    if (depth < 0) {
      return;
    }

    const item = level.items[i];
    if (!item) {
      return;
    }

    const childPath = [...level.path, item.index];

    const existing = this.levels[depth + 1];
    if (existing && samePath(existing.path, childPath)) {
      return;
    }

    this.clearFrom(depth + 1);
    this.pushLevel(childPath, row, fromKeyboard);
  }

  private setHighlight(level: Level, i: number): void {
    level.highlighted = i;
    for (const row of level.dropdown.querySelectorAll<HTMLElement>(".menu-item")) {
      row.classList.toggle("highlight", row.dataset.i === String(i));
    }
  }

  private moveHighlight(level: Level, delta: number): void {
    if (level.items.length === 0) {
      return;
    }
    const from = level.highlighted;
    const next = from < 0
      ? delta > 0 ? 0 : level.items.length - 1
      : (from + delta + level.items.length) % level.items.length;
    this.setHighlight(level, next);
  }

  private rowOf(level: Level, i: number): HTMLElement | null {
    return level.dropdown.querySelector<HTMLElement>(`.menu-item[data-i="${i}"]`);
  }

  private clearFrom(depth: number): void {
    while (this.levels.length > depth) {
      const level = this.levels.pop();
      level?.dropdown.remove();
    }
    if (depth === 0) {
      for (const button of this.root.querySelectorAll(".menu-top.open")) {
        button.classList.remove("open");
      }
    }
  }

  private closeAll(notify: boolean): void {
    if (this.hoverTimer !== undefined) {
      clearTimeout(this.hoverTimer);
      this.hoverTimer = undefined;
    }
    const wasOpen = this.levels.length > 0;
    this.clearFrom(0);
    if (notify && wasOpen) {
      this.handlers.closed();
    }
  }

  private onPointerDownAnywhere(event: PointerEvent): void {
    if (this.levels.length === 0) {
      return;
    }
    const target = event.target as HTMLElement | null;
    if (target?.closest(".menu-dropdown") || target?.closest("#menubar")) {
      return;
    }
    this.closeAll(true);
  }

  private onKeyDown(event: KeyboardEvent): void {
    if (!this.visible || this.top.length === 0) {
      return;
    }

    if (this.levels.length === 0) {
      this.onKeyWhileClosed(event);
      return;
    }

    // While open, the menu owns the keyboard entirely, the way a native menu does. Whatever is
    // not understood is swallowed rather than typed into the editor behind the open menu.
    this.onKeyWhileOpen(event);
    event.preventDefault();
    event.stopPropagation();
  }

  private onKeyWhileClosed(event: KeyboardEvent): void {
    if (event.key === "F10" && !event.altKey && !event.ctrlKey && !event.shiftKey) {
      event.preventDefault();
      event.stopPropagation();
      this.openTopAt(0, true);
      return;
    }

    if (event.altKey && !event.ctrlKey && event.key.length === 1) {
      const wanted = event.key.toLowerCase();
      const topIndex = this.top.findIndex((t) => accelOf(t.caption)?.toLowerCase() === wanted);
      if (topIndex >= 0) {
        event.preventDefault();
        event.stopPropagation();
        this.openTopAt(topIndex, true);
      }
    }
  }

  private onKeyWhileOpen(event: KeyboardEvent): void {
    const level = this.levels[this.levels.length - 1];
    if (!level) {
      return;
    }

    // Alt with a letter reaches across to another top-level menu even while one is open.
    if (event.altKey && !event.ctrlKey && event.key.length === 1) {
      const wanted = event.key.toLowerCase();
      const topIndex = this.top.findIndex((t) => accelOf(t.caption)?.toLowerCase() === wanted);
      if (topIndex >= 0) {
        this.openTopAt(topIndex, true);
      }
      return;
    }

    switch (event.key) {
      case "Escape":
        if (this.levels.length > 1) {
          this.clearFrom(this.levels.length - 1);
        } else {
          this.closeAll(true);
        }
        return;

      case "ArrowDown":
        this.moveHighlight(level, 1);
        return;

      case "ArrowUp":
        this.moveHighlight(level, -1);
        return;

      case "ArrowRight": {
        const item = level.items[level.highlighted];
        if (item?.popup && item.enabled) {
          const row = this.rowOf(level, level.highlighted);
          if (row) {
            this.openSub(level, level.highlighted, row, true);
          }
        } else {
          this.openAdjacentTop(1);
        }
        return;
      }

      case "ArrowLeft":
        if (this.levels.length > 1) {
          this.clearFrom(this.levels.length - 1);
        } else {
          this.openAdjacentTop(-1);
        }
        return;

      case "Enter":
      case " ": {
        const item = level.items[level.highlighted];
        if (!item || !item.enabled) {
          return;
        }
        if (item.popup) {
          const row = this.rowOf(level, level.highlighted);
          if (row) {
            this.openSub(level, level.highlighted, row, true);
          }
          return;
        }
        this.handlers.execute([...level.path, item.index]);
        this.closeAll(true);
        return;
      }

      default: {
        if (event.key.length !== 1 || event.ctrlKey || event.metaKey) {
          return;
        }
        const wanted = event.key.toLowerCase();
        const i = level.items.findIndex((t) => accelOf(t.caption)?.toLowerCase() === wanted);
        if (i < 0) {
          return;
        }
        const row = this.rowOf(level, i);
        if (row) {
          this.onItemActivate(level, i, row);
        }
        return;
      }
    }
  }

  private openAdjacentTop(delta: number): void {
    const current = this.top.findIndex((t) => t.index === this.levels[0]?.path[0]);
    if (current < 0) {
      return;
    }
    const next = (current + delta + this.top.length) % this.top.length;
    this.openTopAt(next, true);
  }
}

function samePath(a: number[], b: number[]): boolean {
  return a.length === b.length && a.every((value, i) => value === b[i]);
}
