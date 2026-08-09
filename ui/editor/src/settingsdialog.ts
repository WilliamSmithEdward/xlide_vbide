/*
 * The settings dialog: a small modal card over the editor, one row per choice, saved the
 * moment a choice changes. There is no OK button because there is nothing to defer — the host
 * writes each change through and echoes it back, and the dialog shows the echo.
 */

import { currentSettings, onSettingsApplied, type EditorSettings } from "./settings.js";

/** The rows, spelled the way the companion editor's own settings describe themselves. */
const OPTIONS = [
  {
    key: "blockLayout" as const,
    kind: "choice" as const,
    label: "Smart Enter block layout",
    description:
      "'Comfy' opens spacer lines around the editable body; 'compact' places the body directly above the closer.",
    choices: [
      { value: "comfy", label: "Comfy" },
      { value: "compact", label: "Compact" },
    ],
  },
  {
    key: "continueCommentOnNewline" as const,
    kind: "toggle" as const,
    label: "Continue comments on Enter",
    description:
      "When the line above is a comment, a new line begins with an apostrophe to continue it.",
  },
  {
    key: "mirrorCommentSpacing" as const,
    kind: "toggle" as const,
    label: "Mirror comment spacing",
    description:
      "A continued comment also repeats the spaces after the apostrophe, so the text lines up.",
  },
  {
    key: "treeFollowsEditor" as const,
    kind: "toggle" as const,
    label: "Explorer follows the editor",
    description:
      "The module you are working on unfolds its procedures in the tree, and everything folds "
      + "away when the last tab closes. Off leaves the tree exactly as you left it.",
  },
  // There is no "indent with tabs" here, and there cannot be a working one: VBA's code store
  // will not hold a tab, and expands any it is handed. Indentation is this many spaces, and
  // Backspace in a line's leading whitespace takes back a whole level of them.
  {
    key: "formatIndentSize" as const,
    kind: "number" as const,
    label: "Indent size",
    description: "One indent level, in spaces. Used by typing, by smart Enter, by Backspace, and by formatting.",
    min: 1,
    max: 8,
  },
  {
    key: "formatCanonicalKeywords" as const,
    kind: "toggle" as const,
    label: "Format: canonical keywords",
    description: "Formatting respells keywords in their canonical case, the way the language spells them.",
  },
  {
    key: "syncEngine" as const,
    kind: "choice" as const,
    label: "Import and export: which one decides",
    description:
      "Both choices import and export the same files, and either way you review the changes before "
      + "anything is written. What differs is which code works out the plan. 'Shared with xlide for "
      + "VS Code' uses the very same code the extension does, so both products agree exactly about "
      + "file names, what kind of module a file holds, and what counts as a change - and it needs "
      + "xlide's engine to be running. 'Built into the add-in' works it out here instead, so import "
      + "and export keep working even if the engine does not. If the engine is unavailable, the "
      + "add-in quietly uses its own and says so in the log.",
    choices: [
      { value: "xlide", label: "Shared with xlide for VS Code" },
      { value: "builtIn", label: "Built into the add-in" },
    ],
  },
];

/**
 * The panes the Panes menu shows and hides. Supplied by the shell, because pane visibility is
 * page-local arrangement rather than a setting the host persists — it lives with the dock
 * layout, and this is the route back for a pane that was closed with its X.
 */
export interface PaneVisibilityControl {
  list(): { name: string; title: string; open: boolean; permanent: boolean }[];
  setOpen(name: string, open: boolean): void;
  /**
   * Moves a pane to a dock side, through the method a real drop calls.
   *
   * Here rather than on a docking-specific type because this is already the object handed to
   * everything that drives panes from outside the docks themselves.
   */
  moveTo(name: string, side: "left" | "right" | "top" | "bottom"): boolean;
}

/**
 * The Panes menu: one checkable row per pane, dropped under its toolbar button. Its own
 * menu rather than a settings section (developer, 2026-08-06) — showing and hiding a pane
 * is a thing done while working, not a preference visited once.
 */
export function openPanesMenu(panes: PaneVisibilityControl, anchor: HTMLElement): void {
  const existing = document.getElementById("panes-menu");
  if (existing) {
    existing.remove();
    return;
  }

  const menu = document.createElement("div");
  menu.id = "panes-menu";
  // The surface's own dropdown chrome, so this menu is the menus the developer already knows.
  menu.className = "menu-dropdown";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "Panes");

  const dismiss = (): void => {
    menu.remove();
    document.removeEventListener("pointerdown", onPointerDown, true);
    document.removeEventListener("keydown", onKey, true);
  };

  const onPointerDown = (event: PointerEvent): void => {
    if (!menu.contains(event.target as Node) && !anchor.contains(event.target as Node)) {
      dismiss();
    }
  };

  const onKey = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  };

  const draw = (): void => {
    menu.replaceChildren();

    for (const pane of panes.list()) {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "menu-item panes-menu-item" + (pane.permanent ? " disabled" : "");
      item.setAttribute("role", "menuitemcheckbox");
      item.setAttribute("aria-checked", String(pane.open));
      item.disabled = pane.permanent;
      item.dataset.pane = pane.name;

      const tick = document.createElement("span");
      tick.className = "panes-menu-tick codicon" + (pane.open ? " codicon-check" : "");
      tick.setAttribute("aria-hidden", "true");

      const label = document.createElement("span");
      label.className = "menu-caption";
      label.textContent = pane.permanent ? `${pane.title} (always shown)` : pane.title;

      item.append(tick, label);
      item.addEventListener("click", () => {
        panes.setOpen(pane.name, !pane.open);
        draw();
      });

      menu.appendChild(item);
    }
  };

  draw();
  document.body.appendChild(menu);

  // Under its button, pulled back on-screen when the button sits near the right edge.
  const box = anchor.getBoundingClientRect();
  const width = menu.offsetWidth;
  menu.style.left = `${Math.max(4, Math.min(box.left, window.innerWidth - width - 4))}px`;
  menu.style.top = `${box.bottom + 2}px`;

  document.addEventListener("pointerdown", onPointerDown, true);
  document.addEventListener("keydown", onKey, true);
  menu.querySelector<HTMLButtonElement>("button:not([disabled])")?.focus();
}

/**
 * Opens the dialog. One at a time: opening while open focuses the one that exists. Every
 * change posts through `update`; the card re-reads its state whenever the host echoes.
 */
export function openSettingsDialog(
  update: (settings: EditorSettings) => void,
  closed: () => void,
): void {
  if (document.getElementById("settings-backdrop")) {
    (document.querySelector("#settings-card select, #settings-card input") as HTMLElement | null)?.focus();
    return;
  }

  const backdrop = document.createElement("div");
  backdrop.id = "settings-backdrop";

  const card = document.createElement("div");
  card.id = "settings-card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-label", "Settings");

  const head = document.createElement("div");
  head.id = "settings-head";

  const title = document.createElement("span");
  title.textContent = "Settings";

  const close = document.createElement("button");
  close.type = "button";
  close.id = "settings-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close settings");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';

  head.append(title, close);
  card.appendChild(head);

  const refreshers: (() => void)[] = [];

  for (const option of OPTIONS) {
    const row = document.createElement("div");
    row.className = "settings-row";

    const text = document.createElement("div");
    text.className = "settings-text";

    const label = document.createElement("label");
    label.className = "settings-label";
    label.textContent = option.label;
    label.htmlFor = `setting-${option.key}`;

    const description = document.createElement("div");
    description.className = "settings-description";
    description.textContent = option.description;

    text.append(label, description);

    let control: HTMLElement;
    if (option.kind === "choice") {
      const select = document.createElement("select");
      select.id = `setting-${option.key}`;
      for (const choice of option.choices) {
        const entry = document.createElement("option");
        entry.value = choice.value;
        entry.textContent = choice.label;
        select.appendChild(entry);
      }

      // BY THE ROW'S OWN KEY, both ways.
      //
      // This wrote to blockLayout and read from blockLayout whatever row it was drawing, which was
      // invisible while blockLayout was the only choice in the list. The second one arrived and
      // showed up empty (its value had been set to "comfy", which is not one of its options) and
      // would have rewritten the block layout when changed (2026-08-09).
      select.addEventListener("change", () => {
        update({ ...currentSettings(), [option.key]: select.value });
      });

      refreshers.push(() => {
        select.value = String(currentSettings()[option.key]);
      });
      control = select;
    } else if (option.kind === "number") {
      const field = document.createElement("input");
      field.type = "number";
      field.id = `setting-${option.key}`;
      field.min = String(option.min);
      field.max = String(option.max);
      field.step = "1";

      field.addEventListener("change", () => {
        const asked = Math.round(Number(field.value));
        const legal = Math.min(option.max, Math.max(option.min, Number.isFinite(asked) ? asked : option.min));
        update({ ...currentSettings(), [option.key]: legal });
      });

      refreshers.push(() => {
        field.value = String(currentSettings()[option.key]);
      });
      control = field;
    } else {
      const toggle = document.createElement("input");
      toggle.type = "checkbox";
      toggle.id = `setting-${option.key}`;

      toggle.addEventListener("change", () => {
        update({ ...currentSettings(), [option.key]: toggle.checked });
      });

      refreshers.push(() => {
        toggle.checked = currentSettings()[option.key];
      });
      control = toggle;
    }

    row.append(text, control);
    card.appendChild(row);
  }

  const foot = document.createElement("div");
  foot.id = "settings-foot";
  foot.textContent = "Changes are saved as you make them.";
  card.appendChild(foot);

  const refresh = (): void => refreshers.forEach((apply) => apply());
  refresh();
  const stopListening = onSettingsApplied(refresh);

  const dismiss = (): void => {
    stopListening();
    document.removeEventListener("keydown", onKey, true);
    backdrop.remove();
    closed();
  };

  const onKey = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  };

  close.addEventListener("click", dismiss);
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) {
      dismiss();
    }
  });
  document.addEventListener("keydown", onKey, true);

  backdrop.appendChild(card);
  document.body.appendChild(backdrop);
  (card.querySelector("select, input") as HTMLElement | null)?.focus();
}
