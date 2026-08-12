/*
 * The settings dialog: a small modal card over the editor, one row per choice, saved the
 * moment a choice changes. There is no OK button because there is nothing to defer - the host
 * writes each change through and echoes it back, and the dialog shows the echo.
 */

import { currentSettings, onSettingsApplied, type EditorSettings } from "./settings.js";
import { openModal } from "./modal.js";

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
    key: "syncEngine" as const,
    kind: "choice" as const,
    // Short enough to stay on ONE line beside its 'i'. The longer version wrapped, which put the
    // symbol next to half a heading and left it looking like it belonged to nothing.
    label: "Import and export",
    description: "Which one works out what to write. Both let you review it first.",
    /*
     * THE LONG ONE GOES BEHIND THE 'i'. At 623 characters it was four times the next longest
     * description in this list and made its row taller than the other six put together, so the
     * settings a developer came to change were pushed off the screen by an explanation of one they
     * probably will not (2026-08-09).
     *
     * AND IT WAS WRITTEN FROM THE INSIDE. "Which code works out the plan", "what counts as a
     * change", the engine as a thing the reader is assumed to know about: all true, none of it a
     * sentence anyone reading a settings dialog is asking for. What they want to know is whether
     * this matters to them and what happens if they pick wrong, and the honest answer to both is
     * "hardly, and nothing" - so that is what it says now, in that order, ending with permission
     * to leave it alone.
     */
    detail:
      "Both do the same job, and both let you check the changes before anything is written.\n\n"
      + "Shared with xlide for VS Code: this add-in and the VS Code extension behave identically, "
      + "so the same folder works the same way in both. Needs xlide's engine running.\n\n"
      + "Built into the add-in: works even if the engine stops.\n\n"
      + "Not sure? Leave it as it is.",
    choices: [
      { value: "xlide", label: "Shared with xlide for VS Code" },
      { value: "builtIn", label: "Built into the add-in" },
    ],
  },
];

/**
 * The panes the Panes menu shows and hides. Supplied by the shell, because pane visibility is
 * page-local arrangement rather than a setting the host persists - it lives with the dock
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
 * menu rather than a settings section (developer, 2026-08-06) - showing and hiding a pane
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

  // The listener is unhooked in `closed`, which the modal runs however the card goes - the
  // close button, Escape, or the backdrop.
  let stopListening = (): void => {};

  const { card, dismiss } = openModal({
    backdropId: "settings-backdrop",
    cardId: "settings-card",
    label: "Settings",
    closed: () => {
      stopListening();
      closed();
    },
  });

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

    // MORE THAN A LINE OF EXPLANATION GOES BEHIND AN 'i' BESIDE THE LABEL.
    //
    // On hover AND on focus, because a tooltip a keyboard cannot reach is a tooltip half the
    // people using this cannot read. It is a real button for the same reason: focusable, on the
    // tab order, and answering Enter and Space without any of that being written here.
    //
    // The text stays in the DOM either way and the control points at it with aria-describedby, so
    // a screen reader announces the explanation with the setting whether or not anything is
    // hovering. Hiding it visually is a visual decision, and it should not be a semantic one.
    let detailId: string | null = null;
    if ("detail" in option && typeof option.detail === "string") {
      detailId = `setting-${option.key}-detail`;

      const more = document.createElement("button");
      more.type = "button";
      more.className = "settings-more";

      // The circle is an inner span so the BUTTON can be the 24px target the pointer needs while
      // the ink stays the 16px this dialog's type size wants. Marked hidden from the accessibility
      // tree: a letter i read out beside the label is noise, and aria-label below is the name.
      const glyph = document.createElement("span");
      glyph.setAttribute("aria-hidden", "true");
      glyph.textContent = "i";
      more.appendChild(glyph);

      more.setAttribute("aria-label", `About ${option.label}`);
      more.setAttribute("aria-expanded", "false");
      more.setAttribute("aria-controls", detailId);

      const detail = document.createElement("div");
      detail.className = "settings-detail";
      detail.id = detailId;
      detail.setAttribute("role", "tooltip");
      detail.textContent = option.detail;

      const show = (open: boolean): void => {
        more.setAttribute("aria-expanded", String(open));
        detail.classList.remove("above");
        detail.classList.toggle("shown", open);

        // ABOVE THE 'i' WHEN THERE IS NO ROOM BELOW IT.
        //
        // This row is the last in the list, so its tooltip opens into whatever is under the
        // dialog, and the editor window on a real machine is not always tall: 640x409 on the
        // developer's own. Measured after showing rather than guessed from the row's position,
        // because the height depends on how the words wrap at this width.
        if (open && detail.getBoundingClientRect().bottom > window.innerHeight - 8) {
          detail.classList.add("above");
        }
      };

      more.addEventListener("pointerenter", () => show(true));
      more.addEventListener("focus", () => show(true));
      more.addEventListener("blur", () => show(false));
      more.addEventListener("click", () => show(more.getAttribute("aria-expanded") !== "true"));

      // Left open while the pointer is on the tooltip itself, so text long enough to want reading
      // twice can be. Leaving either one closes it.
      const leave = (event: PointerEvent): void => {
        const to = event.relatedTarget;
        if (!(to instanceof Node) || (!more.contains(to) && !detail.contains(to))) {
          if (document.activeElement !== more) { show(false); }
        }
      };

      more.addEventListener("pointerleave", leave);
      detail.addEventListener("pointerleave", leave);

      // Escape dismisses it without dismissing the dialog behind it, which is the order a reader
      // expects and the one WCAG asks for.
      more.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && more.getAttribute("aria-expanded") === "true") {
          event.preventDefault();
          event.stopPropagation();
          show(false);
        }
      });

      const beside = document.createElement("span");
      beside.className = "settings-label-row";
      label.replaceWith(beside);
      beside.append(label, more, detail);
    }

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

    // The explanation belongs to the CONTROL, whether or not anything is hovering: a screen
    // reader should read the setting and what it does in one breath, the way a sighted reader gets
    // both from the row.
    if (detailId !== null) {
      control.setAttribute("aria-describedby", detailId);
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
  stopListening = onSettingsApplied(refresh);

  close.addEventListener("click", dismiss);
  (card.querySelector("select, input") as HTMLElement | null)?.focus();
}
