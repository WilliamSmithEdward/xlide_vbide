/*
 * Where a symbol is used, as a list xlide draws itself.
 *
 * The editor has a references window of its own, and it cannot be used here: it renders each
 * result by resolving its address to an editor MODEL, and this surface only holds models for
 * modules with a tab open. So the one reference a developer most needs to be shown - the use in
 * a module they have never opened, the one that breaks when they change something - is the one
 * that window structurally cannot draw.
 *
 * This renders text instead. The host sends the line with each location, so a module with no tab
 * is listed exactly like one with a tab, and clicking either takes you there.
 */

import type { HostLocation } from "./bridge.js";
import { openModal } from "./modal.js";

/** What the dialog needs of the world: where to go when a row is picked. */
export interface ReferencesHandlers {
  navigate(module: string, line: number, column: number, workbook: string | null): void;
}

let open: { dismiss: () => void } | null = null;

/**
 * Shows the references to one symbol, grouped by the module they are in.
 *
 * Replaces whatever is already showing rather than stacking: asking twice is asking again, and
 * two lists of results with no way to tell which is current is worse than one.
 */
export function openReferencesDialog(
  symbol: string,
  locations: readonly HostLocation[],
  handlers: ReferencesHandlers,
): void {
  closeReferencesDialog();

  const { card: dialog, dismiss } = openModal({
    backdropId: "references-backdrop",
    cardId: "references-card",
    label: `References to ${symbol}`,
    closed: () => {
      open = null;
    },
  });
  open = { dismiss };

  const header = document.createElement("div");
  header.id = "references-head";

  const title = document.createElement("h2");
  title.id = "references-title";
  const count = locations.length;
  title.textContent = count === 0
    ? `No references to ${symbol}`
    : `${count} reference${count === 1 ? "" : "s"} to ${symbol}`;

  const close = document.createElement("button");
  close.type = "button";
  close.id = "references-close";
  close.title = "Close";
  close.setAttribute("aria-label", "Close");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';
  close.addEventListener("click", () => closeReferencesDialog());

  header.append(title, close);

  const body = document.createElement("div");
  body.id = "references-body";

  if (count === 0) {
    const empty = document.createElement("p");
    empty.className = "references-empty";
    empty.textContent = "Nothing in this workbook uses it.";
    body.appendChild(empty);
  }

  // Grouped by module, in the order the modules first appear, so the module being worked in -
  // which is where the answer came from - is usually the group at the top.
  const byModule = new Map<string, HostLocation[]>();
  for (const location of locations) {
    const held = byModule.get(location.module);
    if (held) {
      held.push(location);
    } else {
      byModule.set(location.module, [location]);
    }
  }

  for (const [module, inModule] of byModule) {
    const group = document.createElement("div");
    group.className = "references-group";
    group.dataset.module = module;

    const heading = document.createElement("div");
    heading.className = "references-module";
    heading.textContent = `${module}  (${inModule.length})`;
    group.appendChild(heading);

    for (const location of inModule) {
      const row = document.createElement("button");
      row.type = "button";
      row.className = "references-row";
      row.dataset.module = location.module;
      row.dataset.line = String(location.line);

      const where = document.createElement("span");
      where.className = "references-line";
      where.textContent = String(location.line);

      const text = document.createElement("span");
      text.className = "references-preview";
      // Empty rather than absent when the host had no line to send: a row with a number and no
      // text still navigates, and dropping it would under-report the count in the title.
      text.textContent = location.preview ?? "";

      row.append(where, text);
      row.addEventListener("click", () => {
        closeReferencesDialog();
        handlers.navigate(
          location.module,
          location.line,
          location.column,
          location.workbook ?? null);
      });

      group.appendChild(row);
    }

    body.appendChild(group);
  }

  dialog.append(header, body);

  // Focus the first row so the keyboard can walk the list immediately, and Escape has somewhere
  // to be heard from.
  (body.querySelector<HTMLElement>(".references-row") ?? close).focus();
}

export function closeReferencesDialog(): void {
  open?.dismiss();
}
