/*
 * The Tab Order dialog: which control the Tab key reaches next, and a way to change it.
 *
 * The native editor has one (View > Tab Order, menu 469) and this product suppresses it, for the
 * reason every suppression here has: it would act on the native designer's own selection rather
 * than on ours. So the gesture had to come back somewhere, and it comes back here - reached from
 * the canvas's context menu, listing the controls of ONE container in tab order, with Move Up and
 * Move Down.
 *
 * Why a dialog rather than a row in the Properties panel: the panel already carries TabIndex and
 * can already write it, and that is exactly the part nobody can use. A tab order is a sequence,
 * and a sequence is edited by looking at all of it at once - typing 4 into one control's box while
 * MSForms silently renumbers the other eleven is a puzzle, not an edit.
 *
 * The write is a plain TabIndex set through the host's own SetControlProperty, because MSForms
 * does the renumbering itself: giving a control the index of the one above it pushes that one
 * down. The dialog does not compute a new order, it asks for one move.
 */

import { openModal } from "./modal.js";

export interface TabOrderControl {
  name: string;
  type: string;
  tabIndex: number;
}

export interface TabOrderOptions {
  /** The container's own name, "" for the form itself - shown in the card's heading. */
  container: string;

  /** Its children, in whatever order the projection had them. */
  controls: TabOrderControl[];

  /** Writes one control's TabIndex; MSForms renumbers the rest of the container. */
  setIndex: (control: string, index: number) => void;

  /** Runs when the dialog goes, so the caller can drop its handle. */
  closed?: () => void;
}

/** The dialog that is open, for the debug surface and for the caller's own dismissal. */
export interface OpenTabOrder {
  /** The names as the list shows them, top to bottom. */
  order: () => string[];

  /** Moves the named control one place up or down, exactly as the buttons do. */
  move: (name: string, by: -1 | 1) => boolean;

  dismiss: () => void;
}

let standing: OpenTabOrder | null = null;

/** The tab-order dialog that is open, if one is. */
export function openTabOrderDialog(): OpenTabOrder | null {
  return standing;
}

export function showTabOrder(options: TabOrderOptions): OpenTabOrder {
  standing?.dismiss();

  // Sorted here rather than trusted from the projection: the walk reads controls in the
  // collection's order, which is creation order, and the whole point of this dialog is that the
  // two are not the same thing.
  let rows = [...options.controls].sort((a, b) => a.tabIndex - b.tabIndex);
  let selected = rows[0]?.name ?? "";

  const modal = openModal({
    backdropId: "taborder-backdrop",
    cardId: "taborder-card",
    label: `Tab order for ${options.container === "" ? "the form" : options.container}`,
    closed: () => {
      standing = null;
      options.closed?.();
    },
  });

  const heading = document.createElement("h2");
  heading.className = "modal-title";
  heading.textContent = "Tab Order";

  const where = document.createElement("p");
  where.className = "modal-detail";
  where.textContent = options.container === ""
    ? "The controls on the form, in the order Tab reaches them."
    : `The controls in ${options.container}, in the order Tab reaches them.`;

  const list = document.createElement("div");
  list.className = "taborder-list";
  list.setAttribute("role", "listbox");
  list.tabIndex = 0;

  const up = document.createElement("button");
  up.type = "button";
  up.className = "modal-button";
  up.textContent = "Move Up";

  const down = document.createElement("button");
  down.type = "button";
  down.className = "modal-button";
  down.textContent = "Move Down";

  const close = document.createElement("button");
  close.type = "button";
  close.className = "modal-button primary";
  close.textContent = "Close";

  const paint = (): void => {
    list.replaceChildren();
    for (const [at, row] of rows.entries()) {
      const item = document.createElement("div");
      item.className = "taborder-row" + (row.name === selected ? " current" : "");
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", String(row.name === selected));
      item.dataset.control = row.name;

      const rank = document.createElement("span");
      rank.className = "taborder-rank";
      rank.textContent = String(at);

      const name = document.createElement("span");
      name.className = "taborder-name";
      name.textContent = row.name;

      const kind = document.createElement("span");
      kind.className = "taborder-kind";
      kind.textContent = row.type;

      item.append(rank, name, kind);
      item.addEventListener("pointerdown", () => {
        selected = row.name;
        paint();
      });
      list.appendChild(item);
    }

    const at = rows.findIndex((row) => row.name === selected);
    up.disabled = at <= 0;
    down.disabled = at < 0 || at >= rows.length - 1;
  };

  /**
   * One place up or down. The list is re-numbered here for the picture and the HOST is asked for
   * the one write that means it - the moved control takes the index it is going to, and MSForms
   * pushes the other one along, which is what the native dialog does too.
   */
  const move = (name: string, by: -1 | 1): boolean => {
    const at = rows.findIndex((row) => row.name === name);
    const to = at + by;
    if (at < 0 || to < 0 || to >= rows.length) {
      return false;
    }

    const moved = rows[at] as TabOrderControl;
    const displaced = rows[to] as TabOrderControl;
    rows = rows.filter((row) => row.name !== name);
    rows.splice(to, 0, moved);
    // Renumbered by POSITION afterwards, because that is what the host will report back and what
    // a second move has to measure from.
    rows = rows.map((row, index) => ({ ...row, tabIndex: index }));

    selected = name;
    paint();
    options.setIndex(moved.name, displaced.tabIndex);
    return true;
  };

  up.addEventListener("click", () => { move(selected, -1); });
  down.addEventListener("click", () => { move(selected, 1); });
  close.addEventListener("click", () => { modal.dismiss(); });

  list.addEventListener("keydown", (event) => {
    const at = rows.findIndex((row) => row.name === selected);
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const next = rows[at + (event.key === "ArrowDown" ? 1 : -1)];
      if (next) {
        // Ctrl+arrow MOVES where a bare arrow only walks, which is the pairing the canvas keeps
        // between a nudge and a resize.
        if (event.ctrlKey) {
          move(selected, event.key === "ArrowDown" ? 1 : -1);
        } else {
          selected = next.name;
          paint();
        }
      }
    }
  });

  const buttons = document.createElement("div");
  buttons.className = "modal-buttons";
  buttons.append(up, down, close);

  modal.card.append(heading, where, list, buttons);
  paint();
  list.focus();

  standing = {
    order: () => rows.map((row) => row.name),
    move,
    dismiss: modal.dismiss,
  };
  return standing;
}
