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

  /*
   * DRAGGING A ROW WHERE IT GOES, which is the gesture a list like this asks for: the whole
   * sequence is on screen, so "third from the top" is a place a hand can point at, and Move Up
   * pressed five times is the same edit spelled as five (the owner, 2026-08-18: "please add drag
   * and drop to tab order modal"). The buttons and Ctrl+arrow stay, because a pointer is not the
   * only way in and a keyboard user is not to be sent to a mouse.
   *
   * ONE WRITE AT THE DROP, not one per row crossed. The picture follows the pointer, but nothing
   * is asked of the host until the hand lets go - a drag across six controls that ended where it
   * started must cost the form nothing at all.
   *
   * The pointer is captured on the LIST rather than on the row, because selecting repaints and a
   * repainted row is a detached element: capture taken on it is lost mid-gesture, which is the
   * first way this was written and it dropped the drag the instant it began.
   */
  const marker = document.createElement("div");
  marker.className = "taborder-marker";
  marker.hidden = true;

  /** Where a drop at this pointer position would insert, as an index into the visible rows. */
  const dropAt = (clientY: number): number => {
    const items = [...list.querySelectorAll<HTMLElement>(".taborder-row")];
    for (const [at, item] of items.entries()) {
      const box = item.getBoundingClientRect();
      if (clientY < box.top + box.height / 2) {
        return at;
      }
    }

    return items.length;
  };

  /** The insertion line, drawn at the gap a drop would open. */
  const showMarker = (before: number): void => {
    const items = [...list.querySelectorAll<HTMLElement>(".taborder-row")];
    const at = items[before];
    const last = items[items.length - 1];
    marker.style.top = at
      ? `${at.offsetTop}px`
      : `${last === undefined ? 0 : last.offsetTop + last.offsetHeight}px`;
    marker.hidden = false;
  };

  let drag: { name: string; from: number; to: number; y: number; moving: boolean } | null = null;

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
      list.appendChild(item);
    }

    // The insertion line lives in the list and survives every repaint, so it is put back after
    // the rows rather than rebuilt with them.
    list.appendChild(marker);

    const at = rows.findIndex((row) => row.name === selected);
    up.disabled = at <= 0;
    down.disabled = at < 0 || at >= rows.length - 1;
  };

  /**
   * To a PLACE, however far. The list is re-numbered here for the picture and the HOST is asked
   * for the one write that means it - the moved control takes the index it is going to, and
   * MSForms pushes the others along, which is what the native dialog does too. That renumbering
   * is why distance costs nothing: a drag of six places is the same single write as a nudge of
   * one, so there is one write path and Move Up is a special case of it rather than a second
   * implementation.
   */
  const moveTo = (name: string, to: number): boolean => {
    const at = rows.findIndex((row) => row.name === name);
    if (at < 0 || to < 0 || to >= rows.length || to === at) {
      return false;
    }

    const moved = rows[at] as TabOrderControl;
    // The index the destination is CURRENTLY holding, read before the splice: that is the number
    // the moved control has to take for MSForms to push the rest the way the developer means.
    const taking = (rows[to] as TabOrderControl).tabIndex;
    rows = rows.filter((row) => row.name !== name);
    rows.splice(to, 0, moved);
    // Renumbered by POSITION afterwards, because that is what the host will report back and what
    // a second move has to measure from.
    rows = rows.map((row, index) => ({ ...row, tabIndex: index }));

    selected = name;
    paint();
    options.setIndex(moved.name, taking);
    return true;
  };

  /** One place up or down: the buttons' and the keyboard's move, through the one above. */
  const move = (name: string, by: -1 | 1): boolean =>
    moveTo(name, rows.findIndex((row) => row.name === name) + by);

  list.addEventListener("pointerdown", (event) => {
    const item = (event.target as HTMLElement).closest<HTMLElement>(".taborder-row");
    const name = item?.dataset.control;
    if (name === undefined || event.button !== 0) {
      return;
    }

    // Selecting first, so a press that never becomes a drag is exactly the click it used to be.
    selected = name;
    paint();
    const from = rows.findIndex((row) => row.name === name);
    drag = { name, from, to: from, y: event.clientY, moving: false };
    try {
      list.setPointerCapture(event.pointerId);
    } catch {
      // A synthesised pointer has no capture to take; the gesture runs on bubbling alone, which
      // is all the harness's own pointer sequence needs.
    }
  });

  list.addEventListener("pointermove", (event) => {
    if (drag === null) {
      return;
    }

    // A few pixels of slack before this becomes a drag, so a click with an unsteady hand stays a
    // click and the list does not flash an insertion line under every selection.
    if (!drag.moving && Math.abs(event.clientY - drag.y) < 4) {
      return;
    }

    drag.moving = true;
    list.querySelector(`.taborder-row[data-control="${CSS.escape(drag.name)}"]`)
      ?.classList.add("lifting");
    drag.to = dropAt(event.clientY);
    showMarker(drag.to);
  });

  const endDrag = (event: PointerEvent): void => {
    if (drag === null) {
      return;
    }

    const { name, from, to, moving } = drag;
    drag = null;
    marker.hidden = true;
    try {
      if (list.hasPointerCapture(event.pointerId)) {
        list.releasePointerCapture(event.pointerId);
      }
    } catch {
      // Nothing was captured, which is the synthesised-pointer case again.
    }

    if (!moving) {
      return;
    }

    // The drop index counts the row being moved while it is still in the list, so a downward
    // drag lands one place short of where the line was drawn unless it is taken back out first.
    const landing = to > from ? to - 1 : to;
    if (landing === from) {
      paint();
      return;
    }

    moveTo(name, Math.min(landing, rows.length - 1));
  };

  list.addEventListener("pointerup", endDrag);
  list.addEventListener("pointercancel", endDrag);

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
