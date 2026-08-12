/*
 * The one modal scaffold: backdrop, card, captured Escape, backdrop dismissal, a Tab trap,
 * and focus restore.
 *
 * Six dialogs hand-rolled this - help, sponsor, settings, references, sync, and the shell's
 * remove confirm - and the thing a modal SHOULD do was therefore written zero times: every
 * card said aria-modal="true" while Tab walked straight out of it into the editor behind.
 * The trap lives here once now, and so do the pieces the six copies each spelled out. The
 * shared .modal-backdrop and .modal-card rules carry the geometry; a dialog's own ids stay
 * on the elements so its width, padding and probes keep working.
 */

export interface ModalOptions {
  /** The backdrop element's id, kept per dialog so its CSS overrides and probes still bind. */
  backdropId: string;

  /** The card element's id, kept for the same reason. */
  cardId: string;

  /** The card's accessible name. */
  label: string;

  /** "alertdialog" for a question that interrupts; "dialog" (the default) otherwise. */
  role?: "dialog" | "alertdialog";

  /** Runs once after the modal is gone and focus is back where it was. */
  closed?: (() => void) | undefined;
}

export interface OpenModal {
  backdrop: HTMLDivElement;
  card: HTMLDivElement;

  /** Takes the modal down. Idempotent: Escape, the backdrop, and a Close button can share it. */
  dismiss: () => void;
}

/**
 * Everything in the card the keyboard can land on, in DOM order. Queried live on every Tab
 * rather than snapshotted at open, because dialog content changes while it stands - the sync
 * dialog rebuilds its whole plan table on refresh.
 */
function focusRing(card: HTMLElement): HTMLElement[] {
  return [...card.querySelectorAll<HTMLElement>(
    'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])',
  )].filter((one) => !one.hasAttribute("disabled") && one.offsetParent !== null);
}

/**
 * Opens an empty modal and hands back its card for the caller to fill. The caller focuses its
 * own first control once the content is in; the trap keeps the cycle inside from then on.
 */
export function openModal(options: ModalOptions): OpenModal {
  const backdrop = document.createElement("div");
  backdrop.id = options.backdropId;
  backdrop.className = "modal-backdrop";

  const card = document.createElement("div");
  card.id = options.cardId;
  card.className = "modal-card";
  card.setAttribute("role", options.role ?? "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-label", options.label);

  // Where focus came from, to give it back. aria-modal promises the page behind is inert
  // while the card stands; restoring focus on the way out is that promise kept in reverse.
  const wasFocused = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  let standing = true;

  const dismiss = (): void => {
    if (!standing) {
      return;
    }

    standing = false;
    document.removeEventListener("keydown", onKey, true);
    backdrop.remove();

    if (wasFocused?.isConnected) {
      wasFocused.focus();
    }

    // After the restore, so a closed() that chooses its own focus target wins.
    options.closed?.();
  };

  const onKey = (event: KeyboardEvent): void => {
    // Captured, because Monaco answers Escape too and would otherwise swallow it first.
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
      return;
    }

    // The trap aria-modal claims: Tab cycles inside the card, and a focus that escaped by
    // some other route is pulled back in on the next Tab rather than left wandering.
    if (event.key === "Tab") {
      const ring = focusRing(card);
      const first = ring[0];
      const last = ring[ring.length - 1];
      if (!first || !last) {
        event.preventDefault();
        return;
      }

      const active = document.activeElement;
      const inside = active instanceof Node && card.contains(active);
      if (event.shiftKey && (!inside || active === first)) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && (!inside || active === last)) {
        event.preventDefault();
        first.focus();
      }
    }
  };

  // Mousedown, not click: a drag that starts on the card and releases over the backdrop is a
  // missed text selection, not a request to close.
  backdrop.addEventListener("mousedown", (event) => {
    if (event.target === backdrop) {
      dismiss();
    }
  });
  document.addEventListener("keydown", onKey, true);

  backdrop.appendChild(card);
  document.body.appendChild(backdrop);

  return { backdrop, card, dismiss };
}
