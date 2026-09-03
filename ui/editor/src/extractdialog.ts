/*
 * Naming the procedure an extraction is about to make.
 *
 * A refactoring asks exactly one question, and the answer is a name, so the dialog is one field.
 * It stays up when the host refuses: the refusals are about the SELECTION - a jump that crosses
 * the boundary, a With whose receiver stayed behind - and a developer who reads one has to be
 * able to cancel out of it and reselect, rather than watch it flash past as a toast.
 *
 * The refusal reads as an alert (role="alert"), so a screen reader hears it without the focus
 * moving off the field the developer is still in.
 */

import { openModal } from "./modal.js";

/** What the dialog needs of the world: somewhere to send the name. */
export type ExtractRequest = (name: string) => Promise<{ signature: string | null; refused: string | null }>;

/** What a check can read and drive without a pointer. */
export interface ExtractDialogProbe {
  state(): { open: boolean; name: string; busy: boolean; refused: string | null };
  /** Types a name into the field. */
  type(name: string): void;
  /** Presses a named control: extract, cancel. False when unknown. */
  press(control: string): boolean;
}

let live: ExtractDialogProbe | null = null;
let open: { dismiss: () => void } | null = null;

/** The dialog's probe, or null when it is not up. */
export const extractDialogProbe = (): ExtractDialogProbe | null => live;

/**
 * Asks for the new procedure's name and hands it to the host.
 *
 * `suggested` is a starting point rather than a decision - selected and ready to be typed over,
 * because a name a tool invents is nearly always the wrong one and being made to clear it first
 * is a small tax on every single use.
 */
export function openExtractDialog(
  suggested: string,
  request: ExtractRequest,
  heading = "Extract method",
  fieldLabel = "New procedure name",
): void {
  closeExtractDialog();

  const { card, dismiss } = openModal({
    backdropId: "extract-backdrop",
    cardId: "extract-card",
    label: heading,
    closed: () => {
      open = null;
      live = null;
    },
  });
  open = { dismiss };

  const title = document.createElement("h2");
  title.id = "extract-title";
  title.className = "modal-title";
  title.textContent = heading;

  const label = document.createElement("label");
  label.id = "extract-label";
  label.htmlFor = "extract-name";
  label.textContent = fieldLabel;

  const name = document.createElement("input");
  name.type = "text";
  name.id = "extract-name";
  name.spellcheck = false;
  name.autocomplete = "off";
  name.value = suggested;

  // Empty until something is refused, and present the whole time so a reader's live region is
  // there to be updated rather than being inserted with the news already in it.
  const problem = document.createElement("p");
  problem.id = "extract-problem";
  problem.setAttribute("role", "alert");
  problem.hidden = true;

  const buttons = document.createElement("div");
  buttons.id = "extract-buttons";
  buttons.className = "modal-buttons";

  const cancel = document.createElement("button");
  cancel.type = "button";
  cancel.id = "extract-cancel";
  cancel.className = "modal-button";
  cancel.textContent = "Cancel";
  cancel.addEventListener("click", () => closeExtractDialog());

  const extract = document.createElement("button");
  extract.type = "button";
  extract.id = "extract-go";
  extract.className = "modal-button primary";
  extract.textContent = "Extract";

  buttons.append(cancel, extract);
  card.append(title, label, name, problem, buttons);

  let busy = false;

  const refuse = (why: string | null): void => {
    problem.textContent = why ?? "";
    problem.hidden = why === null;
    name.setAttribute("aria-invalid", why === null ? "false" : "true");
    if (why !== null) {
      name.focus();
      name.select();
    }
  };

  const submit = async (): Promise<void> => {
    const wanted = name.value.trim();
    if (busy) {
      return;
    }

    if (wanted.length === 0) {
      refuse("Give the new procedure a name.");
      return;
    }

    busy = true;
    extract.disabled = true;
    extract.textContent = "Extracting...";

    let answer: { signature: string | null; refused: string | null };
    try {
      answer = await request(wanted);
    } catch {
      answer = { signature: null, refused: "The extraction could not be sent, so nothing changed." };
    }

    busy = false;
    extract.disabled = false;
    extract.textContent = "Extract";

    if (answer.refused) {
      refuse(answer.refused);
      return;
    }

    closeExtractDialog();
  };

  extract.addEventListener("click", () => void submit());
  name.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      void submit();
    }
  });

  // Clearing on the next keystroke rather than leaving the last refusal standing under a name
  // that has since changed, which reads as a refusal of the name being typed now.
  name.addEventListener("input", () => {
    if (!problem.hidden) {
      refuse(null);
    }
  });

  live = {
    state: () => ({
      open: open !== null,
      name: name.value,
      busy,
      refused: problem.hidden ? null : problem.textContent,
    }),
    type: (wanted) => {
      name.value = wanted;
      if (!problem.hidden) {
        refuse(null);
      }
    },
    press: (control) => {
      if (control === "extract") {
        void submit();
        return true;
      }

      if (control === "cancel") {
        closeExtractDialog();
        return true;
      }

      return false;
    },
  };

  name.focus();
  name.select();
}

export function closeExtractDialog(): void {
  open?.dismiss();
}
