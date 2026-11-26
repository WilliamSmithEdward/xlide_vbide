/*
 * The update dialog: what the host's self-updater found, and what to do about it.
 *
 * The host owns the truth, the way it does for settings. It asks GitHub, judges the answer
 * against the build that is running, downloads and verifies the installer, and starts it; this
 * card draws the state it sends and posts the button pressed. Nothing here decides whether a
 * version is newer or whether a download is whole - a page that reasoned about that would be a
 * second updater, and the two would disagree one day.
 *
 * One card for the whole path, from "checking" to "the installer is open", so the developer who
 * pressed Install watches the same box fill in rather than a sequence of dialogs.
 */

import { openModal } from "./modal.js";

/** The updater's state, as the host sends it with every change. */
export interface UpdateState {
  /** idle, checking, upToDate, available, skipped, downloading, verified, launched, exited or failed. */
  phase: string;
  /** The version this build is. */
  installed: string;
  /** Loaded from somewhere other than the installer's folder, so installing registers a different copy. */
  developmentBuild: boolean;
  loadedFrom: string | null;
  /** The Office application this session is in, for the sentence that asks to close it. */
  host: string;
  feed: string;
  checkedAt: string | null;
  manual: boolean;
  version: string | null;
  tag: string | null;
  publishedAt: string | null;
  notes: string | null;
  page: string | null;
  size: number;
  received: number;
  path: string | null;
  pid: number;
  exitCode: number | null;
  failure: string | null;
  skippedVersion: string | null;
}

/** The actions the card can ask of the host. `dismiss` is for the record; the host does nothing with it. */
export type UpdateAction = "check" | "install" | "skip" | "notes" | "cancel" | "dismiss";

/** What the dialog needs of the world: a host to send a choice to. */
export interface UpdateHandlers {
  act(action: UpdateAction): void;
}

const IDLE: UpdateState = {
  phase: "idle",
  installed: "",
  developmentBuild: false,
  loadedFrom: null,
  host: "Excel",
  feed: "",
  checkedAt: null,
  manual: false,
  version: null,
  tag: null,
  publishedAt: null,
  notes: null,
  page: null,
  size: 0,
  received: 0,
  path: null,
  pid: 0,
  exitCode: null,
  failure: null,
  skippedVersion: null,
};

let handlers: UpdateHandlers | null = null;
let state: UpdateState = IDLE;

/** Redraws the open card, when one is open. */
let redraw: (() => void) | null = null;

/** Wires the one way out of this module, once, when the bridge exists. */
export function installUpdateHandlers(next: UpdateHandlers): void {
  handlers = next;
}

/** The state as last sent. A snapshot: the host sends a new one for every change. */
export function currentUpdateState(): UpdateState {
  return state;
}

/**
 * Adopts the host's state. `open` is the host asking for the card: an automatic check with
 * news, any outcome of a check asked for by hand, or a failed step. A card already standing
 * redraws either way.
 */
export function applyUpdateState(next: UpdateState, open: boolean): void {
  state = next;
  if (open && !document.getElementById("update-card")) {
    openUpdateDialog();
    return;
  }

  redraw?.();
}

/** One line of what the updater knows, for the About dialog's row. */
export function describeUpdateState(): string {
  switch (state.phase) {
    case "checking":
      return "checking...";
    case "upToDate":
      return state.version && state.version !== state.installed
        ? `newer than the latest release, ${state.version}`
        : "up to date";
    case "available":
      return `${state.version} is available`;
    case "skipped":
      return `${state.version} is available (skipped)`;
    case "downloading":
    case "verified":
    case "launched":
      return `installing ${state.version}`;
    case "exited":
      return `${state.version} is available (the installer closed)`;
    case "failed":
      return "the last check failed";
    default:
      return "not checked yet";
  }
}

/**
 * Opens the card, or focuses the one that is open. With `check`, also asks the host to look
 * now - the wrench's Check for updates - and draws the asking rather than waiting for the
 * host to say so, which it does a moment later.
 */
export function openUpdateDialog(options: { check?: boolean } = {}, closed?: () => void): void {
  const existing = document.getElementById("update-card");
  if (existing) {
    if (options.check) {
      askFor("check");
    }
    redraw?.();
    existing.querySelector<HTMLElement>("#update-close")?.focus();
    return;
  }

  const { card, dismiss } = openModal({
    backdropId: "update-backdrop",
    cardId: "update-card",
    label: "Update xlide",
    closed: () => {
      redraw = null;
      closed?.();
    },
  });

  const head = document.createElement("div");
  head.id = "update-head";

  const title = document.createElement("div");
  title.id = "update-title";

  const close = document.createElement("button");
  close.type = "button";
  close.id = "update-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';
  close.addEventListener("click", dismiss);

  head.append(title, close);

  const body = document.createElement("div");
  body.id = "update-body";

  const buttons = document.createElement("div");
  buttons.id = "update-buttons";
  buttons.className = "modal-buttons";

  card.append(head, body, buttons);

  redraw = (): void => draw(title, body, buttons, dismiss);

  if (options.check) {
    askFor("check");
  }

  redraw();
  (buttons.querySelector<HTMLElement>("button.primary") ?? close).focus();
}

/**
 * Presses one of the open card's buttons by its action name, through the button's own click,
 * for the api. Answers what it did; a name the card is not showing is refused with the names
 * it is.
 */
export function pressUpdateButton(action: string): { did: boolean; detail: string } {
  const card = document.getElementById("update-card");
  if (!card) {
    return { did: false, detail: "the update dialog is not open" };
  }

  const shown = [...card.querySelectorAll<HTMLButtonElement>("button[data-action]")]
    .filter((button) => !button.disabled);
  const wanted = shown.find((button) => button.dataset.action === action);
  if (!wanted) {
    return {
      did: false,
      detail: `the update dialog has no '${action}' button; it shows: ${shown.map((b) => b.dataset.action).join(", ") || "none"}`,
    };
  }

  wanted.click();
  return { did: true, detail: `pressed ${action}` };
}

/** Asks the host, and draws the asking at once where the host will say the same a moment later. */
function askFor(action: UpdateAction): void {
  if (action === "check") {
    state = { ...state, phase: "checking", failure: null, manual: true };
  } else if (action === "install") {
    state = { ...state, phase: "downloading", received: 0, failure: null };
  }

  handlers?.act(action);
  redraw?.();
}

function draw(title: HTMLElement, body: HTMLElement, buttons: HTMLElement, dismiss: () => void): void {
  body.replaceChildren();
  buttons.replaceChildren();

  const line = (text: string, muted = false): void => {
    const paragraph = document.createElement("p");
    paragraph.className = muted ? "update-line update-muted" : "update-line";
    paragraph.textContent = text;
    body.appendChild(paragraph);
  };

  const button = (label: string, action: UpdateAction | "close", primary = false): HTMLButtonElement => {
    const control = document.createElement("button");
    control.type = "button";
    control.className = primary ? "modal-button primary" : "modal-button";
    control.textContent = label;
    control.dataset.action = action;
    control.addEventListener("click", () => {
      if (action === "close") {
        dismiss();
        return;
      }

      askFor(action);
    });
    buttons.appendChild(control);
    return control;
  };

  // Loaded from the repository rather than installed: say so where Install is offered, because
  // the installer registers the copy it writes under Programs\xlide, not this one.
  const developmentNote = (): void => {
    if (state.developmentBuild) {
      line(
        `This is a development build, loaded from ${state.loadedFrom ?? "outside the installer's folder"}. `
        + "The installer registers the copy it installs under Programs\\xlide instead.",
        true,
      );
    }
  };

  const published = state.publishedAt ? ` Published ${state.publishedAt.slice(0, 10)}.` : "";
  const megabytes = (bytes: number): string => (bytes / 1048576).toFixed(1);

  switch (state.phase) {
    case "checking":
      title.textContent = "Update xlide";
      line("Checking for the newest release...");
      button("Close", "close");
      break;

    case "upToDate":
      title.textContent = "Update xlide";
      if (state.version && state.version !== state.installed) {
        line(`You have xlide ${state.installed}, which is newer than the newest release, ${state.version}.`);
      } else {
        line(`You have the latest xlide, ${state.installed}.`);
      }
      button("Check again", "check");
      button("Close", "close", true);
      break;

    case "available":
    case "skipped":
      title.textContent = `xlide ${state.version} is available`;
      line(`You have ${state.installed}.${published}`);
      if (state.phase === "skipped") {
        line("You chose to skip this version. It is offered here only because you asked.", true);
      }
      if (state.notes) {
        const notes = document.createElement("pre");
        notes.className = "update-notes";
        notes.textContent = state.notes;
        body.appendChild(notes);
      }
      developmentNote();
      button("Install now", "install", true);
      button("Release notes", "notes");
      if (state.phase === "available") {
        button("Skip this version", "skip");
      }
      button("Not now", "close");
      break;

    case "downloading": {
      title.textContent = `Installing xlide ${state.version}`;
      line(`Downloading xlide-setup.exe: ${megabytes(state.received)} of ${megabytes(state.size)} MB.`);
      const progress = document.createElement("progress");
      progress.className = "update-progress";
      progress.max = Math.max(1, state.size);
      progress.value = Math.min(state.received, Math.max(1, state.size));
      progress.setAttribute("aria-label", "Download progress");
      body.appendChild(progress);
      button("Cancel", "cancel");
      break;
    }

    case "verified":
      title.textContent = `Installing xlide ${state.version}`;
      line("The download is whole and matches the release's digest. Starting the installer...");
      button("Close", "close");
      break;

    case "launched":
      title.textContent = `Installing xlide ${state.version}`;
      line(`The installer is open in its own window. Save your work and close ${state.host}; `
        + `it installs the moment ${state.host} has closed, and the next time you open the editor it is xlide ${state.version}.`);
      developmentNote();
      button("Close", "close", true);
      break;

    case "exited":
      title.textContent = `xlide ${state.version} was not installed`;
      line(`The installer closed before ${state.host} did`
        + `${state.exitCode === null ? "" : ` (exit code ${state.exitCode})`}, so nothing changed.`);
      if (state.path) {
        const path = document.createElement("p");
        path.className = "update-line update-path";
        path.textContent = `The verified installer is at ${state.path}; running it by hand installs the same build.`;
        body.appendChild(path);
      }
      button("Try again", "install", true);
      button("Close", "close");
      break;

    case "failed":
      title.textContent = "Update xlide";
      line(state.failure ?? "The update did not finish.");
      if (state.version) {
        line(`xlide ${state.version} is available; you have ${state.installed}.`, true);
        button("Try again", "install", true);
      } else {
        button("Try again", "check", true);
      }
      button("Close", "close");
      break;

    default:
      title.textContent = "Update xlide";
      line(`You have xlide ${state.installed}. It has not checked for a newer release yet.`);
      button("Check now", "check", true);
      button("Close", "close");
      break;
  }
}
