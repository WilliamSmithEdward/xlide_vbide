/*
 * The Source Control pane: the project's modules against the branch head, the folder against
 * the project, and the history behind both - with commit, import, export, the remotes, a
 * past-version tab and restore. docs/source-control.md is the design.
 *
 * IT PULLS RATHER THAN BEING PUSHED, the Changes pane's rule for the Changes pane's reason and
 * a heavier one: every answer here runs git.exe, which is seconds rather than milliseconds. So
 * the host is asked when the pane is opened, when a button is pressed, and a quiet moment after
 * the host taps it to say the folder or the repository moved - never on the write path.
 *
 * ONE BRAIN, TWO DOORS. Every request here is the xlide api's `scm` route in the same words,
 * and every action leaves the state the route's would: the pane redraws from the status the
 * host embeds in each action's reply, so what it shows after Commit is what `GET scm` answers.
 */

import type { OpenFiles } from "./changespane.js";
import { drawDiffRows, type SyncDiffLine } from "./diffview.js";
import { installSplitterDrag } from "./livedrag.js";
import { openModal } from "./modal.js";

/** One module live against the branch head, as the host reports it. */
export interface ScmRow {
  module: string;
  kind: string;
  file: string;
  status: string;
  from: string | null;
}

/** One module whose file in the folder differs from the live text. */
export interface ScmOutsideRow {
  module: string;
  kind: string;
  file: string;
  status: string;
}

export interface ScmBranch {
  /** The branch's own name; for a remote's branch, without the remote in front. */
  name: string;
  current: boolean;
  upstream: string | null;
  /** Empty for a local branch; the remote's name for a branch only that remote has. */
  remote: string;
}

export interface ScmCommitFile {
  module: string;
  file: string;
  status: string;
}

export interface ScmCommit {
  hash: string;
  short: string;
  author: string;
  email: string;
  when: string;
  subject: string;
  body: string;
  files: ScmCommitFile[];
}

/** The status reply, which every action's reply embeds as `status`. */
export interface ScmStatus {
  detail: string;
  project: string;
  projectId: string;
  state: string;
  folder: string;
  repository: string;
  gitVersion: string;
  branch: string;
  /** The commit the branch is at, in full; empty on an unborn branch. */
  head: string;
  upstream: string;
  /** The remote pushes go to - origin, else the only one - and its URL; empty with none. */
  remote: string;
  remoteUrl: string;
  ahead: number;
  behind: number;
  dirty: boolean;
  identity: { name: string; email: string } | null;
  rows: ScmRow[];
  outside: ScmOutsideRow[];
  branches: ScmBranch[];
  conflicts: string[];
  lastCommit: { hash: string; short: string; author: string; when: string; subject: string } | null;
  /** Empty when the branch head can be undone from here; else why it cannot. */
  undoBlocked: string;
  suggestedMessage: string;
  covers: string;
  /** A folder the host offers when none is remembered: the sync folder, or one beside the
   * workbook. Read from `suggestedFolder`, then `syncFolder`, then `folder`. */
  suggestedFolder: string;
}

/** How the pane reaches the host: the route's arguments, and the body an action may want. */
export type ScmRequest = (args: Record<string, string>, body?: string) => Promise<Record<string, unknown>>;

/** The editor's blame layer, as the pane's Blame button presses it for the active module. */
export interface BlameControl {
  /** Toggles blame on the active module; false when no module is on screen. */
  toggle(): boolean;
  /** Whether blame is on for the active module. */
  on(): boolean;
}

/** What the pane can be driven and read through, for the dev surface. */
export interface ScmPaneProbe {
  state(): {
    project: string;
    state: string;
    detail: string;
    folder: string;
    repository: string;
    branch: string;
    upstream: { name: string; ahead: number; behind: number } | null;
    /** As the select names them: a remote's branch with the remote in front. */
    branches: string[];
    head: string;
    /** Whether the head commit's Undo is on screen and enabled; `undoBlocked` says why not. */
    undoable: boolean;
    undoBlocked: string;
    dirty: boolean;
    files: string[];
    rows: ScmRow[];
    ticked: string[];
    outside: ScmOutsideRow[];
    commits: { hash: string; short: string; author: string; when: string; subject: string; files: ScmCommitFile[]; open: boolean }[];
    conflicts: string[];
    lastCommit: string | null;
    suggestedMessage: string;
    showing: string | null;
    message: string;
    busy: boolean;
    /** Whether the host has seen the folder or repository move since these rows were read. */
    behind: boolean;
    blameOn: boolean;
    remote: string;
    remoteUrl: string;
    /** The rows' width in pixels, as the divider leaves it. */
    listWidth: number;
  };
  /** Types a URL into the remote line's input, unfolding it first when a remote is attached. */
  setRemoteUrl(url: string): boolean;
  /**
   * Puts the divider where a drag to that width would leave it, floor and ceiling applied.
   * False for a width that is not a number, or while the pane has no width to measure against.
   */
  resizeList(width: number): boolean;
  /**
   * Presses a named control: refresh, commit, undo, export, import, fetch, pull, push, blame,
   * init, abort, browse, use, identity, open, restore. False when it is not on screen or is
   * disabled in this state.
   */
  press(control: string): boolean;
  /**
   * Makes a branch the way the select does: its "New branch..." entry, then the card's name
   * input and Create. False when the select is not usable in this state.
   */
  createBranch(name: string): boolean;
  /** Points the pane at another open file, through the select's own change event. */
  chooseFile(name: string): boolean;
  /** Ticks or unticks a Changes row, through its checkbox's own change event. */
  tick(module: string, on: boolean): boolean;
  /** Types the commit message, through the textarea's own input event. */
  setMessage(text: string): boolean;
  /** Opens a Changes row's comparison, as clicking it does. */
  show(module: string): boolean;
  /** Opens or closes a commit in the history, as clicking its head does. */
  openCommit(short: string): boolean;
  /** Opens a file's comparison under an opened commit, as clicking it does. */
  showFile(short: string, module: string): boolean;
  /** Picks a branch through the select's own change event, which is a checkout. */
  chooseBranch(name: string): boolean;
  /** Types the identity's name and email into the noIdentity state's inputs. */
  setIdentity(name: string, email: string): boolean;
}

const STATUS_WORD: Record<string, string> = {
  modified: "modified",
  added: "added",
  deleted: "deleted",
  renamed: "renamed",
  folderNewer: "folder is newer",
  missingInFolder: "not in the folder",
  missingInProject: "not in the project",
};

/** "2:04 PM" today, "Sep 8" this year, "Sep 8, 2025" before that. */
function whenOf(iso: string): string {
  const at = new Date(iso);
  if (Number.isNaN(at.getTime())) {
    return "";
  }

  const today = new Date();
  if (at.toDateString() === today.toDateString()) {
    return at.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }

  return at.getFullYear() === today.getFullYear()
    ? at.toLocaleDateString([], { month: "short", day: "numeric" })
    : at.toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" });
}

const asString = (value: unknown, fallback = ""): string => (typeof value === "string" ? value : fallback);
const asNumber = (value: unknown): number => (typeof value === "number" && Number.isFinite(value) ? value : 0);
const asArray = (value: unknown): Record<string, unknown>[] =>
  Array.isArray(value) ? value.filter((one): one is Record<string, unknown> => typeof one === "object" && one !== null) : [];

/**
 * A status from the wire, or null when the reply is not one.
 *
 * COERCED FIELD BY FIELD rather than cast: an action's reply embeds the status under `status`,
 * a refusal carries only `error`, and a reply from an older host may lack an array. A pane that
 * read `.length` off a missing array would take the whole draw down with it.
 */
function statusOf(raw: unknown): ScmStatus | null {
  if (typeof raw !== "object" || raw === null) {
    return null;
  }

  const reply = raw as Record<string, unknown>;
  if (typeof reply.state !== "string") {
    return null;
  }

  const identity = typeof reply.identity === "object" && reply.identity !== null
    ? reply.identity as Record<string, unknown>
    : null;
  const last = typeof reply.lastCommit === "object" && reply.lastCommit !== null
    ? reply.lastCommit as Record<string, unknown>
    : null;

  return {
    detail: asString(reply.detail),
    project: asString(reply.project),
    projectId: asString(reply.projectId),
    state: reply.state,
    folder: asString(reply.folder),
    repository: asString(reply.repository),
    gitVersion: asString(reply.gitVersion),
    branch: asString(reply.branch),
    head: asString(reply.head),
    upstream: asString(reply.upstream),
    remote: asString(reply.remote),
    remoteUrl: asString(reply.remoteUrl),
    ahead: asNumber(reply.ahead),
    behind: asNumber(reply.behind),
    dirty: reply.dirty === true,
    identity: identity ? { name: asString(identity.name), email: asString(identity.email) } : null,
    rows: asArray(reply.rows).map((row) => ({
      module: asString(row.module),
      kind: asString(row.kind),
      file: asString(row.file),
      status: asString(row.status),
      from: typeof row.from === "string" && row.from.length > 0 ? row.from : null,
    })),
    outside: asArray(reply.outside).map((row) => ({
      module: asString(row.module),
      kind: asString(row.kind),
      file: asString(row.file),
      status: asString(row.status),
    })),
    branches: asArray(reply.branches).map((row) => ({
      name: asString(row.name),
      current: row.current === true,
      upstream: typeof row.upstream === "string" && row.upstream.length > 0 ? row.upstream : null,
      remote: asString(row.remote),
    })),
    conflicts: Array.isArray(reply.conflicts) ? reply.conflicts.filter((one): one is string => typeof one === "string") : [],
    lastCommit: last
      ? {
        hash: asString(last.hash),
        short: asString(last.short),
        author: asString(last.author),
        when: asString(last.when),
        subject: asString(last.subject),
      }
      : null,
    undoBlocked: asString(reply.undoBlocked),
    suggestedMessage: asString(reply.suggestedMessage),
    covers: asString(reply.covers),
    suggestedFolder: asString(reply.suggestedFolder) || asString(reply.syncFolder),
  };
}

function commitsOf(raw: unknown): ScmCommit[] {
  return asArray(raw).map((commit) => ({
    hash: asString(commit.hash),
    short: asString(commit.short) || asString(commit.hash).slice(0, 7),
    author: asString(commit.author),
    email: asString(commit.email),
    when: asString(commit.when),
    subject: asString(commit.subject),
    body: asString(commit.body),
    files: asArray(commit.files).map((file) => ({
      module: asString(file.module),
      file: asString(file.file),
      status: asString(file.status),
    })),
  }));
}

/** What the comparison strip is showing: a live row against the head, or a commit's file. */
type Showing =
  | { kind: "row"; module: string }
  | { kind: "file"; short: string; hash: string; module: string };

/** How many commits the history lists. The list says when there are more. */
const LOG_LIMIT = 50;

/** Where the rows' width is kept across reloads, once the divider has been dragged. */
const LIST_STORAGE_KEY = "xlide.scm.v1";
/** The narrowest the rows go: a module name and its status word still read. */
const LIST_FLOOR = 180;
/** What the other half keeps whatever the drag: the message box and a comparison worth reading. */
const LIST_KEEP = 260;
/** One arrow key's worth, the dock splitters' step. */
const LIST_STEP = 24;

/** The select's last entry, which asks for a name rather than being a branch. */
const NEW_BRANCH = "+new";

let livePane: ScmPaneProbe | null = null;

/** The Source Control pane's probe, or null before one has been built. */
export const scmPaneProbe = (): ScmPaneProbe | null => livePane;

export class ScmPane {
  private readonly list: HTMLElement;
  private readonly diff: HTMLElement;
  private readonly notice: HTMLElement;
  private readonly message: HTMLTextAreaElement;
  private readonly title: HTMLElement;
  private readonly refresh: HTMLButtonElement;
  private readonly commit: HTMLButtonElement;
  private readonly exportButton: HTMLButtonElement;
  private readonly importButton: HTMLButtonElement;
  private readonly fetch: HTMLButtonElement;
  private readonly pull: HTMLButtonElement;
  private readonly push: HTMLButtonElement;
  private readonly blameButton: HTMLButtonElement;
  private readonly branch: HTMLSelectElement;
  private readonly file: HTMLSelectElement;
  private readonly body: HTMLElement;
  private readonly splitter: HTMLElement;

  /** The rows' width once the divider has been dragged; null is the stylesheet's third. */
  private listWidth: number | null = null;
  private warnedAboutStorage = false;

  /** The stamp the rows were read at, against the newest the host has sent - see stamped(). */
  private drawnStamp = 0;
  private hostStamp = 0;

  /** Whether the remote line shows its input over an attached remote: Change was pressed. */
  private remoteEditing = false;
  private followTimer: ReturnType<typeof setTimeout> | undefined;

  private state: ScmStatus | null = null;
  private commits: ScmCommit[] = [];
  private showing: Showing | null = null;
  private busy = false;

  /**
   * The comparison on screen, kept so a redraw after a read does not take it away: every read
   * ends in draw(), and a developer reading a diff while the folder exports behind them would
   * otherwise be left with "Comparing..." and nothing asking again.
   */
  private shownRows: SyncDiffLine[] | null = null;
  private shownTitle = "";
  private shownDetail = "";
  /** Whether a read has been asked for at all: before the first, the pane offers Refresh. */
  private asked = false;

  /**
   * Rows the developer UNTICKED, by module. Kept as the inverse so a row arriving in a later
   * read is ticked by default - a change the developer has not seen is one they most likely
   * mean to commit, and the tick is where they say otherwise.
   */
  private readonly unticked = new Set<string>();

  /** Commits unfolded in the history, by short hash. */
  private readonly opened = new Set<string>();

  /** The suggestion the message box was last filled from, so a fresh one may replace it. */
  private lastSuggested = "";

  /** What the select was last built from, so an unchanged session leaves an open popup alone. */
  private fileSignature = "";

  constructor(
    private readonly root: HTMLElement,
    private readonly ask: ScmRequest,
    private readonly files: OpenFiles,
    private readonly blame: BlameControl) {
    this.list = root.querySelector("#scm-list") as HTMLElement;
    this.diff = root.querySelector("#scm-diff") as HTMLElement;
    this.notice = root.querySelector("#scm-notice") as HTMLElement;
    this.message = root.querySelector("#scm-message") as HTMLTextAreaElement;
    this.title = root.querySelector("#scm-project") as HTMLElement;
    this.refresh = root.querySelector("#scm-refresh") as HTMLButtonElement;
    this.commit = root.querySelector("#scm-commit") as HTMLButtonElement;
    this.exportButton = root.querySelector("#scm-export") as HTMLButtonElement;
    this.importButton = root.querySelector("#scm-import") as HTMLButtonElement;
    this.fetch = root.querySelector("#scm-fetch") as HTMLButtonElement;
    this.pull = root.querySelector("#scm-pull") as HTMLButtonElement;
    this.push = root.querySelector("#scm-push") as HTMLButtonElement;
    this.blameButton = root.querySelector("#scm-blame") as HTMLButtonElement;
    this.branch = root.querySelector("#scm-branch") as HTMLSelectElement;
    this.file = root.querySelector("#scm-file") as HTMLSelectElement;
    this.body = root.querySelector("#scm-body") as HTMLElement;
    this.splitter = root.querySelector("#scm-splitter") as HTMLElement;

    livePane = this.probe();
    this.dragList();

    // A FILE CHOICE IS A DIFFERENT REPOSITORY, not a filter over one: every project remembers
    // its own folder. What is on screen is dropped first, so a comparison from the old file's
    // rows is never read as the new one's.
    this.file.addEventListener("change", () => {
      this.forgetProject();
      void this.run();
    });

    this.filesChanged();

    this.refresh.addEventListener("click", () => void this.run());
    this.commit.addEventListener("click", () => void this.runCommit());
    this.exportButton.addEventListener("click", () => void this.run({ action: "export" }));
    this.importButton.addEventListener("click", () => void this.run({ action: "import" }));
    this.fetch.addEventListener("click", () => void this.run({ action: "fetch" }));
    this.pull.addEventListener("click", () => void this.run({ action: "pull" }));
    this.push.addEventListener("click", () => void this.run({ action: "push" }));
    this.blameButton.addEventListener("click", () => {
      this.blame.toggle();
      this.drawBlameButton();
    });

    // Picking a branch IS the checkout. The select is redrawn from the status that comes back,
    // so a refused checkout (a dirty workbook) leaves it showing the branch that is still current.
    // The last entry is not a branch: the select goes back to the one that is, and a card asks
    // for the new one's name.
    //
    // NOT ON EVERY ARROW KEY. A closed select fires change for each arrow press and each typed
    // letter, and each would have been a checkout and an import (a review, 2026-09-09). So a
    // change the keyboard walked to waits until Enter, or the select losing focus, says the walk
    // is over, and Escape puts the select back; a pick with the mouse, or from the opened list,
    // acts at once.
    let walking = false;
    let pending = false;
    const act = (): void => {
      pending = false;
      const ref = this.branch.value;
      if (ref === NEW_BRANCH) {
        this.branch.value = this.state?.branch ?? "";
        this.askNewBranch();
        return;
      }

      if (ref && ref !== this.state?.branch) {
        void this.run({ action: "checkout", ref });
      }
    };
    this.branch.addEventListener("mousedown", () => {
      walking = false;
    });
    this.branch.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " " || (event.altKey && (event.key === "ArrowDown" || event.key === "ArrowUp"))) {
        // Enter settles a walk; Space and Alt+Arrow open the list, whose pick is deliberate.
        walking = false;
        if (event.key === "Enter" && pending) {
          event.preventDefault();
          act();
        }
      } else if (event.key === "Escape") {
        if (pending) {
          pending = false;
          this.branch.value = this.state?.branch ?? "";
        }
      } else if (event.key.startsWith("Arrow") || event.key === "Home" || event.key === "End"
        || event.key === "PageUp" || event.key === "PageDown" || (event.key.length === 1 && !event.ctrlKey && !event.metaKey)) {
        walking = true;
      }
    });
    this.branch.addEventListener("change", () => {
      if (walking) {
        pending = true;
        return;
      }

      act();
    });
    this.branch.addEventListener("blur", () => {
      if (pending) {
        act();
      }
    });

    this.message.addEventListener("input", () => this.enable());

    this.draw();
  }

  /**
   * The host saw the folder or the repository move. The pane follows a quiet moment later, once
   * it has been read at all: a burst of exports is one re-read, and a pane nobody has opened
   * stays free. The DOM is not touched here; `behind` is read by the probe.
   */
  stamped(stamp: number): void {
    if (stamp <= this.hostStamp) {
      return;
    }

    this.hostStamp = stamp;
    this.followLater();
  }

  /**
   * Schedules the re-read a stamp asks for: a quiet moment later, only once the pane has been
   * asked at all, and only while it is showing. A hidden pane runs no git - it is behind, and
   * shown() reads it when it comes forward. A pane still busy with an action when the moment
   * comes is re-armed rather than dropped: a commit or a push takes seconds, and the stamp its
   * own export raised, or an outside edit that landed meanwhile, is still owed a read.
   */
  private followLater(): void {
    clearTimeout(this.followTimer);
    if (!this.asked || this.root.hidden) {
      return;
    }

    this.followTimer = setTimeout(() => {
      this.followTimer = undefined;
      if (this.busy) {
        this.followLater();
        return;
      }

      void this.run();
    }, 1500);
  }

  /** Asked for when the pane is opened, which is the only time any of this costs anything. */
  shown(): void {
    this.filesChanged();
    void this.run();
  }

  /** The blame layer moved on its own (the editor's action or the api); the button follows. */
  blameChanged(): void {
    this.drawBlameButton();
  }

  /** The blame layer asked the host and was refused: the words go where every refusal goes. */
  blameFailed(message: string): void {
    this.setNotice(message, true);
    this.drawBlameButton();
  }

  /**
   * Rebuilds the file list from the session as it stands. A closed file drops out, and if it
   * was the one being shown the pane falls back to the file the developer is actually in.
   */
  filesChanged(): void {
    const { names, current } = this.files();
    const signature = JSON.stringify([names, current]);
    if (signature === this.fileSignature) {
      return;
    }

    this.fileSignature = signature;
    const chosen = this.file.value;
    const keep = names.some((name) => name.toLowerCase() === chosen.toLowerCase());

    this.file.replaceChildren();
    for (const name of names) {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      this.file.appendChild(option);
    }

    // One file is not a choice; the name beside it goes the other way, so the workbook is said
    // exactly once either way.
    this.file.hidden = names.length < 2;
    this.title.hidden = !this.file.hidden;
    this.file.value = keep ? chosen : current ?? names[0] ?? "";

    if (this.file.value !== chosen) {
      this.forgetProject();
      if (this.asked) {
        void this.run();
      } else {
        this.draw();
      }
    }
  }

  private forgetProject(): void {
    this.showing = null;
    this.shownRows = null;
    this.state = null;
    this.commits = [];
    this.unticked.clear();
    this.opened.clear();
    this.lastSuggested = "";
    this.message.value = "";
    this.setNotice("", false);
  }

  private withProject(args: Record<string, string>): Record<string, string> {
    return this.file.value ? { project: this.file.value, ...args } : args;
  }

  /**
   * One round trip: the status, or an action whose reply embeds the status. Then, when there
   * is a repository, the history. Busy-gated, so two presses are one ask.
   */
  private async run(args: Record<string, string> = {}, body = ""): Promise<Record<string, unknown> | null> {
    if (this.busy) {
      return null;
    }

    this.busy = true;
    this.asked = true;
    this.enable();
    if (!this.state) {
      this.draw();
    }

    // Stamped BEFORE the ask: a move recorded while the request is in flight may or may not be
    // in this answer, and "there is newer" is the safe direction to err in.
    const asOf = this.hostStamp;
    let answer: Record<string, unknown> | null = null;
    try {
      answer = await this.ask(this.withProject(args), body);

      if (typeof answer.error === "string") {
        this.setNotice(answer.error, true);
      } else {
        const status = statusOf(answer.status ?? answer);
        if (status) {
          this.adopt(status);
          this.drawnStamp = asOf;
        }

        // The input folds away once the remote it typed is the one the status shows.
        if (args.action === "remote") {
          this.remoteEditing = false;
        }

        // An undone commit's message returns to the box - over an empty box or the suggestion
        // the status just put there; a message the developer typed is theirs and stays. Held as
        // typed from then on, so a later suggestion does not replace it either.
        if (args.action === "undo" && typeof answer.message === "string") {
          this.fillMessage(answer.message);
          this.lastSuggested = "";

          // A comparison of a file at the undone commit, and the commit's unfolded state, are
          // about a commit the branch no longer has.
          const gone = asString(answer.short);
          this.opened.delete(gone);
          if (this.showing?.kind === "file" && this.showing.short === gone) {
            this.showing = null;
            this.shownRows = null;
          }
        }

        // An action's own words are worth showing; a bare read's "ready" is not.
        this.setNotice(args.action ? asString(answer.detail) : "", false);
        if (this.state?.state === "ready") {
          this.commits = await this.readLog();
        } else {
          this.commits = [];
        }
      }
    } catch (failed) {
      this.setNotice(failed instanceof Error ? failed.message : String(failed), true);
    } finally {
      this.busy = false;
      this.draw();

      // The comparison on screen is asked again, because its rows may have moved with the
      // status; and a stamp that landed while this read was in flight is still owed one.
      if (this.showing) {
        void this.refreshComparison();
      }
      if (this.hostStamp > this.drawnStamp) {
        this.followLater();
      }
    }

    return answer;
  }

  /** Re-asks the comparison on screen, without taking it down while the answer is on its way. */
  private async refreshComparison(): Promise<void> {
    const showing = this.showing;
    if (!showing) {
      return;
    }

    if (showing.kind === "row") {
      await this.openRow(showing.module, true);
    } else {
      await this.openFileAt(showing.short, showing.hash, showing.module, true);
    }
  }

  private async readLog(): Promise<ScmCommit[]> {
    const answer = await this.ask(this.withProject({ action: "log", limit: String(LOG_LIMIT) }));
    return typeof answer.error === "string" ? [] : commitsOf(answer.commits);
  }

  private adopt(status: ScmStatus): void {
    this.state = status;

    // A row that left the rows leaves the untick set too, so a module that comes back later
    // arrives ticked like any newcomer.
    const named = new Set(status.rows.map((row) => row.module.toLowerCase()));
    for (const held of [...this.unticked]) {
      if (!named.has(held)) {
        this.unticked.delete(held);
      }
    }

    // A comparison of a row that is no longer a row is history, not a lie: it stays until the
    // developer clicks elsewhere, and the read that brought this status asks it again once it
    // has drawn (run's finally), because the rows it lined up may have moved.
    this.fillMessage(status.suggestedMessage);
  }

  /**
   * The message box comes pre-filled from the change-log rounds since the last commit, which is
   * where an agent's work becomes a reviewable commit. Only over an EMPTY box or the previous
   * suggestion: a message the developer typed is theirs.
   */
  private fillMessage(suggested: string): void {
    const current = this.message.value;
    if (current.trim() === "" || current === this.lastSuggested) {
      this.message.value = suggested;
    }
    this.lastSuggested = suggested;
  }

  private async runCommit(): Promise<void> {
    const message = this.message.value.trim();
    const ticked = this.tickedModules();
    if (!message) {
      this.setNotice("Type a commit message first.", true);
      this.message.focus();
      return;
    }

    if (ticked.length === 0) {
      this.setNotice("Tick at least one row to commit.", true);
      return;
    }

    const answer = await this.run({ action: "commit", message }, ticked.join("\n"));
    if (answer && typeof answer.error !== "string") {
      // The message is spent. The box refills from the next suggestion, which after a commit
      // of everything is empty.
      this.message.value = "";
      this.lastSuggested = "";
      this.fillMessage(this.state?.suggestedMessage ?? "");
      this.enable();
    }
  }

  private tickedModules(): string[] {
    return (this.state?.rows ?? [])
      .filter((row) => !this.unticked.has(row.module.toLowerCase()))
      .map((row) => row.module);
  }

  private setNotice(text: string, error: boolean): void {
    this.notice.textContent = text;
    this.notice.classList.toggle("scm-notice-error", error && text.length > 0);
  }

  /** Which controls mean anything in this state. */
  private enable(): void {
    const ready = this.state?.state === "ready" && !this.busy;
    this.refresh.disabled = this.busy;
    this.commit.disabled = !ready || this.tickedModules().length === 0 || this.message.value.trim() === "";
    this.exportButton.disabled = !ready;
    this.importButton.disabled = !ready || (this.state?.outside.length ?? 0) === 0;
    // The remote verbs need a remote, and without one a press would only be refused; the remote
    // line above the rows is where one is attached.
    const attached = ready && (this.state?.remoteUrl ?? "") !== "";
    this.fetch.disabled = !attached;
    this.pull.disabled = !attached;
    this.push.disabled = !attached;
    this.branch.disabled = !ready;
    // Shown with no branch at all, because its last entry is how the first one is made.
    this.branch.hidden = !this.state || this.state.state !== "ready";
    this.message.hidden = !this.state || this.state.state !== "ready";

    // Blame reads the repository the pane is showing; with none there is nothing to paint from,
    // and a press would only be refused.
    this.blameButton.disabled = !this.state || this.state.state !== "ready";
  }

  private drawBlameButton(): void {
    this.blameButton.setAttribute("aria-pressed", this.blame.on() ? "true" : "false");
  }

  /**
   * The divider between the rows and the comparison: a drag, or the arrow keys, to say how much
   * of the pane the rows and the history are worth.
   *
   * A THIRD until it is dragged, whatever the pane's width: the pane docks anywhere, and a pixel
   * default that suits the bottom dock is most of a left one. Once dragged it is a width in
   * pixels, kept across reloads the way the dock splitters' sizes are, between a floor a row can
   * still be read at and a ceiling that leaves the message box and the comparison their room.
   */
  private dragList(): void {
    this.splitter.setAttribute("aria-valuemin", String(LIST_FLOOR));
    this.loadListWidth();

    // The width the drag is asking for, accumulated from the press rather than read back off the
    // list on each move: a list pinned at its ceiling would otherwise start shrinking the moment
    // the pointer turned back, however far past the ceiling it had gone. Listeners run in the
    // order they were added, so this one, added before the drag's own, sees the press first.
    let wanted = 0;
    this.splitter.addEventListener("pointerdown", () => {
      wanted = this.list.getBoundingClientRect().width;
    });
    installSplitterDrag(this.splitter, {
      positionOf: (event) => event.clientX,
      apply: (delta) => {
        wanted += delta;
        this.placeList(wanted);
      },
      // No editor lives in this pane, so a drag's end has nothing to lay out.
      settle: () => undefined,
      persist: () => this.rememberListWidth(),
    });

    this.splitter.addEventListener("keydown", (event) => {
      const step = event.key === "ArrowLeft" ? -LIST_STEP : event.key === "ArrowRight" ? LIST_STEP : 0;
      if (step === 0) {
        return;
      }

      event.preventDefault();
      this.placeList(this.list.getBoundingClientRect().width + step);
      this.rememberListWidth();
    });
  }

  /** The most the rows may take: what leaves the other half its room, and never under the floor. */
  private listCeiling(): number {
    return Math.max(LIST_FLOOR, this.body.clientWidth - LIST_KEEP);
  }

  private placeList(width: number): void {
    // Rounded AFTER the clamp: the ceiling is measured, and a pinned width would otherwise be
    // read out by a screen reader as 596.8000000000001 - the Changes pane's rail found this.
    this.listWidth = Math.round(Math.min(Math.max(LIST_FLOOR, width), this.listCeiling()));
    this.body.style.setProperty("--scm-list-width", `${this.listWidth}px`);
    this.splitter.setAttribute("aria-valuenow", String(this.listWidth));
    this.splitter.setAttribute("aria-valuemax", String(this.listCeiling()));
  }

  private loadListWidth(): void {
    try {
      const raw = localStorage.getItem(LIST_STORAGE_KEY);
      const stored = raw ? (JSON.parse(raw) as { listWidth?: unknown }).listWidth : undefined;
      if (typeof stored !== "number" || !Number.isFinite(stored)) {
        return;
      }

      // Not clamped against the ceiling: the pane may be hidden while the page loads, and a body
      // 0px wide would pin the width to the floor. The stylesheet caps the list on a pane
      // narrower than the width was left in, and the next drag clamps against the width that
      // then exists.
      this.listWidth = Math.max(LIST_FLOOR, Math.round(stored));
      this.body.style.setProperty("--scm-list-width", `${this.listWidth}px`);
      this.splitter.setAttribute("aria-valuenow", String(this.listWidth));
    } catch {
      // Storage off, or holding something that is not ours: the stylesheet's third, as on a
      // first run.
    }
  }

  private rememberListWidth(): void {
    if (this.listWidth === null) {
      return;
    }

    try {
      localStorage.setItem(LIST_STORAGE_KEY, JSON.stringify({ listWidth: this.listWidth }));
    } catch (error) {
      // Storage can be full or off. The width will not survive the reload, which is survivable -
      // but silently, "the divider keeps resetting" is a mystery with nothing behind it. Said
      // once: this runs on every drag, and a broken store would otherwise fill the console.
      if (!this.warnedAboutStorage) {
        this.warnedAboutStorage = true;
        console.warn("[xlide] the Source Control pane's divider could not be saved; it will not survive a reload", error);
      }
    }
  }

  private draw(): void {
    const state = this.state;
    this.title.textContent = state?.project ?? "";
    this.drawBranches();
    this.enable();
    this.drawBlameButton();

    this.list.replaceChildren();

    if (!state) {
      const empty = document.createElement("div");
      empty.className = "scm-empty";
      if (this.busy) {
        empty.textContent = "Reading the repository...";
      } else if (this.asked) {
        // The read came back with no status - a refusal, or a host that did not answer in time.
        // Its words are in the notice; the way on is to ask again, and it is offered here rather
        // than left to a "Reading..." that nothing will ever finish.
        const said = document.createElement("span");
        said.textContent = this.notice.textContent || "The repository could not be read.";
        const retry = document.createElement("button");
        retry.type = "button";
        retry.className = "panel-empty-act";
        retry.textContent = "Retry";
        retry.addEventListener("click", () => void this.run());
        empty.append(said, " ", retry);
      } else {
        // A pane restored open by the saved layout is shown without shown() firing, and a read
        // that runs git is not something to start on a hunch: it says how to ask.
        const said = document.createElement("span");
        said.textContent = "The repository is read when this pane is opened.";
        const read = document.createElement("button");
        read.type = "button";
        read.className = "panel-empty-act";
        read.textContent = "Read it now";
        read.addEventListener("click", () => void this.run());
        empty.append(said, " ", read);
      }
      this.list.appendChild(empty);
      this.drawDiff();
      return;
    }

    if (state.state !== "ready") {
      this.list.appendChild(this.drawState(state));
      this.drawDiff();
      return;
    }

    this.list.appendChild(this.drawRemote(state));
    this.list.appendChild(this.drawSection("Changes", state.rows.length, state.covers));
    if (state.rows.length === 0) {
      const clean = document.createElement("div");
      clean.className = "scm-empty";
      clean.textContent = state.lastCommit
        ? `Nothing to commit. Last commit ${state.lastCommit.short}: ${state.lastCommit.subject}`
        : "Nothing to commit.";
      this.list.appendChild(clean);
    }
    for (const row of state.rows) {
      this.list.appendChild(this.drawRow(row));
    }

    if (state.outside.length > 0) {
      this.list.appendChild(this.drawSection(
        "Folder", state.outside.length,
        "Modules whose file in the folder differs from the live text: a checkout, a pull, or an "
          + "edit in another editor. Import reads them into the project; Export overwrites them."));
      for (const row of state.outside) {
        this.list.appendChild(this.drawOutsideRow(row));
      }
    }

    this.list.appendChild(this.drawSection("History", this.commits.length, ""));
    if (this.commits.length === 0) {
      const none = document.createElement("div");
      none.className = "scm-empty";
      none.textContent = "No commits touch this folder yet.";
      this.list.appendChild(none);
    }
    for (const commit of this.commits) {
      this.list.appendChild(this.drawCommit(commit));
    }
    if (this.commits.length >= LOG_LIMIT) {
      const more = document.createElement("div");
      more.className = "scm-empty";
      more.textContent = `the newest ${LOG_LIMIT} commits`;
      this.list.appendChild(more);
    }

    this.drawDiff();
  }

  private drawBranches(): void {
    const state = this.state;
    const branches = state?.state === "ready" ? state.branches : [];
    const current = state?.branch ?? "";
    const signature = JSON.stringify([branches.map((one) => `${one.remote}/${one.name}`), current]);
    if (this.branch.dataset.signature !== signature) {
      this.branch.dataset.signature = signature;
      this.branch.replaceChildren();
      for (const one of branches.filter((one) => one.remote === "")) {
        const option = document.createElement("option");
        option.value = one.name;
        option.textContent = one.upstream ? `${one.name} (${one.upstream})` : one.name;
        this.branch.appendChild(option);
      }

      // A detached head is not a branch; it is shown as its own entry so the select says the
      // truth rather than the first branch. An unborn branch has no ref yet and is shown the
      // same way.
      if (current && !branches.some((one) => one.remote === "" && one.name === current)) {
        const option = document.createElement("option");
        option.value = current;
        option.textContent = current;
        this.branch.appendChild(option);
      }

      // The branches only a remote has, as of the last fetch. Picking one checks it out as a
      // local branch of the same name tracking it - what git's own checkout of a bare name does
      // when one remote has it - said with the remote in front here, so two remotes holding the
      // name cannot make it ambiguous.
      const remotes = branches.filter((one) => one.remote !== "");
      if (remotes.length > 0) {
        const group = document.createElement("optgroup");
        group.label = "Remote";
        for (const one of remotes) {
          const option = document.createElement("option");
          option.value = `${one.remote}/${one.name}`;
          option.textContent = `${one.remote}/${one.name}`;
          group.appendChild(option);
        }
        this.branch.appendChild(group);
      }

      // How a branch is made: the last entry, which asks for a name rather than being one.
      const fresh = document.createElement("option");
      fresh.value = NEW_BRANCH;
      fresh.textContent = "New branch...";
      this.branch.appendChild(fresh);
    }

    // ALWAYS, not only on a rebuild: a refused checkout leaves the options as they were and the
    // select showing the branch that was picked, which is exactly the one that is not current.
    this.branch.value = current;
    const behind = state && state.upstream
      ? ` ${state.ahead} ahead, ${state.behind} behind ${state.upstream}`
      : "";
    this.branch.title = current ? `Branch ${current}.${behind} Pick another to check it out.` : "";
  }

  private drawSection(name: string, count: number, hint: string): HTMLElement {
    const head = document.createElement("div");
    head.className = "scm-section";
    head.dataset.section = name.toLowerCase();
    if (hint) {
      head.title = hint;
    }

    const label = document.createElement("span");
    label.textContent = name;
    const number = document.createElement("span");
    number.className = "scm-section-count";
    number.textContent = String(count);
    head.append(label, number);
    return head;
  }

  private drawRow(row: ScmRow): HTMLElement {
    const entry = document.createElement("div");
    entry.className = "scm-entry";
    entry.dataset.module = row.module;
    entry.setAttribute("role", "listitem");
    if (this.showing?.kind === "row" && this.showing.module.toLowerCase() === row.module.toLowerCase()) {
      entry.classList.add("scm-entry-showing");
    }

    const tick = document.createElement("input");
    tick.type = "checkbox";
    tick.className = "scm-tick";
    tick.checked = !this.unticked.has(row.module.toLowerCase());
    tick.setAttribute("aria-label", `Commit ${row.module}`);
    tick.addEventListener("click", (event) => event.stopPropagation());
    tick.addEventListener("change", () => {
      if (tick.checked) {
        this.unticked.delete(row.module.toLowerCase());
      } else {
        this.unticked.add(row.module.toLowerCase());
      }
      this.enable();
    });

    const name = document.createElement("span");
    name.className = "scm-module";
    name.textContent = row.module;

    const status = document.createElement("span");
    status.className = `scm-status scm-status-${row.status}`;
    status.textContent = row.from ? `renamed from ${row.from}` : STATUS_WORD[row.status] ?? row.status;

    entry.append(tick, name, status);
    entry.title = `${row.module} (${row.kind}), ${row.file}. Click to compare with the branch head`;
    entry.addEventListener("click", () => void this.openRow(row.module));
    return entry;
  }

  private drawOutsideRow(row: ScmOutsideRow): HTMLElement {
    const entry = document.createElement("div");
    entry.className = "scm-entry scm-entry-still";
    entry.dataset.module = row.module;
    entry.dataset.outside = "true";
    entry.setAttribute("role", "listitem");

    const name = document.createElement("span");
    name.className = "scm-module";
    name.textContent = row.module;

    const status = document.createElement("span");
    status.className = `scm-status scm-status-${row.status}`;
    status.textContent = STATUS_WORD[row.status] ?? row.status;

    entry.append(name, status);
    entry.title = `${row.file}: ${STATUS_WORD[row.status] ?? row.status}`;
    return entry;
  }

  private drawCommit(commit: ScmCommit): HTMLElement {
    const box = document.createElement("div");
    box.className = "scm-commit";
    box.dataset.short = commit.short;
    const open = this.opened.has(commit.short);
    if (open) {
      box.classList.add("scm-commit-open");
    }

    const head = document.createElement("button");
    head.type = "button";
    head.className = "scm-commit-head";
    head.setAttribute("aria-expanded", open ? "true" : "false");

    const short = document.createElement("span");
    short.className = "scm-short";
    short.textContent = commit.short;

    const subject = document.createElement("span");
    subject.className = "scm-subject";
    subject.textContent = commit.subject;

    const author = document.createElement("span");
    author.className = "scm-author";
    author.textContent = commit.author;

    const clock = document.createElement("span");
    clock.className = "scm-when";
    clock.textContent = whenOf(commit.when);

    head.append(short, subject, author, clock);
    head.title = [commit.hash, `${commit.author} <${commit.email}>`, commit.when, "", commit.subject, commit.body]
      .filter((line, index) => index < 4 || line.length > 0)
      .join("\n")
      .trimEnd();
    head.addEventListener("click", () => {
      if (this.opened.has(commit.short)) {
        this.opened.delete(commit.short);
      } else {
        this.opened.add(commit.short);
      }
      this.draw();
    });

    const top = document.createElement("div");
    top.className = "scm-commit-top";
    top.appendChild(head);

    // The branch head's row carries Undo, the pane's amend: the head goes back one commit, its
    // changes return to the rows and its message to the box, and nothing is deleted. Only on the
    // commit that IS the head, and greyed with the reason when this pane cannot undo it - a first
    // commit, a merge, a commit the upstream already holds.
    const state = this.state;
    if (state && state.head !== "" && commit.hash === state.head) {
      const undo = document.createElement("button");
      undo.type = "button";
      undo.id = "scm-undo";
      undo.className = "scm-undo";
      undo.textContent = "Undo";
      undo.disabled = this.busy || state.undoBlocked !== "";
      undo.title = state.undoBlocked !== ""
        ? `Cannot undo: ${state.undoBlocked}`
        : "Take this commit back: the branch goes back one commit, its changes return to the rows, and its "
          + "message to the box. Nothing is deleted.";
      undo.addEventListener("click", (event) => {
        event.stopPropagation();
        void this.run({ action: "undo" });
      });
      top.appendChild(undo);
    }

    box.appendChild(top);

    if (open) {
      if (commit.files.length === 0) {
        const none = document.createElement("div");
        none.className = "scm-empty scm-commit-file";
        none.textContent = "No module files in this commit.";
        box.appendChild(none);
      }
      for (const file of commit.files) {
        box.appendChild(this.drawCommitFile(commit, file));
      }
    }

    return box;
  }

  private drawCommitFile(commit: ScmCommit, file: ScmCommitFile): HTMLElement {
    const entry = document.createElement("button");
    entry.type = "button";
    entry.className = "scm-entry scm-commit-file";
    entry.dataset.module = file.module;
    entry.dataset.short = commit.short;
    if (this.showing?.kind === "file"
      && this.showing.short === commit.short
      && this.showing.module.toLowerCase() === file.module.toLowerCase()) {
      entry.classList.add("scm-entry-showing");
    }

    const name = document.createElement("span");
    name.className = "scm-module";
    name.textContent = file.module || file.file;

    const status = document.createElement("span");
    status.className = `scm-status scm-status-${file.status}`;
    status.textContent = STATUS_WORD[file.status] ?? file.status;

    entry.append(name, status);
    entry.title = `${file.file} at ${commit.short}. Click to compare with the live module`;
    entry.addEventListener("click", () => void this.openFile(commit, file.module || file.file));
    return entry;
  }

  /**
   * The remote line of the ready state: the URL pushes go to, or the input that attaches one.
   * One verb behind both, so Change shows the same input filled in and Save re-points it.
   */
  private drawRemote(state: ScmStatus): HTMLElement {
    const box = document.createElement("div");
    box.className = "scm-state scm-remote";
    const acts = document.createElement("div");
    acts.className = "scm-state-acts";

    const button = (id: string, label: string, run: () => void): HTMLButtonElement => {
      const made = document.createElement("button");
      made.type = "button";
      made.id = id;
      made.className = "panel-empty-act";
      made.textContent = label;
      made.disabled = this.busy;
      made.addEventListener("click", run);
      return made;
    };

    const attached = state.remoteUrl !== "";
    if (attached && !this.remoteEditing) {
      const said = document.createElement("div");
      said.className = "scm-state-text";
      said.textContent = `${state.remote}: ${state.remoteUrl}`;
      acts.appendChild(button("scm-remote-change", "Change", () => {
        this.remoteEditing = true;
        this.draw();
      }));
      box.append(said, acts);
      return box;
    }

    const text = document.createElement("div");
    text.className = "scm-state-text";
    text.textContent = attached
      ? `Re-point ${state.remote}. The first push after that sets the upstream again.`
      : "No remote. Paste the URL of an empty repository on GitHub or elsewhere; the first push "
        + "sets the upstream, and git's credential manager signs you in.";
    const input = document.createElement("input");
    input.id = "scm-remote-url";
    input.type = "url";
    input.placeholder = "https://github.com/you/book.git";
    input.value = state.remoteUrl;
    const label = document.createElement("label");
    label.append("Remote URL", input);

    const attach = button("scm-remote-attach", attached ? "Save" : "Add remote", () => {
      const url = input.value.trim();
      if (!url) {
        this.setNotice("A remote needs a URL.", true);
        return;
      }
      void this.run({ action: "remote", url });
    });
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        attach.click();
      }
    });
    acts.appendChild(attach);
    if (attached) {
      acts.appendChild(button("scm-remote-cancel", "Cancel", () => {
        this.remoteEditing = false;
        this.draw();
      }));
    }

    box.append(text, label, acts);
    return box;
  }

  /**
   * The empty states: each says where the pane stands and offers the one control that moves it
   * on, in the words docs/source-control.md gives them.
   */
  private drawState(state: ScmStatus): HTMLElement {
    const box = document.createElement("div");
    box.className = "scm-state";
    box.dataset.state = state.state;

    const text = document.createElement("div");
    text.className = "scm-state-text";
    const acts = document.createElement("div");
    acts.className = "scm-state-acts";

    const button = (id: string, label: string, run: () => void): HTMLButtonElement => {
      const made = document.createElement("button");
      made.type = "button";
      made.id = id;
      made.className = "panel-empty-act";
      made.textContent = label;
      made.disabled = this.busy;
      made.addEventListener("click", run);
      return made;
    };

    switch (state.state) {
      case "noGit":
        text.textContent = "Git was not found. Install Git for Windows from git-scm.com/download/win, "
          + "then press Refresh. Its credential manager is what signs you in to a remote.";
        break;
      case "noProject":
        text.textContent = "No project is open.";
        break;
      case "unsaved":
        text.textContent = "Save the workbook first. An unsaved workbook has no path for a repository "
          + "folder to be remembered against.";
        break;
      case "noFolder": {
        text.textContent = "Choose the folder this project's modules are exported to. A repository is "
          + "initialised there, or the folder joins the repository above it.";
        acts.appendChild(button("scm-browse", "Choose folder", () => void this.run({ action: "browse" })));
        const offered = state.suggestedFolder || state.folder;
        if (offered) {
          const use = button("scm-use-folder", `Use ${offered}`, () => void this.run({ action: "settings", folder: offered }));
          use.title = offered;
          acts.appendChild(use);
        }
        break;
      }
      case "noRepository": {
        text.textContent = `${state.folder} is not inside a git repository.`;
        const init = button("scm-init", "Initialize", () => void this.run({ action: "init" }));
        init.title = "git init in the folder on a branch named main, with core.autocrlf off and a .gitignore "
          + "for the export's lock files";
        acts.appendChild(init);
        break;
      }
      case "noIdentity": {
        text.textContent = "Git needs to know who commits. The name and email are written to this "
          + "repository's configuration.";
        const name = document.createElement("input");
        name.id = "scm-identity-name";
        name.type = "text";
        name.autocomplete = "name";
        name.value = state.identity?.name ?? "";
        const nameLabel = document.createElement("label");
        nameLabel.append("Name", name);
        const email = document.createElement("input");
        email.id = "scm-identity-email";
        email.type = "email";
        email.autocomplete = "email";
        email.value = state.identity?.email ?? "";
        const emailLabel = document.createElement("label");
        emailLabel.append("Email", email);
        box.append(nameLabel, emailLabel);
        acts.appendChild(button("scm-identity-save", "Save", () => {
          const askedName = name.value.trim();
          const askedEmail = email.value.trim();
          if (!askedName || !askedEmail) {
            this.setNotice("Both the name and the email are needed.", true);
            return;
          }
          void this.run({ action: "identity", name: askedName, email: askedEmail });
        }));
        break;
      }
      case "conflicted": {
        text.textContent = "The folder has conflicts from a pull. Resolve them outside, or abort the "
          + "merge. Import is held until the folder is clean.";
        for (const file of state.conflicts) {
          const line = document.createElement("div");
          line.className = "scm-conflict";
          line.textContent = file;
          box.appendChild(line);
        }
        acts.appendChild(button("scm-abort", "Abort", () => void this.run({ action: "abort" })));
        break;
      }
      default:
        text.textContent = state.detail || `Source control is ${state.state}.`;
        break;
    }

    box.prepend(text);
    if (state.detail && state.state !== "noRepository") {
      const detail = document.createElement("div");
      detail.className = "scm-state-detail";
      detail.textContent = state.detail;
      box.appendChild(detail);
    }
    if (acts.childElementCount > 0) {
      box.appendChild(acts);
    }
    return box;
  }

  /**
   * A Changes row: the live module against the branch head, in the sync dialog's rows. Quiet
   * when it is a re-ask of what is already on screen, which stays up until the answer replaces it.
   */
  private async openRow(module: string, quiet = false): Promise<void> {
    if (!quiet) {
      this.showing = { kind: "row", module };
      this.shownRows = null;
      this.draw();
    }

    const answer = await this.ask(this.withProject({ action: "diff", module }));
    if (this.showing?.kind !== "row" || this.showing.module !== module) {
      return;
    }

    this.drawDiff(
      (answer.rows as SyncDiffLine[] | undefined) ?? [],
      `${module} against ${this.state?.branch || "HEAD"}`,
      asString(answer.error) || asString(answer.detail));
  }

  /** A commit's file: the text at that commit against the live module, with the ways out. */
  private openFile(commit: ScmCommit, module: string): Promise<void> {
    return this.openFileAt(commit.short, commit.hash, module);
  }

  private async openFileAt(short: string, hash: string, module: string, quiet = false): Promise<void> {
    if (!quiet) {
      this.showing = { kind: "file", short, hash, module };
      this.shownRows = null;
      this.draw();
    }

    const answer = await this.ask(this.withProject({ action: "show", module, ref: hash }));
    if (this.showing?.kind !== "file" || this.showing.short !== short || this.showing.module !== module) {
      return;
    }

    this.drawDiff(
      (answer.rows as SyncDiffLine[] | undefined) ?? [],
      `${module} @ ${short} against the live module`,
      asString(answer.error) || asString(answer.detail));
  }

  private drawDiff(rows?: SyncDiffLine[], title?: string, detail?: string): void {
    // A redraw with nothing new to show keeps what is on screen: the rows last answered for the
    // comparison being shown, until a click elsewhere or a fresh answer replaces them.
    if (rows === undefined && this.showing && this.shownRows) {
      rows = this.shownRows;
      title = this.shownTitle;
      detail = this.shownDetail;
    }

    this.diff.replaceChildren();

    if (rows === undefined) {
      const hint = document.createElement("div");
      hint.className = "scm-empty";
      const state = this.state;
      hint.textContent = !state || state.state !== "ready"
        ? ""
        : this.showing
          ? "Comparing..."
          : state.rows.length > 0 || this.commits.length > 0
            ? "Pick a change to compare it with the branch head, or a commit's file to compare it with the live module."
            : "";
      this.diff.appendChild(hint);
      return;
    }

    this.shownRows = rows;
    this.shownTitle = title ?? "";
    this.shownDetail = detail ?? "";

    const head = document.createElement("div");
    head.className = "scm-diff-head";

    const named = document.createElement("span");
    named.className = "scm-diff-title";
    named.textContent = title ?? "";
    head.appendChild(named);

    const showing = this.showing;
    if (showing?.kind === "file") {
      // The two ways out of a past version: read it whole in a tab of its own, or put it back.
      const open = document.createElement("button");
      open.type = "button";
      open.id = "scm-open-version";
      open.textContent = "Open this version";
      open.title = `${showing.module} @ ${showing.short} in a read-only tab`;
      open.addEventListener("click", () => void this.run({ action: "open", module: showing.module, ref: showing.hash }));

      const restore = document.createElement("button");
      restore.type = "button";
      restore.id = "scm-restore";
      restore.textContent = "Restore";
      restore.title = `Write ${showing.module}'s text at ${showing.short} into the module`;
      restore.addEventListener("click", () => this.confirmRestore(showing));

      head.append(open, restore);
    }

    this.diff.appendChild(head);

    const body = document.createElement("div");
    body.className = "scm-diff-body";
    if (rows.length === 0) {
      const hint = document.createElement("div");
      hint.className = "scm-empty";
      hint.textContent = detail || "Nothing differs.";
      body.appendChild(hint);
    } else {
      drawDiffRows(body, rows, "sync");
    }

    this.diff.appendChild(body);
  }

  /**
   * The card behind New branch...: a name, and Create. The branch is cut from the head and stays
   * on its commit, so nothing is imported and a dirty workbook is no bar; the card says so. A
   * refusal - a name git will not take, a name already taken - is shown in the card, with the
   * input live again for another go.
   */
  private askNewBranch(): void {
    if (this.busy || this.state?.state !== "ready") {
      return;
    }

    const from = this.state.branch || "the current commit";
    const at = this.state.head ? "at the same commit" : "before any commit";
    const { card, dismiss } = openModal({
      backdropId: "scm-branch-backdrop",
      cardId: "scm-branch-card",
      label: "New branch",
    });

    const title = document.createElement("div");
    title.id = "scm-branch-title";
    title.className = "modal-title";
    title.textContent = "New branch";

    const said = document.createElement("div");
    said.id = "scm-branch-consequence";
    said.className = "modal-detail";
    said.textContent = `Cut from ${from}, ${at}: the modules stay as they are, only the branch under `
      + "them is new. The first push sets its upstream.";
    card.setAttribute("aria-describedby", said.id);

    const input = document.createElement("input");
    input.id = "scm-branch-name";
    input.type = "text";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.placeholder = "feature/dates";
    const label = document.createElement("label");
    label.id = "scm-branch-label";
    label.append("Branch name", input);

    const error = document.createElement("div");
    error.id = "scm-branch-error";
    error.setAttribute("role", "alert");

    const buttons = document.createElement("div");
    buttons.className = "modal-buttons";

    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.id = "scm-branch-cancel";
    cancel.className = "modal-button";
    cancel.textContent = "Cancel";
    cancel.addEventListener("click", () => dismiss());

    const go = document.createElement("button");
    go.type = "button";
    go.id = "scm-branch-create";
    go.className = "modal-button primary";
    go.textContent = "Create";
    go.addEventListener("click", () => void this.runBranch({ input, go, cancel, error, dismiss }));
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        go.click();
      }
    });

    buttons.append(go, cancel);
    card.append(title, said, label, error, buttons);
    input.focus();
  }

  private async runBranch(card: {
    input: HTMLInputElement;
    go: HTMLButtonElement;
    cancel: HTMLButtonElement;
    error: HTMLElement;
    dismiss: () => void;
  }): Promise<void> {
    const name = card.input.value.trim();
    if (!name) {
      card.error.textContent = "A branch needs a name.";
      card.input.focus();
      return;
    }

    const live = (on: boolean): void => {
      card.go.disabled = !on;
      card.cancel.disabled = !on;
      card.input.disabled = !on;
      card.go.textContent = on ? "Create" : "Creating...";
    };
    live(false);
    card.error.textContent = "";

    // Asked directly rather than through run(), so a read the host's stamp started while the
    // card stood cannot swallow the answer: what comes back IS the status, and it is drawn here.
    let answer: Record<string, unknown>;
    try {
      answer = await this.ask(this.withProject({ action: "branch", name }));
    } catch (failed) {
      answer = { error: failed instanceof Error ? failed.message : String(failed) };
    }

    if (typeof answer.error === "string") {
      card.error.textContent = answer.error;
      live(true);
      card.input.focus();
      return;
    }

    const status = statusOf(answer);
    if (status) {
      this.adopt(status);
    }
    this.setNotice(asString(answer.detail), false);
    card.dismiss();
    this.draw();
  }

  /**
   * Asks first, does it, and says what happened - one card for the whole exchange, the Changes
   * pane's shape. A restore is recoverable by design: it lands as a change-log round labelled
   * with the commit, so the Changes pane can take it back.
   */
  private confirmRestore(showing: Showing & { kind: "file" }): void {
    if (this.busy) {
      return;
    }

    const question = `Restore ${showing.module} to ${showing.short}?`;
    const { card, dismiss } = openModal({
      backdropId: "scm-restore-backdrop",
      cardId: "scm-restore-card",
      label: question,
      role: "alertdialog",
    });

    const asked = document.createElement("div");
    asked.id = "scm-restore-question";
    asked.className = "modal-title";
    asked.textContent = question;

    const said = document.createElement("div");
    said.id = "scm-restore-consequence";
    said.className = "modal-detail";
    said.textContent = "The module's code goes back to how it stood at that commit. What stands now is "
      + "recorded in the Changes pane first, so this can itself be restored away. A form's design is "
      + "not restored, only its code.";
    card.setAttribute("aria-describedby", said.id);

    const buttons = document.createElement("div");
    buttons.className = "modal-buttons";

    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.id = "scm-restore-cancel";
    cancel.className = "modal-button";
    cancel.textContent = "Cancel";
    cancel.addEventListener("click", () => dismiss());

    const go = document.createElement("button");
    go.type = "button";
    go.id = "scm-restore-confirm";
    go.className = "modal-button primary";
    go.textContent = "Restore";
    go.addEventListener("click", () => {
      go.disabled = true;
      cancel.disabled = true;
      go.textContent = "Restoring...";
      void this.runRestore(showing, card, dismiss);
    });

    buttons.append(go, cancel);
    card.append(asked, said, buttons);
    cancel.focus();
  }

  private async runRestore(
    showing: Showing & { kind: "file" }, card: HTMLElement, dismiss: () => void): Promise<void> {
    let outcome: Record<string, unknown>;
    try {
      outcome = await this.ask(this.withProject({ action: "restore", module: showing.module, ref: showing.hash }));
    } catch (failed) {
      outcome = { error: failed instanceof Error ? failed.message : String(failed) };
    }

    // The rows are stale the moment the restore lands, so they are re-read before the outcome
    // is shown and the card sits over a list that already agrees with it.
    await this.run();

    card.replaceChildren();

    const summary = document.createElement("div");
    summary.id = "scm-restore-summary";
    summary.className = "modal-title";
    summary.setAttribute("role", "status");
    summary.textContent = asString(outcome.error) || asString(outcome.detail) || "Done.";

    const detail = document.createElement("div");
    detail.id = "scm-restore-outcome";
    const did = asString(outcome.did);
    const why = asString(outcome.why);
    detail.textContent = did ? (why ? `${showing.module}: ${did} - ${why}` : `${showing.module}: ${did}`) : "";

    const done = document.createElement("button");
    done.type = "button";
    done.id = "scm-restore-done";
    done.className = "modal-button primary";
    done.textContent = "Close";
    done.addEventListener("click", () => dismiss());

    const closing = document.createElement("div");
    closing.className = "modal-buttons";
    closing.appendChild(done);

    card.append(summary, detail, closing);
    done.focus();
  }

  /** The pane as the dev surface reads and drives it. */
  probe(): ScmPaneProbe {
    const press = (button: HTMLButtonElement | null): boolean => {
      if (!button || button.disabled) {
        return false;
      }

      button.click();
      return true;
    };

    return {
      state: () => ({
        project: this.state?.project ?? "",
        state: this.state?.state ?? "",
        detail: this.notice.textContent ?? "",
        folder: this.state?.folder ?? "",
        repository: this.state?.repository ?? "",
        branch: this.state?.branch ?? "",
        upstream: this.state?.upstream
          ? { name: this.state.upstream, ahead: this.state.ahead, behind: this.state.behind }
          : null,
        branches: (this.state?.branches ?? []).map((one) => (one.remote ? `${one.remote}/${one.name}` : one.name)),
        head: this.state?.head ?? "",
        undoable: (() => {
          const undo = this.list.querySelector<HTMLButtonElement>("#scm-undo");
          return undo !== null && !undo.disabled;
        })(),
        undoBlocked: this.state?.undoBlocked ?? "",
        dirty: this.state?.dirty ?? false,
        files: [...this.file.options].map((one) => one.value),
        rows: (this.state?.rows ?? []).map((row) => ({ ...row })),
        ticked: this.tickedModules(),
        outside: (this.state?.outside ?? []).map((row) => ({ ...row })),
        commits: this.commits.map((commit) => ({
          hash: commit.hash,
          short: commit.short,
          author: commit.author,
          when: commit.when,
          subject: commit.subject,
          files: commit.files.map((file) => ({ ...file })),
          open: this.opened.has(commit.short),
        })),
        conflicts: [...(this.state?.conflicts ?? [])],
        lastCommit: this.state?.lastCommit?.short ?? null,
        suggestedMessage: this.state?.suggestedMessage ?? "",
        showing: this.showing === null
          ? null
          : this.showing.kind === "row"
            ? this.showing.module
            : `${this.showing.module}@${this.showing.short}`,
        message: this.message.value,
        busy: this.busy,
        behind: this.hostStamp > this.drawnStamp,
        blameOn: this.blame.on(),
        remote: this.state?.remote ?? "",
        remoteUrl: this.state?.remoteUrl ?? "",
        listWidth: Math.round(this.list.getBoundingClientRect().width),
      }),
      press: (control) => {
        switch (control) {
          case "refresh": return press(this.refresh);
          case "commit": return press(this.commit);
          case "undo": return press(this.list.querySelector<HTMLButtonElement>("#scm-undo"));
          case "export": return press(this.exportButton);
          case "import": return press(this.importButton);
          case "fetch": return press(this.fetch);
          case "pull": return press(this.pull);
          case "push": return press(this.push);
          case "blame": return press(this.blameButton);
          case "init": return press(this.list.querySelector<HTMLButtonElement>("#scm-init"));
          case "abort": return press(this.list.querySelector<HTMLButtonElement>("#scm-abort"));
          case "browse": return press(this.list.querySelector<HTMLButtonElement>("#scm-browse"));
          case "use": return press(this.list.querySelector<HTMLButtonElement>("#scm-use-folder"));
          case "identity": return press(this.list.querySelector<HTMLButtonElement>("#scm-identity-save"));
          case "remote": return press(this.list.querySelector<HTMLButtonElement>("#scm-remote-attach"));
          case "open": return press(this.diff.querySelector<HTMLButtonElement>("#scm-open-version"));
          case "restore": {
            // The real gesture: the head's button, then the confirm the card raises.
            if (!press(this.diff.querySelector<HTMLButtonElement>("#scm-restore"))) {
              return false;
            }
            return press(document.querySelector<HTMLButtonElement>("#scm-restore-confirm"));
          }
          default:
            return false;
        }
      },
      chooseFile: (name) => {
        const option = [...this.file.options].find((one) => one.value.toLowerCase() === name.toLowerCase());
        if (!option) {
          return false;
        }

        this.file.value = option.value;
        this.file.dispatchEvent(new Event("change"));
        return true;
      },
      tick: (module, on) => {
        const box = this.list.querySelector<HTMLInputElement>(
          `.scm-entry[data-module="${CSS.escape(module)}"] > .scm-tick`);
        if (!box) {
          return false;
        }

        if (box.checked !== on) {
          box.checked = on;
          box.dispatchEvent(new Event("change"));
        }
        return true;
      },
      setMessage: (text) => {
        if (this.message.hidden) {
          return false;
        }

        this.message.value = text;
        this.message.dispatchEvent(new Event("input"));
        return true;
      },
      show: (module) => {
        const row = this.list.querySelector<HTMLElement>(
          `.scm-entry[data-module="${CSS.escape(module)}"]:not([data-outside])`);
        if (!row || row.classList.contains("scm-commit-file")) {
          return false;
        }

        row.click();
        return true;
      },
      openCommit: (short) => {
        const head = this.list.querySelector<HTMLButtonElement>(
          `.scm-commit[data-short="${CSS.escape(short)}"] > .scm-commit-top > .scm-commit-head`);
        head?.click();
        return head !== null;
      },
      showFile: (short, module) => {
        const file = this.list.querySelector<HTMLButtonElement>(
          `.scm-commit-file[data-short="${CSS.escape(short)}"][data-module="${CSS.escape(module)}"]`);
        file?.click();
        return file !== null;
      },
      chooseBranch: (name) => {
        const option = [...this.branch.options].find((one) => one.value === name);
        if (!option || this.branch.disabled || this.branch.hidden) {
          return false;
        }

        this.branch.value = option.value;
        this.branch.dispatchEvent(new Event("change"));
        return true;
      },
      createBranch: (name) => {
        if (this.branch.disabled || this.branch.hidden) {
          return false;
        }

        // The real gesture: the select's last entry, then the card it raises.
        this.branch.value = NEW_BRANCH;
        this.branch.dispatchEvent(new Event("change"));
        const input = document.querySelector<HTMLInputElement>("#scm-branch-name");
        const create = document.querySelector<HTMLButtonElement>("#scm-branch-create");
        if (!input || !create) {
          return false;
        }

        input.value = name;
        create.click();
        return true;
      },
      setIdentity: (name, email) => {
        const nameBox = this.list.querySelector<HTMLInputElement>("#scm-identity-name");
        const emailBox = this.list.querySelector<HTMLInputElement>("#scm-identity-email");
        if (!nameBox || !emailBox) {
          return false;
        }

        nameBox.value = name;
        emailBox.value = email;
        return true;
      },
      setRemoteUrl: (url) => {
        let box = this.list.querySelector<HTMLInputElement>("#scm-remote-url");
        if (!box && this.state?.state === "ready" && !this.busy) {
          // An attached remote folds the input behind Change; typing is the gesture that unfolds it.
          this.remoteEditing = true;
          this.draw();
          box = this.list.querySelector<HTMLInputElement>("#scm-remote-url");
        }
        if (!box) {
          return false;
        }

        box.value = url;
        return true;
      },
      resizeList: (width) => {
        if (!Number.isFinite(width) || this.body.clientWidth <= 0) {
          return false;
        }

        this.placeList(width);
        this.rememberListWidth();
        return true;
      },
    };
  }
}
