/*
 * The Changes pane: what happened to this project's module code, by whom, in rounds.
 *
 * IT ONLY SHOWS. There is no revert button here and there is not going to be one. Putting text
 * back is a write, and a pane that writes is a pane that can lose work - so the old text is made
 * REACHABLE instead: read the diff, open what a module held before, copy what you want, and make
 * the change yourself in the editor where normal undo protects you. An agent does the same thing
 * through the write route, and its undo lands in this log like any other round.
 *
 * IT PULLS RATHER THAN BEING PUSHED. Every count in here is a comparison of two whole texts, and
 * this product has learned twice what that costs when it happens per keystroke. So the host is
 * asked when the pane is opened and when the developer asks again - never on the write path.
 */

import { drawDiffRows, type SyncDiffLine } from "./diffview.js";

/** One module's before and after within a round, as the host reports it. */
export interface ChangeEntry {
  module: string;
  kind: string;
  added: number;
  removed: number;
  before: string | null;
  after: string | null;
  held: boolean;
}

/** One round, as the host reports it. */
export interface ChangeRound {
  round: number;
  started: string;
  ended: string;
  by: string;
  label: string | null;
  open: boolean;
  accepted: boolean;
  entries: ChangeEntry[];
}

/** The change log's whole answer. */
export interface ChangeLogState {
  detail: string;
  project: string;
  directory: string;
  acceptedAt: number;
  covers: string;
  rounds: ChangeRound[];
}

/** How the pane reaches the host. One function, because there is only one kind of request. */
export type ChangesRequest = (args: Record<string, string>) => Promise<Record<string, unknown>>;

/** What the pane can be driven and read through, for the dev surface. */
export interface ChangesPaneProbe {
  state(): {
    project: string;
    acceptedAt: number;
    covers: string;
    busy: boolean;
    rounds: {
      round: number;
      by: string;
      label: string | null;
      open: boolean;
      accepted: boolean;
      modules: { module: string; added: number; removed: number; held: boolean }[];
    }[];
    showing: string | null;
  };
  /** Presses a named control: refresh, snapshot, accept. False when unknown. */
  press(control: string): boolean;
  /** Opens one round's module diff, as clicking its row does. */
  show(round: number, module: string): boolean;
}

const KIND_WORD: Record<string, string> = {
  written: "changed",
  added: "added",
  removed: "removed",
};

/** "2:04 PM", or the date as well once it is not today. */
function when(iso: string): string {
  const at = new Date(iso);
  if (Number.isNaN(at.getTime())) {
    return "";
  }

  const today = new Date();
  const sameDay = at.toDateString() === today.toDateString();
  return sameDay
    ? at.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
    : at.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}

let livePane: ChangesPaneProbe | null = null;

/** The Changes pane's probe, or null before one has been built. */
export const changesPaneProbe = (): ChangesPaneProbe | null => livePane;

export class ChangesPane {
  private readonly list: HTMLElement;
  private readonly diff: HTMLElement;
  private readonly covers: HTMLElement;
  private readonly title: HTMLElement;
  private readonly refresh: HTMLButtonElement;
  private readonly snapshot: HTMLButtonElement;
  private readonly accept: HTMLButtonElement;

  private state: ChangeLogState | null = null;
  private showing: { round: number; module: string } | null = null;
  private busy = false;

  constructor(root: HTMLElement, private readonly ask: ChangesRequest) {
    this.list = root.querySelector("#changes-list") as HTMLElement;
    this.diff = root.querySelector("#changes-diff") as HTMLElement;
    this.covers = root.querySelector("#changes-covers") as HTMLElement;
    this.title = root.querySelector("#changes-project") as HTMLElement;
    this.refresh = root.querySelector("#changes-refresh") as HTMLButtonElement;
    this.snapshot = root.querySelector("#changes-snapshot") as HTMLButtonElement;
    this.accept = root.querySelector("#changes-accept") as HTMLButtonElement;

    livePane = this.probe();

    this.refresh.addEventListener("click", () => void this.reload());
    this.snapshot.addEventListener("click", () => void this.reload({ action: "snapshot" }));
    this.accept.addEventListener("click", () => void this.reload({ action: "accept" }));

    this.draw();
  }

  /** Asked for when the pane is opened, which is the only time any of this costs anything. */
  shown(): void {
    void this.reload();
  }

  private async reload(args: Record<string, string> = {}): Promise<void> {
    if (this.busy) {
      return;
    }

    this.busy = true;
    this.setBusy(true);
    try {
      const answer = await this.ask(args);
      this.state = answer as unknown as ChangeLogState;

      // A round the developer just closed is gone from the list as a running one, so a diff
      // opened from it would be pointing at a row that has moved.
      if (args.action) {
        this.showing = null;
      }
    } finally {
      this.busy = false;
      this.setBusy(false);
      this.draw();
    }
  }

  private setBusy(on: boolean): void {
    this.refresh.disabled = on;
    this.snapshot.disabled = on;
    this.accept.disabled = on;
  }

  private draw(): void {
    const state = this.state;
    this.title.textContent = state?.project ?? "";
    this.covers.textContent = state?.covers ?? "";

    this.list.replaceChildren();

    if (!state || state.rounds.length === 0) {
      const empty = document.createElement("div");
      empty.className = "changes-empty";
      empty.textContent = state
        ? "Nothing has been written to this project's modules yet."
        : "Reading the change log...";
      this.list.appendChild(empty);
      this.drawDiff();
      return;
    }

    for (const round of state.rounds) {
      this.list.appendChild(this.drawRound(round));

      // The accepted line, drawn where it falls rather than as a property of a round: everything
      // above it is what has happened since the developer last said yes.
      if (round.round === state.acceptedAt) {
        const line = document.createElement("div");
        line.className = "changes-accepted";
        line.textContent = "accepted";
        this.list.appendChild(line);
      }
    }

    this.drawDiff();
  }

  private drawRound(round: ChangeRound): HTMLElement {
    const box = document.createElement("div");
    box.className = "changes-round";
    box.dataset.round = String(round.round);

    const head = document.createElement("div");
    head.className = "changes-round-head";

    const who = document.createElement("span");
    who.className = `changes-by changes-by-${round.by.toLowerCase() === "developer" ? "developer" : "agent"}`;
    who.textContent = round.by;

    const said = document.createElement("span");
    said.className = "changes-label";
    said.textContent = round.label ?? (round.open ? "still writing" : `${round.entries.length} module(s)`);

    const clock = document.createElement("span");
    clock.className = "changes-when";
    clock.textContent = when(round.ended || round.started);
    clock.title = `Round ${round.round}, ${new Date(round.started).toLocaleString()}`;

    head.append(who, said, clock);
    box.appendChild(head);

    for (const entry of round.entries) {
      box.appendChild(this.drawEntry(round, entry));
    }

    return box;
  }

  private drawEntry(round: ChangeRound, entry: ChangeEntry): HTMLElement {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "changes-entry";
    row.dataset.module = entry.module;
    if (this.showing?.round === round.round && this.showing.module === entry.module) {
      row.classList.add("changes-entry-showing");
    }

    const name = document.createElement("span");
    name.className = "changes-module";
    name.textContent = entry.module;

    const what = document.createElement("span");
    what.className = "changes-kind";
    what.textContent = KIND_WORD[entry.kind] ?? entry.kind;

    const counts = document.createElement("span");
    counts.className = "changes-counts";
    if (entry.held) {
      const added = document.createElement("span");
      added.className = "changes-added";
      added.textContent = `+${entry.added}`;
      const removed = document.createElement("span");
      removed.className = "changes-removed";
      removed.textContent = `-${entry.removed}`;
      counts.append(added, removed);
    } else {
      // A text the log has let go. Saying so is the difference between a log that ages and one
      // that quietly lies about what it can still show.
      counts.textContent = "text let go";
      counts.classList.add("changes-gone");
    }

    row.append(name, what, counts);
    row.addEventListener("click", () => void this.open(round.round, entry.module));
    return row;
  }

  /**
   * Asks the host to line one module's change up, and draws it.
   *
   * The comparison is the HOST's, which is the same one the import dialog uses and the same one
   * that knows to answer a hopelessly different middle as a block instead of building a table
   * nothing can afford. A second implementation in the page would not know that.
   */
  private async open(round: number, module: string): Promise<void> {
    this.showing = { round, module };
    this.draw();

    const answer = await this.ask({ action: "diff", round: String(round), module });

    if (this.showing?.round !== round || this.showing.module !== module) {
      return;
    }

    this.drawDiff(
      (answer.rows as SyncDiffLine[] | undefined) ?? [],
      `${module}, round ${round}`,
      (answer.detail as string | undefined) ?? "");
  }

  private drawDiff(rows?: SyncDiffLine[], title?: string, detail?: string): void {
    this.diff.replaceChildren();

    if (rows === undefined) {
      const hint = document.createElement("div");
      hint.className = "changes-empty";
      hint.textContent = this.state?.rounds.length ? "Choose a module to see what changed." : "";
      this.diff.appendChild(hint);
      return;
    }

    const head = document.createElement("div");
    head.className = "changes-diff-head";
    head.textContent = title ?? "";
    this.diff.appendChild(head);

    const body = document.createElement("div");
    body.className = "changes-diff-body";
    if (rows.length === 0) {
      const hint = document.createElement("div");
      hint.className = "changes-empty";
      hint.textContent = detail && detail !== "held" ? detail : "Nothing on either side.";
      body.appendChild(hint);
    } else {
      drawDiffRows(body, rows, "sync");
    }

    this.diff.appendChild(body);
  }

  /** The pane as the dev surface reads and drives it. */
  probe(): ChangesPaneProbe {
    return {
      state: () => ({
        project: this.state?.project ?? "",
        acceptedAt: this.state?.acceptedAt ?? 0,
        covers: this.state?.covers ?? "",
        busy: this.busy,
        rounds: (this.state?.rounds ?? []).map((round) => ({
          round: round.round,
          by: round.by,
          label: round.label,
          open: round.open,
          accepted: round.accepted,
          modules: round.entries.map((entry) => ({
            module: entry.module,
            added: entry.added,
            removed: entry.removed,
            held: entry.held,
          })),
        })),
        showing: this.showing ? `${this.showing.module}@${this.showing.round}` : null,
      }),
      press: (control) => {
        const button = control === "refresh" ? this.refresh
          : control === "snapshot" ? this.snapshot
          : control === "accept" ? this.accept
          : null;
        button?.click();
        return button !== null;
      },
      show: (round, module) => {
        const row = this.list.querySelector<HTMLElement>(
          `.changes-round[data-round="${round}"] .changes-entry[data-module="${CSS.escape(module)}"]`);
        row?.click();
        return row !== null;
      },
    };
  }
}
