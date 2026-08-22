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
import { openModal } from "./modal.js";

/** One module's before and after within a round, as the host reports it. */
export interface ChangeEntry {
  module: string;
  kind: string;
  added: number;
  removed: number;
  /** What it was called when the round began, when a rename moved it. */
  from: string | null;
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
  acceptedAt: number;
  covers: string;
  rounds: ChangeRound[];
}

/** How the pane reaches the host. One function, because there is only one kind of request. */
export type ChangesRequest = (args: Record<string, string>) => Promise<Record<string, unknown>>;

/** The files this session has open, and the one the developer is looking at. */
export type OpenFiles = () => { names: string[]; current: string | null };

/** What the pane can be driven and read through, for the dev surface. */
export interface ChangesPaneProbe {
  state(): {
    project: string;
    files: string[];
    acceptedAt: number;
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
    /** Whether the comparison is up full size. */
    full: boolean;
    /** Snapshots the full-size card's own rail offers. Zero when it is down or holds one. */
    fullChoices: number;
    /** Whether the rail is up, and how wide it was left. */
    railUp: boolean;
    railWidth: number;
    /** Whether the host has recorded changes since these counts were read. */
    behind: boolean;
  };
  /** Presses a named control: refresh, snapshot, accept. False when unknown. */
  press(control: string): boolean;
  /**
   * Opens one round's module diff, as clicking its row does. `where` says WHICH row: the pane's
   * list, or the full-size card's rail - two controls onto the same comparison, and a harness
   * proving one has not proved the other.
   */
  show(round: number, module: string, where?: "pane" | "full", gesture?: "click" | "double"): boolean;
  /** Points the pane at another open file, through the select's own change event. */
  chooseFile(name: string): boolean;
  /** Opens the comparison on screen full size, or closes it. */
  expand(open: boolean): boolean;
}

const KIND_WORD: Record<string, string> = {
  written: "changed",
  added: "added",
  removed: "removed",
  renamed: "renamed",
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
  private readonly title: HTMLElement;
  private readonly refresh: HTMLButtonElement;
  private readonly snapshot: HTMLButtonElement;
  private readonly accept: HTMLButtonElement;

  private readonly file: HTMLSelectElement;
  private readonly newer: HTMLElement;

  /*
   * WHAT THE COUNTS ARE, AND WHAT THEY ARE NOT.
   *
   * Every number in this pane is a comparison of two whole module texts, so it is worked out when
   * the pane is opened and when the developer asks again - never on the write path. That rule is
   * not negotiable; it is why this pane exists in the shape it does.
   *
   * The cost of it is that the numbers age, and they aged silently: a reading of +54 sat beside a
   * module the editor had already grown to 61 lines, with nothing on screen saying which of the
   * two was current (the owner, 2026-08-22: "shouldn't + number align with number of lines?").
   * They did align - with the module as it was when the pane last looked.
   *
   * So the host taps the pane on the shoulder with a bare integer whenever it records a change,
   * and the pane says "newer changes" if that integer has moved past the one it drew. No counting,
   * no re-read, no text: the pull-only rule holds, and the pane stops implying it is live.
   */
  private drawnStamp = 0;
  private hostStamp = 0;

  private state: ChangeLogState | null = null;
  private showing: { round: number; module: string } | null = null;
  private busy = false;

  /** What the select was last built from, so an unchanged session leaves an open popup alone. */
  private fileSignature = "";

  /** The comparison on screen, kept so it can be shown full size without asking again. */
  private showingRows: SyncDiffLine[] = [];
  private showingTitle = "";
  private full: {
    card: HTMLElement;
    dismiss: () => void;
    title: HTMLElement;
    split: HTMLElement;
    rail: HTMLElement;
    splitter: HTMLElement;
    toggle: HTMLButtonElement;
    body: HTMLElement;
  } | null = null;

  /** How wide the rail was left, and whether it was left up. Kept across opens in this session. */
  private railWidth = 200;
  private railHidden = false;

  constructor(root: HTMLElement, private readonly ask: ChangesRequest, private readonly files: OpenFiles) {
    this.list = root.querySelector("#changes-list") as HTMLElement;
    this.diff = root.querySelector("#changes-diff") as HTMLElement;
    this.title = root.querySelector("#changes-project") as HTMLElement;
    this.refresh = root.querySelector("#changes-refresh") as HTMLButtonElement;
    this.snapshot = root.querySelector("#changes-snapshot") as HTMLButtonElement;
    this.accept = root.querySelector("#changes-accept") as HTMLButtonElement;

    this.file = root.querySelector("#changes-file") as HTMLSelectElement;
    this.newer = root.querySelector("#changes-newer") as HTMLElement;

    livePane = this.probe();

    // A FILE CHOICE IS A DIFFERENT LOG, not a filter over one: every project keeps its own. So
    // what is on screen is dropped first - a comparison from the old file's round would be read
    // as belonging to the new one.
    this.file.addEventListener("change", () => {
      this.showing = null;
      this.state = null;
      void this.reload();
    });

    this.filesChanged();

    this.refresh.addEventListener("click", () => void this.reload());
    this.snapshot.addEventListener("click", () => void this.reload({ action: "snapshot" }));
    this.accept.addEventListener("click", () => void this.reload({ action: "accept" }));

    this.draw();
  }

  /**
   * The host recorded a change. Nothing is re-read - see the note by `drawnStamp`.
   *
   * The DOM is touched only on the edge, when the pane goes from current to behind, because this
   * arrives on every write and a class set on every one of them is work for no change on screen.
   */
  stamped(stamp: number): void {
    if (stamp <= this.hostStamp) {
      return;
    }

    const wasBehind = this.hostStamp > this.drawnStamp;
    this.hostStamp = stamp;
    if (!wasBehind) {
      this.showNewer();
    }
  }

  private showNewer(): void {
    const behind = this.hostStamp > this.drawnStamp;
    this.newer.hidden = !behind;
    this.refresh.classList.toggle("changes-refresh-newer", behind);
    this.refresh.title = behind
      ? "Read the log again - it has moved on since these counts were taken."
      : "Read the log again. These counts are a reading, not a live total.";
  }

  /** Asked for when the pane is opened, which is the only time any of this costs anything. */
  shown(): void {
    this.filesChanged();
    void this.reload();
  }

  /**
   * Rebuilds the file list from the session as it stands.
   *
   * A CLOSED FILE DROPS OUT. A workbook nobody has open is not something this pane offers, and if
   * it was the one being shown the pane falls back to the file the developer is actually in rather
   * than going on answering about a file that is not there. The session lets go of its log at the
   * same moment, so opening the workbook again starts fresh.
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

    // One file is not a choice, the same rule the list panes' file select follows - and the
    // name beside it goes the other way, so the workbook is said exactly once either way.
    this.file.hidden = names.length < 2;
    this.title.hidden = !this.file.hidden;
    this.file.value = keep ? chosen : current ?? names[0] ?? "";

    // Read back, so what is drawn and what is sent cannot disagree: assigning a value the list
    // does not hold leaves a select showing its first option and reports nothing about it.
    if (this.file.value !== chosen) {
      this.showing = null;
      this.state = null;
      void this.reload();
    }
  }

  private async reload(args: Record<string, string> = {}): Promise<void> {
    if (this.busy) {
      return;
    }

    this.busy = true;
    this.setBusy(true);

    // Stamped BEFORE the ask, not after: a change recorded while the request is in flight is one
    // this answer may or may not carry, and claiming it was included would be the pane going
    // quiet about something it might have missed. Erring towards "there is newer" is the safe
    // direction - the cost of a needless refresh is one read, and the cost of a missed one is a
    // number the developer trusts.
    const asOf = this.hostStamp;
    try {
      const answer = await this.ask(
        this.file.value ? { project: this.file.value, ...args } : args);
      this.state = answer as unknown as ChangeLogState;
      this.drawnStamp = asOf;

      // A round the developer just closed is gone from the list as a running one, so a diff
      // opened from it would be pointing at a row that has moved.
      if (args.action) {
        this.showing = null;
      }
    } finally {
      this.showNewer();
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

    this.list.replaceChildren();

    if (!state || state.rounds.length === 0) {
      const empty = document.createElement("div");
      empty.className = "changes-empty";
      empty.textContent = state
        ? "Nothing has been written to this project's modules yet."
        : "Reading the change log...";
      this.list.appendChild(empty);
      this.drawFullList();
      this.drawDiff();
      return;
    }

    // ABOVE THE NEWEST REVIEWED ROUND, not under it. The list runs newest first and `acceptedAt`
    // names the newest round that WAS reviewed, so the mark belongs in front of it: everything
    // below the line has been seen, everything above it has not. Drawn after that round instead,
    // the round itself sat on the unreviewed side - so pressing Accept, which reviews everything,
    // left the mark one row down from the top rather than at it (the owner spotted it, 2026-08-22).
    let marked = state.acceptedAt <= 0;
    for (const round of state.rounds) {
      if (!marked && round.round <= state.acceptedAt) {
        marked = true;
        const line = document.createElement("div");
        line.className = "changes-accepted";
        line.textContent = "accepted";
        this.list.appendChild(line);
      }

      this.list.appendChild(this.drawRound(round));
    }

    // The rail is the same list said shorter, so it is rebuilt from the same pass. One list drawn
    // twice by one rule cannot drift; two lists maintained separately always do.
    this.drawFullList();
    this.drawDiff();
  }

  private drawRound(round: ChangeRound): HTMLElement {
    const box = document.createElement("div");
    box.className = "changes-round";
    box.dataset.round = String(round.round);

    const head = document.createElement("div");
    head.className = "changes-round-head";

    // The round's own number, because everything else names it by that: the api's `round=`, the
    // label a driver reads back, and a developer asking an agent to look at "round 4".
    const number = document.createElement("span");
    number.className = "changes-number";
    number.textContent = String(round.round);

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

    head.append(number, who, said, clock);
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
    // A rename says what it was, because "renamed" on its own leaves the reader looking for a
    // module under a name that is no longer in the tree.
    what.textContent = entry.from
      ? `renamed from ${entry.from}`
      : KIND_WORD[entry.kind] ?? entry.kind;

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
    // A row is a thing you open, and opening a thing properly is a double-click - the same
    // gesture the tree uses to open a module (the owner, 2026-08-22: "double click on row should
    // open popout"). Single click still puts it in the strip, which is the cheap look.
    row.title = `${entry.module}, round ${round.round}. Double-click to see it full size`;
    row.addEventListener("click", () => void this.open(round.round, entry.module));
    row.addEventListener("dblclick", () => this.openFull(round.round, entry.module));
    return row;
  }

  /**
   * Asks the host to line one module's change up, and draws it.
   *
   * The comparison is the HOST's, which is the same one the import dialog uses and the same one
   * that knows to answer a hopelessly different middle as a block instead of building a table
   * nothing can afford. A second implementation in the page would not know that.
   */
  /**
   * The double-click: this row, full size.
   *
   * The first click of the double has already asked for the row, so when its rows are here this
   * is just the expand. When they are not - a double-click on a row nobody had opened yet - the
   * request is in flight, and asking again would be a second round trip over a whole-module
   * comparison to get an answer already on its way. So the intent is REMEMBERED instead, and the
   * draw that lands the rows honours it.
   */
  private openFull(round: number, module: string): void {
    if (this.showing?.round === round
      && this.showing.module === module
      && this.showingRows.length > 0) {
      this.expand();
      return;
    }

    this.growWhenDrawn = { round, module };
    if (this.showing?.round !== round || this.showing.module !== module) {
      void this.open(round, module);
    }
  }

  /** A row asked for full size before its comparison had arrived. */
  private growWhenDrawn: { round: number; module: string } | null = null;

  private async open(round: number, module: string): Promise<void> {
    // A pending full-size intent belongs to ONE row. Opening a different one is the developer
    // having moved on, and a flag left standing would pop that later row open by itself.
    if (this.growWhenDrawn
      && (this.growWhenDrawn.round !== round || this.growWhenDrawn.module !== module)) {
      this.growWhenDrawn = null;
    }

    this.showing = { round, module };
    this.draw();

    const asking = { action: "diff", round: String(round), module };
    const answer = await this.ask(
      this.file.value ? { project: this.file.value, ...asking } : asking);

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

    const named = document.createElement("span");
    named.textContent = title ?? "";
    head.appendChild(named);

    // FULL SIZE, because the pane is a strip along the bottom and a comparison is two columns of
    // code. Anything past a few lines is read three words at a time down there.
    if (rows.length > 0) {
      const grow = document.createElement("button");
      grow.type = "button";
      grow.id = "changes-expand";
      grow.className = "changes-grow";
      grow.title = "Show this comparison full size";
      grow.setAttribute("aria-label", "Show this comparison full size");
      grow.innerHTML = '<span class="codicon codicon-screen-full" aria-hidden="true"></span>';
      grow.addEventListener("click", () => this.expand());
      head.appendChild(grow);
    }

    this.diff.appendChild(head);

    this.showingRows = rows;
    this.showingTitle = title ?? "";

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

    // And the card, when one is up, from the rows the strip was just given.
    this.drawFullDiff();

    // A double-click that landed before its comparison did. Honoured only for the row that was
    // actually asked for: a second click elsewhere while the first was in flight has moved on,
    // and popping open whatever arrived next would be the pane answering a question nobody asked.
    const waiting = this.growWhenDrawn;
    if (waiting
      && this.showing?.round === waiting.round
      && this.showing.module === waiting.module
      && rows.length > 0) {
      this.growWhenDrawn = null;
      this.expand();
    }
  }

  /**
   * Shows the comparison on screen full size, over everything.
   *
   * The pane is a strip along the bottom and a comparison is two columns of code, so anything
   * past a few lines is read three words at a time down there. This is the same rows, drawn by
   * the same renderer, in a card that fills the window.
   *
   * IT CARRIES THE SNAPSHOTS WITH IT. Opening one comparison full size and having to close it to
   * reach the next one is the dialog asking the reader to hold the list in their head; the rail
   * down the left is that list, kept where it can be pointed at (the owner, 2026-08-22: "inside
   * the modal, I'd like a way to change the various snapshots too"). Compact, because the code is
   * what came here to be read.
   */
  expand(): void {
    if (this.full || this.showingRows.length === 0) {
      return;
    }

    const { card, dismiss } = openModal({
      backdropId: "changes-full-backdrop",
      cardId: "changes-full-card",
      label: `What changed in ${this.showingTitle}`,
      closed: () => {
        this.full = null;
      },
    });

    const head = document.createElement("div");
    head.id = "changes-full-head";

    // TOP LEFT, where a panel toggle lives - not out on the divider, where it was a 15px target
    // hunting for a chevron (the owner, 2026-08-22: "I'd prefer a button at the top left").
    const toggle = document.createElement("button");
    toggle.type = "button";
    toggle.id = "changes-full-toggle";
    toggle.className = "toolbar-button";
    toggle.setAttribute("aria-controls", "changes-full-list");
    toggle.innerHTML = '<span class="codicon codicon-list-flat" aria-hidden="true"></span>';
    toggle.addEventListener("click", () => this.showRail(this.railHidden));

    const named = document.createElement("span");
    named.id = "changes-full-title";

    const close = document.createElement("button");
    close.type = "button";
    close.id = "changes-full-close";
    close.title = "Close (Esc)";
    close.setAttribute("aria-label", "Close");
    close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';
    close.addEventListener("click", () => dismiss());

    head.append(toggle, named, close);

    const split = document.createElement("div");
    split.id = "changes-full-split";

    const rail = document.createElement("div");
    rail.id = "changes-full-list";
    rail.setAttribute("role", "listbox");
    rail.setAttribute("aria-label", "Snapshots");
    rail.style.flex = `0 0 ${this.railWidth}px`;
    rail.addEventListener("keydown", (event) => this.railKey(event));

    // The divider between them, the same one the designer's two halves are separated by: a grip to
    // say it can be dragged, and the arrow keys for anyone who cannot drag. It resizes and nothing
    // else now - what put the rail away moved to the head, where a panel toggle belongs.
    const splitter = document.createElement("div");
    splitter.id = "changes-full-splitter";
    splitter.setAttribute("role", "separator");
    splitter.setAttribute("aria-orientation", "vertical");
    splitter.setAttribute("aria-label", "Resize the snapshot list");
    splitter.title = "Drag to resize the snapshots";
    splitter.tabIndex = 0;

    const grip = document.createElement("div");
    grip.id = "changes-full-grip";
    splitter.appendChild(grip);

    const body = document.createElement("div");
    body.id = "changes-full-diff";

    split.append(rail, splitter, body);
    card.append(head, split);

    this.full = { card, dismiss, title: named, split, rail, splitter, toggle, body };
    this.dragRail(splitter);
    this.showRail(!this.railHidden);
    this.drawFullList();
    this.drawFullDiff();
  }

  /**
   * Puts the rail away, or brings it back.
   *
   * The button STAYS either way, in the corner it was pressed in - it is the way home, and a
   * control that removes the only thing which could undo it is a trapdoor. Pressed while the rail
   * is up, which is how a toggle in this product says it is on.
   */
  private showRail(up: boolean): void {
    this.railHidden = !up;
    if (!this.full) {
      return;
    }

    this.full.split.classList.toggle("changes-rail-away", !up);
    this.full.toggle.setAttribute("aria-pressed", up ? "true" : "false");
    this.full.toggle.title = up ? "Hide the snapshots" : "Show the snapshots";
    this.full.toggle.setAttribute("aria-label", this.full.toggle.title);
    this.full.splitter.setAttribute("aria-valuenow", String(up ? Math.round(this.railWidth) : 0));

    if (up) {
      this.full.rail.style.flex = `0 0 ${this.railWidth}px`;
    }
  }

  /** Drag, or the arrow keys, to say how much of the card the snapshots are worth. */
  private dragRail(splitter: HTMLElement): void {
    const widest = (): number => Math.max(240, (this.full?.card.clientWidth ?? 800) * 0.4);
    const settle = (width: number): void => {
      // Rounded AFTER the clamp, not before: the ceiling is a fraction of the card's width, so
      // rounding first left the pinned width - and the `aria-valuenow` read off it - as
      // 596.8000000000001, which is what a screen reader would then say out loud.
      this.railWidth = Math.round(Math.min(Math.max(140, width), widest()));
      if (this.full) {
        this.full.rail.style.flex = `0 0 ${this.railWidth}px`;
        splitter.setAttribute("aria-valuenow", String(this.railWidth));
        splitter.setAttribute("aria-valuemax", String(Math.round(widest())));
      }
    };

    splitter.setAttribute("aria-valuemin", "140");

    let start = 0;
    let startWidth = 0;
    const onMove = (event: PointerEvent): void => settle(startWidth + (event.clientX - start));

    splitter.addEventListener("pointerdown", (event) => {
      // Dragging a rail that is not there is how a divider becomes a mystery. Bring it back first.
      if (this.railHidden) {
        return;
      }

      start = event.clientX;
      startWidth = this.full?.rail.getBoundingClientRect().width ?? this.railWidth;

      // Capture keeps the drag alive when the pointer outruns a 6px divider, but it is a nicety:
      // it throws for a pointer that is no longer down, and taking the listeners with it would
      // turn a lost capture into a dead splitter.
      try {
        splitter.setPointerCapture(event.pointerId);
      } catch {
        /* the drag still tracks through the listeners below */
      }

      splitter.addEventListener("pointermove", onMove);
      const done = (): void => {
        splitter.removeEventListener("pointermove", onMove);
        splitter.removeEventListener("pointerup", done);
        splitter.removeEventListener("pointercancel", done);
      };
      splitter.addEventListener("pointerup", done);
      splitter.addEventListener("pointercancel", done);
    });

    // Arrows only. The divider resizes; putting the rail away is the head's button, and one job
    // per control beats a divider that quietly does two.
    splitter.addEventListener("keydown", (event) => {
      if ((event.key !== "ArrowLeft" && event.key !== "ArrowRight") || this.railHidden) {
        return;
      }

      event.preventDefault();
      settle(this.railWidth + (event.key === "ArrowRight" ? 24 : -24));
    });
  }

  /**
   * The rail: every snapshot this file's log holds, newest first, as one line each.
   *
   * Rebuilt rather than patched, because the list behind it is rebuilt on every refresh and two
   * lists maintained by different rules is how they come to disagree. It is the SAME rows the
   * pane draws, said shorter - a round's number and who wrote it over its modules.
   */
  private drawFullList(): void {
    if (!this.full) {
      return;
    }

    const rail = this.full.rail;

    // The rebuild throws away the element the arrows are standing on, so whether they were
    // standing on one is remembered and given back below. Without it, walking the list with the
    // keyboard drops focus on the first step and the next arrow jumps to the top.
    const had = document.activeElement instanceof Node && rail.contains(document.activeElement);
    rail.replaceChildren();

    const state = this.state;
    const rounds = (state?.rounds ?? []).filter((round) => round.entries.length > 0);

    // ONE SNAPSHOT IS NOT A CHOICE - the same rule the file select follows. A rail offering the
    // single thing already on screen is a column of chrome charging rent for nothing.
    const only = rounds.reduce((count, round) => count + round.entries.length, 0) < 2;
    rail.hidden = only;
    this.full.splitter.hidden = only;
    this.full.toggle.hidden = only;
    if (only) {
      return;
    }

    let marked = (state?.acceptedAt ?? 0) <= 0;
    for (const round of rounds) {
      if (!marked && state && round.round <= state.acceptedAt) {
        marked = true;
        const line = document.createElement("div");
        line.className = "changes-accepted changes-full-accepted";
        line.textContent = "accepted";
        rail.appendChild(line);
      }

      const group = document.createElement("div");
      group.className = "changes-full-group";
      group.setAttribute("role", "group");
      group.setAttribute("aria-label", `Round ${round.round}, by ${round.by}`);
      group.dataset.round = String(round.round);

      const head = document.createElement("div");
      head.className = "changes-full-round";

      const number = document.createElement("span");
      number.className = "changes-number";
      number.textContent = String(round.round);

      const who = document.createElement("span");
      who.className = `changes-by changes-by-${round.by.toLowerCase() === "developer" ? "developer" : "agent"}`;
      who.textContent = round.by;

      head.append(number, who);
      group.appendChild(head);

      for (const entry of round.entries) {
        group.appendChild(this.drawRailEntry(round, entry));
      }

      rail.appendChild(group);
    }

    const here = rail.querySelector<HTMLElement>(".changes-full-showing");
    if (had && here) {
      here.focus();
    }

    here?.scrollIntoView({ block: "nearest" });
  }

  private drawRailEntry(round: ChangeRound, entry: ChangeEntry): HTMLElement {
    const showing = this.showing?.round === round.round && this.showing.module === entry.module;

    const option = document.createElement("div");
    option.className = "changes-full-entry";
    option.dataset.module = entry.module;
    option.dataset.round = String(round.round);
    option.setAttribute("role", "option");
    option.setAttribute("aria-selected", showing ? "true" : "false");

    // Roving tabindex: the rail is ONE stop on the way round the card, and the arrows move within
    // it. Tabbing through forty snapshots to reach the close button is not keyboard support.
    option.tabIndex = showing ? 0 : -1;
    if (showing) {
      option.classList.add("changes-full-showing");
    }

    const name = document.createElement("span");
    name.className = "changes-module";
    name.textContent = entry.module;

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
      counts.textContent = "let go";
      counts.classList.add("changes-gone");
    }

    // The whole line, because a name and its counts are one thing to point at.
    option.title = entry.from
      ? `${entry.module}, round ${round.round}, renamed from ${entry.from}`
      : `${entry.module}, round ${round.round}, ${KIND_WORD[entry.kind] ?? entry.kind}`;

    option.append(name, counts);
    option.addEventListener("click", () => void this.open(round.round, entry.module));
    return option;
  }

  /**
   * The arrows, and what they select.
   *
   * SELECTION FOLLOWS FOCUS, because selecting here only draws a comparison - nothing is written,
   * nothing is spent, and a reader walking the list wants to SEE each one, not press Enter forty
   * times to find out what they are looking at.
   */
  private railKey(event: KeyboardEvent): void {
    const rail = this.full?.rail;
    if (!rail) {
      return;
    }

    const options = [...rail.querySelectorAll<HTMLElement>('[role="option"]')];
    if (options.length === 0) {
      return;
    }

    const at = options.findIndex((one) => one === document.activeElement);
    const going = event.key === "ArrowDown" ? Math.min(options.length - 1, at + 1)
      : event.key === "ArrowUp" ? Math.max(0, at - 1)
      : event.key === "Home" ? 0
      : event.key === "End" ? options.length - 1
      : event.key === "Enter" || event.key === " " ? Math.max(at, 0)
      : -1;

    if (going < 0) {
      return;
    }

    event.preventDefault();
    const next = options[going];
    if (!next) {
      return;
    }

    next.focus();
    next.scrollIntoView({ block: "nearest" });
    const round = Number(next.dataset.round ?? 0);
    const module = next.dataset.module ?? "";
    if (round > 0 && module && !(this.showing?.round === round && this.showing.module === module)) {
      void this.open(round, module);
    }
  }

  /** The comparison in the card: the rows the pane already has, drawn by the same renderer. */
  private drawFullDiff(): void {
    if (!this.full) {
      return;
    }

    this.full.title.textContent = this.showingTitle;
    this.full.card.setAttribute("aria-label", `What changed in ${this.showingTitle}`);
    this.full.body.replaceChildren();
    drawDiffRows(this.full.body, this.showingRows, "sync");
    this.full.body.scrollTop = 0;
  }

  /** The pane as the dev surface reads and drives it. */
  probe(): ChangesPaneProbe {
    return {
      state: () => ({
        project: this.state?.project ?? "",
        files: [...this.file.options].map((one) => one.value),
        acceptedAt: this.state?.acceptedAt ?? 0,
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
        full: this.full !== null,
        fullChoices: this.full && !this.full.rail.hidden
          ? this.full.rail.querySelectorAll('[role="option"]').length
          : 0,
        railUp: this.full !== null && !this.railHidden && !this.full.rail.hidden,
        railWidth: Math.round(this.full && !this.railHidden
          ? this.full.rail.getBoundingClientRect().width
          : 0),
        behind: this.hostStamp > this.drawnStamp,
      }),
      press: (control) => {
        const button = control === "refresh" ? this.refresh
          : control === "snapshot" ? this.snapshot
          : control === "accept" ? this.accept
          // The card's own control, so a driver can put the snapshots away and bring them back.
          : control === "rail" ? this.full?.toggle ?? null
          : null;
        button?.click();
        return button !== null;
      },
      expand: (open) => {
        if (open) {
          this.expand();
          return this.full !== null;
        }

        const was = this.full !== null;
        this.full?.dismiss();
        return was;
      },
      chooseFile: (name) => {
        const option = [...this.file.options].find(
          (one) => one.value.toLowerCase() === name.toLowerCase());
        if (!option) {
          return false;
        }

        this.file.value = option.value;
        this.file.dispatchEvent(new Event("change"));
        return true;
      },
      show: (round, module, where, gesture) => {
        const row = where === "full"
          ? this.full?.rail.querySelector<HTMLElement>(
            `.changes-full-entry[data-round="${round}"][data-module="${CSS.escape(module)}"]`) ?? null
          : this.list.querySelector<HTMLElement>(
            `.changes-round[data-round="${round}"] .changes-entry[data-module="${CSS.escape(module)}"]`);

        if (!row) {
          return false;
        }

        // The REAL gesture: a double-click is a click and then a dblclick, in that order, which
        // is what the row's own two listeners see from a mouse. Dispatching only the dblclick
        // would test a path no developer can take.
        row.click();
        if (gesture === "double") {
          row.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true }));
        }

        return true;
      },
    };
  }
}
