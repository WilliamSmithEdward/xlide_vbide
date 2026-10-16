/*
 * The Bookmarks pane: every module's bookmarks in one list (#111).
 *
 * The native editor's bookmarks are reached one at a time, Next and Previous through the module
 * on screen. A pane is the overview it never had: each mark with its module, its line, and what
 * the line says, so a mark is found by what it is on rather than by hopping to it. A row goes to
 * its line, Delete takes the mark away, and the bar's Clear All takes every one. Every open
 * file's marks show at once; the bar's file filter narrows to one.
 *
 * Drawn from the store on every change it announces, debounced to a frame: marks move with their
 * lines on every keystroke in a marked module, and the list is short.
 */

import type { BookmarkEntry, Bookmarks } from "./bookmarks.js";
import { showContextMenu } from "./contextmenu.js";
import { ScopeSelect, type ScopeEntry, type ScopeFile } from "./scopeselect.js";

export interface BookmarksPaneHost {
  /** Opens the module and puts the caret on the line. */
  navigate(module: string, line: number, project: string | null): void;
  /** The project's display name for the lowercased key a model URI carries, null when unknown. */
  projectName(key: string): string | null;
  /** The open files, by display name, in the explorer's order. */
  files(): string[];
  /** The count the pane's tab wears; zero takes it off. */
  badge(count: number): void;
}

/** What the pane is showing, as the api reads it. */
export interface BookmarksPaneState {
  /** The file the list is narrowed to, or null while it shows every file. */
  file: string | null;
  /** The files the filter offers. Empty while only one file is open and the filter is hidden. */
  files: string[];
  /** The rows as drawn, in order. */
  rows: BookmarkRow[];
}

/** One drawn row, as the api reads it. */
export interface BookmarkRow {
  project: string | null;
  module: string;
  line: number;
  text: string;
}

export class BookmarksPane {
  private readonly list: HTMLElement;
  private readonly clear: HTMLButtonElement;
  private readonly summary: HTMLElement;
  /**
   * Which file's marks to show: the shared scope control, its FILE select only. Every file's
   * marks show by default; the select appears once a second file is open, as it does on the
   * Problems and Tests panes, because one file is not a choice.
   */
  private readonly scope: ScopeSelect;
  private entries: (BookmarkEntry & { project: string | null })[] = [];
  private pending = 0;

  constructor(root: HTMLElement, private readonly store: Bookmarks, private readonly host: BookmarksPaneHost) {
    this.list = root.querySelector("#bookmarks-list") as HTMLElement;
    this.clear = root.querySelector("#bookmarks-clear") as HTMLButtonElement;
    this.summary = root.querySelector("#bookmarks-summary") as HTMLElement;

    this.scope = new ScopeSelect("bookmarks-scope", "Show bookmarks from", "bookmarks", () => this.render());
    this.summary.before(this.scope.files);

    // Narrowed to a file, the button clears what the list shows and says so: clearing marks in
    // files the developer cannot see from here would be a loss they find later.
    this.clear.addEventListener("click", () => {
      if (this.scope.showsEverything()) {
        this.store.clearAll();
        return;
      }
      for (const entry of this.entries) {
        this.store.remove(entry.uri, entry.line);
      }
    });
    store.onDidChange(() => this.schedule());

    this.list.addEventListener("click", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>("[data-i]");
      if (row) {
        this.goTo(Number(row.dataset.i));
      }
    });

    this.list.addEventListener("keydown", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>("[data-i]");
      if (!row) {
        return;
      }
      const i = Number(row.dataset.i);
      if (event.key === "Enter" || event.key === " ") {
        this.goTo(i);
      } else if (event.key === "Delete") {
        this.removeAt(i);
      } else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        this.focusRow(i + (event.key === "ArrowDown" ? 1 : -1));
      } else {
        return;
      }
      event.preventDefault();
    });

    this.list.addEventListener("contextmenu", (event) => {
      const row = (event.target as HTMLElement).closest<HTMLElement>("[data-i]");
      if (!row) {
        return;
      }
      event.preventDefault();
      const i = Number(row.dataset.i);
      showContextMenu(event.clientX, event.clientY, [
        { label: "Go To", run: () => this.goTo(i) },
        { label: "Remove Bookmark", run: () => this.removeAt(i) },
        {},
        { label: "Clear All Bookmarks", run: () => this.store.clearAll() },
      ]);
    });

    this.render();
  }

  /** The rows as drawn, in order. */
  rows(): BookmarkRow[] {
    return this.entries.map((entry) => ({
      project: entry.project, module: entry.module, line: entry.line, text: entry.text,
    }));
  }

  /** The filter and the rows it lets through. */
  state(): BookmarksPaneState {
    return {
      file: this.scope.scopeFile(),
      files: this.scope.files.hidden ? [] : this.host.files(),
      rows: this.rows(),
    };
  }

  /**
   * Narrows to one file by name, or back to every file with null, through the select's own
   * change - what picking it in the list does. Answers false for a file the filter does not offer.
   */
  showFile(file: string | null): boolean {
    const wanted = file === null
      ? [...this.scope.files.options].find((one) => one.value.startsWith("@"))
      : [...this.scope.files.options].find((one) => one.value === `file:${file.toLowerCase()}`);
    if (!wanted || this.scope.files.hidden) {
      return false;
    }
    this.scope.files.value = wanted.value;
    this.scope.files.dispatchEvent(new Event("change"));
    return true;
  }

  /** Goes to the i-th row's line, as a click on it does. */
  goTo(i: number): boolean {
    const entry = this.entries[i];
    if (!entry) {
      return false;
    }
    this.host.navigate(entry.module, entry.line, entry.project);
    return true;
  }

  /** Takes the i-th row's mark away, as Delete on it does, and keeps the keyboard in the list. */
  removeAt(i: number): boolean {
    const entry = this.entries[i];
    if (!entry) {
      return false;
    }
    const hadFocus = this.list.contains(document.activeElement);
    this.store.remove(entry.uri, entry.line);
    this.render();
    if (hadFocus) {
      this.focusRow(Math.min(i, this.entries.length - 1));
    }
    return true;
  }

  /** Presses the bar's button: every mark, or the shown file's while the list is narrowed. */
  clearAll(): void {
    this.clear.click();
  }

  private focusRow(i: number): void {
    this.list.querySelector<HTMLElement>(`[data-i="${i}"]`)?.focus();
  }

  private schedule(): void {
    if (this.pending === 0) {
      this.pending = requestAnimationFrame(() => {
        this.pending = 0;
        this.render();
      });
    }
  }

  render(): void {
    const focused = (document.activeElement as HTMLElement | null)?.closest?.("#bookmarks-list [data-i]");
    const focusedAt = focused ? Number((focused as HTMLElement).dataset.i) : -1;

    const everywhere = this.store.everywhere()
      .map((entry) => ({ ...entry, project: this.host.projectName(entry.projectKey) }))
      .sort((a, b) =>
        (a.project ?? a.projectKey).localeCompare(b.project ?? b.projectKey, undefined, { sensitivity: "base" })
        || a.module.localeCompare(b.module, undefined, { sensitivity: "base" })
        || a.line - b.line);

    // Every open file is offered, with or without marks, so the filter is there before it has
    // anything to narrow - and a file whose marks outlived it still counts under its own name.
    const files = new Map<string, ScopeFile>();
    for (const name of this.host.files()) {
      files.set(name.toLowerCase(), { name, count: 0 });
    }
    const modules = new Map<string, ScopeEntry>();
    for (const entry of everywhere) {
      const file = entry.project ?? entry.projectKey;
      const counted = files.get(file.toLowerCase()) ?? { name: file, count: 0 };
      counted.count++;
      files.set(file.toLowerCase(), counted);
      const key = `${file.toLowerCase()}/${entry.module.toLowerCase()}`;
      const module = modules.get(key) ?? { key, name: entry.module, file, count: 0 };
      module.count++;
      modules.set(key, module);
    }
    this.scope.setEntries([...modules.values()], null, [...files.values()]);

    this.entries = everywhere.filter((entry) => {
      const file = entry.project ?? entry.projectKey;
      return this.scope.admits(`${file.toLowerCase()}/${entry.module.toLowerCase()}`, file);
    });

    // With several files open and all of them showing, every row says which file it is in; a
    // module name alone is only enough inside one file.
    const nameFiles = files.size > 1 && this.scope.showsEverything();

    this.list.replaceChildren();
    this.host.badge(everywhere.length);
    const narrowed = !this.scope.showsEverything();
    this.clear.disabled = this.entries.length === 0;
    (this.clear.lastElementChild as HTMLElement).textContent = narrowed ? "Clear File" : "Clear All";
    this.clear.title = narrowed
      ? `Clear the bookmarks in ${this.scope.scopeName() ?? "this file"}`
      : "Clear all bookmarks (Ctrl+Alt+Shift+K)";
    this.summary.textContent = this.entries.length === 0
      ? ""
      : `${this.entries.length} bookmark${this.entries.length === 1 ? "" : "s"}`;

    if (this.entries.length === 0) {
      const empty = document.createElement("div");
      empty.className = "panel-empty";
      const said = document.createElement("span");
      empty.appendChild(said);

      if (everywhere.length === 0) {
        said.textContent = "No bookmarks. Ctrl+Alt+K marks the caret's line, or use the Bookmarks button on the toolbar.";
      } else {
        // Narrowed to a file with none: say so, and how many are elsewhere, with the way back.
        said.textContent = `No bookmarks in ${this.scope.scopeName() ?? "this file"}. ${everywhere.length} elsewhere.`;
        const showAll = document.createElement("button");
        showAll.type = "button";
        showAll.className = "panel-empty-act";
        showAll.textContent = "Show All";
        showAll.addEventListener("click", () => {
          this.scope.reset();
          this.render();
        });
        empty.appendChild(showAll);
      }
      this.list.appendChild(empty);
      return;
    }

    this.entries.forEach((entry, i) => {
      const row = document.createElement("div");
      row.className = "row bookmark-row";
      row.tabIndex = 0;
      row.dataset.i = String(i);
      row.setAttribute("role", "listitem");

      const mark = document.createElement("span");
      mark.className = "mark codicon codicon-bookmark";
      mark.setAttribute("aria-hidden", "true");

      const body = document.createElement("div");
      body.className = "body";

      const text = document.createElement("div");
      text.className = "message";
      const said = entry.text.trim();
      text.textContent = said.length > 0 ? said : "(blank line)";
      if (said.length === 0) {
        text.classList.add("bookmark-blank");
      }

      const where = document.createElement("div");
      where.className = "where";
      where.textContent = (nameFiles ? `${entry.project ?? entry.projectKey} - ` : "")
        + `${entry.module}, line ${entry.line}`;

      body.append(text, where);
      row.append(mark, body);
      this.list.appendChild(row);
    });

    if (focusedAt >= 0) {
      this.focusRow(Math.min(focusedAt, this.entries.length - 1));
    }
  }
}
