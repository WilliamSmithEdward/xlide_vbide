/*
 * The search widget: ONE floating search UI where two used to duplicate each other
 * (developer, 2026-08-05: "move the search pane into the search popup window, so we're not
 * duplicating UIs"). Monaco's find widget and the bottom panel's Search tab are both gone;
 * this widget floats where Monaco's find sat and carries the whole job.
 *
 * Scope decides the engine. Module scope is live: matches found in the current model as the
 * query is typed, painted with Monaco's own find-match decoration classes, walked with
 * Enter/F3, replaced through executeEdits so a replace-all is one undoable step that flows
 * the normal didChange path to the host. Workbook and All scopes ask the host's engine (the
 * panel's old protocol, unchanged) and render the grouped results inside the widget; rows
 * navigate through the same route problem rows use.
 *
 * Monaco's own find cannot be unbundled - the multicursor feature imports its controller
 * module, which self-registers - so this widget's actions claim the keys instead
 * (dynamically added rules outrank built-ins). The command palette still lists Monaco's
 * find entries; that residue is recorded in the handoff.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";

/** One search hit as the widget draws it; structurally what the host sends. */
export interface SearchMatch {
  workbook?: string | null;
  module: string;
  line: number;
  column: number;
  length: number;
  preview: string;
}

export interface SearchWidgetHandlers {
  /** Ask the host's engine to search the scope. Returns the request id. */
  search(query: string, matchCase: boolean, wholeWord: boolean, scope: string): number;
  /** Replace across the scope; answered like a search, plus the replaced count. */
  replaceAll(query: string, matchCase: boolean, wholeWord: boolean, scope: string, replacement: string): number;
  /** Jump to a result row's line. */
  navigate(module: string, line: number, column: number, selectLine: boolean, workbook?: string): void;
}

export class SearchWidget {
  private readonly editor: monaco.editor.IStandaloneCodeEditor;
  private readonly handlers: SearchWidgetHandlers;

  private readonly root: HTMLElement;
  private readonly findInput: HTMLInputElement;
  private readonly replaceInput: HTMLInputElement;
  private readonly caseButton: HTMLButtonElement;
  private readonly wordButton: HTMLButtonElement;
  private readonly scopeSelect: HTMLSelectElement;
  private readonly counter: HTMLElement;
  private readonly previousButton: HTMLButtonElement;
  private readonly nextButton: HTMLButtonElement;
  private readonly replaceButton: HTMLButtonElement;
  private readonly replaceAllButton: HTMLButtonElement;
  private readonly results: HTMLElement;

  /** Whether the widget is showing; mirrored into a context key so Escape can be claimed
   * inside the editor only while it matters. */
  private readonly openKey: monaco.editor.IContextKey<boolean>;

  private readonly decorations: monaco.editor.IEditorDecorationsCollection;
  private matches: monaco.editor.FindMatch[] = [];
  private current = -1;
  private refindTimer: number | undefined;

  /** The scoped search whose answer is awaited; older answers are ignored. */
  private pendingSearchId = 0;

  constructor(host: HTMLElement, editor: monaco.editor.IStandaloneCodeEditor, handlers: SearchWidgetHandlers) {
    this.editor = editor;
    this.handlers = handlers;
    this.decorations = editor.createDecorationsCollection();
    this.openKey = editor.createContextKey<boolean>("xlideSearchOpen", false);

    this.root = document.createElement("div");
    this.root.id = "search-widget";
    this.root.hidden = true;

    const findRow = document.createElement("div");
    findRow.className = "search-widget-row";
    this.findInput = this.makeInput("Find", "Search text");
    this.caseButton = this.makeToggle("Aa", "Match case");
    this.wordButton = this.makeToggle("ab", "Whole word");
    this.scopeSelect = document.createElement("select");
    this.scopeSelect.setAttribute("aria-label", "Search scope");
    for (const [value, label] of [["module", "Module"], ["project", "Workbook"], ["all", "All workbooks"]]) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      this.scopeSelect.appendChild(option);
    }
    this.counter = document.createElement("span");
    this.counter.className = "search-widget-counter";
    this.counter.setAttribute("aria-live", "polite");
    this.previousButton = this.makeIconButton("arrow-up", "Previous match (Shift+Enter)");
    this.nextButton = this.makeIconButton("arrow-down", "Next match (Enter)");
    const closeButton = this.makeIconButton("close", "Close (Escape)");
    findRow.append(this.findInput, this.caseButton, this.wordButton, this.scopeSelect,
      this.counter, this.previousButton, this.nextButton, closeButton);

    const replaceRow = document.createElement("div");
    replaceRow.className = "search-widget-row";
    this.replaceInput = this.makeInput("Replace", "Replacement text");
    this.replaceButton = this.makeTextButton("Replace");
    this.replaceAllButton = this.makeTextButton("Replace All");
    replaceRow.append(this.replaceInput, this.replaceButton, this.replaceAllButton);

    this.results = document.createElement("div");
    this.results.id = "search-widget-results";
    this.results.setAttribute("role", "list");
    this.results.setAttribute("aria-label", "Search results");
    this.results.hidden = true;

    this.root.append(findRow, replaceRow, this.results);
    host.appendChild(this.root);

    this.wire();
    this.registerActions(closeButton);
  }

  /** Routes the answer to a scoped search; the bridge calls this on searchResult. */
  showSearchResults(id: number, matches: SearchMatch[], truncated: boolean, replaced: number): void {
    // Ids are monotonic, so "not older than what we asked for last" is the acceptance test.
    // Equality was a race: a transport that answers synchronously - the demo does - delivered
    // the result before the id assignment landed, and the panel ignored its own answer.
    if (id < this.pendingSearchId || this.scope() === "module") {
      return;
    }

    this.pendingSearchId = id;
    this.results.replaceChildren();
    this.results.hidden = false;

    if (replaced > 0) {
      this.results.appendChild(this.note(
        `${replaced} occurrence${replaced === 1 ? "" : "s"} replaced.`));
    }

    if (matches.length === 0) {
      this.results.appendChild(this.note(replaced > 0 ? "No matches remain." : "No matches."));
      return;
    }

    this.results.appendChild(this.note(
      `${matches.length}${truncated ? "+" : ""} match${matches.length === 1 ? "" : "es"}`
      + (truncated ? " (narrow the query for the rest)" : "")));

    let groupKey = "";
    for (const match of matches) {
      const key = `${(match.workbook ?? "").toLowerCase()}|${match.module.toLowerCase()}`;
      if (key !== groupKey) {
        groupKey = key;
        const header = document.createElement("div");
        header.className = "search-group";
        header.textContent = match.workbook ? `${match.module} - ${match.workbook}` : match.module;
        this.results.appendChild(header);
      }

      const row = document.createElement("button");
      row.type = "button";
      row.className = "search-row";
      row.setAttribute("role", "listitem");

      const where = document.createElement("span");
      where.className = "search-line";
      where.textContent = String(match.line);
      row.appendChild(where);

      const preview = document.createElement("span");
      preview.className = "search-preview";
      preview.textContent = match.preview.trim();
      preview.title = match.preview;
      row.appendChild(preview);

      const target = match;
      row.addEventListener("click", () => {
        this.handlers.navigate(target.module, target.line, target.column, true, target.workbook ?? undefined);
      });

      this.results.appendChild(row);
    }
  }

  private makeInput(placeholder: string, label: string): HTMLInputElement {
    const input = document.createElement("input");
    input.type = "text";
    input.spellcheck = false;
    input.autocomplete = "off";
    input.placeholder = placeholder;
    input.setAttribute("aria-label", label);
    return input;
  }

  private makeToggle(text: string, title: string): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "search-toggle";
    button.textContent = text;
    button.title = title;
    button.setAttribute("aria-pressed", "false");
    return button;
  }

  private makeIconButton(icon: string, title: string): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "search-widget-icon";
    button.title = title;
    button.setAttribute("aria-label", title);
    const glyph = document.createElement("span");
    glyph.className = `codicon codicon-${icon}`;
    glyph.setAttribute("aria-hidden", "true");
    button.appendChild(glyph);
    return button;
  }

  private makeTextButton(text: string): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "search-widget-run";
    button.textContent = text;
    return button;
  }

  private scope(): string {
    return this.scopeSelect.value;
  }

  private options(): { matchCase: boolean; wholeWord: boolean } {
    return {
      matchCase: this.caseButton.getAttribute("aria-pressed") === "true",
      wholeWord: this.wordButton.getAttribute("aria-pressed") === "true",
    };
  }

  private wire(): void {
    const retoggle = (button: HTMLButtonElement) => {
      button.addEventListener("click", () => {
        button.setAttribute("aria-pressed", String(button.getAttribute("aria-pressed") !== "true"));
        this.queryChanged();
      });
    };
    retoggle(this.caseButton);
    retoggle(this.wordButton);

    this.findInput.addEventListener("input", () => this.queryChanged());
    this.scopeSelect.addEventListener("change", () => this.scopeChanged());

    this.previousButton.addEventListener("click", () => this.previous());
    this.nextButton.addEventListener("click", () => this.next());
    this.replaceButton.addEventListener("click", () => this.replaceCurrent());
    this.replaceAllButton.addEventListener("click", () => this.replaceAllRun());

    this.root.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        this.close();
        return;
      }

      if (event.key === "F3") {
        event.preventDefault();
        if (event.shiftKey) { this.previous(); } else { this.next(); }
        return;
      }

      if (event.key === "Enter" && event.target === this.findInput) {
        event.preventDefault();
        if (this.scope() === "module") {
          if (event.shiftKey) { this.previous(); } else { this.next(); }
        } else {
          this.runScoped(false);
        }
        return;
      }

      if (event.key === "Enter" && event.target === this.replaceInput) {
        event.preventDefault();
        if (this.scope() === "module") {
          this.replaceCurrent();
        } else {
          this.runScoped(false);
        }
      }
    });

    // Module-scope matches ride the text: an edit re-finds after a beat, and a module
    // switch re-finds against the new model.
    this.editor.onDidChangeModelContent(() => {
      if (this.root.hidden || this.scope() !== "module" || this.findInput.value.length === 0) {
        return;
      }

      window.clearTimeout(this.refindTimer);
      this.refindTimer = window.setTimeout(() => this.findInModule(true), 150);
    });
    this.editor.onDidChangeModel(() => {
      if (!this.root.hidden && this.scope() === "module") {
        this.findInModule(false);
      }
    });
  }

  private registerActions(closeButton: HTMLButtonElement): void {
    closeButton.addEventListener("click", () => this.close());

    const KeyMod = monaco.KeyMod;
    const KeyCode = monaco.KeyCode;

    this.editor.addAction({
      id: "xlide.search.open",
      label: "Find",
      keybindings: [KeyMod.CtrlCmd | KeyCode.KeyF],
      run: () => this.open({ scope: "module" }),
    });
    this.editor.addAction({
      id: "xlide.search.replace",
      label: "Replace",
      keybindings: [KeyMod.CtrlCmd | KeyCode.KeyH],
      run: () => this.open({ scope: "module", withReplace: true }),
    });
    this.editor.addAction({
      id: "xlide.search.workbook",
      label: "Search Workbook",
      keybindings: [KeyMod.CtrlCmd | KeyMod.Shift | KeyCode.KeyF],
      run: () => this.open({ scope: "project" }),
    });
    this.editor.addAction({
      id: "xlide.search.next",
      label: "Find Next",
      keybindings: [KeyCode.F3],
      run: () => { if (this.ensureOpenForCycle()) { this.next(); } },
    });
    this.editor.addAction({
      id: "xlide.search.previous",
      label: "Find Previous",
      keybindings: [KeyMod.Shift | KeyCode.F3],
      run: () => { if (this.ensureOpenForCycle()) { this.previous(); } },
    });

    // Escape closes from inside the editor too, but only while the widget shows: the
    // context key keeps the claim from shadowing every other Escape in the editor.
    this.editor.addCommand(KeyCode.Escape, () => this.close(), "xlideSearchOpen");
  }

  /** F3 with the widget closed reopens the last search rather than doing nothing. */
  private ensureOpenForCycle(): boolean {
    if (this.root.hidden) {
      this.open({ scope: "module" });
    }

    return this.findInput.value.length > 0;
  }

  open(options?: { scope?: string; withReplace?: boolean }): void {
    if (options?.scope) {
      this.scopeSelect.value = options.scope;
    }

    // A selected run of text is what the developer wants found; multi-line selections
    // keep the previous query, the way every editor treats them.
    const selection = this.editor.getSelection();
    const model = this.editor.getModel();
    if (model && selection && !selection.isEmpty()
      && selection.startLineNumber === selection.endLineNumber) {
      this.findInput.value = model.getValueInRange(selection);
    }

    this.root.hidden = false;
    this.openKey.set(true);
    this.scopeChanged();
    const focused = options?.withReplace ? this.replaceInput : this.findInput;
    focused.focus();
    focused.select();
  }

  close(): void {
    this.root.hidden = true;
    this.openKey.set(false);
    this.decorations.clear();
    this.matches = [];
    this.current = -1;
    this.editor.focus();
  }

  private queryChanged(): void {
    if (this.scope() === "module") {
      this.findInModule(false);
    }
  }

  private scopeChanged(): void {
    if (this.scope() === "module") {
      this.results.hidden = true;
      this.results.replaceChildren();
      this.findInModule(false);
    } else {
      this.decorations.clear();
      this.matches = [];
      this.current = -1;
      this.counter.textContent = "";
      this.results.hidden = this.results.childElementCount === 0;
    }
  }

  private runScoped(replace: boolean): void {
    const query = this.findInput.value;
    if (query.length === 0) {
      return;
    }

    const { matchCase, wholeWord } = this.options();
    this.results.hidden = false;
    this.results.replaceChildren(this.note("Searching..."));
    this.pendingSearchId = replace
      ? this.handlers.replaceAll(query, matchCase, wholeWord, this.scope(), this.replaceInput.value)
      : this.handlers.search(query, matchCase, wholeWord, this.scope());
  }

  private note(text: string): HTMLElement {
    const note = document.createElement("div");
    note.className = "search-note";
    note.textContent = text;
    return note;
  }

  // ---- The module-scope engine: live matches in the current model. ----

  private findInModule(preserveCurrent: boolean): void {
    const model = this.editor.getModel();
    const query = this.findInput.value;
    if (!model || query.length === 0) {
      this.decorations.clear();
      this.matches = [];
      this.current = -1;
      this.counter.textContent = "";
      return;
    }

    const { matchCase, wholeWord } = this.options();
    const separators = wholeWord
      ? this.editor.getOption(monaco.editor.EditorOption.wordSeparators)
      : null;
    const previous = preserveCurrent && this.current >= 0 ? this.matches[this.current]?.range : null;

    this.matches = model.findMatches(query, false, false, matchCase, separators, false, 10000);

    if (this.matches.length === 0) {
      this.current = -1;
    } else if (previous) {
      this.current = this.nearestMatch(previous.getStartPosition());
    } else {
      this.current = this.nearestMatch(this.editor.getPosition() ?? new monaco.Position(1, 1));
    }

    this.decorate();
    this.updateCounter();
  }

  /** The first match at or after the position, wrapping to the first match. */
  private nearestMatch(position: monaco.Position): number {
    for (let i = 0; i < this.matches.length; i++) {
      if (!this.matches[i].range.getStartPosition().isBefore(position)) {
        return i;
      }
    }

    return 0;
  }

  private decorate(): void {
    this.decorations.set(this.matches.map((match, index) => ({
      range: match.range,
      options: {
        className: index === this.current ? "currentFindMatch" : "findMatch",
        stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
        overviewRuler: {
          color: { id: "editorOverviewRuler.findMatchForeground" },
          position: monaco.editor.OverviewRulerLane.Center,
        },
      },
    })));
  }

  private updateCounter(): void {
    if (this.findInput.value.length === 0) {
      this.counter.textContent = "";
    } else if (this.matches.length === 0) {
      this.counter.textContent = "No results";
    } else {
      this.counter.textContent = `${this.current + 1} of ${this.matches.length}`;
    }
  }

  private next(): void {
    this.step(1);
  }

  private previous(): void {
    this.step(-1);
  }

  private step(direction: number): void {
    if (this.scope() !== "module" || this.matches.length === 0) {
      return;
    }

    this.current = (this.current + direction + this.matches.length) % this.matches.length;
    const range = this.matches[this.current].range;
    this.editor.setSelection(range);
    this.editor.revealRangeInCenterIfOutsideViewport(range);
    this.decorate();
    this.updateCounter();
  }

  private replaceCurrent(): void {
    if (this.scope() !== "module" || this.current < 0 || this.matches.length === 0) {
      return;
    }

    const range = this.matches[this.current].range;
    const start = range.getStartPosition();
    this.editor.executeEdits("xlide-search", [{ range, text: this.replaceInput.value }]);
    this.findInModule(false);
    if (this.matches.length > 0) {
      this.current = this.nearestMatch(start);
      this.step(0);
    }
  }

  private replaceAllRun(): void {
    if (this.scope() !== "module") {
      this.runScoped(true);
      return;
    }

    if (this.matches.length === 0) {
      return;
    }

    // One executeEdits, so the whole replace is one undo step and flows the same
    // didChange path as typing; the host and engine learn of it the ordinary way.
    const replacement = this.replaceInput.value;
    const count = this.matches.length;
    this.editor.executeEdits("xlide-search",
      this.matches.map((match) => ({ range: match.range, text: replacement })));
    this.findInModule(false);
    this.counter.textContent = `${count} replaced`;
  }
}
