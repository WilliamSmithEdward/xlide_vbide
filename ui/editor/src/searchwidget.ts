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
  /** The active group's editor - the widget searches wherever the developer is. */
  private readonly editorOf: () => monaco.editor.IStandaloneCodeEditor;
  private readonly handlers: SearchWidgetHandlers;

  private get editor(): monaco.editor.IStandaloneCodeEditor {
    return this.editorOf();
  }

  private readonly root: HTMLElement;
  private readonly expandButton: HTMLButtonElement;
  private readonly replaceRow: HTMLElement;
  private readonly findInput: HTMLInputElement;
  private readonly replaceInput: HTMLInputElement;
  private readonly caseButton: HTMLButtonElement;
  private readonly wordButton: HTMLButtonElement;
  private readonly scopeSelect: HTMLSelectElement;
  private readonly counter: HTMLElement;
  private readonly previousButton: HTMLButtonElement;
  private readonly nextButton: HTMLButtonElement;
  private readonly findAllButton: HTMLButtonElement;
  private readonly replaceButton: HTMLButtonElement;
  private readonly replaceAllButton: HTMLButtonElement;
  private readonly results: HTMLElement;

  /** Whether the module-scope match table is showing; it rides every re-find while true. */
  private moduleResultsOpen = false;

  /** Whether the widget is showing; mirrored into a context key PER EDITOR so Escape can be
   * claimed inside whichever editor has it, only while it matters. */
  private readonly openKeys: monaco.editor.IContextKey<boolean>[] = [];
  private isOpen = false;

  /** Find-match paint per editor: matches decorate the editor being searched. */
  private readonly decorationsBy = new Map<monaco.editor.IStandaloneCodeEditor, monaco.editor.IEditorDecorationsCollection>();

  private get decorations(): monaco.editor.IEditorDecorationsCollection {
    const editor = this.editor;
    let collection = this.decorationsBy.get(editor);
    if (!collection) {
      collection = editor.createDecorationsCollection([]);
      this.decorationsBy.set(editor, collection);
    }
    return collection;
  }

  private matches: monaco.editor.FindMatch[] = [];
  private current = -1;
  private refindTimer: number | undefined;

  /** The scoped search whose answer is awaited; older answers are ignored. */
  private pendingSearchId = 0;

  /**
   * The last answer to a scoped search, as the host reported it.
   *
   * The module-scope engine below keeps `matches` and `current`; a project or workbook search is
   * answered by the host and its result lived only in the rows this panel drew. So the two fields
   * a reader would naturally check were 0 and -1 for every scope but one.
   */
  private scopedResult: { matches: number; truncated: boolean; replaced: number } | null = null;

  constructor(editorOf: () => monaco.editor.IStandaloneCodeEditor, handlers: SearchWidgetHandlers) {
    this.editorOf = editorOf;
    this.handlers = handlers;

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
    for (const [value, label] of [["module", "Module"], ["project", "Workbook"], ["all", "All workbooks"]] as const) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      this.scopeSelect.appendChild(option);
    }
    this.counter = document.createElement("span");
    this.counter.className = "search-widget-counter";
    this.counter.setAttribute("aria-live", "polite");
    this.findAllButton = this.makeTextButton("Find All");
    this.previousButton = this.makeIconButton("arrow-up", "Previous match (Shift+Enter)");
    this.nextButton = this.makeIconButton("arrow-down", "Next match (Enter)");
    const closeButton = this.makeIconButton("close", "Close (Escape)");
    findRow.append(this.findInput, this.caseButton, this.wordButton, this.scopeSelect,
      this.findAllButton, this.counter, this.previousButton, this.nextButton, closeButton);

    this.replaceRow = document.createElement("div");
    this.replaceRow.className = "search-widget-row";
    this.replaceInput = this.makeInput("Replace", "Replacement text");
    this.replaceButton = this.makeTextButton("Replace");
    this.replaceAllButton = this.makeTextButton("Replace All");
    this.replaceRow.append(this.replaceInput, this.replaceButton, this.replaceAllButton);

    // The expander: find-only until asked, the way the developer wants it and the way
    // every editor's find widget folds. It spans the input rows' left edge; Ctrl+H forces
    // it open, the chevron remembers its state otherwise.
    this.expandButton = document.createElement("button");
    this.expandButton.type = "button";
    this.expandButton.className = "search-widget-expand";
    this.expandButton.title = "Toggle replace";
    this.expandButton.setAttribute("aria-label", "Toggle replace");
    this.expandButton.setAttribute("aria-expanded", "false");
    const chevron = document.createElement("span");
    chevron.className = "codicon codicon-chevron-right";
    chevron.setAttribute("aria-hidden", "true");
    this.expandButton.appendChild(chevron);
    this.replaceRow.hidden = true;

    const rows = document.createElement("div");
    rows.className = "search-widget-rows";
    rows.append(findRow, this.replaceRow);

    const head = document.createElement("div");
    head.className = "search-widget-head";
    head.append(this.expandButton, rows);

    this.results = document.createElement("div");
    this.results.id = "search-widget-results";
    this.results.setAttribute("role", "list");
    this.results.setAttribute("aria-label", "Search results");
    this.results.hidden = true;

    this.root.append(head, this.results);
    closeButton.addEventListener("click", () => this.close());

    this.wire();
  }

  /**
   * Puts the widget over a group's editor area. Called for the active group at boot and
   * whenever the active group changes, so the widget floats where the developer works.
   */
  attachTo(container: HTMLElement): void {
    if (this.root.parentElement !== container) {
      container.appendChild(this.root);
    }
  }

  /**
   * The active editor changed while the widget may be open: match paint belongs to the new
   * editor's model now, and the stale paint on the old one goes.
   */
  onActiveEditorChanged(): void {
    if (this.isOpen && this.scope() === "module") {
      this.clearAllDecorations();
      this.findInModule(false);
    }
  }

  private clearAllDecorations(): void {
    for (const collection of this.decorationsBy.values()) {
      collection.clear();
    }
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

    // KEPT, not only drawn. `matches` and `current` below describe the module-scope engine and
    // are structurally 0 and -1 for every other scope, so a probe reading them after a project
    // search saw "no matches" whatever the panel had just rendered - and a driver that searched
    // nothing looked identical to one that found nothing. Taken from the arguments rather than
    // from the rows, so it is the answer the panel was given and not a re-reading of the DOM.
    this.scopedResult = { matches: matches.length, truncated, replaced };

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
    this.findAllButton.addEventListener("click", () => {
      if (this.scope() === "module") {
        this.moduleResultsOpen = true;
        this.showModuleResults();
      } else {
        this.runScoped(false);
      }
    });

    this.expandButton.addEventListener("click", () => this.setReplaceShown(this.replaceRow.hidden !== false));

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

  }

  /**
   * Wires one editor into the widget: the search actions and keys, the Escape claim, and the
   * re-find listeners. Called once per editor as the workspace creates groups, so Ctrl+F
   * works from any group and acts on the group it was pressed in.
   */
  registerOn(editor: monaco.editor.IStandaloneCodeEditor): void {
    const KeyMod = monaco.KeyMod;
    const KeyCode = monaco.KeyCode;

    // The open flag is per editor, because a standalone editor's commands go into a keybinding
    // service shared by every editor on the page and the when-clause is the only scoping there
    // is. Each editor's Escape rule matches only while ITS context says the widget is open.
    const openKey = editor.createContextKey<boolean>("xlideSearchOpen", this.isOpen);
    this.openKeys.push(openKey);

    editor.addAction({
      id: "xlide.search.open",
      label: "Find",
      keybindings: [KeyMod.CtrlCmd | KeyCode.KeyF],
      run: () => this.open({ scope: "module" }),
    });
    editor.addAction({
      id: "xlide.search.replace",
      label: "Replace",
      keybindings: [KeyMod.CtrlCmd | KeyCode.KeyH],
      run: () => this.open({ scope: "module", withReplace: true }),
    });
    editor.addAction({
      id: "xlide.search.workbook",
      label: "Search Workbook",
      keybindings: [KeyMod.CtrlCmd | KeyMod.Shift | KeyCode.KeyF],
      run: () => this.open({ scope: "project" }),
    });
    editor.addAction({
      id: "xlide.search.next",
      label: "Find Next",
      keybindings: [KeyCode.F3],
      run: () => { if (this.ensureOpenForCycle()) { this.next(); } },
    });
    editor.addAction({
      id: "xlide.search.previous",
      label: "Find Previous",
      keybindings: [KeyMod.Shift | KeyCode.F3],
      run: () => { if (this.ensureOpenForCycle()) { this.previous(); } },
    });

    // Escape closes from inside the editor too, but only while the widget shows: the
    // context key keeps the claim from shadowing every other Escape in the editor.
    editor.addCommand(KeyCode.Escape, () => this.close(), "xlideSearchOpen");

    // Module-scope matches ride the text: an edit re-finds after a beat, and a model switch
    // re-finds against the new model - but only for the editor the widget is bound to.
    editor.onDidChangeModelContent(() => {
      if (editor !== this.editor || this.root.hidden || this.scope() !== "module" || this.findInput.value.length === 0) {
        return;
      }

      window.clearTimeout(this.refindTimer);
      this.refindTimer = window.setTimeout(() => this.findInModule(true), 150);
    });
    editor.onDidChangeModel(() => {
      if (editor === this.editor && !this.root.hidden && this.scope() === "module") {
        this.findInModule(false);
      }
    });
  }

  /** F3 with the widget closed reopens the last search rather than doing nothing. */
  private ensureOpenForCycle(): boolean {
    if (this.root.hidden) {
      this.open({ scope: "module" });
    }

    return this.findInput.value.length > 0;
  }

  /**
   * What the widget is showing, for the debug api's `ui` route.
   *
   * Reported from the fields the widget acts on, not scraped from its inputs: the two agreeing
   * is what makes the report worth anything, and a scrape could not tell a stale render from
   * the state.
   */
  state(): {
    open: boolean;
    query: string;
    replacement: string;
    scope: string;
    matchCase: boolean;
    wholeWord: boolean;
    matches: number;
    current: number;
    replaceShown: boolean;
    scopedMatches: number;
    scopedTruncated: boolean;
    scopedReplaced: number;
  } {
    return {
      open: this.isOpen,
      query: this.findInput.value,
      replacement: this.replaceInput.value,
      scope: this.scopeSelect.value,
      matchCase: this.caseButton.getAttribute("aria-pressed") === "true",
      wholeWord: this.wordButton.getAttribute("aria-pressed") === "true",
      // MODULE SCOPE ONLY, and the names now say so by having siblings. These two come from the
      // live decorations in the current model, which is a different engine from the host search
      // the other scopes use - so they are 0 and -1 whenever the scope is not "module", however
      // many matches the panel is showing.
      matches: this.matches.length,
      current: this.current,
      replaceShown: !this.replaceRow.hidden,
      // The other scopes' answer, as the host reported it. -1 for "nothing has been asked yet",
      // which is a different state from a search that found nothing.
      scopedMatches: this.scopedResult?.matches ?? -1,
      scopedTruncated: this.scopedResult?.truncated ?? false,
      scopedReplaced: this.scopedResult?.replaced ?? 0,
    };
  }

  /** Types a query and runs the search, the way the find box does when a developer types in it. */
  find(query: string, options?: { scope?: string; matchCase?: boolean; wholeWord?: boolean }): void {
    if (!this.isOpen) {
      this.open(options?.scope ? { scope: options.scope } : undefined);
    }

    if (options?.scope && this.scopeSelect.value !== options.scope) {
      // Set AND announced. `change` does not fire for a value assigned in script, and everything
      // that reacts to a scope change hangs off that event: clearing the previous scope's
      // decorations, emptying its match table, re-running the search under the new one. Assigning
      // alone left the select reading "project" over a live module-scope result - the same
      // mistake the toggle buttons below are pressed rather than set to avoid, one control over.
      this.scopeSelect.value = options.scope;
      this.scopeSelect.dispatchEvent(new Event("change", { bubbles: true }));
    }

    // The toggles are pressed through their buttons, because pressing them is what runs the
    // handler that re-searches. Setting aria-pressed alone would leave the widget showing one
    // thing and searching by another - the first version accepted matchCase and applied
    // nothing, and a case-sensitive search for `recalculate` still found `Recalculate`.
    for (const [wanted, button] of [
      [options?.matchCase, this.caseButton],
      [options?.wholeWord, this.wordButton],
    ] as const) {
      if (wanted !== undefined && (button.getAttribute("aria-pressed") === "true") !== wanted) {
        button.click();
      }
    }

    this.findInput.value = query;
    this.findInput.dispatchEvent(new Event("input", { bubbles: true }));
  }

  /**
   * The buttons, as methods, for a driver that has no pointer.
   *
   * `find` above raises an `input` event, and the only handler for that searches when the scope is
   * "module" - so typing a query under any other scope types it and searches nothing. That is the
   * behaviour a person sees too: they type, then press Find All or Enter. These are those presses,
   * going through the same methods the click listeners call rather than round a second path.
   *
   * Replace All is the reason this matters most. It rewrites text across every module of a
   * project, it is the most destructive thing on this surface, and until 2026-08-11 nothing could
   * trigger it except a person with a mouse.
   */
  runFindAll(): void {
    if (this.scope() === "module") {
      this.moduleResultsOpen = true;
      this.showModuleResults();
      return;
    }

    this.runScoped(false);
  }

  /**
   * Types into the replace box and reveals it, which is what a developer does before replacing.
   * The row is hidden until the expander is pressed, and a replace against a hidden box would be
   * a state no person can reach.
   */
  setReplacement(text: string): void {
    if (this.replaceRow.hidden) {
      this.setReplaceShown(true);
    }

    this.replaceInput.value = text;
    this.replaceInput.dispatchEvent(new Event("input", { bubbles: true }));
  }

  runReplaceAll(): void {
    this.replaceAllRun();
  }

  runReplaceCurrent(): void {
    this.replaceCurrent();
  }

  goToNextMatch(): void {
    this.next();
  }

  goToPreviousMatch(): void {
    this.previous();
  }

  open(options?: { scope?: string; withReplace?: boolean }): void {
    if (options?.scope && this.scopeSelect.value !== options.scope) {
      // Set AND announced. `change` does not fire for a value assigned in script, and everything
      // that reacts to a scope change hangs off that event: clearing the previous scope's
      // decorations, emptying its match table, re-running the search under the new one. Assigning
      // alone left the select reading "project" over a live module-scope result - the same
      // mistake the toggle buttons below are pressed rather than set to avoid, one control over.
      this.scopeSelect.value = options.scope;
      this.scopeSelect.dispatchEvent(new Event("change", { bubbles: true }));
    }

    // A selected run of text is what the developer wants found; multi-line selections
    // keep the previous query, the way every editor treats them.
    const selection = this.editor.getSelection();
    const model = this.editor.getModel();
    if (model && selection && !selection.isEmpty()
      && selection.startLineNumber === selection.endLineNumber) {
      this.findInput.value = model.getValueInRange(selection);
    }

    if (options?.withReplace) {
      this.setReplaceShown(true);
    }

    this.root.hidden = false;
    this.isOpen = true;
    for (const key of this.openKeys) {
      key.set(true);
    }
    this.scopeChanged();
    const focused = options?.withReplace ? this.replaceInput : this.findInput;
    focused.focus();
    focused.select();
  }

  /** Shows or folds the replace row; the chevron and its announcement follow. */
  private setReplaceShown(shown: boolean): void {
    this.replaceRow.hidden = !shown;
    this.expandButton.setAttribute("aria-expanded", String(shown));
    const chevron = this.expandButton.querySelector(".codicon");
    chevron?.classList.toggle("codicon-chevron-right", !shown);
    chevron?.classList.toggle("codicon-chevron-down", shown);
  }

  close(): void {
    this.root.hidden = true;
    this.isOpen = false;
    for (const key of this.openKeys) {
      key.set(false);
    }
    this.clearAllDecorations();
    this.matches = [];
    this.current = -1;
    this.moduleResultsOpen = false;
    this.editor.focus();
  }

  private queryChanged(): void {
    if (this.scope() === "module") {
      this.findInModule(false);
    }
  }

  private scopeChanged(): void {
    // A result belongs to the scope it was asked in.
    this.scopedResult = null;

    if (this.scope() === "module") {
      this.moduleResultsOpen = false;
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

    // Cleared as the question goes out, so "asked and not yet answered" reads as -1 rather than
    // as the previous query's count. A caller polling for the answer needs those to differ.
    this.scopedResult = null;

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

      // An emptied query resets the match table too, rather than parking "No matches."
      // over nothing.
      if (this.moduleResultsOpen) {
        this.moduleResultsOpen = false;
        this.results.hidden = true;
        this.results.replaceChildren();
      }

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

    if (this.moduleResultsOpen) {
      this.showModuleResults();
    }
  }

  /**
   * The module-scope match table: every match as a clickable line-and-preview row under the
   * inputs, refreshed by every re-find while open, so the rows always mirror the matches the
   * counter is counting. A row's click selects and reveals its match in the editor.
   */
  private showModuleResults(): void {
    if (this.findInput.value.length === 0) {
      // A Find All with nothing typed is a no-op; the emptied-query reset lives in
      // findInModule, which every query change runs through.
      this.moduleResultsOpen = false;
      return;
    }

    this.results.replaceChildren();
    this.results.hidden = false;

    if (this.matches.length === 0) {
      this.results.appendChild(this.note("No matches."));
      return;
    }

    const shown = Math.min(this.matches.length, 500);
    this.results.appendChild(this.note(
      `${this.matches.length} match${this.matches.length === 1 ? "" : "es"}`
      + (shown < this.matches.length ? ` (showing the first ${shown})` : "")));

    const model = this.editor.getModel();
    for (let i = 0; i < shown; i++) {
      const match = this.matches[i];
      if (!match) {
        continue;
      }
      const range = match.range;

      const row = document.createElement("button");
      row.type = "button";
      row.className = "search-row";
      row.setAttribute("role", "listitem");

      const where = document.createElement("span");
      where.className = "search-line";
      where.textContent = String(range.startLineNumber);
      row.appendChild(where);

      const preview = document.createElement("span");
      preview.className = "search-preview";
      const line = model?.getLineContent(range.startLineNumber) ?? "";
      preview.textContent = line.trim();
      preview.title = line;
      row.appendChild(preview);

      const index = i;
      row.addEventListener("click", () => {
        // The table re-renders with every re-find, so the index matches the live list
        // unless an edit landed in the same beat; a stale click is dropped, not misaimed.
        if (this.scope() === "module" && index < this.matches.length) {
          this.current = index;
          this.step(0);
        }
      });

      this.results.appendChild(row);
    }
  }

  /** The first match at or after the position, wrapping to the first match. */
  private nearestMatch(position: monaco.Position): number {
    for (let i = 0; i < this.matches.length; i++) {
      if (this.matches[i]?.range.getStartPosition().isBefore(position) === false) {
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
    const range = this.matches[this.current]?.range;
    if (!range) {
      return;
    }
    this.editor.setSelection(range);
    this.editor.revealRangeInCenterIfOutsideViewport(range);
    this.decorate();
    this.updateCounter();
  }

  private replaceCurrent(): void {
    if (this.scope() !== "module" || this.current < 0 || this.matches.length === 0) {
      return;
    }

    const range = this.matches[this.current]?.range;
    if (!range) {
      return;
    }
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
