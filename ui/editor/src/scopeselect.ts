/*
 * The scope control: which file's, and which module's, rows a list pane is showing.
 *
 * Both list panes answer for a whole session - the analyzer reads every module of every open
 * project, because VBA compiles per project and a module the developer has never opened still
 * breaks the build, and the runner discovers every test in every open file. That breadth is
 * right and it is also a lot of rows, so each pane carries this to narrow the view without
 * narrowing what the product knows.
 *
 * TWO SELECTS, NOT ONE TREE (the owner, 2026-08-20). The two boundaries in a session are a
 * FILE - a workbook in Excel, a document in Word; "file" is the word that is true in every
 * host - and a MODULE inside one, and they are independent questions. One control offering
 * both as a grouped list makes the reader parse a hierarchy to answer either; two say plainly
 * which is being asked. The file select hides itself when a single file is open, because then
 * it is a choice between one thing and everything.
 *
 * The module list is the SELECTED FILE'S. Narrowing the file narrows what modules exist to
 * choose from, which is the whole reason the two sit together - and a module choice that the
 * new file does not hold falls back to All Modules rather than filtering against nothing.
 *
 * Both lists are rebuilt from the rows on every render, which is what makes them follow the
 * session as files and modules come and go.
 */

export const SCOPE_ALL = "@all";
export const SCOPE_CURRENT = "@current";
export const SCOPE_ALL_FILES = "@allfiles";

/** One module the control can narrow to, as the pane's rows know it. */
export interface ScopeEntry {
  /** Case-folded identity, unique across files. Rows are matched on this. */
  key: string;
  /** The module's own name, for a run target and for what the empty state calls the scope. */
  name: string;
  /** The file it lives in, as the rows name it. */
  file: string;
  /** How many rows it holds right now, shown beside the name. */
  count: number;
}

/** One open file the control can narrow to. */
export interface ScopeFile {
  name: string;
  count: number;
}

export type ScopeKind = "all" | "current" | "file" | "module";

export class ScopeSelect {
  /** The pair, to be put in a toolbar as one thing. */
  readonly element: HTMLElement;
  readonly files: HTMLSelectElement;
  readonly modules: HTMLSelectElement;

  private readonly noun: string;
  private entries: ScopeEntry[] = [];
  private known: ScopeFile[] = [];
  private currentKey: string | null = null;
  private currentName: string | null = null;
  private currentFile: string | null = null;

  /** What each list was last built from, so an unchanged render leaves an open popup alone. */
  private fileSignature = "";
  private moduleSignature = "";

  /**
   * @param id base for the two selects' ids: `<id>-file` and `<id>-module`.
   * @param noun what the counts are counting, said in the labels ("problems", "tests").
   */
  constructor(id: string, ariaLabel: string, noun: string, onChange: () => void) {
    this.noun = noun;

    this.element = document.createElement("span");
    this.element.className = "scope-select-pair";
    this.element.id = id;

    this.files = document.createElement("select");
    this.files.id = `${id}-file`;
    this.files.className = "scope-select scope-select-file";
    this.files.setAttribute("aria-label", `${ariaLabel}: file`);

    this.modules = document.createElement("select");
    this.modules.id = `${id}-module`;
    this.modules.className = "scope-select scope-select-module";
    this.modules.setAttribute("aria-label", `${ariaLabel}: module`);

    // A file choice changes WHICH MODULES EXIST, so the module list is rebuilt under it before
    // anyone is told the scope moved.
    this.files.addEventListener("change", () => {
      this.buildModules(true);
      this.mark();
      onChange();
    });

    this.modules.addEventListener("change", () => {
      this.mark();
      onChange();
    });

    this.element.append(this.files, this.modules);
    this.setEntries([], null, []);
  }

  /**
   * Rebuilds both lists from the rows the pane is holding. `current` is the active tab's
   * module, or null when the active tab is not a module the pane knows; `files` are the open
   * files in the order the pane lists them.
   */
  setEntries(entries: ScopeEntry[], current: { key: string; name: string; file: string } | null, files: ScopeFile[]): void {
    this.entries = entries;
    this.known = files;
    this.currentKey = current?.key ?? null;
    this.currentName = current?.name ?? null;
    this.currentFile = current?.file ?? null;

    const total = files.reduce((sum, file) => sum + file.count, 0);
    const fileSignature = JSON.stringify([total, files.map((file) => [file.name, file.count])]);
    if (fileSignature !== this.fileSignature) {
      this.fileSignature = fileSignature;
      const chosen = this.files.value || SCOPE_ALL_FILES;
      const keep = chosen === SCOPE_ALL_FILES || files.some((file) => fileKey(file.name) === chosen);

      const options = [option(SCOPE_ALL_FILES, `All Files (${total})`)];
      for (const file of files) {
        options.push(option(fileKey(file.name), `${file.name} (${file.count})`));
      }

      this.files.replaceChildren(...options);
      this.files.value = keep ? chosen : SCOPE_ALL_FILES;

      // One file is not a choice; the module select speaks for the whole session on its own.
      this.files.hidden = files.length < 2;
    }

    this.buildModules(false);
    this.mark();
  }

  /**
   * The module list for the file now chosen. `reset` is true when the FILE just changed, where
   * a module from the file being left has to go rather than silently filtering everything away.
   */
  private buildModules(reset: boolean): void {
    const file = this.files.value;
    const mine = file === SCOPE_ALL_FILES
      ? this.entries
      : this.entries.filter((entry) => fileKey(entry.file) === file);
    const total = mine.reduce((sum, entry) => sum + entry.count, 0);

    // A module name is only unique inside its file, so across files a shared name says which -
    // and inside one file it never has to.
    const seen = new Map<string, number>();
    for (const entry of mine) {
      const lower = entry.name.toLowerCase();
      seen.set(lower, (seen.get(lower) ?? 0) + 1);
    }

    const current = this.currentInScope();
    const signature = JSON.stringify([
      file, total, current?.key ?? "", current?.name ?? "", current?.count ?? -1,
      mine.map((entry) => [entry.key, entry.name, entry.file, entry.count]),
    ]);
    if (signature === this.moduleSignature && !reset) {
      return;
    }
    this.moduleSignature = signature;

    const chosen = this.modules.value || SCOPE_ALL;
    const keep = !reset
      && (chosen === SCOPE_ALL || chosen === SCOPE_CURRENT
        || mine.some((entry) => entry.key === chosen));

    const options = [option(SCOPE_ALL, `All Modules (${total})`)];
    if (current !== null) {
      options.push(option(SCOPE_CURRENT, `Current Module: ${current.name} (${current.count})`));
    }

    for (const entry of mine) {
      const shared = (seen.get(entry.name.toLowerCase()) ?? 0) > 1;
      options.push(option(entry.key, `${entry.name}${shared ? ` - ${entry.file}` : ""} (${entry.count})`));
    }

    this.modules.replaceChildren(...options);
    this.modules.value = keep ? chosen : SCOPE_ALL;
  }

  /** The active tab's module, when the chosen file holds it. */
  private currentInScope(): { key: string; name: string; count: number } | null {
    if (this.currentKey === null || this.currentName === null) {
      return null;
    }

    const file = this.files.value;
    if (file !== SCOPE_ALL_FILES && fileKey(this.currentFile ?? "") !== file) {
      return null;
    }

    const held = this.entries.find((entry) => entry.key === this.currentKey);
    return { key: this.currentKey, name: this.currentName, count: held?.count ?? 0 };
  }

  /**
   * A narrowed pane must SAY it is narrowed, or a developer who left it on one module reads an
   * empty list as a clean project. Whichever select is doing the narrowing wears it.
   */
  private mark(): void {
    const narrowedFile = this.files.value !== SCOPE_ALL_FILES;
    const narrowedModule = this.modules.value !== SCOPE_ALL;
    this.files.classList.toggle("scope-narrowed", narrowedFile);
    this.modules.classList.toggle("scope-narrowed", narrowedModule);
    this.files.title = narrowedFile
      ? `Showing ${this.noun} from ${this.fileName() ?? "one file"} only.`
      : `Showing ${this.noun} from every open file.`;
    this.modules.title = narrowedModule
      ? `Showing ${this.noun} from ${this.moduleName() ?? "one module"} only.`
      : `Showing ${this.noun} from every module in view.`;
  }

  /** True when a row with this module key, in this file, belongs in the current scope. */
  admits(key: string, file: string): boolean {
    if (this.files.value !== SCOPE_ALL_FILES && fileKey(file) !== this.files.value) {
      return false;
    }

    const module = this.modules.value;
    if (module === SCOPE_ALL) {
      return true;
    }

    if (module === SCOPE_CURRENT) {
      return this.currentKey !== null && key === this.currentKey;
    }

    return key === module;
  }

  /** True while every row shows - the pane is not narrowed at all. */
  showsEverything(): boolean {
    return this.files.value === SCOPE_ALL_FILES && this.modules.value === SCOPE_ALL;
  }

  /** Both choices in one string, for a caller's own "has anything changed" key. */
  stateKey(): string {
    return `${this.files.value}|${this.modules.value}`;
  }

  /** Which tier the scope is on, so a caller can target a run at the right thing. */
  scopeKind(): ScopeKind {
    if (this.modules.value === SCOPE_CURRENT) {
      return "current";
    }

    if (this.modules.value !== SCOPE_ALL) {
      return "module";
    }

    return this.files.value === SCOPE_ALL_FILES ? "all" : "file";
  }

  /**
   * What the scope is called, for the empty state and for a run target: the module's own name,
   * the file's name when only the file is narrowed, or null while the pane is unscoped.
   */
  scopeName(): string | null {
    switch (this.scopeKind()) {
      case "all":
        return null;
      case "file":
        return this.fileName();
      default:
        return this.moduleName();
    }
  }

  /**
   * The file the scope sits in: the chosen file, or the chosen module's own file, and null
   * while every file is showing. A run has to be told this, because a module name alone does
   * not say which file's copy to run.
   */
  scopeFile(): string | null {
    if (this.files.value !== SCOPE_ALL_FILES) {
      return this.fileName();
    }

    switch (this.scopeKind()) {
      case "current":
        return this.currentFile;
      case "module":
        return this.entries.find((entry) => entry.key === this.modules.value)?.file || null;
      default:
        return null;
    }
  }

  /** Back to everything, from the empty state's own way out. */
  reset(): void {
    this.files.value = SCOPE_ALL_FILES;
    this.buildModules(true);
    this.mark();
  }

  private fileName(): string | null {
    return this.known.find((file) => fileKey(file.name) === this.files.value)?.name ?? null;
  }

  private moduleName(): string | null {
    return this.modules.value === SCOPE_CURRENT
      ? this.currentName
      : this.entries.find((entry) => entry.key === this.modules.value)?.name ?? null;
  }
}

function option(value: string, label: string): HTMLOptionElement {
  const made = document.createElement("option");
  made.value = value;
  made.textContent = label;
  return made;
}

/** The option value for a file. */
function fileKey(name: string): string {
  return `file:${name.toLowerCase()}`;
}
