/*
 * The scope selector: which file's, and which module's, rows a list pane is showing.
 *
 * Both list panes answer for a whole session - the analyzer reads every module of every open
 * project, because VBA compiles per project and a module the developer has never opened still
 * breaks the build, and the runner discovers every test in every open file. That breadth is
 * right and it is also a lot of rows, so each pane carries this control to narrow the view
 * without narrowing what the product knows.
 *
 * Two tiers, because a session has two natural boundaries. A FILE - a workbook in Excel, a
 * document in Word; the word "file" is the one that is true in every host - and a MODULE
 * inside one. Modules are grouped under their file, so a name that appears in two files is
 * two entries rather than an ambiguity, and the file tier only appears when there is more than
 * one file to choose between: with a single file open it would restate All Modules.
 *
 * CURRENT MODULE follows the active tab and so needs no maintenance as tabs open and close.
 * The option list is rebuilt from the rows on every render, which is what makes it follow the
 * session; a named scope whose module or file has left the list falls back to All rather than
 * filtering against something that is no longer there.
 */

export const SCOPE_ALL = "@all";
export const SCOPE_CURRENT = "@current";

/** The value prefix that marks a whole-file scope, so file and module keys cannot collide. */
const FILE_PREFIX = "file:";

/** One module the selector can narrow to, as the pane's rows know it. */
export interface ScopeEntry {
  /** Case-folded identity, unique across files. Rows are matched on this. */
  key: string;
  /** What the option reads under its file's heading - the module's name. */
  label: string;
  /** The module's own name, for a run target and for what the empty state calls the scope. */
  name: string;
  /** The file it lives in, as the rows name it. Empty when the pane has no file dimension. */
  file: string;
  /** How many rows it holds right now, shown beside the label. */
  count: number;
}

/** One open file the selector can narrow to whole. */
export interface ScopeFile {
  /** The file's name as the rows carry it. */
  name: string;
  /** How many rows it holds across all its modules. */
  count: number;
}

export type ScopeKind = "all" | "current" | "file" | "module";

export class ScopeSelect {
  readonly element: HTMLSelectElement;
  private readonly noun: string;
  private currentKey: string | null = null;
  private currentName: string | null = null;
  private currentFile: string | null = null;
  private entries: ScopeEntry[] = [];
  private files: ScopeFile[] = [];

  /** What the option list was last built from, so an unchanged render leaves the popup alone. */
  private signature = "";

  /**
   * @param noun what the counts are counting, said in the option labels ("problems", "tests").
   * @param onChange fires when the developer picks a different scope.
   */
  constructor(id: string, ariaLabel: string, noun: string, onChange: () => void) {
    this.noun = noun;
    this.element = document.createElement("select");
    this.element.id = id;
    this.element.className = "scope-select";
    this.element.setAttribute("aria-label", ariaLabel);
    this.element.addEventListener("change", () => {
      this.markNarrowed();
      onChange();
    });
    this.setEntries([], null, []);
  }

  /**
   * Rebuilds the option list from the rows the pane is holding, keeping the developer's choice
   * where it still exists. `current` is the active tab's module, or null when the active tab is
   * not a module the pane knows; `files` are the open files in the order the pane lists them.
   */
  setEntries(entries: ScopeEntry[], current: { key: string; name: string; file: string } | null, files: ScopeFile[]): void {
    this.entries = entries;
    this.files = files;
    this.currentKey = current?.key ?? null;
    this.currentName = current?.name ?? null;
    this.currentFile = current?.file ?? null;

    const total = entries.reduce((sum, entry) => sum + entry.count, 0);
    const currentEntry = current ? entries.find((entry) => entry.key === current.key) : undefined;
    const signature = JSON.stringify([
      total, current?.key ?? "", current?.name ?? "", currentEntry?.count ?? -1,
      files.map((file) => [file.name, file.count]),
      entries.map((entry) => [entry.key, entry.label, entry.file, entry.count]),
    ]);
    if (signature === this.signature) {
      return;
    }
    this.signature = signature;

    // The chosen scope survives the rebuild when it is still listed; one that has gone - its
    // file closed, its module deleted - takes the pane back to All rather than leaving it
    // filtered against nothing, which reads as an empty pane with no cause.
    const chosen = this.element.value || SCOPE_ALL;
    const keep = chosen === SCOPE_ALL || chosen === SCOPE_CURRENT
      || entries.some((entry) => entry.key === chosen)
      || files.some((file) => fileKey(file.name) === chosen);

    const option = (value: string, label: string) => {
      const made = document.createElement("option");
      made.value = value;
      made.textContent = label;
      return made;
    };

    const children: HTMLElement[] = [
      option(SCOPE_ALL, `All Modules (${total})`),
      option(SCOPE_CURRENT, current
        ? `Current Module: ${current.name} ${currentEntry ? `(${currentEntry.count})` : `(no ${this.noun})`}`
        : "Current Module"),
    ];

    // ONE FILE, NO FILE TIER. A lone "TestFixture.xlsm (10)" beside "All Modules (10)" offers a
    // choice between two spellings of the same thing.
    if (files.length > 1) {
      for (const file of files) {
        const group = document.createElement("optgroup");
        group.label = file.name;
        group.appendChild(option(fileKey(file.name), `All of ${file.name} (${file.count})`));
        for (const entry of entries.filter((one) => sameFile(one.file, file.name))) {
          group.appendChild(option(entry.key, `${entry.label} (${entry.count})`));
        }

        children.push(group);
      }
    } else {
      for (const entry of entries) {
        children.push(option(entry.key, `${entry.label} (${entry.count})`));
      }
    }

    this.element.replaceChildren(...children);
    this.element.value = keep ? chosen : SCOPE_ALL;
    this.markNarrowed();
  }

  /**
   * A narrowed pane must SAY it is narrowed, or a developer who left it on one module reads an
   * empty list as a clean project. The border carries it, and the option text says which.
   */
  private markNarrowed(): void {
    const narrowed = this.element.value !== SCOPE_ALL;
    this.element.classList.toggle("scope-narrowed", narrowed);
    this.element.title = narrowed
      ? `Showing ${this.noun} from ${this.scopeName() ?? "one module"} only.`
      : `Showing ${this.noun} from every module of every open file. Narrow it here.`;
  }

  /** True when a row with this module key, in this file, belongs in the current scope. */
  admits(key: string, file: string): boolean {
    const value = this.element.value;
    if (value === SCOPE_ALL) {
      return true;
    }

    if (value === SCOPE_CURRENT) {
      return this.currentKey !== null && key === this.currentKey;
    }

    if (value.startsWith(FILE_PREFIX)) {
      return fileKey(file) === value;
    }

    return key === value;
  }

  /** True while every row shows - the pane is not narrowed at all. */
  showsEverything(): boolean {
    return this.element.value === SCOPE_ALL;
  }

  /** Which tier the scope is on, so a caller can target a run at the right thing. */
  scopeKind(): ScopeKind {
    const value = this.element.value;
    if (value === SCOPE_ALL) {
      return "all";
    }

    if (value === SCOPE_CURRENT) {
      return "current";
    }

    return value.startsWith(FILE_PREFIX) ? "file" : "module";
  }

  /**
   * What the scope is called, for the empty state and for a run target: the module's own name,
   * the file's name for a whole-file scope, or null while the pane is unscoped. A Current
   * Module scope with no module under it answers null too, and `showsEverything` is how a
   * caller tells those two apart.
   */
  scopeName(): string | null {
    switch (this.scopeKind()) {
      case "all":
        return null;
      case "current":
        return this.currentName;
      case "file":
        return this.files.find((file) => fileKey(file.name) === this.element.value)?.name ?? null;
      default:
        return this.entries.find((entry) => entry.key === this.element.value)?.name ?? null;
    }
  }

  /**
   * The file the scope sits in: a module scope's own file, the file itself for a file scope,
   * the active tab's file for Current Module, and null while unscoped. A run has to be told
   * this, because a module name alone does not say which file's copy to run.
   */
  scopeFile(): string | null {
    switch (this.scopeKind()) {
      case "all":
        return null;
      case "current":
        return this.currentFile;
      case "file":
        return this.scopeName();
      default:
        return this.entries.find((entry) => entry.key === this.element.value)?.file || null;
    }
  }

  /** Back to everything, from the empty state's own way out. */
  reset(): void {
    this.element.value = SCOPE_ALL;
    this.markNarrowed();
  }
}

/** The option value for a whole file. */
function fileKey(name: string): string {
  return `${FILE_PREFIX}${name.toLowerCase()}`;
}

function sameFile(left: string, right: string): boolean {
  return left.toLowerCase() === right.toLowerCase();
}
