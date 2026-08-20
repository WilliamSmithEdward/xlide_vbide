/*
 * The scope selector: which module's rows a list pane is showing.
 *
 * Both list panes answer for a whole workspace - the analyzer reads every module of every open
 * project, because VBA compiles per project and a module the developer has never opened still
 * breaks the build, and the runner discovers every test in the project. That breadth is right
 * and it is also a lot of rows, so each pane carries this control to narrow the view to one
 * module without narrowing what the product knows.
 *
 * Three kinds of scope, and the middle one is the reason this is a selector rather than a
 * filter box: ALL is the default, CURRENT MODULE follows the active tab and so needs no
 * maintenance as tabs open and close, and a named module holds still while the developer works
 * elsewhere. The option list is rebuilt from the rows on every render, which is what makes it
 * follow the workspace; a named scope whose module has left the list falls back to All rather
 * than filtering against a module that is no longer there.
 */

export const SCOPE_ALL = "@all";
export const SCOPE_CURRENT = "@current";

/** One module the selector can narrow to, as the pane's rows know it. */
export interface ScopeEntry {
  /** Case-folded identity, unique across workbooks. Rows are matched on this. */
  key: string;
  /** What the option reads - the workbook named only where the bare name is ambiguous. */
  label: string;
  /** The module's own name, for a run target and for what the empty state calls the scope. */
  name: string;
  /** How many rows it holds right now, shown beside the label. */
  count: number;
}

export class ScopeSelect {
  readonly element: HTMLSelectElement;
  private readonly noun: string;
  private currentKey: string | null = null;
  private currentName: string | null = null;
  private entries: ScopeEntry[] = [];

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
    this.setEntries([], null);
  }

  /**
   * Rebuilds the option list from the rows the pane is holding, keeping the developer's choice
   * where the module is still listed. `current` is the active tab's module, or null when the
   * active tab is not a module the pane knows.
   */
  setEntries(entries: ScopeEntry[], current: { key: string; name: string } | null): void {
    this.entries = entries;
    this.currentKey = current?.key ?? null;
    this.currentName = current?.name ?? null;

    const total = entries.reduce((sum, entry) => sum + entry.count, 0);
    const currentEntry = current ? entries.find((entry) => entry.key === current.key) : undefined;
    const signature = JSON.stringify([
      total, current?.key ?? "", current?.name ?? "", currentEntry?.count ?? -1,
      entries.map((entry) => [entry.key, entry.label, entry.count]),
    ]);
    if (signature === this.signature) {
      return;
    }
    this.signature = signature;

    // The chosen scope survives the rebuild when its module is still listed; a module that has
    // gone - its workbook closed - takes the pane back to All rather than leaving it filtered
    // against nothing, which reads as an empty pane with no cause.
    const chosen = this.element.value || SCOPE_ALL;
    const keep = chosen === SCOPE_ALL || chosen === SCOPE_CURRENT
      || entries.some((entry) => entry.key === chosen);

    const options: HTMLOptionElement[] = [];
    const add = (value: string, label: string) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      options.push(option);
    };

    add(SCOPE_ALL, `All Modules (${total})`);
    add(SCOPE_CURRENT, current
      ? `Current Module: ${current.name} ${currentEntry ? `(${currentEntry.count})` : `(no ${this.noun})`}`
      : "Current Module");
    for (const entry of entries) {
      add(entry.key, `${entry.label} (${entry.count})`);
    }

    this.element.replaceChildren(...options);
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
      : `Showing ${this.noun} from every module. Narrow to one module here.`;
  }

  /** True when a row belonging to this module key belongs in the current scope. */
  admits(key: string): boolean {
    const value = this.element.value;
    if (value === SCOPE_ALL) {
      return true;
    }

    if (value === SCOPE_CURRENT) {
      return this.currentKey !== null && key === this.currentKey;
    }

    return key === value;
  }

  /** True while every row shows - the pane is not narrowed at all. */
  showsEverything(): boolean {
    return this.element.value === SCOPE_ALL;
  }

  /**
   * What the scope is called, for the empty state and for a run target: the module's own name,
   * or null while the pane is unscoped. A Current Module scope with no module under it answers
   * null too, and `showsEverything` is how a caller tells those two apart.
   */
  scopeName(): string | null {
    const value = this.element.value;
    if (value === SCOPE_ALL) {
      return null;
    }

    if (value === SCOPE_CURRENT) {
      return this.currentName;
    }

    return this.entries.find((entry) => entry.key === value)?.name ?? null;
  }

  /** Back to everything, from the empty state's own way out. */
  reset(): void {
    this.element.value = SCOPE_ALL;
    this.markNarrowed();
  }
}
