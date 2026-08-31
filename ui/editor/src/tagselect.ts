/*
 * The tag facet: which tags' tests the Tests pane is showing, chosen from what discovery found.
 *
 * A FACET, NOT A SCOPE. The scope pair narrows to one file or one module - a single choice -
 * where tags are worn many to a test and pressed many at a time, so this is a multi-select: a
 * button wearing the active state, a popup of checkboxes with counts, and "(untagged)" for the
 * tests carrying no tag at all. Checked tags are a union, the same union the outcome chips
 * beside it mean, and nothing checked means the facet is off and everything shows - which is
 * why there is no "all tags" row: clearing IS that.
 *
 * The option list is rebuilt from the scoped rows on every paint, so it follows the session as
 * files and modules come and go - and a chosen tag whose last test left the scope stays listed
 * at (0) rather than vanishing: an active filter that cannot be seen cannot be cleared, and a
 * hidden filter is how a pane lies about an empty list.
 */

/** One tag the facet can press, with how many scoped tests wear it right now. */
export interface TagEntry {
  name: string;
  count: number;
}

/** The wire spelling of "(untagged)": the shim reserves this word to mean tests with no tags. */
const UNTAGGED = "untagged";

export class TagSelect {
  /** The button and its popup, to be seated in a toolbar as one thing. */
  readonly element: HTMLElement;

  private readonly button: HTMLButtonElement;
  private readonly label: HTMLElement;
  private readonly popup: HTMLElement;
  private readonly onChange: () => void;

  /** Chosen tags, case-folded key to the casing the author wrote. */
  private readonly chosen = new Map<string, string>();
  private untaggedChosen = false;

  private entries: TagEntry[] = [];
  private untaggedCount = 0;

  /** What the popup was last built from, so an unchanged render leaves it alone under the pointer. */
  private signature = "";

  constructor(buttonId: string, onChange: () => void) {
    this.onChange = onChange;

    this.element = document.createElement("span");
    this.element.className = "tag-select";

    this.button = document.createElement("button");
    this.button.type = "button";
    this.button.id = buttonId;
    this.button.setAttribute("aria-haspopup", "true");
    this.button.setAttribute("aria-expanded", "false");
    const icon = document.createElement("span");
    icon.className = "codicon codicon-tag";
    icon.setAttribute("aria-hidden", "true");
    this.label = document.createElement("span");
    this.label.className = "tests-label";
    this.button.append(icon, this.label);
    this.button.addEventListener("click", () => this.setOpen(Boolean(this.popup.hidden)));

    this.popup = document.createElement("div");
    this.popup.className = "tag-select-popup";
    this.popup.setAttribute("role", "group");
    this.popup.setAttribute("aria-label", "Filter tests by tag");
    this.popup.hidden = true;
    this.popup.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.stopPropagation();
        this.setOpen(false);
        this.button.focus();
      }
    });

    // Anywhere else closes it. On pointerdown rather than click for the same reason the
    // outcome chips press on pointerdown: the pane repaints under the pointer, and a repaint
    // between press and release breaks the browser's click pairing.
    document.addEventListener("pointerdown", (event) => {
      if (!this.popup.hidden && event.target instanceof Node && !this.element.contains(event.target)) {
        this.setOpen(false);
      }
    });

    this.element.append(this.button, this.popup);
    this.setEntries([], 0);
  }

  /**
   * Rebuilds the option list from the rows the pane is scoping to. Chosen tags the scope no
   * longer holds stay listed at (0) so they remain clearable; the whole control hides when the
   * view holds no tags and nothing is chosen, because then there is nothing to filter by.
   */
  setEntries(entries: TagEntry[], untaggedCount: number): void {
    const merged = new Map<string, TagEntry>();
    for (const entry of entries) {
      merged.set(entry.name.toLowerCase(), entry);
    }

    for (const [key, name] of this.chosen) {
      if (!merged.has(key)) {
        merged.set(key, { name, count: 0 });
      }
    }

    this.entries = [...merged.values()].sort((a, b) => a.name.localeCompare(b.name));
    this.untaggedCount = untaggedCount;
    this.element.hidden = this.entries.length === 0 && !this.untaggedChosen;

    const signature = JSON.stringify([
      this.entries.map((entry) => [entry.name, entry.count]),
      untaggedCount,
      [...this.chosen.keys()].sort(),
      this.untaggedChosen,
    ]);
    if (signature !== this.signature) {
      this.signature = signature;
      this.buildPopup();
    }

    this.dress();
  }

  /** True while any tag - "(untagged)" included - is narrowing the list. */
  active(): boolean {
    return this.chosen.size > 0 || this.untaggedChosen;
  }

  /** True when a row wearing these tags belongs in the filtered view. */
  admits(tags: readonly string[]): boolean {
    if (!this.active()) {
      return true;
    }

    if (this.untaggedChosen && tags.length === 0) {
      return true;
    }

    return tags.some((tag) => this.chosen.has(tag.toLowerCase()));
  }

  /** Whether this one tag is chosen, for a row chip's pressed state. */
  has(tag: string): boolean {
    return this.chosen.has(tag.toLowerCase());
  }

  /** The facet as the wire spells it: a comma list, "untagged" standing in for "(untagged)". */
  wire(): string {
    const parts = [...this.chosen.values()];
    if (this.untaggedChosen) {
      parts.push(UNTAGGED);
    }

    return parts.join(",");
  }

  /** What the facet is called in a sentence, for titles and empty states. */
  words(): string {
    const parts = [...this.chosen.values()];
    if (this.untaggedChosen) {
      parts.push("(untagged)");
    }

    return `tag${parts.length === 1 ? "" : "s"} ${parts.join(", ")}`;
  }

  /** Presses one tag in or out - the row chips' own gesture. */
  toggle(tag: string): void {
    const key = tag.toLowerCase();
    if (this.chosen.has(key)) {
      this.chosen.delete(key);
    } else {
      this.chosen.set(key, tag);
    }

    this.onChange();
  }

  /** Back to everything, from the popup's Clear and from the empty state's way out. */
  clear(): void {
    if (!this.active()) {
      return;
    }

    this.chosen.clear();
    this.untaggedChosen = false;
    this.onChange();
  }

  private setOpen(open: boolean): void {
    this.popup.hidden = !open;
    this.button.setAttribute("aria-expanded", String(open));
    if (open) {
      this.popup.querySelector("input")?.focus();
    }
  }

  private buildPopup(): void {
    // A toggle rebuilds this list under the keyboard, so whichever row held focus gets it back.
    const focused = document.activeElement instanceof HTMLInputElement
      ? document.activeElement.dataset.tag ?? null
      : null;

    const rows: HTMLElement[] = [];
    for (const entry of this.entries) {
      rows.push(this.row(entry.name, entry.count, this.chosen.has(entry.name.toLowerCase()), () => {
        this.toggle(entry.name);
      }));
    }

    if (this.untaggedCount > 0 || this.untaggedChosen) {
      rows.push(this.row("(untagged)", this.untaggedCount, this.untaggedChosen, () => {
        this.untaggedChosen = !this.untaggedChosen;
        this.onChange();
      }));
    }

    if (this.active()) {
      const clear = document.createElement("button");
      clear.type = "button";
      clear.className = "tag-select-clear";
      clear.textContent = "Clear";
      clear.addEventListener("click", () => this.clear());
      rows.push(clear);
    }

    this.popup.replaceChildren(...rows);

    if (focused !== null && !this.popup.hidden) {
      this.popup.querySelector<HTMLInputElement>(`input[data-tag="${CSS.escape(focused)}"]`)?.focus();
    }
  }

  private row(name: string, count: number, checked: boolean, toggle: () => void): HTMLElement {
    const label = document.createElement("label");
    label.className = "tag-select-row";

    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = checked;
    box.dataset.tag = name;
    box.addEventListener("change", toggle);

    const said = document.createElement("span");
    said.className = "tag-select-name";
    said.textContent = name;

    const tally = document.createElement("span");
    tally.className = "tag-select-count";
    tally.textContent = `(${count})`;

    label.append(box, said, tally);
    return label;
  }

  /** The button says what the facet is doing, so a narrowed pane reads narrowed at a glance. */
  private dress(): void {
    const picked = this.chosen.size + (this.untaggedChosen ? 1 : 0);
    const offered = this.entries.length + (this.untaggedCount > 0 ? 1 : 0);
    this.label.textContent = picked === 0
      ? "Tags"
      : picked === 1
        ? `Tags: ${this.untaggedChosen && this.chosen.size === 0 ? "(untagged)" : [...this.chosen.values()][0]}`
        : `Tags: ${picked} of ${Math.max(offered, picked)}`;
    this.button.classList.toggle("tag-narrowed", this.active());
    this.button.title = this.active()
      ? `Showing only tests with ${this.words()}. Click to change or clear.`
      : "Filter the list by tag";
  }
}
