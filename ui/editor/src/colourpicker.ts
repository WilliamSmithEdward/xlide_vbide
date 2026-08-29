/*
 * xlide's own colour picker, for the Properties panel's colour rows.
 *
 * WHY NOT THE PLATFORM'S. `<input type="color">` opens Windows' own colour dialog: a modal in
 * another visual language, in the middle of a surface built to replace exactly that, and it can
 * only ever say `#rrggbb`. Half of what a form's colour rows hold is not an RGB at all - a system
 * colour is a QUESTION ("what does this machine call a button face"), and the native dialog has
 * nowhere to put one.
 *
 * So this picker has two halves. The palette is a ramp of ordinary colours, generated rather than
 * written down, so the grid is inspectable and evenly spaced. The System half is the host's -
 * only it can ask Windows what a button face is today - and picking one writes the QUESTION back,
 * not the answer, which is what keeps a form themed rather than frozen.
 *
 * One picker exists at a time, it closes on Escape, on a click outside, and on a scroll of the
 * panel underneath it, and every choice commits through the same path a typed value takes.
 */

/** One system colour as the host measured it: what it is called, the value a property takes, and
 * what this machine resolves it to right now. */
export interface SystemColour {
  name: string;
  value: string;
  css: string;
}

let systemList: SystemColour[] = [];

/** The host's answer, held for the page's life (setSystemColours, sent at load). */
export function setSystemColours(colours: SystemColour[]): void {
  systemList = colours;
}

/**
 * The palette, generated from one ramp so it is even and inspectable: a row of greys, then eight
 * hues down seven lightnesses. Written as code rather than as sixty-four literals because a
 * hand-typed grid is a grid with a typo in it, and nothing here is anybody's brand palette.
 */
const GREYS = ["#000000", "#404040", "#808080", "#a6a6a6", "#c0c0c0", "#d9d9d9", "#f0f0f0", "#ffffff"];
const HUES = [0, 30, 60, 120, 180, 210, 270, 330];
const LIGHTNESS = [22, 32, 42, 52, 62, 74, 86];

function hsl(hue: number, saturation: number, lightness: number): string {
  const a = (saturation * Math.min(lightness, 100 - lightness)) / 100;
  const channel = (n: number): string => {
    const k = (n + hue / 30) % 12;
    const value = lightness - a * Math.max(-1, Math.min(k - 3, Math.min(9 - k, 1)));
    return Math.round((value * 255) / 100).toString(16).padStart(2, "0");
  };

  return `#${channel(0)}${channel(8)}${channel(4)}`;
}

function palette(): string[] {
  const rows = [GREYS];
  for (const lightness of LIGHTNESS) {
    rows.push(HUES.map((hue) => hsl(hue, 68, lightness)));
  }

  return rows.flat();
}

/** What a probe can see of the open picker: which row it belongs to, what it holds, and how many
 * choices each half offers. Null when none is open. */
export interface ColourPickerState {
  property: string;
  value: string;
  palette: number;
  system: number;
}

let open: {
  root: HTMLElement;
  property: string;
  value: string;
  choose: (spelling: string) => void;
} | null = null;

export function colourPickerState(): ColourPickerState | null {
  return open === null
    ? null
    : {
      property: open.property,
      value: open.value,
      palette: open.root.querySelectorAll(".colour-swatch").length,
      system: open.root.querySelectorAll(".colour-system-row").length,
    };
}

/** Picks by SPELLING - a `#rrggbb` from the palette or a system colour's name - through the same
 * click the hand makes. Answers what it pressed, or null when the open picker has no such
 * choice, which is the honest answer for a probe aiming at a colour the palette does not carry. */
export function pickColour(spelling: string): string | null {
  if (open === null) {
    return null;
  }

  const wanted = spelling.trim().toLowerCase();
  const target = [...open.root.querySelectorAll<HTMLElement>(".colour-swatch, .colour-system-row")]
    .find((one) => (one.dataset.value ?? "").toLowerCase() === wanted);
  if (!target) {
    return null;
  }

  target.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, cancelable: true }));
  return target.dataset.value ?? null;
}

/** Runs after the open picker actually closes - the Properties panel flushes a deferred
 * republish through it, so a picker browsed mid-analysis is not torn down under the hand. */
let afterClose: (() => void) | null = null;

export function onColourPickerClosed(hook: () => void): void {
  afterClose = hook;
}

export function closeColourPicker(): void {
  if (open === null) {
    return;
  }

  open.root.remove();
  open = null;
  document.removeEventListener("pointerdown", onPointerDown, true);
  document.removeEventListener("keydown", onKeyDown, true);
  afterClose?.();
}

function onPointerDown(event: PointerEvent): void {
  if (open && !open.root.contains(event.target as Node)
    && !(event.target as HTMLElement)?.classList?.contains("prop-swatch")) {
    closeColourPicker();
  }
}

function onKeyDown(event: KeyboardEvent): void {
  if (event.key === "Escape" && open) {
    event.preventDefault();
    event.stopPropagation();
    closeColourPicker();
  }
}

/**
 * Opens the picker under an anchor. `value` is what the row holds, in the panel's own spelling -
 * a `#rrggbb` or a system colour's name - and `onPick` gets the spelling chosen, which the row
 * then commits exactly as though it had been typed.
 */
export function openColourPicker(options: {
  anchor: HTMLElement;
  property: string;
  value: string;
  /** Opened from the keyboard: the grid takes focus, so the arrows have somewhere to start. */
  focusFirst?: boolean;
  onPick: (spelling: string) => void;
}): void {
  // Measured BEFORE the close below: closing an earlier picker flushes any deferred panel
  // republish, which rebuilds the rows and detaches this anchor - a rect read after that is
  // all zeros and the picker lands in the corner. The replacement row keeps this geometry.
  const box = options.anchor.getBoundingClientRect();

  closeColourPicker();

  const root = document.createElement("div");
  root.className = "colour-picker";
  root.setAttribute("role", "dialog");
  root.setAttribute("aria-label", `Color for ${options.property}`);

  // The palette: a grid the arrow keys walk, where every cell says its own colour out loud for
  // anyone who cannot see it.
  const grid = document.createElement("div");
  grid.className = "colour-grid";
  grid.setAttribute("role", "listbox");
  grid.setAttribute("aria-label", "Palette");

  const current = options.value.trim().toLowerCase();
  for (const colour of palette()) {
    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = "colour-swatch";
    cell.dataset.value = colour;
    cell.style.background = colour;
    cell.title = colour;
    cell.setAttribute("role", "option");
    cell.setAttribute("aria-label", colour);
    cell.setAttribute("aria-selected", String(colour === current));
    if (colour === current) {
      cell.classList.add("current");
    }

    cell.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      choose(colour);
    });
    grid.appendChild(cell);
  }

  // The System half, which is the reason this picker exists rather than the platform's.
  const system = document.createElement("div");
  system.className = "colour-system";
  system.setAttribute("role", "listbox");
  system.setAttribute("aria-label", "System colors");

  for (const colour of systemList) {
    const row = document.createElement("button");
    row.type = "button";
    row.className = "colour-system-row";
    row.dataset.value = colour.name;
    row.setAttribute("role", "option");
    row.setAttribute("aria-selected", String(colour.name.toLowerCase() === current));
    if (colour.name.toLowerCase() === current) {
      row.classList.add("current");
    }

    const chip = document.createElement("span");
    chip.className = "colour-chip";
    chip.style.background = colour.css;

    const label = document.createElement("span");
    label.className = "colour-system-name";
    label.textContent = colour.name;

    // The value the row will take, said out loud: a system colour is a question, and the answer
    // this machine gives today is worth showing beside it without pretending to be the value.
    const resolved = document.createElement("span");
    resolved.className = "colour-system-css";
    resolved.textContent = colour.css;

    row.append(chip, label, resolved);
    row.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      choose(colour.name);
    });
    system.appendChild(row);
  }

  const heading = document.createElement("div");
  heading.className = "colour-heading";
  heading.textContent = "System";

  // And the way out of both lists: any colour at all, typed.
  const hexRow = document.createElement("div");
  hexRow.className = "colour-hex";
  const hex = document.createElement("input");
  hex.type = "text";
  hex.className = "colour-hex-field";
  hex.value = options.value;
  hex.spellcheck = false;
  hex.setAttribute("aria-label", `Color value for ${options.property}`);
  hex.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      choose(hex.value.trim());
    }
  });
  hexRow.appendChild(hex);

  root.append(grid, heading, system, hexRow);
  document.body.appendChild(root);

  // Under the swatch, and pulled back onto the screen when the row sits low in a short panel.
  // The box was measured at the top, before the close above could recycle the anchor's row.
  const width = root.offsetWidth || 220;
  const height = root.offsetHeight || 260;
  root.style.left = `${Math.max(4, Math.min(box.left - width + box.width, innerWidth - width - 4))}px`;
  root.style.top = box.bottom + height + 4 > innerHeight
    ? `${Math.max(4, box.top - height - 2)}px`
    : `${box.bottom + 2}px`;

  open = { root, property: options.property, value: options.value, choose: options.onPick };
  document.addEventListener("pointerdown", onPointerDown, true);
  document.addEventListener("keydown", onKeyDown, true);

  // Opened from the keyboard, the grid takes focus at whatever is currently set, so the arrows
  // start where the eye already is. Opened by a click, focus stays with the hand.
  if (options.focusFirst) {
    const start = root.querySelector<HTMLElement>(".colour-swatch.current")
      ?? root.querySelector<HTMLElement>(".colour-swatch");
    start?.focus();
  }

  // Arrow keys walk the grid the way they walk any other grid, eight to a row.
  grid.addEventListener("keydown", (event) => {
    const cells = [...grid.querySelectorAll<HTMLElement>(".colour-swatch")];
    const at = cells.indexOf(document.activeElement as HTMLElement);
    const step = event.key === "ArrowRight" ? 1
      : event.key === "ArrowLeft" ? -1
        : event.key === "ArrowDown" ? 8
          : event.key === "ArrowUp" ? -8
            : 0;
    if (step === 0 || at < 0) {
      return;
    }

    event.preventDefault();
    cells[Math.max(0, Math.min(cells.length - 1, at + step))]?.focus();
  });

  function choose(spelling: string): void {
    const picked = open;
    closeColourPicker();
    picked?.choose(spelling);
  }
}
