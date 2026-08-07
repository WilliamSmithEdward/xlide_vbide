/**
 * The About dialog: what this is, which build is running, and where to go next.
 *
 * The version matters more than it looks. Every question that starts "it used to work" is really
 * a question about which build is loaded, and a surface that cannot answer that makes the person
 * asking go and read a registry key.
 */

declare const __XLIDE_BUILD__: string;
declare const __XLIDE_VERSION__: string;
declare const __XLIDE_BUILD_NUMBER__: number;

const REPOSITORY = "https://github.com/WilliamSmithEdward/xlide_vbide";

/**
 * Where the add-in is loaded from. Only the host knows, and it says so once the surface is ready,
 * which can be after someone has already opened this dialog. Held here rather than passed through
 * every caller, because it is one fact about the session and not an argument to anything.
 */
let installPath: string | null = null;

export function setInstallPath(path: string | null): void {
  installPath = path;
  const shown = document.getElementById("help-install-path");
  if (shown) {
    describeInstall(shown);
  }
}

/**
 * Fills an element with the path and, under it, what kind of build lives there. The kind is on its
 * own line rather than in brackets after the path: a development publish is long enough that the
 * bracket wrapped alone onto the next line and read as something that had come apart.
 */
function describeInstall(into: HTMLElement): void {
  into.replaceChildren();

  if (!installPath) {
    into.textContent = "not reported";
    return;
  }

  const where = document.createElement("span");
  where.className = "help-path";
  where.textContent = installPath;
  into.appendChild(where);

  // The installer puts the product under Programs\xlide; anything else was published by hand.
  if (!/\\Programs\\xlide$/i.test(installPath)) {
    const kind = document.createElement("span");
    kind.className = "help-kind";
    kind.textContent = "development build";
    into.appendChild(kind);
  }
}

interface Shortcut {
  keys: string;
  what: string;
}

const SHORTCUTS: Shortcut[] = [
  { keys: "Ctrl+\\", what: "Split the editor" },
  { keys: "Ctrl+F", what: "Find in this module" },
  { keys: "Ctrl+Shift+F", what: "Find across the workbook" },
  { keys: "F2", what: "Object browser" },
  { keys: "F5", what: "Run" },
  { keys: "F8", what: "Step into" },
  { keys: "F9", what: "Toggle a breakpoint" },
];

/**
 * Opens the dialog. One at a time, like the others: opening while open focuses what is there
 * rather than stacking a second card nobody asked for.
 */
export function openHelpDialog(closed?: () => void): void {
  const existing = document.getElementById("help-card");
  if (existing) {
    existing.querySelector<HTMLElement>("#help-close")?.focus();
    return;
  }

  const backdrop = document.createElement("div");
  backdrop.id = "help-backdrop";

  const card = document.createElement("div");
  card.id = "help-card";
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-modal", "true");
  card.setAttribute("aria-label", "About xlide");

  const head = document.createElement("div");
  head.id = "help-head";

  const title = document.createElement("div");
  title.id = "help-title";

  const name = document.createElement("span");
  name.id = "help-name";
  name.textContent = "XLIDE";

  const version = document.createElement("span");
  version.id = "help-version";
  version.textContent = __XLIDE_VERSION__;

  title.append(name, version);

  const close = document.createElement("button");
  close.type = "button";
  close.id = "help-close";
  close.title = "Close (Esc)";
  close.setAttribute("aria-label", "Close");
  close.innerHTML = '<span class="codicon codicon-close" aria-hidden="true"></span>';

  head.append(title, close);

  const blurb = document.createElement("p");
  blurb.className = "help-blurb";
  blurb.textContent =
    "A modern editing surface for VBA, inside the Visual Basic Editor. The native editor keeps " +
    "running underneath as the text of record, the compile target, and the debugger.";

  const facts = document.createElement("dl");
  facts.className = "help-facts";

  for (const [label, value] of [
    ["Version", __XLIDE_VERSION__],
    ["Build", `${__XLIDE_BUILD_NUMBER__}`],
    ["Built", __XLIDE_BUILD__.replace("T", " ")],
  ] as const) {
    const term = document.createElement("dt");
    term.textContent = label;
    const detail = document.createElement("dd");
    detail.textContent = value;
    facts.append(term, detail);
  }

  const whereTerm = document.createElement("dt");
  whereTerm.textContent = "Loaded from";
  const whereDetail = document.createElement("dd");
  whereDetail.id = "help-install-path";
  describeInstall(whereDetail);
  facts.append(whereTerm, whereDetail);

  const keysTitle = document.createElement("div");
  keysTitle.className = "help-section";
  keysTitle.textContent = "Keys worth knowing";

  const keys = document.createElement("dl");
  keys.className = "help-keys";

  for (const shortcut of SHORTCUTS) {
    const term = document.createElement("dt");
    const kbd = document.createElement("kbd");
    kbd.textContent = shortcut.keys;
    term.appendChild(kbd);

    const detail = document.createElement("dd");
    detail.textContent = shortcut.what;
    keys.append(term, detail);
  }

  const foot = document.createElement("div");
  foot.id = "help-foot";

  // Rendered as text rather than an anchor: the page runs under a policy that forbids navigating
  // anywhere, so a link would look like a link and do nothing when pressed.
  const repo = document.createElement("span");
  repo.className = "help-repo";
  repo.textContent = REPOSITORY;

  const copy = document.createElement("button");
  copy.type = "button";
  copy.className = "help-copy";
  copy.textContent = "Copy";
  copy.title = "Copy the address";
  copy.addEventListener("click", () => {
    void navigator.clipboard?.writeText(REPOSITORY).then(
      () => {
        copy.textContent = "Copied";
        window.setTimeout(() => (copy.textContent = "Copy"), 1200);
      },
      () => {
        copy.textContent = "Press Ctrl+C";
      },
    );
  });

  foot.append(repo, copy);

  card.append(head, blurb, facts, keysTitle, keys, foot);
  backdrop.appendChild(card);
  document.body.appendChild(backdrop);

  const dismiss = () => {
    document.removeEventListener("keydown", onKey, true);
    backdrop.remove();
    closed?.();
  };

  function onKey(event: KeyboardEvent): void {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  }

  close.addEventListener("click", dismiss);
  backdrop.addEventListener("mousedown", (event) => {
    if (event.target === backdrop) dismiss();
  });

  // Captured, because Monaco answers Escape too and would otherwise swallow it first.
  document.addEventListener("keydown", onKey, true);

  close.focus();
}
