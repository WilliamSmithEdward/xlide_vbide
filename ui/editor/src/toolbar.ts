/*
 * The command bar above the editor.
 *
 * Two kinds of command live here and they are dispatched differently. Running, stepping and
 * breaking belong to the host: it owns the interpreter, and it is driven through its own menus so
 * that everything depending on it keeps working. Finding, commenting, indenting and formatting
 * belong to the editor and are run in the page, because that is where the text is.
 *
 * The distinction is in the data rather than in the caller, so adding a command is one entry.
 */

export type CommandTarget = "host" | "editor";

export interface ToolbarCommand {
  /** Stable name, sent to the host or used as the editor action id. */
  id: string;
  /** Where the command runs. */
  target: CommandTarget;
  /** Codicon name, from the icon font Monaco already ships. */
  icon: string;
  /** Spoken and hovered label. Includes the key, because that is how it is learned. */
  label: string;
  /** Draws a divider before this one. */
  separatorBefore?: boolean;
  /** Only means something in break mode; drawn disabled until the host says stopped. */
  needsBreak?: boolean;
}

/**
 * Editor action identifiers are Monaco's own. They are strings rather than an enumeration in
 * Monaco itself, and an unknown one is ignored rather than reported, so each is spelled once here.
 */
export const COMMANDS: ToolbarCommand[] = [
  { id: "save", target: "host", icon: "save", label: "Save (Ctrl+S)" },
  { id: "undo", target: "editor", icon: "discard", label: "Undo (Ctrl+Z)" },
  { id: "redo", target: "editor", icon: "redo", label: "Redo (Ctrl+Y)" },

  // Compile leads the run cluster the way it leads the editor's own Debug menu — the menu
  // that is gone as of 2026-08-05, its commands rehomed here.
  { id: "compile", target: "host", icon: "check-all", label: "Compile project", separatorBefore: true },
  { id: "run", target: "host", icon: "play", label: "Run (F5)" },
  { id: "break", target: "host", icon: "debug-pause", label: "Break (Ctrl+F5)" },
  { id: "reset", target: "host", icon: "debug-stop", label: "Reset (Shift+F5)" },

  { id: "stepInto", target: "host", icon: "debug-step-into", label: "Step into (F8)", separatorBefore: true },
  { id: "stepOver", target: "host", icon: "debug-step-over", label: "Step over (Shift+F8)" },
  { id: "stepOut", target: "host", icon: "debug-step-out", label: "Step out" },
  // The rest of the Debug menu's stepping half, greyed outside a break like the Call Stack
  // button: each is disabled in the editor when nothing is stopped, and a click into
  // silence reads as a defect.
  { id: "runToCursor", target: "host", icon: "debug-continue", label: "Run to cursor (break mode)", needsBreak: true },
  { id: "setNextStatement", target: "host", icon: "arrow-circle-right", label: "Set next statement (break mode)", needsBreak: true },
  { id: "showNextStatement", target: "host", icon: "target", label: "Show next statement (break mode)", needsBreak: true },
  { id: "toggleBreakpoint", target: "host", icon: "debug-breakpoint", label: "Toggle breakpoint (F9)" },
  { id: "clearAllBreakpoints", target: "host", icon: "clear-all", label: "Clear all breakpoints" },
  // The View menu's Call Stack, rehomed when the menu went (2026-08-05): a break-mode
  // dialog belongs beside the stepping it narrates — and greyed anywhere else, because
  // the native command is disabled outside a break and a click into silence reads as a
  // defect ("won't appear again", 2026-08-05).
  { id: "callStack", target: "host", icon: "list-tree", label: "Call stack (break mode)", needsBreak: true },

  { id: "xlide.search.open", target: "editor", icon: "search", label: "Find (Ctrl+F)", separatorBefore: true },
  { id: "xlide.search.replace", target: "editor", icon: "replace", label: "Replace (Ctrl+H)" },
  { id: "editor.action.quickCommand", target: "editor", icon: "list-flat", label: "Command palette (F1)" },
  { id: "editor.action.gotoLine", target: "editor", icon: "go-to-file", label: "Go to line (Ctrl+G)" },

  // The comment suite, the way the native editor's Edit toolbar offered it: comment and
  // uncomment as their own buttons for block work, the toggle for the quick single line.
  { id: "editor.action.addCommentLine", target: "editor", icon: "comment", label: "Comment lines", separatorBefore: true },
  { id: "editor.action.removeCommentLine", target: "editor", icon: "comment-draft", label: "Uncomment lines" },
  { id: "editor.action.commentLine", target: "editor", icon: "comment-discussion", label: "Toggle comment (Ctrl+/)" },
  { id: "editor.action.indentLines", target: "editor", icon: "arrow-right", label: "Indent" },
  { id: "editor.action.outdentLines", target: "editor", icon: "arrow-left", label: "Outdent" },
  { id: "editor.action.formatDocument", target: "editor", icon: "symbol-namespace", label: "Format module" },

  { id: "editor.foldAll", target: "editor", icon: "fold", label: "Fold all", separatorBefore: true },
  { id: "editor.unfoldAll", target: "editor", icon: "unfold", label: "Unfold all" },

  { id: "objectBrowser", target: "host", icon: "library", label: "Object browser (F2)", separatorBefore: true },

  { id: "openPanes", target: "editor", icon: "layout", label: "Panes", separatorBefore: true },
  { id: "openSettings", target: "editor", icon: "settings-gear", label: "Settings" },
  { id: "openHelp", target: "editor", icon: "question", label: "About xlide" },
];

/**
 * Builds the toolbar.
 *
 * `available` is asked about every editor command before its button is drawn. Monaco resolves an
 * action id at run time and ignores an unknown one rather than reporting it, so a button for a
 * feature that was not bundled looks identical to one that works and does nothing when pressed.
 * Commands that cannot run are left out and named in the console, which is the difference between
 * a toolbar and a row of decorations.
 */
export function buildToolbar(
  root: HTMLElement,
  run: (command: ToolbarCommand) => void,
  available: (command: ToolbarCommand) => boolean = () => true,
): void {
  root.replaceChildren();

  const missing: string[] = [];
  const entries: { command: ToolbarCommand; button: HTMLButtonElement; divider: HTMLElement | null }[] = [];

  for (const command of COMMANDS) {
    if (command.target === "editor" && !available(command)) {
      missing.push(command.id);
      continue;
    }

    let divider: HTMLElement | null = null;
    if (command.separatorBefore) {
      divider = document.createElement("span");
      divider.className = "toolbar-divider";
      // Decorative: the grouping is already conveyed by the labels.
      divider.setAttribute("aria-hidden", "true");
      root.appendChild(divider);
    }

    const button = document.createElement("button");
    button.type = "button";
    button.className = "toolbar-button";
    button.dataset.command = command.id;
    button.title = command.label;
    // The icon is decorative; the label is what carries the meaning.
    button.setAttribute("aria-label", command.label);

    // Disabled until the host says break mode: the shell flips these with setDebugState.
    if (command.needsBreak) {
      button.dataset.needsBreak = "1";
      button.disabled = true;
    }

    const icon = document.createElement("span");
    icon.className = `codicon codicon-${command.icon}`;
    icon.setAttribute("aria-hidden", "true");

    button.appendChild(icon);
    button.addEventListener("click", () => run(command));
    root.appendChild(button);
    entries.push({ command, button, divider });
  }

  // The overflow control, last and only drawn when it is needed. The pane a developer works in is
  // often half a screen wide, and a strip that simply clips its right-hand end takes commands away
  // with nothing to say they existed. This keeps them one press away instead.
  const overflow = document.createElement("button");
  overflow.type = "button";
  overflow.className = "toolbar-button toolbar-overflow";
  overflow.title = "More commands";
  overflow.setAttribute("aria-label", "More commands");
  overflow.setAttribute("aria-haspopup", "true");
  overflow.hidden = true;
  overflow.innerHTML = '<span class="codicon codicon-more" aria-hidden="true"></span>';
  overflow.addEventListener("click", () => openOverflowMenu(overflow, hidden(), run));
  root.appendChild(overflow);

  function hidden(): ToolbarCommand[] {
    return entries.filter((entry) => entry.button.hidden).map((entry) => entry.command);
  }

  const relayout = (): void => {
    // Measured from a clean slate every time, because the answer when the pane grows is not
    // derivable from the answer when it shrank.
    for (const entry of entries) {
      entry.button.hidden = false;
      if (entry.divider) entry.divider.hidden = false;
    }
    overflow.hidden = true;

    if (root.scrollWidth <= root.clientWidth) {
      return;
    }

    overflow.hidden = false;

    // Drop from the right, which keeps the commands in the order they were designed in and takes
    // the least used away first.
    for (let index = entries.length - 1; index >= 0; index--) {
      if (root.scrollWidth <= root.clientWidth) break;
      const entry = entries[index]!;
      entry.button.hidden = true;
      if (entry.divider) entry.divider.hidden = true;
    }

    // A divider whose whole group went with it would otherwise be left leading the strip.
    for (const entry of entries) {
      if (!entry.divider) continue;
      entry.divider.hidden = entry.button.hidden;
    }
  };

  relayout();

  // The pane is resized by dragging a splitter, not only by resizing the window, so the element
  // itself is what has to be watched.
  observers.get(root)?.disconnect();
  const observer = new ResizeObserver(() => relayout());
  observer.observe(root);
  observers.set(root, observer);

  if (missing.length > 0) {
    console.warn(`[xlide] toolbar commands not available in this build: ${missing.join(", ")}`);
  }
}

/** One observer per toolbar, replaced whenever it is rebuilt, so rebuilds do not accumulate them. */
const observers = new WeakMap<HTMLElement, ResizeObserver>();

/**
 * The commands that did not fit, as a menu under the overflow button. Built from the live buttons
 * rather than from the command list, so a command the host has greyed out for break mode is greyed
 * out here too rather than offering a press that does nothing.
 */
function openOverflowMenu(
  anchor: HTMLElement,
  commands: ToolbarCommand[],
  run: (command: ToolbarCommand) => void,
): void {
  const existing = document.getElementById("toolbar-overflow-menu");
  if (existing) {
    existing.remove();
    return;
  }

  const menu = document.createElement("div");
  menu.id = "toolbar-overflow-menu";
  // The surface's own dropdown chrome, so this is the menu the developer already knows.
  menu.className = "menu-dropdown";
  menu.setAttribute("role", "menu");
  menu.setAttribute("aria-label", "More commands");

  const dismiss = (): void => {
    menu.remove();
    document.removeEventListener("pointerdown", onPointerDown, true);
    document.removeEventListener("keydown", onKey, true);
  };

  function onPointerDown(event: PointerEvent): void {
    if (!menu.contains(event.target as Node) && !anchor.contains(event.target as Node)) {
      dismiss();
    }
  }

  function onKey(event: KeyboardEvent): void {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      dismiss();
    }
  }

  for (const command of commands) {
    const source = document.querySelector<HTMLButtonElement>(`#toolbar [data-command="${command.id}"]`);

    const item = document.createElement("button");
    item.type = "button";
    item.className = "menu-item" + (source?.disabled ? " disabled" : "");
    item.setAttribute("role", "menuitem");
    item.disabled = source?.disabled ?? false;

    const icon = document.createElement("span");
    icon.className = `codicon codicon-${command.icon}`;
    icon.setAttribute("aria-hidden", "true");

    const caption = document.createElement("span");
    caption.className = "menu-caption";
    caption.textContent = command.label;

    item.append(icon, caption);
    item.addEventListener("click", () => {
      dismiss();
      run(command);
    });

    menu.appendChild(item);
  }

  document.body.appendChild(menu);

  // Under its button, pulled back on-screen when the button sits near the right edge.
  const box = anchor.getBoundingClientRect();
  const width = menu.offsetWidth;
  menu.style.left = `${Math.max(4, Math.min(box.left, window.innerWidth - width - 4))}px`;
  menu.style.top = `${box.bottom + 2}px`;

  document.addEventListener("pointerdown", onPointerDown, true);
  document.addEventListener("keydown", onKey, true);
  menu.querySelector<HTMLButtonElement>("button:not([disabled])")?.focus();
}
