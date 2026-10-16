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

import { installEdgeScroll, type EdgeScroll } from "./edgescroll.js";

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
  /**
   * Which file the command is about, when it is about one. The VBE's own References and
   * Project Properties dialogs act on the ACTIVE project, so a command raised from a
   * project's row in the tree has to say which project was clicked or it opens the
   * dialog for whatever tab happens to be in front (the owner, 2026-08-20). Toolbar
   * commands never set this; they mean "whatever is in front" by definition.
   */
  project?: string;
  /**
   * The group the command is drawn in, a key of GROUPS. Consecutive commands of one group share a
   * single button on the strip whose rows open beneath it; see buildGroup. A grouped command has
   * no needsBreak: its row is drawn when the group opens, not kept on the strip to be greyed.
   */
  group?: string;
}

/** The face of each command group: the icon on the strip, and the name it is announced by. */
export const GROUPS: Record<string, { icon: string; label: string }> = {
  bookmarks: { icon: "bookmark", label: "Bookmarks" },
};

/**
 * Editor action identifiers are Monaco's own. They are strings rather than an enumeration in
 * Monaco itself, and an unknown one is ignored rather than reported, so each is spelled once here.
 */
export const COMMANDS: ToolbarCommand[] = [
  { id: "save", target: "host", icon: "save", label: "Save (Ctrl+S)" },
  { id: "undo", target: "editor", icon: "discard", label: "Undo (Ctrl+Z)" },
  { id: "redo", target: "editor", icon: "redo", label: "Redo (Ctrl+Y)" },

  // Compile leads the run cluster the way it leads the editor's own Debug menu - the menu
  // that is gone as of 2026-08-05, its commands rehomed here.
  { id: "compile", target: "host", icon: "check-all", label: "Compile project", separatorBefore: true },
  { id: "run", target: "host", icon: "play", label: "Run (F5)" },
  { id: "break", target: "host", icon: "debug-pause", label: "Break (Ctrl+F5)" },
  { id: "reset", target: "host", icon: "debug-stop", label: "Reset (Shift+F5)" },
  // The Run menu's last item without a home, rehomed here the day that menu went (2026-08-09).
  // A toggle: pressing it again is where you started, which is why the label says so rather than
  // naming a direction the button cannot know it is going in.
  { id: "designMode", target: "host", icon: "edit", label: "Toggle design mode" },

  { id: "stepInto", target: "host", icon: "debug-step-into", label: "Step into (F8)", separatorBefore: true },
  { id: "stepOver", target: "host", icon: "debug-step-over", label: "Step over (Shift+F8)" },
  { id: "stepOut", target: "host", icon: "debug-step-out", label: "Step out (Ctrl+Shift+F8)" },
  // The rest of the Debug menu's stepping half, greyed outside a break like the Call Stack
  // button: each is disabled in the editor when nothing is stopped, and a click into
  // silence reads as a defect.
  { id: "runToCursor", target: "host", icon: "debug-continue", label: "Run to cursor (break mode)", needsBreak: true },
  { id: "setNextStatement", target: "host", icon: "arrow-circle-right", label: "Set next statement (Ctrl+F9, break mode)", needsBreak: true },
  { id: "showNextStatement", target: "host", icon: "target", label: "Show next statement (break mode)", needsBreak: true },
  { id: "toggleBreakpoint", target: "host", icon: "debug-breakpoint", label: "Toggle breakpoint (F9)" },
  { id: "clearAllBreakpoints", target: "host", icon: "clear-all", label: "Clear all breakpoints" },
  // The View menu's Call Stack, rehomed when the menu went (2026-08-05): a break-mode
  // dialog belongs beside the stepping it narrates - and greyed anywhere else, because
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

  // The native Edit toolbar's bookmark four, which the surface hides with the rest of that bar
  // (#111). One button rather than four: the strip is wide already, and these are reached for
  // in bursts rather than one at a time between other commands.
  { id: "xlide.bookmark.toggle", target: "editor", icon: "bookmark", label: "Toggle bookmark (Ctrl+Alt+K)", group: "bookmarks", separatorBefore: true },
  { id: "xlide.bookmark.next", target: "editor", icon: "arrow-down", label: "Next bookmark (Ctrl+Alt+N)", group: "bookmarks" },
  { id: "xlide.bookmark.previous", target: "editor", icon: "arrow-up", label: "Previous bookmark (Ctrl+Alt+P)", group: "bookmarks" },
  { id: "xlide.bookmark.clearAll", target: "editor", icon: "clear-all", label: "Clear all bookmarks (Ctrl+Alt+Shift+K)", group: "bookmarks" },

  { id: "editor.foldAll", target: "editor", icon: "fold", label: "Fold all", separatorBefore: true },
  { id: "editor.unfoldAll", target: "editor", icon: "unfold", label: "Unfold all" },

  { id: "objectBrowser", target: "host", icon: "library", label: "Object browser (F2)", separatorBefore: true },

  { id: "openSync", target: "editor", icon: "repo-sync", label: "Import and export modules", separatorBefore: true },
  { id: "openAgent", target: "editor", icon: "robot", label: "Work with an AI agent" },
  { id: "openPanes", target: "editor", icon: "layout", label: "Panes", separatorBefore: true },
  { id: "restorePaneDefaults", target: "editor", icon: "discard", label: "Restore default pane layout" },
  { id: "openAnalysisRules", target: "editor", icon: "checklist", label: "Analyzer rules" },
  { id: "openSettings", target: "editor", icon: "settings-gear", label: "Settings" },
  { id: "openHelp", target: "editor", icon: "question", label: "About xlide" },
  { id: "openSponsor", target: "editor", icon: "heart", label: "Support xlide" },
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
  // The menu button and the wordmark are the toolbar's fixed ends and they are NOT ours to
  // rebuild: the menu bar owns its own button and rebuilds it whenever the host's menus change,
  // and this runs again on every settings change. Taken out before the wipe and put back around
  // the strip, so the row reads menu, commands, wordmark at every width.
  const lead = root.querySelector("#menubar");
  const trail = root.querySelector("#brand");
  closeGroup();
  root.replaceChildren();

  // The strip is one row that slides. A pane half a screen wide cannot show thirty commands, and
  // the alternatives all cost something worse: clipping takes them away with nothing to say they
  // existed, a menu hides them behind a second decision, and wrapping spends a row of a pane whose
  // height is the point. Chevrons appear only when there is something past the edge.
  const strip = document.createElement("div");
  strip.className = "toolbar-strip";

  if (lead) {
    root.append(lead);
  }
  root.append(strip);
  if (trail) {
    root.append(trail);
  }

  const missing: string[] = [];
  // The group the last drawn button opens, so the commands after it in that group join it.
  let group = null as { name: string; items: ToolbarCommand[] } | null;

  for (const command of COMMANDS) {
    if (command.target === "editor" && !available(command)) {
      missing.push(command.id);
      continue;
    }

    if (command.group !== undefined && command.group === group?.name) {
      group.items.push(command);
      continue;
    }

    if (command.separatorBefore) {
      const divider = document.createElement("span");
      divider.className = "toolbar-divider";
      // Decorative: the grouping is already conveyed by the labels.
      divider.setAttribute("aria-hidden", "true");
      strip.appendChild(divider);
    }

    if (command.group !== undefined) {
      group = { name: command.group, items: [command] };
      strip.appendChild(buildGroup(group.name, group.items, run));
      continue;
    }
    group = null;

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
    strip.appendChild(button);
  }

  // The edges, and everything that makes the strip scroll. Shared with the tab strip so the two
  // cannot end up with different wheel or press-and-hold behaviour.
  scrollers.get(root)?.dispose();
  scrollers.set(root, installEdgeScroll(strip));

  if (missing.length > 0) {
    console.warn(`[xlide] toolbar commands not available in this build: ${missing.join(", ")}`);
  }
}

/** One scroller per toolbar, replaced whenever it is rebuilt, so rebuilds do not accumulate them. */
const scrollers = new WeakMap<HTMLElement, EdgeScroll>();

/*
 * COMMAND GROUPS. One button on the strip stands for several commands, and its rows open beneath
 * it: on hover after a short pause, so a pointer crossing the strip does not throw groups open;
 * on a click, for touch and for anyone who would rather; and on keyboard focus, with Down, Enter
 * or Space moving into the rows. Choosing a row runs it and closes the group.
 *
 * The rows are drawn in the page body rather than in the strip, because the strip clips
 * vertically (it scrolls sideways), and they are drawn when the group opens, so a rebuilt
 * toolbar never leaves a stale set behind.
 */

/** Hover pause before a group opens, and the grace a pointer gets to cross from face to rows. */
const OPEN_DELAY_MS = 150;
const CLOSE_DELAY_MS = 300;

/** Each group button's commands and runner, read when it opens. */
const groups = new WeakMap<HTMLButtonElement, { name: string; items: ToolbarCommand[]; run: (command: ToolbarCommand) => void }>();

/** The open group, if any: one at a time. */
let openGroupState: { button: HTMLButtonElement; flyout: HTMLElement; dispose: () => void } | null = null;
let openTimer = 0;
let closeTimer = 0;
/** Set while Escape hands the focus back to a group's face, so that focus does not reopen it. */
let returningFocus = false;

function buildGroup(name: string, items: ToolbarCommand[], run: (command: ToolbarCommand) => void): HTMLButtonElement {
  const face = GROUPS[name] ?? { icon: items[0]?.icon ?? "ellipsis", label: name };

  const button = document.createElement("button");
  button.type = "button";
  button.className = "toolbar-button toolbar-group";
  button.dataset.group = name;
  button.title = face.label;
  button.setAttribute("aria-label", face.label);
  button.setAttribute("aria-haspopup", "menu");
  button.setAttribute("aria-expanded", "false");

  const icon = document.createElement("span");
  icon.className = `codicon codicon-${face.icon}`;
  icon.setAttribute("aria-hidden", "true");
  button.appendChild(icon);

  groups.set(button, { name, items, run });

  button.addEventListener("pointerenter", (event) => {
    if (event.pointerType === "touch") {
      return;
    }
    window.clearTimeout(closeTimer);
    window.clearTimeout(openTimer);
    openTimer = window.setTimeout(() => openGroup(button, false), OPEN_DELAY_MS);
  });
  button.addEventListener("pointerleave", () => {
    window.clearTimeout(openTimer);
    scheduleClose(button);
  });

  // A click opens and never closes: a pointer that rested long enough to open the group and then
  // clicked would otherwise shut what it just opened. Enter and Space arrive as a click with no
  // pointer behind it (detail 0), and those carry the focus into the rows.
  button.addEventListener("click", (event) => {
    window.clearTimeout(openTimer);
    openGroup(button, event.detail === 0);
  });
  button.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      openGroup(button, true);
    } else if (event.key === "Escape" && openGroupState?.button === button) {
      event.preventDefault();
      event.stopPropagation();
      closeGroup();
    }
  });
  button.addEventListener("focus", () => {
    if (!returningFocus && button.matches(":focus-visible")) {
      openGroup(button, false);
    }
  });
  button.addEventListener("focusout", (event) => closeIfFocusLeft(button, event.relatedTarget));

  return button;
}

function openGroup(button: HTMLButtonElement, focusFirst: boolean): void {
  const group = groups.get(button);
  if (!group) {
    return;
  }

  window.clearTimeout(closeTimer);
  if (openGroupState?.button === button) {
    if (focusFirst) {
      openGroupState.flyout.querySelector<HTMLButtonElement>(".menu-item")?.focus();
    }
    return;
  }
  closeGroup();

  const flyout = document.createElement("div");
  flyout.className = "menu-dropdown toolbar-flyout";
  flyout.dataset.group = group.name;
  flyout.setAttribute("role", "menu");
  flyout.setAttribute("aria-label", button.getAttribute("aria-label") ?? group.name);

  for (const command of group.items) {
    // The strip's labels carry their key in brackets, which is how a hover teaches it; in a row the
    // key is drawn apart, the way every other menu here draws it.
    const [, caption = command.label, key] = /^(.*?)(?: \(([^)]*)\))?$/.exec(command.label) ?? [];

    const row = document.createElement("button");
    row.type = "button";
    row.className = "menu-item toolbar-flyout-item";
    row.dataset.command = command.id;
    row.setAttribute("role", "menuitem");
    row.tabIndex = -1;
    row.title = command.label;

    const icon = document.createElement("span");
    icon.className = `toolbar-flyout-icon codicon codicon-${command.icon}`;
    icon.setAttribute("aria-hidden", "true");

    const text = document.createElement("span");
    text.className = "menu-caption";
    text.textContent = caption;

    row.append(icon, text);
    if (key) {
      const shortcut = document.createElement("span");
      shortcut.className = "menu-shortcut";
      shortcut.textContent = key;
      row.appendChild(shortcut);
    }

    row.addEventListener("pointerenter", () => row.focus({ preventScroll: true }));
    row.addEventListener("click", () => {
      closeGroup();
      group.run(command);
    });
    flyout.appendChild(row);
  }

  flyout.addEventListener("pointerenter", () => window.clearTimeout(closeTimer));
  flyout.addEventListener("pointerleave", () => scheduleClose(button));
  flyout.addEventListener("focusout", (event) => closeIfFocusLeft(button, event.relatedTarget));
  flyout.addEventListener("keydown", (event) => {
    const rows = [...flyout.querySelectorAll<HTMLButtonElement>(".menu-item")];
    const at = rows.indexOf(document.activeElement as HTMLButtonElement);

    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      const delta = event.key === "ArrowDown" ? 1 : -1;
      rows[(at + delta + rows.length) % rows.length]?.focus();
    } else if (event.key === "Home" || event.key === "End") {
      rows[event.key === "Home" ? 0 : rows.length - 1]?.focus();
    } else if (event.key === "Escape") {
      closeGroup();
      // Back on the face, which a keyboard focus would otherwise open straight away again.
      returningFocus = true;
      button.focus();
      returningFocus = false;
    } else {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
  });

  document.body.appendChild(flyout);

  // Under its button, pulled back on-screen when the button sits near the right edge.
  const box = button.getBoundingClientRect();
  flyout.style.left = `${Math.max(4, Math.min(box.left, window.innerWidth - flyout.offsetWidth - 4))}px`;
  flyout.style.top = `${box.bottom + 2}px`;

  // Anything that moves the button out from under the rows closes them: the strip sliding, the
  // window resizing, the page losing the focus to the host.
  const strip = button.closest(".toolbar-strip");
  const onPointerDown = (event: PointerEvent): void => {
    const target = event.target as Node;
    if (!flyout.contains(target) && !button.contains(target)) {
      closeGroup();
    }
  };
  const onMove = (): void => closeGroup();
  document.addEventListener("pointerdown", onPointerDown, true);
  strip?.addEventListener("scroll", onMove);
  window.addEventListener("resize", onMove);
  window.addEventListener("blur", onMove);

  button.setAttribute("aria-expanded", "true");
  button.classList.add("open");
  openGroupState = {
    button,
    flyout,
    dispose: () => {
      document.removeEventListener("pointerdown", onPointerDown, true);
      strip?.removeEventListener("scroll", onMove);
      window.removeEventListener("resize", onMove);
      window.removeEventListener("blur", onMove);
      flyout.remove();
      button.setAttribute("aria-expanded", "false");
      button.classList.remove("open");
    },
  };

  if (focusFirst) {
    flyout.querySelector<HTMLButtonElement>(".menu-item")?.focus();
  }
}

/** Closes the open group, if any. */
export function closeGroup(): void {
  window.clearTimeout(openTimer);
  window.clearTimeout(closeTimer);
  const state = openGroupState;
  openGroupState = null;
  state?.dispose();
}

function scheduleClose(button: HTMLButtonElement): void {
  window.clearTimeout(closeTimer);
  closeTimer = window.setTimeout(() => {
    // Kept open while the keyboard is in it: the pointer wandering off is not the developer leaving.
    const state = openGroupState;
    if (state?.button === button && !state.flyout.contains(document.activeElement)
      && document.activeElement !== button) {
      closeGroup();
    }
  }, CLOSE_DELAY_MS);
}

function closeIfFocusLeft(button: HTMLButtonElement, next: EventTarget | null): void {
  const state = openGroupState;
  if (state?.button !== button) {
    return;
  }
  const node = next as Node | null;
  if (!node || (!state.flyout.contains(node) && !button.contains(node))) {
    closeGroup();
  }
}

/** What the strip offers, grouped commands included, each with the group it is reached through. */
export function toolbarEntries(root: HTMLElement): { id: string; label: string; disabled: boolean; group?: string }[] {
  const entries: { id: string; label: string; disabled: boolean; group?: string }[] = [];
  for (const button of root.querySelectorAll<HTMLButtonElement>("button[data-command], button[data-group]")) {
    const group = groups.get(button);
    if (group) {
      for (const command of group.items) {
        entries.push({ id: command.id, label: command.label, disabled: button.disabled, group: group.name });
      }
    } else {
      entries.push({ id: button.dataset.command ?? "", label: button.title, disabled: button.disabled });
    }
  }
  return entries;
}

/**
 * Presses a grouped command the way a person does: opens its group and clicks the row. Answers
 * null when no group on the strip holds the command.
 */
export function pressGroupedCommand(root: HTMLElement, id: string): { did: boolean; detail: string } | null {
  const wanted = id.trim().toLowerCase();
  for (const button of root.querySelectorAll<HTMLButtonElement>("button[data-group]")) {
    const group = groups.get(button);
    if (!group?.items.some((command) => command.id.toLowerCase() === wanted)) {
      continue;
    }
    openGroup(button, false);
    const row = [...(openGroupState?.flyout.querySelectorAll<HTMLButtonElement>(".menu-item") ?? [])]
      .find((one) => (one.dataset.command ?? "").toLowerCase() === wanted);
    if (!row) {
      closeGroup();
      return { did: false, detail: `the ${group.name} group opened without a row for ${id}` };
    }
    row.click();
    return { did: true, detail: `pressed ${id} in the ${group.name} group` };
  }
  return null;
}

/** Opens a group as a hover would and answers its rows, for looking at it. */
export function openToolbarGroup(root: HTMLElement, name: string): { did: boolean; detail: string; rows: string[] } {
  const button = [...root.querySelectorAll<HTMLButtonElement>("button[data-group]")]
    .find((one) => one.dataset.group === name);
  if (!button) {
    return { did: false, detail: `the toolbar has no ${name} group`, rows: [] };
  }
  openGroup(button, false);
  const rows = [...(openGroupState?.flyout.querySelectorAll<HTMLElement>(".menu-item") ?? [])]
    .map((row) => row.dataset.command ?? "");
  return { did: true, detail: `opened the ${name} group`, rows };
}

