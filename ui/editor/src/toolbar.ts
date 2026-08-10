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
  // The Run menu's last item without a home, rehomed here the day that menu went (2026-08-09).
  // A toggle: pressing it again is where you started, which is why the label says so rather than
  // naming a direction the button cannot know it is going in.
  { id: "designMode", target: "host", icon: "edit", label: "Toggle design mode" },

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

  { id: "openSync", target: "editor", icon: "repo-sync", label: "Import and export modules", separatorBefore: true },
  { id: "openPanes", target: "editor", icon: "layout", label: "Panes", separatorBefore: true },
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

  for (const command of COMMANDS) {
    if (command.target === "editor" && !available(command)) {
      missing.push(command.id);
      continue;
    }

    if (command.separatorBefore) {
      const divider = document.createElement("span");
      divider.className = "toolbar-divider";
      // Decorative: the grouping is already conveyed by the labels.
      divider.setAttribute("aria-hidden", "true");
      strip.appendChild(divider);
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

