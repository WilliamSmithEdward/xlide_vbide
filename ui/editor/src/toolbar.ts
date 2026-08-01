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
}

/**
 * Editor action identifiers are Monaco's own. They are strings rather than an enumeration in
 * Monaco itself, and an unknown one is ignored rather than reported, so each is spelled once here.
 */
export const COMMANDS: ToolbarCommand[] = [
  { id: "run", target: "host", icon: "play", label: "Run (F5)" },
  { id: "break", target: "host", icon: "debug-pause", label: "Break (Ctrl+F5)" },
  { id: "reset", target: "host", icon: "debug-stop", label: "Reset (Shift+F5)" },

  { id: "stepInto", target: "host", icon: "debug-step-into", label: "Step into (F8)", separatorBefore: true },
  { id: "stepOver", target: "host", icon: "debug-step-over", label: "Step over (Shift+F8)" },
  { id: "stepOut", target: "host", icon: "debug-step-out", label: "Step out" },
  { id: "toggleBreakpoint", target: "host", icon: "debug-breakpoint", label: "Toggle breakpoint (F9)" },

  { id: "actions.find", target: "editor", icon: "search", label: "Find (Ctrl+F)", separatorBefore: true },
  { id: "editor.action.startFindReplaceAction", target: "editor", icon: "replace", label: "Replace (Ctrl+H)" },
  { id: "editor.action.quickCommand", target: "editor", icon: "list-flat", label: "Command palette (F1)" },
  { id: "editor.action.gotoLine", target: "editor", icon: "go-to-file", label: "Go to line (Ctrl+G)" },

  { id: "editor.action.commentLine", target: "editor", icon: "comment", label: "Toggle comment (Ctrl+/)", separatorBefore: true },
  { id: "editor.action.indentLines", target: "editor", icon: "arrow-right", label: "Indent" },
  { id: "editor.action.outdentLines", target: "editor", icon: "arrow-left", label: "Outdent" },
  { id: "editor.action.formatDocument", target: "editor", icon: "symbol-namespace", label: "Format module" },

  { id: "editor.foldAll", target: "editor", icon: "fold", label: "Fold all", separatorBefore: true },
  { id: "editor.unfoldAll", target: "editor", icon: "unfold", label: "Unfold all" },
];

export function buildToolbar(root: HTMLElement, run: (command: ToolbarCommand) => void): void {
  root.replaceChildren();

  for (const command of COMMANDS) {
    if (command.separatorBefore) {
      const divider = document.createElement("span");
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

    const icon = document.createElement("span");
    icon.className = `codicon codicon-${command.icon}`;
    icon.setAttribute("aria-hidden", "true");

    button.appendChild(icon);
    button.addEventListener("click", () => run(command));
    root.appendChild(button);
  }
}
