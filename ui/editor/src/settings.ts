/*
 * The developer's settings, as the page holds them.
 *
 * The host owns the truth: it loads the file, sends the settings at ready, and echoes every
 * change back once it is written. The page applies what arrives and asks for what the dialog
 * changes — it never assumes a change took until the echo lands, which is what makes the
 * dialog's state the file's state and not a hope about it.
 */

export interface EditorSettings {
  /** Smart Enter block layout: spacer lines around the body, or body against the closer. */
  blockLayout: "comfy" | "compact";
  /** Enter at the end of a whole-line comment continues the apostrophes. */
  continueCommentOnNewline: boolean;
  /** A continued comment also mirrors the spaces after the apostrophes. */
  mirrorCommentSpacing: boolean;
  /**
   * The tree follows the editor: the module being worked on unfolds its procedures, and
   * everything folds away when the last tab closes. Off leaves the tree to the hand that opened
   * it.
   */
  treeFollowsEditor: boolean;
  /** One indent level, in spaces. Governs typing, smart Enter, and Format Module alike. */
  formatIndentSize: number;
  /** Which planner decides what an import or export will do: "xlide" or "builtIn". */
  syncEngine: string;
}

/** What ships: the companion editor's own defaults. */
export const DEFAULT_SETTINGS: EditorSettings = {
  blockLayout: "comfy",
  continueCommentOnNewline: true,
  mirrorCommentSpacing: true,
  treeFollowsEditor: true,
  formatIndentSize: 4,
  syncEngine: "xlide",
};

let current: EditorSettings = { ...DEFAULT_SETTINGS };

const listeners = new Set<(settings: EditorSettings) => void>();

/** The settings as they stand. A snapshot: mutate nothing, ask again after changes. */
export function currentSettings(): EditorSettings {
  return current;
}

/** Adopts settings the host sent, and tells everyone who asked to hear about it. */
export function applySettings(next: EditorSettings): void {
  current = {
    blockLayout: next.blockLayout === "compact" ? "compact" : "comfy",
    continueCommentOnNewline: next.continueCommentOnNewline !== false,
    mirrorCommentSpacing: next.mirrorCommentSpacing !== false,
    treeFollowsEditor: next.treeFollowsEditor !== false,
    formatIndentSize: Math.min(8, Math.max(1, Math.round(next.formatIndentSize) || 4)),
    syncEngine: next.syncEngine === "builtIn" ? "builtIn" : "xlide",
  };

  for (const listener of listeners) {
    listener(current);
  }
}

/** Hears every applied change, the dialog's way of staying honest about the file's state. */
export function onSettingsApplied(listener: (settings: EditorSettings) => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}
