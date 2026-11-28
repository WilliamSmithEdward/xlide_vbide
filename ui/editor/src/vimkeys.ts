/*
 * Which keys are Vim's while Vim mode is on (#173), decided without an editor so it can be
 * pinned by a test (test/vimkeys.mjs).
 *
 * Plain keys, shifted keys and Escape are Vim's. A Ctrl chord keeps the meaning it has everywhere
 * else in the editor - Ctrl+C copies, Ctrl+V pastes, Ctrl+F finds, Ctrl+Y cuts a line - with four
 * exceptions that have no editor meaning worth keeping: Ctrl+[ is Escape, and outside insert mode
 * Ctrl+Q starts a block selection (Vim's own stand-in where Ctrl+V pastes) and Ctrl+D and Ctrl+U
 * scroll half a page.
 *
 * The keys the host takes before the page sees them (F5, F8, F9, Ctrl+S, Ctrl+R, Ctrl+W and the
 * rest of VbeCommands.ForKey and SurfaceCommandForKey) never arrive here in either mode, so Vim's
 * Ctrl+R is not redo: `:redo` and Ctrl+Shift+Z are.
 */

/** The Ctrl chords Vim may have outside insert mode, as the keymap names them. */
const VIM_CTRL_KEYS: ReadonlySet<string> = new Set(["Ctrl-q", "Ctrl-d", "Ctrl-u"]);

export interface VimKeyContext {
  /** True in insert and replace modes, where a Ctrl chord is the editor's but for Ctrl+[. */
  insertMode: boolean;
  /**
   * The character the key typed when AltGr was down for it, or null. AltGr is how `[`, `{`, `@`
   * and `\` are typed on the layouts that have it, and the browser reports it as Ctrl+Alt.
   */
  altGraphCharacter: string | null;
}

/**
 * What the keymap is to be given for a key it would have been handed as `key`, or null for a key
 * that is not Vim's and goes on to the editor.
 *
 * `key` is the keymap's own name for the press: `'j'` for a character, `Esc`, `Ctrl-d`,
 * `Shift-Ctrl-z`, `Ctrl-Alt-8`.
 *
 * AltGr first, because its name is the wrong one. A `[` typed with AltGr arrives named
 * `Ctrl-Alt-8`, which matches nothing in the keymap; left alone it would fall through to the
 * editor, which types it - in normal mode. The character is what was meant, so the character is
 * what the keymap gets.
 */
export function routeVimKey(key: string, context: VimKeyContext): string | null {
  if (context.altGraphCharacter !== null && context.altGraphCharacter.length === 1) {
    return `'${context.altGraphCharacter}'`;
  }

  if (!/(^|-)Ctrl-/.test(key) || key === "Ctrl-[") {
    return key;
  }

  return !context.insertMode && VIM_CTRL_KEYS.has(key) ? key : null;
}
