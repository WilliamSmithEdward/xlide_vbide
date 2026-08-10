// Type surface of xlide_vscode/src/vbaSourceScan.ts, for the type-check only; the bundle
// carries the real implementation. See ./vbaSmartEnter.d.ts for the why.

/**
 * What a VBA identifier is made of, as a pattern string so a caller can build the flags it needs.
 *
 * ANY LETTER, not an ASCII one: VBA accepts a locale letter, and a combining mark continues a name
 * because Thai and Devanagari build a letter from a base plus a mark. A regex built from this needs
 * the `u` flag, or `\p{L}` is read as a literal 'p{L}' and silently matches nothing.
 */
export const VBA_IDENTIFIER_PATTERN: string;

/** The same, compiled, for finding an identifier anywhere in a line. */
export const VBA_IDENTIFIER_RE: RegExp;

/** The same, anchored, for asking whether a whole string is a legal name. */
export const VBA_IDENTIFIER_NAME_RE: RegExp;
