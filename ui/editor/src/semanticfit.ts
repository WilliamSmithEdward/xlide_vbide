/*
 * Whether a set of semantic tokens can possibly describe the text it is about to be painted on.
 *
 * THE PROBLEM IT EXISTS FOR. A semantic token is a pair of absolute character offsets into the
 * text the ANALYSER was given, and the surface maps them against the model as it stands. Those
 * are the same text almost always and not always: the page holds an edit until typing stops
 * before writing it to the module, and an agent writing through the api moves the module while
 * the analysis of the previous version is still in flight. When the two differ, every offset
 * lands somewhere that means nothing.
 *
 * It does not look like an error. It looks like the colours have come apart - which is what the
 * owner called it (2026-08-22, "colors are glitched"), catching it while a module was being
 * written. Measured on the spot: a comment drawn as five spans in two colours, breaking at
 * "Idemp|otent" and "the ol|d on|e", the fragments painted in the colour of a type.
 *
 * THE RULE IS AN INVARIANT, NOT A GUESS. A semantic token names an identifier, so its range
 * covers a whole one. A range that begins or ends in the MIDDLE of a word is therefore not a
 * range into this text, whatever it was measured against. Checked against the live surface
 * before being relied on: 148 correctly painted spans, not one of them splitting a word.
 *
 * (The four that appeared to were Monaco's own ~50-character chunking of a long token for
 * rendering. Those share a class either side and are not token boundaries at all - which is
 * worth knowing before anyone tries to write this check by counting spans in the DOM.)
 */

/** What counts as part of a word. VBA identifiers, plus the digits and underscores in them. */
const WORD = /[A-Za-z0-9_]/;

/** The little of a text model this needs, so the rule can be tested without one. */
export interface FitText {
  /** The line a character offset falls on, 1-based, with its 1-based column. */
  positionAt(offset: number): { lineNumber: number; column: number };

  /** One line's text, without its ending. */
  lineAt(lineNumber: number): string;
}

/** One token's range, in absolute character offsets into the analysed text. */
export interface FitToken {
  start: number;
  end: number;
}

/**
 * True when every sampled token sits on identifier boundaries, so the set can be painted.
 *
 * SAMPLED, because a corrupted set is corrupted throughout - the offsets all come from one
 * analysis - so reading twenty-odd of them answers the same question as reading sixty thousand.
 * `every` here would turn a colouring pass on a large module into a second walk of it.
 */
export function tokensFitTheText(text: FitText, tokens: readonly FitToken[], samples = 24): boolean {
  if (tokens.length === 0) {
    return true;
  }

  const step = Math.max(1, Math.floor(tokens.length / samples));

  for (let at = 0; at < tokens.length; at += step) {
    const token = tokens[at];
    if (!token || token.end <= token.start) {
      continue;
    }

    const start = text.positionAt(token.start);
    const end = text.positionAt(token.end);
    if (end.lineNumber !== start.lineNumber) {
      // Single-line by construction; a range spanning rows is dropped by the caller anyway.
      continue;
    }

    const line = text.lineAt(start.lineNumber);

    // A WORD CHARACTER ON BOTH SIDES OF THE EDGE is the split. A token beside a dot, a bracket,
    // a space or the end of the line is exactly what a well-placed token looks like.
    const insideStart = line[start.column - 1];
    const beforeStart = line[start.column - 2];
    if (insideStart !== undefined && beforeStart !== undefined
      && WORD.test(insideStart) && WORD.test(beforeStart)) {
      return false;
    }

    const insideEnd = line[end.column - 2];
    const afterEnd = line[end.column - 1];
    if (insideEnd !== undefined && afterEnd !== undefined
      && WORD.test(insideEnd) && WORD.test(afterEnd)) {
      return false;
    }
  }

  return true;
}
