/*
 * Where a diagnostic gets underlined when the analysis gave a position but no extent.
 *
 * A ZERO-WIDTH MARKER HAS NO DEFINED RENDERING. Hand an editor a range whose start equals its end
 * and it is entitled to guess, and the guess is not required to land anywhere near the position
 * the marker names. Monaco's does not: measured on a 13-line module whose only finding was
 * `option-explicit-missing` at (1,1), the marker lookup answered correctly at line 1 and nothing
 * at line 13, while the squiggle rendered on line 13 under `End Sub` - thirteen lines from the
 * Problems pane's own "(1, 1)" (the owner, 2026-08-23: "squiggly in wrong place?").
 *
 * So the guess is taken away. A finding that names a position with no extent is about the LINE it
 * names, and underlining that line's text says so.
 *
 * AND THIS IS THE ANSWER, not a stopgap. It was filed upstream as xlide_vscode#49 - emit a span
 * with some width - and closed as won't fix, so `option-explicit-missing` will go on arriving
 * with start equal to end, and any later rule about "the module" rather than about a piece of
 * text may do the same. That makes the widening permanent, and it is the right place for it:
 * only the side holding the model knows what is on the line being named, and every consumer that
 * draws a diagnostic needs this answered by somebody.
 */

/** As much of a text model as the widening needs, so the rule can be tested without one. */
export interface MarkerText {
  lineCount(): number;

  /** The column after the last non-whitespace character on a line, or 0 when it has none. */
  lastNonWhitespaceColumn(line: number): number;
}

/** One diagnostic's range, in the editor's 1-based line and column. */
export interface MarkerRange {
  startLine: number;
  startColumn: number;
  endLine: number;
  endColumn: number;
}

/**
 * The end a marker should be drawn to.
 *
 * A range with any width at all is returned untouched: the analysis meant it, and second-guessing
 * a real span is how a diagnostic ends up pointing somewhere its message does not.
 */
export function widenIfEmpty(text: MarkerText, marker: MarkerRange): { endLine: number; endColumn: number } {
  if (marker.startLine !== marker.endLine || marker.startColumn !== marker.endColumn) {
    return { endLine: marker.endLine, endColumn: marker.endColumn };
  }

  const line = Math.min(Math.max(marker.startLine, 1), Math.max(text.lineCount(), 1));

  // The line's TEXT, not its full width: trailing whitespace underlined reads as a range that
  // means something. A line with nothing on it has nothing to underline, so it takes one column -
  // which is the one case where a single character is the honest answer.
  const lastText = text.lastNonWhitespaceColumn(line);
  return {
    endLine: line,
    endColumn: lastText > marker.startColumn ? lastText : marker.startColumn + 1,
  };
}
