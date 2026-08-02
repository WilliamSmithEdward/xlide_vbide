// Type surface of xlide_vscode/src/vbaSmartTab.ts, for the type-check only; the bundle
// carries the real implementation. See ./vbaSmartEnter.d.ts for the why.

/**
 * True when Tab at this caret should indent the line (blank line, leading whitespace, or a
 * multi-line selection) rather than insert a tab at the caret. Columns count from zero.
 */
export function smartTabShouldIndentLine(
  lineText: string,
  caretColumn: number,
  selectionIsEmpty: boolean,
  selectionSpansLines: boolean,
): boolean;
