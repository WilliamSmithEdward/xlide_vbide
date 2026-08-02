// Type surface of xlide_vscode/src/analyzer/lexer/strippedLines.ts, for the type-check only;
// the bundle carries the real implementation. See ../../vbaSmartEnter.d.ts for the why.

/** One line with strings and comments blanked, columns preserved. */
export function lexerStrippedLine(line: string): string;

/** Every line of a module with strings and comments blanked, columns preserved. */
export function lexerStrippedLines(source: string): string[];
