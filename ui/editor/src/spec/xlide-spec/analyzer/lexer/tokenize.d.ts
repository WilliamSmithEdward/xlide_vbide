// Type surface of xlide_vscode/src/analyzer/lexer/tokenize.ts, for the type-check only; the
// bundle carries the real implementation. See ../../vbaSmartEnter.d.ts for the why. Only what the
// page reads is declared.

/** A token's category. `comment` covers ' and Rem, and runs on through a ` _` at its end. */
export type TokenKind =
  | "newline" | "comment" | "keyword" | "identifier" | "bracketedIdentifier"
  | "integerLiteral" | "floatLiteral" | "dateLiteral" | "stringLiteral"
  | "operator" | "punctuation" | "colon" | "directive" | "unknown";

/** One lexical token. */
export interface VbaToken {
  kind: TokenKind;
  /** The exact source text, which for a carried comment spans physical lines. */
  rawText: string;
  /** The canonical capitalization, on keyword tokens only. */
  canonicalText?: string;
  /** Absolute UTF-16 offsets of the token's first character and just past its last. */
  start: number;
  end: number;
  /** Zero-based line and UTF-16 column of the token's start. */
  line: number;
  character: number;
}

/**
 * The module's tokens, memoized on the text. A keyword that is one only inside its own statement
 * (Step, Error, Explicit, PtrSafe and the rest) comes back as an identifier everywhere else.
 * Callers must not mutate the array or its tokens.
 */
export function tokenizeCached(src: string): VbaToken[];
