// Type surface of xlide_vscode/src/vbaSmartEnter.ts, for this project's type-check only.
//
// The page bundles the spec repo's smart-editing helpers directly - build.mjs aliases
// "xlide-spec/*" to ../../xlide_vscode/src/* - so the BEHAVIOUR always comes from the spec.
// The spec compiles under its own, laxer compiler settings, and pulling its sources into this
// project's stricter type-check fails on idioms its own settings allow. tsc therefore resolves
// the alias here instead: signatures only, kept in step by hand with the spec's exports, which
// are its stable public smart-editing API. A drifted signature fails HERE at compile time; it
// can never change what runs.

/** A block opener Enter should complete: its closing statement and any body-line prefix. */
export interface VbaSmartBlockOpener {
  endKeyword: string;
  bodyPrefix?: string;
}

/** The body-line replacement Smart Enter builds, and where the editable line sits in it. */
export interface VbaSmartBlockInsertion {
  bodyText: string;
  bodyLineOffset: number;
  replacementText: string;
}

export type VbaSmartBlockLayout = "comfy" | "compact";

/** The parens a bare procedure header is owed, as columns into its line. */
export interface VbaProcedureHeaderParensEdit {
  startCol: number;
  endCol: number;
  newText: string;
}

/** The paired For/Next iterator rename an edit calls for, as an absolute span replacement. */
export interface VbaLoopIteratorSyncEdit {
  span: { start: number; end: number };
  newText: string;
}

export function detectSmartBlockOpener(strippedLine: string): VbaSmartBlockOpener | undefined;

export function isSmartBlockClosedAhead(
  strippedLines: string[],
  openerIdx: number,
  opener: VbaSmartBlockOpener,
): boolean;

export function procedureHeaderParensEdit(line: string): VbaProcedureHeaderParensEdit | undefined;

export function smartBlockInsertion(
  openerLine: string,
  currentBodyLine: string,
  opener: VbaSmartBlockOpener,
  options?: {
    eol?: string;
    insertCloser?: boolean;
    indentUnit?: string;
    layout?: VbaSmartBlockLayout;
  },
): VbaSmartBlockInsertion;

export function commentContinuationText(
  source: string,
  previousLineIndex: number,
  mirrorSpacing: boolean,
): string | undefined;

export function withMemberContinuationText(
  source: string,
  previousLineIndex: number,
): string | undefined;

export function resolveLoopIteratorSyncEdit(
  source: string,
  offset: number,
): VbaLoopIteratorSyncEdit | undefined;
