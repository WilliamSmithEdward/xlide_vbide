// Serves quick fixes by reusing the editor extension's code-action resolver.
//
// Unlike hover and completion, a fix is not resolved from the caret: it is resolved from a finding
// the analyzer already reported, and most fixes need more of that finding than the surface ever
// sees - which argument is missing, where the unclosed block wants its closer. So the fix data
// stays here, on the analysis the engine holds, rather than being carried down to the surface and
// asked for back. The request names a span; every finding overlapping it offers what it can fix.

import { normalizeDiagnosticCode, resolveDiagnosticCodeActions } from '../../../xlide_vscode/src/analyzer';
import type { VbaModuleAnalysisDiagnostic } from '../../../xlide_vscode/src/vbaModuleAnalysis';
import type { CodeActionPayload } from './protocol';

export function codeActionsFor(
    source: string,
    diagnostics: readonly VbaModuleAnalysisDiagnostic[],
    span: { start: number; end: number },
): CodeActionPayload[] {
    const actions: CodeActionPayload[] = [];
    const seen = new Set<string>();

    for (const diagnostic of diagnostics) {
        if (!overlaps(diagnostic.span, span)) {
            continue;
        }

        const code = normalizeDiagnosticCode(diagnostic.code);
        if (!code) {
            continue;
        }

        const fixes = resolveDiagnosticCodeActions(source, {
            code,
            message: diagnostic.message,
            span: diagnostic.span,
            expectedClose: diagnostic.expectedClose,
            insertLine: diagnostic.insertLine,
            expectedCloseReplacementSpan: diagnostic.expectedCloseReplacementSpan,
            expectedCloseReplacementText: diagnostic.expectedCloseReplacementText,
            includeSuppressionAction: true,
            data: diagnostic.data,
        });

        for (const fix of fixes) {
            const action: CodeActionPayload = {
                title: fix.title,
                isPreferred: fix.isPreferred,
                code,
                span: { start: diagnostic.span.start, end: diagnostic.span.end },
                edits: fix.edits.map((edit) => ({
                    start: edit.span.start,
                    end: edit.span.end,
                    text: edit.newText,
                })),
            };

            // Two findings on the same line can each offer the same whole-module fix - adding
            // Option Explicit, most obviously - and the same entry twice in the menu reads as a
            // bug. Identity is the title and the edits, since that is the whole of what applying
            // it would do.
            const identity = JSON.stringify([action.title, action.edits]);
            if (seen.has(identity)) {
                continue;
            }

            seen.add(identity);
            actions.push(action);
        }
    }

    return actions;
}

/**
 * Whether a finding is close enough to the asked-about span to offer its fixes. Touching counts:
 * the caret sitting at either end of a squiggle is the developer pointing at it, and an empty
 * range - which is what a caret is - would otherwise match nothing at all.
 */
function overlaps(
    finding: { start: number; end: number },
    asked: { start: number; end: number },
): boolean {
    return finding.start <= asked.end && asked.start <= finding.end;
}
