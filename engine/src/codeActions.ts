// Serves quick fixes by reusing the editor extension's code-action resolver.
//
// Unlike hover and completion, a fix is not resolved from the caret: it is resolved from a finding
// the analyzer already reported, and most fixes need more of that finding than the surface ever
// sees - which argument is missing, where the unclosed block wants its closer. So the fix data
// stays here, on the analysis the engine holds, rather than being carried down to the surface and
// asked for back. The request names a span; every finding overlapping it offers what it can fix.

import { normalizeDiagnosticCode, resolveDiagnosticCodeActions } from '../../../xlide_vscode/src/analyzer';
// The libraries a project can be given a reference to, with the GUID each one is identified by.
// Upstream's table, imported rather than copied, so a library added there needs no change here -
// and measured before it was taken: it costs this bundle 0.9 KB and four input files, because
// esbuild shakes out the container-format code that module also holds.
import { HOST_LIBRARIES } from '../../../xlide_vscode/src/vba/vbaProjectReferences';
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

        // AND THE ONE FIX THAT IS NOT A TEXT EDIT.
        //
        // `missing-library-reference` says a module names a library the project does not
        // reference. Suppressing it is the only thing the analyzer's own resolver can offer,
        // because the fix writes a record into the PROJECT and the resolver deals in edits to a
        // module - so the editor extension adds this one at its own layer too, as a command.
        //
        // The library is read from the finding's data rather than out of its message: the rule
        // attaches the token it meant, which is the difference between reading a fact and
        // parsing English.
        const missing = diagnostic.data?.addLibraryReference;
        const library = missing ? HOST_LIBRARIES[missing.library.toLowerCase()] : undefined;
        if (missing && library) {
            const action: CodeActionPayload = {
                title: `Add a reference to the ${library.name} object library`,
                // Ahead of the suppression, which is the other thing offered here. Hiding a
                // compile error is a last resort and adding the reference is the fix.
                isPreferred: true,
                code,
                span: { start: diagnostic.span.start, end: diagnostic.span.end },
                edits: [],
                command: 'addLibraryReference',
                // The GUID is the library's identity and the host binds the version it has
                // installed, so no version travels: `AddFromGuid(guid, 0, 0)` bound Word 8.7
                // from MSWORD.OLB on a machine with Office 16 (measured 2026-09-21). A pinned
                // major.minor would be this machine's Office version written into the product.
                arguments: [missing.library.toLowerCase(), library.name, library.guid],
            };

            const identity = JSON.stringify([action.title, action.command, action.arguments]);
            if (!seen.has(identity)) {
                seen.add(identity);
                actions.push(action);
            }
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
