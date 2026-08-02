// Serves the typing-time edits: Smart Enter block completion, canonical casing, and loop
// iterator sync, by reusing the editor extension's own smart-editing helpers.
//
// These are the behaviours a developer feels between keystrokes, and the extension is the
// specification for them, down to the small details: Enter after `Sub Test` leaves `Sub Test()`
// with `End Sub` below, `msgbox` becomes `MsgBox` when the line goes idle, renaming a `For`
// iterator renames its `Next`. Reusing the helpers keeps one behaviour shared between the two
// products instead of two that drift.
//
// The surface applies what comes back as ordinary editor edits, so everything here flows through
// the same write-back path a keystroke does.

import {
    resolveCanonicalCaseEdit,
    resolveCanonicalCaseEdits,
    type CanonicalCaseContext,
} from '../../../xlide_vscode/src/analyzer';
import {
    lexerStrippedLine,
    lexerStrippedLines,
} from '../../../xlide_vscode/src/analyzer/lexer/strippedLines';
import {
    commentContinuationText,
    detectSmartBlockOpener,
    isSmartBlockClosedAhead,
    procedureHeaderParensEdit,
    resolveLoopIteratorSyncEdit,
    smartBlockInsertion,
    withMemberContinuationText,
} from '../../../xlide_vscode/src/vbaSmartEnter';
import { assembleContext } from './moduleContext';
import type {
    CanonicalCaseParams,
    LoopSyncParams,
    ModulePayload,
    SmartEnterParams,
    SmartEnterResult,
    TextEditPayload,
} from './protocol';

// The extension reads these from settings; this surface has no settings page yet, so the
// defaults it ships with are the behaviour here (task #12 owns making them configurable).
const BLOCK_LAYOUT = 'comfy' as const;
const CONTINUE_COMMENT_ON_NEWLINE = true;
const MIRROR_COMMENT_SPACING = true;

interface PhysicalLine {
    text: string;
    start: number;
    end: number;
}

/**
 * What Enter should leave behind, given the text as it stands just after the newline went in.
 * The offset names where the newline was inserted, which is the end of the line Enter was
 * pressed on; the caret in the result is an offset into the text as it stands after the edits.
 */
export function smartEnterFor(params: SmartEnterParams): SmartEnterResult {
    const source = params.source;
    const lines = physicalLines(source);
    const openerIndex = lineIndexAt(lines, params.offset);
    const opener = lines[openerIndex];
    if (!opener) {
        return { edits: [], caret: null };
    }

    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const headerEdit = procedureHeaderParensEdit(opener.text);
    const normalized = headerEdit
        ? `${opener.text.slice(0, headerEdit.startCol)}${headerEdit.newText}${opener.text.slice(headerEdit.endCol)}`
        : opener.text;

    const blockOpener = detectSmartBlockOpener(lexerStrippedLine(normalized));
    const body = lines[openerIndex + 1];
    if (!body || !/^[ \t]*$/.test(body.text)) {
        // Enter mid-line pushed text down, or Enter on the last line: nothing is owed. The
        // header parens are deliberately not applied alone; the extension leaves the line too.
        return { edits: [], caret: null };
    }

    if (!blockOpener) {
        // Not a block: a whole-line comment continues its apostrophes, and a `.member` line
        // inside an open With seeds the next dot. Checked in the extension's order.
        const continuation = CONTINUE_COMMENT_ON_NEWLINE
            ? commentContinuationText(source, openerIndex, MIRROR_COMMENT_SPACING)
            : undefined;
        const lineText = continuation ?? withMemberContinuationText(source, openerIndex);
        if (lineText === undefined || lineText.length === 0) {
            return { edits: [], caret: null };
        }

        return {
            edits: [{ start: body.start, end: body.end, text: lineText }],
            caret: body.start + lineText.length,
        };
    }

    const strippedLines = lexerStrippedLines(source);
    strippedLines[openerIndex] = lexerStrippedLine(normalized);
    const closedAhead = isSmartBlockClosedAhead(strippedLines, openerIndex, blockOpener);

    const insertion = smartBlockInsertion(normalized, body.text, blockOpener, {
        eol,
        insertCloser: !closedAhead,
        layout: BLOCK_LAYOUT,
    });

    const edits: TextEditPayload[] = [];
    let headerShift = 0;
    if (headerEdit) {
        edits.push({
            start: opener.start + headerEdit.startCol,
            end: opener.start + headerEdit.endCol,
            text: headerEdit.newText,
        });
        headerShift = headerEdit.newText.length - (headerEdit.endCol - headerEdit.startCol);
    }

    edits.push({ start: body.start, end: body.end, text: insertion.replacementText });

    // The caret lands at the end of the editable body line: past the spacer line the comfy
    // layout opens with, at the start when the closer was already there.
    const bodyPrefix = insertion.bodyLineOffset > 0 ? eol.length * insertion.bodyLineOffset : 0;
    return {
        edits,
        caret: body.start + headerShift + bodyPrefix + insertion.bodyText.length,
    };
}

/**
 * The case corrections for a span of a module: keywords to their canonical spelling,
 * identifiers to the casing of their declarations, with the same project facts completion uses.
 * The single form corrects only the identifier ending at the span's end, which is what a caret
 * leaves behind when it moves within a line.
 */
export function canonicalCaseFor(
    seeded: readonly ModulePayload[],
    params: CanonicalCaseParams,
): TextEditPayload[] {
    const ctx = assembleContext(seeded, params);
    const caseContext: CanonicalCaseContext = {
        member: {
            codeNames: ctx.codeNames,
            meType: ctx.meType,
            meProjectType: ctx.meProjectType,
            projectClassMembers: ctx.projectClassMembers,
        },
        identifier: {
            codeNames: ctx.codeNameList,
            moduleName: ctx.current.name,
            moduleKind: ctx.moduleKind,
            projectMemberSurfaces: ctx.projectClassMembers,
            projectProcedures: ctx.projectProcedures,
            projectSymbols: ctx.projectSymbols,
        },
        type: {
            projectTypes: ctx.projectTypes,
        },
    };

    const edits: TextEditPayload[] = [];
    if (params.single) {
        const edit = resolveCanonicalCaseEdit(params.source, params.end, caseContext);
        if (edit) {
            edits.push(edit);
        }
    } else {
        edits.push(...resolveCanonicalCaseEdits(
            params.source,
            { start: params.start, end: params.end },
            caseContext,
        ));
    }

    // Leaving a bare `Sub Test` line also completes its parentheses, the way the editor's own
    // rewrite would. Only meaningful when the span is the whole line, which is how it is asked.
    if (params.completeHeader) {
        const line = params.source.slice(params.start, params.end);
        const headerEdit = procedureHeaderParensEdit(line);
        if (headerEdit) {
            edits.push({
                start: params.start + headerEdit.startCol,
                end: params.start + headerEdit.endCol,
                text: headerEdit.newText,
            });
        }
    }

    return edits;
}

/**
 * The paired rename when an edit touches a simple `For`/`For Each` iterator or its `Next` name:
 * the other side of the pair, respelled to match. At most one edit.
 */
export function loopSyncFor(params: LoopSyncParams): TextEditPayload[] {
    const edit = resolveLoopIteratorSyncEdit(params.source, params.offset);
    return edit ? [{ start: edit.span.start, end: edit.span.end, text: edit.newText }] : [];
}

function physicalLines(source: string): PhysicalLine[] {
    const lines: PhysicalLine[] = [];
    let start = 0;
    for (let i = 0; i < source.length; i++) {
        if (source[i] !== '\n') {
            continue;
        }
        const end = i > start && source[i - 1] === '\r' ? i - 1 : i;
        lines.push({ text: source.slice(start, end), start, end });
        start = i + 1;
    }
    lines.push({ text: source.slice(start), start, end: source.length });
    return lines;
}

function lineIndexAt(lines: PhysicalLine[], offset: number): number {
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line && offset >= line.start && offset <= line.end) {
            return i;
        }
    }
    return lines.length - 1;
}
