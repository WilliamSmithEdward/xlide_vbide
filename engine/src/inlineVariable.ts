/*
 * Inline Variable: a local replaced by what it was assigned, and its declaration taken away.
 *
 * The inverse of Extract Variable, and the one of these with a trap VBA sets that no other
 * language does.
 *
 * PARENTHESES ARE NOT FREE HERE. Everywhere else, wrapping an inlined expression in brackets keeps
 * precedence and costs nothing. In VBA, `Foo x` passes `x` by reference and `Foo (x)` passes it by
 * value - the brackets are not grouping, they are an evaluation - so an inliner that parenthesises
 * to be safe about precedence silently changes how every call site binds. There is no way to have
 * both without deciding, per use, whether that use is an argument position, and that is a question
 * about the caller's signature which this cannot always answer.
 *
 * So only an ATOMIC initialiser is inlined: a literal, or one name with its member chain. Those
 * need no brackets at any precedence and mean the same thing in an argument position as out of
 * one. It covers what people actually reach for - `Set ws = Sheet1` used five times, `n = 0` used
 * twice - and everything else is refused in words rather than transformed on a guess.
 *
 * The second trap is evaluation count. Inlining a variable read three times evaluates its
 * initialiser three times, which changes behaviour the moment the expression has a side effect,
 * and VBA says nothing about which expressions are pure. An atomic initialiser is either a literal
 * or a name, and reading either has no side effect - so the atomic rule settles this one too.
 */

import { parseModule, type BodyNode, type ProcedureNode, type VariableGroupNode } from '../../../xlide_vscode/src/analyzer';
import { lineStarts, referencesFor, toLineColumn, type ProjectSymbols } from './navigation';
import type { InlineVariableResult, LocationPayload } from './protocol';

/**
 * An initialiser that can stand anywhere the name stood: a literal, or a single name with an
 * optional member chain and no call.
 *
 * The `(` is what excludes a call and an index alike, deliberately: `Cells(1, 1)` is atomic to a
 * reader and is a call to VBA, so inlining it into three places calls it three times. An operator
 * anywhere is excluded because it would need brackets, which is the trap above.
 */
const ATOMIC = /^(?:"(?:[^"]|"")*"|#[^#]*#|-?\d[\d.]*(?:[eE][+-]?\d+)?[&%@!#]?|(?:True|False|Nothing|Empty|Null)|[\p{L}_][\p{L}\p{M}\p{N}_]*(?:\.[\p{L}_][\p{L}\p{M}\p{N}_]*)*)$/u;

export function inlineVariableFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
): InlineVariableResult {
    const refuse = (why: string): InlineVariableResult => ({ refused: why });

    const module = parseModule(source);
    const procedure = module.members.find(
        (member): member is ProcedureNode => member.kind === 'Procedure'
            && member.span.start <= offset && member.span.end >= offset);

    if (!procedure) {
        return refuse('Put the caret on a local variable inside a procedure.');
    }

    const named = identifierAt(source, offset);
    if (!named) {
        return refuse('There is no name here to inline.');
    }

    if (procedure.params.some((one) => one.name.toLowerCase() === named.toLowerCase())) {
        return refuse(`'${named}' is a parameter, whose value comes from the caller. Inlining it would mean deciding what every caller passes.`);
    }

    let group: VariableGroupNode | undefined;
    let declared: { name: string; line: number; column: number } | undefined;

    for (const node of flatten(procedure.body)) {
        if (node.kind !== 'VariableGroup') {
            continue;
        }

        const found = node.declarations.find((one) => one.name.toLowerCase() === named.toLowerCase());
        if (found?.nameSpan) {
            group = node;
            const at = toLineColumn(lineStarts(source), found.nameSpan.start);
            declared = { name: found.name, line: at.line, column: at.column };
            break;
        }
    }

    if (!group || !declared) {
        return refuse(`'${named}' is not a local declared in '${procedure.name}'. Only a local can be inlined.`);
    }

    if (group.isConst) {
        return refuse(`'${declared.name}' is a Const, which is already a name for a value. Removing it is an edit, not a refactoring.`);
    }

    if (group.declarations.length > 1) {
        return refuse(`'${declared.name}' shares its declaration with other names. Give it a line of its own first.`);
    }

    const starts = lineStarts(source);
    const uses = referencesFor(symbols, moduleName, source, offset, true)
        .filter((one) => one.module.toLowerCase() === moduleName.toLowerCase())
        .filter((one) => !(one.line === declared.line && one.column === declared.column))
        .sort((a, b) => a.line - b.line || a.column - b.column);

    const writes = uses.filter((one) => one.kind === 'write' || one.kind === 'readwrite');
    const reads = uses.filter((one) => one.kind === 'read' || one.kind === 'readwrite');

    if (writes.length === 0) {
        return refuse(`'${declared.name}' is never assigned, so there is nothing to inline in its place.`);
    }

    if (writes.length > 1) {
        return refuse(`'${declared.name}' is assigned ${writes.length} times, so it does not stand for one value. Inlining it would have to choose which.`);
    }

    if (reads.length === 0) {
        return refuse(`'${declared.name}' is assigned and never read, so inlining it would leave nothing behind. Remove it instead.`);
    }

    const assignment = writes[0] as LocationPayload;
    if (reads.some((one) => one.line < assignment.line
        || (one.line === assignment.line && one.column < assignment.column))) {
        return refuse(`'${declared.name}' is read before it is assigned, so its value is not the same everywhere it appears.`);
    }

    /* ---- the initialiser ---------------------------------------------------------------------- */

    const assignedLine = assignment.line;
    const lineStart = starts[assignedLine - 1] as number;
    const lineEnd = assignedLine < starts.length ? (starts[assignedLine] as number) : source.length;
    const text = source.slice(lineStart, lineEnd);

    const equals = text.indexOf('=');
    if (equals < 0) {
        return refuse(`'${declared.name}' is not assigned with '=' on its own line, which is the only shape this can read.`);
    }

    // `Set x = ...` and `x = ...` both, and nothing else on the line: a colon-joined statement or
    // a continuation would put more than one thing in what is read as the value.
    const before = text.slice(0, equals);
    if (!new RegExp(`^\\s*(?:Set\\s+)?${escaped(declared.name)}\\s*$`, 'iu').test(before)) {
        return refuse(`The line assigning '${declared.name}' does more than assign it, so what it stands for cannot be read from it.`);
    }

    const value = text.slice(equals + 1).replace(/\r|\n/g, '').trim();
    if (value.endsWith('_')) {
        return refuse(`'${declared.name}' is assigned over more than one line. Put the value on one line first.`);
    }

    if (value.length === 0) {
        return refuse(`'${declared.name}' is assigned nothing this can read.`);
    }

    if (!ATOMIC.test(value)) {
        return refuse(`'${declared.name}' stands for '${value}', which is not a literal or a plain name. Inlining it would need brackets, and in VBA brackets around an argument pass it by value instead of by reference - so this one has to be done by hand.`);
    }

    /* ---- the edits, back to front --------------------------------------------------------------- */

    const edits: { start: number; end: number; text: string }[] = [];

    for (const read of reads) {
        const at = (starts[read.line - 1] as number) + read.column - 1;
        edits.push({ start: at, end: at + read.length, text: value });
    }

    // The assignment line and the declaration line go whole. Both are statements of their own by
    // the checks above, so nothing else is on them.
    edits.push({ start: lineStart, end: lineEnd, text: '' });

    const declaredFrom = toLineColumn(starts, group.span.start).line;
    const declaredTo = toLineColumn(starts, Math.max(group.span.start, group.span.end - 1)).line;
    if (declaredFrom !== declaredTo) {
        return refuse(`'${declared.name}' is declared over more than one line, which this cannot rewrite safely.`);
    }

    edits.push({
        start: starts[declaredFrom - 1] as number,
        end: declaredFrom < starts.length ? (starts[declaredFrom] as number) : source.length,
        text: '',
    });

    let rewritten = source;
    for (const edit of edits.sort((a, b) => b.start - a.start)) {
        rewritten = rewritten.slice(0, edit.start) + edit.text + rewritten.slice(edit.end);
    }

    return {
        module: moduleName,
        source: rewritten,
        variable: declared.name,
        value,
        replaced: reads.length,
    };
}

/** The identifier covering an offset, as the surface's caret lands in the middle of one. */
function identifierAt(source: string, offset: number): string | null {
    const isPart = (at: number): boolean =>
        at >= 0 && at < source.length && /[\p{L}\p{M}\p{N}_]/u.test(source[at] as string);

    if (!isPart(offset) && !isPart(offset - 1)) {
        return null;
    }

    let start = offset;
    while (isPart(start - 1)) {
        start -= 1;
    }

    let end = offset;
    while (isPart(end)) {
        end += 1;
    }

    const word = source.slice(start, end);
    return /^[\p{L}_]/u.test(word) ? word : null;
}

/** A name as a regular expression, so a name with no special characters still matches literally. */
function escaped(name: string): string {
    return name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Every node in a procedure body, blocks and their contents alike. */
function flatten(body: readonly BodyNode[]): BodyNode[] {
    const out: BodyNode[] = [];
    const visit = (nodes: readonly BodyNode[]): void => {
        for (const node of nodes) {
            out.push(node);
            const nested = (node as { body?: readonly BodyNode[] }).body;
            if (nested) {
                visit(nested);
            }
        }
    };
    visit(body);
    return out;
}
