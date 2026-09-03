/*
 * Extract Method: the first refactoring beyond rename, and the head of a set.
 *
 * Select whole statements, name them, and they become a Private procedure below the one they came
 * from, with the selection replaced by the call. Four decisions make the signature, and each one
 * is read off the analyzer's reference kinds (xlide_vscode#55) rather than guessed here:
 *
 *   read before it is written inside      a parameter, ByVal when the callee never assigns it
 *   written inside and read afterwards    the result, or a ByRef parameter when it cannot be
 *   written inside and never read after   its Dim moves into the new procedure
 *   not a local of this procedure         left alone, a free reference
 *
 * That last one is the whole difference between a refactoring a developer keeps and a six-argument
 * procedure nobody wants: `Sheet1`, module-level state and other procedures are not locals, so they
 * are never passed.
 *
 * THE REFUSALS MATTER MORE THAN THE TRANSFORMATION. An extraction that quietly changes what the
 * code does is worse than one that declines, so every case where the meaning would move - control
 * flow crossing the boundary, a lifetime that would change, a receiver left behind - is refused in
 * words a developer can act on. See docs/extract-method.md for the design and what a spike found
 * the analyzer already answers.
 */

import { parseModule, type BodyNode, type ProcedureNode, type VariableGroupNode } from '../../../xlide_vscode/src/analyzer';
import { lineStarts, referencesFor, toLineColumn, type ProjectSymbols } from './navigation';
import type { ExtractMethodResult, LocationPayload } from './protocol';

/**
 * A VBA identifier, and the same rule the rename box applies: a letter, then letters, digits and
 * underscores. Shared wording rather than a shared constant, because rename's list also refuses a
 * leading underscore for a reason that is its own.
 */
const IDENTIFIER = /^\p{L}[\p{L}\p{M}\p{N}_]{0,254}$/u;

/**
 * The types a Function may return here.
 *
 * VBA assigns an object with `Set` and a value without it, and getting that wrong turns a working
 * extraction into a compile error at the call site. The declared type is the only evidence in
 * reach, so a result whose type is not plainly one of these becomes a ByRef parameter instead -
 * always correct, whatever the variable holds. `Variant` is deliberately absent: it can hold an
 * object, so it needs `Set` exactly when nobody can tell in advance.
 */
const VALUE_TYPES = new Set([
    'boolean', 'byte', 'currency', 'date', 'double', 'integer', 'long', 'longlong', 'longptr',
    'single', 'string',
]);

/** A local or parameter of the procedure being extracted from, with everything the signature needs. */
interface Local {
    name: string;
    /** The `As` clause's type, or the suffix's, or Variant when it has neither. */
    type: string;
    isArray: boolean;
    /** A parameter is already in the caller's signature; a Dim can move. */
    isParameter: boolean;
    /** Static locals outlive the call, which is the one thing extraction must not change. */
    isStatic: boolean;
    isConst: boolean;
    /** The declaring group, so a moved Dim can be cut from the caller. */
    group?: VariableGroupNode;
    /** Where the name is declared, excluded from the reference walk so a Dim is not a use. */
    declaredAt?: { line: number; column: number };
    /** The offset the reference lookup is anchored at. */
    anchor: number;
}

/** What the selection does with one local: enough to decide its part in the signature. */
interface Usage {
    readsBeforeWriting: boolean;
    writesInside: boolean;
    readsInside: boolean;
    readsAfter: boolean;
    usedBefore: boolean;
    usedAfter: boolean;
    usedInside: boolean;
}

export function extractMethodFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    startLine: number,
    endLine: number,
    newName: string,
): ExtractMethodResult {
    const refuse = (why: string): ExtractMethodResult => ({ refused: why });

    if (!IDENTIFIER.test(newName)) {
        return refuse(`'${newName}' is not a VBA name. A name starts with a letter and holds letters, digits and underscores.`);
    }

    const starts = lineStarts(source);
    const lineCount = starts.length;
    if (startLine < 1 || endLine < startLine || endLine > lineCount) {
        return refuse('That is not a range of lines in this module.');
    }

    // Option Explicit off makes an undeclared name an implicit Variant at procedure scope, so
    // extraction would move its scope and with it its lifetime. That is a behaviour change hiding
    // inside a refactoring, and the developer would have no way of seeing it.
    if (!/^[ \t]*Option[ \t]+Explicit\b/im.test(source)) {
        return refuse('This module does not declare Option Explicit, so an undeclared name is an implicit variable whose lifetime extraction would change. Add Option Explicit first.');
    }

    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const selectionStart = starts[startLine - 1] as number;
    const selectionEnd = endLine < lineCount ? (starts[endLine] as number) : source.length;
    const selected = source.slice(selectionStart, selectionEnd);

    const module = parseModule(source);
    const procedure = module.members.find(
        (member): member is ProcedureNode => member.kind === 'Procedure'
            && member.span.start <= selectionStart
            && member.span.end >= selectionEnd,
    );

    if (!procedure) {
        return refuse('Select statements inside one procedure. A selection that reaches outside a procedure, or across two of them, has no single caller to extract from.');
    }

    if (!procedure.closed) {
        return refuse(`'${procedure.name}' has no closing ${closerOf(procedure)}, so there is nothing whole to extract from.`);
    }

    // Inside the BODY: the header declares the signature and the closer ends it, and taking
    // either produces a procedure that cannot be parsed.
    const headerLine = toLineColumn(starts, procedure.span.start).line;
    const closerLine = toLineColumn(starts, Math.max(procedure.span.start, procedure.span.end - 1)).line;
    if (startLine <= headerLine || endLine >= closerLine) {
        return refuse(`Select statements inside '${procedure.name}', not its ${procedure.procKind === 'Sub' ? 'Sub' : 'header'} line or its ${closerOf(procedure)}.`);
    }

    if (!/[^\s]/.test(stripComments(selected))) {
        return refuse('There are no statements in that selection.');
    }

    const nodes = flatten(procedure.body);
    const refusal = crossings(nodes, source, starts, selected, startLine, endLine);
    if (refusal) {
        return refuse(refusal);
    }

    const locals = localsOf(procedure, starts);
    if (locals.some((local) => local.isStatic && usedIn(local, symbols, moduleName, source, startLine, endLine))) {
        return refuse('The selection uses a Static local, whose value outlives the call. Extracting it would change that, so this one has to be done by hand.');
    }

    // A name the module already means something by would compile to the wrong call, or not
    // compile at all. Checked against the module's own members rather than the whole project,
    // because the new procedure is Private.
    const taken = module.members.some((member) => member.kind === 'Procedure'
        && member.name.toLowerCase() === newName.toLowerCase());
    if (taken) {
        return refuse(`'${newName}' is already a procedure in this module.`);
    }

    /* ---- the four decisions ---------------------------------------------------------------- */

    const inputs: Local[] = [];
    const byRef: Local[] = [];
    const moved: Local[] = [];
    let result: Local | undefined;

    for (const local of locals) {
        const usage = usageOf(local, symbols, moduleName, source, startLine, endLine);
        if (!usage.usedInside) {
            continue;
        }

        const needsIncomingValue = usage.readsBeforeWriting;
        const carriesValueBack = usage.writesInside && usage.readsAfter;

        if (carriesValueBack) {
            // One value comes back as the result; VBA has only the one, so the rest go ByRef.
            // A variable the callee must also be TOLD cannot be the result, because a function's
            // return value starts empty - it has to be a ByRef parameter whatever its type.
            if (!result && !needsIncomingValue && !local.isArray && VALUE_TYPES.has(local.type.toLowerCase())) {
                result = local;
            } else {
                byRef.push(local);
            }
            continue;
        }

        if (needsIncomingValue) {
            inputs.push(local);
            continue;
        }

        if (usage.writesInside && !usage.usedAfter && !usage.usedBefore && !local.isParameter && !local.isConst) {
            // Written here, read nowhere else: the variable belongs to the extracted code, so its
            // declaration goes with it.
            moved.push(local);
            continue;
        }

        // Written inside and used before or after without being read after: still the caller's
        // variable, so the callee needs to write through to it.
        if (usage.writesInside) {
            byRef.push(local);
        }
    }

    /* ---- the new procedure ------------------------------------------------------------------ */

    const parameters = [
        ...inputs.map((local) => `${local.isArray ? 'ByRef' : 'ByVal'} ${declarationOf(local)}`),
        ...byRef.map((local) => `ByRef ${declarationOf(local)}`),
    ];

    const returns = result ? ` As ${result.type}` : '';
    const keyword = result ? 'Function' : 'Sub';
    const signature = `Private ${keyword} ${newName}(${parameters.join(', ')})${returns}`;

    const indent = /^[ \t]*/.exec(selected)?.[0] ?? '    ';
    const bodyIndent = indent || '    ';
    const declarations = [...moved, ...(result ? [result] : [])]
        .map((local) => `${bodyIndent}Dim ${declarationOf(local)}`);

    const lines: string[] = [signature];
    if (declarations.length > 0) {
        lines.push(...declarations, '');
    }
    lines.push(...trimTrailingBlanks(splitLines(selected)));
    if (result) {
        lines.push(`${bodyIndent}${newName} = ${result.name}`);
    }
    lines.push(`End ${keyword}`);
    const extracted = lines.join(eol);

    /* ---- the call, and the caller ----------------------------------------------------------- */

    const argumentNames = [...inputs, ...byRef].map((local) => local.isArray ? `${local.name}()` : local.name);
    const call = result
        ? `${indent}${result.name} = ${newName}(${argumentNames.join(', ')})`
        : `${indent}${newName}${argumentNames.length > 0 ? ` ${argumentNames.join(', ')}` : ''}`;

    // Applied back to front so each splice keeps the offsets the next one was measured against.
    const edits: { start: number; end: number; text: string }[] = [
        { start: selectionStart, end: selectionEnd, text: call + eol },
    ];

    for (const local of moved) {
        const cut = cutDeclaration(local, source, starts, startLine, endLine);
        if (cut) {
            edits.push(cut);
        }
    }

    // Below the procedure it came from, where a reader looks for it.
    const after = toLineColumn(starts, Math.max(procedure.span.start, procedure.span.end - 1)).line;
    const insertAt = after < lineCount ? (starts[after] as number) : source.length;
    const leading = insertAt === source.length && !source.endsWith(eol) ? eol : '';
    edits.push({ start: insertAt, end: insertAt, text: `${leading}${eol}${extracted}${eol}` });

    let rewritten = source;
    for (const edit of edits.sort((a, b) => b.start - a.start)) {
        rewritten = rewritten.slice(0, edit.start) + edit.text + rewritten.slice(edit.end);
    }

    return {
        module: moduleName,
        source: rewritten,
        procedure: newName,
        signature,
        from: procedure.name,
    };
}

/** `End Sub`, `End Function`, `End Property`, for a refusal that names what it looked for. */
function closerOf(procedure: ProcedureNode): string {
    return procedure.procKind === 'Sub' ? 'End Sub'
        : procedure.procKind === 'Function' ? 'End Function'
            : 'End Property';
}

/** `name As Long`, `name() As String`, the text a Dim and a parameter both want. */
function declarationOf(local: Local): string {
    return `${local.name}${local.isArray ? '()' : ''} As ${local.type}`;
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

/**
 * Control flow, blocks and receivers that cross the selection's edge.
 *
 * Each of these would compile into something that does not mean what the original meant, and each
 * one is a check a suite can drive. Answers the refusal, or nothing when the selection is whole.
 */
function crossings(
    nodes: readonly BodyNode[],
    source: string,
    starts: readonly number[],
    selected: string,
    startLine: number,
    endLine: number,
): string | undefined {
    const lineOf = (offset: number): number => toLineColumn(starts, offset).line;
    const body = stripComments(selected);

    for (const node of nodes) {
        const first = lineOf(node.span.start);
        const last = lineOf(Math.max(node.span.start, node.span.end - 1));
        const opens = first >= startLine && first <= endLine;
        const closes = last >= startLine && last <= endLine;
        const encloses = first < startLine && last > endLine;

        // A With the selection does not take whole leaves every leading dot without a receiver -
        // including the ordinary case where the whole selection sits INSIDE the block, which is
        // neither opening nor closing within it and so is invisible to the check below. Statements
        // inside a With that never use a dot move perfectly well, so the dots are what decide.
        if (node.kind === 'WithBlock' && (encloses || opens !== closes)) {
            if (encloses && !hasLeadingDot(body)) {
                continue;
            }

            return 'The selection is inside a With block, so its leading dots would have no receiver once they move. Select the whole With block, or take the statements that do not use it.';
        }

        // A block half in and half out: `If` without its `End If`, a `For` whose `Next` is below
        // the selection. Leaf statements have no inside to be half of.
        if ((node as { body?: readonly BodyNode[] }).body && opens !== closes) {
            return `The selection starts inside ${blockName(node.kind)} and ends outside it. Select whole blocks.`;
        }
    }

    if (/\bExit\s+(Sub|Function|Property)\b/i.test(body)) {
        return 'The selection leaves the procedure with Exit, which in a new procedure would only leave that one. Select code that runs to its end.';
    }

    // An Exit that leaves a loop is fine when the loop came too.
    for (const match of body.matchAll(/\bExit\s+(For|Do|While)\b/gi)) {
        const loop = (match[1] as string).toLowerCase();
        const kind = loop === 'for' ? 'ForBlock' : loop === 'do' ? 'DoBlock' : 'WhileBlock';
        const enclosing = nodes.some((node) => node.kind === kind
            && lineOf(node.span.start) >= startLine
            && lineOf(Math.max(node.span.start, node.span.end - 1)) <= endLine);
        if (!enclosing) {
            return `The selection leaves a loop with Exit ${match[1]}, and the loop is outside it. Select the loop too.`;
        }
    }

    if (/\bResume\b/i.test(body)) {
        return 'The selection carries Resume, which resumes the procedure that handled the error. Extracting it would move where it resumes to.';
    }

    // A jump whose landing place is not coming with it, and a landing place jumped to from
    // outside: both leave one half of the pair unable to see the other.
    const labelsInside = new Set(labelsOf(body).map((label) => label.toLowerCase()));
    for (const match of body.matchAll(/\b(?:GoTo|GoSub)\s+([A-Za-z_][\w]*|\d+)/gi)) {
        const target = (match[1] as string).toLowerCase();
        if (target !== '0' && !labelsInside.has(target)) {
            return `The selection jumps to '${match[1]}', which is outside it. Select the label too, or leave the jump behind.`;
        }
    }

    if (labelsInside.size > 0) {
        const outside = stripComments(source.slice(0, starts[startLine - 1] as number))
            + stripComments(source.slice(endLine < starts.length ? (starts[endLine] as number) : source.length));
        for (const match of outside.matchAll(/\b(?:GoTo|GoSub|Resume)\s+([A-Za-z_][\w]*|\d+)/gi)) {
            if (labelsInside.has((match[1] as string).toLowerCase())) {
                return `'${match[1]}' is jumped to from outside the selection, so its label has to stay where it is.`;
            }
        }
    }

    return undefined;
}

/** The words a refusal uses for a block, article and all, from the parser's node kind. */
function blockName(kind: string): string {
    switch (kind) {
        case 'IfBlock': return 'an If block';
        case 'ForBlock': return 'a For loop';
        case 'DoBlock': return 'a Do loop';
        case 'WhileBlock': return 'a While loop';
        case 'WithBlock': return 'a With block';
        case 'SelectBlock': return 'a Select Case block';
        default: return 'a block';
    }
}

/**
 * A member access with nothing before the dot, which is what a `With` block's receiver supplies.
 * The dot has to follow the start of an expression rather than a name or a closing bracket, or
 * every `Sheet1.Cells` in the selection would read as one.
 */
function hasLeadingDot(text: string): boolean {
    return /(^|[\s(,=+\-*/&<>^\\:])\.[A-Za-z_]/m.test(text);
}

/** Line labels: `Fail:` and `10` at the head of a line, the two things a GoTo can name. */
function labelsOf(text: string): string[] {
    const found: string[] = [];
    for (const line of splitLines(text)) {
        const match = /^[ \t]*([A-Za-z_][\w]*|\d+):/.exec(line);
        if (match) {
            found.push(match[1] as string);
        }
    }
    return found;
}

/**
 * The text with comments and string literals blanked, so a keyword quoted in a message is not read
 * as the statement it names. Length is preserved, because the callers measure offsets in it.
 */
function stripComments(text: string): string {
    let out = '';
    let inString = false;
    let inComment = false;

    for (let at = 0; at < text.length; at++) {
        const character = text[at] as string;

        if (character === '\n' || character === '\r') {
            inString = false;
            inComment = false;
            out += character;
            continue;
        }

        if (inComment) {
            out += ' ';
            continue;
        }

        if (character === '"') {
            inString = !inString;
            out += ' ';
            continue;
        }

        if (inString) {
            out += ' ';
            continue;
        }

        if (character === "'") {
            inComment = true;
            out += ' ';
            continue;
        }

        out += character;
    }

    return out;
}

/** Splits on any line ending without deciding which one the file uses. */
function splitLines(text: string): string[] {
    return text.split(/\r\n|\r|\n/);
}

/** Drops the empty lines a whole-line selection collects at its end. */
function trimTrailingBlanks(lines: string[]): string[] {
    const out = [...lines];
    while (out.length > 0 && !/\S/.test(out[out.length - 1] as string)) {
        out.pop();
    }
    return out;
}

/** Every local and parameter of one procedure, with what its declaration says about it. */
function localsOf(procedure: ProcedureNode, starts: readonly number[]): Local[] {
    const locals: Local[] = [];

    for (const parameter of procedure.params) {
        locals.push({
            name: parameter.name,
            type: parameter.asType ?? typeOfSuffix(parameter.typeSuffix),
            isArray: parameter.isArray,
            isParameter: true,
            isStatic: false,
            isConst: false,
            declaredAt: parameter.nameSpan ? toLineColumn(starts, parameter.nameSpan.start) : undefined,
            anchor: parameter.nameSpan?.start ?? parameter.span.start,
        });
    }

    for (const node of flatten(procedure.body)) {
        if (node.kind !== 'VariableGroup') {
            continue;
        }

        const group = node;
        for (const declaration of group.declarations) {
            locals.push({
                name: declaration.name,
                type: declaration.asType ?? typeOfSuffix(declaration.typeSuffix),
                isArray: declaration.isArray,
                isParameter: false,
                isStatic: /^static$/i.test(group.modifier ?? ''),
                isConst: group.isConst,
                group,
                declaredAt: declaration.nameSpan ? toLineColumn(starts, declaration.nameSpan.start) : undefined,
                anchor: declaration.nameSpan?.start ?? declaration.span.start,
            });
        }
    }

    return locals;
}

/** `Long` from `&`, and so on; Variant when a declaration says nothing at all. */
function typeOfSuffix(suffix: string | undefined): string {
    switch (suffix) {
        case '%': return 'Integer';
        case '&': return 'Long';
        case '@': return 'Currency';
        case '!': return 'Single';
        case '#': return 'Double';
        case '$': return 'String';
        default: return 'Variant';
    }
}

/** Every use of one local in its own module, the declaration itself left out. */
function usesOf(
    local: Local,
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
): LocationPayload[] {
    const found = referencesFor(symbols, moduleName, source, local.anchor, true);
    return found
        .filter((one) => one.module.toLowerCase() === moduleName.toLowerCase())
        .filter((one) => !(local.declaredAt
            && one.line === local.declaredAt.line
            && one.column === local.declaredAt.column))
        .sort((a, b) => a.line - b.line || a.column - b.column);
}

/** Whether the selection mentions this local at all. */
function usedIn(
    local: Local,
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    startLine: number,
    endLine: number,
): boolean {
    return usesOf(local, symbols, moduleName, source)
        .some((one) => one.line >= startLine && one.line <= endLine);
}

/**
 * What the selection does with one local, in the terms the signature is decided in.
 *
 * `readsBeforeWriting` is the one that matters most: a local the selection reads before it assigns
 * anything to it needs the caller's value, and one it assigns first does not. That single fact is
 * the difference between a parameter and a local of the new procedure, and it is the analyzer's
 * answer rather than this file's guess.
 */
function usageOf(
    local: Local,
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    startLine: number,
    endLine: number,
): Usage {
    const uses = usesOf(local, symbols, moduleName, source);
    const inside = uses.filter((one) => one.line >= startLine && one.line <= endLine);
    const before = uses.filter((one) => one.line < startLine);
    const after = uses.filter((one) => one.line > endLine);

    const reads = (one: LocationPayload): boolean => one.kind === 'read' || one.kind === 'readwrite';
    const writes = (one: LocationPayload): boolean => one.kind === 'write' || one.kind === 'readwrite';

    // BY LINE, NOT BY COLUMN, because a statement evaluates its right-hand side before it assigns.
    // `total = total + step` reports the write at column 5 and the read at column 13, so reading
    // them in column order says the callee writes total before it reads it - and the extraction
    // then makes it the return value instead of a ByRef parameter, silently dropping the caller's
    // running total. A read anywhere on the first line that writes counts as reading first, which
    // also errs the safe way: passing a value the callee did not need costs an argument, and not
    // passing one it did costs the developer their data.
    const firstWriteLine = Math.min(...inside.filter(writes).map((one) => one.line), Number.POSITIVE_INFINITY);

    return {
        readsBeforeWriting: inside.some((one) => reads(one) && one.line <= firstWriteLine),
        writesInside: inside.some(writes),
        readsInside: inside.some(reads),
        readsAfter: after.some(reads),
        usedBefore: before.length > 0,
        usedAfter: after.length > 0,
        usedInside: inside.length > 0,
    };
}

/**
 * The span to cut when a Dim moves into the new procedure.
 *
 * A group declaring one name goes whole, line and all. A group declaring several - `Dim i As Long,
 * j As Long` - is left alone: cutting one name out of it is a text edit on a line the developer
 * wrote deliberately, and the cost of leaving it is an unused local, which the analyzer already
 * reports.
 */
function cutDeclaration(
    local: Local,
    source: string,
    starts: readonly number[],
    startLine: number,
    endLine: number,
): { start: number; end: number; text: string } | undefined {
    const group = local.group;
    if (!group || group.declarations.length !== 1) {
        return undefined;
    }

    const first = toLineColumn(starts, group.span.start).line;
    const last = toLineColumn(starts, Math.max(group.span.start, group.span.end - 1)).line;
    if (first !== last) {
        return undefined;
    }

    // Inside the selection already: it moves with the text, and cutting it as well would take the
    // line out of the new procedure too.
    if (first >= startLine && first <= endLine) {
        return undefined;
    }

    const from = starts[first - 1] as number;
    const to = first < starts.length ? (starts[first] as number) : source.length;
    return { start: from, end: to, text: '' };
}
