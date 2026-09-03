/*
 * Extract Variable: a selected expression given a name.
 *
 * `Sheets("Data").Range("A1").Value` written twice on one line is the ordinary shape of VBA, and
 * naming it means retyping the whole thing plus a Dim plus a type. That is the refactoring people
 * reach for most after rename, and it needed one fact this product could not derive: what type the
 * selected expression has, and whether assigning it needs `Set`. VBA has no assignment form that
 * works for both an object and a value, so guessing emits a line that does not compile.
 *
 * The analyzer answers it now - `resolveExpressionType` (xlide_vscode#61) - so this is a lookup
 * and two edits: the declaration and its assignment above the statement, and the selection
 * replaced by the name.
 *
 * ONLY THE SELECTION IS REPLACED, not every occurrence of the same text. Extracting an expression
 * that appears three times and evaluating it once is a behaviour change whenever the expression
 * has a side effect - a function call, a Range read that fires a formula - and nothing in VBA
 * says which expressions are pure. A developer who wants the other two replaced can see the name
 * and replace them.
 */

import { parseModule, resolveExpressionType, type BodyNode, type ProcedureNode } from '../../../xlide_vscode/src/analyzer';
import { assembleContext } from './moduleContext';
import { lineStarts, toLineColumn } from './navigation';
import type { ExtractVariableParams, ExtractVariableResult, ModulePayload } from './protocol';

/** The same rule the rename box applies: a letter, then letters, digits and underscores. */
const IDENTIFIER = /^\p{L}[\p{L}\p{M}\p{N}_]{0,254}$/u;

/** Keywords a declaration cannot be named after. The analyzer catches the rest on the next pass. */
const RESERVED = new Set([
    'as', 'byref', 'byval', 'call', 'case', 'const', 'dim', 'do', 'each', 'else', 'elseif', 'end',
    'error', 'exit', 'false', 'for', 'function', 'get', 'goto', 'if', 'in', 'is', 'let', 'loop',
    'me', 'new', 'next', 'nothing', 'null', 'on', 'option', 'private', 'property', 'public',
    'redim', 'resume', 'return', 'select', 'set', 'sub', 'then', 'to', 'true', 'until', 'wend',
    'while', 'with',
]);

export function extractVariableFor(
    seeded: readonly ModulePayload[],
    params: ExtractVariableParams & { source: string },
): ExtractVariableResult {
    const refuse = (why: string): ExtractVariableResult => ({ refused: why });
    const { source, newName } = params;

    if (!IDENTIFIER.test(newName)) {
        return refuse(`'${newName}' is not a VBA name. A name starts with a letter and holds letters, digits and underscores.`);
    }

    if (RESERVED.has(newName.toLowerCase())) {
        return refuse(`'${newName}' is a VBA keyword.`);
    }

    const start = Math.min(params.startOffset, params.endOffset);
    const end = Math.max(params.startOffset, params.endOffset);
    if (start < 0 || end > source.length || start === end) {
        return refuse('Select the expression to give a name to.');
    }

    const module = parseModule(source);
    const procedure = module.members.find(
        (member): member is ProcedureNode => member.kind === 'Procedure'
            && member.span.start <= start && member.span.end >= end);

    if (!procedure) {
        return refuse('Select an expression inside a procedure. A declaration outside one has nowhere to put the variable.');
    }

    // The statement the selection sits in, which is what the declaration goes above. The INNERMOST
    // one: a block's span contains its body, and inserting above the `If` rather than above the
    // line inside it would move the expression out of the branch that guards it.
    const statement = innermost(procedure.body, start, end);
    if (!statement) {
        return refuse('Select an expression inside a statement.');
    }

    if (statement.kind === 'VariableGroup') {
        return refuse('That is part of a declaration, not an expression that can be given a name.');
    }

    // The target of an assignment is the name being written, not a value to be named: extracting
    // it produces a variable holding the OLD value and a line that assigns to it instead.
    const after = source.slice(end, end + 40);
    if (statement.span.start === start && /^\s*=/.test(after)) {
        return refuse(`'${source.slice(start, end).trim()}' is the name being assigned to, not a value. Select what is on the right of the '='.`);
    }

    const taken = declaredNames(module, procedure);
    if (taken.has(newName.toLowerCase())) {
        return refuse(`'${newName}' is already declared here.`);
    }

    const context = assembleContext(seeded, { ...params, moduleName: params.moduleName });
    const resolved = resolveExpressionType(source, { start, end }, {
        moduleName: context.current.name,
        moduleKind: context.moduleKind,
        model: context.hostModel,
        projectClassMembers: context.projectClassMembers,
        projectVisibleSymbols: context.projectSymbols,
    });

    if (!resolved) {
        return refuse('The analyzer could not read that as an expression.');
    }

    if (!resolved.complete) {
        return refuse(`'${source.slice(start, end).trim()}' is not a whole expression. Select all of one - a partial expression cannot be given a name.`);
    }

    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const starts = lineStarts(source);
    const line = toLineColumn(starts, statement.span.start).line;
    const lineStart = starts[line - 1] as number;
    // FROM THE LINE, not from the text before the statement. A statement's span begins at column
    // one - the indent is part of it - so measuring the gap between the line start and the span
    // start is always empty, and every declaration came out hard against the left margin inside a
    // procedure that was indented around it.
    const lineEnd = line < starts.length ? (starts[line] as number) : source.length;
    const indent = /^[ \t]*/.exec(source.slice(lineStart, lineEnd))?.[0] ?? '';
    const expression = source.slice(start, end).trim();

    // `Set` for an object and nothing for a value, which is the fact the analyzer had to supply:
    // VBA has no form that works for both, and the declared type name alone cannot decide it -
    // a project class, a host type and a Variant holding an object all look different.
    const assigns = resolved.isObject ? 'Set ' : '';
    const declaration = `${indent}Dim ${newName} As ${resolved.type}${eol}`
        + `${indent}${assigns}${newName} = ${expression}${eol}`;

    // Back to front, so the first splice keeps the offsets the second was measured against.
    const rewritten = source.slice(0, start) + newName + source.slice(end);
    const withDeclaration = rewritten.slice(0, lineStart) + declaration + rewritten.slice(lineStart);

    return {
        module: params.moduleName,
        source: withDeclaration,
        variable: newName,
        type: resolved.type,
        isObject: resolved.isObject,
        expression,
    };
}

/** The innermost leaf statement covering the span, or undefined when nothing does. */
function innermost(body: readonly BodyNode[], start: number, end: number): BodyNode | undefined {
    let found: BodyNode | undefined;

    const visit = (nodes: readonly BodyNode[]): void => {
        for (const node of nodes) {
            if (node.span.start > start || node.span.end < end) {
                continue;
            }

            const nested = (node as { body?: readonly BodyNode[] }).body;
            if (nested) {
                // A block CONTAINS the statement; the header line is a statement of its own that
                // the walk below reaches. Only descend, so a block never becomes the answer.
                visit(nested);
                continue;
            }

            if (!found || node.span.start > found.span.start) {
                found = node;
            }
        }
    };

    visit(body);
    return found;
}

/** Every name already meaning something here: the module's members and this procedure's locals. */
function declaredNames(module: ReturnType<typeof parseModule>, procedure: ProcedureNode): Set<string> {
    const names = new Set<string>();

    for (const member of module.members) {
        if (member.kind === 'Procedure') {
            names.add(member.name.toLowerCase());
        } else if (member.kind === 'VariableGroup') {
            for (const one of member.declarations) {
                names.add(one.name.toLowerCase());
            }
        }
    }

    for (const parameter of procedure.params) {
        names.add(parameter.name.toLowerCase());
    }

    const locals = (nodes: readonly BodyNode[]): void => {
        for (const node of nodes) {
            if (node.kind === 'VariableGroup') {
                for (const one of node.declarations) {
                    names.add(one.name.toLowerCase());
                }
            }

            const nested = (node as { body?: readonly BodyNode[] }).body;
            if (nested) {
                locals(nested);
            }
        }
    };

    locals(procedure.body);
    return names;
}
