/*
 * Introduce Parameter: a local becomes a parameter, and every caller passes what it was assigned.
 *
 * THE VALUE HAS TO TRAVEL, and where it travels to is the whole difficulty. A local's initialiser
 * is written in the procedure's own scope, so `rate = basePrice * 1.2` means nothing at a call site
 * in another module: `basePrice` is not there. Moving the expression out only works when everything
 * it names is visible from every caller, and that is a check rather than a hope.
 *
 * So the initialiser must be made of things every caller can see: literals, and names that resolve
 * to something the whole project can reach. A local, a parameter or anything Private to this
 * module fails it, and the refusal names what would not travel.
 *
 * THE SIGNATURE IS NOT ALWAYS OURS TO CHANGE. An event handler's shape is fixed by whatever raises
 * it, and a member implementing an interface is fixed by the interface. Adding a parameter to
 * either produces a procedure that no longer binds to the thing it was written for, which VBA
 * reports somewhere else entirely - so both are refused.
 */

import { parseModule, type BodyNode, type ModuleNode, type ProcedureNode, type VariableGroupNode } from '../../../xlide_vscode/src/analyzer';
import { lineStarts, referencesFor, toLineColumn, type ProjectSymbols } from './navigation';
import type { IntroduceParameterResult, LocationPayload } from './protocol';

export function introduceParameterFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
): IntroduceParameterResult {
    const refuse = (why: string): IntroduceParameterResult => ({ modules: [], refused: why });

    const module = parseModule(source);
    const procedure = module.members.find(
        (member): member is ProcedureNode => member.kind === 'Procedure'
            && member.span.start <= offset && member.span.end >= offset);

    if (!procedure) {
        return refuse('Put the caret on a local variable inside a procedure.');
    }

    const named = identifierAt(source, offset);
    if (!named) {
        return refuse('There is no name here to make a parameter.');
    }

    if (procedure.params.some((one) => one.name.toLowerCase() === named.toLowerCase())) {
        return refuse(`'${named}' is already a parameter of '${procedure.name}'.`);
    }

    /* ---- whose signature is this to change ------------------------------------------------------ */

    const underscore = procedure.name.indexOf('_');
    if (underscore > 0) {
        const before = procedure.name.slice(0, underscore);
        if (symbols.byModule.has(before.toLowerCase())) {
            return refuse(`'${procedure.name}' is bound by its name to '${before}' - an event handler or an interface member - so its signature is not this project's to change.`);
        }
    }

    const starts = lineStarts(source);
    const group = declarationOf(procedure, named, starts);
    if (!group) {
        return refuse(`'${named}' is not a local declared in '${procedure.name}'. Only a local can become a parameter.`);
    }

    if (group.group.isConst) {
        return refuse(`'${group.declared}' is a Const, whose value is fixed at compile time. A parameter is not.`);
    }

    if (group.group.declarations.length > 1) {
        return refuse(`'${group.declared}' shares its declaration with other names. Give it a line of its own first.`);
    }

    /* ---- what it is assigned, and whether that can travel --------------------------------------- */

    const uses = referencesFor(symbols, moduleName, source, offset, true)
        .filter((one) => one.module.toLowerCase() === moduleName.toLowerCase())
        .filter((one) => !(one.line === group.line && one.column === group.column))
        .sort((a, b) => a.line - b.line || a.column - b.column);

    const writes = uses.filter((one) => one.kind === 'write' || one.kind === 'readwrite');
    if (writes.length === 0) {
        return refuse(`'${group.declared}' is never assigned, so there is no value for a caller to pass.`);
    }

    if (writes.length > 1) {
        return refuse(`'${group.declared}' is assigned ${writes.length} times, so it does not hold one value the caller could supply.`);
    }

    const assignment = writes[0] as LocationPayload;
    const value = assignedValue(source, starts, assignment, group.declared);
    if (value.refused) {
        return refuse(value.refused);
    }

    // Everything the value names has to be visible from a call site. A local of this procedure is
    // the case that bites: the expression reads perfectly here and means nothing there.
    const local = new Set<string>([
        ...procedure.params.map((one) => one.name.toLowerCase()),
        ...localNamesOf(procedure),
    ]);
    const privateHere = privateNamesOf(module);

    const stranded = [...namesIn(value.text!)]
        .filter((one) => local.has(one) || privateHere.has(one))
        .filter((one) => one !== group.declared.toLowerCase());

    if (stranded.length > 0) {
        return refuse(`'${group.declared}' is assigned from '${value.text}', which uses ${stranded.slice(0, 3).map((one) => `'${one}'`).join(', ')} - not visible where the callers are. A parameter's value has to be one a caller can write.`);
    }

    /* ---- the signature, and every call site ------------------------------------------------------ */

    const type = group.type;
    const parameter = `ByVal ${group.declared} As ${type}`;
    const header = headerOf(source, starts, procedure);
    if (header.refused) {
        return refuse(header.refused);
    }

    const withParameter = header.text!.replace(
        /\(([^)]*)\)/,
        (_whole, inside: string) => `(${inside.trim().length === 0 ? parameter : `${inside}, ${parameter}`})`);

    if (withParameter === header.text) {
        return refuse(`'${procedure.name}' has no parameter list this can add to.`);
    }

    // EVERY EDIT AGAINST THE ORIGINAL TEXT, including the call sites in this same module. The
    // first cut of this rewrote the module and then hunted for its own call sites in the result,
    // where two lines had moved - which is a coordinate problem invented for no reason, since the
    // engine has the original in hand and can measure everything against it once.
    const callers = passAtCallSites(symbols, moduleName, source, procedure, value.text!);

    const edits: { start: number; end: number; text: string }[] = [
        { start: header.from!, end: header.to!, text: withParameter },
        // The declaration and its assignment both go: the caller supplies the value now.
        cutLine(source, starts, group.line),
        cutLine(source, starts, assignment.line),
        ...(callers.here ?? []),
    ];

    let rewritten = source;
    for (const edit of edits.sort((a, b) => b.start - a.start)) {
        rewritten = rewritten.slice(0, edit.start) + edit.text + rewritten.slice(edit.end);
    }

    const modules = [{ module: moduleName, source: rewritten }, ...callers.modules];

    return {
        modules,
        parameter: group.declared,
        type,
        value: value.text!,
        procedure: procedure.name,
        callSites: callers.count,
    };
}

/** The declaration of one local, with where its name sits and what type it carries. */
function declarationOf(procedure: ProcedureNode, named: string, starts: readonly number[]):
    { group: VariableGroupNode; declared: string; type: string; line: number; column: number } | null {
    for (const node of flatten(procedure.body)) {
        if (node.kind !== 'VariableGroup') {
            continue;
        }

        const found = node.declarations.find((one) => one.name.toLowerCase() === named.toLowerCase());
        if (!found?.nameSpan) {
            continue;
        }

        const at = toLineColumn(starts, found.nameSpan.start);
        return {
            group: node,
            declared: found.name,
            type: found.asType ?? 'Variant',
            line: at.line,
            column: at.column,
        };
    }

    return null;
}

/** What the one assignment gives it, when the line is one this can read. */
function assignedValue(
    source: string,
    starts: readonly number[],
    assignment: LocationPayload,
    declared: string,
): { text?: string; refused?: string } {
    const from = starts[assignment.line - 1] as number;
    const to = assignment.line < starts.length ? (starts[assignment.line] as number) : source.length;
    const line = source.slice(from, to);

    const equals = line.indexOf('=');
    if (equals < 0) {
        return { refused: `'${declared}' is not assigned with '=' on its own line, which is the only shape this can read.` };
    }

    const before = line.slice(0, equals);
    if (!new RegExp(`^\\s*(?:Set\\s+)?${declared.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*$`, 'iu').test(before)) {
        return { refused: `The line assigning '${declared}' does more than assign it, so what a caller would pass cannot be read from it.` };
    }

    const text = line.slice(equals + 1).replace(/\r|\n/g, '').trim();
    if (text.length === 0 || text.endsWith('_')) {
        return { refused: `'${declared}' is assigned over more than one line, or nothing this can read.` };
    }

    return { text };
}

/** The procedure's header line, which is where the parameter goes. */
function headerOf(source: string, starts: readonly number[], procedure: ProcedureNode):
    { text?: string; from?: number; to?: number; refused?: string } {
    const line = toLineColumn(starts, procedure.span.start).line;
    const from = starts[line - 1] as number;
    const to = line < starts.length ? (starts[line] as number) : source.length;
    const text = source.slice(from, to);

    if (/_\s*$/.test(text.replace(/\r|\n/g, ''))) {
        return { refused: `'${procedure.name}' declares its parameters over more than one line. Put the header on one line first.` };
    }

    return { text, from, to };
}

/** A whole line removed, newline and all. */
function cutLine(source: string, starts: readonly number[], line: number): { start: number; end: number; text: string } {
    return {
        start: starts[line - 1] as number,
        end: line < starts.length ? (starts[line] as number) : source.length,
        text: '',
    };
}

/**
 * Every call site given the value the local used to be assigned.
 *
 * A call with no arguments gains one; a call with arguments gains another at the end, which is
 * where the parameter was added. A NAMED-argument call is left alone and counted as refused
 * ground: adding a positional argument to `Foo bar:=1` changes what VBA binds where.
 */
function passAtCallSites(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    procedure: ProcedureNode,
    value: string,
): { modules: { module: string; source: string }[]; here: { start: number; end: number; text: string }[]; count: number } {
    if (!procedure.nameSpan) {
        return { modules: [], here: [], count: 0 };
    }

    const uses = referencesFor(symbols, moduleName, source, procedure.nameSpan.start, false);
    const byModule = new Map<string, LocationPayload[]>();

    for (const use of uses) {
        const rows = byModule.get(use.module) ?? [];
        rows.push(use);
        byModule.set(use.module, rows);
    }

    const modules: { module: string; source: string }[] = [];
    let here: { start: number; end: number; text: string }[] = [];
    let count = 0;

    for (const [named, rows] of byModule) {
        const mine = named.toLowerCase() === moduleName.toLowerCase();
        const text = mine ? source : symbols.byModule.get(named.toLowerCase())?.source;
        if (text === undefined) {
            continue;
        }

        const starts = lineStarts(text);
        const edits: { start: number; end: number; text: string }[] = [];

        for (const row of rows) {
            const at = (starts[row.line - 1] as number) + row.column - 1;
            const lineEnd = row.line < starts.length ? (starts[row.line] as number) : text.length;
            const rest = text.slice(at + procedure.name.length, lineEnd).replace(/\r?\n$/, '');

            // A NAMED argument anywhere in the call: left alone, because adding a positional one
            // to `Foo bar:=1` changes what VBA binds where. Counted as untouched rather than
            // rewritten, so the answer's count is of what actually moved.
            if (/:=/.test(rest)) {
                continue;
            }

            const trimmed = rest.trim();

            // AN EMPTY PARAMETER LIST TAKES THE VALUE ALONE. `Total()` becomes `Total(3)`, not
            // `Total(, 3)`, which is what a blanket comma produced - and not `Total(3, 3)`, which
            // is what two replacements produced when the second saw what the first had written.
            // One decision, made once.
            let replacement: string;
            if (trimmed.startsWith('(')) {
                const empty = /^\(\s*\)\s*$/.test(trimmed);
                replacement = empty
                    ? rest.replace(/\(\s*\)(\s*)$/, `(${value})$1`)
                    : rest.replace(/\)(\s*)$/, `, ${value})$1`);
            } else if (trimmed.length === 0) {
                replacement = ` ${value}`;
            } else {
                replacement = `${rest.replace(/\s+$/, '')}, ${value}`;
            }

            edits.push({
                start: at + procedure.name.length,
                end: at + procedure.name.length + rest.length,
                text: replacement,
            });
            count += 1;
        }

        if (edits.length === 0) {
            continue;
        }

        if (mine) {
            here = edits;
            continue;
        }

        let updated = text;
        for (const edit of edits.sort((a, b) => b.start - a.start)) {
            updated = updated.slice(0, edit.start) + edit.text + updated.slice(edit.end);
        }

        modules.push({ module: named, source: updated });
    }

    return { modules, here, count };
}

/** Every local a procedure declares, lowercased. */
function localNamesOf(procedure: ProcedureNode): string[] {
    const names: string[] = [];
    for (const node of flatten(procedure.body)) {
        if (node.kind === 'VariableGroup') {
            for (const one of node.declarations) {
                names.push(one.name.toLowerCase());
            }
        }
    }

    return names;
}

/** Every name Private to this module, which a caller elsewhere cannot see either. */
function privateNamesOf(module: ModuleNode): Set<string> {
    const names = new Set<string>();
    for (const member of module.members) {
        if (member.kind === 'VariableGroup' && !/^(public|global)$/i.test(member.modifier ?? '')) {
            for (const one of member.declarations) {
                names.add(one.name.toLowerCase());
            }
        }

        if (member.kind === 'Procedure' && member.modifiers.some((one) => /^private$/i.test(one))) {
            names.add(member.name.toLowerCase());
        }
    }

    return names;
}

/** Every identifier a stretch of text mentions, lowercased. */
function namesIn(text: string): Set<string> {
    const found = new Set<string>();
    for (const match of text.matchAll(/[\p{L}_][\p{L}\p{M}\p{N}_]*/gu)) {
        found.add(match[0].toLowerCase());
    }

    return found;
}

/** The identifier covering an offset. */
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
