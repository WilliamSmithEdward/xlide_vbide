/*
 * Move to Module: a procedure taken out of one standard module and put into another.
 *
 * THE MOVE IS THE EASY HALF. What decides whether it can happen at all is what the procedure
 * reaches for: a Private variable, Const, Type, Enum, Declare or procedure of the module it is
 * leaving is not reachable from anywhere else, so a procedure that touches one cannot go without
 * taking that with it - and deciding what else to drag along is a judgement, not a rewrite.
 *
 * A PUBLIC member of the old module is a different matter: VBA resolves an unqualified name across
 * the project, so a call to it keeps working from the new home. That is also why most call sites
 * need no editing at all. The ones that DO are the qualified ones - `Helpers.Recalc` names the
 * module, and the module is what changed - and those are rewritten wherever they are.
 *
 * `Option Private Module` turns every Public member of a module into a project-private one, so a
 * module carrying it makes the reasoning above false and the move is refused rather than reasoned
 * about differently.
 */

import { parseModule, type ModuleNode, type ProcedureNode } from '../../../xlide_vscode/src/analyzer';
import { lineStarts, referencesFor, toLineColumn, type ProjectSymbols } from './navigation';
import type { MoveToModuleResult } from './protocol';

export function moveToModuleFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
    targetName: string,
): MoveToModuleResult {
    const refuse = (why: string): MoveToModuleResult => ({ modules: [], refused: why });

    const module = parseModule(source);
    const procedure = module.members.find(
        (member): member is ProcedureNode => member.kind === 'Procedure'
            && member.span.start <= offset && member.span.end >= offset);

    if (!procedure) {
        return refuse('Put the caret inside the procedure to move.');
    }

    if (!procedure.closed) {
        return refuse(`'${procedure.name}' has no closing End, so there is nothing whole to move.`);
    }

    if (targetName.toLowerCase() === moduleName.toLowerCase()) {
        return refuse(`'${procedure.name}' is already in '${moduleName}'.`);
    }

    const target = symbols.byModule.get(targetName.toLowerCase());
    if (!target) {
        return refuse(`This project has no module called '${targetName}'.`);
    }

    // Standard modules both ways. A procedure in a class or a form is a MEMBER of that type -
    // moving it changes what the type can do, which is a design change rather than a relocation -
    // and a document module's procedures are the document's.
    if (!isStandard(symbols, moduleName)) {
        return refuse(`'${moduleName}' is not a standard module, and a procedure in a class, form or document belongs to it. Only a standard module's procedures can move.`);
    }

    if (!isStandard(symbols, targetName)) {
        return refuse(`'${targetName}' is not a standard module, so a procedure moved into it would become one of its members.`);
    }

    const targetSource = target.source ?? '';
    const targetModule = parseModule(targetSource);

    if (hasPrivateModuleOption(module) || hasPrivateModuleOption(targetModule)) {
        return refuse('One of these modules declares Option Private Module, which changes what its Public members can reach from outside. This one has to be done by hand.');
    }

    if (targetModule.members.some((member) => member.kind === 'Procedure'
        && member.name.toLowerCase() === procedure.name.toLowerCase())) {
        return refuse(`'${targetName}' already declares a procedure called '${procedure.name}'.`);
    }

    // An event handler belongs to whatever raises the event, and its name is the binding.
    if (/^[\p{L}_][\p{L}\p{M}\p{N}_]*_[\p{L}_][\p{L}\p{M}\p{N}_]*$/u.test(procedure.name)
        && symbols.byModule.has(procedure.name.slice(0, procedure.name.indexOf('_')).toLowerCase())) {
        return refuse(`'${procedure.name}' reads as an event handler for '${procedure.name.slice(0, procedure.name.indexOf('_'))}', and a handler is bound to the thing that raises it.`);
    }

    /* ---- what it reaches for that would not come with it ---------------------------------------- */

    const kept = privateNamesOf(module, procedure);
    const mentioned = namesIn(source.slice(procedure.span.start, procedure.span.end));
    const stranded = [...kept.keys()].filter((one) => mentioned.has(one));

    if (stranded.length > 0) {
        const names = stranded.slice(0, 4).map((one) => `'${kept.get(one)}'`).join(', ');
        const more = stranded.length > 4 ? `, and ${stranded.length - 4} more` : '';
        return refuse(`'${procedure.name}' uses ${names}${more}, which ${stranded.length === 1 ? 'is' : 'are'} Private to '${moduleName}' and would not be reachable from '${targetName}'.`);
    }

    /* ---- the three edits: out, in, and the qualified call sites ---------------------------------- */

    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const starts = lineStarts(source);
    const first = toLineColumn(starts, procedure.span.start).line;
    const last = toLineColumn(starts, Math.max(procedure.span.start, procedure.span.end - 1)).line;

    const from = starts[first - 1] as number;
    const to = last < starts.length ? (starts[last] as number) : source.length;
    const text = source.slice(from, to).replace(/\s+$/, '');

    // Any blank line the procedure was separated by goes with it, so the module it leaves does not
    // collect a gap for every move.
    let cutFrom = from;
    while (cutFrom >= 2 && /\r?\n\r?\n$/.test(source.slice(Math.max(0, cutFrom - 2), cutFrom))) {
        cutFrom -= source.slice(0, cutFrom).endsWith('\r\n') ? 2 : 1;
    }

    const without = source.slice(0, cutFrom) + source.slice(to);
    const targetEol = targetSource.includes('\r\n') ? '\r\n' : eol;
    const tail = targetSource.length === 0 || targetSource.endsWith(targetEol) ? '' : targetEol;
    const withIt = `${targetSource}${tail}${targetEol}${text.split(/\r\n|\r|\n/).join(targetEol)}${targetEol}`;

    const modules = [
        { module: moduleName, source: without },
        { module: targetName, source: withIt },
    ];

    const qualified = requalified(symbols, moduleName, source, procedure, targetName, targetSource);
    if (qualified.refused) {
        return refuse(qualified.refused);
    }

    for (const one of qualified.modules) {
        // The module being emptied may also carry a qualified call to what it is losing.
        const already = modules.find((held) => held.module.toLowerCase() === one.module.toLowerCase());
        if (already) {
            continue;
        }

        modules.push(one);
    }

    return {
        modules,
        moved: procedure.name,
        from: moduleName,
        to: targetName,
        requalified: qualified.count,
    };
}

/** Whether a module is a standard one, which is the only kind whose procedures can move. */
function isStandard(symbols: ProjectSymbols, moduleName: string): boolean {
    const held = symbols.byModule.get(moduleName.toLowerCase());
    const type = (held?.type ?? '').toLowerCase();
    return type === '' || type === 'standard' || type === 'module';
}

function hasPrivateModuleOption(module: ModuleNode): boolean {
    return module.members.some((member) => member.kind === 'Option'
        && /^private\s+module\b/i.test(member.optionText.trim()));
}

/**
 * Everything Private to this module that a moved procedure would leave behind, by lowercase name.
 *
 * A Public member is not here, deliberately: VBA resolves an unqualified name across the project,
 * so the moved procedure goes on reaching it. What cannot survive the move is what nothing outside
 * this module can see.
 */
function privateNamesOf(module: ModuleNode, moving: ProcedureNode): Map<string, string> {
    const names = new Map<string, string>();
    const isPrivate = (modifiers: readonly (string | undefined)[]): boolean =>
        modifiers.some((one) => /^private$/i.test(one ?? ''));

    for (const member of module.members) {
        switch (member.kind) {
            case 'Procedure':
                if (member !== moving && isPrivate(member.modifiers)) {
                    names.set(member.name.toLowerCase(), member.name);
                }
                break;

            case 'VariableGroup':
                // A module-level `Dim` is Private, and so is an unmarked one.
                if (!/^(public|global)$/i.test(member.modifier ?? '')) {
                    for (const one of member.declarations) {
                        names.set(one.name.toLowerCase(), one.name);
                    }
                }
                break;

            case 'Type':
            case 'Enum':
                if (/^private$/i.test((member as { visibility?: string }).visibility ?? '')) {
                    names.set(member.name.toLowerCase(), member.name);
                }
                break;

            case 'Declare':
                if (isPrivate((member as { modifiers?: string[] }).modifiers ?? [])) {
                    names.set(member.name.toLowerCase(), member.name);
                }
                break;

            default:
                break;
        }
    }

    return names;
}

/** Every identifier a stretch of source mentions, lowercased, comments and strings left out. */
function namesIn(text: string): Set<string> {
    const found = new Set<string>();
    for (const match of blanked(text).matchAll(/[\p{L}_][\p{L}\p{M}\p{N}_]*/gu)) {
        found.add(match[0].toLowerCase());
    }

    return found;
}

/** The text with comments and string literals blanked, so neither can look like a name. */
function blanked(text: string): string {
    let out = '';
    let inString = false;
    let inComment = false;

    for (const character of text) {
        if (character === '\n' || character === '\r') {
            inString = false;
            inComment = false;
            out += character;
            continue;
        }

        if (inComment || inString) {
            if (character === '"') {
                inString = false;
            }

            out += ' ';
            continue;
        }

        if (character === '"') {
            inString = true;
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

/**
 * Call sites that name the OLD module, rewritten to name the new one.
 *
 * `Helpers.Recalc` is the only shape a move breaks, because it says where the procedure lives and
 * that is what changed. An unqualified call is untouched: VBA resolves it across the project and
 * goes on finding it.
 */
function requalified(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    procedure: ProcedureNode,
    targetName: string,
    targetSource: string,
): { modules: { module: string; source: string }[]; count: number; refused?: string } {
    if (!procedure.nameSpan) {
        return { modules: [], count: 0 };
    }

    const uses = referencesFor(symbols, moduleName, source, procedure.nameSpan.start, false);
    const byModule = new Map<string, { line: number; column: number }[]>();

    for (const use of uses) {
        const held = symbols.byModule.get(use.module.toLowerCase());
        const text = use.module.toLowerCase() === moduleName.toLowerCase()
            ? source
            : use.module.toLowerCase() === targetName.toLowerCase()
                ? targetSource
                : held?.source;
        if (text === undefined) {
            continue;
        }

        const starts = lineStarts(text);
        const at = (starts[use.line - 1] as number) + use.column - 1;
        const before = text.slice(Math.max(0, at - moduleName.length - 8), at);
        if (!new RegExp(`\\b${moduleName.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')}\\s*\\.\\s*$`, 'iu').test(before)) {
            continue;
        }

        const rows = byModule.get(use.module) ?? [];
        rows.push({ line: use.line, column: use.column });
        byModule.set(use.module, rows);
    }

    const modules: { module: string; source: string }[] = [];
    let count = 0;

    for (const [named, rows] of byModule) {
        const held = symbols.byModule.get(named.toLowerCase());
        const text = named.toLowerCase() === moduleName.toLowerCase() ? source : held?.source;
        if (text === undefined) {
            continue;
        }

        const starts = lineStarts(text);
        const edits = rows
            .map((row) => (starts[row.line - 1] as number) + row.column - 1)
            .sort((a, b) => b - a);

        let rewritten = text;
        for (const at of edits) {
            // Back from the name, over the dot and any space, to the module's own name.
            const before = rewritten.slice(0, at);
            const qualifier = new RegExp(`\\b${moduleName.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&')}(\\s*\\.\\s*)$`, 'iu')
                .exec(before);
            if (!qualifier) {
                continue;
            }

            const from = before.length - qualifier[0].length;
            rewritten = rewritten.slice(0, from) + targetName + (qualifier[1] ?? '.') + rewritten.slice(at);
            count += 1;
        }

        modules.push({ module: named, source: rewritten });
    }

    return { modules, count };
}
