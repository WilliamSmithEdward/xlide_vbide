/*
 * Implement Interface: every member a class promised, written out for it.
 *
 * `Implements IStore` is a promise VBA checks at compile time and helps with not at all. The class
 * must carry a `Private Sub IStore_Save()` for every member of IStore, with a signature matching
 * the interface's exactly - get one parameter's ByVal wrong and it does not compile, and the
 * message names the member rather than the difference. It is boilerplate a person types by hand,
 * from another module, one member at a time, and the only reward for getting it right is silence.
 *
 * The analyzer already knows both halves: which interfaces a module declares, and what each one
 * declares. So this is a lookup and a rendering, and the rendering is the part that has to be
 * exact.
 *
 * THE PARAMETERS ARE COPIED, NOT REBUILT. Each one is reproduced from the interface's own source
 * text, span for span, because a reconstruction from the parsed fields is a second implementation
 * of VBA's parameter grammar - Optional, ByVal, ParamArray, arrays, defaults and their order - and
 * the one thing this feature must never do is emit a signature that does not match the one it
 * copied.
 *
 * The stubs raise. An empty stub compiles, does nothing, and returns whatever the type's zero is,
 * so a member nobody has written yet is indistinguishable from one that legitimately does nothing;
 * `Err.Raise 5` says so at the moment it matters. Measured against the analyzer: a stub whose only
 * statement raises reports nothing, not even the missing return assignment a Function would
 * otherwise be asked for.
 */

import { parseModule, type ParameterNode, type ProcedureNode, type ProcKind } from '../../../xlide_vscode/src/analyzer';
import type { ProjectSymbols } from './navigation';
import type { ImplementInterfaceResult } from './protocol';

/**
 * The types whose field becomes a Property Let rather than a Property Set.
 *
 * A public field in an interface is a property pair to whoever implements it, and VBA assigns an
 * object with Set and a value without it. Same list and same reason as Extract Method's: the
 * declared type is the only evidence, and `Variant` is absent because it can hold either - a Let
 * and a Set are both legal for one, and Let is the form that works for the values it usually holds.
 */
const VALUE_TYPES = new Set([
    'boolean', 'byte', 'currency', 'date', 'double', 'integer', 'long', 'longlong', 'longptr',
    'single', 'string', 'variant',
]);

/** One member an interface requires of whoever implements it. */
interface Required {
    /** The member's own name, as the interface spells it. */
    name: string;
    kind: ProcKind;
    /** The text between the parentheses, copied from the interface. */
    parameters: string;
    /** The `As` clause, without the keyword; absent on a Sub and a Property Let or Set. */
    returns?: string;
}

export function implementInterfaceFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    interfaceName: string | undefined,
): ImplementInterfaceResult {
    const refuse = (why: string): ImplementInterfaceResult => ({ refused: why });

    // The analyzer's own scan rather than another regex here, so the list this works from is the
    // list the rest of the product believes - including the qualified `Project.IStore` form.
    const declared = symbols.project.moduleImplementsList(moduleName);
    if (declared.length === 0) {
        return refuse(`'${moduleName}' does not declare Implements, so there is no interface to implement. Add 'Implements <name>' below Option Explicit first.`);
    }

    const wanted = interfaceName
        ? declared.filter((one) => bare(one).toLowerCase() === bare(interfaceName).toLowerCase())
        : declared;

    if (wanted.length === 0) {
        return refuse(`'${moduleName}' does not implement '${interfaceName}'. It implements ${list(declared)}.`);
    }

    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const mine = parseModule(source);

    // KEYED BY NAME AND KIND, because a Property Get and a Property Let share a name and are two
    // different members. Keyed by name alone, the Get was written and the Let was skipped as
    // already present - so every property in the interface, and every public field, came out half
    // implemented and the class still would not compile.
    const already = new Set(mine.members
        .filter((member): member is ProcedureNode => member.kind === 'Procedure')
        .map((member) => memberKey(member.name, member.procKind)));

    const written: string[] = [];
    const blocks: string[] = [];

    for (const each of wanted) {
        const name = bare(each);
        const held = symbols.byModule.get(name.toLowerCase());
        if (!held) {
            return refuse(`This project has no module called '${name}', so what it requires cannot be read. Check the name on the Implements line.`);
        }

        const members = requiredOf(held.source ?? '');
        if (members.length === 0) {
            return refuse(`'${name}' declares no public members, so there is nothing to implement.`);
        }

        for (const member of members) {
            const stub = `${name}_${member.name}`;
            const key = memberKey(stub, member.kind);
            if (already.has(key)) {
                continue;
            }

            already.add(key);
            written.push(stub);
            blocks.push(render(stub, member, eol));
        }
    }

    if (written.length === 0) {
        return refuse(`'${moduleName}' already implements every member of ${list(wanted.map(bare))}.`);
    }

    // At the end, where a reader looks for a class's implementation of somebody else's interface,
    // and where nothing already written has to move.
    const tail = source.endsWith(eol) ? '' : eol;
    const rewritten = `${source}${tail}${eol}${blocks.join(eol + eol)}${eol}`;

    return {
        module: moduleName,
        source: rewritten,
        interfaces: wanted.map(bare),
        added: written,
    };
}

/**
 * What makes one member distinct from another: its name AND which of the five it is. A colon,
 * because a VBA identifier cannot hold one, so the two halves cannot run together.
 */
function memberKey(name: string, kind: ProcKind): string {
    return `${name.toLowerCase()}:${kind}`;
}

/** `IStore` out of `IStore` or `ThisProject.IStore`. */
function bare(name: string): string {
    const dot = name.lastIndexOf('.');
    return dot < 0 ? name : name.slice(dot + 1);
}

/** "IStore", or "IStore and ILog", or "IStore, ILog and IClock". */
function list(names: readonly string[]): string {
    if (names.length <= 1) {
        return names.length === 0 ? 'nothing' : `'${names[0]}'`;
    }

    const quoted = names.map((one) => `'${one}'`);
    return `${quoted.slice(0, -1).join(', ')} and ${quoted[quoted.length - 1]}`;
}

/**
 * What one interface requires: its public procedures, and a property pair for each public field.
 *
 * A field is a pair because that is what it is to an implementer - VBA offers no way to implement
 * a variable, and asks for the Get and the Let (or Set) the variable would have behaved as.
 */
function requiredOf(source: string): Required[] {
    const required: Required[] = [];

    for (const member of parseModule(source).members) {
        if (member.kind === 'Procedure') {
            if (isPrivate(member.modifiers)) {
                continue;
            }

            required.push({
                name: member.name,
                kind: member.procKind,
                parameters: parametersOf(member, source),
                ...(member.returnType ? { returns: member.returnType } : {}),
            });
            continue;
        }

        if (member.kind !== 'VariableGroup' || member.isConst || isPrivate([member.modifier])) {
            continue;
        }

        for (const declared of member.declarations) {
            const type = declared.asType ?? 'Variant';
            const array = declared.isArray ? '()' : '';
            required.push({
                name: declared.name,
                kind: 'PropertyGet',
                parameters: '',
                returns: `${type}${array}`,
            });
            required.push({
                name: declared.name,
                kind: VALUE_TYPES.has(type.toLowerCase()) && !declared.isArray ? 'PropertyLet' : 'PropertySet',
                // RHS is VBA's own name for the value on the right of the assignment, and what the
                // editor writes when it generates a property itself.
                parameters: `ByVal RHS As ${type}${array}`,
            });
        }
    }

    return required;
}

/** `Private` on a member means the interface does not require it. Anything else does. */
function isPrivate(modifiers: readonly (string | undefined)[]): boolean {
    return modifiers.some((one) => /^private$/i.test(one ?? ''));
}

/**
 * The interface's own parameter text, copied span for span.
 *
 * Rebuilding it from the parsed fields would be a second implementation of VBA's parameter
 * grammar, and a signature that does not match the interface's is the one failure this feature
 * cannot be allowed to have: it does not compile, and the message names the member rather than
 * the difference.
 */
function parametersOf(procedure: ProcedureNode, source: string): string {
    return procedure.params
        .map((parameter: ParameterNode) => source.slice(parameter.span.start, parameter.span.end).trim())
        .join(', ');
}

/** One stub: the header the interface asks for, a raise, and the closer. */
function render(name: string, member: Required, eol: string): string {
    const opens = keywordOf(member.kind);
    const returns = member.returns && (member.kind === 'Function' || member.kind === 'PropertyGet')
        ? ` As ${member.returns}`
        : '';

    return [
        `Private ${opens} ${name}(${member.parameters})${returns}`,
        // Error 5 is VBA's "invalid procedure call or argument", which is the nearest thing the
        // language has to "not implemented" and what the editor's own generated members raise.
        `    Err.Raise 5, "${name}", "${name} is not implemented yet."`,
        `End ${closerOf(member.kind)}`,
    ].join(eol);
}

function keywordOf(kind: ProcKind): string {
    switch (kind) {
        case 'Sub': return 'Sub';
        case 'Function': return 'Function';
        case 'PropertyGet': return 'Property Get';
        case 'PropertyLet': return 'Property Let';
        case 'PropertySet': return 'Property Set';
    }
}

function closerOf(kind: ProcKind): string {
    return kind === 'Sub' ? 'Sub' : kind === 'Function' ? 'Function' : 'Property';
}
