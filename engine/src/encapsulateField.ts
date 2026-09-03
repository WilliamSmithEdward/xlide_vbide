/*
 * Encapsulate Field: a public variable becomes a private one behind a property pair.
 *
 * `Public Name As String` in a class is a variable anyone can write, cannot be made read-only,
 * cannot validate what it is given, and cannot carry a breakpoint on the write that broke it. The
 * fix is eight lines of the most mechanical code VBA has, and it is the reason the field is still
 * public in most projects.
 *
 * NOTHING THAT USES IT HAS TO CHANGE, which is what makes this one safe. The property keeps the
 * field's name, so `store.Name = "x"` now calls `Property Let Name` and `Debug.Print store.Name`
 * calls `Property Get Name`, in this module and every other one, with no call site rewritten. The
 * variable moves out from under the name and the name goes on meaning what it meant.
 *
 * No type inference anywhere: the type is declared on the line being encapsulated, which is what
 * separates this from Extract Variable (xlide_vscode#61) and is why it could be built first.
 */

import { parseModule, type ProcedureNode, type VariableDeclNode, type VariableGroupNode } from '../../../xlide_vscode/src/analyzer';
import { lineStarts, toLineColumn } from './navigation';
import type { EncapsulateFieldResult } from './protocol';

/**
 * The types whose property pair is a Get and a Let rather than a Get and a Set.
 *
 * Same list and same reason as Implement Interface's: VBA assigns an object with Set and a value
 * without it, and the declared type is the evidence. `Variant` sits with the values because that
 * is how one is nearly always used, and a Variant field that holds an object is a case this
 * refuses rather than guesses at - see the refusal below.
 */
const VALUE_TYPES = new Set([
    'boolean', 'byte', 'currency', 'date', 'double', 'integer', 'long', 'longlong', 'longptr',
    'single', 'string', 'variant',
]);

export function encapsulateFieldFor(
    moduleName: string,
    source: string,
    fieldName: string,
): EncapsulateFieldResult {
    const refuse = (why: string): EncapsulateFieldResult => ({ refused: why });

    const module = parseModule(source);
    let group: VariableGroupNode | undefined;
    let field: VariableDeclNode | undefined;

    for (const member of module.members) {
        if (member.kind !== 'VariableGroup') {
            continue;
        }

        const found = member.declarations.find(
            (one) => one.name.toLowerCase() === fieldName.toLowerCase());
        if (found) {
            group = member;
            field = found;
            break;
        }
    }

    if (!group || !field) {
        return refuse(`'${moduleName}' declares no module-level variable called '${fieldName}'.`);
    }

    if (group.isConst) {
        return refuse(`'${field.name}' is a Const, which is already read-only and cannot be assigned through a property.`);
    }

    if (/^private$/i.test(group.modifier ?? '')) {
        return refuse(`'${field.name}' is already Private, so nothing outside this module can reach it.`);
    }

    if (group.withEvents) {
        return refuse(`'${field.name}' is declared WithEvents, and its event handlers are bound to the variable rather than to a property. This one has to be done by hand.`);
    }

    if (field.isArray) {
        return refuse(`'${field.name}' is an array, and VBA cannot pass one to a Property Let. Expose it through a method instead, or leave it as it is.`);
    }

    // A DECLARATION GROUP OF ONE. `Public a As Long, b As Long` would have to be split, and
    // splitting a line the developer wrote to keep two related names together is a change they
    // did not ask for. Named rather than silently taking the whole line with it.
    if (group.declarations.length > 1) {
        const others = group.declarations
            .filter((one) => one !== field)
            .map((one) => `'${one.name}'`)
            .join(', ');
        return refuse(`'${field.name}' shares its declaration with ${others}. Give it a line of its own first.`);
    }

    const type = field.asType ?? typeOfSuffix(field.typeSuffix);
    const backing = `m_${field.name}`;

    const declared = new Set<string>();
    for (const member of module.members) {
        if (member.kind === 'Procedure') {
            declared.add(member.name.toLowerCase());
        } else if (member.kind === 'VariableGroup') {
            for (const one of member.declarations) {
                declared.add(one.name.toLowerCase());
            }
        }
    }

    if (declared.has(backing.toLowerCase())) {
        return refuse(`'${moduleName}' already declares '${backing}', which is the name the private variable would take.`);
    }

    // A property CANNOT share a name with a procedure, and one that already exists means somebody
    // has started this by hand.
    const clash = module.members.find((member): member is ProcedureNode => member.kind === 'Procedure'
        && member.name.toLowerCase() === field.name.toLowerCase());
    if (clash) {
        return refuse(`'${moduleName}' already declares a ${clash.procKind === 'Sub' ? 'Sub' : 'procedure'} called '${field.name}'.`);
    }

    const eol = source.includes('\r\n') ? '\r\n' : '\n';
    const starts = lineStarts(source);
    const first = toLineColumn(starts, group.span.start).line;
    const last = toLineColumn(starts, Math.max(group.span.start, group.span.end - 1)).line;
    if (first !== last) {
        return refuse(`'${field.name}' is declared over more than one line, which this cannot rewrite safely. Put the declaration on one line first.`);
    }

    const indent = /^[ \t]*/.exec(source.slice(starts[first - 1] as number, starts[first - 1] as number + 40))?.[0] ?? '';
    const writes = VALUE_TYPES.has(type.toLowerCase()) ? 'Let' : 'Set';
    const assigns = writes === 'Let' ? '' : 'Set ';

    // The visibility the field had, so a Friend field becomes a Friend property. Public when the
    // declaration said nothing, which is what a bare `Dim` at module level means.
    const visibility = /^(public|friend)$/i.test(group.modifier ?? '')
        ? (group.modifier as string).replace(/^./, (one) => one.toUpperCase())
        : 'Public';

    const property = [
        `${visibility} Property Get ${field.name}() As ${type}`,
        `${indent}    ${assigns}${field.name} = ${backing}`,
        'End Property',
        '',
        `${visibility} Property ${writes} ${field.name}(ByVal RHS As ${type})`,
        `${indent}    ${assigns}${backing} = RHS`,
        'End Property',
    ].join(eol);

    // The declaration in place, so it stays in the declarations section VBA insists it lives in,
    // and the properties at the end where a reader looks for a class's members.
    const declarationFrom = starts[first - 1] as number;
    const declarationTo = first < starts.length ? (starts[first] as number) : source.length;
    const rewritten = source.slice(0, declarationFrom)
        + `${indent}Private ${backing} As ${type}${eol}`
        + source.slice(declarationTo);

    const tail = rewritten.endsWith(eol) ? '' : eol;
    return {
        module: moduleName,
        source: `${rewritten}${tail}${eol}${property}${eol}`,
        field: field.name,
        backingField: backing,
        accessors: [`Property Get ${field.name}`, `Property ${writes} ${field.name}`],
    };
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
