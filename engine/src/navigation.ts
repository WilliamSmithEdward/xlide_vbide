// Answers "where is this declared" and "where else is this used", across the modules of one
// workbook.
//
// The workbook is the boundary, not a setting. Two open workbooks can each hold a Module1 and a
// Recalculate, and they are unrelated: an answer that crossed from one to the other would send a
// developer to code that has nothing to do with what they clicked, and — once rename is built on
// these same spans — would edit it. Everything here is addressed by (project, module) already, so
// the boundary is the address rather than a rule applied on top of one.
//
// The resolution itself is the extension's, unchanged: a shared, vscode-free resolver that is
// scope- and shadow-aware, understands member access through a receiver, and knows that a bare
// call to an object module's member only binds inside that module.

import { ProjectIndex, resolveProcedureLabelDefinitionAt, resolveTypeReferenceAt } from '../../../xlide_vscode/src/analyzer';
import {
    collectSymbolReferences,
    projectClassMemberAtDefinition,
    sourceMemberDefinitionsAt,
} from '../../../xlide_vscode/src/vbaReferenceResolution';
import type { VbaModuleSymbols } from '../../../xlide_vscode/src/vbaSymbolIndex';
import { buildLiveVbaProjectIndex } from '../../../xlide_vscode/src/vbaProjectAnalysis';
import type { LocationPayload, ModulePayload } from './protocol';

/** One workbook as the resolvers want it: the index, the module list, and the map into it. */
export interface ProjectSymbols {
    project: ProjectIndex;
    modules: VbaModuleSymbols[];
    byModule: Map<string, VbaModuleSymbols>;
}

/**
 * Assembles a workbook's symbols from the text the engine holds — live where a module is being
 * typed in, seeded elsewhere.
 *
 * Live rather than seeded, deliberately and unlike the completion context: an answer here is a
 * line number the surface will scroll to, and the surface shows the live text. Answering from the
 * text as of the last write-back would land the cursor a few lines out in any module with unsaved
 * edits, which is worse than not answering.
 */
export function assembleSymbols(
    modules: readonly { moduleName: string; source: string; type?: string; documentType?: string }[],
): ProjectSymbols {
    const entries: VbaModuleSymbols[] = modules.map((module) => ({
        moduleName: module.moduleName,
        source: module.source,
        type: module.type,
        documentType: module.documentType as VbaModuleSymbols['documentType'],
    }));

    const project = buildLiveVbaProjectIndex(entries.map((entry) => ({
        moduleName: entry.moduleName,
        source: entry.source,
        type: entry.type,
        documentType: entry.documentType,
    })));

    const byModule = new Map<string, VbaModuleSymbols>();
    for (const entry of entries) {
        byModule.set(entry.moduleName.toLowerCase(), entry);
    }

    return { project, modules: entries, byModule };
}

/** Where the identifier at an offset is declared. Empty when nothing there resolves. */
export function definitionsFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
): LocationPayload[] {
    const { project, modules, byModule } = symbols;
    const current = byModule.get(moduleName.toLowerCase());

    // A GoTo label is resolved inside its own procedure and goes no further.
    const label = resolveProcedureLabelDefinitionAt(source, offset);
    if (label) {
        return [locationOf(byModule, moduleName, label.label.span)].filter(isLocation);
    }

    const word = identifierAt(source, offset);
    if (!word) {
        return [];
    }

    // A member reached through a receiver — Sheet1.Refresh, Me.Total, obj.Item — resolves to the
    // member's own declaration wherever it lives, which is what makes this cross-module at all.
    const members = sourceMemberDefinitionsAt(
        source,
        word.text,
        word.end,
        project,
        modules,
        moduleName,
        current?.type,
        current?.documentType,
    );

    if (members.length > 0) {
        return members
            .map((definition) => locationOf(byModule, definition.moduleName, definition.nameSpan))
            .filter(isLocation);
    }

    // `As Widget` and `New Widget` name a type rather than a value, and the type's declaration is
    // where the developer means to go. Only without a qualifier: a qualified name is a member
    // reference, which the branch above owns.
    const qualifier = qualifierBefore(source, word.start);
    if (!qualifier) {
        const reference = resolveTypeReferenceAt(source, offset, {
            projectTypes: project.visibleTypeNames(moduleName),
        });

        if (reference) {
            const found = project
                .resolveTypeDefinitions(moduleName, reference.name)
                .filter((definition) => !reference.qualifier
                    || definition.moduleName.toLowerCase() === reference.qualifier.toLowerCase())
                .map((definition) => definition.nameSpan
                    ? locationOf(byModule, definition.moduleName, definition.nameSpan)
                    : null)
                .filter(isLocation);

            if (found.length > 0) {
                return found;
            }
        }
    }

    const resolved = qualifier
        ? project.resolveQualifiedDefinition(qualifier, word.text)
        : project.resolveDefinition(moduleName, word.text, offset);

    return resolved
        .map((symbol) => locationOf(byModule, symbol.moduleName, symbol.nameSpan))
        .filter(isLocation);
}

/** Every use of the identifier at an offset, across the workbook's modules. */
export function referencesFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
    includeDeclaration: boolean,
): LocationPayload[] {
    const { project, modules, byModule } = symbols;
    const current = byModule.get(moduleName.toLowerCase());

    const word = identifierAt(source, offset);
    if (!word) {
        return [];
    }

    // A type reference takes precedence over bare value resolution, but only when the cursor is
    // not on a member — the unified resolver owns those, and it knows more than this does.
    const onMember = sourceMemberDefinitionsAt(
        source,
        word.text,
        word.end,
        project,
        modules,
        moduleName,
        current?.type,
        current?.documentType,
    ).length > 0
        || (projectClassMemberAtDefinition(project, moduleName, word.text, offset)?.definitions?.length ?? 0) > 0;

    if (!onMember) {
        const reference = resolveTypeReferenceAt(source, offset, {
            projectTypes: project.visibleTypeNames(moduleName),
        });

        if (reference && project.resolveTypeDefinitions(moduleName, reference.name).length > 0) {
            return occurrencesOfTypeName(symbols, reference.name);
        }
    }

    const found = collectSymbolReferences(
        byModule,
        project,
        modules,
        source,
        moduleName,
        current,
        word.text,
        word.end,
        offset,
        includeDeclaration,
    );

    if (!found.hasSymbol) {
        return [];
    }

    return found.references.map((span) => ({
        module: canonicalName(byModule, span.moduleName),
        line: span.line + 1,
        column: span.column + 1,
        length: span.length,
    }));
}

/**
 * Every mention of a type's name across the workbook. A type is named in declarations rather
 * than called, so the scope resolver has nothing to bind; matching the identifier is what the
 * extension does here too.
 */
function occurrencesOfTypeName(symbols: ProjectSymbols, name: string): LocationPayload[] {
    const wanted = name.toLowerCase();
    const out: LocationPayload[] = [];

    for (const module of symbols.modules) {
        const lines = module.source.split(/\r\n|\n|\r/);
        for (let index = 0; index < lines.length; index++) {
            for (const at of identifierPositions(lines[index], wanted)) {
                out.push({ module: module.moduleName, line: index + 1, column: at + 1, length: name.length });
            }
        }
    }

    return out;
}

function* identifierPositions(line: string, wanted: string): Generator<number> {
    const haystack = line.toLowerCase();
    let from = 0;

    for (let at = haystack.indexOf(wanted, from); at >= 0; at = haystack.indexOf(wanted, from)) {
        from = at + 1;
        if (!isIdentifierChar(haystack[at - 1]) && !isIdentifierChar(haystack[at + wanted.length])) {
            yield at;
        }
    }
}

/** The identifier the offset is inside or immediately after, or nothing. */
function identifierAt(source: string, offset: number): { text: string; start: number; end: number } | null {
    let start = Math.min(Math.max(offset, 0), source.length);
    let end = start;

    while (start > 0 && isIdentifierChar(source[start - 1])) {
        start--;
    }
    while (end < source.length && isIdentifierChar(source[end])) {
        end++;
    }

    return end > start ? { text: source.slice(start, end), start, end } : null;
}

/**
 * The receiver an identifier is reached through, when it is written as one word and a dot. A
 * quoted module name — `'Sheet One'.Refresh` — is a name with spaces in it, which is why the
 * quotes are stripped rather than treated as the start of a string.
 */
function qualifierBefore(source: string, start: number): string | undefined {
    let at = start;
    while (at > 0 && (source[at - 1] === ' ' || source[at - 1] === '\t')) {
        at--;
    }

    if (at === 0 || source[at - 1] !== '.') {
        return undefined;
    }

    at--;
    while (at > 0 && (source[at - 1] === ' ' || source[at - 1] === '\t')) {
        at--;
    }

    if (source[at - 1] === "'") {
        const close = at - 1;
        const open = source.lastIndexOf("'", close - 1);
        return open >= 0 ? source.slice(open + 1, close) : undefined;
    }

    let from = at;
    while (from > 0 && isIdentifierChar(source[from - 1])) {
        from--;
    }

    return from < at ? source.slice(from, at) : undefined;
}

function isIdentifierChar(character: string | undefined): boolean {
    return character !== undefined && /[A-Za-z0-9_]/.test(character);
}

function locationOf(
    byModule: Map<string, VbaModuleSymbols>,
    moduleName: string,
    span: { start: number; end: number },
): LocationPayload | null {
    const module = byModule.get(moduleName.toLowerCase());
    if (!module) {
        return null;
    }

    const before = module.source.slice(0, span.start);
    const lineBreaks = before.split(/\r\n|\n|\r/);

    return {
        module: module.moduleName,
        line: lineBreaks.length,
        column: (lineBreaks[lineBreaks.length - 1]?.length ?? 0) + 1,
        length: span.end - span.start,
    };
}

function canonicalName(byModule: Map<string, VbaModuleSymbols>, moduleName: string): string {
    return byModule.get(moduleName.toLowerCase())?.moduleName ?? moduleName;
}

function isLocation(value: LocationPayload | null): value is LocationPayload {
    return value !== null;
}
