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
 * A VBA identifier: a letter, then letters, digits and underscores, up to 255 characters. Checked
 * here rather than on the surface because a rename that produces something VBA cannot parse turns
 * one compiling project into a broken one across several modules at once, and the surface has no
 * business knowing the language's rules.
 */
const IDENTIFIER = /^[A-Za-z][A-Za-z0-9_]{0,254}$/;

/**
 * The reserved words a rename must not produce. Not the full keyword list — a name that merely
 * collides with a statement keyword is caught by the analyzer's own diagnostics on the next pass,
 * and refusing every one of them here would reject names VBA accepts.
 */
const RESERVED = new Set([
    'as', 'byref', 'byval', 'call', 'case', 'const', 'dim', 'do', 'each', 'else', 'elseif', 'end',
    'error', 'exit', 'false', 'for', 'function', 'get', 'goto', 'if', 'in', 'is', 'let', 'loop',
    'me', 'new', 'next', 'nothing', 'null', 'on', 'option', 'private', 'property', 'public',
    'redim', 'resume', 'return', 'select', 'set', 'sub', 'then', 'to', 'true', 'until', 'wend',
    'while', 'with',
]);

/**
 * Every module a rename rewrites, with its new text.
 *
 * Whole texts rather than edit lists: the add-in writes modules, and a module with no tab open
 * has no editor to apply edits to. The engine already holds every module's current text, so
 * producing the result costs nothing and leaves no arithmetic for two sides to get differently.
 */
export function renameFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
    newName: string,
): { modules: { module: string; source: string; replaced: number }[]; oldName?: string; refused?: string; ambiguous?: LocationPayload[] } {
    const word = identifierAt(source, offset);
    if (!word) {
        return { modules: [], refused: 'There is no symbol here to rename.' };
    }

    if (!IDENTIFIER.test(newName)) {
        return {
            modules: [],
            oldName: word.text,
            refused: `'${newName}' is not a VBA name. A name starts with a letter and holds letters, digits and underscores.`,
        };
    }

    if (RESERVED.has(newName.toLowerCase())) {
        return { modules: [], oldName: word.text, refused: `'${newName}' is a VBA keyword.` };
    }

    if (newName.toLowerCase() === word.text.toLowerCase()) {
        // Same name in different case is a recasing, which is a real edit and is allowed. The
        // same name in the same case is not.
        if (newName === word.text) {
            return { modules: [], oldName: word.text, refused: 'That is already its name.' };
        }
    }

    // Anchored at the DECLARATION, wherever the developer started from.
    //
    // Starting from a call site and starting from the declaration must rename the same set, and
    // they did not: asked at `CleanModule.RunTotal`, the member resolver answered with every
    // module declaring a RunTotal, and the rename went through all of them — including a
    // BrokenModule.RunTotal that has nothing to do with it (the developer, 2026-08-06). Resolving
    // to the one declaration first and collecting from there makes every entry point agree, and
    // makes the entry point that was already correct the only one there is.
    const anchor = anchorFor(symbols, moduleName, source, offset);
    if (anchor.ambiguous) {
        return { modules: [], oldName: word.text, refused: anchor.ambiguous };
    }

    const sites = anchor.at
        ? referencesFor(symbols, anchor.at.module, anchor.at.source, anchor.at.offset, true)
        : referencesFor(symbols, moduleName, source, offset, true);

    if (sites.length === 0) {
        return {
            modules: [],
            oldName: word.text,
            refused: `'${word.text}' could not be resolved, so nothing was renamed.`,
        };
    }

    const byModule = new Map<string, LocationPayload[]>();
    for (const site of sites) {
        const key = site.module.toLowerCase();
        const held = byModule.get(key);
        if (held) {
            held.push(site);
        } else {
            byModule.set(key, [site]);
        }
    }

    const out: { module: string; source: string; replaced: number }[] = [];

    for (const [key, locations] of byModule) {
        const module = symbols.byModule.get(key);
        if (!module) {
            // A site in a module the assembly does not hold cannot be rewritten, and a rename
            // that silently skips a module is the failure this whole feature exists to avoid.
            return {
                modules: [],
                oldName: word.text,
                refused: `'${locations[0].module}' could not be read, so nothing was renamed.`,
            };
        }

        const rewritten = replaceAll(module.source, locations, newName);
        if (rewritten === null) {
            return {
                modules: [],
                oldName: word.text,
                refused: `'${module.moduleName}' changed while renaming, so nothing was renamed.`,
            };
        }

        out.push({ module: module.moduleName, source: rewritten, replaced: locations.length });
    }

    return { modules: out, oldName: word.text, ambiguous: leftAlone(symbols, word.text, sites) };
}

/**
 * The declaration a rename should be anchored at, whichever site the developer asked from.
 *
 * One declaration is the answer. None means the symbol is a local, a parameter or something else
 * with no project-level declaration, and the caret's own position is the right anchor. More than
 * one means the name genuinely resolves to several procedures and nothing can prove which was
 * meant — which is the collision the developer asked to be warned about rather than guessed at.
 */
function anchorFor(
    symbols: ProjectSymbols,
    moduleName: string,
    source: string,
    offset: number,
): { at?: { module: string; source: string; offset: number }; ambiguous?: string } {
    const declarations = definitionsFor(symbols, moduleName, source, offset);

    if (declarations.length > 1) {
        const where = [...new Set(declarations.map((one) => one.module))].sort();
        return {
            ambiguous: where.length > 1
                ? `'${where.join("' and '")}' each declare this name, so nothing can tell which was meant. `
                    + 'Rename it from the declaration you mean.'
                : 'This name is declared more than once, so nothing can tell which one was meant.',
        };
    }

    const only = declarations[0];
    if (!only) {
        return {};
    }

    const home = symbols.byModule.get(only.module.toLowerCase());
    if (!home) {
        return {};
    }

    const starts = lineStarts(home.source);
    const line = starts[only.line - 1];
    if (line === undefined) {
        return {};
    }

    return { at: { module: home.moduleName, source: home.source, offset: line + only.column - 1 } };
}

/**
 * Every occurrence of the old name the rename did NOT touch.
 *
 * A bare call is only ambiguous when more than one module declares the name: with a single
 * definition in the workbook there is nothing else it could mean, and the resolver renames it.
 * When there IS a collision the resolver leaves it alone, which is right — nothing can prove
 * which one was meant — but leaving it alone SILENTLY is not: the developer is the only one who
 * knows, and they cannot decide about a call they were never told about.
 *
 * So what is left behind is counted and handed back. Warning about it is the surface's job.
 */
function leftAlone(
    symbols: ProjectSymbols,
    oldName: string,
    renamed: readonly LocationPayload[],
): LocationPayload[] {
    const touched = new Set(renamed.map((where) => `${where.module.toLowerCase()}:${where.line}:${where.column}`));
    const out: LocationPayload[] = [];

    for (const location of occurrencesOfTypeName(symbols, oldName)) {
        if (!touched.has(`${location.module.toLowerCase()}:${location.line}:${location.column}`)) {
            out.push(location);
        }
    }

    return out;
}

/**
 * One module's text with every named span replaced. Applied back to front so that an earlier
 * replacement cannot move the span of a later one, and refused outright if any span does not
 * still hold the name — a rename computed against text that has since moved must not be applied
 * to it by arithmetic that no longer describes it.
 */
function replaceAll(
    source: string,
    locations: readonly LocationPayload[],
    newName: string,
): string | null {
    const starts = lineStarts(source);

    const offsets = locations
        .map((where) => {
            const line = starts[where.line - 1];
            return line === undefined ? null : { start: line + where.column - 1, length: where.length };
        })
        .filter((span): span is { start: number; length: number } => span !== null);

    if (offsets.length !== locations.length) {
        return null;
    }

    offsets.sort((left, right) => right.start - left.start);

    let text = source;
    for (const span of offsets) {
        if (span.start < 0 || span.start + span.length > text.length) {
            return null;
        }
        text = text.slice(0, span.start) + newName + text.slice(span.start + span.length);
    }

    return text;
}

/** Where each line begins, counting the three line endings VBA modules turn up with. */
function lineStarts(source: string): number[] {
    const starts = [0];
    for (let at = 0; at < source.length; at++) {
        const character = source[at];
        if (character === '\n') {
            starts.push(at + 1);
        } else if (character === '\r') {
            starts.push(source[at + 1] === '\n' ? at + 2 : at + 1);
            if (source[at + 1] === '\n') {
                at++;
            }
        }
    }
    return starts;
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
