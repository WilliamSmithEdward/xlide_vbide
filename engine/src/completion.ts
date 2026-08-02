// Serves completions by reusing the editor extension's analyzer resolvers.
//
// The resolvers are pure: source text in, plain completion data out, with project facts passed as
// data. Everything here is the assembly of those facts from the modules the add-in has seeded,
// which is the same assembly the extension performs from its open workbook, minus the editor.
// Reusing the resolvers keeps one completion behaviour shared between the two products instead of
// two that drift.

import {
    eventHandlerDocumentTypeForContext,
    resolveIdentifierCompletions,
    resolveKeywordCompletions,
    resolveMemberCompletions,
    type EventHandlerDocumentType,
    type IdentifierCompletion,
    type IdentifierCompletionContext,
    type KeywordCompletion,
    type MemberCompletion,
    type MemberCompletionContext,
} from '../../../xlide_vscode/src/analyzer';
import {
    buildLiveVbaProjectIndex,
    moduleKindFromType,
    projectEditorSymbolContextForModule,
    type VbaProjectModuleInput,
} from '../../../xlide_vscode/src/vbaProjectAnalysis';
import type { CompletionItemPayload, CompletionParams, ModulePayload } from './protocol';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const CHART = 'Excel.Chart';

interface ModuleEntry {
    name: string;
    type: string;
    documentType?: EventHandlerDocumentType;
    source: string;
}

/** Maps a document module to the host type that `Me` denotes inside it. */
function meTypeFor(entry: ModuleEntry | undefined): string | undefined {
    if (!entry || entry.type !== 'document') {
        return undefined;
    }
    switch (documentTypeFor(entry)) {
        case 'workbook':
            return WORKBOOK;
        case 'chart':
            return CHART;
        default:
            return WORKSHEET;
    }
}

function meProjectTypeFor(entry: ModuleEntry | undefined): string | undefined {
    if (!entry || !['class', 'document', 'userform'].includes(entry.type)) {
        return undefined;
    }
    return entry.name;
}

function documentTypeFor(entry: ModuleEntry | undefined): EventHandlerDocumentType | undefined {
    if (!entry || entry.type !== 'document') {
        return undefined;
    }
    return eventHandlerDocumentTypeForContext({
        moduleName: entry.name,
        moduleKind: 'document',
        documentType: entry.documentType,
    });
}

/** Lowercased document code name -> qualified host type, for receiver resolution. */
function codeNamesFor(entries: ModuleEntry[]): Record<string, string> {
    const out: Record<string, string> = {};
    for (const entry of entries) {
        if (entry.type === 'document') {
            out[entry.name.toLowerCase()] = meTypeFor(entry) ?? WORKSHEET;
        }
    }
    return out;
}

export function completionsFor(
    seeded: readonly ModulePayload[],
    params: CompletionParams,
): CompletionItemPayload[] {
    // The live source replaces the seeded copy of the module being typed in; every other module
    // keeps its seeded text, which carries the project facts.
    const entries: ModuleEntry[] = seeded
        .filter((module) => module.moduleName.toLowerCase() !== params.moduleName.toLowerCase())
        .map((module) => ({
            name: module.moduleName,
            type: module.type ?? 'standard',
            documentType: module.documentType as EventHandlerDocumentType | undefined,
            source: module.source,
        }));

    const current: ModuleEntry = {
        name: params.moduleName,
        type: params.moduleType ?? seededTypeOf(seeded, params.moduleName) ?? 'standard',
        documentType: (params.documentType ?? seededDocumentTypeOf(seeded, params.moduleName)) as
            | EventHandlerDocumentType
            | undefined,
        source: params.source,
    };
    entries.push(current);

    const moduleKind = moduleKindFromType(current.type);

    // Project facts. A project that will not index still completes: members of host receivers
    // and keywords need no index at all.
    let projectClassMembers: MemberCompletionContext['projectClassMembers'];
    let projectProcedures: IdentifierCompletionContext['projectProcedures'];
    let projectSymbols: IdentifierCompletionContext['projectSymbols'];

    try {
        const inputs: VbaProjectModuleInput[] = entries.map((entry) => ({
            moduleName: entry.name,
            source: entry.source,
            type: entry.type,
            documentType: entry.documentType,
        }));

        const project = buildLiveVbaProjectIndex(inputs, {
            moduleName: current.name,
            moduleKind,
            source: current.source,
        });

        const context = projectEditorSymbolContextForModule(project, current.name);
        projectClassMembers = context.analysisOptions.projectClassMembers;
        projectProcedures = context.externalProjectProcedures;
        projectSymbols = context.externalProjectSymbols;
    } catch {
        // Conservative surfaces beat none.
    }

    // After a dot, the members of the receiver are the only sensible answer.
    const members = resolveMemberCompletions(params.source, params.offset, {
        codeNames: codeNamesFor(entries),
        meType: meTypeFor(current),
        meProjectType: meProjectTypeFor(current),
        projectClassMembers,
    });

    if (members.length > 0) {
        return members.map(memberItem);
    }

    // A grammar position that admits only keywords shows only keywords; anywhere else, the
    // identifiers in scope and the keywords share the list, the way the extension offers them.
    const keywords = resolveKeywordCompletions(params.source, params.offset);
    if (keywords.exclusive) {
        return keywords.items.map(keywordItem);
    }

    const identifiers = resolveIdentifierCompletions(params.source, params.offset, {
        codeNames: entries.filter((entry) => entry.type === 'document').map((entry) => entry.name),
        moduleName: current.name,
        moduleKind,
        projectMemberSurfaces: projectClassMembers,
        projectProcedures,
        projectSymbols,
    });

    return [...identifiers.map(identifierItem), ...keywords.items.map(keywordItem)];
}

function seededTypeOf(seeded: readonly ModulePayload[], moduleName: string): string | undefined {
    return seeded.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.type;
}

function seededDocumentTypeOf(seeded: readonly ModulePayload[], moduleName: string): string | undefined {
    return seeded.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.documentType;
}

function memberItem(member: MemberCompletion): CompletionItemPayload {
    return {
        label: member.name,
        kind: member.kind,
        detail: member.signature ?? (member.returns ? `${member.owner}.${member.name} As ${member.returns}` : member.owner),
        documentation: member.documentation,
    };
}

function identifierItem(identifier: IdentifierCompletion): CompletionItemPayload {
    return {
        label: identifier.name,
        kind: identifier.kind,
        detail: identifier.detail,
        documentation: identifier.documentation,
    };
}

function keywordItem(keyword: KeywordCompletion): CompletionItemPayload {
    return {
        label: keyword.label,
        kind: 'keyword',
        detail: keyword.detail,
        documentation: keyword.documentation,
        insertText: keyword.insertText,
        filterText: keyword.filterText,
        sortText: keyword.sortText,
    };
}
