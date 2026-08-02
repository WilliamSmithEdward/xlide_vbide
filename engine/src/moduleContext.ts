// Assembles the project facts the language features share.
//
// Completion asks "what can be typed here" and hover asks "what is this under the cursor", and
// both answers come from the same assembly: the module entries with the live source swapped in,
// the receiver code names, what `Me` denotes, and the indexed project surfaces. One assembly, so
// the two features describe one project rather than two that drift.

import {
    eventHandlerDocumentTypeForContext,
    type EventHandlerDocumentType,
    type IdentifierCompletionContext,
    type MemberCompletionContext,
} from '../../../xlide_vscode/src/analyzer';
import {
    buildLiveVbaProjectIndex,
    moduleKindFromType,
    projectEditorSymbolContextForModule,
    type VbaProjectAnalysisOptions,
    type VbaProjectModuleInput,
} from '../../../xlide_vscode/src/vbaProjectAnalysis';
import type { ModulePayload } from './protocol';

const WORKBOOK = 'Excel.Workbook';
const WORKSHEET = 'Excel.Worksheet';
const CHART = 'Excel.Chart';

export interface ModuleEntry {
    name: string;
    type: string;
    documentType?: EventHandlerDocumentType;
    source: string;
}

/** The live module a request is about: its text as the editor shows it right now. */
export interface LiveModuleRequest {
    moduleName: string;
    source: string;
    moduleType?: string;
    documentType?: string;
}

export interface AssembledContext {
    entries: ModuleEntry[];
    current: ModuleEntry;
    moduleKind: ReturnType<typeof moduleKindFromType>;
    /** Lowercased document code name -> qualified host type, for receiver resolution. */
    codeNames: Record<string, string>;
    /** Document module names, for identifier lists. */
    codeNameList: string[];
    meType?: string;
    meProjectType?: string;
    projectClassMembers?: MemberCompletionContext['projectClassMembers'];
    projectProcedures?: IdentifierCompletionContext['projectProcedures'];
    projectSymbols?: IdentifierCompletionContext['projectSymbols'];
    projectTypes?: VbaProjectAnalysisOptions['projectTypes'];
}

export function assembleContext(
    seeded: readonly ModulePayload[],
    request: LiveModuleRequest,
): AssembledContext {
    // The live source replaces the seeded copy of the module being worked in; every other module
    // keeps its seeded text, which carries the project facts.
    const entries: ModuleEntry[] = seeded
        .filter((module) => module.moduleName.toLowerCase() !== request.moduleName.toLowerCase())
        .map((module) => ({
            name: module.moduleName,
            type: module.type ?? 'standard',
            documentType: module.documentType as EventHandlerDocumentType | undefined,
            source: module.source,
        }));

    const current: ModuleEntry = {
        name: request.moduleName,
        type: request.moduleType ?? seededTypeOf(seeded, request.moduleName) ?? 'standard',
        documentType: (request.documentType ?? seededDocumentTypeOf(seeded, request.moduleName)) as
            | EventHandlerDocumentType
            | undefined,
        source: request.source,
    };
    entries.push(current);

    const context: AssembledContext = {
        entries,
        current,
        moduleKind: moduleKindFromType(current.type),
        codeNames: codeNamesFor(entries),
        codeNameList: entries.filter((entry) => entry.type === 'document').map((entry) => entry.name),
        meType: meTypeFor(current),
        meProjectType: meProjectTypeFor(current),
    };

    // Project facts. A project that will not index still answers: members of host receivers and
    // keywords need no index at all.
    try {
        const inputs: VbaProjectModuleInput[] = entries.map((entry) => ({
            moduleName: entry.name,
            source: entry.source,
            type: entry.type,
            documentType: entry.documentType,
        }));

        const project = buildLiveVbaProjectIndex(inputs, {
            moduleName: current.name,
            moduleKind: context.moduleKind,
            source: current.source,
        });

        const symbols = projectEditorSymbolContextForModule(project, current.name);
        context.projectClassMembers = symbols.analysisOptions.projectClassMembers;
        context.projectTypes = symbols.analysisOptions.projectTypes;
        context.projectProcedures = symbols.externalProjectProcedures;
        context.projectSymbols = symbols.externalProjectSymbols;
    } catch {
        // Conservative surfaces beat none.
    }

    return context;
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

function codeNamesFor(entries: ModuleEntry[]): Record<string, string> {
    const out: Record<string, string> = {};
    for (const entry of entries) {
        if (entry.type === 'document') {
            out[entry.name.toLowerCase()] = meTypeFor(entry) ?? WORKSHEET;
        }
    }
    return out;
}

function seededTypeOf(seeded: readonly ModulePayload[], moduleName: string): string | undefined {
    return seeded.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.type;
}

function seededDocumentTypeOf(seeded: readonly ModulePayload[], moduleName: string): string | undefined {
    return seeded.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.documentType;
}
