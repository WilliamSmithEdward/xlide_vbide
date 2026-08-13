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

type SharedProjectIndex = ReturnType<typeof buildLiveVbaProjectIndex>;

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
    /** A form's controls, host-supplied: the analyzer's implicit members (xlide_vscode#17). */
    implicitMembers?: readonly { name: string; type: string }[];
    projectClassMembers?: MemberCompletionContext['projectClassMembers'];
    projectProcedures?: IdentifierCompletionContext['projectProcedures'];
    projectSymbols?: IdentifierCompletionContext['projectSymbols'];
    projectTypes?: VbaProjectAnalysisOptions['projectTypes'];
}

/**
 * Assembled contexts by seeded-module set and module. The seeded array is replaced whole when a
 * project reseeds, so its identity is the cache key: a hit means the project facts describe the
 * sources the engine holds, and a reseed invalidates everything at once by changing the key.
 *
 * What is cached is built from the seeded copy of the requested module, not its live text. The
 * live text still drives every resolver directly; only the cross-module facts lag until the
 * next write-back, which is the same bargain the editor extension strikes - and what turns the
 * per-keystroke cost of a completion from indexing the whole project into scanning one module.
 */
const contextCache = new WeakMap<readonly ModulePayload[], Map<string, AssembledContext>>();

/*
 * ONE PROJECT INDEX PER SEED, shared by every module's context build.
 *
 * The index used to be rebuilt inside buildContext, once per MODULE per seed, and the analyzer's
 * parse cache holds eight entries - so past eight modules each rebuild evicted every entry as it
 * walked and the NEXT module's build re-parsed the project from scratch. Measured on a twelve
 * 500-line-module project: the first completion in each module cost a median 11ms against 2.6ms
 * warm, 161ms for the first walk across the modules, and every reseed - every write-back that
 * changed text - made all of it cold again (2026-08-12, the audit's C13).
 *
 * Sharing is sound because the per-module builds were near-identical by construction: the
 * override buildContext passed carried the SEEDED source on purpose ("a cache must not embalm
 * whichever keystroke happened to build it"), so the only per-module difference was the current
 * module's own kind spelling - and a module's own contribution is excluded from its own external
 * context anyway (externalProjectProcedures is the point). The per-module HALF of the work,
 * projectEditorSymbolContextForModule, still runs per module below.
 *
 * `null` is remembered too: an index that will not build should not be re-attempted per module,
 * and the conservative empty surfaces are the same answer the old catch produced.
 *
 * The dispatcher's fingerprint index is deliberately NOT taken from here: it builds with
 * ignoreInvalidModules OFF so a project holding one unparseable module reads as "unknown, never
 * reused" rather than as a fingerprint over the valid remainder, and that difference is
 * load-bearing for freshness.
 */
const projectIndexCache = new WeakMap<readonly ModulePayload[], SharedProjectIndex | null>();

function sharedProjectIndex(seeded: readonly ModulePayload[]): SharedProjectIndex | null {
    if (projectIndexCache.has(seeded)) {
        return projectIndexCache.get(seeded) ?? null;
    }

    let built: SharedProjectIndex | null = null;
    try {
        const inputs: VbaProjectModuleInput[] = seeded.map((module) => ({
            moduleName: module.moduleName,
            source: module.source,
            type: module.type ?? 'standard',
            documentType: module.documentType as EventHandlerDocumentType | undefined,
        }));
        built = buildLiveVbaProjectIndex(inputs);
    } catch {
        built = null;
    }

    projectIndexCache.set(seeded, built);
    return built;
}

export function assembleContext(
    seeded: readonly ModulePayload[],
    request: LiveModuleRequest,
): AssembledContext {
    let byModule = contextCache.get(seeded);
    if (!byModule) {
        byModule = new Map();
        contextCache.set(seeded, byModule);
    }

    const key = request.moduleName.toLowerCase();
    const cached = byModule.get(key);
    if (cached) {
        return cached;
    }

    const assembled = buildContext(seeded, request);
    byModule.set(key, assembled);
    return assembled;
}

function buildContext(
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
        // The seeded copy, deliberately: this entry only feeds the cached project facts, and a
        // cache must not embalm whichever keystroke happened to build it. Every resolver gets
        // the live text through its own parameters.
        source: seededSourceOf(seeded, request.moduleName) ?? request.source,
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
        // From the SEEDED module, like every cross-module fact here: the designer facts ride
        // the seed, and a control set that changes without a reseed lags the same way the rest
        // of the project's facts do.
        implicitMembers: seededImplicitMembersOf(seeded, request.moduleName),
    };

    // Project facts, from the ONE index this seed carries (see sharedProjectIndex above); only
    // the per-module symbol view is derived here. A project that will not index still answers:
    // members of host receivers and keywords need no index at all.
    try {
        const project = sharedProjectIndex(seeded);
        if (project) {
            const symbols = projectEditorSymbolContextForModule(project, current.name);
            context.projectClassMembers = symbols.analysisOptions.projectClassMembers;
            context.projectTypes = symbols.analysisOptions.projectTypes;
            context.projectProcedures = symbols.externalProjectProcedures;
            context.projectSymbols = symbols.externalProjectSymbols;
        }
    } catch {
        // Conservative surfaces beat none.
    }

    return context;
}

/** Maps a document module to the host type that `Me` denotes inside it. */
function meTypeFor(entry: ModuleEntry | undefined): string | undefined {
    if (!entry) {
        return undefined;
    }

    // A form's Me is the form: MSForms.UserForm plus the controls and the module's own code,
    // which the analyzer composes once BOTH me fields are set (the vbide issue #2 wiring
    // note). The controls alone are not enough.
    if (entry.type === 'userform') {
        return 'MSForms.UserForm';
    }

    if (entry.type !== 'document') {
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

function seededSourceOf(seeded: readonly ModulePayload[], moduleName: string): string | undefined {
    return seeded.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.source;
}

function seededDocumentTypeOf(seeded: readonly ModulePayload[], moduleName: string): string | undefined {
    return seeded.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.documentType;
}

function seededImplicitMembersOf(
    seeded: readonly ModulePayload[],
    moduleName: string,
): readonly { name: string; type: string }[] | undefined {
    const members = seeded.find(
        (module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.implicitMembers;
    return members && members.length > 0 ? members : undefined;
}
