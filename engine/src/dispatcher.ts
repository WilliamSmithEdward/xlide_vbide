// Maps the wire protocol onto the analyzer.
//
// The analyzer already has a request handler built for exactly this shape: a pure, synchronous
// object that owns project state and answers analysis requests, written so everything crossing its
// boundary is plain data. It was built to sit behind a worker thread's message port. A pipe is the
// same contract with a different transport, so it is reused rather than reimplemented, which keeps
// one analysis path shared with the editor extension instead of two that can disagree.

import { syncPlan, type SyncPlanParams } from './sync.js';
import { AnalysisWorkerState } from '../../../xlide_vscode/src/analysisWorkerLogic';
import type { AnalysisWorkerRequest } from '../../../xlide_vscode/src/analysisWorkerProtocol';
import type { VbaModuleAnalysisDiagnostic } from '../../../xlide_vscode/src/vbaModuleAnalysis';
import {
    buildVbaProjectIndex,
    moduleKindFromType,
    projectAnalysisOptionsForModule,
    projectProcedureSignatures,
} from '../../../xlide_vscode/src/vbaProjectAnalysis';
import type { EventHandlerDocumentType } from '../../../xlide_vscode/src/analyzer';
import { codeActionsFor } from './codeActions';
import { completionsFor } from './completion';
import { forgetProjectWords, outlineFor, projectWordsFor } from './outline';
import { searchModules } from './search';
import { hoverFor } from './hover';
import { canonicalCaseFor, loopSyncFor, smartEnterFor } from './onType';
import {
    assembleSymbols,
    definitionsFor,
    lineStarts,
    referencesFor,
    renameFor,
    renameModuleFor,
    toLineColumn,
    type ProjectSymbols,
} from './navigation';
import { semanticTokensFor } from './semantic';
import { signatureHelpFor } from './signature';
import {
    ErrorCode,
    type CanonicalCaseParams,
    type CanonicalCaseResult,
    type CodeActionParams,
    type CodeActionResult,
    type CompletionParams,
    type CompletionResult,
    type DiagnosticsParams,
    type DidChangeParams,
    type DiagnosticsResult,
    type HoverParams,
    type HoverResult,
    type JsonRpcError,
    type LoopSyncParams,
    type LoopSyncResult,
    type ModulePayload,
    type NavigationParams,
    type NavigationResult,
    type OutlineParams,
    type OutlineResult,
    type ProjectOpenParams,
    type RenameModuleParams,
    type RenameParams,
    type RenameResult,
    type SearchParams,
    type SearchResult,
    type SemanticTokensParams,
    type SemanticTokensResult,
    type SignatureHelpParams,
    type SignatureHelpResult,
    type LiveSourceParams,
    type LiveSourceResult,
    type SmartEnterParams,
    type SmartEnterResult,
} from './protocol';

function liveKey(projectId: string, moduleName: string): string {
    return `${projectId}\0${moduleName.toLowerCase()}`;
}

/**
 * Per seeded module array: what each module's analysis depends on from the REST of the project,
 * as a string that can be compared.
 *
 * This is the exact thing, not an approximation of it. Analysing a module takes the module's own
 * text plus `projectAnalysisOptionsForModule` for it - the project's types, its class members and
 * its procedure signatures as that module sees them. Those two inputs decide the findings, so two
 * requests that agree on both have the same answer and the second need not be computed.
 *
 * Keyed on the seeded array's identity, which is replaced whole on every reseed, so a project
 * that reseeds gets a fresh set and one that does not keeps its own. Built lazily and for the
 * whole project at once, because the index it needs is a project-wide build and doing it per
 * module would be doing it once per module.
 */
const crossModuleFacts = new WeakMap<readonly ModulePayload[], Map<string, string> | null>();

/**
 * The options object as a string that actually reflects its contents.
 *
 * PLAIN `JSON.stringify` HAS A BLIND SPOT HERE. `projectAnalysisOptionsForModule` returns
 * `projectProcedures` as a `Map`, and `JSON.stringify(new Map([...]))` is `"{}"` - not an error,
 * not a warning, just an empty object where every procedure signature in the project should be.
 * A fingerprint that cannot see the signatures is a fingerprint that cannot notice the one thing
 * it exists to notice.
 *
 * Honest about what that did and did not cause: `test/freshness.mjs` passes with this handling
 * REMOVED, because the other fields of the options object happen to move when a signature does,
 * so the memo still invalidated. It was a blind spot rather than an active defect, and it is
 * closed here because relying on a neighbouring field to notice is not a design.
 *
 * Maps and Sets become their entries. Order is insertion order, which is deterministic for the
 * same project, so equal contents still produce equal strings.
 */
function describe(value: unknown): string {
    return JSON.stringify(value, (_key, held: unknown) => {
        if (held instanceof Map) { return { '#map': [...held] }; }
        if (held instanceof Set) { return { '#set': [...held] }; }
        return held;
    });
}

function crossModuleFingerprint(
    seeded: readonly ModulePayload[],
    moduleName: string,
): string | undefined {
    let held = crossModuleFacts.get(seeded);
    if (held === undefined) {
        held = null;
        try {
            const index = buildVbaProjectIndex(seeded.map((module) => ({
                moduleName: module.moduleName,
                source: module.source,
                type: module.type,
                documentType: module.documentType as EventHandlerDocumentType | undefined,
            })));
            const procedures = projectProcedureSignatures(index);

            held = new Map<string, string>();
            for (const module of seeded) {
                // KEPT WHOLE, not digested. A digest would hold 44 bytes instead of the few
                // hundred kilobytes this is for a large project, and it would cost a hash over
                // those bytes on every seed to save a comparison that is already cheap: about
                // 2ms per pass spent to save about 60 microseconds of it. Memory is the cheaper
                // side of that trade here (developer, 2026-08-08).
                //
                // And the instance is what makes the comparison fast. Every request under one
                // seed is handed the SAME string out of this map, so the check against the last
                // answer's copy is a pointer comparison until the project is seeded again.
                held.set(
                    module.moduleName.toLowerCase(),
                    describe(projectAnalysisOptionsForModule(index, module.moduleName, procedures)));
            }
        } catch {
            // Null rather than an empty map, deliberately. An empty map would answer "no facts"
            // for every module, and "no facts" compares equal to itself, so a project whose index
            // will not build would look unchanged forever and freeze its findings. Null means
            // unknown, and unknown is never reused.
            held = null;
        }

        crossModuleFacts.set(seeded, held);
    }

    return held?.get(moduleName.toLowerCase());
}

/** The analyse request, narrowed: the only worker request this engine keeps hold of. */
type AnalyzeRequest = Extract<AnalysisWorkerRequest, { kind: 'analyze' }>;

/** Thrown to answer a request with a JSON-RPC error rather than a result. */
export class RpcError extends Error {
    constructor(
        readonly code: number,
        message: string,
        readonly data?: unknown,
    ) {
        super(message);
        this.name = 'RpcError';
    }

    toJson(): JsonRpcError {
        return { code: this.code, message: this.message, data: this.data };
    }
}

export class Dispatcher {
    private readonly analysis = new AnalysisWorkerState();
    private readonly generations = new Map<string, number>();

    /**
     * The modules each project was last seeded with, kept for completion contexts. The analysis
     * state holds them too, but behind its own boundary; completions need the payloads.
     */
    private readonly seededModules = new Map<string, ModulePayload[]>();

    /**
     * The live text of modules being edited, by project and module, fed by didChange. One
     * string instance per keystroke, shared by every feature that asks about the module - which
     * is what lets the analyzer's identity-keyed token cache actually hit.
     */
    private readonly liveSources = new Map<string, string>();

    /**
     * The last outline answered per module, keyed by the exact text it described. The tree
     * re-asks after every push it notices, and between edits a module's source is the same
     * string, so parsing 26,000 lines again to repeat the same procedures is pure waste.
     */
    private readonly outlineMemo = new Map<string, { source: string; result: OutlineResult }>();

    /**
     * The last colouring answered per module, keyed by the exact text it described - the same
     * bargain the outline strikes, and for a stronger reason: the surface re-asks for the whole
     * module's tokens after every edit, and two passes over 26,000 lines per keystroke is the
     * kind of cost that shows up as the editor feeling slow rather than as anything visible.
     */
    private readonly semanticMemo = new Map<string, { source: string; result: SemanticTokensResult }>();

    /** One workbook's symbol index, kept against the exact module texts it was built from. */
    /**
     * The assembled symbols per project, and what they were assembled FROM.
     *
     * The fingerprint carries each module's name as well as its source. It carried only the
     * sources until 2026-08-10, and a rename is the one edit that changes a name and no text: the
     * memo hit, the old assembly came back, and it still knew the module by the name it no longer
     * had. So a module nothing references could be renamed once and the second rename was refused
     * as not being a module of this workbook, by an assembly that was one rename out of date.
     *
     * It only ever showed on a rename with nothing to replace. Replacing a mention rewrites
     * another module's text, which misses the memo and rebuilds, which is why every rename that
     * did something worked and only the ones that did nothing else broke.
     */
    private readonly symbolsMemo = new Map<string, { fingerprint: string[]; symbols: ProjectSymbols }>();

    /**
     * The last analysis of each module, kept whole: the findings as the analyzer made them, and
     * the text they describe. Quick fixes are resolved from these rather than from anything the
     * surface sends back, because a finding carries fix data - the missing argument's name, the
     * unclosed block's expected closer - that never crosses to the surface at all.
     *
     * The request is kept too, so a fix asked for against text the diagnostics have not caught up
     * with can be answered by analysing that text under the same options the squiggles used.
     */
    private readonly lastAnalysis = new Map<string, {
        source: string;
        diagnostics: readonly VbaModuleAnalysisDiagnostic[];
        request: AnalyzeRequest;
        /**
         * What the last DIAGNOSTICS request was answered under, when it was one: the cross-module
         * facts and the shape of the request. Absent on entries written by a quick fix, which
         * never reuses them for anything but its own source comparison, and which must not let
         * one be reused as a diagnostics answer.
         */
        facts?: string;
        shape?: string;
        /**
         * The caret this answer was computed with, or null for none. Held apart from `shape`
         * because a request with no caret may reuse an answer computed with one, and not the
         * other way round. See the comparison in `diagnostics`.
         */
        caret?: number | null;
        mode?: DiagnosticsResult['mode'];
        /** The positioned reply, so a hit costs nothing rather than a walk of the module. */
        answer?: DiagnosticsResult;
    }>();

    private nextRequestId = 1;
    private initialized = false;

    /** True once the add-in has asked to shut down, so the transport can close. */
    shuttingDown = false;

    handle(method: string, params: unknown): unknown {
        switch (method) {
            case 'initialize':
                this.initialized = true;
                return { engine: 'xlide', protocol: 1 };

            case 'shutdown':
                this.shuttingDown = true;
                return null;

            case 'project/open':
                return this.openProject(this.require<ProjectOpenParams>(params));

            case 'project/close':
                return this.closeProject(this.require<{ projectId: string }>(params));

            case 'module/didClose':
                this.analysis.handle({ kind: 'forget', docKey: this.require<{ documentKey: string }>(params).documentKey });
                return null;

            case 'textDocument/diagnostics':
                return this.diagnostics(this.require<DiagnosticsParams>(params));

            case 'textDocument/completion':
                return this.completion(this.require<CompletionParams>(params));

            case 'textDocument/hover':
                return this.hover(this.require<HoverParams>(params));

            case 'textDocument/signatureHelp':
                return this.signatureHelp(this.require<SignatureHelpParams>(params));

            case 'textDocument/smartEnter':
                return this.smartEnter(this.require<SmartEnterParams>(params));

            case 'textDocument/canonicalCase':
                return this.canonicalCase(this.require<CanonicalCaseParams>(params));

            case 'textDocument/loopSync':
                return this.loopSync(this.require<LoopSyncParams>(params));

            case 'textDocument/codeAction':
                return this.codeAction(this.require<CodeActionParams>(params));

            case 'textDocument/semanticTokens':
                return this.semanticTokens(this.require<SemanticTokensParams>(params));

            case 'textDocument/definition':
                return this.definition(this.require<NavigationParams>(params));

            case 'textDocument/references':
                return this.references(this.require<NavigationParams>(params));

            case 'textDocument/rename':
                return this.rename(this.require<RenameParams>(params));

            case 'workspace/renameModule':
                return this.renameModule(this.require<RenameModuleParams>(params));

            case 'textDocument/outline':
                return this.outline(this.require<OutlineParams>(params));

            case 'sync/plan':
                return this.syncPlan(this.require<SyncPlanParams>(params));

            case 'workspace/search':
                return this.search(this.require<SearchParams>(params));

            case 'textDocument/didChange':                return this.didChange(this.require<DidChangeParams>(params));

            /*
             * WHAT THE ENGINE IS HOLDING for a module, which nothing could see until now.
             *
             * Every finding is computed against this copy, and it is maintained incrementally by
             * didChange rather than re-sent whole. So when a squiggle lands on the wrong line,
             * the question is always whether this copy matches the surface - and there was no
             * way to ask. A finding was seen one line out after a format on 2026-08-08, healed
             * before it could be diagnosed, and the one thing that would have settled it in a
             * single call did not exist.
             *
             * Answers the text and a line count; the caller compares against the surface's.
             */
            case 'debug/liveSource':                      return this.liveSource(this.require<LiveSourceParams>(params));

            default:
                throw new RpcError(ErrorCode.MethodNotFound, `Unknown method: ${method}`);
        }
    }

    /** Every module in scope, live text over seeded, streamed to the matcher. */
    private search(params: SearchParams): SearchResult {
        this.requireInitialized();

        const dispatcher = this;
        function* inScope(): Generator<{ projectId: string; module: string; source: string }> {
            for (const [projectId, modules] of dispatcher.seededModules) {
                if (params.scope !== 'all'
                    && params.projectId
                    && projectId.toLowerCase() !== params.projectId.toLowerCase()) {
                    continue;
                }

                for (const module of modules) {
                    if (params.scope === 'module'
                        && params.module
                        && module.moduleName.toLowerCase() !== params.module.toLowerCase()) {
                        continue;
                    }

                    yield {
                        projectId,
                        module: module.moduleName,
                        source: dispatcher.liveSources.get(liveKey(projectId, module.moduleName)) ?? module.source,
                    };
                }
            }
        }

        return searchModules(inScope(), params);
    }
    private openProject(params: ProjectOpenParams): { modules: number; types: string[]; procedures: string[] } {
        this.requireInitialized();

        this.analysis.handle({
            kind: 'seed',
            workbookKey: params.projectId,
            generation: params.generation,
            modules: params.modules.map((module) => ({
                moduleName: module.moduleName,
                source: module.source,
                type: module.type,
                documentType: module.documentType,
            })),
        });

        this.generations.set(params.projectId, params.generation);
        this.seededModules.set(params.projectId, params.modules.map((module) => ({ ...module })));

        /*
         * A MODULE THAT IS GONE TAKES ITS ANALYSIS WITH IT.
         *
         * Everything memoised per module is keyed by project and NAME. `project/close` prunes all
         * of it, but a component being removed from a project that stays open pruned nothing, so
         * the entries outlived the module. Add a module with a name that has been used before and
         * it inherits the dead one's findings: the memo compares source, facts and shape, and a
         * fresh module whose text matches what the old one last held matches all three.
         *
         * That is not only a test artifact. Delete a module, add another with the same name, and
         * the Problems pane shows the deleted module's errors against the new one's code.
         *
         * It is also what made the end-to-end freshness suite fail about two runs in five: it
         * brings its own two modules, always with the same names, and removes them at the end. A
         * run inherited the previous run's answers and reported findings on a module that had
         * been created three lines earlier (2026-08-08).
         */
        const present = new Set(params.modules.map((module) => module.moduleName.toLowerCase()));
        const prefix = `${params.projectId}\0`;
        for (const key of [...this.lastAnalysis.keys()]) {
            if (key.startsWith(prefix) && !present.has(key.slice(prefix.length))) {
                this.lastAnalysis.delete(key);
                this.semanticMemo.delete(key);
                this.outlineMemo.delete(key);
                this.liveSources.delete(key);
            }
        }

        // The project's own words, for the surface's tokenizer: names that are types and names
        // that are procedures. This is what lets `ROneCOne.Create(...)` read as a type and a
        // call while `values(index, 1)` stays a variable - the distinction the extension makes
        // with its semantic tokens.
        const facts = projectWordsFor(params.projectId, params.modules);
        return { modules: params.modules.length, types: facts.types, procedures: facts.procedures };
    }

    private closeProject(params: { projectId: string }): null {
        this.generations.delete(params.projectId);

        // THE ANALYZER'S OWN PER-DOCUMENT STATE, which nothing else here releases.
        //
        // It keeps an incremental parse per document and drops one when it is told that document
        // closed - `module/didClose`, which this product has never sent, because a module's TAB
        // closing is not the module leaving the project. So the state survived the WORKBOOK
        // closing too: every module a session ever analysed stayed held for the life of the
        // engine, and a developer moving between workbooks all day accumulated all of them
        // (2026-08-09).
        //
        // Released here rather than by starting to send didClose, because this is the point where
        // the documents are certainly gone, and it needs no cooperation from the shim to be right.
        for (const module of this.seededModules.get(params.projectId) ?? []) {
            this.analysis.handle({ kind: 'forget', docKey: `${params.projectId}/${module.moduleName}` });
        }

        this.seededModules.delete(params.projectId);

        const prefix = `${params.projectId}\0`;
        for (const key of this.lastAnalysis.keys()) {
            if (key.startsWith(prefix)) {
                this.lastAnalysis.delete(key);
            }
        }

        for (const key of this.semanticMemo.keys()) {
            if (key.startsWith(prefix)) {
                this.semanticMemo.delete(key);
            }
        }

        this.symbolsMemo.delete(params.projectId);
        forgetProjectWords(params.projectId);
        return null;
    }

    private completion(params: CompletionParams): CompletionResult {
        this.requireInitialized();

        // Not gated on generation: the live text of the module being typed in is what is asked
        // about, and the seeded copies of the others are current enough for the facts they lend.
        const source = this.sourceFor(params);
        if (source === undefined) {
            return { items: [] };
        }

        return { items: completionsFor(this.seededModules.get(params.projectId) ?? [], { ...params, source }) };
    }

    private hover(params: HoverParams): HoverResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { hover: null };
        }

        return { hover: hoverFor(this.seededModules.get(params.projectId) ?? [], { ...params, source }) };
    }

    private signatureHelp(params: SignatureHelpParams): SignatureHelpResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { signature: null };
        }

        return { signature: signatureHelpFor(this.seededModules.get(params.projectId) ?? [], { ...params, source }) };
    }

    private smartEnter(params: SmartEnterParams): SmartEnterResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { edits: [], caret: null };
        }

        return smartEnterFor({ ...params, source });
    }

    private canonicalCase(params: CanonicalCaseParams): CanonicalCaseResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { edits: [] };
        }

        return { edits: canonicalCaseFor(this.seededModules.get(params.projectId) ?? [], { ...params, source }) };
    }

    private loopSync(params: LoopSyncParams): LoopSyncResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { edits: [] };
        }

        return { edits: loopSyncFor({ ...params, source }) };
    }

    private rename(params: RenameParams): RenameResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { modules: [], refused: 'This module is not one the engine holds.' };
        }

        const symbols = this.symbolsFor(params.projectId, params.moduleName, source);
        return renameFor(symbols, params.moduleName, source, params.offset, params.newName);
    }

    private renameModule(params: RenameModuleParams): RenameResult {
        this.requireInitialized();

        // Anchored on the module's own current text, which is what symbolsFor wants; the answer
        // is about every OTHER module, so which one is passed here only decides whose live text
        // wins over its seeded copy.
        // Checked against what the project was SEEDED with, not against the assembly: the
        // assembly synthesises a module it has never heard of so that a module opened before the
        // first seed still answers, and that would let a rename of a module nobody has silently
        // succeed at renaming nothing.
        const seeded = this.seededModules.get(params.projectId) ?? [];
        const wanted = params.moduleName.toLowerCase();
        if (!seeded.some((module) => module.moduleName.toLowerCase() === wanted)) {
            // NAMES WHAT IT DOES HAVE. A refusal that only says "not a module of this workbook"
            // about a module the developer can see in the tree sends the reader to the object
            // model, which agrees with them, and the disagreement is here. Whether the seed is
            // stale or the project is the wrong one is the whole question, and both answers are
            // in this list.
            const known = seeded.map((module) => module.moduleName).sort();
            return {
                modules: [],
                oldName: params.moduleName,
                refused: `'${params.moduleName}' is not a module of this workbook. `
                    + `The engine was seeded with: ${known.join(', ') || '(nothing)'}.`,
            };
        }

        const source = this.sourceFor({ projectId: params.projectId, moduleName: params.moduleName })
            ?? '';
        const symbols = this.symbolsFor(params.projectId, params.moduleName, source);
        const answer = renameModuleFor(symbols, params.moduleName, params.newName);
        return {
            modules: answer.modules,
            oldName: params.moduleName,
            refused: answer.refused,
            module: answer.refused ? undefined : params.moduleName,
        };
    }

    private definition(params: NavigationParams): NavigationResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { locations: [] };
        }

        const symbols = this.symbolsFor(params.projectId, params.moduleName, source);
        return { locations: definitionsFor(symbols, params.moduleName, source, params.offset) };
    }

    private references(params: NavigationParams): NavigationResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { locations: [] };
        }

        const symbols = this.symbolsFor(params.projectId, params.moduleName, source);
        return {
            locations: referencesFor(
                symbols,
                params.moduleName,
                source,
                params.offset,
                params.includeDeclaration ?? true),
        };
    }

    /**
     * A workbook's symbols over its current text: live where the surface is typing, seeded
     * elsewhere, with the request's own module text winning over both.
     *
     * Cached on the exact strings it was built from, compared by identity. Indexing a whole
     * workbook is far too expensive to repeat per navigation, and the surface asks on every
     * Ctrl+click and every open of the references list; between edits the strings are the same
     * instances, so the comparison is a walk over a handful of pointers.
     */
    private symbolsFor(projectId: string, moduleName: string, source: string): ProjectSymbols {
        const seeded = this.seededModules.get(projectId) ?? [];
        const wanted = moduleName.toLowerCase();

        const modules = seeded.map((module) => ({
            moduleName: module.moduleName,
            type: module.type,
            documentType: module.documentType,
            source: module.moduleName.toLowerCase() === wanted
                ? source
                : this.liveSources.get(liveKey(projectId, module.moduleName)) ?? module.source,
        }));

        // A module the surface is asking about that the project was never seeded with: analysed
        // alone rather than not at all, which is what a module opened before the first seed is.
        if (!modules.some((module) => module.moduleName.toLowerCase() === wanted)) {
            modules.push({ moduleName, type: 'standard', documentType: undefined, source });
        }

        // NAME AND SOURCE, because a rename changes the first and not the second. Joined by a
        // character a name cannot hold, so no two different pairs can spell the same fingerprint.
        const fingerprint = modules.map((module) => `${module.moduleName}\u0000${module.source}`);

        const memo = this.symbolsMemo.get(projectId);
        if (memo
            && memo.fingerprint.length === fingerprint.length
            && memo.fingerprint.every((held, index) => held === fingerprint[index])) {
            return memo.symbols;
        }

        const symbols = assembleSymbols(modules);
        this.symbolsMemo.set(projectId, { fingerprint, symbols });
        return symbols;
    }

    private semanticTokens(params: SemanticTokensParams): SemanticTokensResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { tokens: [] };
        }

        const key = liveKey(params.projectId, params.moduleName);
        const memo = this.semanticMemo.get(key);
        if (memo?.source === source) {
            return memo.result;
        }

        const result: SemanticTokensResult = {
            tokens: semanticTokensFor(this.seededModules.get(params.projectId) ?? [], { ...params, source }),
        };

        this.semanticMemo.set(key, { source, result });
        return result;
    }

    private codeAction(params: CodeActionParams): CodeActionResult {
        this.requireInitialized();

        const source = this.sourceFor(params);
        if (source === undefined) {
            return { actions: [] };
        }

        const analysed = this.analysisFor(params, source);
        if (!analysed) {
            return { actions: [] };
        }

        return {
            actions: codeActionsFor(source, analysed, { start: params.start, end: params.end }),
        };
    }

    /**
     * The findings for a module's current text: the ones diagnostics last reported when they
     * describe this exact text, a fresh pass when they do not.
     *
     * The fresh pass is what keeps a fix honest between a keystroke and the diagnostics that
     * follow it. Fixes are spans into the text they will edit, so answering from findings made
     * against older text would place the edit by arithmetic that no longer holds. Null when the
     * module has never been analysed, or when the engine's sources are stale enough that
     * analysing would report on text the developer is not looking at.
     */
    private analysisFor(
        params: CodeActionParams,
        source: string,
    ): readonly VbaModuleAnalysisDiagnostic[] | null {
        const key = liveKey(params.projectId, params.moduleName);
        const memo = this.lastAnalysis.get(key);
        if (memo?.source === source) {
            return memo.diagnostics;
        }

        const request: AnalyzeRequest = {
            kind: 'analyze',
            requestId: this.nextRequestId++,
            docKey: memo?.request.docKey ?? key,
            workbookKey: params.projectId,
            generation: this.generations.get(params.projectId),
            source,
            moduleName: params.moduleName,
            moduleType: params.moduleType ?? memo?.request.moduleType,
            moduleKind: moduleKindFromType(params.moduleType ?? memo?.request.moduleType),
            documentType: params.documentType ?? memo?.request.documentType,
            // Inherited rather than sent: a fix must be offered under the same rules that drew
            // the squiggle, and the last analysis of this module is what drew it.
            severityOverrides: memo?.request.severityOverrides,
        };

        const response = this.analysis.handle(request);
        if (response?.kind !== 'result') {
            // No seed, or the analyzer threw. A quick fix that fails is a lightbulb that does not
            // open, which is what the developer already sees when there is nothing to fix.
            return null;
        }

        this.lastAnalysis.set(key, { source, diagnostics: response.diagnostics, request });
        return response.diagnostics;
    }

    /**
     * What an import or export would do, decided by the companion editor's own planner.
     *
     * Async, unlike everything else here, because the plan reads the FOLDER: their code compares
     * each module against the file beside it and works out which files are stale. The modules
     * arrive with the request rather than being read from a workbook, because this project is open
     * in Excel and the file on disk is stale by definition.
     */
    private syncPlan(params: SyncPlanParams): Promise<unknown> {
        return syncPlan(params);
    }

    private outline(params: OutlineParams): OutlineResult {
        this.requireInitialized();

        // Live source when there is one, the seeded copy otherwise.
        const source = this.sourceFor(params);
        const key = liveKey(params.projectId, params.moduleName);
        if (source !== undefined) {
            const memo = this.outlineMemo.get(key);
            if (memo && memo.source === source) {
                return memo.result;
            }
        }

        const result: OutlineResult = {
            procedures: outlineFor(this.seededModules.get(params.projectId) ?? [], {
                ...params,
                source,
            }),
        };

        if (source !== undefined) {
            this.outlineMemo.set(key, { source, result });
        }

        return result;
    }

    private liveSource(params: LiveSourceParams): LiveSourceResult {
        this.requireInitialized();

        // Through liveKey, not a hand-rolled join: the live map is keyed on a NUL and a lowercased
        // module name, and asking it with a slash and the caller's casing answered "the engine is
        // holding nothing" about a module it was holding.
        const held = this.liveSources.get(liveKey(params.projectId, params.moduleName));
        return {
            held: held !== undefined,
            lines: held === undefined ? 0 : held.split(/\r?\n/).length,
            source: params.includeText === true ? (held ?? null) : null,
        };
    }

    private didChange(params: DidChangeParams): null {
        const key = liveKey(params.projectId, params.moduleName);

        if (params.source !== undefined) {
            this.liveSources.set(key, params.source);
            return null;
        }

        const current = this.liveSources.get(key) ?? this.seededSourceOf(params.projectId, params.moduleName);
        if (current === undefined || !params.edits) {
            return null;
        }

        let text = current;
        for (const edit of params.edits) {
            if (edit.start < 0 || edit.end < edit.start || edit.end > text.length) {
                // A range that does not fit the text held here means the stream desynchronised;
                // holding a mangled copy would answer wrongly forever. Falling back to the
                // seeded copy answers staler but true, until the next full push.
                this.liveSources.delete(key);
                return null;
            }

            text = text.slice(0, edit.start) + edit.text + text.slice(edit.end);
        }

        this.liveSources.set(key, text);
        return null;
    }

    /** The text a request is about: sent with it, held live from didChange, or the seeded copy. */
    private sourceFor(params: { projectId: string; moduleName: string; source?: string }): string | undefined {
        return params.source
            ?? this.liveSources.get(liveKey(params.projectId, params.moduleName))
            ?? this.seededSourceOf(params.projectId, params.moduleName);
    }

    private seededSourceOf(projectId: string, moduleName: string): string | undefined {
        return this.seededModules.get(projectId)
            ?.find((module) => module.moduleName.toLowerCase() === moduleName.toLowerCase())?.source;
    }

    private diagnostics(params: DiagnosticsParams): DiagnosticsResult {
        this.requireInitialized();

        const source = params.projectId !== undefined
            ? this.sourceFor({ projectId: params.projectId, moduleName: params.moduleName, source: params.source })
            : params.source;
        if (source === undefined) {
            return { diagnostics: [] };
        }

        /*
         * THE SAME TEXT UNDER THE SAME PROJECT HAS THE SAME FINDINGS.
         *
         * A full pass asks about every module of a project, and a pass is provoked by a write-back
         * to one of them. The other modules are byte-identical to the last pass and, unless the
         * written module changed a DECLARATION, so is everything they depend on. Re-deriving their
         * findings produces the list already on screen, at full price: on the perf fixture that was
         * 446ms of an 476ms pass, and the pipe is serialised, so it was also 446ms that every
         * completion and hover queued behind (2026-08-08).
         *
         * The comparison is exact rather than a heuristic about what "looks like" a declaration
         * change. `facts` IS the analyzer's cross-module input, and `shape` is every other thing
         * about the request that can move an answer - the caret offset among them, because it
         * suppresses the transient complaints of a half-typed expression and a cached answer from
         * a different caret would put them back.
         */
        const key = params.projectId !== undefined
            ? liveKey(params.projectId, params.moduleName)
            : undefined;

        // No seed, no reuse. Not `?? []`: an empty array is a fresh object every call, so it
        // would key a fresh WeakMap entry and build an empty index per request, to answer
        // "unknown" every time anyway.
        const seeded = params.projectId !== undefined
            ? this.seededModules.get(params.projectId)
            : undefined;

        const facts = seeded !== undefined
            ? crossModuleFingerprint(seeded, params.moduleName)
            : undefined;

        // The caret is NOT part of the shape. It gets its own comparison below, because the two
        // callers use it differently and one of them can accept the other's answer.
        const shape = JSON.stringify([
            params.moduleType ?? null,
            params.documentType ?? null,
            params.severityOverrides ?? null,
        ]);

        const caret = params.activeIncompleteExpressionOffset ?? null;

        if (key !== undefined && facts !== undefined) {
            const memo = this.lastAnalysis.get(key);

            /*
             * ONE ANALYSIS PER DOCUMENT VERSION, WHICH IS THE WHOLE POINT.
             *
             * Two callers ask about the same module. The LIVE path asks on a pause in typing and
             * sends the caret, which holds back the transient complaints of a half-typed
             * expression; those findings become the squiggles. The PASS asks afterwards with no
             * caret, and those findings become the Problems list.
             *
             * Same module, same text, seconds apart, analysed twice. On the 64,802-line fixture
             * that was 3,868ms of a 5,252ms edit spent re-deriving what had just been derived
             * (2026-08-08).
             *
             * Every mature language server treats the problems view as a VIEW OVER THE PUBLISHED
             * DIAGNOSTICS rather than as a second computation, which is also why theirs cannot
             * disagree with the underlines. So:
             *
             *   a request WITHOUT a caret accepts an answer computed WITH one, for the same text.
             *   a request WITH a caret needs the same caret.
             *
             * The asymmetry is the correctness of it. A caret-suppressed answer holds back a
             * transient at a known offset, and that is exactly what the developer is looking at,
             * so the list may show it. The reverse would be wrong: serving an UNsuppressed answer
             * to a live request puts the error back under the cursor mid-expression, which is the
             * thing the suppression exists to prevent.
             *
             * The designated consequence, recorded because it is a real change: the Problems list
             * for the module being edited now suppresses the same transient the squiggle does.
             * The list and the underline agree by construction, which they did not before.
             *
             * `answer` present is part of the test, not an assumption: entries written by a quick
             * fix carry findings and no reply, and one of those must not be served as one.
             */
            const usable = memo?.answer
                && memo.source === source
                && memo.facts === facts
                && memo.shape === shape
                && (caret === null || caret === memo.caret);

            if (usable && memo?.answer) {
                // The ANSWER, not the findings to be positioned again. Positioning walks the whole
                // module to build its line-start table, so re-deriving it for a reply that cannot
                // have changed was a scan of 1.5 MB per memo hit on the largest module.
                return memo.answer;
            }
        }

        const request: AnalyzeRequest = {
            kind: 'analyze',
            requestId: this.nextRequestId++,
            docKey: params.documentKey,
            workbookKey: params.projectId,
            generation: params.generation,
            source,
            moduleName: params.moduleName,
            moduleType: params.moduleType,
            // The semantic rules read the kind, not the type: without this, a class module is
            // analysed as a standard one and every Me, Friend, and event declaration in it is
            // reported as an error. The extension's own client always sends both.
            moduleKind: moduleKindFromType(params.moduleType),
            documentType: params.documentType,
            severityOverrides: params.severityOverrides,
            activeIncompleteExpressionOffset: params.activeIncompleteExpressionOffset,
        };

        const response = this.analysis.handle(request);
        if (!response) {
            throw new RpcError(ErrorCode.InternalError, 'The analyzer returned nothing.');
        }

        if (response.kind === 'needSeed') {
            // The engine is not holding the sources this analysis was asked about. Answering from
            // what it does hold would report findings against text the user is not looking at.
            throw new RpcError(
                ErrorCode.ProjectNotSeeded,
                `No current sources for project '${response.workbookKey}'. Send project/open first.`,
                { projectId: response.workbookKey },
            );
        }

        if (response.kind === 'error') {
            throw new RpcError(ErrorCode.InternalError, response.message);
        }

        const answer = this.positioned(response.diagnostics, source, response.incrementalMode);

        // Kept whole for quick fixes, which need the parts of a finding that do not travel below,
        // and now for the comparison above as well.
        if (key !== undefined) {
            this.lastAnalysis.set(key, {
                source,
                diagnostics: response.diagnostics,
                request,
                facts,
                shape,
                caret,
                mode: response.incrementalMode,
                answer,
            });
        }

        return answer;
    }

    /**
     * Findings with line and column added, counted in `source`.
     *
     * Here rather than inline because this is the only place that knows which text the offsets
     * were counted in - the caller may not have sent one, in which case the choice between the
     * live copy and the seeded one was made inside and the caller cannot see it. A finding
     * measured in one text and drawn in another is how a squiggle lands on the wrong line.
     */
    private positioned(
        diagnostics: readonly VbaModuleAnalysisDiagnostic[],
        source: string,
        mode: DiagnosticsResult['mode'],
    ): DiagnosticsResult {
        // NOTHING TO POSITION, NOTHING TO BUILD. Most modules are clean most of the time, and the
        // table is a walk of the whole module: on the 64,802-line one that was a 1.5 MB scan per
        // pass, per clean module, to place no findings at all.
        if (diagnostics.length === 0) {
            return { diagnostics: [], mode };
        }

        // Built once for the module: per finding it would be a scan of the whole text each time.
        const starts = lineStarts(source);

        return {
            diagnostics: diagnostics.map((diagnostic) => {
                const start = toLineColumn(starts, diagnostic.span.start);
                const end = toLineColumn(starts, diagnostic.span.end);

                return {
                    code: diagnostic.code,
                    message: diagnostic.message,
                    severity: diagnostic.severity,
                    span: { start: diagnostic.span.start, end: diagnostic.span.end },
                    at: {
                        startLine: start.line,
                        startColumn: start.column,
                        endLine: end.line,
                        endColumn: end.column,
                    },
                };
            }),
            mode,
        };
    }

    private requireInitialized(): void {
        if (!this.initialized) {
            throw new RpcError(ErrorCode.InvalidRequest, 'initialize has not been called.');
        }
    }

    private require<T>(params: unknown): T {
        if (params === null || typeof params !== 'object') {
            throw new RpcError(ErrorCode.InvalidParams, 'This method requires parameters.');
        }

        return params as T;
    }
}
