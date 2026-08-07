// Maps the wire protocol onto the analyzer.
//
// The analyzer already has a request handler built for exactly this shape: a pure, synchronous
// object that owns project state and answers analysis requests, written so everything crossing its
// boundary is plain data. It was built to sit behind a worker thread's message port. A pipe is the
// same contract with a different transport, so it is reused rather than reimplemented, which keeps
// one analysis path shared with the editor extension instead of two that can disagree.

import { AnalysisWorkerState } from '../../../xlide_vscode/src/analysisWorkerLogic';
import type { AnalysisWorkerRequest } from '../../../xlide_vscode/src/analysisWorkerProtocol';
import type { VbaModuleAnalysisDiagnostic } from '../../../xlide_vscode/src/vbaModuleAnalysis';
import { moduleKindFromType } from '../../../xlide_vscode/src/vbaProjectAnalysis';
import { codeActionsFor } from './codeActions';
import { completionsFor } from './completion';
import { outlineFor, projectWordsFor } from './outline';
import { searchModules } from './search';
import { hoverFor } from './hover';
import { canonicalCaseFor, loopSyncFor, smartEnterFor } from './onType';
import { assembleSymbols, definitionsFor, referencesFor, type ProjectSymbols } from './navigation';
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
    type SearchParams,
    type SearchResult,
    type SemanticTokensParams,
    type SemanticTokensResult,
    type SignatureHelpParams,
    type SignatureHelpResult,
    type SmartEnterParams,
    type SmartEnterResult,
} from './protocol';

function liveKey(projectId: string, moduleName: string): string {
    return `${projectId}\0${moduleName.toLowerCase()}`;
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
     * string instance per keystroke, shared by every feature that asks about the module — which
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
     * The last colouring answered per module, keyed by the exact text it described — the same
     * bargain the outline strikes, and for a stronger reason: the surface re-asks for the whole
     * module's tokens after every edit, and two passes over 26,000 lines per keystroke is the
     * kind of cost that shows up as the editor feeling slow rather than as anything visible.
     */
    private readonly semanticMemo = new Map<string, { source: string; result: SemanticTokensResult }>();

    /** One workbook's symbol index, kept against the exact module texts it was built from. */
    private readonly symbolsMemo = new Map<string, { sources: string[]; symbols: ProjectSymbols }>();

    /**
     * The last analysis of each module, kept whole: the findings as the analyzer made them, and
     * the text they describe. Quick fixes are resolved from these rather than from anything the
     * surface sends back, because a finding carries fix data — the missing argument's name, the
     * unclosed block's expected closer — that never crosses to the surface at all.
     *
     * The request is kept too, so a fix asked for against text the diagnostics have not caught up
     * with can be answered by analysing that text under the same options the squiggles used.
     */
    private readonly lastAnalysis = new Map<string, {
        source: string;
        diagnostics: readonly VbaModuleAnalysisDiagnostic[];
        request: AnalyzeRequest;
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

            case 'textDocument/outline':
                return this.outline(this.require<OutlineParams>(params));

            case 'workspace/search':
                return this.search(this.require<SearchParams>(params));

            case 'textDocument/didChange':                return this.didChange(this.require<DidChangeParams>(params));

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

        // The project's own words, for the surface's tokenizer: names that are types and names
        // that are procedures. This is what lets `ROneCOne.Create(...)` read as a type and a
        // call while `values(index, 1)` stays a variable — the distinction the extension makes
        // with its semantic tokens.
        const facts = projectWordsFor(params.modules);
        return { modules: params.modules.length, types: facts.types, procedures: facts.procedures };
    }

    private closeProject(params: { projectId: string }): null {
        this.generations.delete(params.projectId);
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

        const memo = this.symbolsMemo.get(projectId);
        if (memo
            && memo.sources.length === modules.length
            && memo.sources.every((held, index) => held === modules[index].source)) {
            return memo.symbols;
        }

        const symbols = assembleSymbols(modules);
        this.symbolsMemo.set(projectId, { sources: modules.map((module) => module.source), symbols });
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

        // Kept whole for quick fixes, which need the parts of a finding that do not travel below.
        if (params.projectId !== undefined) {
            this.lastAnalysis.set(liveKey(params.projectId, params.moduleName), {
                source,
                diagnostics: response.diagnostics,
                request,
            });
        }

        return {
            diagnostics: response.diagnostics.map((diagnostic) => ({
                code: diagnostic.code,
                message: diagnostic.message,
                severity: diagnostic.severity,
                span: { start: diagnostic.span.start, end: diagnostic.span.end },
            })),
            mode: response.incrementalMode,
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
