// Maps the wire protocol onto the analyzer.
//
// The analyzer already has a request handler built for exactly this shape: a pure, synchronous
// object that owns project state and answers analysis requests, written so everything crossing its
// boundary is plain data. It was built to sit behind a worker thread's message port. A pipe is the
// same contract with a different transport, so it is reused rather than reimplemented, which keeps
// one analysis path shared with the editor extension instead of two that can disagree.

import { AnalysisWorkerState } from '../../../xlide_vscode/src/analysisWorkerLogic';
import type { AnalysisWorkerRequest } from '../../../xlide_vscode/src/analysisWorkerProtocol';
import { completionsFor } from './completion';
import { hoverFor } from './hover';
import {
    ErrorCode,
    type CompletionParams,
    type CompletionResult,
    type DiagnosticsParams,
    type DiagnosticsResult,
    type HoverParams,
    type HoverResult,
    type JsonRpcError,
    type ModulePayload,
    type ProjectOpenParams,
} from './protocol';

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

            default:
                throw new RpcError(ErrorCode.MethodNotFound, `Unknown method: ${method}`);
        }
    }

    private openProject(params: ProjectOpenParams): { modules: number } {
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
        return { modules: params.modules.length };
    }

    private closeProject(params: { projectId: string }): null {
        this.generations.delete(params.projectId);
        this.seededModules.delete(params.projectId);
        return null;
    }

    private completion(params: CompletionParams): CompletionResult {
        this.requireInitialized();

        // Not gated on generation: the request carries the live source of the module being typed
        // in, and the seeded copies of the others are current enough for the facts they lend.
        return { items: completionsFor(this.seededModules.get(params.projectId) ?? [], params) };
    }

    private hover(params: HoverParams): HoverResult {
        this.requireInitialized();

        // Same liveness rule as completion.
        return { hover: hoverFor(this.seededModules.get(params.projectId) ?? [], params) };
    }

    private diagnostics(params: DiagnosticsParams): DiagnosticsResult {
        this.requireInitialized();

        const request: AnalysisWorkerRequest = {
            kind: 'analyze',
            requestId: this.nextRequestId++,
            docKey: params.documentKey,
            workbookKey: params.projectId,
            generation: params.generation,
            source: params.source,
            moduleName: params.moduleName,
            moduleType: params.moduleType,
            documentType: params.documentType,
            severityOverrides: params.severityOverrides,
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
