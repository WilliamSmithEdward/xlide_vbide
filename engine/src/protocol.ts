// Wire protocol between the add-in and this engine.
//
// JSON-RPC 2.0, one object per line. Newline framing rather than length headers because every
// message is a single line of JSON and a reader that splits on newlines cannot desynchronise on a
// miscounted byte.
//
// Positions are UTF-16 character offsets into a module's source, not line and column. That is the
// analyzer's own currency, so the conversion happens once, at the add-in boundary, where the line
// and column the editor reports are turned into an offset. Doing it here instead would mean
// carrying a line index for every module on both sides of the pipe.

/** A request from the add-in. */
export interface JsonRpcRequest {
    jsonrpc: '2.0';
    /** Absent for notifications, which are not answered. */
    id?: number | string;
    method: string;
    params?: unknown;
}

/** A response to a request that carried an id. */
export interface JsonRpcResponse {
    jsonrpc: '2.0';
    id: number | string;
    result?: unknown;
    error?: JsonRpcError;
}

export interface JsonRpcError {
    code: number;
    message: string;
    data?: unknown;
}

/** Standard JSON-RPC codes, plus the one case this engine adds. */
export const ErrorCode = {
    ParseError: -32700,
    InvalidRequest: -32600,
    MethodNotFound: -32601,
    InvalidParams: -32602,
    InternalError: -32603,
    /** The engine has no current source for the project, or a stale one. Send project/open again. */
    ProjectNotSeeded: -32000,
} as const;

/** One module as the add-in reads it out of the editor. */
export interface ModulePayload {
    moduleName: string;
    source: string;
    /** Host component type: standard, class, document, or userform. */
    type?: string;
    /** For document modules, which document kind, so event handlers resolve. */
    documentType?: string;
}

/** project/open: replaces everything the engine knows about a project. */
export interface ProjectOpenParams {
    projectId: string;
    /**
     * Increases whenever the module set or any module's text changes. The engine refuses analysis
     * carrying a generation it has not been seeded with, rather than answering from stale sources.
     */
    generation: number;
    modules: ModulePayload[];
}

/** textDocument/diagnostics: analyse one module. */
export interface DiagnosticsParams {
    /** Stable key for the module being edited, used to hold incremental state. */
    documentKey: string;
    projectId?: string;
    generation?: number;
    source: string;
    moduleName: string;
    moduleType?: string;
    documentType?: string;
    /** Per-rule severity overrides keyed by diagnostic code. */
    severityOverrides?: Record<string, string>;
}

/** A single finding. Spans are UTF-16 offsets into the analysed source. */
export interface DiagnosticPayload {
    code?: string;
    message: string;
    severity: string;
    span: { start: number; end: number };
}

export interface DiagnosticsResult {
    diagnostics: DiagnosticPayload[];
    /** Whether the analyser reused per-procedure work or ran a full pass. */
    mode?: 'full' | 'incremental';
}
