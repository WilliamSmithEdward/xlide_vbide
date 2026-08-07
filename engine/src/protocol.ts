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
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    moduleName: string;
    moduleType?: string;
    documentType?: string;
    /** Per-rule severity overrides keyed by diagnostic code. */
    severityOverrides?: Record<string, string>;
    /**
     * The caret, when the module is the one being typed in. The analyzer holds back the
     * transient complaints of an expression mid-edit at this position, so typing does not
     * squiggle against itself.
     */
    activeIncompleteExpressionOffset?: number;
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

/**
 * textDocument/completion: what can be typed at an offset.
 *
 * The source travels with the request because the developer is mid-keystroke: the engine's seeded
 * copy of this module is as old as the last write-back, and completions against that would offer
 * members of a receiver that no longer exists on the line. The other modules' seeded sources are
 * current enough for the project facts they contribute.
 */
export interface CompletionParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offset of the caret into that source. */
    offset: number;
    moduleType?: string;
    documentType?: string;
}

/** One completion. The kind is the analyzer's own vocabulary; the surface maps it to icons. */
export interface CompletionItemPayload {
    label: string;
    kind: string;
    detail?: string;
    documentation?: string;
    /** May contain editor snippet placeholders; the surface inserts it as a snippet when it does. */
    insertText?: string;
    filterText?: string;
    sortText?: string;
}

export interface CompletionResult {
    items: CompletionItemPayload[];
}

/**
 * textDocument/hover: describe the identifier at an offset. The source travels with the request
 * for the same liveness reason completion's does.
 */
export interface HoverParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offset into that source. */
    offset: number;
    moduleType?: string;
    documentType?: string;
}

/** A resolved hover. The span is UTF-16 offsets into the request's source. */
export interface HoverPayload {
    /** Declaration line, rendered as VBA code. */
    signature: string;
    /** Plain-text facts: origin module, visibility, source note. */
    details: string[];
    /** Markdown documentation when any is known. */
    documentation?: string;
    span: { start: number; end: number };
}

export interface HoverResult {
    hover: HoverPayload | null;
}

/**
 * textDocument/signatureHelp: the call tip for the argument list the caret is inside. Same
 * liveness rule as completion.
 */
export interface SignatureHelpParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offset into that source. */
    offset: number;
    moduleType?: string;
    documentType?: string;
}

/** One parameter slot, its label exactly as it appears in the signature line. */
export interface SignatureParameterPayload {
    label: string;
    documentation?: string;
}

/** A resolved call tip: the signature line and which parameter is active. */
export interface SignatureInfoPayload {
    label: string;
    parameters: SignatureParameterPayload[];
    activeParameter: number;
    documentation?: string;
    details?: string[];
}

export interface SignatureHelpResult {
    signature: SignatureInfoPayload | null;
}

/** A text replacement, offsets into the request's source. An insertion has start === end. */
export interface TextEditPayload {
    start: number;
    end: number;
    text: string;
}

/**
 * textDocument/smartEnter: what Enter should leave behind. The source is the text just after
 * the newline went in; the offset names where it went in, which is the end of the line Enter
 * was pressed on. Same liveness rule as completion.
 */
export interface SmartEnterParams {
    projectId: string;
    moduleName: string;
    /** The module text just after the newline, when sent; the live copy otherwise. */
    source?: string;
    /** UTF-16 offset at which the newline was inserted. */
    offset: number;
    moduleType?: string;
    documentType?: string;
}

export interface SmartEnterResult {
    edits: TextEditPayload[];
    /** Where the caret belongs in the text as it stands after the edits; null with no edits. */
    caret: number | null;
}

/**
 * textDocument/canonicalCase: the case corrections for a span. The single form corrects only
 * the identifier ending at the span's end. completeHeader also finishes a bare procedure
 * header's parentheses, and is only meaningful when the span is a whole line.
 */
export interface CanonicalCaseParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    start: number;
    end: number;
    single?: boolean;
    completeHeader?: boolean;
    moduleType?: string;
    documentType?: string;
}

export interface CanonicalCaseResult {
    edits: TextEditPayload[];
}

/**
 * textDocument/loopSync: when the edit at an offset touched a simple For/For Each iterator or
 * its Next name, the paired rename for the other side.
 */
export interface LoopSyncParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offset just after the edit. */
    offset: number;
    moduleType?: string;
    documentType?: string;
}

export interface LoopSyncResult {
    edits: TextEditPayload[];
}

/**
 * textDocument/codeAction: the quick fixes offered over a span.
 *
 * A span rather than an offset, because the surface asks about a selection as readily as a caret,
 * and an empty selection is a span whose ends meet. No diagnostics travel with the request: the
 * engine resolves fixes from the analysis it holds, which carries the fix data the surface never
 * saw. Same liveness rule as completion — the source of the module being typed in travels with
 * the request, since a fix must edit the text on screen and not the text last written back.
 */
export interface CodeActionParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offsets into that source. An empty selection has start === end. */
    start: number;
    end: number;
    moduleType?: string;
    documentType?: string;
}

/** One quick fix: what to call it, the finding it answers, and the edits that apply it. */
export interface CodeActionPayload {
    title: string;
    /** Whether the editor should offer this one first. */
    isPreferred?: boolean;
    /** The diagnostic code it answers, so the surface can attach it to the right squiggle. */
    code?: string;
    /** The finding's own span, for the same reason. */
    span: { start: number; end: number };
    edits: TextEditPayload[];
}

export interface CodeActionResult {
    actions: CodeActionPayload[];
}

/**
 * textDocument/definition and textDocument/references: where the identifier at an offset is
 * declared, and everywhere else it is used.
 *
 * Both are answered within one workbook and never across two. Two open workbooks can each hold a
 * Module1 and a Recalculate, and they are unrelated. Same liveness rule as completion for the
 * module being typed in; the other modules answer from the live text the engine holds for them,
 * because the answer is a line the surface will scroll to and the surface shows live text.
 */
export interface NavigationParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offset into that source. */
    offset: number;
    moduleType?: string;
    documentType?: string;
    /** references only: whether the declaration itself is one of the answers. */
    includeDeclaration?: boolean;
}

/** One place in the workbook: which module, and a 1-based line and column into its live text. */
export interface LocationPayload {
    module: string;
    line: number;
    column: number;
    length: number;
    /**
     * The line the reference sits on, trimmed for a results list.
     *
     * Carried because the surface's own references list renders TEXT rather than editor models,
     * and only modules with a tab open have a model. Sending the line is what lets a reference in
     * a module nobody has opened be shown at all — which is the reference most worth showing.
     */
    preview?: string;
}

export interface NavigationResult {
    locations: LocationPayload[];
}

/**
 * textDocument/rename: the new text of every module a rename changes, across the workbook and
 * never past it.
 *
 * Whole module texts rather than edits, because the add-in writes modules and a module with no
 * tab open has no editor to apply an edit list to. The engine holds every module's current text
 * anyway, so producing the result of the edit is free and leaves nothing for two sides to
 * disagree about.
 */
export interface RenameParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    /** UTF-16 offset of the symbol being renamed. */
    offset: number;
    newName: string;
    moduleType?: string;
    documentType?: string;
}

/** One module a rename rewrites: its name, and what it says afterwards. */
export interface RenamedModulePayload {
    module: string;
    source: string;
    /** How many occurrences were replaced in it, for the summary the surface shows. */
    replaced: number;
}

export interface RenameResult {
    modules: RenamedModulePayload[];
    /** The name as it stood, so the surface can say what it renamed. */
    oldName?: string;
    /** Present when nothing was renamed, saying why in words a developer can act on. */
    refused?: string;
    /**
     * Present when what was renamed is a MODULE rather than a symbol inside one. Its name lives
     * on the component rather than in any module's text, so the add-in has that half to do; the
     * texts here are what the rest of the workbook must say once it has.
     */
    module?: string;
    /**
     * Uses of the old name the rename deliberately did not touch, because another module declares
     * the same name and nothing can prove which one a bare call meant. Reported rather than
     * silently skipped: the developer is the only one who knows, and they cannot decide about a
     * call nobody told them about.
     */
    ambiguous?: LocationPayload[];
}

/**
 * workspace/renameModule: the new text of every module that mentions a module being renamed.
 *
 * The module's own name is not in its text — it belongs to the component, which the add-in
 * renames through the object model. This works out only what the OTHER modules must say
 * afterwards, so the add-in can compute everything before it changes anything.
 */
export interface RenameModuleParams {
    projectId: string;
    /** The module being renamed, by its current name. */
    moduleName: string;
    newName: string;
}

/**
 * textDocument/semanticTokens: the analysed colouring of a whole module — which identifiers name
 * types, which kind of type each names, and which are host globals nothing has shadowed. Same
 * liveness rule as completion.
 */
export interface SemanticTokensParams {
    projectId: string;
    moduleName: string;
    /** The module text, when sent; the engine's live copy from didChange otherwise. */
    source?: string;
    moduleType?: string;
    documentType?: string;
}

/**
 * One coloured span. The type is the analyzer's own vocabulary — class, enum, struct, type,
 * variable — and the modifiers are the standard semantic-token ones, of which only
 * `defaultLibrary` is used, for host globals. Offsets into the request's source.
 */
export interface SemanticTokenPayload {
    start: number;
    end: number;
    type: string;
    modifiers?: string[];
}

export interface SemanticTokensResult {
    /** In position order: the surface encodes them as deltas, so the order is the contract. */
    tokens: SemanticTokenPayload[];
}

/**
 * textDocument/outline: a module's procedures, in declaration order. The source is optional:
 * present for the module being edited (the liveness rule), absent to use the seeded copy, which
 * is how the tree asks about modules that are not open.
 */
export interface OutlineParams {
    projectId: string;
    moduleName: string;
    source?: string;
    moduleType?: string;
    documentType?: string;
}

/** One procedure: the kind as the tree spells it ("Sub", "Property Get"), and its 1-based line. */
export interface OutlineProcedure {
    name: string;
    kind: string;
    line: number;
}

export interface OutlineResult {
    procedures: OutlineProcedure[];
}

/**
 * textDocument/didChange: the live text of a module, kept engine-side so requests can carry an
 * offset and nothing else. A full source replaces; edits apply to what is held, offsets into
 * the text as it stood, in the bottom-up order the editor reports them. A notification: no id,
 * no answer, ordered with everything else on the pipe.
 */
export interface DidChangeParams {
    projectId: string;
    moduleName: string;
    source?: string;
    edits?: TextEditPayload[];
}

/**
 * workspace/search: find text across the modules the engine holds. Scope narrows by project
 * and module; the engine searches its live text where it has one, seeded text otherwise.
 */
export interface SearchParams {
    scope: 'module' | 'project' | 'all';
    projectId?: string;
    module?: string;
    query: string;
    matchCase?: boolean;
    wholeWord?: boolean;
}

/** One hit: where, and the line it sits on, trimmed for the results list. */
export interface SearchMatchPayload {
    projectId: string;
    module: string;
    line: number;
    column: number;
    length: number;
    preview: string;
}

export interface SearchResult {
    matches: SearchMatchPayload[];
    truncated: boolean;
}
