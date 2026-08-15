import * as monaco from "monaco-editor/editor/editor.api.js";
import { DocumentStore, docKeyOf, type DocumentId } from "./documents.js";
import type { ExplorerProject } from "./explorer.js";
import { setInstallPath } from "./helpdialog.js";
import type { MenuItem } from "./menubar.js";
import type { Shell, ShellFinding, ShellProperty } from "./shell.js";
import type { SearchWidget } from "./searchwidget.js";
import type { ToolbarCommand } from "./toolbar.js";
import type { Workspace } from "./workspace.js";
import { takeFormattingMark } from "./format.js";
import { applySettings, type EditorSettings, type IncomingSettings } from "./settings.js";
import { type XlideTheme } from "./theme.js";
import { updateVbaLanguageFacts } from "./vba.js";

/*
 * Position convention
 * -------------------
 * Every line and column that crosses this boundary is 1-BASED, in both directions.
 * Monaco is also 1-based for lines (lineNumber) and 1-based for columns (column: column 1 is
 * before the first character), so the host numbers map straight onto monaco.IRange with no
 * arithmetic. The conversion helpers below exist to make that explicit and to keep the
 * clamping in one place, not to shift any origin.
 *
 * An end position is exclusive in the usual text-range sense: a range that starts and ends at
 * the same position is an insertion point.
 */

export interface HostRange {
  /** 1-based. */
  startLine: number;
  /** 1-based; column 1 is before the first character on the line. */
  startColumn: number;
  /** 1-based. */
  endLine: number;
  /** 1-based. */
  endColumn: number;
}

export interface HostTextChange extends HostRange {
  text: string;
}

export type HostSeverity = "error" | "warning" | "info" | "hint";

export interface HostMarker extends HostRange {
  severity: HostSeverity;
  message: string;
  code?: string;
}

export type HostMessage =
  | { type: "openDocument"; moduleName: string; project?: string | null; text: string }
  | { type: "clearDocuments" }
  | { type: "syncDocument"; moduleName: string; project?: string | null; text: string }
  | { type: "notice"; text: string; sticky?: boolean }
  | { type: "editorCommand"; id: string }
  | { type: "immediateResult"; text: string; failed: boolean }
  | { type: "setModules"; modules: string[]; projects?: (string | null)[]; active: string | null; activeProject?: string | null; dirty?: boolean[]; faces?: (string | null)[]; activeFace?: string | null }
  | { type: "setFindings"; findings: ShellFinding[] }
  | { type: "setProjects"; projects: ExplorerProject[] }
  | { type: "setDiagnostics"; moduleName: string; project?: string | null; markers: HostMarker[] }
  | { type: "setCurrentLine"; line: number | null }
  | { type: "setBreakpoints"; lines: number[] }
  | { type: "confirmClose"; name: string; project?: string | null }
  | { type: "formMarkup"; moduleName: string; project?: string | null; markup: string | null; reason: string | null; form?: FormMarkupBox | null; controls?: FormMarkupControl[] | null }
  | { type: "formMarkupApplied"; moduleName: string; project?: string | null; ok: boolean; added: string[]; removed: string[]; set: number; refused?: string | null }
  | { type: "formMarkupLint"; moduleName: string; project?: string | null; findings: FormMarkupLintFinding[]; draftForm?: FormMarkupBox | null; draft?: FormMarkupControl[] | null }
  | { type: "designerApplySave"; moduleName: string; project?: string | null }
  | { type: "revealLine"; line: number }
  | { type: "setCaret"; line: number; column: number }
  | { type: "setMenu"; path: number[]; items: MenuItem[] }
  | { type: "setChrome"; menuBar: boolean }
  | { type: "setInstallPath"; path: string | null }
  | { type: "setProperties"; component: string; kind: string; properties: ShellProperty[] }
  | { type: "completionResult"; id: number; items: HostCompletionItem[] }
  | { type: "hoverResult"; id: number; hover: HostHoverPayload | null }
  | { type: "signatureHelpResult"; id: number; signature: HostSignatureInfo | null }
  | { type: "canonicalCaseResult"; id: number; edits: HostTextEdit[] }
  | { type: "codeActionResult"; id: number; actions: HostCodeAction[] }
  | { type: "semanticTokensResult"; id: number; tokens: HostSemanticToken[]; failed?: boolean }
  | { type: "navigationResult"; id: number; locations: HostLocation[] }
  | { type: "renameResult"; id: number; oldName?: string | null; newName?: string | null; modules: string[]; replaced: number; refused?: string | null }
  | { type: "outlineResult"; id: number; procedures: HostProcedure[]; failed?: boolean }
  | { type: "syncResult"; id: number; json: string }
  | { type: "setLanguageFacts"; types: string[]; procedures: string[] }
  | { type: "setLocals"; stopped: boolean; context: string | null; rows: { expression: string; value: string; kind: string }[] }
  | { type: "setWatches"; stopped: boolean; rows: { expression: string; value: string; kind: string; context: string }[] }
  | { type: "setDebugState"; mode: string }
  | { type: "obLibrariesResult"; id: number; libraries: ObLibrary[] }
  | { type: "obTypesResult"; id: number; types: ObType[] }
  | { type: "obMembersResult"; id: number; members: ObMember[] }
  | { type: "searchResult"; id: number; matches: HostSearchMatch[]; truncated: boolean; replaced?: number }
  | ({ type: "setSettings" } & IncomingSettings);

/**
 * One library the Object Browser lists: a referenced type library, or an open workbook's
 * project - the kind says which, and only a project's members can be navigated to.
 */
export interface ObLibrary {
  name: string;
  description: string;
  kind: "project" | "library";
}

/** One browsable type of a library. */
export interface ObType {
  name: string;
  kind: string;
}

/**
 * One member of a type, its signature spelled the way VBA would. The line is where the
 * member lives in its module - meaningful only for project members, zero elsewhere.
 */
export interface ObMember {
  name: string;
  kind: string;
  signature: string;
  description: string;
  line: number;
}

/** The form's own box as the markup layer projects it. Bounds are points, like the markup's;
 * colours arrive as CSS, already through the system palette; insides are the designer's own
 * client area - the parity numbers. */
export interface FormMarkupBox {
  caption?: string | null;
  width?: number | null;
  height?: number | null;
  backColor?: string | null;
  foreColor?: string | null;
  insideWidth?: number | null;
  insideHeight?: number | null;
}

/**
 * One control of the form projection, flat with a parent NAME - the walk's shape. Bounds are
 * points relative to the parent's client area (MSForms' own model); the canvas composes them
 * by nesting rather than by arithmetic.
 */
export interface FormMarkupControl {
  type: string;
  name: string;
  caption?: string | null;
  left?: number | null;
  top?: number | null;
  width?: number | null;
  height?: number | null;
  parent?: string | null;
  fontName?: string | null;
  fontSize?: number | null;
  fontBold?: boolean | null;
  fontItalic?: boolean | null;
  backColor?: string | null;
  foreColor?: string | null;
  insideWidth?: number | null;
  insideHeight?: number | null;
  tabs?: string[] | null;
}

/** A form's projection as one answer: the markup text and the spec it prints. */
export interface FormMarkupPayload {
  markup: string | null;
  reason: string | null;
  form?: FormMarkupBox | null;
  controls?: FormMarkupControl[] | null;
}

/** One squiggle for the markup document: 1-based line, reason, "error" or "warning". */
export interface FormMarkupLintFinding {
  line: number;
  message: string;
  severity: string;
}

/** The DRAFT a lint round trip parsed out of the document: dialect fields only - the text
 * has no walk beside it - for the canvas to preview while the document is dirty. */
export interface FormMarkupDraft {
  form: FormMarkupBox | null;
  controls: FormMarkupControl[];
}

/** How an apply of the designer tab's document ended; the fresh projection follows it. */
export interface FormMarkupApplied {
  ok: boolean;
  added: string[];
  removed: string[];
  set: number;
  refused?: string | null;
}

/** One procedure in a module's outline: the kind as the tree spells it, and its 1-based line. */
export interface HostProcedure {
  name: string;
  kind: string;
  line: number;
}

/** One search hit as the host answers it; workbook is the display name, or null unsaid. */
export interface HostSearchMatch {
  workbook?: string | null;
  module: string;
  line: number;
  column: number;
  length: number;
  preview: string;
}
/** A text replacement, UTF-16 offsets into the live source; an insertion has start === end. */
export interface HostTextEdit {
  start: number;
  end: number;
  text: string;
}

/**
 * One quick fix from the host's engine: what to call it, the finding it answers, and the edits
 * that apply it. The code and span belong to the finding, so the fix can be attached to the
 * squiggle it belongs to rather than floating free of it.
 */
export interface HostCodeAction {
  title: string;
  isPreferred: boolean;
  code?: string | null;
  start: number;
  end: number;
  edits: HostTextEdit[];
}

/**
 * One coloured span from the host's engine, UTF-16 offsets into the live source. The type is the
 * analyzer's vocabulary; the only modifier it uses is `defaultLibrary`, for host globals.
 */
export interface HostSemanticToken {
  start: number;
  end: number;
  type: string;
  modifiers?: string[] | null;
}

/**
 * One place in the workbook: which module, its workbook when the host names one, and a 1-based
 * line and column into that module's live text.
 */
export interface HostLocation {
  module: string;
  workbook?: string | null;
  line: number;
  column: number;
  length: number;
  /** The line it sits on, so a module with no tab open can still be listed. */
  preview?: string | null;
}

/** What a rename did, or the reason it did nothing. */
export interface HostRenameAnswer {
  modules: string[];
  replaced: number;
  refused: string | null;
}

/** One parameter slot, its label exactly as it appears in the signature line. */
export interface HostSignatureParameter {
  label: string;
  documentation?: string | null;
}

/** A resolved call tip from the host's engine. */
export interface HostSignatureInfo {
  label: string;
  parameters: HostSignatureParameter[];
  activeParameter: number;
  documentation?: string | null;
  details?: string[] | null;
}

/** A resolved hover from the host's engine. Spans are UTF-16 offsets into the live source. */
export interface HostHoverPayload {
  signature: string;
  details: string[];
  documentation?: string | null;
  start: number;
  end: number;
}

/** One completion from the host's engine. The kind is the analyzer's own vocabulary. */
export interface HostCompletionItem {
  label: string;
  kind: string;
  detail?: string | null;
  documentation?: string | null;
  insertText?: string | null;
  filterText?: string | null;
  sortText?: string | null;
}

/**
 * Where the page's start-up time went, measured from navigation start.
 *
 * Reported rather than inferred. The surface sits over a code pane, so every millisecond here is
 * time the developer spends looking at a pane that has not been replaced yet, and the only way to
 * know which stage owns it is to have the page say so.
 */
export interface BootTimings {
  /** Fetching, parsing and evaluating the bundle: the module body runs at this mark. */
  scriptMs: number;
  /** Constructing the editor widget. */
  createMs: number;
  /** Everything, to the moment this message is posted. */
  totalMs: number;
  /** Build stamp of the bundle actually running, so a cached stale one is visible. */
  build?: string;
  /** When the bundle's bytes had fully arrived; scriptMs minus this is compile-and-run. */
  fetchMs?: number;
  /** Bytes that crossed the wire for the bundle. Zero means the browser's cache served it. */
  transferBytes?: number;
  /** When the document itself had arrived: everything before this is the browser starting. */
  htmlMs?: number;
  /** When the bundle's request left: the gap after htmlMs is parse-and-queue, not serving. */
  requestStartMs?: number;
}

export type ClientMessage =
  | { type: "ready"; timings?: BootTimings }
  /**
   * A change to a module's text. `source` names what made it when that was not the keyboard:
   * the host holds findings back about a line being typed, and a format that moves a single
   * line is indistinguishable from a keystroke by shape alone.
   */
  | { type: "contentChanged"; moduleName: string; project?: string; revision: number; changes: HostTextChange[]; fullLength: number; source?: "format" }
  | { type: "selectionChanged"; startLine: number; startColumn: number; endLine: number; endColumn: number }
  | { type: "breakpointToggleRequested"; line: number }
  | { type: "activateModule"; moduleName: string; project?: string; face?: string }
  | { type: "navigate"; module: string; line: number; column: number; project?: string }
  | { type: "command"; name: string }
  | { type: "evaluate"; text: string }
  | { type: "openExternal"; url: string }
  | { type: "requestDocument"; module: string; project?: string }
  | { type: "panel"; name: string; open: boolean }
  | { type: "menu"; path: number[] }
  | { type: "menuExecute"; path: number[] }
  | { type: "editProperty"; component: string; name: string; value: string }
  | { type: "selectComponent"; name: string }
  | { type: "closeModule"; name: string; project?: string; action?: string; face?: string }
  | { type: "requestFormMarkup"; module: string; project?: string }
  | { type: "applyFormMarkup"; module: string; project?: string; markup: string }
  | { type: "lintFormMarkup"; module: string; project?: string; markup: string }
  | { type: "insertComponent"; kind: number; project?: string }
  | { type: "removeComponent"; name: string; project?: string }
  | { type: "completion"; id: number; offset: number }
  | { type: "hover"; id: number; offset: number }
  | { type: "signatureHelp"; id: number; offset: number }
  | { type: "canonicalCase"; id: number; start: number; end: number; single?: boolean; completeHeader?: boolean }
  | { type: "codeAction"; id: number; start: number; end: number }
  | { type: "semanticTokens"; id: number; module: string; project?: string }
  | { type: "definition"; id: number; offset: number }
  | { type: "references"; id: number; offset: number; includeDeclaration: boolean }
  | { type: "rename"; id: number; offset: number; newName: string }
  | { type: "renameModule"; id: number; module: string; project?: string; newName: string }
  | { type: "undoRename"; id: number }
  | { type: "outline"; id: number; module: string; project?: string }
  | ({ type: "sync"; id: number; body: string } & Record<string, string | number>)
  | { type: "obLibraries"; id: number }
  | { type: "obTypes"; id: number; library: string }
  | { type: "obMembers"; id: number; library: string; typeName: string }
  | { type: "close" }
  | { type: "search"; id: number; query: string; matchCase: boolean; wholeWord: boolean; scope: string }
  | { type: "replaceAll"; id: number; query: string; matchCase: boolean; wholeWord: boolean; scope: string; replacement: string }
  // Spread rather than restated, and that is the whole fix. Written out member by member, this
  // shape fell one member short of the inbound one - syncEngine - and the compiler had no way to
  // notice, because neither shape referred to the other or to EditorSettings. The effect was not a
  // setting that could not be changed but a setting that got RESET: every other change posted a
  // payload with no syncEngine in it, and the host read the absence as the default.
  | ({ type: "updateSettings" } & EditorSettings)
  | { type: "trace"; text: string };

export interface HostTransport {
  post(message: ClientMessage): void;
  subscribe(handler: (message: HostMessage) => void): void;
}

interface WebView2Host {
  postMessage(message: unknown): void;
  addEventListener(type: "message", handler: (event: { data: unknown }) => void): void;
}

declare global {
  interface Window {
    chrome?: { webview?: WebView2Host };
  }
}

/** The owner the surface sets its squiggles under; anything else on a model is not ours. */
export const MARKER_OWNER = "xlide";

const SEVERITY: Record<HostSeverity, monaco.MarkerSeverity> = {
  error: monaco.MarkerSeverity.Error,
  warning: monaco.MarkerSeverity.Warning,
  info: monaco.MarkerSeverity.Info,
  hint: monaco.MarkerSeverity.Hint,
};


function fromMonacoRange(range: monaco.IRange): HostRange {
  return {
    startLine: range.startLineNumber,
    startColumn: range.startColumn,
    endLine: range.endLineNumber,
    endColumn: range.endColumn,
  };
}

/**
 * One pending host round trip's bookkeeping: allocate an id, arm the timeout that resolves the
 * fallback, remember the resolver, and settle or expire exactly once.
 *
 * This table was hand-copied ten times across the bridge - ten maps, ten counters, twelve
 * request bodies and ten resolve blocks all carrying the same eight lines - and a missed
 * clearTimeout or a wrong map in one copy is a leaked timer or a promise that resolves empty
 * after the budget, both of which read as a slow host rather than a bug (the audit's B15).
 * The timeout and the fallback stay per-call on purpose: 2s for the keystroke-riding asks, 8s
 * where the host thread legitimately stalls while a large module is shown, 30s for renames,
 * 120s for a sync that can include a person choosing a folder. Those budgets are deliberate
 * and must not collapse into one constant.
 *
 * The fallback is a function so every expiry resolves a FRESH value - the rename fallback is
 * an object a caller could hold onto, and sharing one instance across timeouts would let one
 * caller's reading leak into another's.
 */
class RequestTable<T> {
  private readonly waiting = new Map<number, {
    resolve: (value: T) => void;
    timer: ReturnType<typeof setTimeout>;
  }>();

  private nextId = 1;

  /** Sends one request via `post`, which receives the allocated id to put on the wire. */
  ask(fallback: () => T, timeoutMs: number, post: (id: number) => void): Promise<T> {
    const id = this.nextId++;

    return new Promise<T>((resolve) => {
      const timer = setTimeout(() => {
        this.waiting.delete(id);
        resolve(fallback());
      }, timeoutMs);

      this.waiting.set(id, { resolve, timer });
      post(id);
    });
  }

  /** Settles one id with its answer. An id already expired changes nothing, by design. */
  settle(id: number, value: T): void {
    const waiter = this.waiting.get(id);
    if (waiter) {
      this.waiting.delete(id);
      clearTimeout(waiter.timer);
      waiter.resolve(value);
    }
  }
}

export class EditorBridge {
  private readonly transport: HostTransport;

  /** The open documents, one live model each. Public: the workspace shows them. */
  readonly documents: DocumentStore;

  /**
   * The frame and the editor grid, assigned right after construction and before the ready
   * message lets the host speak. Assigned rather than constructor-taken because the three -
   * bridge, workspace, shell - reference each other, and the bridge is built first.
   */
  shell: Shell | null = null;
  workspace: Workspace | null = null;

  /** The floating search widget; assigned after construction, the same way openSettings is.
   * Search answers from the host route here - the widget owns the whole search UI. */
  searchWidget: SearchWidget | null = null;
  private readonly disposables: monaco.IDisposable[] = [];

  /**
   * Debug decorations - the stopped line and the breakpoint dots - ride the host-active
   * document's MODEL, not an editor: they are visible in whichever group shows that model,
   * and they survive the model moving between groups.
   */
  private currentLineDecor: { model: monaco.editor.ITextModel; ids: string[] } | null = null;
  private breakpointDecor: { model: monaco.editor.ITextModel; ids: string[] } | null = null;

  /** Lines that already carry a breakpoint, so the hover dot is not drawn over a real one. */
  private breakpointLines = new Set<number>();

  /**
   * The squiggles the host last sent, per document.
   *
   * Kept because replacing a model's text drags every marker to the end of the replacement.
   * Markers are anchored to positions in the text, and a whole-document edit is, as far as the
   * editor is concerned, the entire text being deleted and different text arriving: everything
   * anchored inside it collapses to one point. They are set again afterwards, at the positions the
   * host gave, which are still correct because the text either did not change or changed only in
   * ways that do not move lines.
   */
  private readonly markersByDoc = new Map<string, HostMarker[]>();

  /** Monotonic counters over locally originated edits, one per document. */
  private readonly revisions = new Map<string, number>();

  /** The document the host says is active - the one its native active pane shows. */
  private hostActive: DocumentId | null = null;

  // One pending-request table per host round trip; the three rename methods share one, the
  // way they always shared a counter. Budgets and fallbacks live at the call sites.
  private readonly pendingCompletions = new RequestTable<HostCompletionItem[]>();
  private readonly pendingHovers = new RequestTable<HostHoverPayload | null>();
  private readonly pendingSignatures = new RequestTable<HostSignatureInfo | null>();
  private readonly pendingCanonicalCases = new RequestTable<HostTextEdit[]>();
  private readonly pendingCodeActions = new RequestTable<HostCodeAction[]>();
  private readonly pendingSemanticTokens = new RequestTable<HostSemanticToken[] | null>();
  private readonly pendingNavigations = new RequestTable<HostLocation[]>();
  private readonly pendingRenames = new RequestTable<HostRenameAnswer>();
  private readonly pendingOutlines = new RequestTable<HostProcedure[] | null>();
  private readonly pendingSyncs = new RequestTable<Record<string, unknown>>();
  /** Echo suppression: true while a host edit is being written into the model. */
  private applyingHostEdit = false;
  /** Once the host names a theme, the OS preference stops overriding it. */

  constructor(
    transport: HostTransport,
    documents: DocumentStore = new DocumentStore(),
  ) {
    this.transport = transport;
    this.documents = documents;

    // Content listeners ride the MODEL, not the editor: an edit is attributed to the document
    // it changed, whichever editor made it, which is what lets several models live at once.
    this.documents.onModelCreated = (id, model) => {
      const key = docKeyOf(id.module, id.project);
      this.revisions.set(key, 0);
      model.onDidChangeContent((event) => this.onModelContentChanged(id, model, event));
    };

    this.documents.onModelClosing = (id, model) => {
      const key = docKeyOf(id.module, id.project);
      this.revisions.delete(key);
      this.markersByDoc.delete(key);
      if (this.currentLineDecor?.model === model) {
        this.currentLineDecor = null;
      }
      if (this.breakpointDecor?.model === model) {
        this.breakpointDecor = null;
      }
    };

    transport.subscribe((message) => this.handle(message));
  }

  /** The active group's editor, which is what "the editor" means bridge-wide. */
  private ed(): monaco.editor.IStandaloneCodeEditor | null {
    return this.workspace?.activeEditor() ?? null;
  }

  /**
   * Wires one group's editor into the bridge: the caret authority, the breakpoint margin,
   * and its hover preview. Called once per editor as the workspace creates groups.
   *
   * Only the ACTIVE group's caret reaches the host - the caret decides what a Run acts on,
   * and there is one native caret - and only the active group's margin toggles breakpoints,
   * because the toggle targets the host-active module and a background group's margin would
   * aim at the wrong one.
   */
  attachEditor(editor: monaco.editor.IStandaloneCodeEditor): void {
    const hover = editor.createDecorationsCollection([]);

    this.disposables.push(
      editor.onDidChangeCursorSelection((event) => {
        if (editor === this.ed()) {
          this.onSelectionChanged(event.selection);
        }
      }),
      editor.onMouseDown((event) => {
        if (editor === this.ed()) {
          this.onMouseDown(event);
        }
      }),
      editor.onMouseMove((event) => {
        if (editor === this.ed()) {
          this.onMouseMove(event, hover);
        } else {
          hover.clear();
        }
      }),
      editor.onMouseLeave(() => hover.clear()),
    );
  }

  start(timings?: BootTimings): void {
    this.transport.post(timings ? { type: "ready", timings } : { type: "ready" });
  }

  /** Asks the host to show a module. The tree names the workbook it means; a tab cannot yet.
   * face "design" asks for a form's designer tab instead of its code pane. */
  activateModule(moduleName: string, project?: string, face?: string): void {
    this.transport.post({
      type: "activateModule",
      moduleName,
      ...(project ? { project } : {}),
      ...(face ? { face } : {}),
    });
  }

  /**
   * Asks the host for a form's projection - the markup text and the spec it prints - and
   * routes the answer to whoever watches that form. The answer is a broadcast rather than a
   * reply: an apply, a later request, or another view of the same form all land here, and the
   * designer tab wants each one, not only the one it asked for.
   */
  requestFormMarkup(module: string, project: string | null): void {
    this.transport.post({ type: "requestFormMarkup", module, ...(project ? { project } : {}) });
  }

  /** Watches a form's projection answers; returns the unwatch. One watcher per form. */
  onFormMarkup(module: string, project: string | null, watcher: (payload: FormMarkupPayload) => void): () => void {
    const key = docKeyOf(module, project);
    this.formMarkupWatchers.set(key, watcher);
    return () => {
      if (this.formMarkupWatchers.get(key) === watcher) {
        this.formMarkupWatchers.delete(key);
      }
    };
  }

  private readonly formMarkupWatchers = new Map<string, (payload: FormMarkupPayload) => void>();

  /** Applies the designer tab's document to the form; the outcome and a fresh projection
   * come back through the watchers. */
  applyFormMarkup(module: string, project: string | null, markup: string): void {
    this.transport.post({ type: "applyFormMarkup", module, ...(project ? { project } : {}), markup });
  }

  /** Watches a form's apply outcomes; returns the unwatch. One watcher per form. */
  onFormMarkupApplied(module: string, project: string | null, watcher: (outcome: FormMarkupApplied) => void): () => void {
    const key = docKeyOf(module, project);
    this.formMarkupAppliedWatchers.set(key, watcher);
    return () => {
      if (this.formMarkupAppliedWatchers.get(key) === watcher) {
        this.formMarkupAppliedWatchers.delete(key);
      }
    };
  }

  private readonly formMarkupAppliedWatchers = new Map<string, (outcome: FormMarkupApplied) => void>();

  /** Asks the host to lint the document as it stands; findings return through the watcher. */
  lintFormMarkup(module: string, project: string | null, markup: string): void {
    this.transport.post({ type: "lintFormMarkup", module, ...(project ? { project } : {}), markup });
  }

  /** Watches a form's lint answers; returns the unwatch. One watcher per form. */
  onFormMarkupLint(module: string, project: string | null, watcher: (findings: FormMarkupLintFinding[], draft: FormMarkupDraft | null) => void): () => void {
    const key = docKeyOf(module, project);
    this.formMarkupLintWatchers.set(key, watcher);
    return () => {
      if (this.formMarkupLintWatchers.get(key) === watcher) {
        this.formMarkupLintWatchers.delete(key);
      }
    };
  }

  private readonly formMarkupLintWatchers = new Map<string, (findings: FormMarkupLintFinding[], draft: FormMarkupDraft | null) => void>();

  /** The host's Ctrl+S asking a designer tab to apply-then-save; keyed by the form. */
  onDesignerApplySave(module: string, project: string | null, watcher: () => void): () => void {
    const key = docKeyOf(module, project);
    this.designerApplySaveWatchers.set(key, watcher);
    return () => {
      if (this.designerApplySaveWatchers.get(key) === watcher) {
        this.designerApplySaveWatchers.delete(key);
      }
    };
  }

  private readonly designerApplySaveWatchers = new Map<string, () => void>();

  /** Tells the host which panel is showing, so it only watches what is being looked at. */
  panelChanged(name: string, open: boolean): void {
    this.transport.post({ type: "panel", name, open });
  }

  /** Asks the host to evaluate a line entered in the Immediate panel. */
  evaluate(text: string): void {
    this.transport.post({ type: "evaluate", text });
  }

  /**
   * Asks the host to open an address outside the surface.
   *
   * This page may not navigate anywhere, by policy - the whole point of it is that it is a fixed
   * document over a host, not a browser. The host has no such restriction and opens the address
   * in whatever the machine uses, but only if it is one of the few it holds: what is sent from
   * here is a request, not an instruction.
   */
  openExternal(url: string): void {
    this.transport.post({ type: "openExternal", url });
  }

  /**
   * Asks the host to go to a place: a finding, or a procedure picked in the tree. The host
   * shows the module and moves the native caret; the surface's own caret is placed here, the
   * moment the module is the one showing, so the click ends with the cursor at the target and
   * the editor focused, ready to type. A procedure asks for its whole line, so where the block
   * starts is visible at a glance.
   */
  navigate(module: string, line: number, column: number, selectLine = false, project?: string): void {
    this.pendingCaret = { module, project: project ?? null, line, column, selectLine };
    this.transport.post({ type: "navigate", module, line, column, ...(project ? { project } : {}) });
    this.applyPendingCaret();
  }

  /** A caret waiting for its module to be shown; applied when the active document matches. */
  private pendingCaret: { module: string; project: string | null; line: number; column: number; selectLine: boolean } | null = null;

  /**
   * Places the waiting caret if the shown module is the one it belongs to. A navigation that
   * named no workbook matches the module by name alone - a finding that could not say still
   * navigates - and one that named it must match both parts.
   */
  private applyPendingCaret(): void {
    const pending = this.pendingCaret;
    const model = this.model();
    if (!pending || !model) {
      return;
    }

    const shown = this.documents.idOf(model);
    if (!shown
      || shown.module.toLowerCase() !== pending.module.toLowerCase()
      || (pending.project !== null
        && (shown.project ?? "").toLowerCase() !== pending.project.toLowerCase())) {
      return;
    }

    this.pendingCaret = null;
    const editor = this.ed();
    if (!editor) {
      return;
    }

    const line = Math.min(Math.max(pending.line, 1), model.getLineCount());
    if (pending.selectLine) {
      editor.setSelection(new monaco.Selection(line, 1, line, model.getLineMaxColumn(line)));
    } else {
      const column = Math.min(Math.max(pending.column, 1), model.getLineMaxColumn(line));
      editor.setPosition({ lineNumber: line, column });
    }
    editor.revealLineInCenterIfOutsideViewport(line);
    editor.focus();
  }

  /** Asks the host for a menu's items; [] is the bar itself. */
  requestMenu(path: number[]): void {
    this.transport.post({ type: "menu", path });
  }

  /** Asks the host to run the menu item a position chain leads to. */
  executeMenu(path: number[]): void {
    this.transport.post({ type: "menuExecute", path });
  }

  /** Asks the host to write a property the developer edited. */
  editProperty(component: string, name: string, value: string): void {
    this.transport.post({ type: "editProperty", component, name, value });
  }

  /** Tells the host the explorer's selection changed, which the properties panel follows. */
  selectComponent(name: string): void {
    this.transport.post({ type: "selectComponent", name });
  }

  /** Asks the host to close a module's pane, which is what closes its tab. The host holds a
   * close whose module has unsaved changes and asks back with confirmClose; the developer's
   * choice returns through the same message as the action - "save" or "discard". face
   * "design" closes a form's designer tab, which has no unsaved-text question host-side. */
  closeModule(name: string, project?: string, action?: string, face?: string): void {
    this.transport.post({
      type: "closeModule",
      name,
      ...(project ? { project } : {}),
      ...(action ? { action } : {}),
      ...(face ? { face } : {}),
    });
  }

  /**
   * Asks the host to adopt and persist settings. The host echoes them back as setSettings once
   * written, and the echo is what the page applies: a choice is real when the file has it.
   */
  updateSettings(settings: EditorSettings): void {
    this.transport.post({ type: "updateSettings", ...settings });
  }

  /**
   * Asks the host for a new component: 1 module, 2 class module, 3 form. Named workbook when
   * the request came from a workbook's own menu; the active project otherwise.
   */
  insertComponent(kind: number, project?: string): void {
    this.transport.post({ type: "insertComponent", kind, ...(project ? { project } : {}) });
  }

  /**
   * Asks the host to remove a component, the developer having already been asked whether they
   * mean it. The page does not close the tab first: the host prunes it from the pane list it
   * republishes, so the tab goes because the component went, in that order and once.
   */
  removeComponent(name: string, project?: string): void {
    this.transport.post({ type: "removeComponent", name, ...(project ? { project } : {}) });
  }

  /** Asks the host to toggle the breakpoint on a line, same as clicking the margin. */
  toggleBreakpoint(line: number): void {
    this.transport.post({ type: "breakpointToggleRequested", line });
  }

  /**
   * Asks the host what can be typed at an offset. Resolves empty rather than rejecting when the
   * host is slow or gone, because a completion that fails is a list that does not open, not an
   * error anybody should see.
   */
  requestCompletions(offset: number): Promise<HostCompletionItem[]> {
    return this.pendingCompletions.ask(() => [], 2000, (id) =>
      this.transport.post({ type: "completion", id, offset }));
  }

  /**
   * Asks the host what the identifier at an offset is. Resolves null rather than rejecting when
   * the host is slow or gone: a hover that fails is a tooltip that does not appear, not an error.
   */
  requestHover(offset: number): Promise<HostHoverPayload | null> {
    return this.pendingHovers.ask(() => null, 2000, (id) =>
      this.transport.post({ type: "hover", id, offset }));
  }

  /**
   * Asks the host for the call tip at an offset. Resolves null rather than rejecting: a tip that
   * fails is a tip that does not show.
   */
  requestSignatureHelp(offset: number): Promise<HostSignatureInfo | null> {
    return this.pendingSignatures.ask(() => null, 2000, (id) =>
      this.transport.post({ type: "signatureHelp", id, offset }));
  }

  /**
   * Asks the host what can be fixed over a span. Resolves empty rather than rejecting: a fix that
   * fails is a lightbulb that does not appear, which is what an unfixable line looks like anyway.
   */
  requestCodeActions(start: number, end: number): Promise<HostCodeAction[]> {
    return this.pendingCodeActions.ask(() => [], 2000, (id) =>
      this.transport.post({ type: "codeAction", id, start, end }));
  }

  /**
   * Asks the host for the case corrections over a span. Resolves empty rather than rejecting: a
   * recase that fails is a line left as typed, and the next pass over it will ask again.
   */
  requestCanonicalCase(
    start: number,
    end: number,
    options: { single?: boolean; completeHeader?: boolean } = {},
  ): Promise<HostTextEdit[]> {
    return this.pendingCanonicalCases.ask(() => [], 2000, (id) =>
      this.transport.post({
        type: "canonicalCase",
        id,
        start,
        end,
        ...(options.single ? { single: true } : {}),
        ...(options.completeHeader ? { completeHeader: true } : {}),
      }));
  }

  /**
   * Asks the host for a module's procedures, for its node in the tree. Resolves null - never
   * empty - when no answer comes: a timeout is not a statement that the module has no
   * procedures, and the difference decides whether the tree keeps what it already shows. The
   * window is generous because the host thread legitimately stalls for seconds while a large
   * module is being shown, and the answer queued behind that stall is still a good answer.
   */
  /**
   * Asks the host what an import or export would do, or asks for it to be done.
   *
   * The arguments go through untouched and the answer comes back as the host's own JSON, because
   * the host answers this and the debug api's `sync` route from the SAME call. Shaping it here
   * would be a second opinion about what a plan is, and second opinions drift.
   *
   * Resolves to an object carrying `error` rather than rejecting: a folder that has gone away is
   * something the dialog says, not something that throws.
   */
  requestSync(args: Record<string, string>, body = ""): Promise<Record<string, unknown>> {
    // The budget is long because it covers reading every module of a project and, for a
    // browse, a developer deciding where to put them.
    return this.pendingSyncs.ask(
      () => ({ error: "the host did not answer in time" }),
      120000,
      (id) => this.transport.post({ type: "sync", id, body, ...args }));
  }

  requestOutline(module: string, project?: string): Promise<HostProcedure[] | null> {
    return this.pendingOutlines.ask(() => null, 8000, (id) =>
      this.transport.post({ type: "outline", id, module, ...(project ? { project } : {}) }));
  }

  /**
   * Asks the host where the symbol at an offset is declared, or everywhere in the workbook it is
   * used. Resolves empty rather than rejecting: navigation that fails is a click that does not
   * move the cursor.
   */
  requestNavigation(
    offset: number,
    references: boolean,
    includeDeclaration = true,
  ): Promise<HostLocation[]> {
    return this.pendingNavigations.ask(() => [], 8000, (id) =>
      this.transport.post(references
        ? { type: "references", id, offset, includeDeclaration }
        : { type: "definition", id, offset }));
  }

  /**
   * Asks the host to rename a symbol everywhere it is used in the workbook.
   *
   * The HOST does the renaming, not the page: modules with no tab open have no model to edit,
   * and they are exactly the ones a rename must not miss. So nothing comes back but a summary,
   * and the open tabs are refreshed by the ordinary document sync.
   */
  /**
   * Asks the host to put the last rename back.
   *
   * The host holds it, not the page: a rename edits several modules and the editor's undo stack
   * is per model, so Ctrl+Z in the module you are looking at would reverse that module's share
   * and leave every other one renamed - a half-renamed project, which is worse than no undo.
   */
  requestRenameUndo(): Promise<HostRenameAnswer> {
    return this.pendingRenames.ask(
      () => ({ modules: [], replaced: 0, refused: "The undo timed out, so nothing changed." }),
      30000,
      (id) => this.transport.post({ type: "undoRename", id }));
  }

  requestModuleRename(module: string, project: string | null, newName: string): Promise<HostRenameAnswer> {
    return this.pendingRenames.ask(
      () => ({ modules: [], replaced: 0, refused: "The rename timed out, so nothing changed." }),
      30000,
      (id) => this.transport.post({ type: "renameModule", id, module, newName, ...(project ? { project } : {}) }));
  }

  requestRename(offset: number, newName: string): Promise<HostRenameAnswer> {
    return this.pendingRenames.ask(
      () => ({ modules: [], replaced: 0, refused: "The rename timed out, so nothing changed." }),
      30000,
      (id) => this.transport.post({ type: "rename", id, offset, newName }));
  }

  /**
   * The variables in scope where execution is stopped, or empty when it is not stopped.
   *
   * The Locals panel's own reading, kept so a hover can answer with a value. Empty the moment
   * execution resumes, because a value from a frame that has gone is worse than no value: it
   * looks current and is not.
   */
  private stoppedLocals: { expression: string; value: string; kind: string }[] = [];

  /**
   * What the debugger says this name is worth right now, or null when nothing is stopped, the
   * name is not in scope, or the caret is not in the module execution stopped in.
   *
   * That last guard matters. The Locals panel holds ONE frame, and a name in some other module
   * can easily match one of its variables; showing that value on a hover in a procedure that is
   * not running would be a confident lie. The stopped module is the one the surface is showing,
   * because the host brings it forward when it breaks.
   */
  localValue(model: monaco.editor.ITextModel, name: string): { value: string; kind: string } | null {
    if (this.stoppedLocals.length === 0 || model !== this.hostActiveModel()) {
      return null;
    }

    const wanted = name.toLowerCase();
    const found = this.stoppedLocals.find((row) => row.expression.toLowerCase() === wanted);

    // A row with no value is a parent node in the panel's tree, not a reading.
    return found && found.value.length > 0 ? { value: found.value, kind: found.kind } : null;
  }

  /** The model a host location names, or null when that module has no tab open. */
  modelForLocation(location: HostLocation): monaco.editor.ITextModel | null {
    return this.documents.get(location.module, location.workbook ?? null);
  }

  /** Documents asked for and not yet arrived, so a second ask joins the first. */
  private readonly pendingDocuments = new Map<string, Promise<monaco.editor.ITextModel | null>>();

  /**
   * Makes sure the page has a module's text, asking the host if it does not.
   *
   * The page holds a module's text once it has been ACTIVATED - not because its pane is open - so
   * a workspace opened onto eight modules holds one. Anything that DRAWS a module without going
   * to it needs this: the editor's peek window resolves each result to a model, and with no model
   * it draws an empty window (2026-08-07).
   *
   * Resolves with whatever the page has when it settles, including null: a module that cannot be
   * read is one result that will not preview, not a reason to answer nothing.
   */
  ensureDocument(module: string, project: string | null): Promise<monaco.editor.ITextModel | null> {
    const held = this.documents.get(module, project);
    if (held) {
      return Promise.resolve(held);
    }

    const key = docKeyOf(module, project);
    const already = this.pendingDocuments.get(key);
    if (already) {
      return already;
    }

    const waiting = new Promise<monaco.editor.ITextModel | null>((resolve) => {
      const timer = setTimeout(() => {
        this.documentWaiters.delete(key);
        this.pendingDocuments.delete(key);
        resolve(this.documents.get(module, project));
      }, 4000);

      this.documentWaiters.set(key, () => {
        clearTimeout(timer);
        this.documentWaiters.delete(key);
        this.pendingDocuments.delete(key);
        resolve(this.documents.get(module, project));
      });

      this.transport.post({
        type: "requestDocument",
        module,
        ...(project ? { project } : {}),
      });
    });

    this.pendingDocuments.set(key, waiting);
    return waiting;
  }

  /** Callbacks waiting on a document's text, by document key. */
  private readonly documentWaiters = new Map<string, () => void>();

  /** Goes where a host location names, opening its module if it is not already open. */
  navigateTo(location: HostLocation): void {
    this.navigate(
      location.module,
      location.line,
      location.column,
      false,
      location.workbook ?? undefined);
  }

  /**
   * Asks the host to colour a model. Addressed by the model's own document, not by whichever is
   * host-active: the editor colours every model it is showing, and a split shows two.
   *
   * Resolves null rather than empty when the host cannot answer, so the caller can keep the
   * colouring already on screen - a module that suddenly loses its analysed colours reads as the
   * analysis having broken, which is exactly what it would be lying about.
   */
  requestSemanticTokens(model: monaco.editor.ITextModel): Promise<HostSemanticToken[] | null> {
    const shown = this.documents.idOf(model);
    if (!shown) {
      return Promise.resolve(null);
    }

    return this.pendingSemanticTokens.ask(() => null, 8000, (id) =>
      this.transport.post({
        type: "semanticTokens",
        id,
        module: shown.module,
        ...(shown.project ? { project: shown.project } : {}),
      }));
  }

  /** True while a host edit is being written into the model, so listeners can tell it from typing. */
  get isApplyingHostEdit(): boolean {
    return this.applyingHostEdit;
  }

  /**
   * Runs a toolbar command.
   *
   * An editor command is run here; a host command is sent on. The editor's caret has to reach the
   * host before a host command runs, because the host runs the procedure its own caret is in, and
   * it is sent by the same selection message the caret already produces.
   */
  runCommand(command: ToolbarCommand): void {
    if (command.target === "host") {
      this.transport.post({ type: "command", name: command.id });
      return;
    }

    const editor = this.ed();
    if (!editor) {
      return;
    }

    // Focus first. An editor action taken while the button has focus operates on an editor that
    // does not have it, and the ones that open a widget put it somewhere the developer cannot type.
    editor.focus();

    // Undo and redo are not actions. They are built into the editor rather than registered like
    // the rest, so looking them up finds nothing and they have to be triggered by name.
    if (command.id === "undo" || command.id === "redo") {
      editor.trigger("xlide", command.id, null);
      return;
    }

    editor.getAction(command.id)?.run();
  }

  dispose(): void {
    for (const disposable of this.disposables) {
      disposable.dispose();
    }
    this.disposables.length = 0;
  }

  handle(message: HostMessage): void {
    switch (message.type) {
      case "openDocument":
        this.openDocument(message.moduleName, message.project ?? null, message.text);
        return;
      case "clearDocuments":
        this.clearDocuments();
        return;
      case "syncDocument":
        this.syncDocument(message.moduleName, message.project ?? null, message.text);
        return;
      case "notice":
        this.shell?.notify(message.text, message.sticky === true);
        return;
      case "editorCommand":
        this.runEditorCommand(message.id);
        return;
      case "immediateResult":
        this.shell?.appendImmediate(message.text, message.failed ? "failed" : "result");
        return;
      case "setModules": {
        // The open list is the models' truth too: a pane closed anywhere disposes its model
        // here, undo history and all, the same moment its tab leaves the strip. Designer tabs
        // ride the list wearing their face; their keys differ from the code panes', so the
        // model store - which only ever holds code documents - is untouched by them.
        const open: DocumentId[] = message.modules.map((module, index) => ({
          module,
          project: (message.projects ?? [])[index] ?? null,
          ...((message.faces ?? [])[index] === "design" ? { face: "design" as const } : {}),
        }));
        this.documents.closeMissing(open.filter((id) => !id.face));

        this.hostActive = message.active
          ? {
            module: message.active,
            project: message.activeProject ?? null,
            ...(message.activeFace === "design" ? { face: "design" as const } : {}),
          }
          : null;

        this.workspace?.setOpen(open, message.dirty ?? [], this.hostActive);
        this.shell?.setActiveModule(message.active, message.activeProject ?? null);

        // A navigation whose module just became the shown one lands its caret now.
        this.applyPendingCaret();
        return;
      }
      case "setFindings":
        this.shell?.setFindings(message.findings);
        return;
      case "setProjects":
        this.shell?.setProjects(message.projects);
        return;
      case "setDiagnostics":
        this.setDiagnostics(message.moduleName, message.project ?? null, message.markers);
        return;
      case "setCurrentLine":
        this.setCurrentLine(message.line);
        return;
      case "setBreakpoints":
        this.setBreakpoints(message.lines);
        return;
      case "confirmClose":
        this.shell?.confirmClose(message.name, message.project ?? null);
        return;
      case "formMarkup":
        // Keyed by the FORM, not the tab: the answer describes the module, and whoever draws
        // it - today the designer tab's two halves - subscribes under the form's own key.
        this.formMarkupWatchers.get(docKeyOf(message.moduleName, message.project ?? null))?.(message);
        return;
      case "formMarkupApplied":
        this.formMarkupAppliedWatchers.get(docKeyOf(message.moduleName, message.project ?? null))?.(message);
        return;
      case "formMarkupLint":
        this.formMarkupLintWatchers.get(docKeyOf(message.moduleName, message.project ?? null))?.(
          message.findings,
          message.draft ? { form: message.draftForm ?? null, controls: message.draft } : null);
        return;
      case "designerApplySave":
        this.designerApplySaveWatchers.get(docKeyOf(message.moduleName, message.project ?? null))?.();
        return;
      case "revealLine":
        this.ed()?.revealLineInCenterIfOutsideViewport(message.line);
        return;
      case "setCaret":
        // The caret decides what an editor command acts on, and the host copies it into the
        // native pane before running one, so this is how anything outside the page aims a
        // Run or a Step at a particular procedure.
        this.ed()?.setPosition({ lineNumber: message.line, column: message.column });
        this.ed()?.revealLineInCenterIfOutsideViewport(message.line);
        return;
      case "setMenu":
        this.shell?.setMenu(message.path, message.items);
        return;
      case "setChrome":
        this.shell?.setMenuBarVisible(message.menuBar);
        return;
      case "setInstallPath":
        setInstallPath(message.path);
        return;
      case "setProperties":
        this.shell?.setProperties(message.component, message.kind, message.properties);
        return;
      case "completionResult":
        this.pendingCompletions.settle(message.id, message.items);
        return;
      case "hoverResult":
        this.pendingHovers.settle(message.id, message.hover);
        return;
      case "signatureHelpResult":
        this.pendingSignatures.settle(message.id, message.signature);
        return;
      case "canonicalCaseResult":
        this.pendingCanonicalCases.settle(message.id, message.edits);
        return;
      case "codeActionResult":
        this.pendingCodeActions.settle(message.id, message.actions);
        return;
      case "renameResult":
        this.pendingRenames.settle(message.id, {
          modules: message.modules,
          replaced: message.replaced,
          refused: message.refused ?? null,
        });
        return;
      case "navigationResult":
        this.pendingNavigations.settle(message.id, message.locations);
        return;
      case "semanticTokensResult":
        // A failed answer is a shrug, not a statement of colourlessness.
        this.pendingSemanticTokens.settle(message.id, message.failed ? null : message.tokens);
        return;
      case "syncResult": {
        let answer: Record<string, unknown>;
        try {
          answer = JSON.parse(message.json) as Record<string, unknown>;
        } catch {
          answer = { error: "the host's answer could not be read" };
        }
        this.pendingSyncs.settle(message.id, answer);
        return;
      }
      case "outlineResult":
        // A failed answer is a shrug, not a statement of emptiness.
        this.pendingOutlines.settle(message.id, message.failed ? null : message.procedures);
        return;
      case "setLanguageFacts":
        updateVbaLanguageFacts(message.types, message.procedures);
        return;
      case "searchResult":
        this.searchWidget?.showSearchResults(message.id, message.matches, message.truncated, message.replaced ?? 0);
        return;
      case "setLocals":
        // Kept as well as shown. The panel draws them; the hover uses them for data tips, which
        // is the same reading answering a different question.
        this.stoppedLocals = (message.stopped ?? false) ? (message.rows ?? []) : [];
        this.shell?.setLocals(message.stopped ?? false, message.context ?? null, message.rows ?? []);
        return;
      case "setWatches":
        this.shell?.setWatches(message.stopped, message.rows ?? []);
        return;
      case "setDebugState":
        // Leaving break mode drops the values immediately, without waiting for a locals push to
        // say so. Reset produced no such push, so a hover went on reporting `label = "twin"`
        // from a frame that no longer existed: a reading that looks current and is not is worse
        // than none at all (2026-08-07).
        if (message.mode !== "break") {
          this.stoppedLocals = [];
        }

        this.shell?.setDebugMode(message.mode);
        return;
      case "obLibrariesResult":
      case "obTypesResult":
      case "obMembersResult":
        // The Object Browser is its own page in its own window now; its answers never
        // arrive on the editor's transport.
        return;
      case "setSettings":
        // Handed over whole. Coercing here as well as in applySettings meant two normalisations
        // of the same six values, and the second one silently decided what a missing field meant.
        applySettings(message);
        return;
      default: {
        const unknown: never = message;
        console.warn("[xlide] unhandled host message", unknown);
      }
    }
  }

  /**
   * The theme follows the operating system, and nothing overrides it.
   *
   * There used to be a pin here for a host-sent `setTheme` to set. Nothing ever sent one - not the
   * shim, not the page, not a probe - so the flag was read false on every branch that read it, and
   * the whole path was a protocol the README described and the product did not have.
   */
  applyOsTheme(theme: XlideTheme): void {
    monaco.editor.setTheme(theme);
  }

  /**
   * Runs a command the host named: one of Monaco's actions, or one of the surface's own.
   *
   * The host reaches for this when a route it owns lands on something the surface owns: a claimed
   * key, or a menu item whose native version acts on the covered pane. Undo and redo are built
   * into the editor rather than registered, so they are triggered by name; the panel commands are
   * not editor commands at all and go to the shell.
   */
  /**
   * Writes a line into the host's log, because the page has no log of its own that support can
   * read. For the moments when the question is which side of the bridge went quiet.
   */
  private nextSearchId = 1;

  /** Sends a search to the host; the answer returns as a searchResult with the same id. */
  requestSearch(query: string, matchCase: boolean, wholeWord: boolean, scope: string): number {
    const id = this.nextSearchId++;
    this.transport.post({ type: "search", id, query, matchCase, wholeWord, scope });
    return id;
  }

  /** Sends a replace-all; the answer lists what remained matched, plus the replaced count. */
  requestReplaceAll(query: string, matchCase: boolean, wholeWord: boolean, scope: string, replacement: string): number {
    const id = this.nextSearchId++;
    this.transport.post({ type: "replaceAll", id, query, matchCase, wholeWord, scope, replacement });
    return id;
  }
  trace(text: string): void {
    this.transport.post({ type: "trace", text });
  }

  /** Opens the settings dialog; wired by the page's entry point, which owns the dialog. */
  openSettings: (() => void) | null = null;

  private runEditorCommand(id: string): void {
    this.trace(`editorCommand ${id}`);

    if (id === "xlide.openSettings") {
      this.openSettings?.();
      return;
    }

    if (id === "xlide.panel.immediate") {
      this.shell?.showImmediate();
      return;
    }

    if (id === "xlide.panel.locals") {
      this.shell?.showLocalsPanel();
      return;
    }

    if (id === "xlide.panel.properties") {
      this.shell?.revealProperties();
      return;
    }

    // Tab cycling arrives from the host because the browser swallows Ctrl+PageDown for its own
    // tab switching before the page could ever see the key. Cycling is within the active group.
    if (id === "xlide.tab.close") {
      // The workspace knows which tab is active and which workbook it belongs to, so nothing has
      // to be guessed from a name. It goes through the same gate the tab's own X uses, so a
      // module with unsaved changes still gets the question.
      this.workspace?.closeActive();
      return;
    }

    if (id === "xlide.tab.next" || id === "xlide.tab.previous") {
      const target = this.workspace?.cycleTab(id === "xlide.tab.next" ? 1 : -1);
      this.trace(`cycle -> ${target ?? "(nothing)"}`);
      return;
    }

    if (id === "xlide.split.right" || id === "xlide.split.down") {
      this.workspace?.splitActive(id === "xlide.split.right" ? "right" : "down");
      return;
    }

    const editor = this.ed();
    if (!editor) {
      return;
    }

    editor.focus();

    if (id === "undo" || id === "redo") {
      editor.trigger("xlide", id, null);
      return;
    }

    editor.getAction(id)?.run();
  }

  private model(): monaco.editor.ITextModel | null {
    return this.ed()?.getModel() ?? null;
  }

  /**
   * How many text models exist in the page, against how many documents are open.
   *
   * The two must match. A model that outlives the document it belonged to is the leak this
   * surface is most likely to grow - models are created per open module and disposed when
   * its pane closes, and nothing else would notice one that stayed. Monaco is bundled
   * rather than global, so a probe cannot count them without this.
   */
  modelCensus(): { models: number; documents: number } {
    return {
      models: monaco.editor.getModels().length,
      documents: this.documents.all().length,
    };
  }

  /**
   * What the page is still waiting on, for the debug api's `ui` route.
   *
   * A document request that never comes home fails as a four-second silence, and the surface
   * shows nothing while it does. Without this, "the tab is blank" and "the text has not arrived
   * yet" look identical from outside, which is two rounds of reading the wrong code.
   */
  waitingOn(): { documents: string[]; hostActive: DocumentId | null; stoppedRows: number } {
    return {
      documents: [...this.pendingDocuments.keys()],
      hostActive: this.hostActive ? { ...this.hostActive } : null,
      stoppedRows: this.stoppedLocals.length,
    };
  }

  /** The host-active document's model, which is what engine answers are computed against. */
  hostActiveModel(): monaco.editor.ITextModel | null {
    return this.hostActive
      ? this.documents.get(this.hostActive.module, this.hostActive.project)
      : null;
  }

  /** Shows the empty workspace: every pane is closed and the editor should say so. */
  private clearDocuments(): void {
    this.markersByDoc.clear();
    this.setCurrentLine(null);
    this.setBreakpoints([]);

    this.applyingHostEdit = true;
    try {
      this.workspace?.clear();
      this.documents.clear();
    } finally {
      this.applyingHostEdit = false;
    }

    this.hostActive = null;
    this.pendingCaret = null;
    this.shell?.setWorkspaceEmpty(true);
  }

  /**
   * A module is open: its model exists from here until its pane closes. Idempotent - a model
   * that already exists adopts the text in place (the host re-opens everything after a page
   * reload, and re-sends a clean document whose module changed underneath), keeping its undo
   * stack and caret. Which group shows it is the workspace's business, decided when the tab
   * list arrives.
   */
  private openDocument(moduleName: string, project: string | null, text: string): void {
    this.documents.open(moduleName, project, text, (model, adopted) => this.adoptText(model, adopted));
    this.shell?.setWorkspaceEmpty(false);

    // A group left with nothing to show asked for this; now it can.
    this.workspace?.documentOpened(moduleName, project);

    // And anything that asked for the text itself rather than for a tab.
    this.documentWaiters.get(docKeyOf(moduleName, project))?.();
  }

  /**
   * Adopts the host's version of a document in place.
   *
   * The host owns the text; this surface is a view of it. When the two differ the host is right,
   * and the difference is usually its own doing: it respells keywords as it takes a module in.
   *
   * Applied as one edit rather than by setting the value, so the undo stack survives and the
   * caret stays where the developer left it. Setting the value discards both, and doing that
   * while somebody is typing moves them to the top of the module mid-word.
   */
  private adoptText(model: monaco.editor.ITextModel, text: string): void {
    const showing = this.workspace?.editorShowing(model) ?? null;
    const selections = showing?.getSelections() ?? null;

    this.applyingHostEdit = true;
    try {
      model.pushEditOperations(
        selections,
        [{ range: model.getFullModelRange(), text, forceMoveMarkers: false }],
        () => selections);
    } finally {
      this.applyingHostEdit = false;
    }

    if (showing && selections) {
      // Clamped by Monaco to the new text, so a position past the end lands at the end rather
      // than being rejected.
      showing.setSelections(selections);
    }

    // Set again, because replacing the text collapsed them all onto its end. Without this a
    // defect reported on line six is drawn under the last line of the module.
    const id = this.documents.idOf(model);
    if (id) {
      this.applyMarkers(model, this.markersByDoc.get(docKeyOf(id.module, id.project)) ?? []);
    }
  }

  private syncDocument(moduleName: string, project: string | null, text: string): void {
    const model = this.documents.get(moduleName, project);
    if (model && model.getValue() !== text) {
      this.adoptText(model, text);
    }
  }

  private onModelContentChanged(id: DocumentId, model: monaco.editor.ITextModel, event: monaco.editor.IModelContentChangedEvent): void {
    if (this.applyingHostEdit || model.isDisposed()) {
      return;
    }

    const key = docKeyOf(id.module, id.project);
    const revision = (this.revisions.get(key) ?? 0) + 1;
    this.revisions.set(key, revision);

    // Monaco reports changes bottom-up so that earlier ranges stay valid; the order is
    // preserved here and the host must apply them in the same order.
    const changes: HostTextChange[] = event.changes.map((change) => ({
      ...fromMonacoRange(change.range),
      text: change.text,
    }));

    // The changes are the message, for every module at every size; the host rebuilds its
    // shadow by applying them. Small modules used to ship their whole text as well and the
    // host preferred that copy, which quietly inverted the coverage: the rebuild - the path
    // whose drift would corrupt the write-back - ran only above 64,000 characters, where
    // nothing tested it, while every keystroke paid an O(document) build, serialize and
    // parse for information the ranges already carry (the audit's C14, measured 2026-08-12:
    // 88ms against 68ms for the same typing). The length rides along so a divergence is
    // seen the moment it happens rather than believed impossible. The message names its
    // document (decision 12): the edit belongs to the module it changed, whichever editor
    // made it.
    // WHERE THE CHANGE CAME FROM, when it was not the keyboard. The host holds back findings
    // about a line while it is being typed, and tells typing from anything else by the shape of
    // the change: one line, no newline. A format that moves a single line has that shape, and
    // running one made the squiggle on that line vanish until the caret was moved off it. The
    // editor's change event carries no source, so the formatter marks its own edit.
    this.transport.post({
      type: "contentChanged",
      moduleName: id.module,
      ...(id.project === null ? {} : { project: id.project }),
      revision,
      changes,
      fullLength: model.getValueLength(),
      ...(takeFormattingMark(model) ? { source: "format" as const } : {}),
    });
  }

  private onSelectionChanged(selection: monaco.Selection): void {
    this.shell?.setPosition(selection.positionLineNumber, selection.positionColumn);
    this.transport.post({
      type: "selectionChanged",
      startLine: selection.startLineNumber,
      startColumn: selection.startColumn,
      endLine: selection.endLineNumber,
      endColumn: selection.endColumn,
    });
  }

  private onMouseDown(event: monaco.editor.IEditorMouseEvent): void {
    if (event.target.type !== monaco.editor.MouseTargetType.GUTTER_GLYPH_MARGIN) {
      return;
    }
    const line = event.target.position?.lineNumber;
    if (line === undefined) {
      return;
    }
    // The page never toggles the breakpoint itself; the host owns the set and answers with
    // setBreakpoints.
    this.transport.post({ type: "breakpointToggleRequested", line });
  }

  private setDiagnostics(moduleName: string, project: string | null, markers: HostMarker[]): void {
    this.markersByDoc.set(docKeyOf(moduleName, project), markers);

    const model = this.documents.get(moduleName, project);
    if (model) {
      this.applyMarkers(model, markers);
    }
  }

  private applyMarkers(model: monaco.editor.ITextModel, markers: HostMarker[]): void {
    const converted: monaco.editor.IMarkerData[] = markers.map((marker) => ({
      severity: SEVERITY[marker.severity] ?? monaco.MarkerSeverity.Error,
      message: marker.message,
      startLineNumber: marker.startLine,
      startColumn: marker.startColumn,
      endLineNumber: marker.endLine,
      endColumn: marker.endColumn,
      ...(marker.code === undefined ? {} : { code: marker.code }),
    }));
    monaco.editor.setModelMarkers(model, MARKER_OWNER, converted);
  }

  /**
   * Replaces one held set of model decorations. Clears the previous model's when the target
   * moved - the stopped line must not survive on a module the debugger has left - and applies
   * the new set on the model that carries it now.
   */
  private applyModelDecor(
    held: { model: monaco.editor.ITextModel; ids: string[] } | null,
    model: monaco.editor.ITextModel | null,
    decorations: monaco.editor.IModelDeltaDecoration[],
  ): { model: monaco.editor.ITextModel; ids: string[] } | null {
    if (held && held.model !== model && !held.model.isDisposed()) {
      held.model.deltaDecorations(held.ids, []);
    }

    if (!model) {
      return null;
    }

    const previous = held && held.model === model ? held.ids : [];
    return { model, ids: model.deltaDecorations(previous, decorations) };
  }

  private setCurrentLine(line: number | null): void {
    const model = line === null ? null : this.hostActiveModel();
    this.currentLineDecor = this.applyModelDecor(
      this.currentLineDecor,
      model,
      line === null || model === null ? [] : [
        {
          range: new monaco.Range(line, 1, line, 1),
          options: {
            isWholeLine: true,
            className: "xlide-current-line",
            glyphMarginClassName: "xlide-current-line-glyph",
            overviewRuler: {
              color: "#ffd24a",
              position: monaco.editor.OverviewRulerLane.Full,
            },
            stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
          },
        },
      ]);
  }

  private onMouseMove(event: monaco.editor.IEditorMouseEvent, hover: monaco.editor.IEditorDecorationsCollection): void {
    const line = event.target.type === monaco.editor.MouseTargetType.GUTTER_GLYPH_MARGIN
      ? event.target.position?.lineNumber
      : undefined;

    // Nothing under the pointer, or a line that already has one: either way there is nothing
    // useful to preview.
    if (line === undefined || this.breakpointLines.has(line)) {
      hover.clear();
      return;
    }

    // The preview tells the truth before the click: a dim dot where a breakpoint can go, an
    // orange cross where one cannot - and a click on a refused line does nothing at all, so
    // the red dot only ever appears where it is real (the developer's design, 2026-08-04).
    const breakable = lineCanCarryBreakpoint(this.model()?.getLineContent(line) ?? "");

    hover.set([
      {
        range: new monaco.Range(line, 1, line, 1),
        options: breakable
          ? {
            isWholeLine: false,
            glyphMarginClassName: "xlide-breakpoint-hover",
            stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
          }
          : {
            isWholeLine: false,
            glyphMarginClassName: "xlide-breakpoint-refused codicon codicon-close",
            stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
          },
      },
    ]);
  }

  private setBreakpoints(lines: number[]): void {
    const sorted = [...new Set(lines)].sort((a, b) => a - b);
    this.breakpointLines = new Set(sorted);

    this.breakpointDecor = this.applyModelDecor(
      this.breakpointDecor,
      this.hostActiveModel(),
      sorted.map((line) => ({
        range: new monaco.Range(line, 1, line, 1),
        options: {
          isWholeLine: false,
          glyphMarginClassName: "xlide-breakpoint-glyph",
          glyphMarginHoverMessage: { value: "Breakpoint" },
          overviewRuler: {
            color: "#e51400",
            position: monaco.editor.OverviewRulerLane.Left,
          },
          stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
        },
      })),
    );
  }
}

export function webView2Transport(): HostTransport | null {
  // window.chrome exists in any Chromium browser; only the webview member proves a host.
  const webview = window.chrome?.webview;
  if (!webview) {
    return null;
  }
  return {
    post(message) {
      webview.postMessage(message);
    },
    subscribe(handler) {
      webview.addEventListener("message", (event) => {
        const data = event.data;
        const parsed = typeof data === "string" ? safeParse(data) : data;
        if (isHostMessage(parsed)) {
          handler(parsed);
        } else {
          console.warn("[xlide] dropped malformed host message", data);
        }
      });
    },
  };
}

function safeParse(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

/**
 * Whether a line can carry a breakpoint. The mirror of the host's CanBreakOn, so the hover
 * preview and the host's verdict on the click always agree; a preview that promises what the
 * click then refuses is worse than no preview.
 */
function lineCanCarryBreakpoint(line: string): boolean {
  const code = line.trim();
  if (code.length === 0 || code.startsWith("'") || startsWithWord(code, "Rem")) {
    return false;
  }

  if (startsWithWord(code, "Option", "Attribute", "Declare", "Dim", "Const", "Type", "Enum")
    || /^end\s+(type|enum)\b/i.test(code)) {
    return false;
  }

  for (const modifier of ["Public", "Private", "Friend", "Static", "Global"]) {
    if (startsWithWord(code, modifier)) {
      const rest = code.slice(modifier.length).trimStart();
      return startsWithWord(rest, "Sub", "Function", "Property");
    }
  }

  return true;
}

function startsWithWord(text: string, ...words: string[]): boolean {
  for (const word of words) {
    if (!text.toLowerCase().startsWith(word.toLowerCase())) {
      continue;
    }
    const next = text[word.length];
    if (next === undefined || !/[A-Za-z0-9_]/.test(next)) {
      return true;
    }
  }
  return false;
}

function isHostMessage(value: unknown): value is HostMessage {
  return typeof value === "object" && value !== null && typeof (value as { type?: unknown }).type === "string";
}

const DEMO_MODULE = `Attribute VB_Name = "Module1"
Option Explicit

' xlide demo module. The host is absent, so this text is local.
Private Const Banner As String = "he said ""hello"" and left"
Private Const Started As Date = #2026-01-31#
Private Const Mask As Long = &HFF00&
Private Const Bits As Integer = &O777

Public Sub Recalculate(ByVal sheetName As String, Optional ByVal force As Boolean = False)
    Dim rowIndex As Long
    Dim total As Double
    Dim values() As Variant

    ReDim values(1 To 10)

    If force Then
        Debug.Print Banner, Started, Mask, Bits
    ElseIf sheetName = vbNullString Then
        Exit Sub
    Else
        total = 1.5E-3 + 42 + 0.25
    End If

    For rowIndex = 1 To 10 Step 1
        total = total + _
            rowIndex * 2
    Next rowIndex

    Do While total > 0
        total = total - 1
    Loop

    Select Case sheetName
        Case "Sheet1"
            Rem legacy comment form
        Case Else
            With Application
                .Calculate
            End With
    End Select
End Sub

Public Function Describe(ByRef target As Object) As String
    If target Is Nothing Then
        Describe = "nothing"
    Else
        Describe = TypeName(target)
    End If
End Function
`;

const DEMO_MODULE_2 = `Attribute VB_Name = "Module2"
Option Explicit

' The demo's second document, so tab switching and split groups are exercisable here.
Public Sub Report(ByVal title As String)
    Dim total As Double
    total = 2.5

    Debug.Print title; total
End Sub
`;

/**
 * Loopback transport used when the page is opened outside WebView2. It logs everything the
 * page would have sent and replays a scripted set of host messages so the surface is testable
 * in a plain browser.
 */
export function demoTransport(): HostTransport {
  let deliver: ((message: HostMessage) => void) | null = null;

  const send = (message: HostMessage): void => {
    console.log("[xlide demo] host -> page", message);
    deliver?.(message);
  };

  // The demo's open tabs and which of them carry unsaved changes, so the close-confirm
  // loop is exercisable in a plain browser: a dirty close is answered with the question,
  // and the answer closes the tab the way the host would. The active tab follows
  // activateModule, so switching and splitting behave here the way they do in the host.
  const openModules = ["Module1", "Module2"];
  const dirtyModules = new Set(openModules);
  let activeModule: string | null = openModules[0] ?? null;

  // Every module the demo can show, by name, so activating one that is closed can re-open it
  // with its text - which is what the host does, and what a drag of a closed module's tree row
  // onto the editor relies on.
  const moduleTexts: Record<string, string> = { Module1: DEMO_MODULE, Module2: DEMO_MODULE_2 };

  // The demo's two modules both live in Book1.xlsm, which is what the tree says - so setModules
  // must say it too. The host always carries the project per open module; without it here the
  // tabs read project:null while the tree row reads Book1.xlsm, and the two cannot be recognised
  // as the same document. A tree-row drag of an OPEN module then found no tab to move (found
  // 2026-08-12 while building the drag targets, where the demo's own inconsistency masqueraded
  // as a feature bug).
  const DEMO_WORKBOOK = "Book1.xlsm";
  const sendModules = (): void => {
    if (activeModule !== null && !openModules.includes(activeModule)) {
      activeModule = openModules[0] ?? null;
    }

    send({
      type: "setModules",
      modules: [...openModules],
      projects: openModules.map(() => DEMO_WORKBOOK),
      active: activeModule,
      activeProject: activeModule === null ? null : DEMO_WORKBOOK,
      dirty: openModules.map((name) => dirtyModules.has(name)),
    });
  };

  // The demo's tree, held rather than sent once, so removing a component can take it out of both
  // lists and republish - which is what the host does, and what makes the removal exercisable
  // here at all. Two workbooks, and a name that lives in only one of them.
  const demoProjects: ExplorerProject[] = [
    {
      name: "Book1.xlsm",
      components: [
        { name: "Sheet1", kind: 100 },
        { name: "ThisWorkbook", kind: 100 },
        { name: "Module1", kind: 1 },
        { name: "SalesRow", kind: 2 },
      ],
    },
    {
      name: "Book2.xlsm",
      components: [
        { name: "ThisWorkbook", kind: 100 },
        { name: "Helpers", kind: 1 },
      ],
    },
  ];

  /*
   * ONE WORKBOOK ON REQUEST, with `?books=1`.
   *
   * The tree behaves differently at one workbook than at several: closing the last tab folds the
   * workbooks away, unless there is only one, where folding the only row would hide the whole
   * project behind a click with one possible answer. Two branches, and the demo could only ever
   * reach the first, so the interesting one was verifiable against a real Excel and nowhere else.
   *
   * The demo exists to make the page drivable without a host, and a demo that cannot reach half
   * the states the page has is only half a demo.
   */
  const soleWorkbook = new URLSearchParams(location.search).get("books") === "1";

  /*
   * THE HOST SAYS WHEN NOTHING IS OPEN, and the demo did not.
   *
   * Closing the last pane makes the real session send `clearDocuments`, which is what puts the
   * surface into its empty state and folds the tree back. The demo only ever republished a shorter
   * tab list, so a page driven here could never reach the empty workspace at all: every headless
   * check about what happens when the last tab closes was measuring a state the demo cannot enter,
   * and one of them passed for exactly that reason (2026-08-10).
   */
  const sendEmptyWhenNothingIsLeft = (): void => {
    if (openModules.length === 0) {
      send({ type: "clearDocuments" });
    }
  };

  const sendProjects = (): void => {
    const shown = soleWorkbook ? demoProjects.slice(0, 1) : demoProjects;
    send({ type: "setProjects", projects: shown.map((one) => ({ ...one })) });
  };

  return {
    post(message) {
      console.log("[xlide demo] page -> host", message);
      if (message.type === "ready") {
        send({ type: "openDocument", moduleName: "Module1", project: DEMO_WORKBOOK, text: DEMO_MODULE });
        send({ type: "openDocument", moduleName: "Module2", project: DEMO_WORKBOOK, text: DEMO_MODULE_2 });
        sendModules();
        send({
          type: "setSettings",
          blockLayout: "comfy",
          continueCommentOnNewline: true,
          mirrorCommentSpacing: true,
          treeFollowsEditor: true,
          formatIndentSize: 4,
        });
        send({
          type: "setLocals",
          stopped: true,
          context: "VBAProject.Module1.Demo",
          rows: [
            { expression: "counter", value: "42", kind: "Long" },
            { expression: "label", value: '"hello"', kind: "String" },
            { expression: "cells", value: "", kind: "Object/Range" },
          ],
        });
        send({
          type: "setWatches",
          stopped: true,
          rows: [
            { expression: "counter > 40", value: "True", kind: "Boolean", context: "Module1.Demo" },
            { expression: "label", value: '"hello"', kind: "String", context: "Module1.Demo" },
          ],
        });
        sendProjects();
        send({
          type: "setDiagnostics",
          moduleName: "Module1",
          markers: [
            {
              startLine: 11,
              startColumn: 9,
              endLine: 11,
              endColumn: 17,
              severity: "warning",
              message: "Variable 'rowIndex' shadows an outer declaration.",
              code: "XL0101",
            },
          ],
        });
        send({ type: "setBreakpoints", lines: [17, 30] });
        send({ type: "setCurrentLine", line: 17 });
        send({ type: "revealLine", line: 17 });
        send({ type: "setDebugState", mode: "break" });
        send({
          type: "setProperties",
          component: "Module1",
          kind: "Module",
          properties: [{ name: "(Name)", value: "Module1", writable: true, boolean: false }],
        });
      }
      if (message.type === "activateModule") {
        // Activating a CLOSED module opens it, the way the host does - it re-enters the open
        // list and its text is (re)sent - so a drag of a closed module's tree row lands a tab
        // that has content. A module the demo does not know is ignored rather than opened blank.
        if (!openModules.includes(message.moduleName) && moduleTexts[message.moduleName]) {
          openModules.push(message.moduleName);
          send({
            type: "openDocument",
            moduleName: message.moduleName,
            project: DEMO_WORKBOOK,
            text: moduleTexts[message.moduleName]!,
          });
        }

        if (openModules.includes(message.moduleName)) {
          activeModule = message.moduleName;
        }
        sendModules();
      }
      if (message.type === "closeModule") {
        if (!message.action && dirtyModules.has(message.name)) {
          // The host holds a dirty close and asks; the page's modal answers.
          send({ type: "confirmClose", name: message.name, project: message.project ?? null });
        } else {
          if (message.action === "save") {
            // Saving the workbook cleans every module in it, the way the editor saves.
            dirtyModules.clear();
          } else {
            dirtyModules.delete(message.name);
          }

          const at = openModules.indexOf(message.name);
          if (at >= 0) {
            openModules.splice(at, 1);
          }

          sendModules();
          sendEmptyWhenNothingIsLeft();
        }
      }
      if (message.type === "removeComponent") {
        // What the host does, in the order it does it: the component leaves the project, its tab
        // leaves the strip, and both lists are republished. A document module is refused here too,
        // so the demo cannot be used to prove a removal the real host would never perform.
        const owner = message.project
          ? demoProjects.find((one) => one.name === message.project)
          : demoProjects.find((one) => one.components.some((part) => part.name === message.name));
        const doomed = owner?.components.find((part) => part.name === message.name);

        if (owner && doomed && doomed.kind !== 100) {
          owner.components = owner.components.filter((part) => part !== doomed);

          const at = openModules.indexOf(message.name);
          if (at >= 0) {
            openModules.splice(at, 1);
          }
          dirtyModules.delete(message.name);

          sendModules();
          sendProjects();
        }
      }
      if (message.type === "selectComponent") {
        // A worksheet-shaped answer, so the boolean dropdowns and the header are reachable.
        send({
          type: "setProperties",
          component: message.name,
          kind: "Worksheet",
          properties: [
            { name: "(Name)", value: message.name, writable: true, boolean: false },
            { name: "Name", value: "Sheet1", writable: true, boolean: false },
            { name: "StandardWidth", value: "8.43", writable: true, boolean: false },
            { name: "EnableCalculation", value: "True", writable: true, boolean: true },
            { name: "Visible", value: "-1", writable: true, boolean: false },
            { name: "Parent", value: "[object]", writable: false, boolean: false },
          ],
        });
      }
      if (message.type === "editProperty") {
        send({ type: "notice", text: `${message.name} set to '${message.value}'` });
      }
      if (message.type === "breakpointToggleRequested") {
        // The real host owns the breakpoint set; the demo just echoes the single line back.
        send({ type: "setBreakpoints", lines: [message.line] });
      }
      if (message.type === "menu") {
        send({ type: "setMenu", path: message.path, items: demoMenuItems(message.path) });
      }
      if (message.type === "menuExecute") {
        send({ type: "notice", text: `menu [${message.path.join(", ")}] executed` });
      }
      if (message.type === "hover") {
        send({
          type: "hoverResult",
          id: message.id,
          hover: {
            signature: "ThisWorkbook As Workbook",
            details: ["Excel host global"],
            start: message.offset,
            end: message.offset,
          },
        });
      }
      // One fix over the demo's one squiggle, so the lightbulb and its menu are exercisable in a
      // plain browser. The span is the marker's, which is how the real host answers: the fix
      // belongs to the finding, not to wherever the caret happened to be.
      if (message.type === "codeAction") {
        const shadowed = DEMO_MODULE.indexOf("rowIndex");
        const answers = message.start <= shadowed + "rowIndex".length && shadowed <= message.end;

        send({
          type: "codeActionResult",
          id: message.id,
          actions: answers
            ? [{
              title: "Rename to 'outerRowIndex'",
              isPreferred: true,
              code: "XL0101",
              start: shadowed,
              end: shadowed + "rowIndex".length,
              edits: [{ start: shadowed, end: shadowed + "rowIndex".length, text: "outerRowIndex" }],
            }]
            : [],
        });
      }
      // The demo's own answers for both, so the peek windows and the context-menu entries are
      // exercisable in a plain browser: every mention of the word under the caret, in both
      // modules, which is what makes it a cross-module answer rather than a local one.
      if (message.type === "definition" || message.type === "references") {
        // The offset is into whichever module is host-active, so the word has to be read out of
        // THAT module's text. Reading it out of Module1's regardless answered about a word that
        // was not under the caret whenever Module2 was showing - and, because Module1 is searched
        // first, made an answer in another module impossible to produce here. That is the one case
        // the surface has now got wrong twice.
        const asked = activeModule === "Module2" ? DEMO_MODULE_2 : DEMO_MODULE;
        const at = asked.slice(0, message.offset).search(/[A-Za-z0-9_]*$/);
        const word = /^[A-Za-z0-9_]+/.exec(asked.slice(at))?.[0] ?? "";
        const locations: HostLocation[] = [];

        if (word.length > 1) {
          for (const [module, text] of [["Module1", DEMO_MODULE], ["Module2", DEMO_MODULE_2]] as const) {
            text.split("\n").forEach((line, index) => {
              const column = line.indexOf(word);
              if (column >= 0) {
                locations.push({
                  module,
                  line: index + 1,
                  column: column + 1,
                  length: word.length,
                  // The host sends the line with every location; the demo must too, or the
                  // references list looks empty here and correct everywhere else.
                  preview: line.trim(),
                });
              }
            });
          }
        }

        send({
          type: "navigationResult",
          id: message.id,
          // A definition is one place; references are all of them.
          locations: message.type === "definition" ? locations.slice(0, 1) : locations,
        });
      }
      // The demo's own colouring, so the legend, the delta encoding and the theme rules are
      // exercisable in a plain browser: the host globals the module names, plus its one type.
      if (message.type === "semanticTokens") {
        const text = message.module === "Module2" ? DEMO_MODULE_2 : DEMO_MODULE;
        const tokens: HostSemanticToken[] = [];
        for (const [name, type, modifiers] of [
          ["Application", "variable", ["defaultLibrary"]],
          ["TypeName", "variable", ["defaultLibrary"]],
          ["Object", "class", []],
        ] as [string, string, string[]][]) {
          for (let at = text.indexOf(name); at >= 0; at = text.indexOf(name, at + 1)) {
            tokens.push({ start: at, end: at + name.length, type, modifiers });
          }
        }

        send({
          type: "semanticTokensResult",
          id: message.id,
          tokens: tokens.sort((left, right) => left.start - right.start),
        });
      }
      if (message.type === "signatureHelp") {
        send({
          type: "signatureHelpResult",
          id: message.id,
          signature: {
            label: "MsgBox(Prompt, [Buttons], [Title]) As VbMsgBoxResult",
            parameters: [{ label: "Prompt" }, { label: "[Buttons]" }, { label: "[Title]" }],
            activeParameter: 0,
          },
        });
      }
      // The demo persists nothing, so the echo IS the write: what the dialog asks for comes
      // straight back, which exercises the same applied-on-echo path the host uses.
      if (message.type === "search" || message.type === "replaceAll") {
        send({
          type: "searchResult",
          id: message.id,
          matches: message.type === "replaceAll" ? [] : [
            { workbook: "Book1.xlsm", module: "Module1", line: 4, column: 8, length: message.query.length, preview: "    Const Banner As String = (demo match)" },
            { workbook: "Book1.xlsm", module: "Module2", line: 12, column: 5, length: message.query.length, preview: "    total = total + 1 (demo match)" },
          ],
          truncated: false,
          replaced: message.type === "replaceAll" ? 2 : 0,
        });
      }
      if (message.type === "updateSettings") {
        send({
          type: "setSettings",
          blockLayout: message.blockLayout === "compact" ? "compact" : "comfy",
          continueCommentOnNewline: message.continueCommentOnNewline,
          mirrorCommentSpacing: message.mirrorCommentSpacing,
          treeFollowsEditor: message.treeFollowsEditor !== false,
          formatIndentSize: message.formatIndentSize,
        });
      }
      // The demo has no engine; answering the recase requests empty keeps a keystroke an
      // ordinary keystroke. Smart Enter, Smart Tab, and loop sync are page-local and need no
      // answers at all, which also makes the demo the place they can be exercised without a host.
      if (message.type === "canonicalCase") {
        send({ type: "canonicalCaseResult", id: message.id, edits: [] });
      }
      if (message.type === "outline") {
        send({
          type: "outlineResult",
          id: message.id,
          procedures: message.module === "Module1"
            ? [
              { name: "Recalculate", kind: "Sub", line: 10 },
              { name: "Describe", kind: "Function", line: 44 },
            ]
            : [],
        });
      }
      if (message.type === "completion") {
        send({
          type: "completionResult",
          id: message.id,
          items: [
            { label: "Worksheets", kind: "property", detail: "Workbook.Worksheets As Sheets" },
            { label: "Range", kind: "property", detail: "Worksheet.Range As Range" },
            { label: "Application", kind: "global", detail: "Excel.Application" },
            { label: "MsgBox", kind: "runtime", detail: "MsgBox(Prompt, [Buttons], ...)" },
            { label: "If", kind: "keyword", detail: "If ... Then", insertText: "If ${1:condition} Then\n\t$0\nEnd If" },
          ],
        });
      }
    },
    subscribe(handler) {
      deliver = handler;
    },
  };
}

/** A menu tree shaped like the host's, so every rendering case is reachable in a browser. */
function demoMenuItems(path: number[]): MenuItem[] {
  const item = (index: number, caption: string, extra: Partial<MenuItem> = {}): MenuItem => ({
    index,
    caption,
    enabled: true,
    separator: false,
    popup: false,
    checked: false,
    ...extra,
  });

  if (path.length === 0) {
    return ["&File", "&Edit", "&View", "&Insert", "F&ormat", "&Debug", "&Run", "&Tools", "&Window", "&Help"]
      .map((caption, i) => item(i + 1, caption, { popup: true }));
  }

  if (path.length === 1) {
    return [
      item(1, "&Save Book1", { shortcut: "Ctrl+S" }),
      item(2, "&Import File...", { separator: true }),
      item(3, "&Export File...", { enabled: false }),
      item(4, "Print && Export", { popup: true, separator: true }),
      item(5, "&Toolbar", { checked: true }),
      item(6, "&Close and Return", { separator: true, shortcut: "Alt+Q" }),
    ];
  }

  return [
    item(1, "Deep &One"),
    item(2, "Deep &Two", { popup: path.length < 3 }),
  ];
}
