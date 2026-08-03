import * as monaco from "monaco-editor/editor/editor.api.js";
import type { ExplorerProject } from "./explorer.js";
import type { MenuItem } from "./menubar.js";
import type { Shell, ShellFinding, ShellProperty } from "./shell.js";
import type { ToolbarCommand } from "./toolbar.js";
import { THEME_DARK, THEME_LIGHT, type XlideTheme } from "./theme.js";
import { VBA_LANGUAGE_ID, updateVbaLanguageFacts } from "./vba.js";

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
  | { type: "loadDocument"; moduleName: string; text: string }
  | { type: "clearDocument" }
  | { type: "syncDocument"; moduleName: string; text: string }
  | { type: "notice"; text: string }
  | { type: "editorCommand"; id: string }
  | { type: "immediateResult"; text: string; failed: boolean }
  | { type: "setModules"; modules: string[]; projects?: (string | null)[]; active: string | null; activeProject?: string | null }
  | { type: "setFindings"; findings: ShellFinding[] }
  | { type: "setProjects"; projects: ExplorerProject[] }
  | { type: "applyEdit"; revision: number; changes: HostTextChange[] }
  | { type: "setTheme"; theme: XlideTheme }
  | { type: "setDiagnostics"; markers: HostMarker[] }
  | { type: "setCurrentLine"; line: number | null }
  | { type: "setBreakpoints"; lines: number[] }
  | { type: "revealLine"; line: number }
  | { type: "setMenu"; path: number[]; items: MenuItem[] }
  | { type: "setChrome"; menuBar: boolean }
  | { type: "setProperties"; component: string; kind: string; properties: ShellProperty[] }
  | { type: "completionResult"; id: number; items: HostCompletionItem[] }
  | { type: "hoverResult"; id: number; hover: HostHoverPayload | null }
  | { type: "signatureHelpResult"; id: number; signature: HostSignatureInfo | null }
  | { type: "canonicalCaseResult"; id: number; edits: HostTextEdit[] }
  | { type: "outlineResult"; id: number; procedures: HostProcedure[]; failed?: boolean }
  | { type: "setLanguageFacts"; types: string[]; procedures: string[] };

/** One procedure in a module's outline: the kind as the tree spells it, and its 1-based line. */
export interface HostProcedure {
  name: string;
  kind: string;
  line: number;
}

/** A text replacement, UTF-16 offsets into the live source; an insertion has start === end. */
export interface HostTextEdit {
  start: number;
  end: number;
  text: string;
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
  | { type: "contentChanged"; revision: number; changes: HostTextChange[]; fullLength: number; fullText?: string }
  | { type: "selectionChanged"; startLine: number; startColumn: number; endLine: number; endColumn: number }
  | { type: "breakpointToggleRequested"; line: number }
  | { type: "activateModule"; moduleName: string; project?: string }
  | { type: "navigate"; module: string; line: number; column: number; project?: string }
  | { type: "command"; name: string }
  | { type: "evaluate"; text: string }
  | { type: "panel"; name: string; open: boolean }
  | { type: "menu"; path: number[] }
  | { type: "menuExecute"; path: number[] }
  | { type: "editProperty"; component: string; name: string; value: string }
  | { type: "selectComponent"; name: string }
  | { type: "closeModule"; name: string; project?: string }
  | { type: "insertComponent"; kind: number; project?: string }
  | { type: "completion"; id: number; offset: number }
  | { type: "hover"; id: number; offset: number }
  | { type: "signatureHelp"; id: number; offset: number }
  | { type: "canonicalCase"; id: number; start: number; end: number; single?: boolean; completeHeader?: boolean }
  | { type: "outline"; id: number; module: string; project?: string }
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

const MARKER_OWNER = "xlide";

const SEVERITY: Record<HostSeverity, monaco.MarkerSeverity> = {
  error: monaco.MarkerSeverity.Error,
  warning: monaco.MarkerSeverity.Warning,
  info: monaco.MarkerSeverity.Info,
  hint: monaco.MarkerSeverity.Hint,
};

function toMonacoRange(range: HostRange): monaco.IRange {
  return {
    startLineNumber: range.startLine,
    startColumn: range.startColumn,
    endLineNumber: range.endLine,
    endColumn: range.endColumn,
  };
}

function fromMonacoRange(range: monaco.IRange): HostRange {
  return {
    startLine: range.startLineNumber,
    startColumn: range.startColumn,
    endLine: range.endLineNumber,
    endColumn: range.endColumn,
  };
}

export class EditorBridge {
  private readonly editor: monaco.editor.IStandaloneCodeEditor;
  private readonly transport: HostTransport;
  private readonly shell: Shell | null;
  private readonly disposables: monaco.IDisposable[] = [];
  private readonly currentLine: monaco.editor.IEditorDecorationsCollection;
  private readonly breakpoints: monaco.editor.IEditorDecorationsCollection;

  /**
   * The faint dot shown under the pointer in the breakpoint margin.
   *
   * The margin is a narrow strip with nothing in it, and nothing about it says it can be clicked.
   * Showing where the breakpoint would land is what makes the target findable, and it also draws
   * the edges of the strip, which is the part that was being guessed at.
   */
  private readonly breakpointHover: monaco.editor.IEditorDecorationsCollection;

  /** Lines that already carry a breakpoint, so the hover dot is not drawn over a real one. */
  private breakpointLines = new Set<number>();

  /**
   * The squiggles the host last sent.
   *
   * Kept because replacing the model's text drags every marker to the end of the replacement.
   * Markers are anchored to positions in the text, and a whole-document edit is, as far as the
   * editor is concerned, the entire text being deleted and different text arriving: everything
   * anchored inside it collapses to one point. They are set again afterwards, at the positions the
   * host gave, which are still correct because the text either did not change or changed only in
   * ways that do not move lines.
   */
  private lastMarkers: HostMarker[] = [];

  /** Monotonic counter over locally originated edits; reset by loadDocument, adopted from applyEdit. */
  private revision = 0;

  /** Completion requests awaiting their answers, by request identifier. */
  private readonly pendingCompletions = new Map<number, {
    resolve: (items: HostCompletionItem[]) => void;
    timer: ReturnType<typeof setTimeout>;
  }>();

  /** Hover requests awaiting their answers, by request identifier. */
  private readonly pendingHovers = new Map<number, {
    resolve: (hover: HostHoverPayload | null) => void;
    timer: ReturnType<typeof setTimeout>;
  }>();

  /** Call-tip requests awaiting their answers, by request identifier. */
  private readonly pendingSignatures = new Map<number, {
    resolve: (signature: HostSignatureInfo | null) => void;
    timer: ReturnType<typeof setTimeout>;
  }>();

  /** Canonical-case requests awaiting their answers, by request identifier. */
  private readonly pendingCanonicalCases = new Map<number, {
    resolve: (edits: HostTextEdit[]) => void;
    timer: ReturnType<typeof setTimeout>;
  }>();

  /** Outline requests awaiting their answers, by request identifier. */
  private readonly pendingOutlines = new Map<number, {
    resolve: (procedures: HostProcedure[] | null) => void;
    timer: ReturnType<typeof setTimeout>;
  }>();

  private nextCompletionId = 1;
  private nextHoverId = 1;
  private nextSignatureId = 1;
  private nextCanonicalCaseId = 1;
  private nextOutlineId = 1;
  /** Echo suppression: true while a host edit is being written into the model. */
  private applyingHostEdit = false;
  /** Once the host names a theme, the OS preference stops overriding it. */
  private themePinned = false;

  constructor(editor: monaco.editor.IStandaloneCodeEditor, transport: HostTransport, shell: Shell | null = null) {
    this.editor = editor;
    this.transport = transport;
    this.shell = shell;
    this.currentLine = editor.createDecorationsCollection([]);
    this.breakpoints = editor.createDecorationsCollection([]);
    this.breakpointHover = editor.createDecorationsCollection([]);

    this.disposables.push(
      editor.onDidChangeModelContent((event) => this.onContentChanged(event)),
      editor.onDidChangeCursorSelection((event) => this.onSelectionChanged(event.selection)),
      editor.onMouseDown((event) => this.onMouseDown(event)),
      editor.onMouseMove((event) => this.onMouseMove(event)),
      editor.onMouseLeave(() => this.breakpointHover.clear()),
    );

    transport.subscribe((message) => this.handle(message));
  }

  get isThemePinned(): boolean {
    return this.themePinned;
  }

  start(timings?: BootTimings): void {
    this.transport.post(timings ? { type: "ready", timings } : { type: "ready" });
  }

  /** Asks the host to show a module. The tree names the workbook it means; a tab cannot yet. */
  activateModule(moduleName: string, project?: string): void {
    this.transport.post({ type: "activateModule", moduleName, ...(project ? { project } : {}) });
  }

  /** Tells the host which panel is showing, so it only watches what is being looked at. */
  panelChanged(name: string, open: boolean): void {
    this.transport.post({ type: "panel", name, open });
  }

  /** Asks the host to evaluate a line entered in the Immediate panel. */
  evaluate(text: string): void {
    this.transport.post({ type: "evaluate", text });
  }

  /**
   * Asks the host to go to a place: a finding, or a procedure picked in the tree. The host
   * shows the module and moves the native caret; the surface's own caret is placed here, the
   * moment the module is the one showing, so the click ends with the cursor at the target and
   * the editor focused, ready to type. A procedure asks for its whole line, so where the block
   * starts is visible at a glance.
   */
  navigate(module: string, line: number, column: number, selectLine = false, project?: string): void {
    this.pendingCaret = { module, line, column, selectLine };
    this.transport.post({ type: "navigate", module, line, column, ...(project ? { project } : {}) });
    this.applyPendingCaret();
  }

  /** A caret waiting for its module to arrive; applied on the next matching loadDocument. */
  private pendingCaret: { module: string; line: number; column: number; selectLine: boolean } | null = null;

  /**
   * Places the waiting caret if the shown module is the one it belongs to. A load of any other
   * module supersedes the navigation, and the wait is abandoned rather than left to fire on
   * some later visit.
   */
  private applyPendingCaret(): void {
    const pending = this.pendingCaret;
    const model = this.model();
    if (!pending || !model) {
      return;
    }

    const uri = monaco.Uri.parse(`xlide:/${encodeURIComponent(pending.module)}`);
    if (model.uri.toString() !== uri.toString()) {
      return;
    }

    this.pendingCaret = null;
    const line = Math.min(Math.max(pending.line, 1), model.getLineCount());
    if (pending.selectLine) {
      this.editor.setSelection(new monaco.Selection(line, 1, line, model.getLineMaxColumn(line)));
    } else {
      const column = Math.min(Math.max(pending.column, 1), model.getLineMaxColumn(line));
      this.editor.setPosition({ lineNumber: line, column });
    }
    this.editor.revealLineInCenterIfOutsideViewport(line);
    this.editor.focus();
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

  /** Asks the host to close a module's pane, which is what closes its tab. */
  closeModule(name: string, project?: string): void {
    this.transport.post({ type: "closeModule", name, ...(project ? { project } : {}) });
  }

  /**
   * Asks the host for a new component: 1 module, 2 class module, 3 form. Named workbook when
   * the request came from a workbook's own menu; the active project otherwise.
   */
  insertComponent(kind: number, project?: string): void {
    this.transport.post({ type: "insertComponent", kind, ...(project ? { project } : {}) });
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
    const id = this.nextCompletionId++;

    return new Promise<HostCompletionItem[]>((resolve) => {
      const timer = setTimeout(() => {
        this.pendingCompletions.delete(id);
        resolve([]);
      }, 2000);

      this.pendingCompletions.set(id, { resolve, timer });
      this.transport.post({ type: "completion", id, offset });
    });
  }

  /**
   * Asks the host what the identifier at an offset is. Resolves null rather than rejecting when
   * the host is slow or gone: a hover that fails is a tooltip that does not appear, not an error.
   */
  requestHover(offset: number): Promise<HostHoverPayload | null> {
    const id = this.nextHoverId++;

    return new Promise<HostHoverPayload | null>((resolve) => {
      const timer = setTimeout(() => {
        this.pendingHovers.delete(id);
        resolve(null);
      }, 2000);

      this.pendingHovers.set(id, { resolve, timer });
      this.transport.post({ type: "hover", id, offset });
    });
  }

  /**
   * Asks the host for the call tip at an offset. Resolves null rather than rejecting: a tip that
   * fails is a tip that does not show.
   */
  requestSignatureHelp(offset: number): Promise<HostSignatureInfo | null> {
    const id = this.nextSignatureId++;

    return new Promise<HostSignatureInfo | null>((resolve) => {
      const timer = setTimeout(() => {
        this.pendingSignatures.delete(id);
        resolve(null);
      }, 2000);

      this.pendingSignatures.set(id, { resolve, timer });
      this.transport.post({ type: "signatureHelp", id, offset });
    });
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
    const id = this.nextCanonicalCaseId++;

    return new Promise<HostTextEdit[]>((resolve) => {
      const timer = setTimeout(() => {
        this.pendingCanonicalCases.delete(id);
        resolve([]);
      }, 2000);

      this.pendingCanonicalCases.set(id, { resolve, timer });
      this.transport.post({
        type: "canonicalCase",
        id,
        start,
        end,
        ...(options.single ? { single: true } : {}),
        ...(options.completeHeader ? { completeHeader: true } : {}),
      });
    });
  }

  /**
   * Asks the host for a module's procedures, for its node in the tree. Resolves null — never
   * empty — when no answer comes: a timeout is not a statement that the module has no
   * procedures, and the difference decides whether the tree keeps what it already shows. The
   * window is generous because the host thread legitimately stalls for seconds while a large
   * module is being shown, and the answer queued behind that stall is still a good answer.
   */
  requestOutline(module: string, project?: string): Promise<HostProcedure[] | null> {
    const id = this.nextOutlineId++;

    return new Promise<HostProcedure[] | null>((resolve) => {
      const timer = setTimeout(() => {
        this.pendingOutlines.delete(id);
        resolve(null);
      }, 8000);

      this.pendingOutlines.set(id, { resolve, timer });
      this.transport.post({ type: "outline", id, module, ...(project ? { project } : {}) });
    });
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

    // Focus first. An editor action taken while the button has focus operates on an editor that
    // does not have it, and the ones that open a widget put it somewhere the developer cannot type.
    this.editor.focus();

    // Undo and redo are not actions. They are built into the editor rather than registered like
    // the rest, so looking them up finds nothing and they have to be triggered by name.
    if (command.id === "undo" || command.id === "redo") {
      this.editor.trigger("xlide", command.id, null);
      return;
    }

    this.editor.getAction(command.id)?.run();
  }

  dispose(): void {
    for (const disposable of this.disposables) {
      disposable.dispose();
    }
    this.disposables.length = 0;
  }

  handle(message: HostMessage): void {
    switch (message.type) {
      case "loadDocument":
        this.loadDocument(message.moduleName, message.text);
        return;
      case "clearDocument":
        this.clearDocument();
        return;
      case "syncDocument":
        this.syncDocument(message.moduleName, message.text);
        return;
      case "notice":
        this.shell?.notify(message.text);
        return;
      case "editorCommand":
        this.runEditorCommand(message.id);
        return;
      case "immediateResult":
        this.shell?.appendImmediate(message.text, message.failed ? "failed" : "result");
        return;
      case "setModules":
        this.shell?.setModules(message.modules, message.projects ?? [], message.active, message.activeProject ?? null);
        return;
      case "setFindings":
        this.shell?.setFindings(message.findings);
        return;
      case "setProjects":
        this.shell?.setProjects(message.projects);
        return;
      case "applyEdit":
        this.applyEdit(message.revision, message.changes);
        return;
      case "setTheme":
        this.themePinned = true;
        monaco.editor.setTheme(message.theme === THEME_LIGHT ? THEME_LIGHT : THEME_DARK);
        return;
      case "setDiagnostics":
        this.setDiagnostics(message.markers);
        return;
      case "setCurrentLine":
        this.setCurrentLine(message.line);
        return;
      case "setBreakpoints":
        this.setBreakpoints(message.lines);
        return;
      case "revealLine":
        this.editor.revealLineInCenterIfOutsideViewport(message.line);
        return;
      case "setMenu":
        this.shell?.setMenu(message.path, message.items);
        return;
      case "setChrome":
        this.shell?.setMenuBarVisible(message.menuBar);
        return;
      case "setProperties":
        this.shell?.setProperties(message.component, message.kind, message.properties);
        return;
      case "completionResult": {
        const waiter = this.pendingCompletions.get(message.id);
        if (waiter) {
          this.pendingCompletions.delete(message.id);
          clearTimeout(waiter.timer);
          waiter.resolve(message.items);
        }
        return;
      }
      case "hoverResult": {
        const waiter = this.pendingHovers.get(message.id);
        if (waiter) {
          this.pendingHovers.delete(message.id);
          clearTimeout(waiter.timer);
          waiter.resolve(message.hover);
        }
        return;
      }
      case "signatureHelpResult": {
        const waiter = this.pendingSignatures.get(message.id);
        if (waiter) {
          this.pendingSignatures.delete(message.id);
          clearTimeout(waiter.timer);
          waiter.resolve(message.signature);
        }
        return;
      }
      case "canonicalCaseResult": {
        const waiter = this.pendingCanonicalCases.get(message.id);
        if (waiter) {
          this.pendingCanonicalCases.delete(message.id);
          clearTimeout(waiter.timer);
          waiter.resolve(message.edits);
        }
        return;
      }
      case "outlineResult": {
        const waiter = this.pendingOutlines.get(message.id);
        if (waiter) {
          this.pendingOutlines.delete(message.id);
          clearTimeout(waiter.timer);
          // A failed answer is a shrug, not a statement of emptiness.
          waiter.resolve(message.failed ? null : message.procedures);
        }
        return;
      }
      case "setLanguageFacts":
        updateVbaLanguageFacts(message.types, message.procedures);
        return;
      default: {
        const unknown: never = message;
        console.warn("[xlide] unhandled host message", unknown);
      }
    }
  }

  applyOsTheme(theme: XlideTheme): void {
    if (this.themePinned) {
      return;
    }
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
  trace(text: string): void {
    this.transport.post({ type: "trace", text });
  }

  private runEditorCommand(id: string): void {
    this.trace(`editorCommand ${id}`);

    if (id === "xlide.panel.immediate") {
      this.shell?.showImmediate();
      return;
    }

    if (id === "xlide.panel.properties") {
      this.shell?.revealProperties();
      return;
    }

    // Tab cycling arrives from the host because the browser swallows Ctrl+PageDown for its own
    // tab switching before the page could ever see the key.
    if (id === "xlide.tab.next" || id === "xlide.tab.previous") {
      const target = this.shell?.cycleTab(id === "xlide.tab.next" ? 1 : -1);
      this.trace(`cycle -> ${target ?? "(nothing)"}`);
      return;
    }

    this.editor.focus();

    if (id === "undo" || id === "redo") {
      this.editor.trigger("xlide", id, null);
      return;
    }

    this.editor.getAction(id)?.run();
  }

  private model(): monaco.editor.ITextModel | null {
    return this.editor.getModel();
  }

  /** Shows the empty workspace: every pane is closed and the editor should say so. */
  private clearDocument(): void {
    this.lastMarkers = [];
    this.currentLine.clear();
    this.breakpoints.clear();
    this.breakpointHover.clear();

    const existing = this.model();
    this.applyingHostEdit = true;
    try {
      this.editor.setModel(null);
      existing?.dispose();
    } finally {
      this.applyingHostEdit = false;
    }

    this.revision = 0;
    this.pendingCaret = null;
    this.shell?.setWorkspaceEmpty(true);
  }

  private loadDocument(moduleName: string, text: string): void {
    // A different module's squiggles are not this one's.
    this.lastMarkers = [];
    this.shell?.setWorkspaceEmpty(false);

    const existing = this.model();
    // A fresh model per module keeps the URI meaningful for markers and disposes the old
    // undo stack, which must not survive a module switch.
    const uri = monaco.Uri.parse(`xlide:/${encodeURIComponent(moduleName)}`);
    const previous = monaco.editor.getModel(uri);
    this.applyingHostEdit = true;
    try {
      if (previous) {
        previous.setValue(text);
        this.editor.setModel(previous);
      } else {
        this.editor.setModel(monaco.editor.createModel(text, VBA_LANGUAGE_ID, uri));
      }
      if (existing && existing.uri.toString() !== uri.toString()) {
        existing.dispose();
      }
    } finally {
      this.applyingHostEdit = false;
    }
    this.revision = 0;
    this.currentLine.clear();
    this.breakpoints.clear();

    // A navigation that asked for this module lands its caret now; one that asked for a
    // different module has been superseded by this load and is dropped.
    if (this.pendingCaret && this.pendingCaret.module === moduleName) {
      this.applyPendingCaret();
    } else {
      this.pendingCaret = null;
    }
  }

  /**
   * Adopts the host's version of the module in place.
   *
   * The host owns the text; this surface is a view of it. When the two differ the host is right,
   * and the difference is usually its own doing: it respells keywords as it takes a module in.
   *
   * Applied as one edit rather than by setting the value, so the undo stack survives and the
   * caret stays where the developer left it. Setting the value discards both, and doing that
   * while somebody is typing moves them to the top of the module mid-word.
   */
  private syncDocument(moduleName: string, text: string): void {
    const model = this.model();
    if (!model || model.getValue() === text) {
      return;
    }

    // A message for a module that is no longer shown is stale by definition.
    const uri = monaco.Uri.parse(`xlide:/${encodeURIComponent(moduleName)}`);
    if (model.uri.toString() !== uri.toString()) {
      return;
    }

    const selections = this.editor.getSelections();

    this.applyingHostEdit = true;
    try {
      model.pushEditOperations(
        selections ?? null,
        [{ range: model.getFullModelRange(), text, forceMoveMarkers: false }],
        () => selections ?? null);
    } finally {
      this.applyingHostEdit = false;
    }

    if (selections) {
      // Clamped by Monaco to the new text, so a position past the end lands at the end rather
      // than being rejected.
      this.editor.setSelections(selections);
    }

    // Set again, because replacing the text collapsed them all onto its end. Without this a
    // defect reported on line six is drawn under the last line of the module.
    this.setDiagnostics(this.lastMarkers);
  }

  private applyEdit(revision: number, changes: HostTextChange[]): void {
    const model = this.model();
    if (!model) {
      return;
    }
    const operations: monaco.editor.IIdentifiedSingleEditOperation[] = changes.map((change) => ({
      range: toMonacoRange(change),
      text: change.text,
      forceMoveMarkers: true,
    }));

    this.applyingHostEdit = true;
    try {
      // pushEditOperations keeps the change on the undo stack; executeEdits via the editor
      // would also move the cursor, which the host edit must not do.
      model.pushEditOperations(null, operations, () => null);
    } finally {
      this.applyingHostEdit = false;
    }
    // The host is the revision authority once it has written to the document.
    this.revision = revision;
  }

  private onContentChanged(event: monaco.editor.IModelContentChangedEvent): void {
    if (this.applyingHostEdit) {
      return;
    }
    const model = this.model();
    if (!model) {
      return;
    }
    this.revision += 1;
    // Monaco reports changes bottom-up so that earlier ranges stay valid; the order is
    // preserved here and the host must apply them in the same order.
    const changes: HostTextChange[] = event.changes.map((change) => ({
      ...fromMonacoRange(change.range),
      text: change.text,
    }));

    // A small module travels whole, which is simplest. A large one travels as its changes:
    // building and shipping the full text per keystroke is what typing latency is made of,
    // and the host reconstructs the same text from the ranges. The length rides along so a
    // divergence would be seen the moment it happened rather than believed impossible.
    const fullLength = model.getValueLength();
    const message: Extract<ClientMessage, { type: "contentChanged" }> = {
      type: "contentChanged",
      revision: this.revision,
      changes,
      fullLength,
    };
    if (fullLength < 64_000) {
      message.fullText = model.getValue();
    }

    this.transport.post(message);
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

  private setDiagnostics(markers: HostMarker[]): void {
    this.lastMarkers = markers;

    const model = this.model();
    if (!model) {
      return;
    }
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

  private setCurrentLine(line: number | null): void {
    if (line === null) {
      this.currentLine.clear();
      return;
    }
    this.currentLine.set([
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

  private onMouseMove(event: monaco.editor.IEditorMouseEvent): void {
    const line = event.target.type === monaco.editor.MouseTargetType.GUTTER_GLYPH_MARGIN
      ? event.target.position?.lineNumber
      : undefined;

    // Nothing under the pointer, or a line that already has one: either way there is nothing
    // useful to preview.
    if (line === undefined || this.breakpointLines.has(line)) {
      this.breakpointHover.clear();
      return;
    }

    this.breakpointHover.set([
      {
        range: new monaco.Range(line, 1, line, 1),
        options: {
          isWholeLine: false,
          glyphMarginClassName: "xlide-breakpoint-hover",
          stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
        },
      },
    ]);
  }

  private setBreakpoints(lines: number[]): void {
    const sorted = [...new Set(lines)].sort((a, b) => a - b);
    this.breakpointLines = new Set(sorted);
    this.breakpointHover.clear();
    this.breakpoints.set(
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

  return {
    post(message) {
      console.log("[xlide demo] page -> host", message);
      if (message.type === "ready") {
        send({ type: "loadDocument", moduleName: "Module1", text: DEMO_MODULE });
        send({ type: "setModules", modules: ["Module1", "Module2"], active: "Module1" });
        send({
          type: "setProjects",
          projects: [
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
          ],
        });
        send({
          type: "setDiagnostics",
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
        send({
          type: "setProperties",
          component: "Module1",
          kind: "Module",
          properties: [{ name: "(Name)", value: "Module1", writable: true, boolean: false }],
        });
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
