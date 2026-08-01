import * as monaco from "monaco-editor/editor/editor.api.js";
import { THEME_DARK, THEME_LIGHT, type XlideTheme } from "./theme.js";
import { VBA_LANGUAGE_ID } from "./vba.js";

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
  | { type: "applyEdit"; revision: number; changes: HostTextChange[] }
  | { type: "setTheme"; theme: XlideTheme }
  | { type: "setDiagnostics"; markers: HostMarker[] }
  | { type: "setCurrentLine"; line: number | null }
  | { type: "setBreakpoints"; lines: number[] }
  | { type: "revealLine"; line: number };

export type ClientMessage =
  | { type: "ready" }
  | { type: "contentChanged"; revision: number; changes: HostTextChange[]; fullText: string }
  | { type: "selectionChanged"; startLine: number; startColumn: number; endLine: number; endColumn: number }
  | { type: "breakpointToggleRequested"; line: number };

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
  private readonly disposables: monaco.IDisposable[] = [];
  private readonly currentLine: monaco.editor.IEditorDecorationsCollection;
  private readonly breakpoints: monaco.editor.IEditorDecorationsCollection;

  /** Monotonic counter over locally originated edits; reset by loadDocument, adopted from applyEdit. */
  private revision = 0;
  /** Echo suppression: true while a host edit is being written into the model. */
  private applyingHostEdit = false;
  /** Once the host names a theme, the OS preference stops overriding it. */
  private themePinned = false;

  constructor(editor: monaco.editor.IStandaloneCodeEditor, transport: HostTransport) {
    this.editor = editor;
    this.transport = transport;
    this.currentLine = editor.createDecorationsCollection([]);
    this.breakpoints = editor.createDecorationsCollection([]);

    this.disposables.push(
      editor.onDidChangeModelContent((event) => this.onContentChanged(event)),
      editor.onDidChangeCursorSelection((event) => this.onSelectionChanged(event.selection)),
      editor.onMouseDown((event) => this.onMouseDown(event)),
    );

    transport.subscribe((message) => this.handle(message));
  }

  get isThemePinned(): boolean {
    return this.themePinned;
  }

  start(): void {
    this.transport.post({ type: "ready" });
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

  private model(): monaco.editor.ITextModel | null {
    return this.editor.getModel();
  }

  private loadDocument(moduleName: string, text: string): void {
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
    this.transport.post({
      type: "contentChanged",
      revision: this.revision,
      changes,
      fullText: model.getValue(),
    });
  }

  private onSelectionChanged(selection: monaco.Selection): void {
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

  private setBreakpoints(lines: number[]): void {
    const sorted = [...new Set(lines)].sort((a, b) => a - b);
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
      }
      if (message.type === "breakpointToggleRequested") {
        // The real host owns the breakpoint set; the demo just echoes the single line back.
        send({ type: "setBreakpoints", lines: [message.line] });
      }
    },
    subscribe(handler) {
      deliver = handler;
    },
  };
}
