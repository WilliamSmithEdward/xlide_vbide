// VBA typing automation: Smart Enter, Smart Tab, Smart Backspace, canonical casing, and loop
// iterator sync - the extension's typing feel, re-created over Monaco.
//
// The structural half runs HERE, in the page, on the extension's own pure helpers, bundled
// straight from its source: what Enter leaves behind, what Tab means at this caret, what
// Backspace does to an empty comment continuation, which `Next` renames with its `For`. The
// extension computes these beside the document and applies them within a frame, and that
// immediacy is the feature; a round trip to the host's engine put an answer's latency between
// Enter and its `End If`, and under analysis load the wait was visible. Text the page already
// holds needs no round trip to describe.
//
// Canonical casing stays in the engine: respelling an identifier needs the project's
// declarations, and only the engine holds every module.

import * as monaco from "monaco-editor/editor/editor.api.js";
import {
  commentContinuationText,
  detectSmartBlockOpener,
  isSmartBlockClosedAhead,
  procedureHeaderParensEdit,
  resolveLoopIteratorSyncEdit,
  smartBlockInsertion,
  withMemberContinuationText,
} from "xlide-spec/vbaSmartEnter";
import { lexerStrippedLine, lexerStrippedLines } from "xlide-spec/analyzer/lexer/strippedLines";
import { smartTabShouldIndentLine } from "xlide-spec/vbaSmartTab";
import type { EditorBridge, HostTextEdit } from "./bridge.js";
import { currentSettings } from "./settings.js";

/** How long a touched line rests before its recase pass, the extension's own figure. */
const CANONICAL_LINE_IDLE_DELAY_MS = 200;

/**
 * One indent level: that many spaces, the width the developer asked for.
 *
 * The same width the editor's own `tabSize`/`indentSize` follow, so everything that indents
 * agrees. Never a tab: VBA's code store will not hold one, and expands any it is handed.
 */
function indentUnitOf(): string {
  return " ".repeat(Math.max(1, currentSettings().formatIndentSize));
}

/** The change a plain Enter makes: a newline plus whatever indent the editor added. */
const PLAIN_ENTER = /^\r?\n[ \t]*$/;

/** Lines that could be half of a For/Next pair, checked before scanning the document at all. */
const LOOP_LINE = /^[ \t]*(?:For|Next)\b/i;

/** Distinguishes each editor's key rules from every other editor's. See the scope key below. */
let nextEditorScope = 1;

// The extension reads these from settings, and so does this surface now: the host loads the
// developer's file, the page applies what arrives, and every Enter reads the choice as it
// stands. See settings.ts for the store and settingsdialog.ts for where choices are made.

export function installTypingAutomation(
  editor: monaco.editor.IStandaloneCodeEditor,
  bridge: EditorBridge,
): void {
  new TypingAutomation(editor, bridge);
}

class TypingAutomation {
  /** True while one of this controller's own edits is going into the model. */
  private applying = false;

  /** Lines the developer has touched, by model and line: the gate every recase runs behind. */
  private readonly touched = new Set<string>();

  /** One pending recase timer per touched line, collapsed as keystrokes land on it. */
  private readonly timers = new Map<string, ReturnType<typeof setTimeout>>();

  /** Where the caret last rested, so the line it leaves can be recased behind it. */
  private candidate: { model: monaco.editor.ITextModel; position: monaco.Position } | null = null;

  constructor(
    private readonly editor: monaco.editor.IStandaloneCodeEditor,
    private readonly bridge: EditorBridge,
  ) {
    editor.onDidChangeModelContent((event) => this.onContentChanged(event));
    editor.onDidChangeCursorPosition((event) => this.onCursorChanged(event));
    editor.onDidBlurEditorText(() => this.flushCandidateLine());
    editor.onDidChangeModel(() => this.reset());

    // Tab and Backspace, rebound under the extension's own when-clauses. Both fall through to
    // the editor's stock command whenever the smart answer is "nothing", so the keys never go
    // dead - they only gain the VBE-flavoured cases.
    //
    // The scope key is what makes them safe with more than one editor group. A standalone
    // editor's addCommand registers into a keybinding service SHARED by every editor on the
    // page, and the when-clause is the only scoping there is: with two groups, two identical
    // Backspace rules matched, the later one won everywhere, and its handler deleted in ITS
    // editor - so Backspace looked dead in the group being typed in (2026-08-06). A context
    // key created on THIS editor is true only in this editor's context, so each rule matches
    // exactly the editor it belongs to.
    const scope = `xlideTypingScope${nextEditorScope++}`;
    editor.createContextKey(scope, true);

    editor.addCommand(
      monaco.KeyCode.Tab,
      () => this.smartTab(),
      `${scope} && editorTextFocus && !editorReadonly && !suggestWidgetVisible && !inSnippetMode`
      + " && !editorTabMovesFocus && !inlineSuggestionVisible");
    editor.addCommand(
      monaco.KeyCode.Backspace,
      () => this.smartBackspace(),
      `${scope} && editorTextFocus && !editorReadonly && !suggestWidgetVisible && !inSnippetMode`);

    const model = editor.getModel();
    const position = editor.getPosition();
    this.candidate = model && position ? { model, position } : null;
  }

  private onContentChanged(event: monaco.editor.IModelContentChangedEvent): void {
    const model = this.editor.getModel();
    if (!model) {
      return;
    }

    if (this.bridge.isApplyingHostEdit) {
      // The host rewrote text under us, so the line numbers in the touch map may no longer
      // name the lines they did. Forgetting them beats recasing the wrong line.
      this.forgetModel(model);
      return;
    }

    if (this.applying) {
      return;
    }

    this.trackForRecase(model, event.changes);

    if (event.changes.length !== 1) {
      return;
    }
    const change = event.changes[0];
    if (!change) {
      return;
    }

    if (change.rangeLength === 0 && PLAIN_ENTER.test(change.text)) {
      // After the event settles, the way the extension's document listener runs after the
      // editor's own Enter: the editor's auto-indent is already on the body line, and editing
      // from inside a content event would re-enter this listener mid-flight.
      const openerLineNumber = change.range.startLineNumber;
      queueMicrotask(() => this.smartEnter(model, openerLineNumber));
      return;
    }

    if (!/[\r\n]/.test(change.text)) {
      this.maybeLoopSync(model, change);
    }
  }

  /**
   * Marks the lines a change touched and schedules their recase: at once for the line an Enter
   * departed, after an idle rest for a line still being typed on.
   */
  private trackForRecase(
    model: monaco.editor.ITextModel,
    changes: readonly monaco.editor.IModelContentChange[],
  ): void {
    const touchedNow = new Set<number>();
    const immediate = new Set<number>();

    for (const change of changes) {
      const line = Math.min(change.range.startLineNumber, model.getLineCount());
      touchedNow.add(line);
      this.touched.add(this.lineKey(model, line));

      if (change.rangeLength === 0 && PLAIN_ENTER.test(change.text)) {
        immediate.add(line);
        this.scheduleRecase(model, line, 0);
      }
    }

    for (const line of touchedNow) {
      if (!immediate.has(line)) {
        this.scheduleRecase(model, line, CANONICAL_LINE_IDLE_DELAY_MS);
      }
    }
  }

  private scheduleRecase(model: monaco.editor.ITextModel, line: number, delayMs: number): void {
    const key = this.lineKey(model, line);
    const existing = this.timers.get(key);
    if (existing !== undefined) {
      clearTimeout(existing);
    }

    this.timers.set(key, setTimeout(() => {
      this.timers.delete(key);
      void this.recaseLine(model, line, false);
    }, delayMs));
  }

  /** Asks for a whole line's case corrections; leaving a line also completes a bare header. */
  private async recaseLine(
    model: monaco.editor.ITextModel,
    line: number,
    completeHeader: boolean,
  ): Promise<void> {
    if (this.editor.getModel() !== model || !this.touched.has(this.lineKey(model, line))) {
      return;
    }
    if (line < 1 || line > model.getLineCount()) {
      return;
    }

    const start = model.getOffsetAt({ lineNumber: line, column: 1 });
    const end = model.getOffsetAt({ lineNumber: line, column: model.getLineMaxColumn(line) });
    const version = model.getVersionId();

    const edits = await this.bridge.requestCanonicalCase(
      start,
      end,
      completeHeader ? { completeHeader: true } : {},
    );
    if (edits.length === 0) {
      return;
    }

    if (!this.applyIfCurrent(model, version, edits, "typing: recase")) {
      // More typing landed while the answer was in flight. The next quiet moment asks again,
      // against the text as it stands then.
      if (this.editor.getModel() === model && line <= model.getLineCount()) {
        this.scheduleRecase(model, line, CANONICAL_LINE_IDLE_DELAY_MS);
      }
    }
  }

  /** Asks for the single correction of the word ending where the caret rested. */
  private async recaseWordAt(
    model: monaco.editor.ITextModel,
    position: monaco.Position,
  ): Promise<void> {
    if (this.editor.getModel() !== model || !this.touched.has(this.lineKey(model, position.lineNumber))) {
      return;
    }
    if (position.lineNumber > model.getLineCount()) {
      return;
    }

    const column = Math.min(position.column, model.getLineMaxColumn(position.lineNumber));
    const offset = model.getOffsetAt({ lineNumber: position.lineNumber, column });
    const version = model.getVersionId();

    const edits = await this.bridge.requestCanonicalCase(offset, offset, { single: true });
    if (edits.length === 0) {
      return;
    }

    // A dropped single-word recase self-heals: the line stays touched and idles again.
    this.applyIfCurrent(model, version, edits, "typing: recase");
  }

  /**
   * What the Enter that just went in should leave behind, decided the way the extension decides
   * it: the block opener completes with its closer and an indented body line, a whole-line
   * comment continues its apostrophes, a `.member` line inside an open With seeds the next dot.
   * The editor-produced auto-indent is kept when it is already deeper than the opener asks.
   */
  private smartEnter(model: monaco.editor.ITextModel, openerLineNumber: number): void {
    if (this.editor.getModel() !== model || this.applying || this.bridge.isApplyingHostEdit) {
      return;
    }

    const bodyLineNumber = openerLineNumber + 1;
    if (openerLineNumber < 1 || bodyLineNumber > model.getLineCount()) {
      return;
    }

    // Enter mid-line pushed text down, or something landed on the body line already: nothing
    // is owed. The header parens are deliberately not applied alone; the extension waits too.
    const bodyLine = model.getLineContent(bodyLineNumber);
    if (!/^[ \t]*$/.test(bodyLine)) {
      return;
    }

    const openerLine = model.getLineContent(openerLineNumber);
    const headerEdit = procedureHeaderParensEdit(openerLine);
    const normalized = headerEdit
      ? `${openerLine.slice(0, headerEdit.startCol)}${headerEdit.newText}${openerLine.slice(headerEdit.endCol)}`
      : openerLine;

    const opener = detectSmartBlockOpener(lexerStrippedLine(normalized));
    if (!opener) {
      // Not a block: a whole-line comment continues, then a With-member line, in the
      // extension's order. The helpers index lines from zero.
      const settings = currentSettings();
      const source = model.getValue();
      const continuation = settings.continueCommentOnNewline
        ? commentContinuationText(source, openerLineNumber - 1, settings.mirrorCommentSpacing)
        : undefined;
      const lineText = continuation ?? withMemberContinuationText(source, openerLineNumber - 1);
      if (!lineText) {
        return;
      }

      this.applyBodyLine(model, bodyLineNumber, bodyLine, [], lineText, lineText.length, 0);
      return;
    }

    const source = model.getValue();
    const strippedLines = lexerStrippedLines(source);
    strippedLines[openerLineNumber - 1] = lexerStrippedLine(normalized);
    const closedAhead = isSmartBlockClosedAhead(strippedLines, openerLineNumber - 1, opener);

    const insertion = smartBlockInsertion(normalized, bodyLine, opener, {
      eol: model.getEOL(),
      insertCloser: !closedAhead,
      layout: currentSettings().blockLayout,

      // THE DEVELOPER'S INDENT, not the analyzer's default.
      //
      // smartBlockInsertion takes an indentUnit and this call never passed one, so a block body
      // was indented with the analyzer's fallback - a tab - whatever the setting said. With
      // "indent with tabs" OFF, pressing Enter after a plain line gave spaces and pressing it
      // after `If ... Then` gave a tab, in the same file, from the same key (2026-08-08).
      indentUnit: indentUnitOf(),
    });

    const headerOperations: monaco.editor.IIdentifiedSingleEditOperation[] = headerEdit
      ? [{
        range: new monaco.Range(
          openerLineNumber, headerEdit.startCol + 1,
          openerLineNumber, headerEdit.endCol + 1),
        text: headerEdit.newText,
      }]
      : [];

    this.applyBodyLine(
      model,
      bodyLineNumber,
      bodyLine,
      headerOperations,
      insertion.replacementText,
      insertion.bodyText.length,
      insertion.bodyLineOffset);
  }

  /**
   * Replaces the editor-created body line and places the caret at the end of the editable body
   * text, `bodyLineOffset` lines below it. Grouped with the Enter itself on the undo stack: one
   * undo returns the developer to the line as typed.
   */
  private applyBodyLine(
    model: monaco.editor.ITextModel,
    bodyLineNumber: number,
    bodyLine: string,
    extraOperations: monaco.editor.IIdentifiedSingleEditOperation[],
    replacementText: string,
    caretColumnOffset: number,
    bodyLineOffset: number,
  ): void {
    const operations: monaco.editor.IIdentifiedSingleEditOperation[] = [
      ...extraOperations,
      {
        range: new monaco.Range(bodyLineNumber, 1, bodyLineNumber, bodyLine.length + 1),
        text: replacementText,
      },
    ];

    this.applying = true;
    try {
      this.editor.executeEdits("xlide-smart-enter", operations);
      this.editor.pushUndoStop();

      const caretLine = bodyLineNumber + bodyLineOffset;
      if (caretLine <= model.getLineCount()) {
        const caret = new monaco.Position(caretLine, caretColumnOffset + 1);
        this.editor.setPosition(caret);
        this.candidate = { model, position: caret };
      }
    } finally {
      this.applying = false;
    }
  }

  /**
   * Tab, the way the extension's Smart Tab reads it: at a blank line, in the leading
   * whitespace, or over a multi-line selection it indents the line; inside line content it
   * stays an ordinary tab. An empty continued-comment marker is cleared first either way.
   */
  private smartTab(): void {
    const model = this.editor.getModel();
    if (!model) {
      return;
    }

    this.clearEmptyContinuedComment(model);

    const selection = this.editor.getSelection();
    if (!selection) {
      this.editor.trigger("xlide", "tab", null);
      return;
    }

    const position = selection.getPosition();
    const lineText = model.getLineContent(position.lineNumber);
    const selections = this.editor.getSelections() ?? [selection];
    const spansLines = selections.some((s) => s.startLineNumber !== s.endLineNumber);

    // The helper counts columns from zero, the way the extension's editor does.
    if (smartTabShouldIndentLine(lineText, position.column - 1, selection.isEmpty(), spansLines)) {
      this.editor.trigger("xlide", "editor.action.indentLines", null);
    } else {
      this.editor.trigger("xlide", "tab", null);
    }
  }

  /**
   * Backspace on the empty marker a continued comment left behind clears the whole marker, so
   * leaving a comment is one keystroke rather than one per apostrophe. Anything else stays an
   * ordinary Backspace.
   */
  private smartBackspace(): void {
    const model = this.editor.getModel();
    if (!model || !this.clearEmptyContinuedComment(model)) {
      this.editor.trigger("xlide", "deleteLeft", null);
    }
  }

  /**
   * The extension's clearEmptyContinuedComment: a caret at the end of a line that holds only a
   * comment marker continued from the line above deletes the marker. True when it did.
   */
  private clearEmptyContinuedComment(model: monaco.editor.ITextModel): boolean {
    const selections = this.editor.getSelections();
    if (!selections || selections.length !== 1) {
      return false;
    }

    const selection = selections[0];
    if (!selection || !selection.isEmpty()) {
      return false;
    }

    const position = selection.getPosition();
    if (position.lineNumber === 1) {
      return false;
    }

    const line = model.getLineContent(position.lineNumber);
    const before = line.slice(0, position.column - 1);
    const after = line.slice(position.column - 1);
    if (after.trim().length > 0) {
      return false;
    }

    // Any apostrophe run, mirroring commentContinuationText's capture, so 2- and
    // 4+-apostrophe continued comments clear too.
    const match = /^(\s*)('+) ?$/.exec(before);
    const indent = match?.[1];
    const apostrophes = match?.[2];
    if (indent === undefined || apostrophes === undefined) {
      return false;
    }

    const previous = model.getLineContent(position.lineNumber - 1).trimStart();
    if (!previous.startsWith(apostrophes)) {
      return false;
    }

    const markerStart = indent.length;
    this.applying = true;
    try {
      this.editor.executeEdits("xlide-smart-backspace", [{
        range: new monaco.Range(
          position.lineNumber, markerStart + 1,
          position.lineNumber, position.column),
        text: "",
      }]);
    } finally {
      this.applying = false;
    }
    return true;
  }

  /** The paired rename when an edit touches a simple For/Next iterator: local, synchronous. */
  private maybeLoopSync(model: monaco.editor.ITextModel, change: monaco.editor.IModelContentChange): void {
    const line = Math.min(change.range.startLineNumber, model.getLineCount());
    if (!LOOP_LINE.test(model.getLineContent(line))) {
      return;
    }

    const column = Math.min(change.range.startColumn + change.text.length, model.getLineMaxColumn(line));
    const offset = model.getOffsetAt({ lineNumber: line, column });

    const edit = resolveLoopIteratorSyncEdit(model.getValue(), offset);
    if (!edit) {
      return;
    }

    this.applyIfCurrent(
      model,
      model.getVersionId(),
      [{ start: edit.span.start, end: edit.span.end, text: edit.newText }],
      "typing: loop sync");
  }

  /**
   * The recase moments the caret's travel decides: leaving a line recases the whole line and
   * completes a bare header; resting elsewhere on the same line recases the word left behind.
   */
  private onCursorChanged(event: monaco.editor.ICursorPositionChangedEvent): void {
    const model = this.editor.getModel();
    if (!model) {
      this.candidate = null;
      return;
    }

    if (this.applying || this.bridge.isApplyingHostEdit) {
      this.candidate = { model, position: event.position };
      return;
    }

    const previous = this.candidate;
    if (previous && previous.model === model) {
      if (previous.position.lineNumber !== event.position.lineNumber) {
        void this.recaseLine(model, previous.position.lineNumber, true);
      } else if (
        event.reason === monaco.editor.CursorChangeReason.Explicit
        && !previous.position.equals(event.position)
      ) {
        // Navigation within the line, not the advance typing makes: the word behind the old
        // caret is finished now. Typing's own advances are covered by the idle pass.
        void this.recaseWordAt(model, previous.position);
      }
    }

    this.candidate = { model, position: event.position };
  }

  /** The caret is leaving the editor entirely; settle the line it was on. */
  private flushCandidateLine(): void {
    const candidate = this.candidate;
    if (!candidate) {
      return;
    }
    void this.recaseLine(candidate.model, candidate.position.lineNumber, true);
  }

  /**
   * Applies the host's edits if the text is still the text they were computed against, and
   * says so in the host's log when it is not. Offsets convert to ranges only after the version
   * check, so they always mean what the engine meant.
   */
  private applyIfCurrent(
    model: monaco.editor.ITextModel,
    version: number,
    edits: HostTextEdit[],
    label: string,
  ): boolean {
    if (this.editor.getModel() !== model || model.getVersionId() !== version) {
      this.bridge.trace(`${label}: ${edits.length} edit(s) stale, dropped`);
      return false;
    }

    const operations = edits.map((edit) => ({
      range: monaco.Range.fromPositions(model.getPositionAt(edit.start), model.getPositionAt(edit.end)),
      text: edit.text,
    }));

    this.applying = true;
    try {
      // On the undo stack without moving the caret: a recase is not a cursor event.
      model.pushEditOperations(null, operations, () => null);
    } finally {
      this.applying = false;
    }
    return true;
  }

  /** A model switch is a new world: old line numbers name nothing here. */
  private reset(): void {
    for (const timer of this.timers.values()) {
      clearTimeout(timer);
    }
    this.timers.clear();
    this.touched.clear();

    const model = this.editor.getModel();
    const position = this.editor.getPosition();
    this.candidate = model && position ? { model, position } : null;
  }

  private forgetModel(model: monaco.editor.ITextModel): void {
    const prefix = `${model.uri.toString()}\n`;
    for (const [key, timer] of this.timers) {
      if (key.startsWith(prefix)) {
        clearTimeout(timer);
        this.timers.delete(key);
      }
    }
    for (const key of [...this.touched]) {
      if (key.startsWith(prefix)) {
        this.touched.delete(key);
      }
    }
  }

  private lineKey(model: monaco.editor.ITextModel, line: number): string {
    return `${model.uri.toString()}\n${line}`;
  }
}
