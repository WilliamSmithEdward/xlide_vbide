// VBA typing automation: the editor-side halves of Smart Enter, canonical casing, and loop
// iterator sync. The decisions live in the host's engine, which answers with plain edits; this
// controller owns the moments — which keystroke asks, when a line has gone idle, when the caret
// leaves a line — and applies what comes back if the text has not moved in the meantime.
//
// It mirrors the extension's typing automation and canonical-case controller, re-expressed over
// the bridge: the same triggers, the same touch-and-idle line tracking, the same
// leave-a-line-and-it-recases feel. An answer that arrives after more typing is dropped, never
// merged: the next pass over the line asks again against the text as it stands.

import * as monaco from "monaco-editor/editor/editor.api.js";
import type { EditorBridge, HostTextEdit } from "./bridge.js";

/** How long a touched line rests before its recase pass, the extension's own figure. */
const CANONICAL_LINE_IDLE_DELAY_MS = 200;

/** The change a plain Enter makes: a newline plus whatever indent the editor added. */
const PLAIN_ENTER = /^\r?\n[ \t]*$/;

/** Lines that could be half of a For/Next pair, checked before asking the host at all. */
const LOOP_LINE = /^[ \t]*(?:For|Next)\b/i;

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

    if (PLAIN_ENTER.test(change.text)) {
      void this.smartEnter(model, change);
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

  /** Asks what the Enter that just went in should leave behind, and puts the caret there. */
  private async smartEnter(
    model: monaco.editor.ITextModel,
    change: monaco.editor.IModelContentChange,
  ): Promise<void> {
    const offset = model.getOffsetAt({
      lineNumber: change.range.startLineNumber,
      column: change.range.startColumn,
    });
    const version = model.getVersionId();

    const result = await this.bridge.requestSmartEnter(offset);
    if (!result || result.edits.length === 0) {
      return;
    }

    if (this.editor.getModel() !== model || model.getVersionId() !== version) {
      this.bridge.trace("typing: smart enter stale, dropped");
      return;
    }

    const operations = result.edits.map((edit) => ({
      range: monaco.Range.fromPositions(model.getPositionAt(edit.start), model.getPositionAt(edit.end)),
      text: edit.text,
    }));

    this.applying = true;
    try {
      // Grouped with the Enter itself: one undo returns the developer to the line as typed.
      this.editor.executeEdits("xlide-smart-enter", operations);
      this.editor.pushUndoStop();

      if (result.caret !== null) {
        const caret = model.getPositionAt(result.caret);
        this.editor.setPosition(caret);
        this.candidate = { model, position: caret };
      }
    } finally {
      this.applying = false;
    }
  }

  private maybeLoopSync(model: monaco.editor.ITextModel, change: monaco.editor.IModelContentChange): void {
    const line = Math.min(change.range.startLineNumber, model.getLineCount());
    if (!LOOP_LINE.test(model.getLineContent(line))) {
      return;
    }

    const column = Math.min(change.range.startColumn + change.text.length, model.getLineMaxColumn(line));
    const offset = model.getOffsetAt({ lineNumber: line, column });
    void this.loopSync(model, offset);
  }

  private async loopSync(model: monaco.editor.ITextModel, offset: number): Promise<void> {
    const version = model.getVersionId();
    const edits = await this.bridge.requestLoopSync(offset);
    if (edits.length === 0) {
      return;
    }

    // Dropped is fine: the pair resyncs on the next keystroke that touches it.
    this.applyIfCurrent(model, version, edits, "typing: loop sync");
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
