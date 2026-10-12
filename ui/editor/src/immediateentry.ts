/**
 * The Immediate pane's entry line: a one-line instance of the code editor (#97).
 *
 * It was a plain text box, so everything had to be typed out in full while the native editor's
 * own Immediate window offers the same member lists, quick info and parameter info its code
 * window does. As an editor on the VBA language it gets the colouring and every provider the
 * code window has; the providers recognise this model by its URI scheme and ask the host to
 * answer in the scope the line will run in - the stopped procedure in break mode, the module on
 * screen otherwise.
 *
 * Enter runs the line and Up and Down walk the history, unless the completion list is open,
 * when they belong to the list the way they do in the code window. The line stays one line: a
 * newline pasted in becomes a space.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import "monaco-editor/features/placeholderText/register.js";
import { VBA_LANGUAGE_ID } from "./vba.js";

/** The scheme the providers recognise the entry line by. */
export const IMMEDIATE_SCHEME = "xlide-immediate";

export const isImmediateModel = (model: monaco.editor.ITextModel): boolean =>
  model.uri.scheme === IMMEDIATE_SCHEME;

export class ImmediateEntry {
  readonly editor: monaco.editor.IStandaloneCodeEditor;
  private readonly model: monaco.editor.ITextModel;
  private readonly history: string[] = [];
  private historyIndex = 0;

  constructor(host: HTMLElement, private readonly onSubmit: (text: string) => void) {
    this.model = monaco.editor.createModel("", VBA_LANGUAGE_ID, monaco.Uri.parse(`${IMMEDIATE_SCHEME}:/line`));
    this.editor = monaco.editor.create(host, {
      model: this.model,
      automaticLayout: true,
      lineNumbers: "off",
      glyphMargin: false,
      folding: false,
      lineDecorationsWidth: 0,
      lineNumbersMinChars: 0,
      minimap: { enabled: false },
      scrollbar: { vertical: "hidden", horizontal: "hidden", alwaysConsumeMouseWheel: false },
      overviewRulerLanes: 0,
      hideCursorInOverviewRuler: true,
      overviewRulerBorder: false,
      renderLineHighlight: "none",
      scrollBeyondLastLine: false,
      wordWrap: "off",
      contextmenu: false,
      links: false,
      occurrencesHighlight: "off",
      selectionHighlight: false,
      renderWhitespace: "none",
      fontFamily: '"Cascadia Mono", Consolas, "Courier New", monospace',
      fontSize: 12,
      lineHeight: 18,
      padding: { top: 2, bottom: 2 },
      // The lists and tips are taller than the line; drawn fixed they can leave its box.
      fixedOverflowWidgets: true,
      suggest: { showStatusBar: false },
      quickSuggestions: { other: true, comments: false, strings: false },
      ariaLabel: "Evaluate an expression or run a statement",
      placeholder: "? 1 + 1",
    });

    // ONE LINE. A newline typed or pasted becomes a space: the native window runs one line per
    // Enter, and a line broken in two would be two statements the developer never typed.
    this.model.onDidChangeContent(() => {
      const text = this.model.getValue();
      if (/[\r\n]/.test(text)) {
        const position = this.editor.getPosition();
        this.model.setValue(text.replace(/\r?\n|\r/g, " "));
        if (position) {
          this.editor.setPosition({ lineNumber: 1, column: this.model.getLineMaxColumn(1) });
        }
      }
    });

    const listClosed = "!suggestWidgetVisible && !parameterHintsVisible";
    this.editor.addCommand(monaco.KeyCode.Enter, () => this.submit(), "!suggestWidgetVisible");
    this.editor.addCommand(monaco.KeyCode.UpArrow, () => this.walk(-1), listClosed);
    this.editor.addCommand(monaco.KeyCode.DownArrow, () => this.walk(1), listClosed);
  }

  /** The line as typed. */
  get value(): string {
    return this.model.getValue();
  }

  /** Puts text on the line with the caret at its end. */
  setValue(text: string): void {
    this.model.setValue(text.replace(/\r?\n|\r/g, " "));
    this.editor.setPosition({ lineNumber: 1, column: this.model.getLineMaxColumn(1) });
  }

  focus(): void {
    this.editor.focus();
  }

  hasFocus(): boolean {
    return this.editor.hasTextFocus();
  }

  /** The placeholder the host's own idiom suggests. */
  setPlaceholder(text: string): void {
    this.editor.updateOptions({ placeholder: text });
  }

  /** Runs the line, as Enter does. */
  submit(): void {
    const text = this.model.getValue().trim();
    if (text.length === 0) {
      return;
    }

    this.history.push(text);
    this.historyIndex = this.history.length;
    this.model.setValue("");
    this.onSubmit(text);
  }

  /** The last few lines are usually variations on each other, so they are worth walking. */
  private walk(step: number): void {
    if (this.history.length === 0) {
      return;
    }

    this.historyIndex = Math.min(this.history.length, Math.max(0, this.historyIndex + step));
    this.setValue(this.history[this.historyIndex] ?? "");
  }
}
