/*
 * Blame, painted on the module: who last touched each line, at the end of the line, and which
 * lines are not committed yet.
 *
 * The bookmark layer's shape, for the bookmark layer's reason: the hints are Monaco decorations
 * on the MODEL, so they ride edits and survive tab switches for free, and one store serves every
 * editor group. What is held per document is the ON flag - blame stays on for a module across a
 * close and reopen - and the last reading, keyed to the text version it was fetched for.
 *
 * The rows are the host's: `git blame --porcelain` on the committed file, mapped onto the live
 * text by the same diff that draws the pane's rows, so an attribute line the file carries and
 * the module does not, and a local edit, both come out right without this side knowing about
 * either. A reading goes stale when the text moves, so the layer asks again a quiet moment after
 * the last edit, and when the host says the repository moved.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";
import { blameHint, blameHover, blameReadingOf, UNCOMMITTED_HINT, type BlameReading } from "./blameformat.js";
import { docKeyOf, type DocumentId } from "./documents.js";

/** How the layer reaches the host: the `scm` route's arguments. */
export type BlameRequest = (args: Record<string, string>) => Promise<Record<string, unknown>>;

/** How long after the last edit the reading is asked for again. */
const REFETCH_AFTER_EDIT_MS = 1000;

export class Blame {
  /** Documents blame is on for, by document key. Outlives the model: a reopened module keeps it. */
  private readonly on = new Set<string>();

  /** Live decoration ids per model URI. */
  private readonly ids = new Map<string, string[]>();

  /** The last reading per model URI, and the text version it was fetched for. */
  private readonly readings = new Map<string, { version: number; reading: BlameReading }>();

  /** Models already being watched, so an editor swap does not double-register. */
  private readonly watched = new WeakSet<monaco.editor.ITextModel>();

  /** The pending refetch per model URI, so a burst of edits is one ask. */
  private readonly timers = new Map<string, ReturnType<typeof setTimeout>>();

  /** The host's last refusal per model URI, kept until a reading lands; what a probe reads
   * when the layer was turned on and is off again a moment later. */
  private readonly refusals = new Map<string, string>();

  /** Told whenever a document's on/off flag moves, so the pane's button can follow. */
  onChanged: (() => void) | null = null;

  /** Told the host's words when it refuses a reading, so they land where every refusal does. */
  onError: ((message: string) => void) | null = null;

  constructor(
    private readonly ask: BlameRequest,
    private readonly idOf: (model: monaco.editor.ITextModel) => DocumentId | null) {
  }

  private keyOf(model: monaco.editor.ITextModel): string | null {
    const id = this.idOf(model);
    return id && !id.face ? docKeyOf(id.module, id.project) : null;
  }

  /**
   * Whether blame is on for a model, how many lines carry a hint, how many are uncommitted, and
   * how many decorations are actually on the model - the reading and the paint are two claims,
   * and a probe that read only the first would call a layer that painted nothing healthy.
   */
  blameOn(model: monaco.editor.ITextModel): {
    on: boolean; lines: number; uncommitted: number; painted: number; refused: string | null;
  } {
    const key = this.keyOf(model);
    const held = this.readings.get(model.uri.toString());
    const on = key !== null && this.on.has(key);
    return {
      on,
      lines: on && held ? held.reading.lines.length : 0,
      uncommitted: on && held ? held.reading.uncommitted.length : 0,
      painted: (this.ids.get(model.uri.toString()) ?? []).length,
      refused: this.refusals.get(model.uri.toString()) ?? null,
    };
  }

  /** Turns blame on or off for a model's document. False when the model is nobody's document. */
  setOn(model: monaco.editor.ITextModel, on: boolean): boolean {
    const key = this.keyOf(model);
    if (key === null) {
      return false;
    }

    if (on) {
      this.on.add(key);
      void this.fetch(model);
    } else {
      this.on.delete(key);
      this.clearTimer(model);
      this.decorate(model, null);
    }

    this.onChanged?.();
    return true;
  }

  toggle(model: monaco.editor.ITextModel): boolean {
    const key = this.keyOf(model);
    return key !== null && this.setOn(model, !this.on.has(key));
  }

  /** The host says the repository or the folder moved: every painted model is read again. */
  stamped(): void {
    for (const model of monaco.editor.getModels()) {
      const key = this.keyOf(model);
      if (key !== null && this.on.has(key)) {
        void this.fetch(model);
      }
    }
  }

  /** Adopts a model: paints it if its document is on, watches edits and disposal. */
  adopt(model: monaco.editor.ITextModel): void {
    if (this.watched.has(model)) {
      return;
    }

    this.watched.add(model);
    const uri = model.uri.toString();

    const key = this.keyOf(model);
    if (key !== null && this.on.has(key)) {
      void this.fetch(model);
    }

    // The decorations follow the edit; the READING does not. A line typed into is still
    // painted with its last author until the host has looked again, which is why the refetch
    // waits for the typing to pause rather than running per keystroke.
    model.onDidChangeContent(() => {
      const owner = this.keyOf(model);
      if (owner === null || !this.on.has(owner)) {
        return;
      }

      this.clearTimer(model);
      this.timers.set(uri, setTimeout(() => {
        this.timers.delete(uri);
        void this.fetch(model);
      }, REFETCH_AFTER_EDIT_MS));
    });

    model.onWillDispose(() => {
      this.clearTimer(model);
      this.ids.delete(uri);
      this.readings.delete(uri);
      this.refusals.delete(uri);
    });
  }

  /** Wires the blame action into one editor. Called once per editor group. */
  attach(editor: monaco.editor.IStandaloneCodeEditor): void {
    const model = editor.getModel();
    if (model) {
      this.adopt(model);
    }
    editor.onDidChangeModel(() => {
      const next = editor.getModel();
      if (next) {
        this.adopt(next);
      }
    });

    editor.addAction({
      id: "xlide.blame.toggle",
      label: "Toggle Blame",
      contextMenuGroupId: "1_xlide",
      contextMenuOrder: 3,
      run: () => {
        const target = editor.getModel();
        if (target) {
          this.toggle(target);
        }
      },
    });
  }

  private clearTimer(model: monaco.editor.ITextModel): void {
    const uri = model.uri.toString();
    const timer = this.timers.get(uri);
    if (timer !== undefined) {
      clearTimeout(timer);
      this.timers.delete(uri);
    }
  }

  private async fetch(model: monaco.editor.ITextModel): Promise<void> {
    const id = this.idOf(model);
    if (!id || id.face) {
      return;
    }

    const version = model.getVersionId();
    const answer = await this.ask({
      action: "blame",
      module: id.module,
      ...(id.project ? { project: id.project } : {}),
    });

    // Turned off, edited past, or closed while the host was looking: the answer describes a
    // text that is gone, and painting it would put the wrong author on the wrong line.
    if (model.isDisposed() || !this.on.has(docKeyOf(id.module, id.project)) || model.getVersionId() !== version) {
      return;
    }

    // A refusal turns the layer OFF for the document and says why: a button left pressed over a
    // module with nothing painted claims a state the host just declined, and the words are the
    // only way to learn that there is no repository, no git, or nothing committed yet.
    if (typeof answer.error === "string") {
      this.on.delete(docKeyOf(id.module, id.project));
      this.clearTimer(model);
      this.decorate(model, null);
      this.refusals.set(model.uri.toString(), answer.error);
      this.onError?.(answer.error);
      this.onChanged?.();
      return;
    }

    const reading = blameReadingOf(answer);
    this.refusals.delete(model.uri.toString());
    this.readings.set(model.uri.toString(), { version, reading });
    this.decorate(model, reading);
    this.onChanged?.();
  }

  private decorate(model: monaco.editor.ITextModel, reading: BlameReading | null): void {
    const key = model.uri.toString();
    const lineCount = model.getLineCount();
    const decorations: monaco.editor.IModelDeltaDecoration[] = [];

    // At the END of the line, injected rather than appended: the text underneath is untouched,
    // the caret never stops on it, and the hint cannot be selected or copied out with the code.
    const at = (line: number, content: string, className: string, hover: string | null): void => {
      if (line < 1 || line > lineCount) {
        return;
      }

      const column = model.getLineMaxColumn(line);
      decorations.push({
        range: new monaco.Range(line, column, line, column),
        options: {
          after: {
            content,
            inlineClassName: className,
            cursorStops: monaco.editor.InjectedTextCursorStops.None,
          },
          // LOAD-BEARING: the model drops an injected text whose range is empty unless the
          // decoration says so (textModel getInjectedTextInInterval), and a hint pinned to the
          // end of a line is exactly an empty range. Without this the decorations sit on the
          // model, count as painted, and the view never draws one (measured headless, 2026-09-08).
          showIfCollapsed: true,
          ...(hover ? { hoverMessage: { value: hover } } : {}),
          stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
        },
      });
    };

    if (reading) {
      for (const row of reading.lines) {
        const hint = blameHint(row);
        if (hint) {
          at(row.line, hint, "xlide-blame", blameHover(row));
        }
      }

      for (const line of reading.uncommitted) {
        at(line, UNCOMMITTED_HINT, "xlide-blame-uncommitted", "Not committed");
      }
    }

    this.ids.set(key, model.deltaDecorations(this.ids.get(key) ?? [], decorations));
    if (!reading) {
      this.ids.delete(key);
      this.readings.delete(key);
    }
  }
}
