/*
 * Bookmarks, the surface's own. The native editor's bookmarks live in the pane the surface
 * covers, and their menu - the Edit menu's last survivor - is gone (2026-08-05), so the job
 * moved here whole: toggle a mark on a line, hop between marks, clear them all.
 *
 * The marks are Monaco decorations, which ride the text as it is edited, so bookmarks shift
 * with their lines for free. Models stay alive for as long as their modules are open
 * (decision 12), so a live model's marks simply persist on it; a model being disposed - its
 * pane closed - has its lines captured at the last moment they exist, and restored if the
 * module returns, keyed by the model's URI.
 *
 * One store serves every editor group: the marks belong to the DOCUMENT, and jumping walks
 * the marks of whichever document the acting editor shows. Like the native editor's, they
 * last as long as the session; nothing persists them.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";

export class Bookmarks {
  /** Lines by model URI, for models that have been disposed. */
  private readonly held = new Map<string, number[]>();

  /** Live decoration ids per model URI. */
  private readonly ids = new Map<string, string[]>();

  /** Models already being watched for disposal, so an editor swap does not double-register. */
  private readonly watched = new WeakSet<monaco.editor.ITextModel>();

  /** Current bookmark lines of a model, read from its live decorations. */
  private lines(model: monaco.editor.ITextModel): number[] {
    const found = (this.ids.get(model.uri.toString()) ?? [])
      .map((id) => model.getDecorationRange(id)?.startLineNumber)
      .filter((line): line is number => line !== undefined);
    return [...new Set(found)].sort((a, b) => a - b);
  }

  private decorate(model: monaco.editor.ITextModel, marks: number[]): void {
    const key = model.uri.toString();
    this.ids.set(key, model.deltaDecorations(this.ids.get(key) ?? [], marks.map((line) => ({
      range: new monaco.Range(line, 1, line, 1),
      options: {
        glyphMarginClassName: "xlide-bookmark codicon codicon-bookmark",
        glyphMarginHoverMessage: { value: "Bookmark" },
        stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
      },
    }))));
  }

  /**
   * The marks on a model right now, for the debug api's `ui` route.
   *
   * Read from the live decorations rather than from the held map, because the decorations ARE
   * the truth: they ride the text as it is edited, so the held copy is only what was last
   * saved off. A reporter that read the map would describe where the marks were put, not where
   * they are - which is the whole behaviour worth checking.
   */
  marksOn(model: monaco.editor.ITextModel): number[] {
    return this.lines(model);
  }

  /** Adopts a model: restores held marks if its module was closed before, watches disposal. */
  adopt(model: monaco.editor.ITextModel): void {
    if (this.watched.has(model)) {
      return;
    }

    this.watched.add(model);
    const key = model.uri.toString();

    const restored = this.held.get(key);
    if (restored && restored.length > 0) {
      this.held.delete(key);
      this.decorate(model, restored);
    }

    model.onWillDispose(() => {
      const marks = this.lines(model);
      if (marks.length > 0) {
        this.held.set(key, marks);
      }
      this.ids.delete(key);
    });
  }

  /** Wires the bookmark actions into one editor. Called once per editor group. */
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

    const jump = (line: number): void => {
      editor.setPosition({ lineNumber: line, column: 1 });
      editor.revealLineInCenterIfOutsideViewport(line);
      editor.focus();
    };

    editor.addAction({
      id: "xlide.bookmark.toggle",
      label: "Toggle Bookmark",
      keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Alt | monaco.KeyCode.KeyK],
      contextMenuGroupId: "1_xlide",
      contextMenuOrder: 2,
      run: () => {
        const target = editor.getModel();
        const at = editor.getPosition()?.lineNumber;
        if (!target || at === undefined) {
          return;
        }

        const marks = this.lines(target);
        this.decorate(target, marks.includes(at) ? marks.filter((line) => line !== at) : [...marks, at]);
      },
    });

    editor.addAction({
      id: "xlide.bookmark.next",
      label: "Next Bookmark",
      keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Alt | monaco.KeyCode.KeyN],
      run: () => {
        const target = editor.getModel();
        if (!target) {
          return;
        }

        const marks = this.lines(target);
        const at = editor.getPosition()?.lineNumber ?? 0;
        const next = marks.find((line) => line > at) ?? marks[0];
        if (next !== undefined) {
          jump(next);
        }
      },
    });

    editor.addAction({
      id: "xlide.bookmark.previous",
      label: "Previous Bookmark",
      keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Alt | monaco.KeyCode.KeyP],
      run: () => {
        const target = editor.getModel();
        if (!target) {
          return;
        }

        const marks = this.lines(target);
        const at = editor.getPosition()?.lineNumber ?? 0;
        const previous = [...marks].reverse().find((line) => line < at) ?? marks[marks.length - 1];
        if (previous !== undefined) {
          jump(previous);
        }
      },
    });

    editor.addAction({
      id: "xlide.bookmark.clearAll",
      label: "Clear All Bookmarks",
      keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Alt | monaco.KeyMod.Shift | monaco.KeyCode.KeyK],
      run: () => {
        // Every module's, the way the native command cleared them: the held map, and the
        // marks on every model that still lives.
        this.held.clear();
        for (const model of monaco.editor.getModels()) {
          if (this.ids.has(model.uri.toString())) {
            this.decorate(model, []);
          }
        }
      },
    });
  }
}
