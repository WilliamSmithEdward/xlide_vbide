/*
 * Bookmarks, the surface's own. The native editor's bookmarks live in the pane the surface
 * covers, and their menu — the Edit menu's last survivor — is gone (2026-08-05), so the job
 * moved here whole: toggle a mark on a line, hop between marks, clear them all.
 *
 * The marks are Monaco decorations, which ride the text as it is edited, so bookmarks shift
 * with their lines for free. The surface shows one model at a time and replaces it on every
 * module switch, so each model's marks are captured as it goes — onWillDispose, the last
 * moment its decoration positions exist — and restored when its module returns, keyed by the
 * model's URI, which names the module.
 *
 * Like the native editor's, they last as long as the session; nothing persists them.
 */

import * as monaco from "monaco-editor/editor/editor.api.js";

export function installBookmarks(editor: monaco.editor.IStandaloneCodeEditor): void {
  /** Lines by model URI, for models not currently shown. */
  const held = new Map<string, number[]>();

  /** Live decoration ids on the current model. */
  let ids: string[] = [];

  let watching: monaco.IDisposable | null = null;

  const lines = (): number[] => {
    const model = editor.getModel();
    if (!model) {
      return [];
    }

    const found = ids
      .map((id) => model.getDecorationRange(id)?.startLineNumber)
      .filter((line): line is number => line !== undefined);
    return [...new Set(found)].sort((a, b) => a - b);
  };

  const decorate = (marks: number[]): void => {
    const model = editor.getModel();
    if (!model) {
      ids = [];
      return;
    }

    ids = model.deltaDecorations(ids, marks.map((line) => ({
      range: new monaco.Range(line, 1, line, 1),
      options: {
        glyphMarginClassName: "xlide-bookmark codicon codicon-bookmark",
        glyphMarginHoverMessage: { value: "Bookmark" },
        stickiness: monaco.editor.TrackedRangeStickiness.NeverGrowsWhenTypingAtEdges,
      },
    })));
  };

  const adopt = (): void => {
    watching?.dispose();
    watching = null;
    ids = [];

    const model = editor.getModel();
    if (!model) {
      return;
    }

    decorate(held.get(model.uri.toString()) ?? []);
    watching = model.onWillDispose(() => {
      held.set(model.uri.toString(), lines());
    });
  };

  editor.onDidChangeModel(adopt);
  adopt();

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
      const at = editor.getPosition()?.lineNumber;
      if (at === undefined) {
        return;
      }

      const marks = lines();
      decorate(marks.includes(at) ? marks.filter((line) => line !== at) : [...marks, at]);
    },
  });

  editor.addAction({
    id: "xlide.bookmark.next",
    label: "Next Bookmark",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Alt | monaco.KeyCode.KeyN],
    run: () => {
      const marks = lines();
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
      const marks = lines();
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
      // Every module's, the way the native command cleared them: the held map and the
      // marks on the model in hand.
      held.clear();
      decorate([]);
    },
  });
}
