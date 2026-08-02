import * as monaco from "monaco-editor/editor/editor.api.js";

// monaco-editor 0.56 ships contributions as opt-in feature registrations; editor.api.js alone
// is a bare widget. Only the ones this surface actually uses are pulled in, which is what
// keeps the bundle from growing to the full editor.main.js size.
import "monaco-editor/features/bracketMatching/register.js";
import "monaco-editor/features/clipboard/register.js";
import "monaco-editor/features/codicon/register.js";
import "monaco-editor/features/comment/register.js";
import "monaco-editor/features/contextmenu/register.js";
import "monaco-editor/features/cursorUndo/register.js";
import "monaco-editor/features/find/register.js";
import "monaco-editor/features/folding/register.js";
import "monaco-editor/features/format/register.js";
import "monaco-editor/features/gotoError/register.js";
import "monaco-editor/features/gotoLine/register.js";
import "monaco-editor/features/hover/register.js";
import "monaco-editor/features/indentation/register.js";
import "monaco-editor/features/lineSelection/register.js";
import "monaco-editor/features/linesOperations/register.js";
import "monaco-editor/features/multicursor/register.js";
import "monaco-editor/features/parameterHints/register.js";
import "monaco-editor/features/quickCommand/register.js";
import "monaco-editor/features/smartSelect/register.js";
// Suggest is the completion widget itself; snippet expands the placeholders completions insert.
// A completion provider without these registered answers into a void.
//
// The suggest controller is imported by its own path, deliberately. The feature's register
// module imports ONLY the inline-completions adapter and never the controller, so registering
// the feature grows the bundle and changes nothing on screen; this was found by asking a
// running editor for editor.action.triggerSuggest and being told there is no such action.
import "monaco-editor/features/snippet/register.js";
import "monaco-editor/features/suggest/register.js";
import "monaco-editor/editor/contrib/suggest/browser/suggestController.js";
import "monaco-editor/features/tokenization/register.js";
import "monaco-editor/features/wordHighlighter/register.js";
import "monaco-editor/features/wordOperations/register.js";
import "monaco-editor/features/wordPartOperations/register.js";

import "./styles.css";
import { EditorBridge, demoTransport, webView2Transport, type HostCompletionItem } from "./bridge.js";
import { showContextMenu } from "./contextmenu.js";
import { DEFAULT_FORMAT_OPTIONS, registerFormatting } from "./format.js";
import { Shell } from "./shell.js";
import { defineThemes, preferredTheme, watchPreferredTheme } from "./theme.js";
import { VBA_LANGUAGE_ID, registerVba } from "./vba.js";

// Stamped by the build; reported to the host so the log names the running bundle.
declare const __XLIDE_BUILD__: string;

// The worker is a sibling of index.html in dist, addressed relative to the document base so a
// virtual host mapping, a static server and a subdirectory deployment all resolve the same.
globalThis.MonacoEnvironment = {
  getWorker: () => new Worker(new URL("./editor.worker.js", document.baseURI), { type: "module" }),
};

// Taken at the top of the module body, which esbuild places after every import has been evaluated.
// So this is the cost of fetching, parsing and running the whole bundle, and it is the number that
// decides whether the surface is worth putting over a pane at all.
const scriptMs = performance.now();

function boot(): void {
  const container = document.getElementById("container");
  if (!container) {
    throw new Error("missing #container");
  }

  registerVba();
  defineThemes();
  registerFormatting(() => DEFAULT_FORMAT_OPTIONS);

  const editor = monaco.editor.create(container, {
    value: "",
    language: VBA_LANGUAGE_ID,
    theme: preferredTheme(),
    automaticLayout: true,
    glyphMargin: true,
    lineNumbersMinChars: 4,
    minimap: { enabled: false },
    scrollBeyondLastLine: false,
    renderLineHighlight: "line",
    renderWhitespace: "selection",
    fontFamily: '"Cascadia Mono", Consolas, "Courier New", monospace',
    fontSize: 13,
    tabSize: 4,
    insertSpaces: true,
    detectIndentation: false,
    autoIndent: "full",
    wordWrap: "off",
    smoothScrolling: false,
    fixedOverflowWidgets: true,
  });

  const createMs = performance.now();

  const transport = webView2Transport();

  // The shell is built before the bridge, because the bridge routes host messages into it. The
  // handlers close over the bridge, which does not exist yet, so they reach it through a variable
  // that is assigned immediately below: nothing can call them before then, because both a tab and
  // a finding need the host to have sent something first.
  let bridge: EditorBridge;
  const shell = new Shell(document.body, {
    activateModule: (name) => bridge.activateModule(name),
    navigate: (module, line, column) => bridge.navigate(module, line, column),
    layoutChanged: () => editor.layout(),
    command: (command) => bridge.runCommand(command),
    // Undo and redo are built in rather than registered, so they never resolve as actions and
    // would be dropped by a check that only knows about registered ones.
    commandAvailable: (command) =>
      command.id === "undo" || command.id === "redo" || editor.getAction(command.id) !== null,
    evaluate: (text) => bridge.evaluate(text),
    panelChanged: (name, open) => bridge.panelChanged(name, open),
    menuRequest: (path) => bridge.requestMenu(path),
    menuExecute: (path) => bridge.executeMenu(path),
    menuClosed: () => editor.focus(),
    editProperty: (component, name, value) => bridge.editProperty(component, name, value),
    selectComponent: (name) => bridge.selectComponent(name),
    closeModule: (name) => bridge.closeModule(name),
    insertComponent: (kind) => bridge.insertComponent(kind),
  });

  bridge = new EditorBridge(editor, transport ?? demoTransport(), shell);

  // Reachable from a devtools console, which is how the page half of a host defect gets isolated
  // from the transport half.
  (globalThis as { xlideBridge?: EditorBridge }).xlideBridge = bridge;

  // Completions come from the host's engine: the analyzer that verified the Excel object model
  // decides what a receiver offers, and the page only renders the answer. Triggered on the dot
  // for member access, and by ordinary typing for identifiers and keywords.
  monaco.languages.registerCompletionItemProvider(VBA_LANGUAGE_ID, {
    triggerCharacters: ["."],
    provideCompletionItems: async (model, position) => {
      if (model !== editor.getModel()) {
        return { suggestions: [] };
      }

      const items = await bridge.requestCompletions(model.getOffsetAt(position));
      const word = model.getWordUntilPosition(position);
      const range = new monaco.Range(
        position.lineNumber,
        word.startColumn,
        position.lineNumber,
        word.endColumn);

      return { suggestions: items.map((item) => toSuggestion(item, range)) };
    },
  });

  // Hovers come from the same engine: the identifier under the cursor described by its
  // declaration line, its origin, and its documentation. The signature renders as VBA code, the
  // way the extension renders it.
  monaco.languages.registerHoverProvider(VBA_LANGUAGE_ID, {
    provideHover: async (model, position) => {
      if (model !== editor.getModel()) {
        return null;
      }

      const hover = await bridge.requestHover(model.getOffsetAt(position));
      if (!hover) {
        return null;
      }

      const start = model.getPositionAt(hover.start);
      const end = model.getPositionAt(hover.end);
      const contents: monaco.IMarkdownString[] = [
        { value: "```vba\n" + hover.signature + "\n```" },
      ];

      if (hover.details.length > 0) {
        contents.push({ value: hover.details.join("  \n") });
      }

      if (hover.documentation) {
        contents.push({ value: hover.documentation });
      }

      return {
        range: new monaco.Range(start.lineNumber, start.column, end.lineNumber, end.column),
        contents,
      };
    },
  });

  // The host's own commands, present in the editor's context menu and the command palette so
  // they are discoverable where a developer already looks for commands.
  const hostActions: Array<[string, string, string]> = [
    ["xlide.run", "Run Sub/UserForm (F5)", "run"],
    ["xlide.toggleBreakpoint", "Toggle Breakpoint (F9)", "toggleBreakpoint"],
    ["xlide.runToCursor", "Run To Cursor (Ctrl+F8)", "runToCursor"],
  ];

  for (const [id, label, command] of hostActions) {
    editor.addAction({
      id,
      label,
      contextMenuGroupId: "1_xlide",
      contextMenuOrder: 1,
      run: () => bridge.runCommand({ id: command, target: "host", icon: "", label }),
    });
  }

  // The margin's own menu. The editor would otherwise offer its text menu there, which is a menu
  // for a place the click was not.
  editor.onContextMenu((event) => {
    const kind = event.target.type;
    if (kind !== monaco.editor.MouseTargetType.GUTTER_GLYPH_MARGIN
      && kind !== monaco.editor.MouseTargetType.GUTTER_LINE_NUMBERS) {
      return;
    }

    event.event.preventDefault();
    event.event.stopPropagation();

    const line = event.target.position?.lineNumber;
    showContextMenu(event.event.posx, event.event.posy, [
      {
        label: "Toggle Breakpoint",
        enabled: line !== undefined,
        run: () => bridge.toggleBreakpoint(line ?? 1),
      },
      {
        label: "Clear All Breakpoints",
        run: () => bridge.runCommand({
          id: "clearAllBreakpoints",
          target: "host",
          icon: "",
          label: "Clear All Breakpoints",
        }),
      },
    ]);
  });

  watchPreferredTheme((theme) => bridge.applyOsTheme(theme));

  if (!transport) {
    console.log("[xlide demo] window.chrome.webview absent, running the loopback demo");
  }

  bridge.start({
    scriptMs: Math.round(scriptMs),
    createMs: Math.round(createMs - scriptMs),
    totalMs: Math.round(performance.now()),
    build: __XLIDE_BUILD__,
  });

  // After ready, so the reply cannot arrive before the host considers the page up. The bar needs
  // its top-level items before anything is clicked: they carry the Alt accelerators.
  shell.requestMenus();
}

/** The analyzer's completion kinds, mapped onto the editor's icons. */
const COMPLETION_KINDS: Record<string, monaco.languages.CompletionItemKind> = {
  method: monaco.languages.CompletionItemKind.Method,
  property: monaco.languages.CompletionItemKind.Property,
  event: monaco.languages.CompletionItemKind.Event,
  constant: monaco.languages.CompletionItemKind.Constant,
  enum: monaco.languages.CompletionItemKind.Enum,
  enumMember: monaco.languages.CompletionItemKind.EnumMember,
  global: monaco.languages.CompletionItemKind.Variable,
  codeName: monaco.languages.CompletionItemKind.File,
  variable: monaco.languages.CompletionItemKind.Variable,
  parameter: monaco.languages.CompletionItemKind.Variable,
  value: monaco.languages.CompletionItemKind.Value,
  procedure: monaco.languages.CompletionItemKind.Function,
  function: monaco.languages.CompletionItemKind.Function,
  runtime: monaco.languages.CompletionItemKind.Function,
  module: monaco.languages.CompletionItemKind.Module,
  type: monaco.languages.CompletionItemKind.Struct,
  object: monaco.languages.CompletionItemKind.Class,
  collection: monaco.languages.CompletionItemKind.Class,
  keyword: monaco.languages.CompletionItemKind.Keyword,
};

function toSuggestion(item: HostCompletionItem, range: monaco.Range): monaco.languages.CompletionItem {
  const insertText = item.insertText ?? item.label;

  const suggestion: monaco.languages.CompletionItem = {
    label: item.label,
    kind: COMPLETION_KINDS[item.kind] ?? monaco.languages.CompletionItemKind.Text,
    insertText,
    range,
  };

  if (item.detail) {
    suggestion.detail = item.detail;
  }
  if (item.documentation) {
    suggestion.documentation = { value: item.documentation };
  }
  if (item.filterText) {
    suggestion.filterText = item.filterText;
  }
  if (item.sortText) {
    suggestion.sortText = item.sortText;
  }

  // The engine's snippet placeholders use the editor's own snippet syntax; plain text must not
  // be parsed as it, or a literal dollar sign in an insertion would vanish.
  if (insertText.includes("${") || insertText.includes("$0")) {
    suggestion.insertTextRules = monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet;
  }

  return suggestion;
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", boot, { once: true });
} else {
  boot();
}
