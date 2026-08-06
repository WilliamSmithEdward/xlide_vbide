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
import { DocumentStore } from "./documents.js";
import { SearchWidget } from "./searchwidget.js";
import { registerFormatting } from "./format.js";
import { currentSettings } from "./settings.js";
import { openPanesMenu, openSettingsDialog } from "./settingsdialog.js";
import { openHelpDialog } from "./helpdialog.js";
import { Bookmarks } from "./bookmarks.js";
import { bootObjectBrowserPage } from "./objectbrowser.js";
import { Shell } from "./shell.js";
import { defineThemes, preferredTheme, watchPreferredTheme } from "./theme.js";
import { installTypingAutomation } from "./typing.js";
import { Workspace } from "./workspace.js";
import { VBA_LANGUAGE_ID, registerVba } from "./vba.js";

// Stamped by the build; reported to the host so the log names the running bundle.
declare const __XLIDE_BUILD__: string;

// Read out of Directory.Build.props at build time, so the surface and the add-in cannot disagree
// about which release this is.
declare const __XLIDE_VERSION__: string;

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
  const editorArea = document.getElementById("editor-area");
  const emptyView = document.getElementById("empty-view");
  if (!editorArea || !emptyView) {
    throw new Error("missing #editor-area or #empty-view");
  }

  // The wordmark's version, filled in here rather than baked into the skeleton, so the number
  // lives in exactly one place in the bundle.
  const brandVersion = document.getElementById("brand-version");
  if (brandVersion) {
    brandVersion.textContent = __XLIDE_VERSION__;
  }

  // A page exception is invisible without a DevTools client attached, which is exactly the
  // situation during a live test: the surface misbehaves, the shim log says nothing, and the
  // developer is left describing symptoms. Both failure channels are forwarded to the host,
  // where they land in the log beside everything else that happened at that moment. Errors
  // only - never ordinary console noise, which would drown the log it is trying to help.
  //
  // BOUNDED, because this ships to users. The host log collapses consecutive IDENTICAL lines
  // but not ones that vary, and every forwarded error is also a message across the bridge
  // handled on the host's user interface thread - so a fault thrown once per animation frame
  // with a changing value in its text would flood the log and push the editor around while
  // doing it. The first errors are the ones that explain a session anyway: the first failure
  // usually causes the rest.
  const errorBudget = 20;
  let errorsForwarded = 0;
  let lastReported = "";

  const reportToHost = (what: string, detail: string) => {
    const line = `page error: ${what}: ${detail}`;

    // A repeat costs nothing and buys nothing; it must not spend the budget either.
    if (line === lastReported || errorsForwarded > errorBudget) {
      return;
    }

    lastReported = line;
    errorsForwarded += 1;

    try {
      bridge?.trace(errorsForwarded > errorBudget
        ? `page error: too many page errors this session; the rest are in DevTools only`
        : line);
    } catch {
      // The bridge may not exist yet; a page that fails before it is built is a page that
      // never reached ready, which the loader already reports.
    }
  };

  window.addEventListener("error", (event) => {
    const where = event.filename ? ` (${event.filename}:${event.lineno}:${event.colno})` : "";
    reportToHost("uncaught", `${event.message}${where}\n${event.error?.stack ?? ""}`.trim());
  });

  window.addEventListener("unhandledrejection", (event) => {
    const reason = event.reason;
    reportToHost("unhandled rejection",
      reason instanceof Error ? `${reason.message}\n${reason.stack ?? ""}`.trim() : String(reason));
  });

  registerVba();
  defineThemes();

  // The formatter reads the developer's choices at each run, so a change in the settings
  // dialog is the very next Format Module's behaviour.
  registerFormatting(() => {
    const settings = currentSettings();
    return {
      indentSize: settings.formatIndentSize,
      useTabs: settings.formatUseTabs,
      canonicalKeywords: settings.formatCanonicalKeywords,
    };
  });

  // What every group's editor is created with. The companion editor's minimap, its settings
  // included: blocks rather than characters, the slider always visible (the developer's
  // request: it is the indicator, not a control to be discovered). The find widget floats
  // over the text instead of reserving a band above it (2026-08-04).
  const editorOptions: monaco.editor.IStandaloneEditorConstructionOptions = {
    value: "",
    language: VBA_LANGUAGE_ID,
    theme: preferredTheme(),
    automaticLayout: true,
    glyphMargin: true,
    lineNumbersMinChars: 4,
    minimap: {
      enabled: true,
      renderCharacters: false,
      showMarkSectionHeaders: false,
      showRegionSectionHeaders: false,
      showSlider: "always",
    },
    scrollBeyondLastLine: false,
    find: {
      addExtraSpaceOnTop: false,
    },
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
  };

  const transport = webView2Transport();
  const documents = new DocumentStore();
  const bookmarks = new Bookmarks();

  // Bridge, workspace, and shell reference each other, so they are built in dependency
  // order and stitched by assignment: the bridge first (it only needs the transport), the
  // workspace next (its editor factory wires each new editor into everything), the search
  // widget on the workspace's active editor, and the shell last — its toolbar keeps only
  // the commands that resolve as actions at build time, so every per-editor action must be
  // registered before it looks. Nothing host-driven runs before bridge.start().
  const bridge = new EditorBridge(transport ?? demoTransport(), documents);

  let searchWidget: SearchWidget;
  let workspace: Workspace;
  // Declared before the workspace: its constructor announces the first active group, and the
  // callback below must find an undefined shell, not a const still in its dead zone.
  let shell: Shell | undefined;

  /** Everything a new group's editor gets, the moment the workspace creates it. */
  const wireEditor = (editor: monaco.editor.IStandaloneCodeEditor): void => {
    bridge.attachEditor(editor);
    searchWidget.registerOn(editor);
    installTypingAutomation(editor, bridge);
    bookmarks.attach(editor);
    registerHostActions(editor, bridge);
    installMarginMenu(editor, bridge);
  };

  searchWidget = new SearchWidget(() => workspace.activeEditor(), {
    search: (query, matchCase, wholeWord, scope) => bridge.requestSearch(query, matchCase, wholeWord, scope),
    replaceAll: (query, matchCase, wholeWord, scope, replacement) =>
      bridge.requestReplaceAll(query, matchCase, wholeWord, scope, replacement),
    navigate: (module, line, column, selectLine, workbook) =>
      bridge.navigate(module, line, column, selectLine, workbook),
  });
  bridge.searchWidget = searchWidget;

  workspace = new Workspace(editorArea, emptyView, documents, {
    createEditor: (groupBody) => {
      const editor = monaco.editor.create(groupBody, editorOptions);
      wireEditor(editor);
      return editor;
    },
    activate: (id) => bridge.activateModule(id.module, id.project ?? undefined),
    close: (id, action) => bridge.closeModule(id.module, id.project ?? undefined, action),
    activeChanged: (id, editor) => {
      // The search widget floats over the active group and searches its editor.
      searchWidget.attachTo(editor.getContainerDomNode());
      searchWidget.onActiveEditorChanged();
      shell?.setActiveModule(id?.module ?? null, id?.project ?? null);
    },
    layoutChanged: () => workspace?.editors().forEach((editor) => editor.layout()),
  });
  bridge.workspace = workspace;

  const createMs = performance.now();

  // Automatic layout rides ResizeObserver and tracks the window live; the settle here is
  // only the safety net for a final frame the observer missed. It waits for the resize to
  // pause — running it per event doubled every layout of a drag, which read as latency and
  // churn (2026-08-05) — and a measure that finds nothing changed costs nothing.
  //
  // The live-resize class is the minimap's peace: its canvas repaints a frame behind the
  // layout that moved it, so during a drag its blocks were alternately stale, clipped, and
  // redrawn — a flicker at the right edge. Faded out while events stream and back in at
  // the settle, the same deliberate quiet native apps keep during a live resize.
  let resizeSettled: ReturnType<typeof setTimeout> | undefined;
  window.addEventListener("resize", () => {
    document.body.classList.add("live-resize");
    clearTimeout(resizeSettled);
    resizeSettled = setTimeout(() => {
      workspace.editors().forEach((editor) => editor.layout());
      document.body.classList.remove("live-resize");
    }, 150);
  });

  shell = new Shell(document.body, {
    activateModule: (name, workbook) => bridge.activateModule(name, workbook),
    moduleIsOpen: (name) => bridge.documents.all()
      .some((id) => id.module.toLowerCase() === name.toLowerCase()),
    navigate: (module, line, column, selectLine, workbook) =>
      bridge.navigate(module, line, column, selectLine, workbook),
    layoutChanged: () => workspace.editors().forEach((editor) => editor.layout()),
    command: (command) => {
      // The settings dialog is the page's own, not a Monaco action and not the host's.
      if (command.id === "openSettings") {
        openSettingsDialog((next) => bridge.updateSettings(next), () => workspace.activeEditor().focus());
        return;
      }

      // The Panes menu: its own dropdown under its own toolbar button, beside settings
      // (developer, 2026-08-06). Showing and hiding a pane is done while working, not
      // visited once like a preference.
      if (command.id === "openHelp") {
        openHelpDialog(() => workspace.activeEditor().focus());
        return;
      }

      if (command.id === "openPanes") {
        const button = document.querySelector<HTMLElement>('#toolbar [data-command="openPanes"]');
        if (button && shell) {
          openPanesMenu(shell.paneVisibility(), button);
        }
        return;
      }

      bridge.runCommand(command);
    },
    // Undo and redo are built in rather than registered, so they never resolve as actions and
    // would be dropped by a check that only knows about registered ones. The settings dialog
    // and the Panes menu are the page's own and always exist.
    commandAvailable: (command) =>
      command.id === "undo" || command.id === "redo"
      || command.id === "openSettings" || command.id === "openPanes" || command.id === "openHelp"
      || workspace.activeEditor().getAction(command.id) !== null,
    evaluate: (text) => bridge.evaluate(text),
    panelChanged: (name, open) => bridge.panelChanged(name, open),
    menuRequest: (path) => bridge.requestMenu(path),
    menuExecute: (path) => bridge.executeMenu(path),
    menuClosed: () => workspace.activeEditor().focus(),
    editProperty: (component, name, value) => bridge.editProperty(component, name, value),
    selectComponent: (name) => bridge.selectComponent(name),
    closeModule: (name, workbook, action) => bridge.closeModule(name, workbook, action),
    insertComponent: (kind, project) => bridge.insertComponent(kind, project),
    requestOutline: (module, workbook) => bridge.requestOutline(module, workbook),
    trace: (text) => bridge.trace(text),
  });
  bridge.shell = shell;

  // Ctrl+W closes the active group's active tab from anywhere in the surface. The host's key
  // hook claims it first when it is listening; this is the page's own answer for every moment
  // it is not, so the shortcut never depends on which corner of the surface has focus.
  document.addEventListener("keydown", (event) => {
    if (!event.ctrlKey || event.shiftKey || event.altKey) {
      return;
    }

    // Ctrl+\ splits, the studio's own key for it.
    if (event.code === "Backslash") {
      event.preventDefault();
      workspace.splitActive("right");
      return;
    }

    if (event.code !== "KeyW" && event.code !== "F4") {
      return;
    }

    event.preventDefault();
    workspace.closeActive();
  }, { capture: true });

  // Reachable from a devtools console, which is how the page half of a host defect gets isolated
  // from the transport half.
  (globalThis as { xlideBridge?: EditorBridge }).xlideBridge = bridge;

  // Tools > Options routes here from the host: the native Options dialog is superseded, and
  // the product's settings are where the choices that matter live.
  bridge.openSettings = () =>
    openSettingsDialog((next) => bridge.updateSettings(next), () => workspace.activeEditor().focus());

  // Completions come from the host's engine: the analyzer that verified the Excel object model
  // decides what a receiver offers, and the page only renders the answer. Triggered on the dot
  // for member access, and by ordinary typing for identifiers and keywords.
  //
  // Engine requests are offset-only against the HOST-ACTIVE module (decision 12), so every
  // provider answers only for its model. A background group's model gets no engine answers —
  // honest, where an answer computed against the wrong module's text would not be.
  monaco.languages.registerCompletionItemProvider(VBA_LANGUAGE_ID, {
    triggerCharacters: ["."],
    provideCompletionItems: async (model, position) => {
      if (model !== bridge.hostActiveModel()) {
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
      if (model !== bridge.hostActiveModel()) {
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

  // Call tips, triggered the way the extension triggers them: the opening paren, the comma, and
  // the space, because VBA calls procedures without parentheses too.
  monaco.languages.registerSignatureHelpProvider(VBA_LANGUAGE_ID, {
    signatureHelpTriggerCharacters: ["(", ",", " "],
    signatureHelpRetriggerCharacters: [","],
    provideSignatureHelp: async (model, position) => {
      if (model !== bridge.hostActiveModel()) {
        return null;
      }

      const info = await bridge.requestSignatureHelp(model.getOffsetAt(position));
      if (!info) {
        return null;
      }

      const signature: monaco.languages.SignatureInformation = {
        label: info.label,
        parameters: info.parameters.map((parameter) => ({
          label: parameter.label,
          ...(parameter.documentation ? { documentation: { value: parameter.documentation } } : {}),
        })),
        ...(info.documentation ? { documentation: { value: info.documentation } } : {}),
      };

      return {
        value: {
          signatures: [signature],
          activeSignature: 0,
          activeParameter: info.activeParameter,
        },
        dispose: () => { },
      };
    },
  });

  watchPreferredTheme((theme) => bridge.applyOsTheme(theme));

  if (!transport) {
    console.log("[xlide demo] window.chrome.webview absent, running the loopback demo");
  }

  // The bundle's own resource entry splits the two costs that scriptMs lumps together:
  // everything before responseEnd is fetching, everything after is compiling and running.
  // A transfer size of zero is the browser's cache answering — the number that says whether
  // a second boot is allowed to be cheaper than the first.
  const bundleEntry = performance
    .getEntriesByType("resource")
    .find((entry) => entry.name.endsWith("/editor.js")) as PerformanceResourceTiming | undefined;
  const navigationEntry = performance
    .getEntriesByType("navigation")[0] as PerformanceNavigationTiming | undefined;

  bridge.start({
    scriptMs: Math.round(scriptMs),
    createMs: Math.round(createMs - scriptMs),
    totalMs: Math.round(performance.now()),
    build: __XLIDE_BUILD__,
    ...(bundleEntry
      ? {
        fetchMs: Math.round(bundleEntry.responseEnd),
        transferBytes: Math.round(bundleEntry.transferSize),
        requestStartMs: Math.round(bundleEntry.requestStart),
      }
      : {}),
    ...(navigationEntry ? { htmlMs: Math.round(navigationEntry.responseEnd) } : {}),
  });

  // After ready, so the reply cannot arrive before the host considers the page up. The bar needs
  // its top-level items before anything is clicked: they carry the Alt accelerators.
  shell.requestMenus();
}

/**
 * The host's own commands, present in each editor's context menu and the command palette so
 * they are discoverable where a developer already looks for commands. The navigation pair
 * rode the View menu until it went (2026-08-05); their keys are claimed host-side while the
 * surface has focus, and bound here as well for every moment it does not. Registered on
 * every group's editor, so the palette works wherever it opens.
 */
function registerHostActions(editor: monaco.editor.IStandaloneCodeEditor, bridge: EditorBridge): void {
  const hostActions: Array<[string, string, string, number[]?]> = [
    ["xlide.run", "Run Sub/UserForm (F5)", "run"],
    ["xlide.toggleBreakpoint", "Toggle Breakpoint (F9)", "toggleBreakpoint"],
    ["xlide.runToCursor", "Run To Cursor (Ctrl+F8)", "runToCursor"],
    ["xlide.goToDefinition", "Go to Definition (Shift+F2)", "goToDefinition",
      [monaco.KeyMod.Shift | monaco.KeyCode.F2]],
    ["xlide.lastPosition", "Last Position (Ctrl+Shift+F2)", "lastPosition",
      [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.F2]],
  ];

  for (const [id, label, command, keys] of hostActions) {
    editor.addAction({
      id,
      label,
      contextMenuGroupId: "1_xlide",
      contextMenuOrder: 1,
      ...(keys ? { keybindings: keys } : {}),
      run: () => bridge.runCommand({ id: command, target: "host", icon: "", label }),
    });
  }
}

/**
 * The margin's own menu, per editor. The editor would otherwise offer its text menu there,
 * which is a menu for a place the click was not.
 */
function installMarginMenu(editor: monaco.editor.IStandaloneCodeEditor, bridge: EditorBridge): void {
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

// One bundle, two documents: the editor surface, and the Object Browser palette the host
// opens in its own floating window. The palette wants none of the editor's machinery —
// no Monaco boot, no shell, no bridge — so it takes its own door before any of that starts.
const entry = new URLSearchParams(window.location.search).get("view") === "objbrowser"
  ? bootObjectBrowserPage
  : boot;

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", entry, { once: true });
} else {
  entry();
}
