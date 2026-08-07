import * as monaco from "monaco-editor/editor/editor.api.js";

// monaco-editor 0.56 ships contributions as opt-in feature registrations; editor.api.js alone
// is a bare widget. Only the ones this surface actually uses are pulled in, which is what
// keeps the bundle from growing to the full editor.main.js size.
import "monaco-editor/features/bracketMatching/register.js";
import "monaco-editor/features/clipboard/register.js";
// The lightbulb and its menu. A code-action provider without this registered answers into a void,
// the same way a completion provider does without suggest.
import "monaco-editor/features/codeAction/register.js";
import "monaco-editor/features/codicon/register.js";
import "monaco-editor/features/comment/register.js";
import "monaco-editor/features/contextmenu/register.js";
import "monaco-editor/features/cursorUndo/register.js";
import "monaco-editor/features/find/register.js";
import "monaco-editor/features/folding/register.js";
import "monaco-editor/features/format/register.js";
import "monaco-editor/features/gotoError/register.js";
import "monaco-editor/features/gotoLine/register.js";
// Ctrl+click on a symbol. The commands the same feature is named for — F12, Shift+F12, the peek
// windows, and their right-click entries — live in a module its register never imports, so they
// come in by their own path. That is the third feature here whose register module covers less
// than its name does; the pattern is worth expecting rather than rediscovering.
import "monaco-editor/features/gotoSymbol/register.js";
import "monaco-editor/editor/contrib/gotoSymbol/browser/goToCommands.js";
import "monaco-editor/features/hover/register.js";
import "monaco-editor/features/indentation/register.js";
import "monaco-editor/features/lineSelection/register.js";
import "monaco-editor/features/linesOperations/register.js";
import "monaco-editor/features/multicursor/register.js";
import "monaco-editor/features/parameterHints/register.js";
import "monaco-editor/features/quickCommand/register.js";
// The rename input box and its F2 binding.
import "monaco-editor/features/rename/register.js";
// What asks a semantic-tokens provider for tokens and paints the answer. The feature's register
// module imports only the VIEWPORT contribution, which serves registerDocumentRangeSemanticTokens
// providers; the whole-document feature is a separate module it never touches, so registering the
// feature alone leaves a document provider registered and never called. Found the same way the
// suggest controller below was: by watching a running editor ask for nothing.
import "monaco-editor/features/semanticTokens/register.js";
import "monaco-editor/editor/contrib/semanticTokens/browser/documentSemanticTokens.js";
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
import { EditorBridge, MARKER_OWNER, demoTransport, webView2Transport, type HostCompletionItem, type HostLocation, type HostRenameAnswer } from "./bridge.js";
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

/**
 * The semantic-token legend, which is the extension's own. An index into these lists is what
 * crosses to the editor, so the order is a contract with the theme rather than a preference:
 * theme.ts colours a token by matching its type and modifiers against its rule scopes.
 */
const SEMANTIC_TOKEN_TYPES = ["class", "enum", "struct", "type", "variable"];

/** `defaultLibrary` marks a host-injected global — Application, ThisWorkbook, ActiveSheet. */
const SEMANTIC_TOKEN_MODIFIERS = ["defaultLibrary"];

/**
 * The host's answers as the editor wants them, dropping any whose module has no model. A location
 * the editor cannot resolve to a model renders as an empty row in the peek window, which reads as
 * a result that is there and says nothing — worse than one result fewer.
 */
function toEditorLocations(
  bridge: EditorBridge,
  locations: readonly HostLocation[],
): monaco.languages.Location[] {
  const out: monaco.languages.Location[] = [];

  for (const location of locations) {
    const model = bridge.modelForLocation(location);
    if (!model) {
      continue;
    }

    out.push({
      uri: model.uri,
      range: new monaco.Range(
        location.line,
        location.column,
        location.line,
        location.column + location.length),
    });
  }

  return out;
}

/**
 * What to tell the developer a rename did. The count of modules matters more than the count of
 * uses: a rename that reached four modules is a rename that reached modules they cannot see, and
 * that is the fact worth putting on screen.
 */
function renameSummary(answer: HostRenameAnswer, newName: string): string {
  const uses = `${answer.replaced} use${answer.replaced === 1 ? "" : "s"}`;
  const modules = answer.modules.length === 1
    ? answer.modules[0]
    : `${answer.modules.length} modules: ${answer.modules.join(", ")}`;
  return `Renamed to ${newName} — ${uses} in ${modules}.`;
}

/** A range of nothing, which is what a refused rename has to carry alongside its reason. */
function emptyRangeAt(position: monaco.Position): monaco.Range {
  return new monaco.Range(position.lineNumber, position.column, position.lineNumber, position.column);
}

function modifierBits(modifiers: readonly string[] | null | undefined): number {
  let bits = 0;
  for (const modifier of modifiers ?? []) {
    const at = SEMANTIC_TOKEN_MODIFIERS.indexOf(modifier);
    if (at >= 0) {
      bits |= 1 << at;
    }
  }
  return bits;
}

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
    // Asked for outright. The default defers to the theme, and a standalone theme has no way to
    // say yes: the flag is hardcoded off on every one of them, so a provider would be registered,
    // asked nothing, and paint nothing.
    "semanticHighlighting.enabled": true,
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

  // Quick fixes. Every fix answers a finding, so the squiggles already on screen say whether
  // there can be any: no marker touching the range means no round trip, which matters because
  // the lightbulb asks again every time the caret settles.
  monaco.languages.registerCodeActionProvider(VBA_LANGUAGE_ID, {
    provideCodeActions: async (model, range) => {
      const none = { actions: [], dispose: () => { } };
      if (model !== bridge.hostActiveModel()) {
        return none;
      }

      const markers = monaco.editor
        .getModelMarkers({ resource: model.uri, owner: MARKER_OWNER })
        .filter((marker) => monaco.Range.areIntersectingOrTouching(marker, range));
      if (markers.length === 0) {
        return none;
      }

      const offered = await bridge.requestCodeActions(
        model.getOffsetAt(range.getStartPosition()),
        model.getOffsetAt(range.getEndPosition()));

      const actions: monaco.languages.CodeAction[] = offered.map((action) => {
        const start = model.getPositionAt(action.start);
        const end = model.getPositionAt(action.end);
        // The finding this fix answers, so the menu groups it under the right squiggle and the
        // marker's own "fix this" affordance finds it.
        const answered = markers.filter((marker) =>
          marker.code === action.code
          && marker.startLineNumber === start.lineNumber
          && marker.startColumn === start.column
          && marker.endLineNumber === end.lineNumber
          && marker.endColumn === end.column);

        return {
          title: action.title,
          kind: "quickfix",
          isPreferred: action.isPreferred,
          ...(answered.length > 0 ? { diagnostics: answered } : {}),
          edit: {
            edits: action.edits.map((edit) => ({
              resource: model.uri,
              versionId: model.getVersionId(),
              textEdit: {
                range: monaco.Range.fromPositions(
                  model.getPositionAt(edit.start),
                  model.getPositionAt(edit.end)),
                text: edit.text,
              },
            })),
          },
        };
      });

      return { actions, dispose: () => { } };
    },
  }, {
    // Declared, not decorative: the editor gates Ctrl+. and Shift+Alt+. on a context key built
    // from exactly this list, so a provider that omits it draws a lightbulb nobody can open from
    // the keyboard.
    providedCodeActionKinds: ["quickfix"],
  });

  // Go to definition, across the modules of one workbook and never past it.
  //
  // A module the developer already has open answers as a location, so Ctrl+click, F12 and the
  // peek window all work in place. A module that is not open has no model for the editor to
  // point at, so the host is asked to go there instead — which is the same path the search
  // results and the outline tree already take, and it opens the module on the way.
  monaco.languages.registerDefinitionProvider(VBA_LANGUAGE_ID, {
    provideDefinition: async (model, position) => {
      if (model !== bridge.hostActiveModel()) {
        return null;
      }

      const found = await bridge.requestNavigation(model.getOffsetAt(position), false);
      const locations = toEditorLocations(bridge, found);
      const elsewhere = found[0];

      if (locations.length === 0 && elsewhere) {
        bridge.navigateTo(elsewhere);
      }

      return locations;
    },
  });

  // Find all references, over the same answers.
  //
  // Only modules that are open can be shown: the peek window renders text, and a module with no
  // tab has no model to render. The answer itself covers the whole workbook — what is missing is
  // a way for the page to ask the host to open a module without also moving the caret into it.
  monaco.languages.registerReferenceProvider(VBA_LANGUAGE_ID, {
    provideReferences: async (model, position, context) => {
      if (model !== bridge.hostActiveModel()) {
        return null;
      }

      return toEditorLocations(
        bridge,
        await bridge.requestNavigation(
          model.getOffsetAt(position), true, context.includeDeclaration));
    },
  });

  // Rename, across every module of the workbook that uses the symbol, whether its tab is open or
  // not (the developer, 2026-08-06).
  //
  // The HOST does the renaming, so this returns no edits. A module with no tab has no model to
  // edit, and those are exactly the ones a rename must not miss — so the work goes where the
  // modules are, and the open tabs are refreshed by the ordinary document sync that follows.
  monaco.languages.registerRenameProvider(VBA_LANGUAGE_ID, {
    resolveRenameLocation: (model, position): monaco.languages.RenameLocation & monaco.languages.Rejection => {
      if (model !== bridge.hostActiveModel()) {
        return { range: emptyRangeAt(position), text: "", rejectReason: "Rename works in the module the editor is showing." };
      }

      const word = model.getWordAtPosition(position);
      if (!word) {
        return { range: emptyRangeAt(position), text: "", rejectReason: "There is no symbol here to rename." };
      }

      return {
        range: new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn),
        text: word.word,
      };
    },
    provideRenameEdits: async (model, position, newName) => {
      if (model !== bridge.hostActiveModel()) {
        return { edits: [], rejectReason: "Rename works in the module the editor is showing." };
      }

      const answer = await bridge.requestRename(model.getOffsetAt(position), newName);
      if (answer.refused) {
        return { edits: [], rejectReason: answer.refused };
      }

      bridge.shell?.notify(renameSummary(answer, newName));
      return { edits: [] };
    },
  });

  // Semantic colouring, over the grammar rather than instead of it. The grammar already paints
  // the project's words from the lists project/open hands it; what it cannot do is tell a class
  // from an enum from a user-defined type, or tell a host global from a local that shadows its
  // name. Those need the analysis, and this is where it arrives.
  monaco.languages.registerDocumentSemanticTokensProvider(VBA_LANGUAGE_ID, {
    getLegend: () => ({ tokenTypes: SEMANTIC_TOKEN_TYPES, tokenModifiers: SEMANTIC_TOKEN_MODIFIERS }),
    provideDocumentSemanticTokens: async (model) => {
      const tokens = await bridge.requestSemanticTokens(model);
      if (!tokens) {
        // Null keeps what is already painted. Returning an empty set would strip the colouring
        // from a module whose analysis merely took too long.
        return null;
      }

      // The editor's wire format: five numbers per token, each row relative to the row before.
      const data = new Uint32Array(tokens.length * 5);
      let previousLine = 1;
      let previousColumn = 1;
      let at = 0;

      for (const token of tokens) {
        const start = model.getPositionAt(token.start);
        const end = model.getPositionAt(token.end);
        if (end.lineNumber !== start.lineNumber) {
          // A token the editor cannot express: its rows are single-line by construction.
          continue;
        }

        const typeIndex = SEMANTIC_TOKEN_TYPES.indexOf(token.type);
        if (typeIndex < 0) {
          continue;
        }

        data[at++] = start.lineNumber - previousLine;
        data[at++] = start.lineNumber === previousLine
          ? start.column - previousColumn
          : start.column - 1;
        data[at++] = end.column - start.column;
        data[at++] = typeIndex;
        data[at++] = modifierBits(token.modifiers);

        previousLine = start.lineNumber;
        previousColumn = start.column;
      }

      return { data: data.subarray(0, at) };
    },
    releaseDocumentSemanticTokens: () => { },
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
 * The commands in each editor's context menu and command palette.
 *
 * Only what the VBA host alone can do is sent to it: running a procedure, and the breakpoints
 * the debugger owns. Everything about the code itself is xlide's, because xlide knows more
 * about it — the editor's Go to Definition crosses modules, understands members reached through
 * a receiver, and reads the text as typed rather than as last written back, none of which the
 * host's own does (the developer, 2026-08-06: everything should be xlide).
 *
 * The VBA keys are kept on the xlide commands. Shift+F2 is what a VBA developer's hands already
 * do; what changes is what answers.
 */
function registerHostActions(editor: monaco.editor.IStandaloneCodeEditor, bridge: EditorBridge): void {
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

  // The VBA keys, answered by the editor. Both editor commands are registered by their features
  // rather than as editor actions, so they are triggered by id — which falls through to the
  // command registry, where those live.
  const editorActions: Array<[string, string, string, number]> = [
    ["xlide.goToDefinition", "Go to Definition (Shift+F2)", "editor.action.revealDefinition",
      monaco.KeyMod.Shift | monaco.KeyCode.F2],
    ["xlide.findReferences", "Find All References (Shift+F12)", "editor.action.goToReferences",
      monaco.KeyMod.Shift | monaco.KeyCode.F12],
    // The VBE's Last Position steps back through where the caret has been. So does this, and it
    // steps back through every move rather than only the ones that were jumps.
    ["xlide.lastPosition", "Last Position (Ctrl+Shift+F2)", "cursorUndo",
      monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.F2],
  ];

  for (const [id, label, command, key] of editorActions) {
    editor.addAction({
      id,
      label,
      contextMenuGroupId: "1_xlide",
      contextMenuOrder: 2,
      keybindings: [key],
      run: (target) => { target.trigger("xlide", command, null); },
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
