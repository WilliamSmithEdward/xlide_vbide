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
// Ctrl+click on a symbol. The commands the same feature is named for - F12, Shift+F12, the peek
// windows, and their right-click entries - live in a module its register never imports, so they
// come in by their own path. That is the third feature here whose register module covers less
// than its name does; the pattern is worth expecting rather than rediscovering.
import "monaco-editor/features/gotoSymbol/register.js";
import "monaco-editor/editor/contrib/gotoSymbol/browser/goToCommands.js";
// And the window those commands open. goToCommands asks ReferencesController for the peek, and
// the STANDALONE editor registers that controller in a module of its own that nothing else
// imports - so Find All References ran, resolved, and showed nothing. Fourth feature here whose
// pieces arrive separately.
import "monaco-editor/editor/standalone/browser/referenceSearch/standaloneReferenceSearch.js";
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

// Not a contribution but the registry the right-click menu is built from, and the expression the
// menu asks before drawing an entry. Reached directly because one entry the editor registers
// unconditionally has to be moved rather than merely not registered - see foldPeekIntoTheMenu.
import { MenuId, MenuRegistry } from "monaco-editor/platform/actions/common/actions.js";
import { ContextKeyExpr } from "monaco-editor/platform/contextkey/common/contextkey.js";

import "./styles.css";
import { EditorBridge, MARKER_OWNER, demoTransport, webView2Transport, type HostCompletionItem, type HostLocation, type HostRenameAnswer } from "./bridge.js";
import { showContextMenu } from "./contextmenu.js";
import { installDevSurface } from "./devsurface.js";
import { openReferencesDialog } from "./referencesdialog.js";
import { DocumentStore, docKeyOf, docUriOf, type DocumentId } from "./documents.js";
import { DesignerView } from "./designerview.js";
import { SearchWidget } from "./searchwidget.js";
import { registerFormatting } from "./format.js";
import { currentSettings, onSettingsApplied } from "./settings.js";
import { openPanesMenu, openSettingsDialog } from "./settingsdialog.js";
import { ChangesPane } from "./changespane.js";
import { openSyncDialog } from "./syncdialog.js";
import { openHelpDialog } from "./helpdialog.js";
import { openSponsorDialog } from "./sponsordialog.js";
import { Bookmarks } from "./bookmarks.js";
import { bootObjectBrowserPage } from "./objectbrowser.js";
import { Shell } from "./shell.js";
import { registerTestDirectiveCompletion } from "./testdirectives.js";
import { TestsPane } from "./testspane.js";
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
// `function` is appended rather than sorted in: an index into this list is the wire format,
// so appending is what keeps every already-painted token meaning what it meant.
// `enumMember` (2026-08-19): a resolved host constant - xlLandscape, wdOrientPortrait,
// acNormal - painted the tint the companion renders them in (xlide_vscode#35).
const SEMANTIC_TOKEN_TYPES = ["class", "enum", "struct", "type", "variable", "function", "enumMember"];

/** `defaultLibrary` marks a host-injected global - Application, ThisWorkbook, ActiveSheet. */
const SEMANTIC_TOKEN_MODIFIERS = ["defaultLibrary"];

/**
 * Where each URI the editor was handed came from, so the opener below can name a module to the
 * host the way the host spelled it. A URI has to lowercase the workbook to be a stable identity;
 * the host wants its own spelling back. Bounded by the workbook: one entry per module ever
 * offered as a target.
 */
const offeredTargets = new Map<string, HostLocation>();

/**
 * The host's answers as the editor wants them.
 *
 * A module with no tab open has no model, and its URI is built rather than the location dropped.
 * Dropping it is how the definition provider came to navigate from inside itself: with nothing
 * left to return there was no way to reach the module but to go there - and a provider that moves
 * the caret cancels the very request that asked for it, because the editor watches the position
 * while a provider runs. That is why Peek Definition jumped instead of peeking. The provider
 * answers the question; the opener is what acts on the answer.
 */
function toEditorLocations(
  bridge: EditorBridge,
  locations: readonly HostLocation[],
): monaco.languages.Location[] {
  return locations.map((location) => {
    // An open module's own model URI, not a rebuilt one: both come from the same parts, but the
    // model already exists and its spelling is the one the editor knows it by.
    const uri = bridge.modelForLocation(location)?.uri
      ?? docUriOf(location.module, location.workbook ?? null);
    offeredTargets.set(uri.toString(), location);

    return {
      uri,
      range: new monaco.Range(
        location.line,
        location.column,
        location.line,
        location.column + location.length),
    };
  });
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
  return `Renamed to ${newName}: ${uses} in ${modules}.`;
}

/**
 * Every use of the symbol at the caret, in xlide's own list. One function, so the key and the
 * right-click entry cannot show two different things.
 */
async function referencesAt(
  bridge: EditorBridge,
  model: monaco.editor.ITextModel,
  position: monaco.Position,
): Promise<{ word: string; found: HostLocation[] } | null> {
  // The same refusal every language provider opens with: this answers for the module the host
  // believes is active and for no other, because that is what the menu entry does.
  if (model !== bridge.hostActiveModel()) {
    return null;
  }

  return {
    word: model.getWordAtPosition(position)?.word ?? "",
    found: await bridge.requestNavigation(model.getOffsetAt(position), true, true),
  };
}

async function showReferences(
  bridge: EditorBridge,
  editor: monaco.editor.ICodeEditor,
): Promise<void> {
  const model = editor.getModel();
  const position = editor.getPosition();
  if (!model || !position) {
    return;
  }

  await showReferencesAt(bridge, model, position);
}

/**
 * Opens the references dialog for one position, which is the FEATURE - the dialog renders a
 * module with no tab open, which monaco's own peek cannot. The menu entry (Shift+F12) and the
 * `references` act's open form both come here, through the same referencesAt lookup, so the two
 * cannot list a different set. Answers whether a dialog was opened.
 */
async function showReferencesAt(
  bridge: EditorBridge,
  model: monaco.editor.ITextModel,
  position: monaco.Position,
): Promise<boolean> {
  const answer = await referencesAt(bridge, model, position);
  if (!answer) {
    return false;
  }

  openReferencesDialog(answer.word || "this symbol", answer.found, {
    navigate: (module, line, column, workbook) =>
      bridge.navigate(module, line, column, false, workbook ?? undefined),
  });

  return true;
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
      /*
       * ALWAYS ON, and no longer a setting.
       *
       * Two paths canonicalise keywords before the formatter is ever asked and neither consults a
       * setting: the HOST respells them as it takes a module, and the page runs its own recase per
       * touched line. Typing `public sub go()` with the switch OFF still produced `Public Sub go()`
       * (2026-08-09). So the switch promised something it could not deliver, on a row in a dialog
       * that was already too tall.
       *
       * The respell stays because it still covers a real window: paste, then Format Module inside
       * the 200ms before the recase timer fires.
       */
      canonicalKeywords: true,
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
    // Room above the first line and below the last. Line one sat hard against the tab strip, so
    // the top of a module read as clipped rather than as the start of something (the developer,
    // 2026-08-07). Below matters too: the last line otherwise sits on the status bar, which
    // scrollBeyondLastLine: false is what leaves it doing.
    padding: {
      top: 8,
      bottom: 8,
    },
    find: {
      addExtraSpaceOnTop: false,
    },
    renderLineHighlight: "line",
    renderWhitespace: "selection",
    fontFamily: '"Cascadia Mono", Consolas, "Courier New", monospace',
    fontSize: 13,
    // The width is the DEVELOPER'S choice, not a constant; the character is not a choice at all.
    //
    // Spaces, always: VBA's code store will not hold a tab, and expands any it is handed, so a
    // module indented with tabs read back as spaces and the page disagreed with the workbook for
    // as long as it stayed open. The "indent with tabs" setting was removed the day that was
    // measured (2026-08-07).
    //
    // `useTabStops` is what makes that bearable, and is load-bearing rather than incidental:
    // Backspace in a line's leading whitespace takes back a whole indent level rather than one
    // space, so spaces behave like tabs where a developer actually feels the difference.
    // detectIndentation stays off, or monaco would guess from the file and overrule the width.
    tabSize: currentSettings().formatIndentSize,
    insertSpaces: true,
    useTabStops: true,
    detectIndentation: false,
    autoIndent: "full",
    // A BLANK LINE KEEPS THE INDENT IT WAS GIVEN, which is what the editor underneath does.
    //
    // On by default, this removes indentation the editor inserted as soon as the caret leaves
    // the line without typing anything - so pressing Enter twice left the first of the two
    // lines at column 1, and arrowing back up landed there instead of at the indent the code
    // around it sits on (the owner, 2026-08-21). The native VBE keeps that whitespace and so
    // does the module it writes, so trimming it was this editor disagreeing with the thing it
    // covers.
    trimAutoWhitespace: false,
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

  // Indentation follows the setting for the WHOLE session, not only for editors made after the
  // change. A developer who switches to spaces mid-session and finds the open module still
  // inserting tabs has been given a switch that half works.
  onSettingsApplied((next) => {
    for (const editor of workspace.editors()) {
      editor.updateOptions({ tabSize: next.formatIndentSize });
    }
  });

  // Bridge, workspace, and shell reference each other, so they are built in dependency
  // order and stitched by assignment: the bridge first (it only needs the transport), the
  // workspace next (its editor factory wires each new editor into everything), the search
  // widget on the workspace's active editor, and the shell last - its toolbar keeps only
  // the commands that resolve as actions at build time, so every per-editor action must be
  // registered before it looks. Nothing host-driven runs before bridge.start().
  const bridge = new EditorBridge(transport ?? demoTransport(), documents);

  let searchWidget: SearchWidget;
  let workspace: Workspace;
  // Declared before the workspace: its constructor announces the first active group, and the
  // callback below must find an undefined shell, not a const still in its dead zone.
  let shell: Shell | undefined;
  let changesPane: ChangesPane | undefined;

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

  // The designer views, one per form, living as long as their tab does. The workspace mounts
  // a view's root into whichever group shows its tab; membership comes back from the host's
  // list through designerRetain, which is what disposes a view - a move between groups never
  // touches the host's list, so a move can never be mistaken for a close.
  const designerViews = new Map<string, DesignerView>();

  /**
   * Applies and saves every designer document holding unwritten edits, answering who was
   * saved and who refused. The two consumers are sync's export flush - the dialog's wrapper
   * here and the debug api's pre-step through the designerSaveDirty act - so both doors ship
   * what is on screen, never the last save (decision 15's sibling rule; the Run command keeps
   * the same one).
   */
  const saveDirtyDesigners = async (): Promise<{ saved: string[]; refused: string[] }> => {
    const saved: string[] = [];
    const refused: string[] = [];
    for (const view of designerViews.values()) {
      if (!view.isDirty()) {
        continue;
      }

      const outcome = await view.saveNow();
      if (outcome !== null && !outcome.ok) {
        refused.push(`${view.id.module}: ${outcome.refused ?? "the apply was refused"}`);
      } else {
        saved.push(view.id.module);
      }
    }

    return { saved, refused };
  };

  const designerViewFor = (id: DocumentId): DesignerView => {
    const key = docKeyOf(id.module, id.project, id.face);
    let view = designerViews.get(key);
    if (!view) {
      view = new DesignerView(id, {
        request: () => bridge.requestFormMarkup(id.module, id.project ?? null),
        watch: (listener) => bridge.onFormMarkup(id.module, id.project ?? null, listener),
        apply: (markup) => bridge.applyFormMarkup(id.module, id.project ?? null, markup),
        watchApplied: (listener) => bridge.onFormMarkupApplied(id.module, id.project ?? null, listener),
        dirtyChanged: (dirty) => workspace?.setFaceDirty(id, dirty),
        lint: (markup) => bridge.lintFormMarkup(id.module, id.project ?? null, markup),
        watchLint: (listener) => bridge.onFormMarkupLint(id.module, id.project ?? null, listener),
        // The language's own vocabulary, not this form's: the module travels only so the host
        // can describe the Form itself from a live one.
        requestVocabulary: () => bridge.requestFormMarkupVocabulary(id.module, id.project ?? null),
        watchVocabulary: (listener) => bridge.onFormMarkupVocabulary(listener),
        // The RAW File Save - "saveOnly" skips the host's designer branch, which is what
        // asked this view to apply in the first place; "save" here would loop forever. F5
        // comes back through the sibling that saves and then launches the form.
        saveWorkbook: (run) => bridge.runCommand({
          id: run ? "saveOnlyThenRun" : "saveOnly", target: "host", icon: "", label: "Save",
        }),
        watchApplySave: (listener) => bridge.onDesignerApplySave(id.module, id.project ?? null, listener),
        eventStub: (control) => {
          // The gesture ends IN the code half, focused on the Sub (the owner, 2026-08-19).
          // Marked BEFORE the request, and the reveal happens when the host's own caret
          // arrives with the stub's real line - a placeholder navigate raced that caret and
          // whoever landed last won (measured: 5ms apart, either order).
          bridge.expectStubReveal(id.module, id.project ?? null);
          bridge.designerEventStub(id.module, id.project ?? null, control);
        },
        selection: (control) => bridge.designerSelection(id.module, id.project ?? null, control),
        zorder: (control, front) => bridge.designerZOrder(id.module, id.project ?? null, control, front),
        setProperty: (control, property, value) =>
          bridge.designerSetProperty(id.module, id.project ?? null, control, property, value),
        // The whole settings object with one field replaced, which is the call every control
        // in the settings dialog makes: the grid's switch on the canvas is one more of them.
        changeSetting: (key, value) => bridge.updateSettings({ ...currentSettings(), [key]: value }),
        // Read at CALL time, not captured: `shell` is built after the views' factory is
        // declared, so binding it here would bind the undefined it holds now.
        handlers: () => bridge.requestOutline(id.module, id.project ?? undefined),
        // The handler is in this form's OWN code half, so the navigation names the same module.
        // It lands on the code face because a navigation takes the slot from a designer tab -
        // which is the fix that made clicking a Sub in the tree work at all.
        openHandler: (line) => bridge.navigate(id.module, line, 1, true, id.project ?? undefined),
        autoSize: (control) => bridge.requestDesignerAutoSize(id.module, id.project ?? null, control),
        notify: (text) => { shell?.notify(text); },
      });
      designerViews.set(key, view);
    }
    return view;
  };

  workspace = new Workspace(editorArea, emptyView, documents, {
    createEditor: (groupBody) => {
      const editor = monaco.editor.create(groupBody, editorOptions);
      wireEditor(editor);
      return editor;
    },
    activate: (id) => bridge.activateModule(id.module, id.project ?? undefined, id.face),
    /*
     * A DESIGNER TAB'S UNAPPLIED EDITS GET THE SAME QUESTION a module's unsaved text gets.
     *
     * Closing one used to be unconditional, and the host's own comment said the page asked its
     * own question - which it did not (found in the 2026-08-16 hunt). So a developer who moved
     * three controls and pressed Ctrl+W lost the three moves with nothing said. The question is
     * the page's because the state is: unapplied markup lives in the view, and nothing host-side
     * can know to hold the close.
     *
     * Save means what Ctrl+S means on this tab - apply the document to the form, then save the
     * workbook - which is `applyNow`, the same path the key takes. The close follows the apply
     * rather than racing it.
     */
    close: (id, action) => {
      const view = id.face === "design"
        ? designerViews.get(docKeyOf(id.module, id.project, id.face))
        : undefined;
      if (view && !action && view.canvasSnapshot().dirty) {
        shell?.confirmClose(id.module, id.project ?? null, (chosen) => {
          if (chosen === "save") {
            view.applyNow();
          }

          bridge.closeModule(id.module, id.project ?? undefined, "discard", id.face);
        });
        return;
      }

      bridge.closeModule(id.module, id.project ?? undefined, action, id.face);
    },
    activeChanged: (id, editor) => {
      // The search widget floats over the active group and searches its editor.
      searchWidget.attachTo(editor.getContainerDomNode());
      searchWidget.onActiveEditorChanged();
      shell?.setActiveModule(id?.module ?? null, id?.project ?? null);
    },
    layoutChanged: () => {
      workspace?.editors().forEach((editor) => editor.layout());
      designerViews.forEach((view) => view.layout());
    },
    designerBody: (id) => (id.face === "design" ? designerViewFor(id).root : null),
    designerShown: (body) => {
      for (const view of designerViews.values()) {
        if (view.root === body) {
          view.shown();
          return;
        }
      }
    },
    designerRetain: (openKeys) => {
      for (const [key, view] of [...designerViews]) {
        if (!openKeys.has(key)) {
          designerViews.delete(key);
          view.dispose();
        }
      }
    },
  });
  bridge.workspace = workspace;

  const createMs = performance.now();

  // Automatic layout rides ResizeObserver and tracks the window live; the settle here is
  // only the safety net for a final frame the observer missed. It waits for the resize to
  // pause - running it per event doubled every layout of a drag, which read as latency and
  // churn (2026-08-05) - and a measure that finds nothing changed costs nothing.
  //
  // The live-resize class is the minimap's peace: its canvas repaints a frame behind the
  // layout that moved it, so during a drag its blocks were alternately stale, clipped, and
  // redrawn - a flicker at the right edge. Faded out while events stream and back in at
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
    openDesigner: (name, workbook) => bridge.activateModule(name, workbook, "design"),
    navigate: (module, line, column, selectLine, workbook) =>
      bridge.navigate(module, line, column, selectLine, workbook),
    // The drop ends in the same state the row's own click ends in - the module's activate, or
    // the procedure's navigate with its line - with only the PLACEMENT added: the workspace
    // remembers the group and index the drop chose and the arriving tab lands there.
    dragFromTree: (payload, start, became) =>
      workspace.beginDocumentDrag(
        { module: payload.module, project: payload.workbook ?? null },
        payload.member ? `${payload.module} - ${payload.member}` : payload.module,
        start,
        {
          became,
          open: payload.line === undefined
            ? () => bridge.activateModule(payload.module, payload.workbook)
            : () => bridge.navigate(payload.module, payload.line!, 1, true, payload.workbook),
        }),
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
      if (command.id === "openSync") {
        openSyncDialog(
          async (args, body) => {
            // EXPORT FLUSHES THE DESIGNERS FIRST (the owner, 2026-08-19): the document is the
            // transaction log and the form only catches up on a save, so an export that
            // skipped the save shipped the LAST save - the same bug Run had, fixed the same
            // way. A document that refuses to apply stops the export and says why, exactly
            // as it stops a run; exporting the stale form instead would be shipping a lie.
            if (args.direction === "export") {
              const flushed = await saveDirtyDesigners();
              if (flushed.refused.length > 0) {
                return {
                  error: "the export did not run: a designer document refuses to apply - "
                    + flushed.refused.join("; "),
                } as never;
              }
            }
            return bridge.requestSync(args, body);
          },
          () => workspace.activeEditor().focus(),
          // WHICH PROJECTS ARE OPEN, so the dialog can be pointed at one explicitly and say so
          // in every request. The tree's own list, and the workbook of the document being looked
          // at as the one to start on - which is the answer to "which file am I working on" that
          // the developer would give.
          {
            names: (shell?.currentProjects() ?? []).map((one) => one.name),
            current: workspace.activeDocument()?.project ?? null,
          });
        return;
      }
      if (command.id === "openHelp") {
        openHelpDialog(() => workspace.activeEditor().focus());
        return;
      }


      // Beside About, because it answers the question About raises. The addresses are opened by
      // the HOST: this page is not allowed to navigate anywhere, and the host will only open the
      // three it holds.
      if (command.id === "openSponsor") {
        openSponsorDialog(
          { openExternal: (url) => bridge.openExternal(url) },
          () => workspace.activeEditor().focus());
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
      || command.id === "openSponsor" || command.id === "openSync"
      || workspace.activeEditor().getAction(command.id) !== null,
    evaluate: (text) => bridge.evaluate(text),
    panelChanged: (name, open) => bridge.panelChanged(name, open),
    menuRequest: (path) => bridge.requestMenu(path),
    menuExecute: (path) => bridge.executeMenu(path),
    menuClosed: () => workspace.activeEditor().focus(),
    /*
     * A ROW ON A FORM'S OWN TAB EDITS THE DOCUMENT; everything else still writes the component.
     *
     * The designer tab is a transaction - drag, resize, delete and the rest all edit the text and
     * wait for Ctrl+S - and the panel was the one surface writing past it, straight to the form
     * over COM. That left its edits off the undo stack entirely and left the form holding values
     * the document had given back, which is both of the owner's reports on 2026-08-17 ("ctrl+z is
     * not undoing it", "ctrl+z updates the markdown editor, but not the designer"). See task #68.
     *
     * DESIGNATED DEVIATION, and docs/userform-designer.md says so: the native Properties window
     * changes the form the moment you leave the row, and this does not while a designer tab is
     * open. It changes the document, and Ctrl+S is the apply - the same bargain every other
     * gesture on the tab already makes.
     *
     * A module's (Name), a worksheet, a form with no designer tab open: no document exists to
     * write, so those go where they always went. So does anything the document cannot spell.
     */
    editProperty: (component, name, value) => {
      const active = workspace.activeDocument();
      const view = active?.face === "design"
        ? designerViews.get(docKeyOf(active.module, active.project, active.face))
        : undefined;
      if (view && view.spells(component, name) && view.writeProperty(component, name, value)) {
        return;
      }

      bridge.editProperty(component, name, value);
    },
    pickPicture: (component, name) => bridge.pickPicture(component, name),
    selectComponent: (name) => bridge.selectComponent(name),
    renameModule: (name, workbook, newName) => {
      void bridge.requestModuleRename(name, workbook, newName).then((answer) => {
        bridge.shell?.notify(answer.refused
          ?? `Renamed ${name} to ${newName}: ${answer.replaced} mention${answer.replaced === 1 ? "" : "s"}`
            + ` in ${answer.modules.length} module${answer.modules.length === 1 ? "" : "s"}.`);
      });
    },
    closeModule: (name, workbook, action) => bridge.closeModule(name, workbook, action),
    insertComponent: (kind, project) => bridge.insertComponent(kind, project),
    removeComponent: (name, project) => bridge.removeComponent(name, project),
    requestOutline: (module, workbook) => bridge.requestOutline(module, workbook),
    trace: (text) => bridge.trace(text),
    testsShown: () => bridge.testsAction("show"),
    changesShown: () => changesPane?.shown(),
  });
  bridge.shell = shell;

  // The Tests pane: a projection of the host's setTests message, its gestures posted back
  // through the same channel the debug api's tests route drives - one brain, two doors.
  const testsPane = new TestsPane(document.querySelector("#tests") as HTMLElement, {
    act: (action, test, file) => bridge.testsAction(action, test, file),
    // The file goes with the module: two open files can each hold an InvoiceTests, and a
    // navigation that named only the module would open whichever answered first.
    navigate: (module, line, file) => bridge.navigate(module, line, 1, true, file),
  });
  bridge.testsChanged = (tests) => testsPane.paint(tests);

  // The Changes pane: the change log, read when the pane is looked at and at no other time.
  // Its request goes through the same channel the debug api's `changes` route answers, so the
  // pane and a driver are reading one reply rather than two shapes that can drift.
  changesPane = new ChangesPane(
    document.querySelector("#changes") as HTMLElement,
    (args) => bridge.requestChanges(args));

  // Both panes' Current Module scope follows the same tab, and the shell is where every road
  // to a changed active module already meets - the host's own setActive and the workspace's
  // tab activation both land there.
  shell.activeModuleChanged = (module, project) => testsPane.setActiveModule(module, project);

  // Asked once at boot, so a pane restored open by the saved layout holds its state - the
  // support chip included - without waiting for a run, a press, or the first analysis pass
  // to land (the owner: "doesn't show until test is run"). A boot before any project is open
  // answers nothing, and auto-rediscovery corrects the picture at the first pass.
  bridge.testsAction("refresh");

  // Ctrl+W closes the active group's active tab from anywhere in the surface. The host's key
  /*
   * NO SURFACE FALLS THROUGH TO THE BROWSER'S OWN CONTEXT MENU.
   *
   * This page is an EDITOR, and the menu the host would otherwise show offers Back, Refresh,
   * Save as, Print, Send tab to your devices and Inspect. Refresh and Back throw the developer
   * out of what they were doing; the rest are meaningless here. It surfaced on the empty end of a
   * tab strip (the owner, 2026-08-18) and the same hole was open on every surface that had not
   * thought to write a handler - a panel's background, a status bar, the gap between two docks.
   *
   * ONE BACKSTOP RATHER THAN A HANDLER PER SURFACE. Every place with something to offer already
   * calls preventDefault and shows its own menu, and this runs after them on the way up; it does
   * not replace them, it catches everything they do not claim. A surface that grows a menu later
   * needs nothing from here.
   *
   * A TEXT FIELD IS NOT AN EXCEPTION, though it was for one build. The thought was that the
   * browser's menu at least carries Cut, Copy and Paste - and then the owner right-clicked a
   * Properties value and got Emoji, IMPORT PASSWORDS, Writing direction, Send tab to your devices
   * and Inspect above them. A password importer in a VBA editor is not a convenience. So the
   * field gets the four items it actually needs, ours, and nothing else.
   *
   * Paste is the only one that cannot be done with `execCommand`, which browsers block for
   * reading: it goes through `navigator.clipboard.readText`, measured working in this host before
   * being offered, and says so rather than failing silently if it is ever refused. The text is
   * put in through `setRangeText` and an `input` event, because the row's own listener is what
   * commits a property edit - inserting the characters without it would show a value the model
   * never heard about.
   *
   * Monaco is a real exception and needs nothing: it brings its own menu and claims the event
   * before this ever runs.
   */
  document.addEventListener("contextmenu", (event) => {
    event.preventDefault();

    const field = (event.target as HTMLElement | null)
      ?.closest<HTMLInputElement | HTMLTextAreaElement>("input, textarea");
    if (!field || field.disabled || field.readOnly) {
      return;
    }

    const has = field.selectionStart !== field.selectionEnd;
    const run = (command: string): void => {
      field.focus();
      document.execCommand(command);
    };

    showContextMenu(event.clientX, event.clientY, [
      { label: "Undo", run: () => { run("undo"); } },
      {},
      { label: "Cut", enabled: has, run: () => { run("cut"); } },
      { label: "Copy", enabled: has, run: () => { run("copy"); } },
      {
        label: "Paste",
        run: () => {
          field.focus();
          navigator.clipboard.readText().then((text) => {
            const from = field.selectionStart ?? field.value.length;
            const to = field.selectionEnd ?? from;
            field.setRangeText(text, from, to, "end");
            field.dispatchEvent(new Event("input", { bubbles: true }));
          }, () => {
            shell?.notify("This host will not let the page read the clipboard; Ctrl+V still will.");
          });
        },
      },
      {},
      { label: "Select All", run: () => { run("selectAll"); } },
    ]);
  });

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
  // provider answers only for its model. A background group's model gets no engine answers -
  // honest, where an answer computed against the wrong module's text would not be.
  const completionProvider: monaco.languages.CompletionItemProvider = {
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
  };

  monaco.languages.registerCompletionItemProvider(VBA_LANGUAGE_ID, completionProvider);
  registerTestDirectiveCompletion(VBA_LANGUAGE_ID);

  // Hovers come from the same engine: the identifier under the cursor described by its
  // declaration line, its origin, and its documentation. The signature renders as VBA code, the
  // way the extension renders it.
  const hoverProvider: monaco.languages.HoverProvider = {
    provideHover: async (model, position) => {
      if (model !== bridge.hostActiveModel()) {
        return null;
      }

      // What the debugger says this is worth, when execution is stopped in this module.
      //
      // First in the hover and answered even when the analyzer has nothing: while stopped, the
      // value is the thing being looked for, and a local has no declaration the analyzer would
      // describe anyway. This is the reading the Locals panel is already showing, asked about
      // one name instead of all of them.
      const word = model.getWordAtPosition(position);
      const live = word ? bridge.localValue(model, word.word) : null;

      const hover = await bridge.requestHover(model.getOffsetAt(position));

      if (!hover) {
        return live && word
          ? {
            range: new monaco.Range(position.lineNumber, word.startColumn, position.lineNumber, word.endColumn),
            contents: [{ value: "```vba\n" + word.word + " = " + live.value + "\n```" }],
          }
          : null;
      }

      const start = model.getPositionAt(hover.start);
      const end = model.getPositionAt(hover.end);
      const contents: monaco.IMarkdownString[] = [];

      if (live && word) {
        contents.push({ value: "```vba\n" + word.word + " = " + live.value + "\n```" });
      }

      contents.push({ value: "```vba\n" + hover.signature + "\n```" });

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
  };

  monaco.languages.registerHoverProvider(VBA_LANGUAGE_ID, hoverProvider);

  // Quick fixes. Every fix answers a finding, so the squiggles already on screen say whether
  // there can be any: no marker touching the range means no round trip, which matters because
  // the lightbulb asks again every time the caret settles.
  const codeActionProvider: monaco.languages.CodeActionProvider = {
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
  };

  monaco.languages.registerCodeActionProvider(VBA_LANGUAGE_ID, codeActionProvider, {
    // Declared, not decorative: the editor gates Ctrl+. and Shift+Alt+. on a context key built
    // from exactly this list, so a provider that omits it draws a lightbulb nobody can open from
    // the keyboard.
    providedCodeActionKinds: ["quickfix"],
  });

  // Go to definition, across the modules of one workbook and never past it.
  //
  // Nothing but an answer comes out of here. Every command that asks - F12, Ctrl+click, Shift+F2,
  // and Peek Definition - comes through this one provider, and it cannot tell which is asking, so
  // acting on the answer here means acting the same way for all of them. Peek asked and was
  // taken to the definition instead (the developer, 2026-08-07). Going anywhere is the opener's
  // job, below.
  const definitionProvider: monaco.languages.DefinitionProvider = {
    provideDefinition: async (model, position) => {
      if (model !== bridge.hostActiveModel()) {
        return null;
      }

      const found = await bridge.requestNavigation(model.getOffsetAt(position), false);

      // The text is fetched before the answer is given, for any module the page does not already
      // hold. Peek draws each result by resolving it to a MODEL, and the page holds a module's
      // text once it has been ACTIVATED - not because its pane is open - so peeking into a module
      // nobody had opened drew an empty window (2026-08-07). Asked for without activating
      // anything, which is what makes a peek a peek.
      await Promise.all(found
        .filter((location) => !bridge.modelForLocation(location))
        .map((location) => bridge.ensureDocument(location.module, location.workbook ?? null)));

      return toEditorLocations(bridge, found);
    },
  };

  monaco.languages.registerDefinitionProvider(VBA_LANGUAGE_ID, definitionProvider);

  // And where a location in another module is opened.
  //
  // The standalone editor's own open handler answers only for the model already attached to the
  // editor that asked; for every other document it returns null and the jump silently does not
  // happen. This is the seam the editor offers for exactly that, and the host fills it: the host
  // owns the modules, including the ones with no tab, so it opens the module, brings its group
  // forward and places the caret - the same path the tree and the references list already take.
  //
  // The answer MUST be true. Anything else is read as "not handled" and falls through to the
  // built-in handler, which then fails, so a navigation that worked reads as one that did nothing.
  monaco.editor.registerEditorOpener({
    openCodeEditor: (source, resource, selectionOrPosition) => {
      // The asking editor's own model stays the editor's business: moving within one module is a
      // move, and routing it through the host would make it a round trip that ends where it began.
      if (source.getModel()?.uri.toString() === resource.toString()) {
        return false;
      }

      const target = offeredTargets.get(resource.toString());
      if (!target) {
        return false;
      }

      const at = selectionOrPosition === undefined
        ? { lineNumber: target.line, column: target.column }
        : "lineNumber" in selectionOrPosition
          ? { lineNumber: selectionOrPosition.lineNumber, column: selectionOrPosition.column }
          : {
            lineNumber: selectionOrPosition.startLineNumber,
            column: selectionOrPosition.startColumn,
          };

      bridge.navigate(target.module, at.lineNumber, at.column, false, target.workbook ?? undefined);
      return true;
    },
  });

  // No reference provider is registered, deliberately.
  //
  // Registering one buys the editor's own Go to References and Peek References, and both draw
  // their results by resolving each one to a MODEL - which this surface only has for modules with
  // a tab open. So the use in a module nobody has opened, the one worth being shown, is the one
  // they cannot draw: the editor's window showed 3 of the 4 uses xlide's list showed (the
  // developer, 2026-08-06). Find All References below answers the same question over the same
  // engine answers, and renders the line the host sends, so an unopened module lists like any
  // other.
  //
  // Leaving the provider unregistered is also what takes those two commands off the right-click
  // menu and off Shift+F12: both are gated on there being a reference provider at all, so the
  // menu has no entry to duplicate and the key is free for xlide's (the developer, 2026-08-07).

  // Rename, across every module of the workbook that uses the symbol, whether its tab is open or
  // not (the developer, 2026-08-06).
  //
  // The HOST does the renaming, so this returns no edits. A module with no tab has no model to
  // edit, and those are exactly the ones a rename must not miss - so the work goes where the
  // modules are, and the open tabs are refreshed by the ordinary document sync that follows.
  const renameProvider: monaco.languages.RenameProvider = {
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
  };

  monaco.languages.registerRenameProvider(VBA_LANGUAGE_ID, renameProvider);

  // Semantic colouring, over the grammar rather than instead of it. The grammar already paints
  // the project's words from the lists project/open hands it; what it cannot do is tell a class
  // from an enum from a user-defined type, tell a host global from a local that shadows its
  // name, or know that a form's designer declares controls whose method calls should paint as
  // calls (xlide_vscode#20). Those need the analysis, and this is where it arrives.
  //
  // AND WHERE IT REFRESHES. Monaco re-queries tokens on edits and on this event, nothing else -
  // so a query that raced the project seed painted a form's control methods plain and STAYED
  // that way in an untouched module (caught 2026-08-19: NameBox.SetFocus blue while its hover
  // said "TextBox method"; reloading painted it). Diagnostics arriving is the analysis saying
  // something moved, debounced so a pass burst costs one re-query of the visible models.
  const semanticRefresh = new monaco.Emitter<void>();
  let semanticRefreshTimer: ReturnType<typeof setTimeout> | undefined;
  bridge.semanticsMayHaveMoved = () => {
    clearTimeout(semanticRefreshTimer);
    semanticRefreshTimer = setTimeout(() => semanticRefresh.fire(), 250);
  };
  monaco.languages.registerDocumentSemanticTokensProvider(VBA_LANGUAGE_ID, {
    onDidChange: semanticRefresh.event,
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
  const signatureProvider: monaco.languages.SignatureHelpProvider = {
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
  };

  monaco.languages.registerSignatureHelpProvider(VBA_LANGUAGE_ID, signatureProvider);

  // And the surface itself, as data and as actions: the debug api's `ui` and `act` routes. See
  // devsurface.ts for why this is a module rather than a script written fresh per question.
  installDevSurface({
    workspace,
    explorer: shell.explorerTree(),
    bridge,
    designer: {
      viewFor: (module, project) => {
        for (const view of designerViews.values()) {
          if (view.id.module.toLowerCase() === module.toLowerCase()
            && (project === null || (view.id.project ?? "").toLowerCase() === project.toLowerCase())) {
            return view;
          }
        }
        return null;
      },
      saveDirty: saveDirtyDesigners,
    },
    search: searchWidget,
    bookmarks,
    providers: {
      hover: hoverProvider,
      completion: completionProvider,
      signature: signatureProvider,
      codeAction: codeActionProvider,
      definition: definitionProvider,
      rename: renameProvider,
    },
    referencesAt: async ({ line, column }) => {
      const editor = workspace.activeEditor();
      const model = editor.getModel();
      return model ? referencesAt(bridge, model, new monaco.Position(line, column)) : null;
    },
    openReferences: async ({ line, column }) => {
      const editor = workspace.activeEditor();
      const model = editor.getModel();
      return model ? showReferencesAt(bridge, model, new monaco.Position(line, column)) : false;
    },
    panes: shell.paneVisibility(),
    openSettings: () => bridge.openSettings?.(),
    statusNotice: () => shell.currentNotice(),
    statusPosition: () => shell.currentPosition(),
    statusModule: () => shell.currentModule(),
    pressToolbar: (id) => shell.pressToolbarCommand(id),
    toolbarCommands: () => shell.toolbarCommandsShown(),
    properties: () => shell.propertiesShown(),
    editProperty: (name, value) => shell.editPropertyByName(name, value),
    openSponsors: () => openSponsorDialog(
      { openExternal: (url) => bridge.openExternal(url) },
      () => workspace.activeEditor().focus()),
    // The same call every control in the settings dialog makes: the whole settings object with
    // one field replaced, posted to the host. Not the host's settings route, which is the path
    // that could never have had this defect.
    changeSetting: (key, value) => {
      const settings = currentSettings();
      if (!Object.hasOwn(settings, key)) {
        return false;
      }

      bridge.updateSettings({ ...settings, [key]: value });
      return true;
    },
  });

  // Once, not per editor: the menu registry is the editor's, and every editor draws from it.
  foldPeekIntoTheMenu();

  // Shift+F2 is what a VBA developer's hands already do, so it names the editor's OWN Go to
  // Definition rather than an xlide action that forwards to it. A forwarding action would appear
  // in the command palette as a second Go to Definition - two entries, two keys, one thing - which
  // is the shape this menu has been getting rid of all morning. A rule adds the key and nothing
  // else, and the palette keeps one entry.
  monaco.editor.addKeybindingRule({
    keybinding: monaco.KeyMod.Shift | monaco.KeyCode.F2,
    command: "editor.action.revealDefinition",
  });

  watchPreferredTheme((theme) => bridge.applyOsTheme(theme));

  if (!transport) {
    console.log("[xlide demo] window.chrome.webview absent, running the loopback demo");
  }

  // The bundle's own resource entry splits the two costs that scriptMs lumps together:
  // everything before responseEnd is fetching, everything after is compiling and running.
  // A transfer size of zero is the browser's cache answering - the number that says whether
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
 * Peek Definition, on the menu itself rather than behind a slide-out (the developer, 2026-08-07).
 *
 * The editor registers a "Peek" submenu unconditionally and hangs six commands off it. This
 * surface answers exactly one of them: peek declaration, type definition, implementations,
 * references and call hierarchy are each gated on a provider xlide does not register. So the
 * slide-out holds one item, and reaching it costs a second click and a hover on a menu with room
 * for it.
 *
 * The entry is MOVED, not copied - the same command object, so the label and the Alt+F12 the menu
 * draws beside it are the editor's own and cannot drift from what the key actually does.
 *
 * The submenu itself is left alone. An entry whose children all resolve away is not drawn at all
 * (`menuService`: `if (submenuActions.length > 0)`), so Peek disappears because it is empty rather
 * than because it was hidden - and if xlide ever answers implementations or type definitions, it
 * comes back holding exactly those. Hiding the submenu outright would have left them unreachable.
 *
 * This reaches into a registry entry the editor owns, so it names what it is looking for and does
 * nothing when it is not there: a shape that moves under a version bump brings the slide-out back,
 * and nothing worse.
 */
function foldPeekIntoTheMenu(): void {
  for (const entry of MenuRegistry.getMenuItems(MenuId.EditorContextPeek)) {
    if (entry.command?.id !== "editor.action.peekDefinition") {
      continue;
    }

    MenuRegistry.appendMenuItem(MenuId.EditorContext, {
      command: entry.command,
      // Its own precondition, carried over: inside a peek window there is nothing to peek from,
      // and the entry should be as absent there as it was in the slide-out.
      when: entry.when,
      group: "navigation",
      // Between Go to Definition at 1.1 and Find All References at 1.35.
      order: 1.2,
    });

    entry.when = ContextKeyExpr.false();
  }
}

/**
 * The commands in each editor's context menu and command palette.
 *
 * Only what the VBA host alone can do is sent to it: running a procedure, and the breakpoints
 * the debugger owns. Everything about the code itself is xlide's, because xlide knows more
 * about it - the editor's Go to Definition crosses modules, understands members reached through
 * a receiver, and reads the text as typed rather than as last written back, none of which the
 * host's own does (the developer, 2026-08-06: everything should be xlide).
 *
 * The VBA keys are kept on the xlide commands. Shift+F2 is what a VBA developer's hands already
 * do; what changes is what answers.
 *
 * No command carries its key in its NAME. The menu draws keys in the column at the right, from
 * the keybinding service, so a name that also names the key says it twice and reads as the key
 * being part of what the command is called (the developer, 2026-08-07: the hint on the right and
 * not in the name is best practice). Every command here therefore declares its key.
 */
function registerHostActions(editor: monaco.editor.IStandaloneCodeEditor, bridge: EditorBridge): void {
  // These three keys never reach the page. The browser's accelerator hook claims F5, F8 and F9
  // before the document is offered them, because the surface covers the pane the editor used to
  // receive them through (`AddInSession.OnSurfaceKey`). Declaring them here is therefore a
  // statement of what the key means, which is what the menu draws - and a fallback that does the
  // same thing if the hook ever stops claiming them.
  const hostActions: Array<[string, string, string, number]> = [
    ["xlide.run", "Run Sub/UserForm", "run", monaco.KeyCode.F5],
    ["xlide.toggleBreakpoint", "Toggle Breakpoint", "toggleBreakpoint", monaco.KeyCode.F9],
    ["xlide.runToCursor", "Run To Cursor", "runToCursor",
      monaco.KeyMod.CtrlCmd | monaco.KeyCode.F8],
  ];

  for (const [id, label, command, key] of hostActions) {
    editor.addAction({
      id,
      label,
      contextMenuGroupId: "1_xlide",
      contextMenuOrder: 1,
      keybindings: [key],
      run: () => bridge.runCommand({ id: command, target: "host", icon: "", label }),
    });
  }

  // The VBE's Last Position steps back through where the caret has been. So does this, and it
  // steps back through every move rather than only the ones that were jumps. It keeps a menu
  // entry because the editor's own menu has no equivalent to duplicate - and its own name,
  // because "Cursor Undo" is not what a VBA developer is looking for.
  editor.addAction({
    id: "xlide.lastPosition",
    label: "Last Position",
    contextMenuGroupId: "1_xlide",
    contextMenuOrder: 2,
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.F2],
    run: (target) => { target.trigger("xlide", "cursorUndo", null); },
  });

  // Go to Definition, invoked ON the definition, has nowhere to go - so the editor runs its
  // "alternative definition command", which by default goes to references. That command opens the
  // window this surface cannot use, and it is now gated off anyway, so it would only report having
  // found nothing.
  //
  // Set to nothing rather than to xlide's own list, which was tried and could not work: the editor
  // looks the id up among its SYMBOL NAVIGATION commands and runs it as one of those, so an id
  // that names anything else is not found and the fallback silently never fires (the developer,
  // 2026-08-06: it does nothing). Nothing is what it now honestly does; Find All References is a
  // key and a menu entry away.
  editor.updateOptions({
    gotoLocation: { alternativeDefinitionCommand: "" },
  });

  // Undo Rename, which the editor's own undo cannot be.
  //
  // A rename edits every module that uses the symbol, and the undo stack is PER MODEL: Ctrl+Z in
  // the module on screen reverses that module's share and leaves the rest renamed, which is a
  // half-renamed project - worse than no undo at all. So the reversal is the host's, over the
  // same modules the rename touched, and it is a command of its own rather than a key that
  // already means something narrower.
  editor.addAction({
    id: "xlide.undoRename",
    label: "Undo Rename",
    contextMenuGroupId: "1_modification",
    contextMenuOrder: 1.5,
    run: async () => {
      const answer = await bridge.requestRenameUndo();
      bridge.shell?.notify(answer.refused
        ?? `Rename put back: ${answer.modules.length} module`
          + `${answer.modules.length === 1 ? "" : "s"}.`);
    },
  });

  // Find All References, in xlide's own list.
  //
  // The editor's window renders each result by resolving it to a MODEL, and this surface only has
  // models for modules with a tab open - so the use in a module nobody has opened, which is the
  // one worth being shown, is the one it cannot draw. This list renders the line the host sends
  // instead, so an unopened module is listed like any other and clicking it opens it.
  editor.addAction({
    id: "xlide.findAllReferences",
    label: "Find All References",
    contextMenuGroupId: "navigation",
    contextMenuOrder: 1.35,
    keybindings: [monaco.KeyMod.Shift | monaco.KeyCode.F12],
    run: (target) => void showReferences(bridge, target),
  });
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
// opens in its own floating window. The palette wants none of the editor's machinery -
// no Monaco boot, no shell, no bridge - so it takes its own door before any of that starts.
const entry = new URLSearchParams(window.location.search).get("view") === "objbrowser"
  ? bootObjectBrowserPage
  : boot;

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", entry, { once: true });
} else {
  entry();
}
