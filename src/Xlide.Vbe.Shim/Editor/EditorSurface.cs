using System.Text.Json;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.WebView;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The modern editing surface, positioned over a code pane.
///
/// The native pane underneath is never removed. It remains the text of record, the compile target,
/// and what the debugger drives, so everything the host does keeps working. It is simply not what
/// the developer looks at.
///
/// Placement follows the pane rather than being set once. Panes are document children of the
/// editor: they move, resize, and restack whenever the user rearranges anything, and a surface that
/// sampled a rectangle at creation would drift away from the pane within seconds.
/// </summary>
internal sealed class EditorSurface : IDisposable
{
    private readonly nint _host;

    private OverlayWindow? _overlay;
    private WebView2Surface? _browser;

#if DEBUG
    /// <summary>The page itself, for the debug api's eval route. Debug builds only.</summary>
    internal WebView2Surface? Browser => _browser;

    /// <summary>
    /// The overlay the page is drawn in — NOT <see cref="Host"/>, which is the document area
    /// the overlay is a child of. The surface covers more than the document area (it draws
    /// the menu bar and the toolbar), so a screenshot crop that used the parent landed tens
    /// of pixels high. Debug builds only.
    /// </summary>
    internal nint SurfaceWindow => _overlay?.Handle ?? 0;
#endif

    /// <summary>
    /// One open module as this surface mirrors it: the workbook it belongs to (display name,
    /// the identity the page's tabs use), the text as last agreed, and whether the developer
    /// has typed something that has not been written back yet.
    /// </summary>
    private sealed class OpenDoc
    {
        public required string Module { get; init; }
        public string? Project { get; init; }
        public required string Text { get; set; }
        public bool Unwritten { get; set; }
    }

    /// <summary>
    /// Every module the surface holds live, keyed by (workbook, module). One model per key on
    /// the page, one mirror per key here — two workbooks' Module1 are two documents (decision
    /// 12), where a name-keyed table would have merged them.
    /// </summary>
    private readonly Dictionary<string, OpenDoc> _docs = new(StringComparer.Ordinal);

    /// <summary>The active document's key — the module the native active pane shows.</summary>
    private string? _activeKey;

    /// <summary>The identity two documents are the same by, spelled like the page's tab key.</summary>
    private static string DocKey(string module, string? project) =>
        $"{(project ?? string.Empty).ToLowerInvariant()}\0{module.ToLowerInvariant()}";

    private OpenDoc? ActiveDoc => _activeKey is { } key && _docs.TryGetValue(key, out var doc) ? doc : null;

    /// <summary>
    /// Messages waiting for the page to be ready, newest per kind, in the order the kinds were
    /// first sent.
    ///
    /// Everything the host has to say arrives before the page can hear it. The engine answers in
    /// tens of milliseconds and the page takes a couple of seconds, so the module's text, its
    /// squiggles, and the first findings are all produced while there is nothing listening.
    /// Dropping them left a surface that was correctly positioned and permanently empty, which
    /// reads as a rendering fault rather than as messages that arrived too early.
    ///
    /// Only the newest of each kind is kept, because each one replaces the last: there is no value
    /// in replaying three sets of findings to a page that has seen none of them. Order is kept
    /// because there is one dependency between kinds: loading a document resets the model, and
    /// squiggles set against the model being replaced go with it.
    /// </summary>
    private readonly Dictionary<string, string> _pending = [];
    private readonly List<string> _pendingOrder = [];

    /// <summary>
    /// How long after the last keystroke the module is written back.
    ///
    /// Long enough that typing a word is one write rather than five, short enough that tabbing away
    /// or reaching for the mouse has already saved. Anything that must see the text sooner flushes
    /// explicitly rather than waiting for this.
    /// </summary>
    private const uint WriteDelayMilliseconds = 400;

    /// <summary>
    /// How long the typing rests before the live analysis pass. Just inside the write delay, so
    /// findings refresh from the text as typed without waiting for the module write-back. A
    /// large module waits longer: its pass costs more, and running it between keystrokes is
    /// what a completion would otherwise queue behind. Leaving the line runs it at once.
    /// </summary>
    private const uint LiveAnalysisDelayMilliseconds = 350;
    private const uint LiveAnalysisDelayLargeMilliseconds = 1500;
    private const int LargeModuleCharacters = 200_000;

    /// <summary>True when the text changed since the last live pass, the line-leave trigger's gate.</summary>
    private bool _changedSinceLiveAnalysis;

    private bool _loaded;

    private EditorSurface(nint host) => _host = host;

    /// <summary>The window this surface is a child of, which is the editor's document area.</summary>
    public nint Host => _host;

    /// <summary>The active module — the one the native active pane shows — or null when nothing is.</summary>
    public string? Module => ActiveDoc?.Module;

    /// <summary>The active module's workbook, by the display name the page's tabs use.</summary>
    public string? Project => ActiveDoc?.Project;

    /// <summary>Raised when the surface reports the developer changed a module's text:
    /// (module, workbook display name or null, the whole text).</summary>
    public Action<string, string?, string>? TextChanged { get; set; }

    /// <summary>Raised when the developer picks a module, with the workbook when the picker
    /// knows it (the tree does; the tab strip does not yet).</summary>
    public Action<string, string?>? ModuleRequested { get; set; }

    /// <summary>Raised when the developer wants to be taken to a place, with the workbook when
    /// the asker knows it.</summary>
    public Action<string, int, int, string?>? NavigateRequested { get; set; }

    /// <summary>Raised when the developer chooses a command the editor owns.</summary>
    public Action<string>? CommandRequested { get; set; }

    /// <summary>Raised when the developer enters a line in the Immediate panel.</summary>
    public Action<string>? EvaluateRequested { get; set; }

    /// <summary>Raised when the developer asks for the last rename to be put back.</summary>
    public Action<int>? RenameUndoRequested { get; set; }

    /// <summary>Raised with an address the page wants opened outside itself.</summary>
    public Action<string>? ExternalOpenRequested { get; set; }

    /// <summary>Raised when the page wants a module's text without being taken to it.</summary>
    public Action<string, string?>? DocumentRequested { get; set; }

    /// <summary>Raised with the panel that is showing, and whether the panel is open.</summary>
    public Action<string, bool>? PanelChanged { get; set; }

    /// <summary>Raised when the developer opens a menu, with the path to it; empty is the bar.</summary>
    public Action<int[]>? MenuRequested { get; set; }

    /// <summary>Raised when the developer chooses a menu item, with the path to it.</summary>
    public Action<int[]>? MenuExecuteRequested { get; set; }

    /// <summary>Raised when the developer edits a property: component, property name, new value.</summary>
    public Action<string, string, string>? PropertyEditRequested { get; set; }

    /// <summary>Raised when the developer selects a component in the explorer without opening it.</summary>
    public Action<string>? ComponentSelected { get; set; }

    /// <summary>Raised when the developer closes a module's tab, with its workbook when the
    /// tab carries one. The third value is their answer for unsaved changes — "save" or
    /// "discard" — and null on the first ask, before any question has been put.</summary>
    public Action<string, string?, string?>? ModuleCloseRequested { get; set; }

    /// <summary>Raised when the developer changed a setting in the page's dialog.</summary>
    public Action<ProductSettings>? SettingsChangeRequested { get; set; }

    /// <summary>
    /// Raised when the developer asks for a new component: (kind, workbook). Kind is 1 module,
    /// 2 class, 3 form; the workbook names the project whose menu asked, or null for the active
    /// one.
    /// </summary>
    public Action<int, string?>? ComponentInsertRequested { get; set; }

    /// <summary>Raised when the editor wants completions: request identifier, caret offset.</summary>
    public Action<int, int>? CompletionRequested { get; set; }

    /// <summary>Raised when the page asks what the identifier at an offset is: (requestId, offset).</summary>
    public Action<int, int>? HoverRequested { get; set; }

    /// <summary>Raised when the page asks for the call tip at an offset: (requestId, offset).</summary>
    public Action<int, int>? SignatureHelpRequested { get; set; }

    /// <summary>Raised when the page asks what Enter should leave behind: (requestId, offset).</summary>
    public Action<int, int>? SmartEnterRequested { get; set; }

    /// <summary>
    /// Raised when the page asks for case corrections:
    /// (requestId, start, end, single, completeHeader).
    /// </summary>
    public Action<int, int, int, bool, bool>? CanonicalCaseRequested { get; set; }

    /// <summary>Raised when the page asks for the paired loop rename: (requestId, offset).</summary>
    public Action<int, int>? LoopSyncRequested { get; set; }

    /// <summary>Raised when the page asks what can be fixed over a span: (requestId, start, end).</summary>
    public Action<int, int, int>? CodeActionsRequested { get; set; }

    /// <summary>
    /// Raised when the page asks where a symbol is declared or used:
    /// (requestId, offset, wantsReferences, includeDeclaration).
    /// </summary>
    public Action<int, int, bool, bool>? NavigationRequested { get; set; }

    /// <summary>Raised when the page asks for a module's procedures: (requestId, moduleName,
    /// workbook or null).</summary>
    public Action<int, string, string?>? OutlineRequested { get; set; }

    /// <summary>Raised when the page asks for a module's colouring: (requestId, moduleName,
    /// workbook or null).</summary>
    public Action<int, string, string?>? SemanticTokensRequested { get; set; }

    /// <summary>Raised when the page asks to rename a symbol: (requestId, offset, newName).</summary>
    public Action<int, int, string>? RenameRequested { get; set; }

    /// <summary>Raised when the page asks to rename a MODULE and everything that names it:
    /// (requestId, moduleName, workbook or null, newName).</summary>
    public Action<int, string, string?, string>? ModuleRenameRequested { get; set; }

    /// <summary>Runs an action on the host thread, which owns the browser and the object model.</summary>
    public void RunOnHostThread(Action action) => _overlay?.RunOnHostThread(action);

    /// <summary>Answers one completion request. Never held: an answer outlives no keystroke.</summary>
    public void ShowCompletions(int requestId, SurfaceCompletionItem[] items)
    {
        ArgumentNullException.ThrowIfNull(items);

        if (!_loaded)
        {
            return;
        }

        // The request was already logged with its item count; this line proves the answer made it
        // back across the thread marshal and onto the wire to the page.
        Log.Info($"completion: delivering {requestId}");

        Post(JsonSerializer.Serialize(
            new CompletionResultMessage("completionResult", requestId, items),
            EditorMessageContext.Default.CompletionResultMessage));
    }

    /// <summary>Answers one hover request. Never held: a hover outlives no cursor move.</summary>
    public void ShowHover(int requestId, SurfaceHoverPayload? hover)
    {
        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new HoverResultMessage("hoverResult", requestId, hover),
            EditorMessageContext.Default.HoverResultMessage));
    }

    /// <summary>Answers one call-tip request. Never held: a tip outlives no keystroke.</summary>
    public void ShowSignatureHelp(int requestId, SurfaceSignatureInfo? signature)
    {
        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SignatureHelpResultMessage("signatureHelpResult", requestId, signature),
            EditorMessageContext.Default.SignatureHelpResultMessage));
    }

    /// <summary>Answers one Smart Enter request. Never held: an answer outlives no keystroke.</summary>
    public void ShowSmartEnter(int requestId, SurfaceTextEdit[] edits, int? caret)
    {
        ArgumentNullException.ThrowIfNull(edits);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SmartEnterResultMessage("smartEnterResult", requestId, edits, caret),
            EditorMessageContext.Default.SmartEnterResultMessage));
    }

    /// <summary>Answers one canonical-case request. Never held.</summary>
    public void ShowCanonicalCase(int requestId, SurfaceTextEdit[] edits)
    {
        ArgumentNullException.ThrowIfNull(edits);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new CanonicalCaseResultMessage("canonicalCaseResult", requestId, edits),
            EditorMessageContext.Default.CanonicalCaseResultMessage));
    }

    /// <summary>Answers one loop-sync request. Never held.</summary>
    public void ShowLoopSync(int requestId, SurfaceTextEdit[] edits)
    {
        ArgumentNullException.ThrowIfNull(edits);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new LoopSyncResultMessage("loopSyncResult", requestId, edits),
            EditorMessageContext.Default.LoopSyncResultMessage));
    }

    /// <summary>Answers one quick-fix request. Never held: the caret moves and the offer lapses.</summary>
    public void ShowCodeActions(int requestId, SurfaceCodeAction[] actions)
    {
        ArgumentNullException.ThrowIfNull(actions);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new CodeActionResultMessage("codeActionResult", requestId, actions),
            EditorMessageContext.Default.CodeActionResultMessage));
    }

    /// <summary>Answers one rename request: what changed, or why nothing did.</summary>
    public void ShowRenamed(int requestId, string? oldName, string? newName, string[] modules, int replaced, string? refused)
    {
        ArgumentNullException.ThrowIfNull(modules);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new RenameResultMessage("renameResult", requestId, oldName, newName, modules, replaced, refused),
            EditorMessageContext.Default.RenameResultMessage));
    }

    /// <summary>Answers one navigation request. Never held: the caret moves and the answer lapses.</summary>
    public void ShowLocations(int requestId, SurfaceLocation[] locations)
    {
        ArgumentNullException.ThrowIfNull(locations);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new NavigationResultMessage("navigationResult", requestId, locations),
            EditorMessageContext.Default.NavigationResultMessage));
    }

    /// <summary>Answers one colouring request. Failed keeps what the page already shows.</summary>
    public void ShowSemanticTokens(int requestId, SurfaceSemanticToken[] tokens, bool failed)
    {
        ArgumentNullException.ThrowIfNull(tokens);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SemanticTokensResultMessage("semanticTokensResult", requestId, tokens, failed),
            EditorMessageContext.Default.SemanticTokensResultMessage));
    }

    /// <summary>
    /// Tells the tokenizer the project's words. Held for the page when it has not loaded yet,
    /// and replaced rather than queued when it has: only the latest lists describe the project.
    /// </summary>
    public void SetLanguageFacts(string[] types, string[] procedures)
    {
        ArgumentNullException.ThrowIfNull(types);
        ArgumentNullException.ThrowIfNull(procedures);

        Send("setLanguageFacts", JsonSerializer.Serialize(
            new SetLanguageFactsMessage("setLanguageFacts", types, procedures),
            EditorMessageContext.Default.SetLanguageFactsMessage));
    }

    /// <summary>Answers one outline request. Never held. Failed means "no answer", not "empty".</summary>
    public void ShowOutline(int requestId, SurfaceOutlineProcedure[] procedures, bool failed = false)
    {
        ArgumentNullException.ThrowIfNull(procedures);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new OutlineResultMessage("outlineResult", requestId, procedures, failed),
            EditorMessageContext.Default.OutlineResultMessage));
    }

    /// <summary>Answers an import/export request with the service's own JSON, verbatim.</summary>
    public void ShowSyncResult(int requestId, string json)
    {
        ArgumentNullException.ThrowIfNull(json);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SyncResultMessage("syncResult", requestId, json),
            EditorMessageContext.Default.SyncResultMessage));
    }

    /// <summary>
    /// The dialog asking what an import or export would do, or asking for it to be done.
    /// </summary>
    public Action<int, IReadOnlyDictionary<string, string>, string>? SyncRequested { get; set; }

    /// <summary>Raised once, when the page has loaded and everything held for it has been sent.</summary>
    public Action? Ready { get; set; }

    /// <summary>
    /// True once the page is up. Consulted before covering native chrome the page replaces: a menu
    /// bar covered by a surface whose page never arrived is a menu bar the developer cannot reach.
    /// </summary>
    public bool IsReady => _loaded;

    /// <summary>
    /// Asked about each key the editor might own, before the document sees it. Return true to
    /// claim it.
    /// </summary>
    public Func<uint, bool>? KeyPressed { get; set; }

    /// <summary>
    /// Where the caret is in the surface, one-based.
    ///
    /// Kept here rather than pushed into the editor as it moves. The editor's own caret only has to
    /// agree at the moment something reads it, which is when a command runs, and moving it on every
    /// cursor movement would put an automation call on the path of every arrow key.
    /// </summary>
    public int CaretLine { get; private set; } = 1;

    /// <summary>Caret column in the surface, one-based.</summary>
    public int CaretColumn { get; private set; } = 1;

    /// <summary>
    /// Creates the surface over the editor frame. Returns null when the editing bundle is not
    /// present, which is an ordinary state: the add-in works without it and shows the native pane.
    /// </summary>
    public static EditorSurface? Create(nint host, PixelRect bounds)
    {
        var directory = ShimModule.Directory;
        if (directory is null)
        {
            return null;
        }

        var root = WebViewPaths.EditorContentRoot(directory);
        if (!File.Exists(WebViewPaths.EditorEntryDocument(directory)))
        {
            Log.Info($"editor surface: no bundle at {root}, the native pane stays visible");
            return null;
        }

        var surface = new EditorSurface(host);

        surface._overlay = OverlayWindow.Create(host, bounds);
        if (surface._overlay is null)
        {
            return null;
        }

        surface._overlay.Resized += size => surface._browser?.SetBounds(size);
        surface._overlay.Elapsed = surface.FlushEdits;
        surface._overlay.Polled = () => surface.Polled?.Invoke();
        surface._overlay.SettleDue = () => surface.PlacementSettled?.Invoke();

        // Asked for the editing document, then left alone. Creating a browser is asynchronous in two
        // stages, so mapping the content root and navigating from here would run before the browser
        // exists and quietly do nothing. The surface performs both once it is ready.
        surface._browser = WebView2Surface.Start(
            surface._overlay.Handle,
            surface._overlay.ClientBounds());

        if (surface._browser is null)
        {
            surface._overlay.Dispose();
            return null;
        }

#if DEBUG
        surface._browser.DebugName = "editor";
#endif

        surface._browser.MessageReceived = surface.OnMessage;
        surface._browser.AcceleratorPressed = key => surface.KeyPressed?.Invoke(key) ?? false;
        surface._overlay.LoaderTicked = () => surface.LoadingPulse?.Invoke();
        surface._overlay.AnalysisDue = () =>
        {
            surface._changedSinceLiveAnalysis = false;
            surface.LiveAnalysisDue?.Invoke();
        };

        Log.Info($"editor surface: created, serving from {root}");
        return surface;
    }

    /// <summary>
    /// Raised on the host thread a few times a second while the loader is showing. The session
    /// re-asserts placement here, because the loading phase has no pane and therefore no window
    /// events, while the editor is still arranging itself underneath the loader.
    /// </summary>
    public Action? LoadingPulse { get; set; }

    /// <summary>True once the loader has been showing implausibly long; placement consults it.</summary>
    public bool IsLoaderStalled => _overlay?.LoaderStalled ?? false;

    /// <summary>Raised on the host thread when the typing has rested and the live text should be analysed.</summary>
    public Action? LiveAnalysisDue { get; set; }

    /// <summary>
    /// Raised on the host thread with each change to the shown module's text: (module, whole
    /// text or null, edits or null). Whoever mirrors the text — the engine — listens here.
    /// </summary>
    public Action<string, string?, EngineTextEdit[]?>? LiveTextPushed { get; set; }

    /// <summary>
    /// Raised when a change event was plain typing confined to one line — no newline in it,
    /// every range on that line — with the 1-based line. This is what begins the hold that
    /// keeps fresh verdicts off the line still being typed.
    /// </summary>
    public Action<int>? LineTyped { get; set; }

    /// <summary>Raised on every caret update with the settled 1-based line, after the caret
    /// properties reflect it. This is what releases the hold.</summary>
    public Action<int>? CaretLineSettled { get; set; }

    /// <summary>
    /// Raised when an edit added or removed lines in a module: everything anchored in it below
    /// afterLine moves by delta. This is how line-anchored bookkeeping — breakpoints — follows
    /// the text.
    /// </summary>
    public Action<string, int, int>? LinesShifted { get; set; }

    /// <summary>The page asked to search: id, query, matchCase, wholeWord, scope.</summary>
    public Action<int, string, bool, bool, string>? SearchRequested { get; set; }

    /// <summary>The page asked to replace across a scope: the search's shape plus the replacement.</summary>
    public Action<int, string, bool, bool, string, string>? ReplaceAllRequested { get; set; }

    /// <summary>Answers a search, echoing its id; replaced counts a replace-all's edits.</summary>
    public void ShowSearchResults(int id, SurfaceSearchMatch[] matches, bool truncated, int replaced = 0)
    {
        ArgumentNullException.ThrowIfNull(matches);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SearchResultMessage("searchResult", id, matches, truncated, replaced),
            EditorMessageContext.Default.SearchResultMessage));
    }

    /// <summary>Moves the surface over a pane, or hides it when there is nothing to cover.</summary>
    public void Follow(PixelRect bounds, bool visible)
    {
#if DEBUG
        var startedAt = Environment.TickCount64;
#endif
        // Growing gives the browser its new size BEFORE the window reaches it.
        //
        // The overlay grows in one call and the browser catches up in its own time, so for a
        // frame the strip between the old edge and the new one is the overlay's own ground —
        // a flat band at the right edge on every tick of a drag (the developer, 2026-08-06).
        // The window's paint handler already covers that strip deliberately, because what lies
        // under it is the old native editor and THAT bleeding through is worse. Painting it
        // sooner is not the answer; not exposing it is.
        //
        // A child may be larger than its parent — it is simply clipped — so sizing it first
        // means the parent grows into a child that already covers the new area. Only when
        // growing: shrinking the child first would expose the same band on the inside edge,
        // which is the same defect facing the other way.
        if (visible && _overlay is { } growing)
        {
            var current = growing.ClientBounds();
            if (bounds.Width > current.Width || bounds.Height > current.Height)
            {
                _browser?.SetBounds(new PixelRect(0, 0, bounds.Width, bounds.Height));
            }
        }

        _overlay?.Place(bounds, visible);
#if DEBUG
        var placedAt = Environment.TickCount64;
#endif

        // The browser's size normally rides the WM_SIZE the placement above produces; it is
        // asserted here as well, so a size message that never arrived — a raced resize storm —
        // cannot leave the page laid out for a width the window no longer has, with its right
        // edge (minimap, scrollbar) falling outside the visible surface (2026-08-05). The
        // browser side skips the call when nothing changed, so the assertion is free.
        if (_overlay is { } overlay)
        {
            _browser?.SetBounds(overlay.ClientBounds());
        }

#if DEBUG
        var doneAt = Environment.TickCount64;
        Diagnostics.PerfCounters.Follow(placedAt - startedAt, doneAt - placedAt, _browser is not null);
#endif
    }

    /// <summary>
    /// Punches holes in the surface where a native window must show through, in the frame's
    /// client space — one tenant today: the Object Browser. An empty set makes it whole.
    /// </summary>
    public void SetCutouts(ReadOnlySpan<PixelRect> holes) => _overlay?.SetCutouts(holes);

    /// <summary>Raised once, on the host thread, a quiet moment after the last armed frame
    /// event — when whoever owns placement runs its full pass once instead of per event.</summary>
    public Action? PlacementSettled { get; set; }

    /// <summary>Starts or restarts the settle debounce; each frame event pushes it out.</summary>
    public void ArmPlacementSettle(uint milliseconds) => _overlay?.StartSettleTimer(milliseconds);

    /// <summary>
    /// Puts keyboard focus back on the surface. Activating a native pane takes it, and a surface
    /// that does not have focus hears no keys: the shortcut that switched module would work once
    /// and then fall silent.
    /// </summary>
    public void Focus() => _browser?.Focus();

    /// <summary>
    /// Makes a module the active document, opening it live on the page if it is not already.
    ///
    /// A document already open keeps what it has: unwritten edits outrank the read that produced
    /// this call (the page's model is ahead of the module, not behind it), and a clean document
    /// whose module changed underneath adopts the new text as a sync — an in-place edit that
    /// keeps the model's undo stack and caret — never as a reload.
    /// </summary>
    public void Show(string moduleName, string? project, string text)
    {
        var key = DocKey(moduleName, project);
        _activeKey = key;

        if (!_docs.TryGetValue(key, out var doc))
        {
            _docs[key] = new OpenDoc { Module = moduleName, Project = project, Text = text };
            PostOpenDocument(moduleName, project, text);
            return;
        }

        if (doc.Unwritten)
        {
            return;
        }

        if (!string.Equals(doc.Text, text, StringComparison.Ordinal))
        {
            doc.Text = text;
            PostSyncDocument(moduleName, project, text);
        }
    }

    /// <summary>
    /// Gives the page a module's text WITHOUT making it the active document.
    ///
    /// Show() is the only other way a document reaches the page, and it moves the active one,
    /// because it is what activating a pane calls. The page needs a third thing: text for a module
    /// it is going to draw but not go to — a definition it is peeking at, a reference it is
    /// listing. Peeking one and being taken there instead is exactly the bug this answers.
    ///
    /// A document the page already has is left alone. It may be ahead of the module (unwritten
    /// edits) and this is not a sync; it is a copy for something that had none.
    /// </summary>
    public void Publish(string moduleName, string? project, string text)
    {
        var key = DocKey(moduleName, project);
        if (_docs.ContainsKey(key))
        {
            return;
        }

        _docs[key] = new OpenDoc { Module = moduleName, Project = project, Text = text };
        PostOpenDocument(moduleName, project, text);
    }

    private void PostOpenDocument(string moduleName, string? project, string text)
    {
        if (!_loaded)
        {
            // Not queued: the ready handler re-opens every live document, which also covers a
            // page that reloaded mid-session and lost its models.
            return;
        }

        Post(JsonSerializer.Serialize(
            new OpenDocumentMessage("openDocument", moduleName, project, text),
            EditorMessageContext.Default.OpenDocumentMessage));
    }

    private void PostSyncDocument(string moduleName, string? project, string text)
    {
        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SyncDocumentMessage("syncDocument", moduleName, project, text),
            EditorMessageContext.Default.SyncDocumentMessage));
    }

    /// <summary>
    /// Shows nothing: every pane is closed, and the workspace says so instead of showing the last
    /// module as if it were still open. Unwritten edits are flushed first, not dropped.
    /// </summary>
    public void Clear()
    {
        FlushEdits();
        _docs.Clear();
        _activeKey = null;

        // Squiggles held for a page that never arrived describe models that no longer exist.
        foreach (var kind in _pendingOrder.Where(k => k.StartsWith("setDiagnostics:", StringComparison.Ordinal)).ToArray())
        {
            Drop(kind);
        }

        // Posted, never held: the empty document table already IS the cleared state, and the
        // ready handler re-opens from the table alone. A held clear replayed after ready's
        // re-opens would wrongly clear documents shown again since (the order a boot flushes
        // in is not the order a session said things in).
        if (_loaded)
        {
            Post(JsonSerializer.Serialize(
                new ClearDocumentMessage("clearDocuments"),
                EditorMessageContext.Default.ClearDocumentMessage));
        }
    }

    /// <summary>
    /// Drops every document that is no longer in the editor's open list, flushing any unwritten
    /// edits first — a pane closed natively mid-debounce still gets its write. The page prunes
    /// its own models from the same open list, so no message is needed.
    /// </summary>
    public void PruneDocuments(IReadOnlyCollection<(string Module, string? Project)> open)
    {
        ArgumentNullException.ThrowIfNull(open);

        var keep = new HashSet<string>(open.Select(pair => DocKey(pair.Module, pair.Project)), StringComparer.Ordinal);

        // Removed from the table BEFORE any flush callback runs: a flush reaches WriteModule,
        // which republishes, which prunes again — and the re-entrant pass must see a table
        // without the closing documents, not a half-walked one.
        List<OpenDoc>? closing = null;
        foreach (var (key, doc) in _docs.ToArray())
        {
            if (keep.Contains(key))
            {
                continue;
            }

            _docs.Remove(key);
            if (_activeKey == key)
            {
                _activeKey = null;
            }

            if (doc.Unwritten)
            {
                doc.Unwritten = false;
                (closing ??= []).Add(doc);
            }
        }

        foreach (var doc in closing ?? [])
        {
            TextChanged?.Invoke(doc.Module, doc.Project, doc.Text);
        }
    }

    /// <summary>
    /// Adopts the editor's version of a module without disturbing what the developer is doing.
    ///
    /// The editor is the text of record and it rewrites what it is given: it respells keywords and
    /// normalises spacing as it takes a module in. So the moment after an edit is written, the two
    /// disagree, and every later comparison would see a difference that is not the developer's.
    /// Adopting its version closes that immediately, at the one moment the difference is known to
    /// be the editor's doing and not an edit in flight.
    /// </summary>
    public void Sync(string moduleName, string? project, string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (!_docs.TryGetValue(DocKey(moduleName, project), out var doc))
        {
            return;
        }

        doc.Text = text;
        PostSyncDocument(moduleName, project, text);
    }

    /// <summary>
    /// Tells the developer something briefly.
    ///
    /// For the cases where an action is legitimately declined and silence would read as a fault.
    /// Not held for a page that is not up: a notice about something that happened before the
    /// surface existed is not worth showing when it finally does.
    /// </summary>
    public void Notify(string text)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(text);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new NoticeMessage("notice", text),
            EditorMessageContext.Default.NoticeMessage));
    }

    /// <summary>
    /// Runs one of the surface's own commands.
    ///
    /// Used for keys the host would otherwise take. A key claimed at the browser's accelerator hook
    /// never reaches the document, so the command the developer wanted has to be asked for
    /// explicitly rather than left to the page's own key handling.
    /// </summary>
    public void RunEditorCommand(string id)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);

        if (!_loaded)
        {
            Log.Info($"surface: command {id} dropped, the page is not up");
            return;
        }

        Post(JsonSerializer.Serialize(
            new EditorCommandMessage("editorCommand", id),
            EditorMessageContext.Default.EditorCommandMessage));
    }

    /// <summary>
    /// Answers a menu request with the items the editor holds right now.
    ///
    /// Never held: a reply only exists because the page asked, so the page is up, and a menu's
    /// contents are true at the moment they are read. Replaying them later would drop a menu down
    /// that nobody has open.
    /// </summary>
    public void ShowMenu(int[] path, SurfaceMenuItem[] items)
    {
        ArgumentNullException.ThrowIfNull(path);
        ArgumentNullException.ThrowIfNull(items);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SetMenuMessage("setMenu", path, items),
            EditorMessageContext.Default.SetMenuMessage));
    }

    /// <summary>The chrome state last sent, so a placement pass repeating it costs nothing.</summary>
    private bool? _chromeSent;

    /// <summary>Shows or withdraws the surface's own menu bar.</summary>
    public void SetChrome(bool menuBar)
    {
        // Placement re-derives on every resize tick and almost never changes this; a message
        // per tick is a cross-process hop the page answers by doing nothing.
        if (_chromeSent == menuBar)
        {
            return;
        }

        _chromeSent = menuBar;

        Send("setChrome", JsonSerializer.Serialize(
            new SetChromeMessage("setChrome", menuBar),
            EditorMessageContext.Default.SetChromeMessage));
    }

    /// <summary>
    /// Tells the surface where this library is loaded from, so its About dialog can say whether it
    /// is the installed build or one published straight out of the repository.
    /// </summary>
    public void ShowInstallPath(string? path)
    {
        Send("setInstallPath", JsonSerializer.Serialize(
            new SetInstallPathMessage("setInstallPath", path),
            EditorMessageContext.Default.SetInstallPathMessage));
    }

    /// <summary>Replaces the properties panel's contents with the selected component's properties.</summary>
    public void ShowProperties(string component, string kind, SurfacePropertyEntry[] properties)
    {
        ArgumentNullException.ThrowIfNull(component);
        ArgumentNullException.ThrowIfNull(kind);
        ArgumentNullException.ThrowIfNull(properties);

        Send("setProperties", JsonSerializer.Serialize(
            new SetPropertiesMessage("setProperties", component, kind, properties),
            EditorMessageContext.Default.SetPropertiesMessage));
    }

    /// <summary>Adds a line to the Immediate panel's output.</summary>
    public void ShowImmediateResult(string text, bool failed)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new ImmediateResultMessage("immediateResult", text, failed),
            EditorMessageContext.Default.ImmediateResultMessage));
    }

    /// <summary>True when any open document holds typing that has not reached its module.</summary>
    public bool HasUnwrittenEdits => _docs.Values.Any(doc => doc.Unwritten);

    /// <summary>The active document's text, or null when nothing is shown.</summary>
    public string? Text => ActiveDoc?.Text;

    /// <summary>A document's text as the surface holds it, or null when it is not open.</summary>
    public string? TextOf(string moduleName, string? project) =>
        _docs.TryGetValue(DocKey(moduleName, project), out var doc) ? doc.Text : null;

    /// <summary>True when this one document holds typing that has not reached its module.</summary>
    public bool HasUnwritten(string moduleName, string? project) =>
        _docs.TryGetValue(DocKey(moduleName, project), out var doc) && doc.Unwritten;

    /// <summary>
    /// Every document this surface holds text for, with enough to tell them apart.
    ///
    /// Which modules have TEXT and which merely have tabs are different lists — text arrives when
    /// a module is activated, a tab exists because its pane does — and two defects came from
    /// nothing ever showing the difference (2026-08-07).
    /// </summary>
    public IReadOnlyList<(string Module, string? Project, int Lines, bool Unwritten, bool Active)> DocumentTable =>
        [.. _docs.Select(entry => (
            entry.Value.Module,
            entry.Value.Project,
            entry.Value.Text.Length == 0 ? 0 : entry.Value.Text.Split('\n').Length,
            entry.Value.Unwritten,
            entry.Key == _activeKey))];

    /// <summary>Every open document, as (module, workbook display name) pairs.</summary>
    public IReadOnlyList<(string Module, string? Project)> OpenDocuments =>
        [.. _docs.Values.Select(doc => (doc.Module, doc.Project))];

    /// <summary>Replaces the squiggles shown on one open document's model.</summary>
    public void ShowDiagnostics(string moduleName, string? project, EditorMarker[] markers)
    {
        ArgumentNullException.ThrowIfNull(markers);

        // Keyed per document rather than one held slot: with several models live, the newest
        // set for one module must not evict the newest set for another.
        Send($"setDiagnostics:{(project ?? string.Empty).ToLowerInvariant()}/{moduleName.ToLowerInvariant()}", JsonSerializer.Serialize(
            new SetDiagnosticsMessage("setDiagnostics", moduleName, project, markers),
            EditorMessageContext.Default.SetDiagnosticsMessage));
    }

    /// <summary>Sends the developer's settings, for the page's dialog and typing behaviour.</summary>
    public void ShowSettings(ProductSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        Send("setSettings", JsonSerializer.Serialize(
            new SetSettingsMessage(
                "setSettings",
                settings.BlockLayout,
                settings.ContinueCommentOnNewline,
                settings.MirrorCommentSpacing,
                settings.TreeFollowsEditor,
                settings.FormatIndentSize,
                settings.FormatCanonicalKeywords,
                settings.SyncEngine),
            EditorMessageContext.Default.SetSettingsMessage));
    }

    /// <summary>Replaces the tab strip: every module the editor has open, and which one is shown.</summary>
    public void ShowModules(string[] modules, string?[] projects, string? active, string? activeProject, bool[]? dirty = null)
    {
        ArgumentNullException.ThrowIfNull(modules);
        ArgumentNullException.ThrowIfNull(projects);

        Send("setModules", JsonSerializer.Serialize(
            new SetModulesMessage("setModules", modules, projects, active, activeProject, dirty),
            EditorMessageContext.Default.SetModulesMessage));
    }

    /// <summary>Replaces the project explorer's contents.</summary>
    public void ShowProjects(SurfaceProject[] projects)
    {
        ArgumentNullException.ThrowIfNull(projects);

        Send("setProjects", JsonSerializer.Serialize(
            new SetProjectsMessage("setProjects", projects),
            EditorMessageContext.Default.SetProjectsMessage));
    }

    /// <summary>
    /// Marks the line execution is stopped on, or clears the mark.
    ///
    /// Not held for a page that is still loading: where execution was stopped seconds ago is not
    /// where it is stopped now, and a stale arrow is worse than none.
    /// </summary>
    public void ShowCurrentLine(int? line)
    {
        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SetCurrentLineMessage("setCurrentLine", line),
            EditorMessageContext.Default.SetCurrentLineMessage));
    }

    /// <summary>
    /// Replaces what the Locals panel shows. Stopped false is the idle state; stopped true with
    /// no rows is a break whose variables cannot be read, shown honestly as an empty break.
    /// Not held for a loading page for the same reason as the current line: stale variables are
    /// worse than none.
    /// </summary>
    public void ShowLocals(bool stopped, string? context, SurfaceLocalRow[] rows)
    {
        ArgumentNullException.ThrowIfNull(rows);

        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SetLocalsMessage("setLocals", stopped, context, rows),
            EditorMessageContext.Default.SetLocalsMessage));
    }

    /// <summary>Marks a line whose breakpoint the host refused: a brief orange cross.</summary>
    public void ShowBreakpointRefused(int line)
    {
        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new BreakpointRefusedMessage("breakpointRefused", line),
            EditorMessageContext.Default.BreakpointRefusedMessage));
    }

    /// <summary>Which debug mode the editor is in; held so a late page still learns it.</summary>
    public void ShowDebugMode(string mode)
    {
        Send("setDebugState", JsonSerializer.Serialize(
            new SetDebugStateMessage("setDebugState", mode),
            EditorMessageContext.Default.SetDebugStateMessage));
    }

    /// <summary>Replaces the Watch panel's rows; not stopped clears it to its idle text.</summary>
    public void ShowWatches(bool stopped, SurfaceWatchRow[] rows)
    {
        ArgumentNullException.ThrowIfNull(rows);

        Send("setWatches", JsonSerializer.Serialize(
            new SetWatchesMessage("setWatches", stopped, rows),
            EditorMessageContext.Default.SetWatchesMessage));
    }

    /// <summary>
    /// Asks the developer what to do about a module's unsaved changes before its tab closes.
    /// Not held: the request answers a close the page just asked for, so the page is there.
    /// </summary>
    public void ConfirmClose(string module, string? project)
    {
        if (!_loaded)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new ConfirmCloseMessage("confirmClose", module, project),
            EditorMessageContext.Default.ConfirmCloseMessage));
    }

    /// <summary>Replaces the breakpoints shown on the module currently displayed.</summary>
    public void ShowBreakpoints(int[] lines)
    {
        ArgumentNullException.ThrowIfNull(lines);

        Send("setBreakpoints", JsonSerializer.Serialize(
            new SetBreakpointsMessage("setBreakpoints", lines),
            EditorMessageContext.Default.SetBreakpointsMessage));
    }

    /// <summary>Raised on every tick of the poll timer, on the host thread.</summary>
    public Action? Polled { get; set; }

    /// <summary>Starts or stops the poll timer, which is how execution state is watched.</summary>
    public void Poll(uint milliseconds)
    {
        if (milliseconds == 0)
        {
            _overlay?.StopPollTimer();
        }
        else
        {
            _overlay?.StartPollTimer(milliseconds);
        }
    }

    /// <summary>Raised when the developer clicks the glyph margin to toggle a breakpoint.</summary>
    public Action<int>? BreakpointToggleRequested { get; set; }

    /// <summary>
    /// The text of a one-based line as the surface currently has it, or null when there is no such
    /// line. This is the surface's copy, which is what the developer is looking at.
    /// </summary>
    public string? LineAt(int line)
    {
        if (ActiveDoc?.Text is not { } text || line < 1)
        {
            return null;
        }

        var start = 0;
        for (var current = 1; ; current++)
        {
            var end = text.IndexOf('\n', start);
            if (current == line)
            {
                var lineText = end < 0 ? text[start..] : text[start..end];
                return lineText.TrimEnd('\r');
            }

            if (end < 0)
            {
                return null;
            }

            start = end + 1;
        }
    }

    /// <summary>Replaces the panel's contents, across every module.</summary>
    public void ShowFindings(SurfaceFinding[] findings)
    {
        ArgumentNullException.ThrowIfNull(findings);

        Send("setFindings", JsonSerializer.Serialize(
            new SetFindingsMessage("setFindings", findings),
            EditorMessageContext.Default.SetFindingsMessage));
    }

    /// <summary>
    /// Scrolls a one-based line into view. Not held: where the developer wanted to be some seconds
    /// before the page existed is not where they want to be now.
    /// </summary>
    public void Reveal(int line)
    {
        if (!_loaded || line < 1)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new RevealLineMessage("revealLine", line),
            EditorMessageContext.Default.RevealLineMessage));
    }

    /// <summary>Puts the caret on a line and reveals it. See SetCaretMessage for why both.</summary>
    public void SetCaret(int line, int column)
    {
        if (!_loaded || line < 1)
        {
            return;
        }

        Post(JsonSerializer.Serialize(
            new SetCaretMessage("setCaret", line, Math.Max(1, column)),
            EditorMessageContext.Default.SetCaretMessage));
    }

    /// <summary>
    /// Writes the developer's edits back to the module now, if there are any.
    ///
    /// Called before anything that reads the module: running, stepping, switching module, or
    /// shutting down. Without it the host compiles and runs the text as it was before the developer
    /// started typing, which looks like an editor that does not save.
    /// </summary>
    public void FlushEdits()
    {
        _overlay?.StopWriteTimer();

        // Flags drop before the callbacks run: a flush reaches WriteModule and republishes,
        // and the re-entrant pass must find nothing left to flush.
        List<OpenDoc>? unwritten = null;
        foreach (var doc in _docs.Values)
        {
            if (doc.Unwritten)
            {
                doc.Unwritten = false;
                (unwritten ??= []).Add(doc);
            }
        }

        foreach (var doc in unwritten ?? [])
        {
            TextChanged?.Invoke(doc.Module, doc.Project, doc.Text);
        }
    }

    /// <summary>
    /// Forgets one document's unwritten edits instead of writing them. For the moment the
    /// developer has chosen to discard: the module is about to be put back to its saved text,
    /// and the debounced write of the abandoned text must not land on top of it. Only this
    /// document forgets — a sibling tab's pending edits are not the developer's answer here.
    /// </summary>
    public void DiscardEdits(string moduleName, string? project)
    {
        if (_docs.TryGetValue(DocKey(moduleName, project), out var doc))
        {
            doc.Unwritten = false;
        }

        // The write timer stays armed if some other document still owes a write.
        if (!_docs.Values.Any(other => other.Unwritten))
        {
            _overlay?.StopWriteTimer();
        }
    }

    /// <summary>Sends a message, or holds it until the page is ready for it.</summary>
    private void Send(string kind, string json)
    {
        if (_loaded)
        {
            Post(json);
            return;
        }

        Log.Verbose($"page <- {kind} held for a page that is not ready, {json.Length} chars");

        if (!_pending.ContainsKey(kind))
        {
            _pendingOrder.Add(kind);
        }

        _pending[kind] = json;
    }

    private void Drop(string kind)
    {
        if (_pending.Remove(kind))
        {
            _pendingOrder.Remove(kind);
        }
    }

    private void Flush()
    {
        foreach (var kind in _pendingOrder)
        {
            Post(_pending[kind]);
        }

        _pendingOrder.Clear();
        _pending.Clear();
    }

    private void OnMessage(string payload)
    {
        try
        {
            using var document = JsonDocument.Parse(payload);
            if (!document.RootElement.TryGetProperty("type", out var type))
            {
                return;
            }

            // Traffic in, one line per message. Bursts of identical shapes collapse in the log;
            // the payload length keeps two different edits from reading as a repeat.
            Log.Verbose($"page -> {type.GetString()}, {payload.Length} chars");

            switch (type.GetString())
            {
                case "ready":
                    _loaded = true;

                    // A fresh page has the default chrome, whatever was last said to the
                    // page before it; the cache would otherwise skip re-saying it.
                    _chromeSent = null;
                    Log.Info($"editor surface: ready{DescribeTimings(document.RootElement)}");
#if DEBUG
                    PageBuildStamp = document.RootElement.TryGetProperty("timings", out var readyTimings)
                        && readyTimings.TryGetProperty("build", out var readyBuild)
                        ? readyBuild.GetString()
                        : null;
#endif
                    // Every live document is (re)opened before anything held is flushed: a page
                    // that reloaded mid-session lost its models, and the messages behind it —
                    // squiggles, tabs — land on models, so the models come first. An open the
                    // page already has is adopted idempotently.
                    foreach (var doc in _docs.Values.ToArray())
                    {
                        PostOpenDocument(doc.Module, doc.Project, doc.Text);
                    }

                    Flush();

                    // Only now is the browser worth looking at: the page is styled and has its
                    // flushed state. The loader retires in the same breath.
                    _browser?.Reveal();
                    _overlay?.HideLoader();
                    Ready?.Invoke();
                    break;

                case "selectionChanged":
                    if (document.RootElement.TryGetProperty("startLine", out var line)
                        && line.TryGetInt32(out var caretLine))
                    {
                        // Leaving an edited line settles it: the deferred pass runs now, while
                        // staying on the line keeps the engine free for the completions and
                        // call tips the typing is asking for.
                        if (caretLine != CaretLine && _changedSinceLiveAnalysis)
                        {
                            _overlay?.StartAnalyseTimer(1);
                        }

                        CaretLine = caretLine;
                        CaretColumn = document.RootElement.TryGetProperty("startColumn", out var column)
                            && column.TryGetInt32(out var caretColumn)
                            ? caretColumn
                            : 1;

                        // After the properties update, so whatever this triggers reads the
                        // caret as settled, not as it was.
                        CaretLineSettled?.Invoke(caretLine);
                    }

                    break;

                case "contentChanged":
                {
                    // The message names the document it edits (decision 12); the active one
                    // stands in only for a page too old to say, which no shipped pairing is.
                    var editedDoc = document.RootElement.TryGetProperty("moduleName", out var editedName)
                        && editedName.GetString() is { Length: > 0 } editedModule
                        ? (_docs.TryGetValue(
                            DocKey(
                                editedModule,
                                document.RootElement.TryGetProperty("project", out var editedProject)
                                    ? editedProject.GetString()
                                    : null),
                            out var addressed)
                            ? addressed
                            : null)
                        : ActiveDoc;

                    if (editedDoc is not null)
                    {
                        // A small module arrives whole; a large one arrives as its changes and
                        // is reconstructed here, because building and shipping the full text of
                        // a large module per keystroke is what typing latency is made of.
                        var updated = document.RootElement.TryGetProperty("fullText", out var text)
                            ? text.GetString()
                            : null;

                        EngineTextEdit[]? parsedEdits = null;
                        if (updated is null
                            && document.RootElement.TryGetProperty("changes", out var changeSet)
                            && changeSet.ValueKind == JsonValueKind.Array)
                        {
                            parsedEdits = ParseChanges(editedDoc.Text, changeSet);
                            updated = parsedEdits is null ? null : ApplyEdits(editedDoc.Text, parsedEdits);
                        }

                        if (updated is not null)
                        {
                            if (document.RootElement.TryGetProperty("fullLength", out var lengthElement)
                                && lengthElement.TryGetInt32(out var expectedLength)
                                && updated.Length != expectedLength)
                            {
                                Log.Error($"surface: reconstructed {updated.Length} character(s) where the page holds {expectedLength}");
                            }

                            // Held, not written. Writing a module replaces its changed lines,
                            // which resets the project, so doing it per keystroke would reset
                            // the project on every keystroke. The edit is written once the
                            // developer stops typing, and immediately before anything that has
                            // to see it.
                            editedDoc.Text = updated;
                            editedDoc.Unwritten = true;
                            _overlay?.StartWriteTimer(WriteDelayMilliseconds);

                            // Breakpoints are line-anchored bookkeeping, and edits move lines.
                            // Each change that adds or removes lines shifts every anchor below
                            // it, so a dot stays on the statement it was set on instead of
                            // drifting onto whatever scrolled into its number — the ghost dot
                            // no click could remove (2026-08-04).
                            if (document.RootElement.TryGetProperty("changes", out var shiftSet)
                                && shiftSet.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var change in shiftSet.EnumerateArray())
                                {
                                    if (!change.TryGetProperty("startLine", out var startElement)
                                        || !startElement.TryGetInt32(out var startLine)
                                        || !change.TryGetProperty("endLine", out var endElement)
                                        || !endElement.TryGetInt32(out var endLine))
                                    {
                                        continue;
                                    }

                                    var newlines = 0;
                                    var body = change.TryGetProperty("text", out var textElement)
                                        ? textElement.GetString() ?? string.Empty
                                        : string.Empty;
                                    foreach (var character in body)
                                    {
                                        if (character == '\n')
                                        {
                                            newlines++;
                                        }
                                    }

                                    var delta = newlines - (endLine - startLine);
                                    if (delta != 0)
                                    {
                                        LinesShifted?.Invoke(editedDoc.Module, startLine, delta);
                                    }
                                }
                            }

                            // The findings shown must describe this text, not the text as of
                            // the last write: a deleted error must go, and it must go soon —
                            // where soon costs little, and on the line's own goodbye otherwise.
                            _changedSinceLiveAnalysis = true;
                            _overlay?.StartAnalyseTimer(updated.Length >= LargeModuleCharacters
                                ? LiveAnalysisDelayLargeMilliseconds
                                : LiveAnalysisDelayMilliseconds);

                            // The engine mirrors this text and answers requests from its copy,
                            // so the change stream continues to it — small edits for a large
                            // module, the whole text for a small one — ordered ahead of any
                            // request about the text it makes.
                            LiveTextPushed?.Invoke(editedDoc.Module, parsedEdits is null ? updated : null, parsedEdits);

                            // Read from the raw change set rather than the parsed edits, because
                            // small modules skip the parse and this must fire for them too.
                            //
                            // A change the page has named the source of is not typing, whatever
                            // its shape. Formatting a module whose lines are already right except
                            // one produces a one-line change with no newline in it, which is
                            // exactly what a keystroke produces, and reading it as one armed the
                            // hold and took the squiggle off that line until the caret was moved
                            // away (2026-08-07). The page marks its formatter's edits because the
                            // editor's change event carries no source of its own.
                            var typedByHand = !document.RootElement.TryGetProperty("source", out var madeBy)
                                || madeBy.ValueKind != JsonValueKind.String;

                            if (typedByHand
                                && document.RootElement.TryGetProperty("changes", out var typedSet)
                                && SingleLineTypedIn(typedSet) is { } typedLine)
                            {
                                LineTyped?.Invoke(typedLine);
                            }
                        }
                    }

                    break;
                }

                case "activateModule":
                    if (document.RootElement.TryGetProperty("moduleName", out var requested)
                        && requested.GetString() is { Length: > 0 } name)
                    {
                        var requestedProject = document.RootElement.TryGetProperty("project", out var wanted)
                            ? wanted.GetString()
                            : null;
                        Log.Info($"surface: activate {name} requested"
                            + (requestedProject is null ? string.Empty : $" in {requestedProject}"));
                        ModuleRequested?.Invoke(name, requestedProject);
                    }

                    break;

                case "navigate":
                    OnNavigate(document.RootElement);
                    break;

                case "breakpointToggleRequested":
                    if (document.RootElement.TryGetProperty("line", out var breakpointLine)
                        && breakpointLine.TryGetInt32(out var toggled))
                    {
                        BreakpointToggleRequested?.Invoke(toggled);
                    }

                    break;

                case "panel":
                    if (document.RootElement.TryGetProperty("name", out var panel)
                        && panel.GetString() is { Length: > 0 } panelName)
                    {
                        var open = !document.RootElement.TryGetProperty("open", out var isOpen)
                            || isOpen.ValueKind != JsonValueKind.False;

                        PanelChanged?.Invoke(panelName, open);
                    }

                    break;

                case "evaluate":
                    if (document.RootElement.TryGetProperty("text", out var expression)
                        && expression.GetString() is { Length: > 0 } entered)
                    {
                        EvaluateRequested?.Invoke(entered);
                    }

                    break;

                case "openExternal":
                    if (document.RootElement.TryGetProperty("url", out var address)
                        && address.GetString() is { Length: > 0 } addressAsked)
                    {
                        ExternalOpenRequested?.Invoke(addressAsked);
                    }

                    break;

                case "requestDocument":
                    if (document.RootElement.TryGetProperty("module", out var documentName)
                        && documentName.GetString() is { Length: > 0 } documentAsked)
                    {
                        DocumentRequested?.Invoke(
                            documentAsked,
                            document.RootElement.TryGetProperty("project", out var documentOwner)
                                ? documentOwner.GetString()
                                : null);
                    }

                    break;

                case "command":
                    if (document.RootElement.TryGetProperty("name", out var command)
                        && command.GetString() is { Length: > 0 } commandName)
                    {
                        CommandRequested?.Invoke(commandName);
                    }

                    break;

                case "menu":
                    if (PathOf(document.RootElement) is { } menuPath)
                    {
                        MenuRequested?.Invoke(menuPath);
                    }

                    break;

                case "menuExecute":
                    // An empty path names the bar itself, which can be asked about but not run.
                    if (PathOf(document.RootElement) is { Length: > 0 } executePath)
                    {
                        MenuExecuteRequested?.Invoke(executePath);
                    }

                    break;

                case "selectComponent":
                    if (document.RootElement.TryGetProperty("name", out var selected)
                        && selected.GetString() is { Length: > 0 } selectedName)
                    {
                        ComponentSelected?.Invoke(selectedName);
                    }

                    break;


                case "closeModule":
                    if (document.RootElement.TryGetProperty("name", out var closing)
                        && closing.GetString() is { Length: > 0 } closingName)
                    {
                        var closingProject = document.RootElement.TryGetProperty("project", out var closingOwner)
                            ? closingOwner.GetString()
                            : null;
                        var closingAction = document.RootElement.TryGetProperty("action", out var closingChoice)
                            ? closingChoice.GetString()
                            : null;
                        ModuleCloseRequested?.Invoke(closingName, closingProject, closingAction);
                    }

                    break;

                case "updateSettings":
                {
                    var layout = document.RootElement.TryGetProperty("blockLayout", out var layoutValue)
                        ? layoutValue.GetString() ?? "comfy"
                        : "comfy";
                    var continueComment = !document.RootElement.TryGetProperty("continueCommentOnNewline", out var continueValue)
                        || continueValue.ValueKind is not (JsonValueKind.False);
                    var mirrorSpacing = !document.RootElement.TryGetProperty("mirrorCommentSpacing", out var mirrorValue)
                        || mirrorValue.ValueKind is not (JsonValueKind.False);
                    var treeFollows = !document.RootElement.TryGetProperty("treeFollowsEditor", out var treeValue)
                        || treeValue.ValueKind is not JsonValueKind.False;
                    var indentSize = document.RootElement.TryGetProperty("formatIndentSize", out var indentValue)
                        && indentValue.TryGetInt32(out var asked)
                        ? asked
                        : 4;
                    var canonicalKeywords = !document.RootElement.TryGetProperty("formatCanonicalKeywords", out var keywordsValue)
                        || keywordsValue.ValueKind is not (JsonValueKind.False);

                    var syncEngine = document.RootElement.TryGetProperty("syncEngine", out var engineValue)
                        && engineValue.ValueKind == JsonValueKind.String
                            ? engineValue.GetString() ?? "xlide"
                            : "xlide";

                    SettingsChangeRequested?.Invoke(new ProductSettings
                    {
                        BlockLayout = layout,
                        ContinueCommentOnNewline = continueComment,
                        MirrorCommentSpacing = mirrorSpacing,
                        TreeFollowsEditor = treeFollows,
                        FormatIndentSize = indentSize,
                        FormatCanonicalKeywords = canonicalKeywords,
                        SyncEngine = syncEngine,
                    }.Normalized());
                    break;
                }

                case "insertComponent":
                    if (document.RootElement.TryGetProperty("kind", out var componentKind)
                        && componentKind.TryGetInt32(out var insertKind))
                    {
                        var targetProject = document.RootElement.TryGetProperty("project", out var projectElement)
                            ? projectElement.GetString()
                            : null;

                        ComponentInsertRequested?.Invoke(insertKind, targetProject);
                    }

                    break;

                case "completion":
                    if (document.RootElement.TryGetProperty("id", out var completionId)
                        && completionId.TryGetInt32(out var requestId)
                        && document.RootElement.TryGetProperty("offset", out var caretOffset)
                        && caretOffset.TryGetInt32(out var offset)
                        && offset >= 0)
                    {
                        CompletionRequested?.Invoke(requestId, offset);
                    }

                    break;

                case "hover":
                    if (document.RootElement.TryGetProperty("id", out var hoverId)
                        && hoverId.TryGetInt32(out var hoverRequestId)
                        && document.RootElement.TryGetProperty("offset", out var hoverOffsetElement)
                        && hoverOffsetElement.TryGetInt32(out var hoverOffset)
                        && hoverOffset >= 0)
                    {
                        HoverRequested?.Invoke(hoverRequestId, hoverOffset);
                    }

                    break;

                case "signatureHelp":
                    if (document.RootElement.TryGetProperty("id", out var signatureId)
                        && signatureId.TryGetInt32(out var signatureRequestId)
                        && document.RootElement.TryGetProperty("offset", out var signatureOffsetElement)
                        && signatureOffsetElement.TryGetInt32(out var signatureOffset)
                        && signatureOffset >= 0)
                    {
                        SignatureHelpRequested?.Invoke(signatureRequestId, signatureOffset);
                    }

                    break;

                case "smartEnter":
                    if (document.RootElement.TryGetProperty("id", out var smartEnterId)
                        && smartEnterId.TryGetInt32(out var smartEnterRequestId)
                        && document.RootElement.TryGetProperty("offset", out var smartEnterOffsetElement)
                        && smartEnterOffsetElement.TryGetInt32(out var smartEnterOffset)
                        && smartEnterOffset >= 0)
                    {
                        SmartEnterRequested?.Invoke(smartEnterRequestId, smartEnterOffset);
                    }

                    break;

                case "canonicalCase":
                    if (document.RootElement.TryGetProperty("id", out var caseId)
                        && caseId.TryGetInt32(out var caseRequestId)
                        && document.RootElement.TryGetProperty("start", out var caseStartElement)
                        && caseStartElement.TryGetInt32(out var caseStart)
                        && document.RootElement.TryGetProperty("end", out var caseEndElement)
                        && caseEndElement.TryGetInt32(out var caseEnd)
                        && caseStart >= 0
                        && caseEnd >= caseStart)
                    {
                        var single = document.RootElement.TryGetProperty("single", out var singleElement)
                            && singleElement.ValueKind == JsonValueKind.True;
                        var completeHeader = document.RootElement.TryGetProperty("completeHeader", out var headerElement)
                            && headerElement.ValueKind == JsonValueKind.True;

                        CanonicalCaseRequested?.Invoke(caseRequestId, caseStart, caseEnd, single, completeHeader);
                    }

                    break;

                case "codeAction":
                    if (document.RootElement.TryGetProperty("id", out var fixId)
                        && fixId.TryGetInt32(out var fixRequestId)
                        && document.RootElement.TryGetProperty("start", out var fixStartElement)
                        && fixStartElement.TryGetInt32(out var fixStart)
                        && document.RootElement.TryGetProperty("end", out var fixEndElement)
                        && fixEndElement.TryGetInt32(out var fixEnd)
                        && fixStart >= 0
                        && fixEnd >= fixStart)
                    {
                        CodeActionsRequested?.Invoke(fixRequestId, fixStart, fixEnd);
                    }

                    break;

                case "undoRename":
                    if (document.RootElement.TryGetProperty("id", out var undoId)
                        && undoId.TryGetInt32(out var undoRequestId))
                    {
                        RenameUndoRequested?.Invoke(undoRequestId);
                    }

                    break;

                case "rename":
                    if (document.RootElement.TryGetProperty("id", out var renameId)
                        && renameId.TryGetInt32(out var renameRequestId)
                        && document.RootElement.TryGetProperty("offset", out var renameOffsetElement)
                        && renameOffsetElement.TryGetInt32(out var renameOffset)
                        && renameOffset >= 0
                        && document.RootElement.TryGetProperty("newName", out var newNameElement)
                        && newNameElement.GetString() is { Length: > 0 } renameNewName)
                    {
                        RenameRequested?.Invoke(renameRequestId, renameOffset, renameNewName);
                    }

                    break;

                case "renameModule":
                    if (document.RootElement.TryGetProperty("id", out var modRenameId)
                        && modRenameId.TryGetInt32(out var modRenameRequestId)
                        && document.RootElement.TryGetProperty("module", out var modRenameElement)
                        && modRenameElement.GetString() is { Length: > 0 } modRenameModule
                        && document.RootElement.TryGetProperty("newName", out var modNewNameElement)
                        && modNewNameElement.GetString() is { Length: > 0 } modNewName)
                    {
                        var modProject = document.RootElement.TryGetProperty("project", out var modOwner)
                            ? modOwner.GetString()
                            : null;
                        ModuleRenameRequested?.Invoke(modRenameRequestId, modRenameModule, modProject, modNewName);
                    }

                    break;

                case "definition":
                case "references":
                    if (document.RootElement.TryGetProperty("id", out var navId)
                        && navId.TryGetInt32(out var navRequestId)
                        && document.RootElement.TryGetProperty("offset", out var navOffsetElement)
                        && navOffsetElement.TryGetInt32(out var navOffset)
                        && navOffset >= 0)
                    {
                        // The declaration counts as a use unless the page says otherwise, which
                        // is what the editor's own "find all references" means by the word.
                        var includeDeclaration =
                            !document.RootElement.TryGetProperty("includeDeclaration", out var declElement)
                            || declElement.ValueKind != JsonValueKind.False;

                        NavigationRequested?.Invoke(
                            navRequestId,
                            navOffset,
                            type.GetString() == "references",
                            includeDeclaration);
                    }

                    break;

                case "loopSync":
                    if (document.RootElement.TryGetProperty("id", out var loopId)
                        && loopId.TryGetInt32(out var loopRequestId)
                        && document.RootElement.TryGetProperty("offset", out var loopOffsetElement)
                        && loopOffsetElement.TryGetInt32(out var loopOffset)
                        && loopOffset >= 0)
                    {
                        LoopSyncRequested?.Invoke(loopRequestId, loopOffset);
                    }

                    break;

                case "sync":
                    if (document.RootElement.TryGetProperty("id", out var syncId)
                        && syncId.TryGetInt32(out var syncRequestId))
                    {
                        // Everything but the id is passed through as the service's own arguments,
                        // so the dialog and the debug api hand it identical requests.
                        var syncArguments = new Dictionary<string, string>(StringComparer.Ordinal);
                        foreach (var syncField in document.RootElement.EnumerateObject())
                        {
                            if (syncField.Name is "type" or "id" or "body")
                            {
                                continue;
                            }

                            if (syncField.Value.ValueKind == JsonValueKind.String
                                && syncField.Value.GetString() is { Length: > 0 } argument)
                            {
                                syncArguments[syncField.Name] = argument;
                            }
                        }

                        var syncBody = document.RootElement.TryGetProperty("body", out var syncBodyElement)
                            ? syncBodyElement.GetString() ?? string.Empty
                            : string.Empty;
                        SyncRequested?.Invoke(syncRequestId, syncArguments, syncBody);
                    }

                    break;

                case "outline":
                    if (document.RootElement.TryGetProperty("id", out var outlineId)
                        && outlineId.TryGetInt32(out var outlineRequestId)
                        && document.RootElement.TryGetProperty("module", out var outlineModuleElement)
                        && outlineModuleElement.GetString() is { Length: > 0 } outlineModule)
                    {
                        var outlineProject = document.RootElement.TryGetProperty("project", out var outlineOwner)
                            ? outlineOwner.GetString()
                            : null;
                        OutlineRequested?.Invoke(outlineRequestId, outlineModule, outlineProject);
                    }

                    break;

                case "semanticTokens":
                    if (document.RootElement.TryGetProperty("id", out var colourId)
                        && colourId.TryGetInt32(out var colourRequestId)
                        && document.RootElement.TryGetProperty("module", out var colourModuleElement)
                        && colourModuleElement.GetString() is { Length: > 0 } colourModule)
                    {
                        var colourProject = document.RootElement.TryGetProperty("project", out var colourOwner)
                            ? colourOwner.GetString()
                            : null;
                        SemanticTokensRequested?.Invoke(colourRequestId, colourModule, colourProject);
                    }

                    break;

                case "search":
                case "replaceAll":
                    if (document.RootElement.TryGetProperty("id", out var searchId)
                        && searchId.TryGetInt32(out var searchRequestId)
                        && document.RootElement.TryGetProperty("query", out var queryElement)
                        && queryElement.GetString() is { Length: > 0 } query)
                    {
                        var matchCase = document.RootElement.TryGetProperty("matchCase", out var caseElement)
                            && caseElement.ValueKind is JsonValueKind.True;
                        var wholeWord = document.RootElement.TryGetProperty("wholeWord", out var wordElement)
                            && wordElement.ValueKind is JsonValueKind.True;
                        var scope = document.RootElement.TryGetProperty("scope", out var scopeElement)
                            ? scopeElement.GetString() ?? "module"
                            : "module";

                        if (type.GetString() == "search")
                        {
                            SearchRequested?.Invoke(searchRequestId, query, matchCase, wholeWord, scope);
                        }
                        else
                        {
                            var replacement = document.RootElement.TryGetProperty("replacement", out var replacementElement)
                                ? replacementElement.GetString() ?? string.Empty
                                : string.Empty;
                            ReplaceAllRequested?.Invoke(searchRequestId, query, matchCase, wholeWord, scope, replacement);
                        }
                    }

                    break;

                case "trace":
                    // The page's own words, in the only log support can read.
                    if (document.RootElement.TryGetProperty("text", out var trace)
                        && trace.GetString() is { Length: > 0 } traced)
                    {
                        Log.Info($"page: {traced}");
                    }

                    break;

                case "editProperty":
                    if (document.RootElement.TryGetProperty("component", out var owner)
                        && owner.GetString() is { Length: > 0 } ownerName
                        && document.RootElement.TryGetProperty("name", out var propertyName)
                        && propertyName.GetString() is { Length: > 0 } property
                        && document.RootElement.TryGetProperty("value", out var newValue)
                        && newValue.GetString() is { } value)
                    {
                        PropertyEditRequested?.Invoke(ownerName, property, value);
                    }

                    break;
            }
        }
        catch (JsonException)
        {
            Log.Warn("editor surface: a message from the page was not valid");
        }
    }

    /// <summary>
    /// Converts the page's change ranges to offsets, computed once against the text the ranges
    /// describe. The ranges arrive bottom-up, and stay in that order. Null when a range is
    /// malformed, and the caller keeps the text it has rather than guessing.
    /// </summary>
    private static EngineTextEdit[]? ParseChanges(string text, JsonElement changes)
    {
        var lineStarts = TextPositions.LineStarts(text);
        var edits = new List<EngineTextEdit>(changes.GetArrayLength());

        foreach (var change in changes.EnumerateArray())
        {
            if (!change.TryGetProperty("startLine", out var startLineElement)
                || !startLineElement.TryGetInt32(out var startLine)
                || !change.TryGetProperty("startColumn", out var startColumnElement)
                || !startColumnElement.TryGetInt32(out var startColumn)
                || !change.TryGetProperty("endLine", out var endLineElement)
                || !endLineElement.TryGetInt32(out var endLine)
                || !change.TryGetProperty("endColumn", out var endColumnElement)
                || !endColumnElement.TryGetInt32(out var endColumn)
                || !change.TryGetProperty("text", out var replacement)
                || replacement.GetString() is not { } inserted)
            {
                return null;
            }

            edits.Add(new EngineTextEdit(
                TextPositions.ToOffset(lineStarts, startLine, startColumn),
                TextPositions.ToOffset(lineStarts, endLine, endColumn),
                inserted));
        }

        return [.. edits];
    }

    /// <summary>
    /// The one line a change event typed on, or null when it is anything else: a newline, a
    /// paste across lines, a multi-cursor edit. Only plain typing begins the active-line hold,
    /// because everything else is precisely the "leaving the line" a hold ends on.
    /// </summary>
    private static int? SingleLineTypedIn(JsonElement changes)
    {
        if (changes.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        int? line = null;
        foreach (var change in changes.EnumerateArray())
        {
            if (!change.TryGetProperty("startLine", out var startLineElement)
                || !startLineElement.TryGetInt32(out var startLine)
                || !change.TryGetProperty("endLine", out var endLineElement)
                || !endLineElement.TryGetInt32(out var endLine)
                || startLine != endLine
                || (line is { } known && known != startLine)
                || !change.TryGetProperty("text", out var replacement)
                || replacement.GetString() is not { } inserted
                || inserted.Contains('\n', StringComparison.Ordinal)
                || inserted.Contains('\r', StringComparison.Ordinal))
            {
                return null;
            }

            line = startLine;
        }

        return line;
    }

    /// <summary>
    /// Applies bottom-up edits from the tail forward, so every earlier offset still means what
    /// it meant. Null when an edit is out of bounds.
    /// </summary>
    private static string? ApplyEdits(string text, EngineTextEdit[] edits)
    {
        var updated = text;

        foreach (var edit in edits)
        {
            if (edit.Start < 0 || edit.End < edit.Start || edit.End > updated.Length)
            {
                return null;
            }

            updated = string.Concat(updated.AsSpan(0, edit.Start), edit.Text, updated.AsSpan(edit.End));
        }

        return updated;
    }

    /// <summary>
    /// The position chain in a menu message, or null when it is malformed. Positions are one-based
    /// because the editor's collections are.
    /// </summary>
    private static int[]? PathOf(JsonElement message)
    {
        if (!message.TryGetProperty("path", out var path) || path.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        var result = new int[path.GetArrayLength()];
        var i = 0;

        foreach (var element in path.EnumerateArray())
        {
            if (!element.TryGetInt32(out result[i]) || result[i] < 1)
            {
                return null;
            }

            i++;
        }

        return result;
    }

    private void OnNavigate(JsonElement message)
    {
        if (!message.TryGetProperty("module", out var module)
            || module.GetString() is not { Length: > 0 } component)
        {
            return;
        }

        var line = message.TryGetProperty("line", out var lineValue) && lineValue.TryGetInt32(out var l) ? l : 1;
        var column = message.TryGetProperty("column", out var columnValue) && columnValue.TryGetInt32(out var c) ? c : 1;
        var project = message.TryGetProperty("project", out var projectValue) ? projectValue.GetString() : null;

        NavigateRequested?.Invoke(component, line, column, project);
    }

    /// <summary>
    /// Renders what the page reported about its own start-up, so the cost of putting a surface over
    /// a pane is a measured number in the log rather than an impression.
    /// </summary>
#if DEBUG
    /// <summary>
    /// The bundle stamp the page reported when it booted, kept for the debug api's doctor
    /// route. A page and a shim built at different times is the commonest cause of "my fix
    /// is not in the log", and the only cure is being able to ask.
    /// </summary>
    internal string? PageBuildStamp { get; private set; }
#endif

    private static string DescribeTimings(JsonElement message)
    {
        if (!message.TryGetProperty("timings", out var timings) || timings.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        var script = timings.TryGetProperty("scriptMs", out var s) && s.TryGetInt32(out var scriptMs) ? scriptMs : -1;
        var create = timings.TryGetProperty("createMs", out var c) && c.TryGetInt32(out var createMs) ? createMs : -1;
        var total = timings.TryGetProperty("totalMs", out var t) && t.TryGetInt32(out var totalMs) ? totalMs : -1;
        var build = timings.TryGetProperty("build", out var b) ? b.GetString() : null;

        // The split that decides where start-up work goes: fetch is the host serving bytes,
        // bundle-minus-fetch is the browser compiling them, and transfer 0 is the cache.
        var detail = string.Empty;
        if (timings.TryGetProperty("fetchMs", out var f) && f.TryGetInt32(out var fetchMs))
        {
            var transfer = timings.TryGetProperty("transferBytes", out var tr) && tr.TryGetInt32(out var bytes)
                ? bytes
                : -1;
            var html = timings.TryGetProperty("htmlMs", out var h) && h.TryGetInt32(out var htmlMs) ? htmlMs : -1;
            var request = timings.TryGetProperty("requestStartMs", out var rq) && rq.TryGetInt32(out var requestMs)
                ? requestMs
                : -1;
            detail = $", html {html}ms, request {request}ms, fetch {fetchMs}ms"
                + $", compile+run {(script >= 0 ? script - fetchMs : -1)}ms"
                + $", transfer {(transfer == 0 ? "cache" : $"{transfer} bytes")}";
        }

        return $" in {total}ms (bundle {script}ms, editor {create}ms{detail}{(build is null ? string.Empty : $", build {build}")})";
    }

    private void Post(string json)
    {
        try
        {
            // Traffic out. The type sits at the front of every payload this side writes, so the
            // first stretch of the text names it without parsing.
            if (Log.VerboseEnabled)
            {
                var head = json.Length <= 48 ? json : json[..48];
                Log.Verbose($"page <- {head}…, {json.Length} chars");
            }

            _browser?.PostMessage(json);
        }
        catch (Exception ex)
        {
            Log.Error("editor surface: a message could not be sent", ex);
        }
    }

    public void Dispose()
    {
        _browser?.Dispose();
        _browser = null;

        _overlay?.Dispose();
        _overlay = null;

        _docs.Clear();
        _activeKey = null;
        _loaded = false;
        _pending.Clear();
        _pendingOrder.Clear();
    }
}
