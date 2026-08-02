using System.Text.Json;
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
    private string? _module;

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

    private string? _text;
    private bool _unwritten;
    private bool _loaded;

    private EditorSurface(nint host) => _host = host;

    /// <summary>The window this surface is a child of, which is the editor's document area.</summary>
    public nint Host => _host;

    /// <summary>The component currently shown, or null when nothing is.</summary>
    public string? Module => _module;

    /// <summary>Raised when the surface reports the developer changed the text.</summary>
    public Action<string, string>? TextChanged { get; set; }

    /// <summary>Raised when the developer picks a module from the tab strip.</summary>
    public Action<string>? ModuleRequested { get; set; }

    /// <summary>Raised when the developer picks a finding and wants to be taken to it.</summary>
    public Action<string, int, int>? NavigateRequested { get; set; }

    /// <summary>Raised when the developer chooses a command the editor owns.</summary>
    public Action<string>? CommandRequested { get; set; }

    /// <summary>Raised when the developer enters a line in the Immediate panel.</summary>
    public Action<string>? EvaluateRequested { get; set; }

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

    /// <summary>Raised when the developer closes a module's tab.</summary>
    public Action<string>? ModuleCloseRequested { get; set; }

    /// <summary>Raised when the developer asks for a new component: 1 module, 2 class, 3 form.</summary>
    public Action<int>? ComponentInsertRequested { get; set; }

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

        surface._browser.MessageReceived = surface.OnMessage;
        surface._browser.AcceleratorPressed = key => surface.KeyPressed?.Invoke(key) ?? false;

        Log.Info($"editor surface: created, serving from {root}");
        return surface;
    }

    /// <summary>Moves the surface over a pane, or hides it when there is nothing to cover.</summary>
    public void Follow(PixelRect bounds, bool visible) => _overlay?.Place(bounds, visible);

    /// <summary>
    /// Puts keyboard focus back on the surface. Activating a native pane takes it, and a surface
    /// that does not have focus hears no keys: the shortcut that switched module would work once
    /// and then fall silent.
    /// </summary>
    public void Focus() => _browser?.Focus();

    /// <summary>Shows a module's text.</summary>
    public void Show(string moduleName, string text)
    {
        // Squiggles belong to the module they were computed for. Carrying a held set across a
        // switch would decorate the new module at the old one's positions.
        if (_module != moduleName)
        {
            Drop("setDiagnostics");
        }

        _module = moduleName;
        _text = text;

        // The text just arrived from the module, so there is nothing of the developer's to write.
        _unwritten = false;
        _overlay?.StopWriteTimer();

        Send("loadDocument", JsonSerializer.Serialize(
            new LoadDocumentMessage("loadDocument", moduleName, text),
            EditorMessageContext.Default.LoadDocumentMessage));
    }

    /// <summary>
    /// Shows nothing: every pane is closed, and the workspace says so instead of showing the last
    /// module as if it were still open.
    /// </summary>
    public void Clear()
    {
        _module = null;
        _text = null;
        _unwritten = false;
        _overlay?.StopWriteTimer();
        Drop("setDiagnostics");

        // The same kind as loading, deliberately: whichever of the two was said last is the truth
        // a page that has not booted yet should wake up to.
        Send("loadDocument", JsonSerializer.Serialize(
            new ClearDocumentMessage("clearDocument"),
            EditorMessageContext.Default.ClearDocumentMessage));
    }

    /// <summary>
    /// Adopts the editor's version of the module without disturbing what the developer is doing.
    ///
    /// The editor is the text of record and it rewrites what it is given: it respells keywords and
    /// normalises spacing as it takes a module in. So the moment after an edit is written, the two
    /// disagree, and every later comparison would see a difference that is not the developer's.
    /// Adopting its version closes that immediately, at the one moment the difference is known to
    /// be the editor's doing and not an edit in flight.
    /// </summary>
    public void Sync(string moduleName, string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (_module != moduleName)
        {
            return;
        }

        _text = text;

        if (_loaded)
        {
            Post(JsonSerializer.Serialize(
                new SyncDocumentMessage("syncDocument", moduleName, text),
                EditorMessageContext.Default.SyncDocumentMessage));
        }
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

    /// <summary>Shows or withdraws the surface's own menu bar.</summary>
    public void SetChrome(bool menuBar)
    {
        Send("setChrome", JsonSerializer.Serialize(
            new SetChromeMessage("setChrome", menuBar),
            EditorMessageContext.Default.SetChromeMessage));
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

    /// <summary>True when the developer has typed something that has not reached the module.</summary>
    public bool HasUnwrittenEdits => _unwritten;

    /// <summary>The surface's copy of the module text, or null when it is showing nothing.</summary>
    public string? Text => _text;

    /// <summary>Replaces the squiggles shown on the module currently displayed.</summary>
    public void ShowDiagnostics(EditorMarker[] markers)
    {
        ArgumentNullException.ThrowIfNull(markers);

        Send("setDiagnostics", JsonSerializer.Serialize(
            new SetDiagnosticsMessage("setDiagnostics", markers),
            EditorMessageContext.Default.SetDiagnosticsMessage));
    }

    /// <summary>Replaces the tab strip: every module the editor has open, and which one is shown.</summary>
    public void ShowModules(string[] modules, string? active)
    {
        ArgumentNullException.ThrowIfNull(modules);

        Send("setModules", JsonSerializer.Serialize(
            new SetModulesMessage("setModules", modules, active),
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
        if (_text is null || line < 1)
        {
            return null;
        }

        var start = 0;
        for (var current = 1; ; current++)
        {
            var end = _text.IndexOf('\n', start);
            if (current == line)
            {
                var text = end < 0 ? _text[start..] : _text[start..end];
                return text.TrimEnd('\r');
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

        if (!_unwritten || _module is null || _text is null)
        {
            return;
        }

        _unwritten = false;
        TextChanged?.Invoke(_module, _text);
    }

    /// <summary>Sends a message, or holds it until the page is ready for it.</summary>
    private void Send(string kind, string json)
    {
        if (_loaded)
        {
            Post(json);
            return;
        }

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

            switch (type.GetString())
            {
                case "ready":
                    _loaded = true;
                    Log.Info($"editor surface: ready{DescribeTimings(document.RootElement)}");
                    Flush();
                    Ready?.Invoke();
                    break;

                case "selectionChanged":
                    if (document.RootElement.TryGetProperty("startLine", out var line)
                        && line.TryGetInt32(out var caretLine))
                    {
                        CaretLine = caretLine;
                        CaretColumn = document.RootElement.TryGetProperty("startColumn", out var column)
                            && column.TryGetInt32(out var caretColumn)
                            ? caretColumn
                            : 1;
                    }

                    break;

                case "contentChanged":
                    if (_module is not null
                        && document.RootElement.TryGetProperty("fullText", out var text)
                        && text.GetString() is { } updated)
                    {
                        // Held, not written. Writing a module replaces every line of it, which
                        // resets the project, so doing it per keystroke would reset the project on
                        // every keystroke. The edit is written once the developer stops typing,
                        // and immediately before anything that has to see it.
                        _text = updated;
                        _unwritten = true;
                        _overlay?.StartWriteTimer(WriteDelayMilliseconds);
                    }

                    break;

                case "activateModule":
                    if (document.RootElement.TryGetProperty("moduleName", out var requested)
                        && requested.GetString() is { Length: > 0 } name)
                    {
                        Log.Info($"surface: activate {name} requested");
                        ModuleRequested?.Invoke(name);
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
                        ModuleCloseRequested?.Invoke(closingName);
                    }

                    break;

                case "insertComponent":
                    if (document.RootElement.TryGetProperty("kind", out var componentKind)
                        && componentKind.TryGetInt32(out var insertKind))
                    {
                        ComponentInsertRequested?.Invoke(insertKind);
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

        NavigateRequested?.Invoke(component, line, column);
    }

    /// <summary>
    /// Renders what the page reported about its own start-up, so the cost of putting a surface over
    /// a pane is a measured number in the log rather than an impression.
    /// </summary>
    private static string DescribeTimings(JsonElement message)
    {
        if (!message.TryGetProperty("timings", out var timings) || timings.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        var script = timings.TryGetProperty("scriptMs", out var s) && s.TryGetInt32(out var scriptMs) ? scriptMs : -1;
        var create = timings.TryGetProperty("createMs", out var c) && c.TryGetInt32(out var createMs) ? createMs : -1;
        var total = timings.TryGetProperty("totalMs", out var t) && t.TryGetInt32(out var totalMs) ? totalMs : -1;

        return $" in {total}ms (bundle {script}ms, editor {create}ms)";
    }

    private void Post(string json)
    {
        try
        {
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

        _module = null;
        _text = null;
        _loaded = false;
        _pending.Clear();
        _pendingOrder.Clear();
    }
}
