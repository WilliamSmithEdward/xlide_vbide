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

    /// <summary>Text waiting for the page to become ready. Null once it has been sent.</summary>
    private string? _pending;

    /// <summary>
    /// Squiggles waiting for the page, held for the same reason the text is.
    ///
    /// Analysis finishes long before the page loads: the engine answers in tens of milliseconds and
    /// the page takes hundreds, so the first pass always lands early. Dropping it left the module
    /// showing no defects until something changed the text, which reads as an analyzer that found
    /// nothing rather than a message that arrived too soon.
    /// </summary>
    private EditorMarker[]? _pendingMarkers;

    private bool _loaded;

    private EditorSurface(nint host) => _host = host;

    /// <summary>The window this surface is a child of, which is the editor's document area.</summary>
    public nint Host => _host;

    /// <summary>The component currently shown, or null when nothing is.</summary>
    public string? Module => _module;

    /// <summary>Raised when the surface reports the developer changed the text.</summary>
    public Action<string, string>? TextChanged { get; set; }

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

        // Asked for the editing document, then left alone. Creating a browser is asynchronous in two
        // stages, so mapping the content root and navigating from here would run before the browser
        // exists and quietly do nothing. The surface performs both once it is ready.
        surface._browser = WebView2Surface.Start(
            surface._overlay.Handle,
            surface._overlay.ClientBounds(),
            SurfaceContent.Editor);

        if (surface._browser is null)
        {
            surface._overlay.Dispose();
            return null;
        }

        surface._browser.MessageReceived = surface.OnMessage;

        Log.Info($"editor surface: created, serving from {root}");
        return surface;
    }

    /// <summary>Moves the surface over a pane, or hides it when there is nothing to cover.</summary>
    public void Follow(PixelRect bounds, bool visible) => _overlay?.Place(bounds, visible);

    /// <summary>Shows a module's text.</summary>
    public void Show(string moduleName, string text)
    {
        // Squiggles belong to the module they were computed for. Carrying held ones across a switch
        // would decorate the new module at the old one's positions.
        if (_module != moduleName)
        {
            _pendingMarkers = null;
        }

        _module = moduleName;
        _pending = text;

        // The page loads asynchronously and is almost always still loading when the first module
        // arrives, so what to show is held rather than dropped and sent the moment it is ready.
        // Discarding it left the surface correctly positioned and permanently blank, which reads as
        // a rendering fault rather than a message that arrived too early.
        if (!_loaded)
        {
            return;
        }

        Flush();
    }

    private void Flush()
    {
        if (_module is not null && _pending is not null)
        {
            Post(JsonSerializer.Serialize(
                new LoadDocumentMessage("loadDocument", _module, _pending),
                EditorMessageContext.Default.LoadDocumentMessage));

            _pending = null;
        }

        // After the text, never before. Loading a document resets the model, and markers set
        // against the model being replaced are discarded with it.
        if (_pendingMarkers is { } markers)
        {
            _pendingMarkers = null;
            ShowDiagnostics(markers);
        }
    }

    /// <summary>Replaces the squiggles shown on the module currently displayed.</summary>
    public void ShowDiagnostics(EditorMarker[] markers)
    {
        ArgumentNullException.ThrowIfNull(markers);

        if (!_loaded)
        {
            _pendingMarkers = markers;
            return;
        }

        Post(JsonSerializer.Serialize(
            new SetDiagnosticsMessage("setDiagnostics", markers),
            EditorMessageContext.Default.SetDiagnosticsMessage));
    }

    /// <summary>Scrolls a one-based line into view. Ignored until the page is ready.</summary>
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
                    break;

                case "contentChanged":
                    if (_module is not null
                        && document.RootElement.TryGetProperty("fullText", out var text)
                        && text.GetString() is { } updated)
                    {
                        TextChanged?.Invoke(_module, updated);
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
        _loaded = false;
    }
}
