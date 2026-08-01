using System.Text.Json;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.WebView;

namespace Xlide.Vbe.Shim.UI;

/// <summary>
/// Carries messages between the add-in and whatever panel is currently shown.
///
/// The two ends have unrelated lifetimes. The panel is created by the editor when it opens a tool
/// window and destroyed when it closes one; analysis runs whether a panel exists or not. Coupling
/// them directly would mean the analysis path holding a reference to a window that may already be
/// gone, so instead the panel registers itself here.
///
/// Registration is not readiness. A browser surface exists the moment it is asked for and cannot
/// receive anything until its content surface has been created and its document has loaded, which
/// is two asynchronous stages later. Analysis routinely finishes inside that window: on this
/// machine the engine reports in 30ms and the panel document loads in about 200ms, so treating
/// registration as readiness dropped every finding of the first run and left the panel reading
/// "Problems none" while the engine had already proved a defect.
///
/// So the panel says when it is ready, and the last findings are kept and replayed to it. They are
/// kept rather than consumed, because a panel that is closed and reopened has to come up populated
/// too, and nothing re-runs analysis on its behalf.
/// </summary>
internal static class PanelBus
{
    private static readonly Lock Gate = new();
    private static WebView2Surface? _surface;

    /// <summary>Whether the registered panel's document has announced itself.</summary>
    private static bool _ready;

    /// <summary>The last findings sent, kept for replay. Null until analysis has produced any.</summary>
    private static string? _latest;

    /// <summary>Raised when the user asks the panel to go to a finding.</summary>
    public static Action<string, int, int>? NavigateRequested { get; set; }

    /// <summary>Registers the surface a panel is rendering into.</summary>
    public static void Attach(WebView2Surface surface)
    {
        ArgumentNullException.ThrowIfNull(surface);

        lock (Gate)
        {
            _surface = surface;

            // A newly attached panel has not loaded its document yet, whatever the last one did.
            _ready = false;
        }

        surface.MessageReceived = OnPageMessage;
    }

    /// <summary>Unregisters a surface that is being torn down.</summary>
    public static void Detach(WebView2Surface surface)
    {
        lock (Gate)
        {
            if (ReferenceEquals(_surface, surface))
            {
                _surface = null;
                _ready = false;
            }
        }
    }

    /// <summary>Sends findings to the panel, or holds them until one is ready for them.</summary>
    public static void PublishFindings(IReadOnlyList<Finding> findings)
    {
        try
        {
            var payload = JsonSerializer.Serialize(
                new FindingsMessage("findings", [.. findings.Select(ToPayload)]),
                PanelJsonContext.Default.FindingsMessage);

            WebView2Surface? target;

            lock (Gate)
            {
                _latest = payload;
                target = _ready ? _surface : null;
            }

            target?.PostMessage(payload);
        }
        catch (Exception ex)
        {
            Log.Error("panel: findings could not be sent", ex);
        }
    }

    /// <summary>Handles a message the panel document posted.</summary>
    private static void OnPageMessage(string payload)
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
                    OnPageReady();
                    break;

                case "navigate":
                    OnNavigateRequested(document.RootElement);
                    break;
            }
        }
        catch (JsonException)
        {
            Log.Warn("panel: a message from the page was not valid");
        }
    }

    private static void OnPageReady()
    {
        WebView2Surface? target;
        string? backlog;

        lock (Gate)
        {
            _ready = true;
            target = _surface;
            backlog = _latest;
        }

        Log.Info("panel: ready");

        if (target is not null && backlog is not null)
        {
            target.PostMessage(backlog);
        }
    }

    private static void OnNavigateRequested(JsonElement message)
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

    private static FindingPayload ToPayload(Finding finding) => new(
        finding.Module,
        finding.Code,
        finding.Message,
        finding.Severity,
        finding.StartLine,
        finding.StartColumn);
}
