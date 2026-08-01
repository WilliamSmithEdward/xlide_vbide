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
/// gone, so instead the panel registers itself here and messages sent while none is registered are
/// simply dropped.
/// </summary>
internal static class PanelBus
{
    private static readonly Lock Gate = new();
    private static WebView2Surface? _surface;
    private static string? _pending;

    /// <summary>Registers the surface a panel is rendering into.</summary>
    public static void Attach(WebView2Surface surface)
    {
        string? backlog;

        lock (Gate)
        {
            _surface = surface;
            backlog = _pending;
            _pending = null;
        }

        // Analysis usually finishes before the panel has loaded, so the last message sent while
        // nothing was listening is replayed. Without it the panel comes up empty and stays empty
        // until something changes, which reads as the product not working.
        if (backlog is not null)
        {
            surface.PostMessage(backlog);
        }
    }

    /// <summary>Unregisters a surface that is being torn down.</summary>
    public static void Detach(WebView2Surface surface)
    {
        lock (Gate)
        {
            if (ReferenceEquals(_surface, surface))
            {
                _surface = null;
            }
        }
    }

    /// <summary>Sends findings to the panel, or holds them until one appears.</summary>
    public static void PublishFindings(IReadOnlyList<Finding> findings)
    {
        try
        {
            var payload = JsonSerializer.Serialize(
                new FindingsMessage("findings", [.. findings.Select(ToPayload)]),
                PanelJsonContext.Default.FindingsMessage);

            WebView2Surface? surface;

            lock (Gate)
            {
                surface = _surface;
                if (surface is null)
                {
                    _pending = payload;
                }
            }

            surface?.PostMessage(payload);
        }
        catch (Exception ex)
        {
            Log.Error("panel: findings could not be sent", ex);
        }
    }

    private static FindingPayload ToPayload(Finding finding) => new(
        finding.Module,
        finding.Code,
        finding.Message,
        finding.Severity,
        finding.StartLine,
        finding.StartColumn);
}
