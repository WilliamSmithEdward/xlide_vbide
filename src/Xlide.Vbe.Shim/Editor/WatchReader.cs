using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Reads what the editor's own Watches window holds: each watch is a list item whose name is
/// the row's four columns run together. The editor exposes no debugger object - no watches, no
/// values - so its own window is the only place this information surfaces.
///
/// The window it reads is a GHOST PALETTE, prepared exactly the way the Locals one is (see
/// AddInSession.PrepareLocalsGhost and lesson 29): floated through the object model, made
/// layered at alpha zero, and parked off the virtual screen. The editor feeds any window with
/// a paintable surface, and a layered window always has one, so the ghost is fed through every
/// break and step while being impossible to see.
///
/// The connection, the guarded walk, the fault backoff and the disposal live on
/// GhostWindowReader, shared with the Locals reader; this class keeps what is particular to
/// Watches - only list items matter, and a row splits into four columns.
/// </summary>
internal sealed class WatchReader : GhostWindowReader
{
    /// <summary>One row of the Watches window, its four columns separated.</summary>
    public readonly record struct WatchRow(string Expression, string Value, string Type, string Context);

    private WatchReader(nint window)
        : base(window, "watch")
    {
    }

    /// <summary>
    /// Prepares a reader for the window, or null when the accessibility interface cannot reach
    /// it. A host without it loses the panel; the police pass hides the native window.
    /// </summary>
    public static WatchReader? Create(nint window)
    {
        if (window == 0)
        {
            return null;
        }

        var reader = new WatchReader(window);
        return reader.Connect() ? reader : null;
    }

    /// <summary>The window's current rows, or null when it cannot be read.</summary>
    public IReadOnlyList<WatchRow>? Read()
    {
        List<WatchRow>? rows = null;

        var walked = TryWalk(
            type => type == UiAutomationIds.ListItemControl,
            (_, text) =>
            {
                if (ParseRow(text) is { } row)
                {
                    (rows ??= []).Add(row);
                }
            });

        return walked ? rows is null ? [] : rows : null;
    }

    /// <summary>
    /// One row's columns, split back apart.
    ///
    /// A REAL watch row's accessible name carries no "Expression" header word - measured
    /// 2026-08-05 against a watch added through the native dialog: " counter Value
    /// &lt;Out of context&gt; Type Empty Context BreakProbe.BreakHere " - the leading space is
    /// the watch-type icon column's empty cell. The expression is everything before the FIRST
    /// " Value " (a watch expression can contain spaces: "counter &gt; 40"); the value can be
    /// anything, including the header words inside a string literal, so it takes everything
    /// between that first " Value " and the LAST " Type " that still sits before the LAST
    /// " Context ". A name with the headers missing or out of order parses as nothing.
    /// </summary>
    internal static WatchRow? ParseRow(string text)
    {
        const string valueHeader = " Value ";
        const string typeHeader = " Type ";
        const string contextHeader = " Context ";

        var valueAt = text.IndexOf(valueHeader, StringComparison.Ordinal);
        if (valueAt < 0)
        {
            return null;
        }

        var expression = text[..valueAt].Trim();
        if (expression.Length == 0)
        {
            return null;
        }

        var contextAt = text.LastIndexOf(contextHeader, StringComparison.Ordinal);
        if (contextAt <= valueAt)
        {
            return null;
        }

        var typeAt = text.LastIndexOf(typeHeader, contextAt, StringComparison.Ordinal);
        if (typeAt <= valueAt)
        {
            return null;
        }

        var value = text[(valueAt + valueHeader.Length)..typeAt].TrimEnd();
        var type = text[(typeAt + typeHeader.Length)..contextAt].TrimEnd();
        var context = text[(contextAt + contextHeader.Length)..].TrimEnd();

        return new WatchRow(expression, value, type, context);
    }
}
