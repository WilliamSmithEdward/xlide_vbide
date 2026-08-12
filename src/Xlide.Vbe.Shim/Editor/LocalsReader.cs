using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Reads what the editor's own Locals window holds: each variable is a list item whose name is
/// the row's columns run together, and an edit control names the broken procedure. The editor
/// exposes no debugger object - no stack, no frames, no variables - so its own window is the
/// only place this information surfaces.
///
/// The window it reads is the GHOST PALETTE (see AddInSession.PrepareLocalsGhost): the native
/// Locals window floated through the object model, made layered at alpha zero, and parked off
/// the virtual screen. The editor only feeds a window with a paintable surface (lesson 25 -
/// hidden never fills, covered fills unreliably and never on a step), and a layered window
/// renders into its own surface regardless of position or occlusion, so the ghost is fed
/// faithfully through every break and step while being impossible to see. Probed 2026-08-04:
/// counter tracked 1 through 4 across steps at alpha 0, off screen.
///
/// The connection, the guarded walk, the fault backoff and the disposal live on
/// GhostWindowReader, shared with the Watches reader; this class keeps what is particular to
/// Locals - which control types matter, and how a row's text splits back into columns.
/// </summary>
internal sealed class LocalsReader : GhostWindowReader
{
    /// <summary>One row of the Locals window, its three columns separated.</summary>
    public readonly record struct LocalRow(string Expression, string Value, string Type);

    /// <summary>A reading: which procedure is broken, and the variables in scope.</summary>
    public sealed record LocalsSnapshot(string? Context, IReadOnlyList<LocalRow> Rows);

    private LocalsReader(nint window)
        : base(window, "locals")
    {
    }

    /// <summary>
    /// Prepares a reader for the window, or null when the accessibility interface cannot reach
    /// it. A host without it loses the panel; the police pass hides the native window.
    /// </summary>
    public static LocalsReader? Create(nint window)
    {
        if (window == 0)
        {
            return null;
        }

        var reader = new LocalsReader(window);
        return reader.Connect() ? reader : null;
    }

    /// <summary>
    /// The window's current content, or null when it cannot be read.
    ///
    /// The placeholder rows the window shows outside a break - "&lt;No Variables&gt;" - do not
    /// parse as rows, so an idle window reads as an empty snapshot rather than as a variable
    /// with an angle-bracketed name.
    /// </summary>
    public LocalsSnapshot? Read()
    {
        string? context = null;
        List<LocalRow>? rows = null;

        var walked = TryWalk(
            type => type is UiAutomationIds.ListItemControl
                or UiAutomationIds.EditControl
                or UiAutomationIds.PaneControl,
            (type, text) =>
            {
                if (type == UiAutomationIds.ListItemControl)
                {
                    if (ParseRow(text) is { } row)
                    {
                        (rows ??= []).Add(row);
                    }
                }
                else if (type == UiAutomationIds.EditControl && text.Length > 0)
                {
                    context = text;
                }
                else if (context is null && IsContextText(text))
                {
                    // The context box reaches the accessibility tree as a bare pane, not an
                    // edit (measured 2026-08-05: pane named "VBAProject.BreakProbe.BreakHere"
                    // beside panes named "" and "..."), so a pane may carry the context but
                    // never outranks a real edit. An empty edit does not count as a context:
                    // the panel hides its strip on null, and an empty non-null string would
                    // show a blank strip instead.
                    context = text;
                }
            });

        return walked ? new LocalsSnapshot(context, rows is null ? [] : rows) : null;
    }

    /// <summary>
    /// Whether a pane's text reads as the broken procedure's path. The real one is dotted
    /// ("VBAProject.Module1.Test"); its neighbours are empty or the call-stack button's "...".
    /// </summary>
    internal static bool IsContextText(string text) =>
        text.Contains('.', StringComparison.Ordinal) && text.Any(char.IsLetter);

    /// <summary>
    /// One row's columns, split back apart.
    ///
    /// The row's accessible name is its columns run together with the header words in between:
    /// "Expression counter Value 42 Type Long". The expression is a single token - an identifier
    /// or an indexed path, never containing a space - while the value can be anything, including
    /// the words Value or Type inside a string literal, so the value takes everything between
    /// the first " Value " and the LAST " Type ". The placeholder the idle window shows has no
    /// expression token and parses as nothing.
    /// </summary>
    internal static LocalRow? ParseRow(string text)
    {
        const string expressionHeader = "Expression ";
        const string valueHeader = " Value ";
        const string typeHeader = " Type ";

        if (!text.StartsWith(expressionHeader, StringComparison.Ordinal))
        {
            return null;
        }

        var valueAt = text.IndexOf(valueHeader, expressionHeader.Length, StringComparison.Ordinal);
        if (valueAt < 0)
        {
            return null;
        }

        var expression = text[expressionHeader.Length..valueAt];
        if (expression.Length == 0 || expression.Contains(' '))
        {
            return null;
        }

        var typeAt = text.LastIndexOf(typeHeader, StringComparison.Ordinal);
        if (typeAt <= valueAt)
        {
            return null;
        }

        var value = text[(valueAt + valueHeader.Length)..typeAt];
        var type = text[(typeAt + typeHeader.Length)..].TrimEnd();

        return new LocalRow(expression, value.TrimEnd(), type);
    }
}
