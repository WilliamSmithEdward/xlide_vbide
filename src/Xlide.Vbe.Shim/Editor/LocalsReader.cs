using Xlide.Vbe.Shim.Com;
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
    public readonly record struct LocalRow(
        string Expression, string Value, string Type,
        int Depth = 0, bool Expandable = false, bool Expanded = false);

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
        HashSet<string> seen = new(StringComparer.Ordinal);

        using var walker = RawWalker();
        var walked = TryWalk(
            type => type is UiAutomationIds.ListItemControl
                or UiAutomationIds.EditControl
                or UiAutomationIds.PaneControl,
            (type, text, element) =>
            {
                if (type == UiAutomationIds.ListItemControl)
                {
                    if (ParseRow(text) is { } row)
                    {
                        var depth = RowDepth(element, walker?.Target);
                        using var item = ComHandle<IUIAutomationElement>.Borrow(element);
                        var bounds = item is null ? null : ReadBounds(item.Target);
                        // VBE's provider can return the same child twice at the same rectangle
                        // after an object is expanded. It is one visible row, not two variables.
                        if (bounds is null || seen.Add(RowKey(text, depth, bounds.Value)))
                        {
                            (rows ??= []).Add(row with { Depth = depth });
                        }
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

        if (!walked)
        {
            return null;
        }

        if (rows is not null)
        {
            for (var i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                var expanded = i + 1 < rows.Count && rows[i + 1].Depth > row.Depth;
                rows[i] = row with { Expanded = expanded, Expandable = expanded || row.Value.Length == 0 };
            }
        }

        return new LocalsSnapshot(context, rows is null ? [] : rows);
    }

    private ComHandle<IUIAutomationTreeWalker>? RawWalker()
    {
        try
        {
            return Automation is { } automation && automation.GetRawViewWalker(out var pointer) >= 0 && pointer != 0
                ? ComHandle<IUIAutomationTreeWalker>.Own(pointer) : null;
        }
        catch
        {
            // The flat reading still works when an accessibility provider refuses a walker.
            return null;
        }
    }

    private static int RowDepth(nint element, IUIAutomationTreeWalker? walker)
    {
        if (walker is null) return 0;
        var depth = 0;
        var current = ComHandle<IUIAutomationElement>.Borrow(element);
        try
        {
            for (var i = 0; i < 16; i++)
            {
                if (current is null || walker.GetParentElement(current.Pointer, out var parentPointer) < 0 || parentPointer == 0) break;
                var parent = ComHandle<IUIAutomationElement>.Own(parentPointer);
                if (parent is null || parent.Target.GetCurrentPropertyValue(UiAutomationIds.ControlTypeProperty, out var kind) < 0
                    || kind.AsInt32() != UiAutomationIds.ListItemControl)
                {
                    parent?.Dispose();
                    break;
                }
                depth++;
                current.Dispose();
                current = parent;
            }
        }
        finally
        {
            current?.Dispose();
        }
        return depth;
    }

    private static (double X, double Y, double Width, double Height)? ReadBounds(IUIAutomationElement item) =>
        item.GetCurrentPropertyValue(UiAutomationIds.BoundingRectangleProperty, out var bounds) >= 0
            ? bounds.TakeRectangle() : null;

    private static string RowKey(string text, int depth, (double X, double Y, double Width, double Height) bounds) =>
        $"{depth}\u001f{bounds.X}\u001f{bounds.Y}\u001f{bounds.Width}\u001f{bounds.Height}\u001f{text}";

    /// <summary>Ask VBE's own Locals control to expand/collapse a row at the break.</summary>
    public void Toggle(int rowIndex, string expression, int depth)
    {
        if (rowIndex < 0 || rowIndex > 10000) return;
        var current = -1;
        HashSet<string> seen = new(StringComparer.Ordinal);
        using var walker = RawWalker();
        nint listWindow = 0;
        (double X, double Y, double Width, double Height)? listRect = null;
        (double X, double Y, double Width, double Height)? rowRect = null;
        _ = TryWalk(type => type is UiAutomationIds.ListControl or UiAutomationIds.ListItemControl,
            (type, text, element) =>
            {
                using var item = ComHandle<IUIAutomationElement>.Borrow(element);
                if (item is null) return;
                if (type == UiAutomationIds.ListControl)
                {
                    if (item.Target.GetCurrentPropertyValue(UiAutomationIds.NativeWindowHandleProperty, out var handle) >= 0)
                        listWindow = handle.AsInt32();
                    listRect = ReadBounds(item.Target);
                }
                else if (ParseRow(text) is { } row)
                {
                    var bounds = ReadBounds(item.Target);
                    var rowDepth = RowDepth(element, walker?.Target);
                    if (bounds is not null && !seen.Add(RowKey(text, rowDepth, bounds.Value))) return;
                    if (++current == rowIndex && rowDepth == depth
                        && string.Equals(row.Expression, expression, StringComparison.Ordinal))
                        rowRect = bounds;
                }
            });

        if (listWindow == 0 || listRect is null || rowRect is null) return;
        var x = (int)(rowRect.Value.X - listRect.Value.X) + 6 + Math.Clamp(depth, 0, 10) * 14;
        var y = (int)(rowRect.Value.Y - listRect.Value.Y + rowRect.Value.Height / 2);
        if (x < 0 || y < 0 || x >= listRect.Value.Width || y >= listRect.Value.Height) return;
        var point = (nint)((y << 16) | (x & 0xffff));
        Win32.PostMessage(listWindow, 0x0201, 1, point);
        Win32.PostMessage(listWindow, 0x0202, 0, point);
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
    /// "Expression counter Value 42 Type Long". The expression can also be a native member label
    /// such as "Item 1"; the value can be anything, including
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
        if (expression.Length == 0)
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
