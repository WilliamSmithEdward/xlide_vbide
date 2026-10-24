using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
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
        if (ReadAll(tailDue: TailDue) is not { } reading)
        {
            return null;
        }

        List<LocalRow> rows;
        using (reading)
        {
            rows = reading.Rows.Select(row => row.Row).ToList();
            if (reading.Truncated && reading.TailRead)
            {
                _tail = rows.Skip(reading.Reachable).ToList();
                _tailFor = reading.Structure;
                _tailAt = Environment.TickCount64;
            }
            else if (reading.Truncated)
            {
                rows.AddRange(_tail ?? []);
            }
            else
            {
                _tail = null;
                _tailFor = null;
            }
        }

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var expanded = i + 1 < rows.Count && rows[i + 1].Depth > row.Depth;
            rows[i] = row with { Expanded = expanded, Expandable = expanded || row.Value.Length == 0 };
        }

        return new LocalsSnapshot(reading.Context, rows);
    }

    /// <summary>
    /// The rows recovered below an expanded block at the last recovery, and the structure of
    /// the reachable rows they were recovered against. A step that changes only a value in the
    /// tail leaves that structure as it was, so the tail is also read again once a second.
    /// </summary>
    private List<LocalRow>? _tail;
    private string? _tailFor;
    private long _tailAt;
    private const int TailRefreshMs = 1000;

    private Func<string, bool> TailDue => structure =>
        _tail is null || structure != _tailFor || Environment.TickCount64 - _tailAt > TailRefreshMs;

    /// <summary>More rows below an expanded block than any procedure declares.</summary>
    private const int MaxTailRows = 256;

    private ComHandle<IUIAutomationTreeWalker>? _walker;

    /// <summary>A row as a walk found it: parsed, its identity and expand state, its live element.</summary>
    private sealed record WalkedRow(LocalRow Row, string Text, string? Identity, int State, ComHandle<IUIAutomationElement>? Element);

    /// <summary>One reading's rows, context and list element; owns every live element it holds.</summary>
    private sealed class RowsReading : IDisposable
    {
        public string? Context;
        public List<WalkedRow> Rows = [];
        public ComHandle<IUIAutomationElement>? List;
        public nint ListWindow;

        /// <summary>The rows the list itself reaches; any after them were recovered by selection.</summary>
        public int Reachable;

        /// <summary>The reachable rows end inside an expanded block, so rows may be missing after them.</summary>
        public bool Truncated;

        /// <summary>Whether this reading recovered the rows below the block itself.</summary>
        public bool TailRead;

        public string Structure = string.Empty;

        /// <summary>The identities of the rows read so far, so a recovery stops where it meets one.</summary>
        public HashSet<string> Seen = new(StringComparer.Ordinal);

        /// <summary>
        /// The rows that enclose the last row read, outermost first, each with its place among
        /// its parent's rows; and how many rows each open level has had so far.
        /// </summary>
        public List<(int Row, int Place)> Path = [];
        public List<int> Counts = [];

        public void Dispose()
        {
            foreach (var row in Rows) row.Element?.Dispose();
            Rows.Clear();
            List?.Dispose();
            List = null;
        }
    }

    /// <summary>
    /// Every row in display order with its depth: the rows the list reaches, and when they end
    /// inside an expanded block and `tailDue` says so, the rows recovered below it.
    ///
    /// The provider's own child lists cannot give the nesting (mapped 2026-10-17): an expanded
    /// row lists its members, or lists nothing at all once it has been collapsed and expanded
    /// again, and a collapsed row may list the row after it. Its parent links name the wrong
    /// row too, below an expanded block. What holds is the list's chain, which runs through
    /// every row in display order and stops at the first fall in depth. Inside a chain, then,
    /// a row is one deeper than an expanded row above it and level with any other.
    /// </summary>
    private RowsReading? ReadAll(Func<string, bool> tailDue)
    {
        var reading = new RowsReading();
        var walked = TryWalk(
            type => type is UiAutomationIds.ListItemControl
                or UiAutomationIds.ListControl
                or UiAutomationIds.EditControl
                or UiAutomationIds.PaneControl,
            element =>
            {
                var text = element.Text;
                if (element.ControlType == UiAutomationIds.ListItemControl)
                {
                    if (ParseRow(text) is { } row)
                    {
                        reading.Rows.Add(new WalkedRow(row, text, element.Identity, element.State,
                            ComHandle<IUIAutomationElement>.Borrow(element.Element)));
                    }
                }
                else if (element.ControlType == UiAutomationIds.ListControl)
                {
                    if (reading.List is null)
                    {
                        reading.List = ComHandle<IUIAutomationElement>.Borrow(element.Element);
                        reading.ListWindow = element.Window;
                    }
                }
                else if (element.ControlType == UiAutomationIds.EditControl && text.Length > 0)
                {
                    reading.Context = text;
                }
                else if (reading.Context is null && IsContextText(text))
                {
                    // The context box reaches the accessibility tree as a bare pane, not an
                    // edit (measured 2026-08-05: pane named "VBAProject.BreakProbe.BreakHere"
                    // beside panes named "" and "..."), so a pane may carry the context but
                    // never outranks a real edit. An empty edit does not count as a context:
                    // the panel hides its strip on null, and an empty non-null string would
                    // show a blank strip instead.
                    reading.Context = text;
                }
            },
            live: true,
            descendRows: false);

        if (!walked)
        {
            reading.Dispose();
            return null;
        }

        try
        {
            reading.Structure = string.Join('\n', reading.Rows.Select(row => $"{row.Identity} {row.State} {row.Row.Expression}"));
            for (var i = 0; i < reading.Rows.Count; i++)
            {
                Place(reading, i, i == 0 ? 0 : ChainDepth(reading.Rows[i - 1]));
            }

            reading.Reachable = reading.Rows.Count;
            reading.Truncated = reading.Rows.Count > 0 && reading.Rows[^1].Row.Depth > 0;
            if (reading.Truncated && tailDue(reading.Structure))
            {
                RecoverTail(reading);
                reading.TailRead = true;
            }
        }
        catch (Exception ex)
        {
            Log.Verbose($"locals: the rows' nesting could not be read: {ex.Message}");
        }

        return reading;
    }

    /// <summary>The depth of the row after this one in the same chain.</summary>
    private static int ChainDepth(WalkedRow previous) =>
        previous.Row.Depth + (previous.State == UiAutomationIds.Expanded ? 1 : 0);

    /// <summary>
    /// Sets a row's depth and moves the reading's path to it: the rows enclosing it, each with
    /// its place among its parent's rows.
    /// </summary>
    private static void Place(RowsReading reading, int index, int depth)
    {
        var rows = reading.Rows;
        depth = Math.Clamp(depth, 0, reading.Path.Count);
        rows[index] = rows[index] with { Row = rows[index].Row with { Depth = depth } };
        if (rows[index].Identity is { } identity) reading.Seen.Add(identity);

        // A row one deeper than the path opens a new block, whose count starts again.
        var keep = depth < reading.Path.Count ? depth + 1 : depth;
        if (reading.Path.Count > depth) reading.Path.RemoveRange(depth, reading.Path.Count - depth);
        if (reading.Counts.Count > keep) reading.Counts.RemoveRange(keep, reading.Counts.Count - keep);
        while (reading.Counts.Count <= depth) reading.Counts.Add(0);
        reading.Path.Add((index, reading.Counts[depth]++));
    }

    /// <summary>
    /// The depth of a row recovered below an expanded block, the first of its chain.
    ///
    /// Its parent link does not name its parent: the provider builds the parent from the
    /// parent's place among ITS parent's rows and returns whichever row is shown at that place
    /// (measured 2026-10-17: c's members, below an expanded Workbook, named the row at 4.5 - c
    /// is the sixth local - which was one of the Workbook's members). That place is enough. The
    /// parent is one of the rows enclosing the row above, and the one whose place matches is
    /// it. A parent that is the list itself makes the row top-level.
    /// </summary>
    private int RecoveredDepth(RowsReading reading, ComHandle<IUIAutomationElement> element)
    {
        if (Walker() is not { } walker
            || walker.Target.GetParentElementBuildCache(element.Pointer, ElementRequest, out var parentPointer) < 0
            || parentPointer == 0)
        {
            return Math.Max(0, reading.Path.Count - 2);
        }

        using var parent = ComHandle<IUIAutomationElement>.Own(parentPointer);
        if (parent is null) return Math.Max(0, reading.Path.Count - 2);
        var controlType = parent.Target.GetCachedPropertyValue(UiAutomationIds.ControlTypeProperty, out var kind) >= 0
            ? kind.AsInt32() : 0;
        if (controlType != UiAutomationIds.ListItemControl) return 0;

        var place = parent.Target.GetCachedPropertyValue(UiAutomationIds.RuntimeIdProperty, out var runtimeId) >= 0
            ? DisplayIndex(runtimeId.TakeRuntimeId()) : null;
        // A row that can head a block reads Expanded, or - recovered through the selection,
        // where the provider reports every row as a leaf - is an object, whose value is empty.
        for (var level = reading.Path.Count - 1; level >= 0 && place is not null; level--)
        {
            var (row, at) = reading.Path[level];
            var candidate = reading.Rows[row];
            if (at == place && (candidate.State == UiAutomationIds.Expanded || candidate.Row.Value.Length == 0))
            {
                return level + 1;
            }
        }

        // Unmatched, the likeliest answer: one level up from the row above.
        return Math.Max(0, reading.Path.Count - 2);
    }

    /// <summary>The last part of a row's runtime identity: its place on screen, or for a parent link, its place among its parent's rows.</summary>
    private static int? DisplayIndex(string? identity) =>
        identity is not null && identity.LastIndexOf('.') is var dot and >= 0
        && int.TryParse(identity.AsSpan(dot + 1), System.Globalization.NumberStyles.None,
            System.Globalization.CultureInfo.InvariantCulture, out var index)
            ? index : null;

    private ComHandle<IUIAutomationTreeWalker>? Walker()
    {
        if (_walker is null && Automation is { } automation
            && automation.GetRawViewWalker(out var pointer) >= 0 && pointer != 0)
        {
            _walker = ComHandle<IUIAutomationTreeWalker>.Own(pointer);
        }

        return _walker;
    }

    /// <summary>
    /// The rows the provider loses below an expanded block, found by selecting them.
    ///
    /// From the last member of an expanded object nothing links onward, though the native
    /// list paints every row after it. A selected row is reachable all the same, through the
    /// list's selection. So the lost row is selected the way a person would select it: the row
    /// above is selected, which scrolls it into view, the list moves down a line, and a click
    /// lands on the row below, in its Type column, which edits nothing and holds no expander.
    /// The selection then names the lost row, and the rows the provider chains after it come
    /// with it, up to the next fall in depth, where the same step repeats. When a click selects
    /// nothing new, the list has ended. The ghost is invisible, so the selection and the
    /// scrolling are seen by nobody; the scroll position is put back.
    /// </summary>
    private void RecoverTail(RowsReading reading)
    {
        if (reading.List is not { } list || reading.ListWindow == 0 || reading.Rows.Count == 0
            || reading.Rows[^1].Element is not { } last || Walker() is not { } walker)
        {
            return;
        }

        using var selection = Pattern<IUIAutomationSelectionPattern>(list,
            UiAutomationIds.SelectionPatternId, UiAutomationIds.SelectionPattern);
        using var scroll = Pattern<IUIAutomationScrollPattern>(list,
            UiAutomationIds.ScrollPatternId, UiAutomationIds.ScrollPattern);
        if (selection is null || scroll is null || CurrentBounds(list.Target) is not { } box)
        {
            return;
        }

        var restore = scroll.Target.GetCurrentVerticalScrollPercent(out var percent) >= 0 ? percent : (double?)null;
        var anchor = ComHandle<IUIAutomationElement>.Borrow(last.Pointer);
        var recovered = 0;
        try
        {
            while (anchor is not null && recovered < MaxTailRows)
            {
                using (var item = Pattern<IUIAutomationSelectionItemPattern>(anchor,
                    UiAutomationIds.SelectionItemPatternId, UiAutomationIds.SelectionItemPattern))
                {
                    if (item is null || item.Target.Select() < 0) break;
                }

                // Select scrolls a row only as far as the edge; one more line makes room below.
                _ = scroll.Target.Scroll(UiAutomationIds.ScrollNoAmount, UiAutomationIds.ScrollSmallIncrement);
                if (CurrentBounds(anchor.Target) is not { } row || double.IsInfinity(row.Y)) break;

                var x = (int)(row.X - box.X + row.Width * 0.85);
                var y = (int)(row.Y - box.Y + row.Height * 1.5);
                if (x < 0 || y < 0 || x >= box.Width || y >= box.Height) break;
                var point = (nint)((y << 16) | (x & 0xffff));
                Win32.PostMessage(reading.ListWindow, 0x0201, 1, point);
                Win32.PostMessage(reading.ListWindow, 0x0202, 0, point);

                // The question below is answered on the host thread after the click it queued
                // behind, so the selection it reads is the click's.
                using var selected = SelectedRow(selection);
                if (selected is null
                    || selected.Target.BuildUpdatedCache(ElementRequest, out var currentPointer) < 0 || currentPointer == 0)
                {
                    break;
                }

                var before = reading.Rows.Count;
                var current = ComHandle<IUIAutomationElement>.Own(currentPointer);
                while (current is not null && recovered < MaxTailRows)
                {
                    // A chain is taken only while it runs to the next row on screen.
                    var first = reading.Rows.Count == before;
                    if (Describe(current.Pointer, 0) is not { } found || found.Identity is null
                        || reading.Seen.Contains(found.Identity) || ParseRow(found.Text) is not { } parsed
                        || (!first && DisplayIndex(found.Identity) != DisplayIndex(reading.Rows[^1].Identity) + 1))
                    {
                        current.Dispose();
                        break;
                    }

                    var depth = first ? RecoveredDepth(reading, current) : ChainDepth(reading.Rows[^1]);
                    reading.Rows.Add(new WalkedRow(parsed, found.Text, found.Identity, found.State, current));
                    Place(reading, reading.Rows.Count - 1, depth);
                    recovered++;

                    current = walker.Target.GetNextSiblingElementBuildCache(current.Pointer, ElementRequest, out var nextPointer) >= 0
                        ? ComHandle<IUIAutomationElement>.Own(nextPointer) : null;
                }

                if (reading.Rows.Count == before) break;
                anchor.Dispose();
                anchor = ComHandle<IUIAutomationElement>.Borrow(reading.Rows[^1].Element!.Pointer);
            }
        }
        catch (Exception ex)
        {
            Log.Verbose($"locals: recovering the rows below an expanded object stopped: {ex.Message}");
        }
        finally
        {
            anchor?.Dispose();
            if (restore is { } back) _ = scroll.Target.SetScrollPercent(UiAutomationIds.NoScroll, back);
        }

        Log.Verbose($"locals: recovered {recovered} row(s) below an expanded object");
    }

    private static ComHandle<TPattern>? Pattern<TPattern>(ComHandle<IUIAutomationElement> element, int patternId, in Guid interfaceId)
        where TPattern : class =>
        element.Target.GetCurrentPatternAs(patternId, in interfaceId, out var pointer) >= 0 && pointer != 0
            ? ComHandle<TPattern>.Own(pointer) : null;

    private static (double X, double Y, double Width, double Height)? CurrentBounds(IUIAutomationElement element) =>
        element.GetCurrentPropertyValue(UiAutomationIds.BoundingRectangleProperty, out var bounds) >= 0
            ? bounds.TakeRectangle() : null;

    private static ComHandle<IUIAutomationElement>? SelectedRow(ComHandle<IUIAutomationSelectionPattern> selection)
    {
        if (selection.Target.GetCurrentSelection(out var arrayPointer) < 0 || arrayPointer == 0) return null;
        using var array = ComHandle<IUIAutomationElementArray>.Own(arrayPointer);
        return array is not null && array.Target.GetLength(out var count) >= 0 && count > 0
            && array.Target.GetElement(0, out var pointer) >= 0
            ? ComHandle<IUIAutomationElement>.Own(pointer) : null;
    }

    /// <summary>
    /// Ask VBE's own Locals control to expand or collapse a row at the break: a click on the
    /// row's expander, the row first selected into view when the list has it scrolled away.
    /// The row index counts the rows recovered below an expanded block too.
    /// </summary>
    public void Toggle(int rowIndex, string expression, int depth)
    {
        if (rowIndex < 0 || rowIndex > 10000) return;
        using var reading = ReadAll(tailDue: _ => false);
        if (reading is null) return;

        // The tail is recovered only when the row asked for lies in it.
        if (rowIndex >= reading.Rows.Count && reading.Truncated)
        {
            RecoverTail(reading);
        }

        if (rowIndex >= reading.Rows.Count || reading.List is not { } list || reading.ListWindow == 0) return;
        var target = reading.Rows[rowIndex];
        if (target.Element is not { } element || target.Row.Depth != depth
            || !string.Equals(target.Row.Expression, expression, StringComparison.Ordinal)
            || CurrentBounds(list.Target) is not { } box)
        {
            return;
        }

        var row = CurrentBounds(element.Target);
        var before = row;
        if (row is not { } seen || double.IsInfinity(seen.Y) || seen.Y < box.Y || seen.Y + seen.Height > box.Y + box.Height)
        {
            using var item = Pattern<IUIAutomationSelectionItemPattern>(element,
                UiAutomationIds.SelectionItemPatternId, UiAutomationIds.SelectionItemPattern);
            _ = item?.Target.Select();
            row = CurrentBounds(element.Target);
        }

        // The expander's offset is the control's own, laid out at 96 dots per inch and scaled by
        // the window's. Unscaled, the click fell short of every nested row's expander under 130%
        // scaling (an RDP session, 2026-10-23: rows 26px tall, depth 1 clicked at 22 of ~29),
        // while a top-level row's still landed, so only a member far down an object failed.
        var dpi = Win32.GetDpiForWindow(reading.ListWindow);
        var scale = dpi == 0 ? 1.0 : dpi / 96.0;
        var x = row is { } at ? (int)(at.X - box.X + (6 + Math.Clamp(depth, 0, 10) * 14) * scale) : -1;
        var y = row is { } on ? (int)(on.Y - box.Y + on.Height / 2) : -1;
        Log.Verbose($"locals: toggle '{expression}' row {rowIndex}: list {box}, row {before?.ToString() ?? "none"}"
            + (before == row ? string.Empty : $" then {row?.ToString() ?? "none"}")
            + $", {dpi} dpi, click {x},{y}");
        if (row is not { } rect || double.IsInfinity(rect.Y)) return;
        if (x < 0 || y < 0 || x >= box.Width || y >= box.Height) return;
        var point = (nint)((y << 16) | (x & 0xffff));
        Win32.PostMessage(reading.ListWindow, 0x0201, 1, point);
        Win32.PostMessage(reading.ListWindow, 0x0202, 0, point);

        // The rows below an expanded block are read again at once, not a second from now.
        _tail = null;
    }

    public override void Dispose()
    {
        _walker?.Dispose();
        _walker = null;
        base.Dispose();
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
