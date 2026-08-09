using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Reads what the editor's own Locals window holds, through the accessibility interface: each
/// variable is a list item whose name is the row's columns run together, and an edit control
/// names the broken procedure. The editor exposes no debugger object — no stack, no frames, no
/// variables — so its own window is the only place this information surfaces.
///
/// The window it reads is the GHOST PALETTE (see AddInSession.PrepareLocalsGhost): the native
/// Locals window floated through the object model, made layered at alpha zero, and parked off
/// the virtual screen. The editor only feeds a window with a paintable surface (lesson 25 —
/// hidden never fills, covered fills unreliably and never on a step), and a layered window
/// renders into its own surface regardless of position or occlusion, so the ghost is fed
/// faithfully through every break and step while being impossible to see. Probed 2026-08-04:
/// counter tracked 1 through 4 across steps at alpha 0, off screen (probed 2026-08-04).
///
/// Everything here — creation, connection, every read — runs on the ghost reading thread, never
/// the host's: the rows are served by an accessibility provider inside this process, and a
/// client on the provider's own thread is the unsupported configuration. The day the panel
/// actually died (2026-08-05) the crash was an undersized variant out-parameter — the story is
/// on ComVariantBlock — but the thread separation stands on its own; GhostReaderThread has the
/// full rationale.
/// </summary>
internal sealed class LocalsReader : IDisposable
{
    /// <summary>One row of the Locals window, its three columns separated.</summary>
    public readonly record struct LocalRow(string Expression, string Value, string Type);

    /// <summary>A reading: which procedure is broken, and the variables in scope.</summary>
    public sealed record LocalsSnapshot(string? Context, IReadOnlyList<LocalRow> Rows);

    private readonly nint _window;

    private ComHandle<IUIAutomation>? _automation;
    private ComHandle<IUIAutomationElement>? _element;
    private nint _condition;

    /// <summary>
    /// Consecutive read failures, and when the next attempt may run. A fault does not end the
    /// reader for the session — the first failure of a streak is logged in full, the rest wait
    /// out a pause quietly, and a later success announces the recovery.
    /// </summary>
    private int _consecutiveFailures;
    private long _retryAt;

    private LocalsReader(nint window) => _window = window;

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

    private bool Connect()
    {
        var hr = Win32.CoCreateInstance(
            in UiAutomationIds.AutomationClass,
            0,
            Win32.ClassContextInProcessServer,
            in UiAutomationIds.Automation,
            out var automation);

        if (hr < 0 || automation == 0)
        {
            Log.Info($"locals: the accessibility interface is unavailable, 0x{hr:X8}");
            return false;
        }

        _automation = ComHandle<IUIAutomation>.Own(automation);
        if (_automation is null)
        {
            return false;
        }

        if (_automation.Target.ElementFromHandle(_window, out var elementPointer) < 0 || elementPointer == 0)
        {
            Log.Info("locals: the window has no accessible element");
            return false;
        }

        _element = ComHandle<IUIAutomationElement>.Own(elementPointer);
        if (_element is null)
        {
            return false;
        }

        // Kept for the reader's lifetime: a read happens on every debug poll in a break, and the
        // condition never changes.
        if (_automation.Target.CreateTrueCondition(out _condition) < 0 || _condition == 0)
        {
            _condition = 0;
            return false;
        }

        return true;
    }

    /// <summary>
    /// The window's current content, or null when it cannot be read.
    ///
    /// The placeholder rows the window shows outside a break — "&lt;No Variables&gt;" — do not
    /// parse as rows, so an idle window reads as an empty snapshot rather than as a variable
    /// with an angle-bracketed name.
    /// </summary>
    public LocalsSnapshot? Read()
    {
        if (_element is null || _condition == 0 || Environment.TickCount64 < _retryAt)
        {
            return null;
        }

        // Named so a fault can say which call it died in: a native fault inside the
        // accessibility library carries no managed frames, and its bare NullReferenceException
        // points at everything and nothing (2026-08-05).
        var stage = "findAll";

        try
        {
            if (_element.Target.FindAll(UiAutomationIds.Descendants, _condition, out var arrayPointer) < 0
                || arrayPointer == 0)
            {
                return null;
            }

            stage = "length";
            using var array = ComHandle<IUIAutomationElementArray>.Own(arrayPointer);
            if (array is null || array.Target.GetLength(out var count) < 0)
            {
                return null;
            }

            string? context = null;
            List<LocalRow>? rows = null;
            var poisoned = 0;

            for (var i = 0; i < count; i++)
            {
                // Each element fends for itself: one descendant whose property fetch faults
                // natively must cost that element alone, not the whole reading.
                try
                {
                    stage = $"element {i} of {count}";
                    if (array.Target.GetElement(i, out var childPointer) < 0 || childPointer == 0)
                    {
                        continue;
                    }

                    using var child = ComHandle<IUIAutomationElement>.Own(childPointer);
                    if (child is null)
                    {
                        continue;
                    }

                    stage = $"control type {i}";
                    if (child.Target.GetCurrentPropertyValue(UiAutomationIds.ControlTypeProperty, out var kind) < 0)
                    {
                        continue;
                    }

                    var controlType = kind.AsInt32();
                    if (controlType != UiAutomationIds.ListItemControl
                        && controlType != UiAutomationIds.EditControl
                        && controlType != UiAutomationIds.PaneControl)
                    {
                        continue;
                    }

                    stage = $"name {i}";
                    if (child.Target.GetCurrentPropertyValue(UiAutomationIds.NameProperty, out var name) < 0)
                    {
                        continue;
                    }

                    var text = name.TakeString();
                    if (text is null)
                    {
                        continue;
                    }

                    if (controlType == UiAutomationIds.ListItemControl)
                    {
                        if (ParseRow(text) is { } row)
                        {
                            (rows ??= []).Add(row);
                        }
                    }
                    else if (controlType == UiAutomationIds.EditControl && text.Length > 0)
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
                }
                catch
                {
                    poisoned++;
                }
            }

            if (poisoned > 0)
            {
                Log.Verbose($"locals: skipped {poisoned} unreadable element(s)");
            }

            if (_consecutiveFailures > 0)
            {
                Log.Info($"locals: recovered after {_consecutiveFailures} failed read(s)");
                _consecutiveFailures = 0;
            }

            return new LocalsSnapshot(context, rows is null ? [] : rows);
        }
        catch (Exception ex)
        {
            // Backed off rather than stopped: this runs on every requested read, and a
            // recurring fault would write the same line several times a second — so the first
            // failure of a streak is logged in full and the rest wait out a pause quietly.
            _consecutiveFailures++;
            _retryAt = Environment.TickCount64 + 5000;
            if (_consecutiveFailures == 1)
            {
                Log.Error($"locals: the read died at {stage}, backing off", ex);
            }

            return null;
        }
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
    /// "Expression counter Value 42 Type Long". The expression is a single token — an identifier
    /// or an indexed path, never containing a space — while the value can be anything, including
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

    public void Dispose()
    {
        if (_condition != 0)
        {
            Marshal.Release(_condition);
            _condition = 0;
        }

        _element?.Dispose();
        _element = null;

        _automation?.Dispose();
        _automation = null;
    }
}
