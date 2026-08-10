using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Reads what the editor's own Watches window holds, through the accessibility interface: each
/// watch is a list item whose name is the row's four columns run together. The editor exposes
/// no debugger object - no watches, no values - so its own window is the only place this
/// information surfaces.
///
/// The window it reads is a GHOST PALETTE, prepared exactly the way the Locals one is (see
/// AddInSession.PrepareLocalsGhost and lesson 29): floated through the object model, made
/// layered at alpha zero, and parked off the virtual screen. The editor feeds any window with
/// a paintable surface, and a layered window always has one, so the ghost is fed through every
/// break and step while being impossible to see.
/// </summary>
internal sealed class WatchReader : IDisposable
{
    /// <summary>One row of the Watches window, its four columns separated.</summary>
    public readonly record struct WatchRow(string Expression, string Value, string Type, string Context);

    private readonly nint _window;

    private ComHandle<IUIAutomation>? _automation;
    private ComHandle<IUIAutomationElement>? _element;
    private nint _condition;

    /// <summary>Failure streak and backoff, the same manner as the Locals reader's.</summary>
    private int _consecutiveFailures;
    private long _retryAt;

    private WatchReader(nint window) => _window = window;

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
            Log.Info($"watch: the accessibility interface is unavailable, 0x{hr:X8}");
            return false;
        }

        _automation = ComHandle<IUIAutomation>.Own(automation);
        if (_automation is null)
        {
            return false;
        }

        if (_automation.Target.ElementFromHandle(_window, out var elementPointer) < 0 || elementPointer == 0)
        {
            Log.Info("watch: the window has no accessible element");
            return false;
        }

        _element = ComHandle<IUIAutomationElement>.Own(elementPointer);
        if (_element is null)
        {
            return false;
        }

        if (_automation.Target.CreateTrueCondition(out _condition) < 0 || _condition == 0)
        {
            _condition = 0;
            return false;
        }

        return true;
    }

    /// <summary>The window's current rows, or null when it cannot be read.</summary>
    public IReadOnlyList<WatchRow>? Read()
    {
        if (_element is null || _condition == 0 || Environment.TickCount64 < _retryAt)
        {
            return null;
        }

        try
        {
            if (_element.Target.FindAll(UiAutomationIds.Descendants, _condition, out var arrayPointer) < 0
                || arrayPointer == 0)
            {
                return null;
            }

            using var array = ComHandle<IUIAutomationElementArray>.Own(arrayPointer);
            if (array is null || array.Target.GetLength(out var count) < 0)
            {
                return null;
            }

            List<WatchRow>? rows = null;
            var poisoned = 0;

            for (var i = 0; i < count; i++)
            {
                // Each element fends for itself, the same manner as the Locals reader: one
                // descendant whose property fetch faults natively costs that element alone.
                try
                {
                    if (array.Target.GetElement(i, out var childPointer) < 0 || childPointer == 0)
                    {
                        continue;
                    }

                    using var child = ComHandle<IUIAutomationElement>.Own(childPointer);
                    if (child is null)
                    {
                        continue;
                    }

                    if (child.Target.GetCurrentPropertyValue(UiAutomationIds.ControlTypeProperty, out var kind) < 0
                        || kind.AsInt32() != UiAutomationIds.ListItemControl)
                    {
                        continue;
                    }

                    if (child.Target.GetCurrentPropertyValue(UiAutomationIds.NameProperty, out var name) < 0)
                    {
                        continue;
                    }

                    var text = name.TakeString();
                    if (text is not null && ParseRow(text) is { } row)
                    {
                        (rows ??= []).Add(row);
                    }
                }
                catch
                {
                    poisoned++;
                }
            }

            if (poisoned > 0)
            {
                Log.Verbose($"watch: skipped {poisoned} unreadable element(s)");
            }

            if (_consecutiveFailures > 0)
            {
                Log.Info($"watch: recovered after {_consecutiveFailures} failed read(s)");
                _consecutiveFailures = 0;
            }

            return rows is null ? [] : rows;
        }
        catch (Exception ex)
        {
            // Backed off rather than stopped, the same manner as the Locals reader: the first
            // failure of a streak is logged in full, the rest wait out a pause quietly.
            _consecutiveFailures++;
            _retryAt = Environment.TickCount64 + 5000;
            if (_consecutiveFailures == 1)
            {
                Log.Error("watch: the window could not be read, backing off", ex);
            }

            return null;
        }
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
