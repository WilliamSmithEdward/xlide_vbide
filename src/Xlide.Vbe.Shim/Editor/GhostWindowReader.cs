using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The shared half of reading a ghost palette through the accessibility interface: connect to
/// the window's element, walk its descendants handing each readable one to the caller, back
/// off on faults, and let go of everything on dispose.
///
/// The Locals and Watches readers were two transcriptions of this - the audit's B16 - and the
/// dangerous part is that the transcription included the hardening that came out of the
/// 2026-08-05 crash: the per-element try/catch, the poisoned counter, the 5000ms backoff with
/// its first-of-streak log, and the named stage that turns a bare native fault into a line
/// that says which call died. A correction to this walk had to be made twice, and the copy in
/// front of the maintainer was the only one that got it (the Watches copy never carried the
/// stage names at all). It lives here once now; a reader keeps only its control-type rule and
/// its row parser.
///
/// Everything here runs on the ghost reading thread, never the host's: the rows are served by
/// an accessibility provider inside this process, and a client on the provider's own thread is
/// the unsupported configuration. GhostReaderThread has the full rationale, and the ghost
/// palette story - the layered window at alpha zero, parked off screen, fed faithfully because
/// it has a paintable surface - is told at AddInSession.PrepareLocalsGhost and lesson 29.
/// </summary>
internal abstract class GhostWindowReader : IDisposable
{
    private readonly nint _window;

    /// <summary>The log prefix, so a locals line and a watch line stay tellable apart.</summary>
    private readonly string _who;

    private ComHandle<IUIAutomation>? _automation;
    private ComHandle<IUIAutomationElement>? _element;
    private nint _condition;
    private ComHandle<IUIAutomationCacheRequest>? _cacheRequest;
    private ComHandle<IUIAutomationCacheRequest>? _liveRequest;
    private ComHandle<IUIAutomationCacheRequest>? _elementRequest;

    /// <summary>Deeper than any Locals tree a person expands by hand; a cycle cannot recurse forever.</summary>
    private const int MaxDepth = 32;

    /// <summary>
    /// Consecutive read failures, and when the next attempt may run. A fault does not end the
    /// reader for the session - the first failure of a streak is logged in full, the rest wait
    /// out a pause quietly, and a later success announces the recovery.
    /// </summary>
    private int _consecutiveFailures;
    private long _retryAt;

    protected GhostWindowReader(nint window, string who)
    {
        _window = window;
        _who = who;
    }

    protected IUIAutomation? Automation => _automation?.Target;

    protected bool Connect()
    {
        var hr = Win32.CoCreateInstance(
            in UiAutomationIds.AutomationClass,
            0,
            Win32.ClassContextInProcessServer,
            in UiAutomationIds.Automation,
            out var automation);

        if (hr < 0 || automation == 0)
        {
            Log.Info($"{_who}: the accessibility interface is unavailable, 0x{hr:X8}");
            return false;
        }

        _automation = ComHandle<IUIAutomation>.Own(automation);
        if (_automation is null)
        {
            return false;
        }

        if (_automation.Target.ElementFromHandle(_window, out var elementPointer) < 0 || elementPointer == 0)
        {
            Log.Info($"{_who}: the window has no accessible element");
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

        // The whole raw subtree, with every property a reader uses, in ONE call. The walk used
        // to ask each element for each property one call at a time, and every one of those
        // calls crossed to the host thread. Measured 2026-10-17 with a Workbook and a Worksheet
        // expanded: 136 rows read in 500ms, on every poll of the break and twice per expand
        // click. Cached, 270 rows read in 110-150ms and an expand shows in about 200ms.
        if (_automation.Target.CreateCacheRequest(out var requestPointer) < 0 || requestPointer == 0)
        {
            return false;
        }

        _cacheRequest = ComHandle<IUIAutomationCacheRequest>.Own(requestPointer);
        if (_cacheRequest is null
            || _cacheRequest.Target.AddProperty(UiAutomationIds.ControlTypeProperty) < 0
            || _cacheRequest.Target.AddProperty(UiAutomationIds.NameProperty) < 0
            || _cacheRequest.Target.AddProperty(UiAutomationIds.BoundingRectangleProperty) < 0
            || _cacheRequest.Target.AddProperty(UiAutomationIds.NativeWindowHandleProperty) < 0
            || _cacheRequest.Target.AddProperty(UiAutomationIds.RuntimeIdProperty) < 0
            || _cacheRequest.Target.AddProperty(UiAutomationIds.ExpandCollapseStateProperty) < 0
            // Without the pattern itself in the request, the cached state of an expanded row
            // never read Expanded, though the live property did (measured 2026-10-17).
            || _cacheRequest.Target.AddPattern(UiAutomationIds.ExpandCollapsePatternId) < 0
            || _cacheRequest.Target.SetTreeScope(UiAutomationIds.Subtree) < 0
            // The raw view, as the old walk's FindAll with a true condition saw it; the default
            // filter is the control view, which can drop the panes the context box reads as.
            || _cacheRequest.Target.SetTreeFilter(_condition) < 0
            || _cacheRequest.Target.SetAutomationElementMode(UiAutomationIds.ElementModeNone) < 0)
        {
            return false;
        }

        // The same request with live elements behind the cached values, for the rare walk that
        // has to act on what it finds (a click, a selection) rather than only read it.
        if (_cacheRequest.Target.Clone(out var livePointer) < 0 || livePointer == 0)
        {
            return false;
        }

        _liveRequest = ComHandle<IUIAutomationCacheRequest>.Own(livePointer);
        if (_liveRequest is null
            || _liveRequest.Target.SetAutomationElementMode(UiAutomationIds.ElementModeFull) < 0
            || _liveRequest.Target.Clone(out var singlePointer) < 0 || singlePointer == 0)
        {
            return false;
        }

        // And one more for a single element at a time: a parent, or the next row.
        _elementRequest = ComHandle<IUIAutomationCacheRequest>.Own(singlePointer);
        return _elementRequest is not null
            && _elementRequest.Target.SetTreeScope(UiAutomationIds.ElementScope) >= 0;
    }

    /// <summary>A live request that fetches one element's properties, for the tree walker's *BuildCache calls.</summary>
    protected nint ElementRequest => _elementRequest?.Pointer ?? 0;

    /// <summary>An element's cached values as a walk would offer them, or null for one with no name.</summary>
    protected static WalkedElement? Describe(nint pointer, int depth)
    {
        using var element = ComHandle<IUIAutomationElement>.Borrow(pointer);
        if (element is null) return null;
        var target = element.Target;
        var identity = target.GetCachedPropertyValue(UiAutomationIds.RuntimeIdProperty, out var runtimeId) >= 0
            ? runtimeId.TakeRuntimeId() : null;
        var controlType = target.GetCachedPropertyValue(UiAutomationIds.ControlTypeProperty, out var kind) >= 0
            ? kind.AsInt32() : 0;
        if (target.GetCachedPropertyValue(UiAutomationIds.NameProperty, out var name) < 0 || name.TakeString() is not { } text)
        {
            return null;
        }

        var bounds = target.GetCachedPropertyValue(UiAutomationIds.BoundingRectangleProperty, out var rectangle) >= 0
            ? rectangle.TakeRectangle() : null;
        var window = target.GetCachedPropertyValue(UiAutomationIds.NativeWindowHandleProperty, out var handle) >= 0
            ? handle.AsInt32() : 0;
        var state = target.GetCachedPropertyValue(UiAutomationIds.ExpandCollapseStateProperty, out var expand) >= 0
            && expand.Type == UiVariant.TypeInt32 ? expand.AsInt32() : -1;
        return new WalkedElement(controlType, text, depth, bounds, window, identity, pointer, state);
    }

    /// <summary>
    /// An element the walk offers: its control type, its accessible name, how many list items
    /// enclose it by the walk's path (a Watches row's nesting), its rectangle, its window
    /// handle, its runtime identity, the element itself - borrowed for the length of the
    /// callback, and live only on a live walk - and its expand state, or -1 when it has none.
    /// </summary>
    protected readonly record struct WalkedElement(
        int ControlType, string Text, int Depth,
        (double X, double Y, double Width, double Height)? Bounds, nint Window,
        string? Identity, nint Element, int State = -1);

    /// <summary>
    /// One guarded walk over the window's accessible descendants, in document order: control
    /// types the reader does not want are skipped before their name is read, and every wanted
    /// one reaches `take`. False when the walk could not run at all - not connected, inside the
    /// backoff, or died - which is the caller's cue to answer null rather than empty.
    ///
    /// `live` walks with live elements behind the values. `descendRows` false keeps the walk out
    /// of list items' own child lists, for a reader that works out the nesting itself
    /// (LocalsReader).
    /// </summary>
    protected bool TryWalk(
        Func<int, bool> wants, Action<WalkedElement> take, bool live = false, bool descendRows = true)
    {
        var request = live ? _liveRequest : _cacheRequest;
        if (_element is null || request is null || Environment.TickCount64 < _retryAt)
        {
            return false;
        }

        // Named so a fault can say which call it died in: a native fault inside the
        // accessibility library carries no managed frames, and its bare NullReferenceException
        // points at everything and nothing (2026-08-05).
        var stage = "cache";

        try
        {
            if (_element.Target.BuildUpdatedCache(request.Pointer, out var rootPointer) < 0 || rootPointer == 0)
            {
                return false;
            }

            using var root = ComHandle<IUIAutomationElement>.Own(rootPointer);
            if (root is null)
            {
                return false;
            }

            // The editor's list providers do not keep their child lists straight (mapped
            // 2026-10-17 on the Locals window from outside the process). The list itself lists
            // every row in display order, an expanded row's members included, and an expanded
            // row may list its members a second time; a collapsed row may list the row after it.
            // So a row's own children are taken only while it reads Expanded, and an element is
            // offered once, at the first place it appears. LocalsReader goes further and does
            // not descend at all, because those child lists also go empty after an expand and a
            // collapse; it works the nesting out from the list's own chain of rows.
            var poisoned = 0;
            HashSet<string> visited = new(StringComparer.Ordinal);
            Visit(root.Target, 0);

            void Visit(IUIAutomationElement parent, int depth)
            {
                stage = $"children at depth {depth}";
                if (depth > MaxDepth || parent.GetCachedChildren(out var arrayPointer) < 0 || arrayPointer == 0)
                {
                    return;
                }

                using var array = ComHandle<IUIAutomationElementArray>.Own(arrayPointer);
                if (array is null || array.Target.GetLength(out var count) < 0)
                {
                    return;
                }

                for (var i = 0; i < count; i++)
                {
                    // Each element fends for itself: one descendant whose cached value faults
                    // natively must cost that element alone, not the whole reading.
                    try
                    {
                        stage = $"element {i} of {count} at depth {depth}";
                        if (array.Target.GetElement(i, out var childPointer) < 0 || childPointer == 0)
                        {
                            continue;
                        }

                        using var child = ComHandle<IUIAutomationElement>.Own(childPointer);
                        if (child is not null)
                        {
                            Offer(child, depth);
                        }
                    }
                    catch
                    {
                        poisoned++;
                    }
                }
            }

            void Offer(ComHandle<IUIAutomationElement> element, int depth)
            {
                var target = element.Target;
                var identity = target.GetCachedPropertyValue(UiAutomationIds.RuntimeIdProperty, out var runtimeId) >= 0
                    ? runtimeId.TakeRuntimeId() : null;
                if (identity is not null && !visited.Add(identity))
                {
                    return;
                }

                var controlType = target.GetCachedPropertyValue(UiAutomationIds.ControlTypeProperty, out var kind) >= 0
                    ? kind.AsInt32() : 0;
                var walked = controlType != 0 && wants(controlType) ? Describe(element.Pointer, depth) : null;
                if (walked is { } offered)
                {
                    take(offered);
                }

                if (controlType != UiAutomationIds.ListItemControl)
                {
                    Visit(target, depth);
                }
                // A provider without the expand state keeps the old behaviour: descend, and let
                // the runtime ids drop the copies.
                else if (descendRows
                    && (walked?.State ?? -1) is var state && (state == -1 || state == UiAutomationIds.Expanded))
                {
                    Visit(target, depth + 1);
                }
            }

            if (poisoned > 0)
            {
                Log.Verbose($"{_who}: skipped {poisoned} unreadable element(s)");
            }

            if (_consecutiveFailures > 0)
            {
                Log.Info($"{_who}: recovered after {_consecutiveFailures} failed read(s)");
                _consecutiveFailures = 0;
            }

            return true;
        }
        catch (Exception ex)
        {
            // Backed off rather than stopped: this runs on every requested read, and a
            // recurring fault would write the same line several times a second - so the first
            // failure of a streak is logged in full and the rest wait out a pause quietly.
            _consecutiveFailures++;
            _retryAt = Environment.TickCount64 + 5000;
            if (_consecutiveFailures == 1)
            {
                Log.Error($"{_who}: the read died at {stage}, backing off", ex);
            }

            return false;
        }
    }

    public virtual void Dispose()
    {
        _cacheRequest?.Dispose();
        _cacheRequest = null;
        _liveRequest?.Dispose();
        _liveRequest = null;
        _elementRequest?.Dispose();
        _elementRequest = null;

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
