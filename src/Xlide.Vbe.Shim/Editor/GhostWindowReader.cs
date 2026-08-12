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

        return true;
    }

    /// <summary>An element the walk offers: its control type, and its accessible name.</summary>
    protected delegate void TakeElement(int controlType, string text);

    /// <summary>
    /// One guarded walk over the window's accessible descendants: control types the reader does
    /// not want are skipped before their name is fetched, and every wanted one reaches `take`.
    /// False when the walk could not run at all - not connected, inside the backoff, or died -
    /// which is the caller's cue to answer null rather than empty.
    /// </summary>
    protected bool TryWalk(Func<int, bool> wants, TakeElement take)
    {
        if (_element is null || _condition == 0 || Environment.TickCount64 < _retryAt)
        {
            return false;
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
                return false;
            }

            stage = "length";
            using var array = ComHandle<IUIAutomationElementArray>.Own(arrayPointer);
            if (array is null || array.Target.GetLength(out var count) < 0)
            {
                return false;
            }

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
                    if (!wants(controlType))
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

                    take(controlType, text);
                }
                catch
                {
                    poisoned++;
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
