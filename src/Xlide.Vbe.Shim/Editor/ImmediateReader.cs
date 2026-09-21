using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Reads what the editor's own Immediate window holds, so Debug.Print reaches the surface.
///
/// Debug.Print writes into that window and nowhere else. It cannot be intercepted: Debug is
/// intrinsic to the language, the window exposes no handle through the object model, and asking
/// the window for its text answers with its caption. Hiding the window without reading it would
/// take Debug.Print away from the developer entirely, which is not a restyling.
///
/// The one thing that can read it is the interface a screen reader uses. The window is a document
/// to that interface, and a document's text is readable in full, including while the window is
/// hidden, which is exactly the arrangement here.
///
/// Reading is polled rather than subscribed. There is a change event for text, and it fires on the
/// editor's own thread while that thread is inside the developer's running code; taking the
/// callback there would reenter a thread already busy. Reading what changed after the fact costs
/// one call and cannot deadlock.
/// </summary>
internal sealed class ImmediateReader : IDisposable
{
    private readonly nint _window;

    private ComHandle<IUIAutomation>? _automation;
    private ComHandle<IUIAutomationTextPattern>? _text;

    /// <summary>Everything read so far, so only what is new is reported.</summary>
    private string _seen = string.Empty;

    /// <summary>The last raw reading, so the log records each change of state exactly once.</summary>
    private string? _lastRaw;

    private bool _failed;
    private bool _evaluating;

    private ImmediateReader(nint window) => _window = window;

    /// <summary>Raised with each run of text the window has gained.</summary>
    public Action<string>? Appended { get; set; }

    /// <summary>
    /// Prepares a reader for a window, or null when the accessibility interface cannot be reached.
    /// A host without it is not a failure worth stopping for; Debug.Print simply stays where it is.
    /// </summary>
    public static ImmediateReader? Create(nint window)
    {
        if (window == 0)
        {
            return null;
        }

        var reader = new ImmediateReader(window);
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
            Log.Info($"immediate: the accessibility interface is unavailable, 0x{hr:X8}");
            return false;
        }

        _automation = ComHandle<IUIAutomation>.Own(automation);
        if (_automation is null)
        {
            return false;
        }

        if (_automation.Target.ElementFromHandle(_window, out var elementPointer) < 0 || elementPointer == 0)
        {
            Log.Info("immediate: the window has no accessible element");
            return false;
        }

        using var element = ComHandle<IUIAutomationElement>.Own(elementPointer);
        if (element is null)
        {
            return false;
        }

        // The window itself carries no text; the document inside it does.
        if (_automation.Target.CreateTrueCondition(out var condition) < 0 || condition == 0)
        {
            return false;
        }

        try
        {
            if (element.Target.FindFirst(UiAutomationIds.Descendants, condition, out var documentPointer) < 0
                || documentPointer == 0)
            {
                Log.Info("immediate: the window contains no document");
                return false;
            }

            using var document = ComHandle<IUIAutomationElement>.Own(documentPointer);
            if (document is null)
            {
                return false;
            }

            // The pattern arrives as a bare unknown, so it is asked for the interface that has
            // the text on it.
            if (document.Target.GetCurrentPattern(UiAutomationIds.TextPatternId, out var patternPointer) < 0
                || patternPointer == 0)
            {
                Log.Info("immediate: the document exposes no text");
                return false;
            }

            try
            {
                if (Marshal.QueryInterface(patternPointer, in UiAutomationIds.TextPattern, out var typed) < 0)
                {
                    return false;
                }

                _text = ComHandle<IUIAutomationTextPattern>.Own(typed);
                return _text is not null;
            }
            finally
            {
                Marshal.Release(patternPointer);
            }
        }
        finally
        {
            Marshal.Release(condition);
        }
    }

    /// <summary>
    /// Reports whatever the window has gained since the last read.
    ///
    /// Three things about the reading are not what they look like. The text ends in the line the
    /// caret sits on, and new output is inserted BEFORE that tail, so the raw text never simply
    /// grows: compared raw, every reading looked like a rewrite, which first replayed the whole
    /// buffer under every evaluation and then, made silent, swallowed the output instead. The
    /// tail is trimmed away and lines are all that is compared. An empty reading is ignored
    /// outright: the window is hidden and nobody can clear it, but a project reset makes it read
    /// as empty for a moment, and adopting that moment as the baseline is what a replay grows
    /// from. When old lines roll out, the surviving suffix identifies the newly printed output.
    /// A rewrite with no overlap is adopted without replaying it.
    /// </summary>
    public void Poll()
    {
        if (_failed || _text is null || _evaluating)
        {
            return;
        }

        var raw = ReadAll();

        // Once per change of state, the raw truth: what the window says it holds and what the
        // baseline is. This is what tells a dead capture apart from a window nothing wrote to.
        if (raw != _lastRaw)
        {
            _lastRaw = raw;
            Log.Info($"immediate: window {(raw is null ? "unreadable" : $"{raw.Length} char(s) '{Tail(raw)}'")}, seen {_seen.Length}");
        }

        var current = Normalise(raw);
        if (current is null || current.Length == 0 || current == _seen)
        {
            return;
        }

        var addition = Xlide.Vbe.Core.Editor.ImmediateTranscript.Added(_seen, current);
        _seen = current;

        if (addition is { Length: > 0 })
        {
            Appended?.Invoke(addition);
        }
    }

    /// <summary>
    /// The reading without the caret's own tail, which is presentation and not output.
    ///
    /// Only the caret line and NUL are removed. Trimming output spaces made the next poll
    /// replay a numeric result's trailing space as a spurious Debug.Print line (#21).
    /// </summary>
    private static string? Normalise(string? text)
    {
        if (text is null)
        {
            return null;
        }

        return text.EndsWith("\r\n\0", StringComparison.Ordinal) ? text[..^3] : text.TrimEnd('\0');
    }

    /// <summary>The end of a reading with every non-printable made visible, for the log.</summary>
    private static string Tail(string text)
    {
        var tail = text.Length > 60 ? text[^60..] : text;
        var printable = new System.Text.StringBuilder(tail.Length + 8);

        foreach (var c in tail)
        {
            if (c == '\r')
            {
                printable.Append("\\r");
            }
            else if (c == '\n')
            {
                printable.Append("\\n");
            }
            else if (c < ' ' || c > '~')
            {
                printable.Append($"\\u{(int)c:X4}");
            }
            else
            {
                printable.Append(c);
            }
        }

        return printable.ToString();
    }

    /// <summary>Treats everything currently in the window as already seen.</summary>
    public void Reset() => _seen = Normalise(ReadAll()) ?? string.Empty;

    /// <summary>
    /// The window's whole text as it stands, which nothing outside this class could see.
    ///
    /// The reader existed to push each NEW run of output to the page, and that is all anybody
    /// could observe: what the window says right now was reachable only by having watched every
    /// message that ever arrived. So a probe could type into the Immediate window and never read
    /// it, which is why the panel had a route and no suite (2026-08-07).
    /// </summary>
    public string? Text() => Normalise(ReadAll());

    /// <summary>
    /// Evaluates in the native paused scope. No project edits, clipboard use or global keys.
    /// Pending output is delivered to XLIDE before replacing the native working buffer. Keeping
    /// that buffer short avoids stale accessibility offsets when VBE rolls its history over.
    /// </summary>
    public ImmediateEvaluator.Result Evaluate(string line, Func<bool> stillPaused)
    {
        if (_evaluating)
        {
            return new("An Immediate evaluation is already in progress.", true);
        }

        if (_failed || _text is null || !Win32.IsWindow(_window)
            || Win32.GetWindowThreadProcessId(_window, out var owner) == 0
            || owner != Win32.GetCurrentProcessId())
        {
            return new("The native Immediate window is unavailable. The debugger was not reset.", true);
        }

        var marker = "'xlide:" + Guid.NewGuid().ToString("N");
        var command = line.Trim() + " " + marker;
        if (command.Length > 1023 || line.Any(char.IsControl))
        {
            return new("Enter one Immediate line, at most " + (1023 - marker.Length - 1)
                + " characters, without control characters.", true);
        }

        Poll(); // Deliver pending Debug.Print output before taking ownership of the transcript.
        _evaluating = true;
        try
        {
            if (!stillPaused())
            {
                return new("The debugger is no longer paused; the line was not executed.", true);
            }

            if (_text.Target.GetDocumentRange(out var pointer) < 0 || pointer == 0)
            {
                return new("The native Immediate caret could not be located.", true);
            }

            using (var range = ComHandle<IUIAutomationTextRange>.Own(pointer))
            {
                if (range is null || range.Target.Select() < 0)
                {
                    return new("The native Immediate caret could not be positioned.", true);
                }
            }

            // The first character replaces the selected native transcript. XLIDE's visible
            // history is retained; Poll above already delivered everything pending in it.
            foreach (var character in command)
            {
                Win32.SendMessage(_window, 0x0102, character, 1); // WM_CHAR
            }

            // Never press Enter unless the entire requested command landed on its own line.
            // This also catches a stale caret, an incomplete native line, or provider failure.
            var typed = Text();
            if (typed is null || !("\n" + typed).EndsWith("\n" + command, StringComparison.Ordinal))
            {
                return new("The native Immediate window did not accept the complete line; it was not executed.", true);
            }

            using var errors = new NativeImmediateErrors();
            Win32.SendMessage(_window, 0x0100, 13, 1); // WM_KEYDOWN: Enter executes; WM_CHAR does not.
            Win32.SendMessage(_window, 0x0101, 13, unchecked((nint)0xC0000001));
            if (errors.Message is { } error)
            {
                return new(error, true);
            }

            var output = Xlide.Vbe.Core.Editor.ImmediateTranscript.Output(ReadAll(), marker);
            return output is null
                ? new("The native evaluation completed, but its output could not be read in full. Do not retry statements with side effects automatically.", true)
                : new(output, false, HasOutput: true);
        }
        catch (Exception ex)
        {
            Log.Error("immediate: native evaluation failed", ex);
            return new("Native Immediate evaluation failed: " + ex.Message, true);
        }
        finally
        {
            Reset(); // The caller displays the result; the poller must not echo it a second time.
            _evaluating = false;
        }
    }

    private string? ReadAll()
    {
        var text = _text;
        if (text is null)
        {
            return null;
        }

        try
        {
            if (text.Target.GetDocumentRange(out var rangePointer) < 0 || rangePointer == 0)
            {
                return null;
            }

            using var range = ComHandle<IUIAutomationTextRange>.Own(rangePointer);
            if (range is null)
            {
                return null;
            }

            if (range.Target.GetText(1024 * 1024, out var value) < 0 || value == 0)
            {
                return null;
            }

            try
            {
                var raw = Marshal.PtrToStringBSTR(value);
                if (raw.Length < 1024) { return raw; }

                // The VBE provider caps each GetText call, even with maxLength = -1.
                // Read smaller ranges so a long history cannot hide the newest command.
                if (range.Target.Clone(out var copy) < 0 || copy == 0) { return null; }
                using var chunk = ComHandle<IUIAutomationTextRange>.Own(copy);
                if (chunk is null || chunk.Target.MoveEndpointByRange(1, range.Pointer, 0) < 0)
                {
                    return null;
                }

                var all = new System.Text.StringBuilder();
                while (all.Length < 1024 * 1024)
                {
                    if (chunk.Target.MoveEndpointByUnit(1, 0, 512, out var moved) < 0 || moved == 0
                        || chunk.Target.GetText(UiAutomationIds.WholeRange, out var part) < 0 || part == 0)
                    {
                        return null;
                    }

                    try { all.Append(Marshal.PtrToStringBSTR(part).TrimEnd('\0')); }
                    finally { Marshal.FreeBSTR(part); }
                    if (chunk.Target.CompareEndpoints(1, range.Pointer, 1, out var compared) < 0)
                    {
                        return null;
                    }
                    if (compared >= 0) { return all.Append('\0').ToString(); }
                    if (chunk.Target.MoveEndpointByRange(0, chunk.Pointer, 1) < 0) { return null; }
                }

                return null;
            }
            finally
            {
                Marshal.FreeBSTR(value);
            }
        }
        catch (Exception ex)
        {
            // Stopped rather than repeated. This runs on a timer, and a fault that recurs would
            // write the same line to the log several times a second.
            _failed = true;
            Log.Error("immediate: the window could not be read, no longer trying", ex);
            return null;
        }
    }

    public void Dispose()
    {
        Appended = null;

        _text?.Dispose();
        _text = null;

        _automation?.Dispose();
        _automation = null;
    }
}
