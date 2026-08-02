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

    private bool _failed;

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
    /// The window is a scrolling buffer that the editor trims and the developer can clear, so what
    /// it holds is not always what it held plus something. When the new text continues the old, the
    /// difference is new output; when it does not, the buffer was cleared or trimmed and the whole
    /// of it is treated as new rather than trying to work out which part survived.
    /// </summary>
    public void Poll()
    {
        if (_failed || _text is null)
        {
            return;
        }

        var current = ReadAll();
        if (current is null || current == _seen)
        {
            return;
        }

        var addition = current.StartsWith(_seen, StringComparison.Ordinal)
            ? current[_seen.Length..]
            : current;

        _seen = current;

        if (addition.Length > 0)
        {
            Appended?.Invoke(addition);
        }
    }

    /// <summary>Treats everything currently in the window as already seen.</summary>
    public void Reset() => _seen = ReadAll() ?? string.Empty;

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

            if (range.Target.GetText(UiAutomationIds.WholeRange, out var value) < 0 || value == 0)
            {
                return null;
            }

            try
            {
                return Marshal.PtrToStringBSTR(value);
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
