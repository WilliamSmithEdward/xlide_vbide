using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Observes window activity in this process and reports it to one listener.
///
/// The editor raises almost nothing through its object model. There is no notification that a code
/// pane appeared, moved, took focus, or closed, so those facts are recovered from window events
/// instead.
///
/// Two constraints shape this type. The callback the operating system invokes must be a plain
/// function pointer, which under ahead-of-time compilation means a static method with no captured
/// state, so the instance is reached through a static field rather than through the callback
/// arguments. And the hook must be removed before the host begins tearing itself down, because a
/// hook that outlives its library is a callback into freed code.
/// </summary>
internal sealed unsafe class WindowEventHook : IDisposable
{
    /// <summary>
    /// The single live hook. There is one editor per add-in connection and one hook per editor, so
    /// a static reference is honest about the shape rather than a shortcut. It is only ever written
    /// from the host user interface thread.
    /// </summary>
    private static WindowEventHook? _current;

    private readonly Action<WindowEvent> _listener;
    private nint _hook;

    private WindowEventHook(Action<WindowEvent> listener) => _listener = listener;

    /// <summary>
    /// Starts observing the calling process. Returns null when the hook could not be installed, in
    /// which case the caller keeps working without live layout tracking rather than failing.
    /// </summary>
    public static WindowEventHook? Install(Action<WindowEvent> listener)
    {
        if (_current is not null)
        {
            Log.Warn("window events: a hook is already installed, refusing to install a second");
            return null;
        }

        var hook = new WindowEventHook(listener);
        _current = hook;

        // Scoped to this process and delivered on our own message loop. Injecting the callback into
        // the thread that raised the event would mean running our code inside the host's internals
        // for no benefit: the thread whose events matter is the one we are already on.
        var processId = (uint)Environment.ProcessId;

        hook._hook = Win32.SetWinEventHook(
            Win32.EventObjectCreate,
            Win32.EventObjectNameChange,
            0,
            (nint)(delegate* unmanaged<nint, uint, nint, int, int, uint, uint, void>)&OnWindowEvent,
            processId,
            0,
            Win32.WinEventOutOfContext);

        if (hook._hook == 0)
        {
            Log.Error($"window events: the hook could not be installed, error {Marshal.GetLastWin32Error()}");
            _current = null;
            return null;
        }

        Log.Info("window events: hook installed");
        return hook;
    }

    [UnmanagedCallersOnly]
    private static void OnWindowEvent(
        nint hook,
        uint eventId,
        nint window,
        int objectId,
        int childId,
        uint eventThread,
        uint eventTime)
    {
        // Nothing may escape into the operating system's dispatch, and nothing may be slow: this
        // runs on the host's user interface thread, between the host's own messages.
        try
        {
            var current = _current;
            if (current is null || window == 0)
            {
                return;
            }

            // Events about a child object inside a window are not events about the window. The one
            // exception worth taking is the caret, which is how the text cursor becomes observable.
            if (objectId != Win32.ObjIdWindow && objectId != Win32.ObjIdCaret)
            {
                return;
            }

            current._listener(new WindowEvent(eventId, window, objectId));
        }
        catch (Exception ex)
        {
            Log.Error("window events: a listener failed", ex);
        }
    }

    public void Dispose()
    {
        var hook = _hook;
        _hook = 0;

        if (hook != 0)
        {
            Win32.UnhookWinEvent(hook);
            Log.Info("window events: hook removed");
        }

        if (ReferenceEquals(_current, this))
        {
            _current = null;
        }
    }
}

/// <summary>One observed window event, reduced to the parts this add-in acts on.</summary>
/// <param name="EventId">The raw event identifier.</param>
/// <param name="Window">The window the event concerns.</param>
/// <param name="ObjectId">Which object within the window the event concerns.</param>
internal readonly record struct WindowEvent(uint EventId, nint Window, int ObjectId)
{
    public bool IsCreate => EventId == Win32.EventObjectCreate;

    public bool IsDestroy => EventId == Win32.EventObjectDestroy;

    public bool IsShow => EventId == Win32.EventObjectShow;

    public bool IsHide => EventId == Win32.EventObjectHide;

    public bool IsFocus => EventId == Win32.EventObjectFocus;

    public bool IsLocationChange => EventId == Win32.EventObjectLocationChange;

    public bool IsNameChange => EventId == Win32.EventObjectNameChange;

    /// <summary>True when the event describes the text cursor rather than the window frame.</summary>
    public bool IsCaret => ObjectId == Win32.ObjIdCaret;

    /// <summary>
    /// True when the event can change where a pane sits or whether it is the one being edited.
    /// Anything positioned over a pane has to react to all of these.
    /// </summary>
    public bool AffectsLayout => IsCreate || IsDestroy || IsShow || IsHide || IsLocationChange;

    /// <summary>The event as one word, for the development log.</summary>
    public string Describe() => EventId switch
    {
        Win32.EventObjectCreate => "create",
        Win32.EventObjectDestroy => "destroy",
        Win32.EventObjectShow => "show",
        Win32.EventObjectHide => "hide",
        Win32.EventObjectReorder => "reorder",
        Win32.EventObjectFocus => "focus",
        Win32.EventObjectLocationChange => "move",
        Win32.EventObjectNameChange => "rename",
        _ => $"event {EventId:X}",
    };
}
