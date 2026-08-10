using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The ghost palettes' own reading thread: the Locals and Watches readers live here, and every
/// accessibility call they make happens here, never on the host thread.
///
/// The day the panels died (2026-08-05, "locals: skipped 4 unreadable element(s)", 24 sessions
/// straight) the crash itself turned out to be an undersized variant out-parameter - see
/// ComVariantBlock in Interop/UiAutomation.cs for that story. The thread stays because the old
/// arrangement was wrong independently: the variable rows are served by the editor's own
/// accessibility provider inside this process, and a client asking from the provider's OWN
/// thread is the configuration the accessibility framework documents as unsupported - it
/// re-enters the provider mid-call and worked on borrowed luck. A separate thread asks the way
/// every real client does: serviced while the editor's thread pumps, which during a break it
/// does constantly - that is what makes the debugger interactive.
///
/// The host thread therefore only ever ASKS (<see cref="RequestRead"/>) and LOOKS
/// (<see cref="Locals"/>, <see cref="Watches"/>); it never waits. If a read wedges, the
/// readings go stale and the log names it, but the editor never hangs. Readings are cleared at
/// break exit so a new break starts empty rather than showing the previous break's variables.
/// </summary>
internal sealed class GhostReaderThread : IDisposable
{
    private readonly nint _localsWindow;
    private readonly nint _watchWindow;

    /// <summary>Set to ask for a fresh read; set once more to let the loop see the stop flag.</summary>
    private readonly AutoResetEvent _wake = new(initialState: true);

    private Thread? _thread;
    private volatile bool _stopping;

    private LocalsReader.LocalsSnapshot? _locals;
    private IReadOnlyList<WatchReader.WatchRow>? _watches;

    private GhostReaderThread(nint localsWindow, nint watchWindow)
    {
        _localsWindow = localsWindow;
        _watchWindow = watchWindow;
    }

    /// <summary>The latest Locals reading, or null when none has succeeded since the last clear.</summary>
    public LocalsReader.LocalsSnapshot? Locals => Volatile.Read(ref _locals);

    /// <summary>The latest Watches reading, or null when none has succeeded since the last clear.</summary>
    public IReadOnlyList<WatchReader.WatchRow>? Watches => Volatile.Read(ref _watches);

    /// <summary>Starts the thread, or returns null when there is no ghost to read.</summary>
    public static GhostReaderThread? Start(nint localsWindow, nint watchWindow)
    {
        if (localsWindow == 0 && watchWindow == 0)
        {
            return null;
        }

        var readers = new GhostReaderThread(localsWindow, watchWindow);
        readers._thread = new Thread(readers.Run)
        {
            IsBackground = true,
            Name = "xlide-ghost-reader",
        };
        readers._thread.Start();
        return readers;
    }

    /// <summary>Asks for a fresh read soon. Free to call every poll tick; wakes coalesce.</summary>
    public void RequestRead() => _wake.Set();

    /// <summary>
    /// Forgets both readings. Called at break exit: the next break must start with nothing
    /// rather than with the previous break's variables, which are exactly stale enough to
    /// mislead ("stale variables are worse than none").
    /// </summary>
    public void ClearReadings()
    {
        Volatile.Write(ref _locals, null);
        Volatile.Write(ref _watches, null);
    }

    private void Run()
    {
        // The automation object, the elements, and every read stay on this one thread, so the
        // apartment can simply be the multithreaded one - the same shape as any external
        // accessibility client.
        var initialized = Win32.CoInitializeEx(0, Win32.ApartmentMultithreaded) >= 0;

        LocalsReader? locals = null;
        WatchReader? watches = null;

        try
        {
            locals = _localsWindow != 0 ? LocalsReader.Create(_localsWindow) : null;
            if (_localsWindow != 0)
            {
                Log.Info(locals is null
                    ? "locals: the ghost palette could not be read; the panel sits idle"
                    : $"locals: ghost palette {_localsWindow:X} feeding the panel");
            }

            watches = _watchWindow != 0 ? WatchReader.Create(_watchWindow) : null;
            if (_watchWindow != 0)
            {
                Log.Info(watches is null
                    ? "watch: the ghost palette could not be read; the panel sits idle"
                    : $"watch: ghost palette {_watchWindow:X} feeding the panel");
            }

            while (!_stopping)
            {
                _wake.WaitOne();
                if (_stopping)
                {
                    break;
                }

                if (locals?.Read() is { } snapshot)
                {
                    Volatile.Write(ref _locals, snapshot);
                }

                if (watches?.Read() is { } reading)
                {
                    Volatile.Write(ref _watches, reading);
                }
            }
        }
        catch (Exception ex)
        {
            // Nothing on this thread may escape: an unhandled exception here would take the
            // host process down. The panels go quiet; the log says why.
            Log.Error("ghost reader: the reading thread died", ex);
        }
        finally
        {
            locals?.Dispose();
            watches?.Dispose();

            if (initialized)
            {
                Win32.CoUninitialize();
            }
        }
    }

    public void Dispose()
    {
        _stopping = true;
        _wake.Set();

        // Bounded: the thread may be inside a cross-thread accessibility call that THIS thread
        // has to pump to answer, so waiting forever here would deadlock the host. An abandoned
        // background thread costs nothing worse than a stale log line at process exit.
        if (_thread is { } thread && !thread.Join(500))
        {
            Log.Warn("ghost reader: the reading thread did not stop in time; abandoned");
        }

        _thread = null;
        ClearReadings();
    }
}
