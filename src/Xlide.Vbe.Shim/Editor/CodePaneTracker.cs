using System.Runtime.InteropServices;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>One code pane window the editor is currently showing.</summary>
/// <param name="Window">Handle of the pane window.</param>
/// <param name="Component">Component the pane is showing, or null when it could not be matched.</param>
/// <param name="Bounds">Screen rectangle of the pane.</param>
/// <param name="IsVisible">Whether the pane is currently shown.</param>
internal readonly record struct CodePane(nint Window, string? Component, PixelRect Bounds, bool IsVisible);

/// <summary>
/// Keeps an up to date picture of which code panes exist, where they are, and what they show.
///
/// This is the map an editor surface is positioned by, so it has to be correct rather than
/// approximately correct. Two sources are combined because neither is sufficient alone. Window
/// enumeration finds the pane windows and their rectangles, which the object model does not expose.
/// The object model lists which components actually have panes open, which the windows cannot say
/// reliably, because a pane window shares its class with the Immediate window and its caption is
/// localised.
///
/// Everything here runs on the host user interface thread.
/// </summary>
internal sealed class CodePaneTracker : IDisposable
{
    /// <summary>Window class the editor uses for code panes, and for the Immediate window.</summary>
    internal const string PaneClass = "VbaWindow";

    /// <summary>Window class of the editor's own frame.</summary>
    internal const string FrameClass = "wndclass_desked_gsk";

    private static CodePaneTracker? _enumerating;

    private readonly DispatchObject _editor;
    private readonly List<CodePane> _panes = [];
    private readonly List<nint> _candidates = [];

    private WindowEventHook? _hook;
    private bool _refreshing;

    /// <summary>
    /// Component names the editor last reported as having a pane open.
    ///
    /// Reading this costs several calls into the object model, and a single splitter drag raises a
    /// location change for every affected window on every mouse move. Re-reading it per event would
    /// turn one gesture into hundreds of cross-process calls on the thread the user is typing on.
    /// It changes only when a pane opens or closes, so it is cached and invalidated by exactly
    /// those events.
    /// </summary>
    private List<string>? _openComponents;

    public CodePaneTracker(DispatchObject editor) => _editor = editor;

    /// <summary>Panes as of the last refresh.</summary>
    public IReadOnlyList<CodePane> Panes => _panes;

    /// <summary>Raised after the set of panes, or any pane's position, has changed.</summary>
    public event Action<IReadOnlyList<CodePane>>? Changed;

    public void Start()
    {
        _hook = WindowEventHook.Install(OnWindowEvent);
        Refresh();
    }

    private void OnWindowEvent(WindowEvent windowEvent)
    {
        if (!windowEvent.AffectsLayout || windowEvent.IsCaret)
        {
            return;
        }

        // A pane appearing or disappearing is the only thing that can change which components have
        // panes open, so it is the only thing that invalidates the expensive half of a refresh.
        // Moves and resizes need nothing but window rectangles.
        if (windowEvent.IsCreate || windowEvent.IsDestroy)
        {
            _openComponents = null;
        }

        Refresh();
    }

    /// <summary>Rebuilds the picture from both sources.</summary>
    public void Refresh()
    {
        // Reading window rectangles can itself raise events on some systems. Re-entering here would
        // recurse without bound.
        if (_refreshing)
        {
            return;
        }

        _refreshing = true;

        try
        {
            var open = _openComponents ??= ReadOpenComponents();
            var found = FindPaneWindows();
            var updated = new List<CodePane>(found.Count);

            foreach (var window in found)
            {
                var caption = ReadWindowText(window);
                if (CodePaneCaption.IsKnownNonCodePane(caption))
                {
                    continue;
                }

                var component = CodePaneCaption.MatchComponent(caption, open);
                if (component is null)
                {
                    // Unmatched windows are dropped rather than guessed at. Placing a surface over
                    // the wrong window is worse than placing none.
                    continue;
                }

                updated.Add(new CodePane(window, component, ReadBounds(window), Win32.IsWindowVisible(window)));
            }

            // Events arrive in bursts and most of them change nothing. Telling listeners that
            // nothing changed would make every one of them re-do its own work, which for a surface
            // positioned over a pane means moving a window that is already in the right place.
            if (SameAs(updated))
            {
                return;
            }

            _panes.Clear();
            _panes.AddRange(updated);

            Changed?.Invoke(_panes);
        }
        catch (COMException ex)
        {
            // The editor can be mid-teardown, in which case the object model stops answering. That
            // is expected and not worth surfacing as an error.
            Log.Info($"code panes: refresh abandoned, the editor stopped answering, 0x{ex.HResult:X8}");
        }
        catch (Exception ex)
        {
            Log.Error("code panes: refresh failed", ex);
        }
        finally
        {
            _refreshing = false;
        }
    }

    /// <summary>True when a freshly built list is identical to the one already published.</summary>
    private bool SameAs(List<CodePane> updated)
    {
        if (updated.Count != _panes.Count)
        {
            return false;
        }

        for (var i = 0; i < updated.Count; i++)
        {
            if (updated[i] != _panes[i])
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Names of the components the editor reports as having a pane open. This is the authority on
    /// what a pane window may legitimately be showing.
    /// </summary>
    private List<string> ReadOpenComponents()
    {
        var names = new List<string>();

        using var panes = _editor.GetObject("CodePanes");
        if (panes is null)
        {
            return names;
        }

        var count = panes.GetInt32("Count");
        for (var i = 1; i <= count; i++)
        {
            using var pane = panes.GetItem(i);
            using var module = pane?.GetObject("CodeModule");
            using var component = module?.GetObject("Parent");

            var name = component?.GetString("Name");
            if (!string.IsNullOrEmpty(name))
            {
                names.Add(name);
            }
        }

        return names;
    }

    /// <summary>
    /// Every visible window of the pane class, anywhere under the editor frame.
    ///
    /// Enumerated rather than walked as direct children: the panes are not all at the same depth,
    /// and asking only for direct children finds none of the docked ones.
    /// </summary>
    internal static unsafe HashSet<nint> VisiblePanes()
    {
        _visible = [];

        try
        {
            var ours = Win32.GetCurrentProcessId();

            nint frame = 0;
            while ((frame = Win32.FindWindowEx(0, frame, FrameClass, null)) != 0)
            {
                Win32.GetWindowThreadProcessId(frame, out var owner);
                if (owner != ours)
                {
                    continue;
                }

                Win32.EnumChildWindows(
                    frame,
                    (nint)(delegate* unmanaged<nint, nint, int>)&OnVisiblePane,
                    0);
            }

            return _visible;
        }
        finally
        {
            _visible = null;
        }
    }

    private static HashSet<nint>? _visible;

    [UnmanagedCallersOnly]
    private static int OnVisiblePane(nint window, nint parameter)
    {
        if (_visible is not null && Win32.IsWindowVisible(window) && ReadClassName(window) == PaneClass)
        {
            _visible.Add(window);
        }

        return 1;
    }

    private unsafe List<nint> FindPaneWindows()
    {
        _candidates.Clear();

        var frame = FindEditorFrame();
        if (frame == 0)
        {
            return _candidates;
        }

        // The callback is a plain function pointer under ahead-of-time compilation, so the instance
        // is reached through a static field for the duration of the call. Enumeration is synchronous
        // and single threaded, so the field cannot be observed by anyone else.
        _enumerating = this;

        try
        {
            Win32.EnumChildWindows(
                frame,
                (nint)(delegate* unmanaged<nint, nint, int>)&OnChildWindow,
                0);
        }
        finally
        {
            _enumerating = null;
        }

        return _candidates;
    }

    [UnmanagedCallersOnly]
    private static int OnChildWindow(nint window, nint parameter)
    {
        try
        {
            var tracker = _enumerating;
            if (tracker is not null && ReadClassName(window) == PaneClass)
            {
                tracker._candidates.Add(window);
            }
        }
        catch
        {
            // Nothing may escape into the enumeration.
        }

        return 1;
    }

    /// <summary>Finds the editor's own frame window, which every pane is a descendant of.</summary>
    private unsafe nint FindEditorFrame()
    {
        using var main = _editor.GetObject("MainWindow");
        if (main is null)
        {
            return 0;
        }

        // The object model exposes no window handle, so the frame is located by walking up from any
        // window the editor owns. Starting from a pane we already know is cheapest when we have one.
        foreach (var pane in _panes)
        {
            if (Win32.IsWindow(pane.Window))
            {
                var root = Win32.GetAncestor(pane.Window, Win32.GaRoot);
                if (root != 0 && ReadClassName(root) == FrameClass)
                {
                    return root;
                }
            }
        }

        return FindFrameByEnumeration();
    }

    private static unsafe nint FindFrameByEnumeration()
    {
        // GetTopWindow-style enumeration over the desktop would reach other processes. Walking our
        // own windows is enough: the editor frame is a top-level window in this process.
        var found = 0L;
        var self = (uint)Environment.ProcessId;

        for (var window = Win32.GetTopWindow(0); window != 0; window = Win32.GetWindow(window, Win32.GwHwndNext))
        {
            uint owner;
            Win32.GetWindowThreadProcessId(window, &owner);
            if (owner != self)
            {
                continue;
            }

            if (ReadClassName(window) == FrameClass)
            {
                found = window;
                break;
            }
        }

        return (nint)found;
    }

    private static unsafe string ReadClassName(nint window)
    {
        const int capacity = 64;
        var buffer = stackalloc char[capacity];
        var length = Win32.GetClassName(window, buffer, capacity);
        return length <= 0 ? string.Empty : new string(buffer, 0, length);
    }

    private static unsafe string ReadWindowText(nint window)
    {
        const int capacity = 256;
        var buffer = stackalloc char[capacity];
        var length = Win32.GetWindowText(window, buffer, capacity);
        return length <= 0 ? string.Empty : new string(buffer, 0, length);
    }

    /// <summary>
    /// The pane's client area in screen coordinates: the part that shows text, without the pane's
    /// own borders.
    ///
    /// The client area rather than the window rectangle, because a pane maximised inside the
    /// editor's document area has a window rectangle that reaches outside it. Its own borders and
    /// the space its caption would occupy sit above and to the left of the area anything is drawn
    /// in. A surface placed on the window rectangle is pushed up out of view by exactly that much,
    /// and shows the module scrolled down by the first couple of lines.
    /// </summary>
    private static unsafe PixelRect ReadBounds(nint window)
    {
        Rect client;
        if (!Win32.GetClientRect(window, &client))
        {
            return default;
        }

        var corners = stackalloc Point[2];
        corners[0] = new Point { X = client.Left, Y = client.Top };
        corners[1] = new Point { X = client.Right, Y = client.Bottom };

        // A destination of zero means the screen.
        Marshal.SetLastSystemError(0);
        if (Win32.MapWindowPoints(window, 0, corners, 2) == 0 && Marshal.GetLastSystemError() != 0)
        {
            return default;
        }

        return new PixelRect(
            Math.Min(corners[0].X, corners[1].X),
            Math.Min(corners[0].Y, corners[1].Y),
            Math.Max(corners[0].X, corners[1].X),
            Math.Max(corners[0].Y, corners[1].Y));
    }

    public void Dispose()
    {
        _hook?.Dispose();
        _hook = null;
        _panes.Clear();
    }
}
