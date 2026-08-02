using System.Runtime.InteropServices;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// A plain window that sits over a code pane and hosts the editing surface.
///
/// It is a child of the editor's own frame rather than of the pane. A child of the pane would be
/// clipped, moved, and destroyed by code that has no idea it exists, and the pane is recreated more
/// often than it looks. A sibling positioned over it survives all of that and is ours to place.
///
/// Nothing is painted here. The window exists to own a rectangle and give the browser surface a
/// parent; every pixel inside it comes from the surface.
/// </summary>
internal sealed unsafe class OverlayWindow : IDisposable
{
    private const string ClassName = "XlideEditorOverlay";

    /// <summary>Timer identifiers, scoped to this window.</summary>
    private const nuint WriteTimerId = 1;
    private const nuint PollTimerId = 2;

    /// <summary>Message that carries a posted action's handle in its LPARAM.</summary>
    private const uint WmRunAction = 0x8000 + 0x71;
    private static bool _classRegistered;
    private static readonly Lock ClassGate = new();

    private nint _handle;
    private GCHandle _self;

    /// <summary>
    /// Where the overlay was last put, so an unchanged position costs nothing.
    ///
    /// Window events arrive for anything that moves anywhere in the editor, and most of them leave
    /// the pane exactly where it was. Repositioning anyway would resize the browser and force a
    /// relayout of the document on every one of them.
    /// </summary>
    private PixelRect _placed;
    private bool _shown;

    private OverlayWindow()
    {
    }

    public nint Handle => _handle;

    /// <summary>Client area, in the coordinates the browser surface expects.</summary>
    public PixelRect ClientBounds()
    {
        Rect rect;
        if (_handle == 0 || !Win32.GetClientRect(_handle, &rect))
        {
            return default;
        }

        return new PixelRect(0, 0, rect.Right - rect.Left, rect.Bottom - rect.Top);
    }

    /// <summary>Raised when the window has been resized, so the surface can follow.</summary>
    public Action<PixelRect>? Resized { get; set; }

    /// <summary>Raised once, on the host thread, after <see cref="StartWriteTimer"/> elapses.</summary>
    public Action? Elapsed { get; set; }

    /// <summary>Raised repeatedly, on the host thread, while the poll timer runs.</summary>
    public Action? Polled { get; set; }

    /// <summary>
    /// Starts or restarts a one-shot timer. Restarting is what makes it a debounce: each call
    /// pushes the deadline out, so a burst of keystrokes produces one callback rather than one per
    /// key.
    /// </summary>
    public void StartWriteTimer(uint milliseconds)
    {
        if (_handle != 0)
        {
            Win32.SetTimer(_handle, WriteTimerId, milliseconds, 0);
        }
    }

    /// <summary>Cancels the one-shot timer if it is running.</summary>
    public void StopWriteTimer()
    {
        if (_handle != 0)
        {
            Win32.KillTimer(_handle, WriteTimerId);
        }
    }

    /// <summary>Starts a repeating timer, or changes its interval.</summary>
    public void StartPollTimer(uint milliseconds)
    {
        if (_handle != 0)
        {
            Win32.SetTimer(_handle, PollTimerId, milliseconds, 0);
        }
    }

    /// <summary>Stops the repeating timer.</summary>
    public void StopPollTimer()
    {
        if (_handle != 0)
        {
            Win32.KillTimer(_handle, PollTimerId);
        }
    }

    /// <summary>
    /// Runs an action on the thread that owns this window, which is the only thread the editor's
    /// object model and the browser may be touched from. An engine answer arrives on a pool
    /// thread, and this is its way back.
    /// </summary>
    public void RunOnHostThread(Action action)
    {
        ArgumentNullException.ThrowIfNull(action);

        if (_handle == 0)
        {
            return;
        }

        var handle = GCHandle.Alloc(action);

        if (!Win32.PostMessage(_handle, WmRunAction, 0, GCHandle.ToIntPtr(handle)))
        {
            handle.Free();
        }
    }

    /// <summary>Creates the overlay as a child of <paramref name="parent"/>.</summary>
    public static OverlayWindow? Create(nint parent, PixelRect bounds)
    {
        if (parent == 0 || !EnsureClassRegistered())
        {
            return null;
        }

        var overlay = new OverlayWindow();
        overlay._self = GCHandle.Alloc(overlay);

        var handle = Win32.CreateWindowEx(
            0,
            ClassName,
            null,
            Win32.WsChild | Win32.WsClipChildren | Win32.WsClipSiblings,
            bounds.Left,
            bounds.Top,
            bounds.Width,
            bounds.Height,
            parent,
            0,
            ShimModule.Handle,
            GCHandle.ToIntPtr(overlay._self));

        if (handle == 0)
        {
            Log.Error($"overlay: could not be created, error {Marshal.GetLastWin32Error()}");
            overlay._self.Free();
            return null;
        }

        overlay._handle = handle;
        return overlay;
    }

    /// <summary>Moves and sizes the overlay, and shows or hides it.</summary>
    public void Place(PixelRect bounds, bool visible)
    {
        if (_handle == 0)
        {
            return;
        }

        if (!visible)
        {
            if (_shown)
            {
                _shown = false;
                Win32.ShowWindow(_handle, Win32.SwHide);
            }

            return;
        }

        // Raising and moving are separated deliberately.
        //
        // The overlay has to be raised on every call, whether or not it has moved. Its siblings are
        // the code panes, and the editor puts a pane on top of its siblings whenever it activates
        // one, which happens every time the user picks a different module. Skipping the raise
        // because the rectangle had not changed let the newly activated pane cover the surface, and
        // the surface looked like it had vanished.
        //
        // Moving is the expensive half: it resizes the browser and relayouts the document. Window
        // events arrive for anything that moves anywhere in the editor and most leave the pane
        // exactly where it was, so that half is skipped when nothing changed.
        var moved = !_shown || !bounds.Equals(_placed);

        var flags = Win32.SwpNoActivate | (moved ? 0 : Win32.SwpNoMove | Win32.SwpNoSize);

        Win32.SetWindowPos(
            _handle,
            Win32.HwndTop,
            bounds.Left,
            bounds.Top,
            bounds.Width,
            bounds.Height,
            flags);

        if (moved)
        {
            _placed = bounds;
            Log.Info($"overlay: placed at {bounds.Left},{bounds.Top} {bounds.Width}x{bounds.Height} in its parent");
        }

        if (!_shown)
        {
            _shown = true;
            Win32.ShowWindow(_handle, Win32.SwShowNoActivate);
        }
    }

    private static bool EnsureClassRegistered()
    {
        lock (ClassGate)
        {
            if (_classRegistered)
            {
                return true;
            }

            fixed (char* className = ClassName)
            {
                var windowClass = new WndClassExW
                {
                    Size = (uint)sizeof(WndClassExW),
                    Style = Win32.CsHRedraw | Win32.CsVRedraw,
                    WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&WindowProc,
                    Instance = ShimModule.Handle,
                    Cursor = Win32.LoadCursor(0, Win32.IdcArrow),
                    Background = 0,
                    ClassName = className,
                };

                if (Win32.RegisterClassEx(&windowClass) != 0)
                {
                    _classRegistered = true;
                    return true;
                }

                // Registering twice in one process is not a failure; the class is already there.
                var error = Marshal.GetLastWin32Error();
                if (error == Win32.ErrorClassAlreadyExists)
                {
                    _classRegistered = true;
                    return true;
                }

                Log.Error($"overlay: the window class could not be registered, error {error}");
                return false;
            }
        }
    }

    [UnmanagedCallersOnly]
    private static nint WindowProc(nint window, uint message, nint wParam, nint lParam)
    {
        try
        {
            switch (message)
            {
                case Win32.WmNcCreate:
                {
                    // The only chance to associate the window with its managed object before any
                    // other message arrives.
                    var create = (CreateStructW*)lParam;
                    Win32.SetWindowLongPtr(window, Win32.GwlpUserData, create->CreateParams);
                    break;
                }

                case Win32.WmSize:
                {
                    var overlay = FromHandle(window);
                    if (overlay is not null)
                    {
                        overlay.Resized?.Invoke(overlay.ClientBounds());
                    }

                    return 0;
                }

                case Win32.WmTimer:
                {
                    var overlay = FromHandle(window);
                    if (overlay is null)
                    {
                        return 0;
                    }

                    if ((nuint)wParam == WriteTimerId)
                    {
                        // One shot. A window timer repeats until it is killed, and the work behind
                        // this only needs doing once per burst.
                        overlay.StopWriteTimer();
                        overlay.Elapsed?.Invoke();
                    }
                    else if ((nuint)wParam == PollTimerId)
                    {
                        overlay.Polled?.Invoke();
                    }

                    return 0;
                }

                case WmRunAction:
                {
                    var stored = GCHandle.FromIntPtr(lParam);

                    try
                    {
                        (stored.Target as Action)?.Invoke();
                    }
                    finally
                    {
                        stored.Free();
                    }

                    return 0;
                }

                case Win32.WmEraseBackground:
                    // The browser paints every pixel. Erasing first would flash.
                    return 1;

                case Win32.WmNcDestroy:
                {
                    var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
                    if (stored != 0)
                    {
                        var handle = GCHandle.FromIntPtr(stored);
                        if (handle.IsAllocated)
                        {
                            if (handle.Target is OverlayWindow overlay)
                            {
                                overlay._handle = 0;
                            }

                            handle.Free();
                        }

                        Win32.SetWindowLongPtr(window, Win32.GwlpUserData, 0);
                    }

                    break;
                }
            }
        }
        catch (Exception)
        {
            // Nothing may escape into the window procedure.
        }

        return Win32.DefWindowProc(window, message, wParam, lParam);
    }

    private static OverlayWindow? FromHandle(nint window)
    {
        var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        if (stored == 0)
        {
            return null;
        }

        var handle = GCHandle.FromIntPtr(stored);
        return handle.IsAllocated ? handle.Target as OverlayWindow : null;
    }

    public void Dispose()
    {
        var handle = _handle;

        if (handle != 0)
        {
            Win32.KillTimer(handle, WriteTimerId);
            Win32.KillTimer(handle, PollTimerId);
        }

        _handle = 0;
        Elapsed = null;
        Polled = null;

        if (handle != 0 && Win32.IsWindow(handle))
        {
            // Destroying the window frees the handle to this object from its own message handler.
            Win32.DestroyWindow(handle);
            return;
        }

        if (_self.IsAllocated)
        {
            _self.Free();
        }
    }
}
