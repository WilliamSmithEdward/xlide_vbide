using System.Runtime.InteropServices;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.ToolWindow;

/// <summary>
/// The child window the container gives us a place for, and everything the browser is layered on
/// top of.
///
/// It does almost nothing itself. It exists because the browser needs a window to be parented to
/// that we control, rather than one the container reparents, resizes, and destroys on its own
/// schedule. Owning the window means the browser's bounds are ours to keep correct.
/// </summary>
internal sealed unsafe class HostWindow : IDisposable
{
    private const string ClassName = "XlideToolWindowHost";

    private static bool _classRegistered;

    private nint _handle;
    private GCHandle _self;

    private HostWindow()
    {
    }

    /// <summary>Handle of the child window, or zero once it has been destroyed.</summary>
    public nint Handle => _handle;

    /// <summary>Raised when the client area changes size, in client coordinates.</summary>
    public Action<PixelRect>? Resized { get; set; }

    /// <summary>Raised when the window takes focus, so the browser can be given it.</summary>
    public Action? FocusReceived { get; set; }

    /// <summary>
    /// Raised when the window is going away. The container is free to destroy the parent without
    /// deactivating us first, and a browser left parented to a destroyed window keeps its processes
    /// alive for the life of the host.
    /// </summary>
    public Action? Destroying { get; set; }

    /// <summary>
    /// Creates the child window inside the container's window.
    ///
    /// The styles are not decoration. WS_CLIPCHILDREN keeps this window from painting over the
    /// browser's own child windows, and WS_CLIPSIBLINGS keeps it from painting over whatever else
    /// the container has parented alongside it. Without either, a resize produces flicker that
    /// looks like a rendering bug in the browser.
    /// </summary>
    public static HostWindow? Create(nint parent, PixelRect bounds)
    {
        if (parent == 0 || !Win32.IsWindow(parent))
        {
            Log.Error("tool window: the container supplied no usable parent window");
            return null;
        }

        if (!EnsureClassRegistered())
        {
            return null;
        }

        var window = new HostWindow();
        window._self = GCHandle.Alloc(window);

        var handle = Win32.CreateWindowEx(
            0,
            ClassName,
            null,
            Win32.WsChild | Win32.WsVisible | Win32.WsClipChildren | Win32.WsClipSiblings,
            bounds.Left,
            bounds.Top,
            bounds.Width,
            bounds.Height,
            parent,
            0,
            ShimModule.Handle,
            GCHandle.ToIntPtr(window._self));

        if (handle == 0)
        {
            var error = Marshal.GetLastWin32Error();
            Log.Error($"tool window: CreateWindowEx failed with error {error}");
            window._self.Free();
            return null;
        }

        // The window procedure stores the handle on WM_NCCREATE, but only if that message was
        // dispatched to us. Assigning here as well makes the object usable either way.
        window._handle = handle;
        Log.Info($"tool window: created child window 0x{handle:X} under parent 0x{parent:X}, {bounds.Width}x{bounds.Height}");
        return window;
    }

    /// <summary>Moves and sizes the window inside its parent.</summary>
    public void SetBounds(PixelRect bounds)
    {
        if (_handle == 0)
        {
            return;
        }

        Win32.SetWindowPos(
            _handle,
            0,
            bounds.Left,
            bounds.Top,
            bounds.Width,
            bounds.Height,
            Win32.SwpNoZOrder | Win32.SwpNoActivate);
    }

    /// <summary>Current client area, which is what the browser is sized to.</summary>
    public PixelRect ClientBounds()
    {
        if (_handle == 0)
        {
            return default;
        }

        Rect rect;
        return Win32.GetClientRect(_handle, &rect)
            ? new PixelRect(rect.Left, rect.Top, rect.Right, rect.Bottom)
            : default;
    }

    public void Dispose()
    {
        var handle = _handle;
        _handle = 0;

        if (handle != 0 && Win32.IsWindow(handle))
        {
            // Destroying the window drives WM_NCDESTROY, which frees the handle to this object.
            Win32.DestroyWindow(handle);
            return;
        }

        if (_self.IsAllocated)
        {
            _self.Free();
        }
    }

    private static bool EnsureClassRegistered()
    {
        if (_classRegistered)
        {
            return true;
        }

        fixed (char* className = ClassName)
        {
            var description = default(WndClassExW);
            description.Size = (uint)sizeof(WndClassExW);
            description.Style = Win32.CsHRedraw | Win32.CsVRedraw | Win32.CsDblClks;
            description.WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&WindowProc;
            description.Instance = ShimModule.Handle;
            description.Cursor = Win32.LoadCursor(0, Win32.IdcArrow);
            description.Background = Win32.ColorButtonFace;
            description.ClassName = className;

            if (Win32.RegisterClassEx(&description) != 0)
            {
                _classRegistered = true;
                return true;
            }

            // A second registration of the same class in the same process is expected if the
            // library is ever asked for two controls before the first one registered.
            var error = Marshal.GetLastWin32Error();
            if (error == Win32.ErrorClassAlreadyExists)
            {
                _classRegistered = true;
                return true;
            }

            Log.Error($"tool window: RegisterClassEx failed with error {error}");
            return false;
        }
    }

    /// <summary>
    /// The window procedure. It runs on the host user interface thread and is entered before any
    /// managed code of ours has a chance to run for a given window, so it must tolerate not yet
    /// knowing which object the window belongs to.
    /// </summary>
    [UnmanagedCallersOnly]
    private static nint WindowProc(nint window, uint message, nint wParam, nint lParam)
    {
        try
        {
            if (message == Win32.WmNcCreate)
            {
                var create = (CreateStructW*)lParam;
                if (create is not null && create->CreateParams != 0)
                {
                    Win32.SetWindowLongPtr(window, Win32.GwlpUserData, create->CreateParams);

                    var target = FromHandle(window);
                    if (target is not null)
                    {
                        target._handle = window;
                    }
                }

                return Win32.DefWindowProc(window, message, wParam, lParam);
            }

            var host = FromHandle(window);

            switch (message)
            {
                case Win32.WmSize when host is not null:
                    host.Resized?.Invoke(host.ClientBounds());
                    return 0;

                case Win32.WmSetFocus when host is not null:
                    host.FocusReceived?.Invoke();
                    return 0;

                case Win32.WmEraseBackground when host is not null:
                    // The browser covers the whole client area. Erasing underneath it only produces
                    // a flash of the class background on every resize.
                    return 1;

                case Win32.WmDestroy when host is not null:
                    // The browser is torn down before the window it lives in disappears, not after.
                    host.Destroying?.Invoke();
                    host._handle = 0;
                    host.Resized = null;
                    host.FocusReceived = null;
                    host.Destroying = null;
                    break;

                case Win32.WmNcDestroy:
                    var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
                    Win32.SetWindowLongPtr(window, Win32.GwlpUserData, 0);
                    if (stored != 0)
                    {
                        var handle = GCHandle.FromIntPtr(stored);
                        if (handle.IsAllocated)
                        {
                            if (handle.Target is HostWindow owner)
                            {
                                owner._handle = 0;
                                owner._self = default;
                            }

                            handle.Free();
                        }
                    }

                    break;
            }

            return Win32.DefWindowProc(window, message, wParam, lParam);
        }
        catch (Exception ex)
        {
            // An exception crossing back into the window manager terminates the host. Swallow it,
            // record it, and let the default procedure produce a sane answer.
            Log.Error($"tool window: window procedure failed on message 0x{message:X}", ex);
            return Win32.DefWindowProc(window, message, wParam, lParam);
        }
    }

    private static HostWindow? FromHandle(nint window)
    {
        var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        if (stored == 0)
        {
            return null;
        }

        var handle = GCHandle.FromIntPtr(stored);
        return handle.IsAllocated ? handle.Target as HostWindow : null;
    }
}
