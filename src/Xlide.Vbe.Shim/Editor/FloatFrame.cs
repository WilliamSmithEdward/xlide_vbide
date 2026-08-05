using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// A top-level frame of our own that adopts a native child window.
///
/// Built for the Object Browser, which the editor cannot float itself: it is an MDI document
/// window, not a dockable tool window, and LinkedWindows.Remove answers it with a silent
/// shrug (measured 2026-08-05 — the call succeeds and nothing moves). Adopted, the native
/// window is reparented into this frame and stripped of its own caption and sizing border, so
/// it fills a real window that can be moved, sized, and closed beside the surface — never on
/// it. Closing returns the child to the document area, styles restored, for its next summons.
///
/// The frame is OWNED by the editor's frame rather than parented to it, which is what keeps
/// it above the editor the way a palette sits, without living inside its client area.
/// </summary>
internal sealed unsafe class FloatFrame : IDisposable
{
    private const string ClassName = "XlideFloatFrame";

    private static readonly object ClassGate = new();
    private static bool _classRegistered;

    private GCHandle _self;
    private nint _handle;
    private nint _child;
    private long _childStyle;
    private nint _restoreParent;

    /// <summary>Called when the developer closes the frame; the owner returns the child.</summary>
    private Action? _closed;

    /// <summary>Closes arriving before this tick are the editor's reflex, not the developer's.</summary>
    private long _ignoreCloseUntil;

    /// <summary>
    /// Swallows any close request for a moment. Hiding the adopted window in the object model
    /// makes the editor send WM_CLOSE to the window's root — this frame — as a reflex; the
    /// developer's own close comes seconds later, not milliseconds.
    /// </summary>
    public void IgnoreCloseBriefly(uint milliseconds) =>
        _ignoreCloseUntil = Environment.TickCount64 + milliseconds;

    private FloatFrame()
    {
    }

    public nint Handle => _handle;

    /// <summary>
    /// Creates the frame over <paramref name="owner"/> and adopts <paramref name="child"/> into
    /// it. Null when the frame cannot be created; the child is untouched then.
    /// </summary>
    public static FloatFrame? Adopt(nint owner, nint child, string title, Action closed)
    {
        if (owner == 0 || child == 0 || !EnsureClassRegistered())
        {
            return null;
        }

        var frame = new FloatFrame();
        frame._self = GCHandle.Alloc(frame);
        frame._closed = closed;
        frame._child = child;
        frame._restoreParent = Win32.GetParent(child);

        // Centred on the owner, sized for browsing: the editor keeps no float rectangle for a
        // window it does not believe can float.
        Rect ownerRect;
        Win32.GetWindowRect(owner, &ownerRect);
        var width = Math.Min(940, Math.Max(560, ownerRect.Right - ownerRect.Left - 160));
        var height = Math.Min(700, Math.Max(420, ownerRect.Bottom - ownerRect.Top - 160));
        var left = ownerRect.Left + ((ownerRect.Right - ownerRect.Left) - width) / 2;
        var top = ownerRect.Top + ((ownerRect.Bottom - ownerRect.Top) - height) / 2;

        var handle = Win32.CreateWindowEx(
            0,
            ClassName,
            title,
            Win32.WsOverlappedWindow | Win32.WsClipChildren,
            left,
            top,
            width,
            height,
            owner,
            0,
            ShimModule.Handle,
            GCHandle.ToIntPtr(frame._self));

        if (handle == 0)
        {
            frame._self.Free();
            return null;
        }

        frame._handle = handle;

        // The child keeps its own caption and sizing border from its MDI life; inside this
        // frame they would draw a second title bar. Stripped here, restored at release.
        frame._childStyle = (long)Win32.GetWindowLongPtr(child, Win32.GwlStyle);
        var stripped = frame._childStyle
            & ~(long)(Win32.WsCaption | Win32.WsThickFrame | Win32.WsSysMenu
                      | Win32.WsMinimizeBox | Win32.WsMaximizeBox);
        Win32.SetWindowLongPtr(child, Win32.GwlStyle, (nint)stripped);

        Win32.SetParent(child, handle);
        frame.FitChild();
        Win32.SetWindowPos(child, 0, 0, 0, 0, 0,
            Win32.SwpNoMove | Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate | Win32.SwpFrameChanged);

        Win32.ShowWindow(handle, Win32.SwShowNoActivate);
        Log.Info($"float frame: adopted {child:X} as '{title}'");
        return frame;
    }

    /// <summary>Brings the frame to the top of its band, for a summons while already open.</summary>
    public void Present()
    {
        if (_handle != 0)
        {
            Win32.SetWindowPos(_handle, Win32.HwndTop, 0, 0, 0, 0,
                Win32.SwpNoMove | Win32.SwpNoSize | Win32.SwpNoActivate);
        }
    }

    /// <summary>
    /// Returns the child to the parent it came from, styles restored, and destroys the frame.
    /// The caller decides what the child's visibility should be once it is home.
    /// </summary>
    public void Release()
    {
        var child = _child;
        _child = 0;

        if (child != 0 && Win32.IsWindow(child))
        {
            Win32.SetWindowLongPtr(child, Win32.GwlStyle, (nint)_childStyle);
            Win32.SetParent(child, _restoreParent);
            Win32.SetWindowPos(child, 0, 0, 0, 0, 0,
                Win32.SwpNoMove | Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate | Win32.SwpFrameChanged);
        }

        Dispose();
    }

    /// <summary>Sizes the child to the frame's client area.</summary>
    private void FitChild()
    {
        if (_child == 0 || _handle == 0)
        {
            return;
        }

        Rect client;
        if (Win32.GetClientRect(_handle, &client))
        {
            Win32.SetWindowPos(_child, 0, 0, 0, client.Right - client.Left, client.Bottom - client.Top,
                Win32.SwpNoZOrder | Win32.SwpNoActivate);
        }
    }

    private static FloatFrame? FromHandle(nint window)
    {
        var data = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        return data != 0 ? GCHandle.FromIntPtr(data).Target as FloatFrame : null;
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
                    Style = 0,
                    WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&WindowProc,
                    Instance = ShimModule.Handle,
                    Cursor = Win32.LoadCursor(0, Win32.IdcArrow),
                    Background = 0,
                    ClassName = className,
                };

                if (Win32.RegisterClassEx(&windowClass) != 0
                    || Marshal.GetLastWin32Error() == Win32.ErrorClassAlreadyExists)
                {
                    _classRegistered = true;
                    return true;
                }

                Log.Error($"float frame: the window class could not be registered, error {Marshal.GetLastWin32Error()}");
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
                    var create = (CreateStructW*)lParam;
                    Win32.SetWindowLongPtr(window, Win32.GwlpUserData, create->CreateParams);
                    break;
                }

                case Win32.WmSize:
                    FromHandle(window)?.FitChild();
                    return 0;

                case Win32.WmClose:
                {
                    var frame = FromHandle(window);
                    if (frame is not null && Environment.TickCount64 < frame._ignoreCloseUntil)
                    {
                        Log.Verbose("float frame: swallowed the editor's reflex close");
                        return 0;
                    }

                    // The owner returns the child and destroys the frame; the close box only
                    // says the developer is done.
                    frame?._closed?.Invoke();
                    return 0;
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error("float frame: the window procedure failed", ex);
        }

        return Win32.DefWindowProc(window, message, wParam, lParam);
    }

    public void Dispose()
    {
        var handle = _handle;
        _handle = 0;

        if (handle != 0)
        {
            Win32.DestroyWindow(handle);
        }

        if (_self.IsAllocated)
        {
            _self.Free();
        }
    }
}
