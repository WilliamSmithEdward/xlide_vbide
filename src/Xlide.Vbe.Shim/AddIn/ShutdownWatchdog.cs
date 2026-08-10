using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// A message-only window that outlives the session to notice a shutdown that never finished.
///
/// OnBeginShutdown stops the session before the host shows its save prompt, and when the
/// developer cancels there, no callback says so: the host keeps running and the add-in is a
/// corpse. This window is the one thing deliberately left alive. Its timer keeps firing on the
/// host thread - which only still pumps if the shutdown was in fact abandoned - and whoever
/// listens can look around and stand the session up again.
/// </summary>
internal sealed unsafe class ShutdownWatchdog : IDisposable
{
    private const string ClassName = "XlideShutdownWatchdog";
    private const nuint TickTimerId = 1;
    private const uint TickMilliseconds = 1500;

    private static bool _classRegistered;
    private static readonly Lock ClassGate = new();

    private nint _handle;
    private GCHandle _self;

    /// <summary>Raised on the host thread at every tick, for as long as the watchdog lives.</summary>
    public Action? Ticked { get; set; }

    private ShutdownWatchdog()
    {
    }

    public static ShutdownWatchdog? Create(Action ticked)
    {
        ArgumentNullException.ThrowIfNull(ticked);

        if (!EnsureClassRegistered())
        {
            return null;
        }

        var watchdog = new ShutdownWatchdog { Ticked = ticked };
        watchdog._self = GCHandle.Alloc(watchdog);

        var handle = Win32.CreateWindowEx(
            0,
            ClassName,
            null,
            0,
            0,
            0,
            0,
            0,
            Win32.HwndMessage,
            0,
            ShimModule.Handle,
            GCHandle.ToIntPtr(watchdog._self));

        if (handle == 0)
        {
            Log.Error($"watchdog: could not be created, error {Marshal.GetLastWin32Error()}");
            watchdog._self.Free();
            return null;
        }

        watchdog._handle = handle;
        Win32.SetTimer(handle, TickTimerId, TickMilliseconds, 0);
        return watchdog;
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
                    WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&WindowProc,
                    Instance = ShimModule.Handle,
                    ClassName = className,
                };

                if (Win32.RegisterClassEx(&windowClass) != 0
                    || Marshal.GetLastWin32Error() == Win32.ErrorClassAlreadyExists)
                {
                    _classRegistered = true;
                    return true;
                }

                Log.Error($"watchdog: the window class could not be registered, error {Marshal.GetLastWin32Error()}");
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

                case Win32.WmTimer:
                {
                    var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
                    if (stored != 0
                        && GCHandle.FromIntPtr(stored) is { IsAllocated: true } handle
                        && handle.Target is ShutdownWatchdog watchdog)
                    {
                        watchdog.Ticked?.Invoke();
                    }

                    return 0;
                }

                case Win32.WmNcDestroy:
                {
                    var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
                    if (stored != 0)
                    {
                        var handle = GCHandle.FromIntPtr(stored);
                        if (handle.IsAllocated)
                        {
                            if (handle.Target is ShutdownWatchdog watchdog)
                            {
                                watchdog._handle = 0;
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

    public void Dispose()
    {
        var handle = _handle;
        _handle = 0;
        Ticked = null;

        if (handle != 0 && Win32.IsWindow(handle))
        {
            Win32.KillTimer(handle, TickTimerId);
            Win32.DestroyWindow(handle);
            return;
        }

        if (_self.IsAllocated)
        {
            _self.Free();
        }
    }
}
