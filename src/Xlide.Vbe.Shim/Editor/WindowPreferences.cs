using System.Text.Json;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>The same window preferences in every Office host. WINDOWPLACEMENT keeps normal
/// bounds while maximized; its workspace coordinates go back through that same Win32 API.</summary>
internal sealed unsafe class WindowPreferences(UserStateStore store, string key)
{
    private bool _restored;
    private bool _restoring;
    private string? _saved;
    private bool _wasMaximized;

    public void Restore(nint window)
    {
        if (_restored || window == 0) return;
        _restoring = true;
        try
        {
            var json = store.ReadValue(UserStateStore.UiFile, key);
            var geometry = json is null ? null
                : JsonSerializer.Deserialize(json, WindowGeometryContext.Default.WindowGeometry);
            if (geometry is null || geometry.Width <= 0 || geometry.Height <= 0
                || geometry.Width > 100000 || geometry.Height > 100000
                || Math.Abs((long)geometry.Left) > 100000 || Math.Abs((long)geometry.Top) > 100000) return;

            var rectangle = new Rect { Left = geometry.Left, Top = geometry.Top,
                Right = geometry.Left + geometry.Width, Bottom = geometry.Top + geometry.Height };
            var monitor = new MonitorInfo { Size = (uint)sizeof(MonitorInfo) };
            if (!Win32.GetMonitorInfo(Win32.MonitorFromRect(&rectangle, 2), &monitor)) return;
            var x = monitor.Work.Left - monitor.Monitor.Left;
            var y = monitor.Work.Top - monitor.Monitor.Top;
            var fitted = geometry.Fit(new PixelRect(monitor.Work.Left - x, monitor.Work.Top - y,
                monitor.Work.Right - x, monitor.Work.Bottom - y));
            var placement = new WindowPlacement { Length = (uint)sizeof(WindowPlacement),
                ShowCommand = fitted.Maximized ? 3u : 1u,
                Normal = new Rect { Left = fitted.Left, Top = fitted.Top,
                    Right = fitted.Left + fitted.Width, Bottom = fitted.Top + fitted.Height } };
            if (!Win32.SetWindowPlacement(window, &placement))
                Log.Warn($"preferences: {key} placement was refused");
        }
        catch (Exception ex) { Log.Error($"preferences: {key} could not be restored", ex); }
        finally { _restoring = false; _restored = true; _wasMaximized = Win32.IsZoomed(window); }
    }

    public void CaptureStateChange(nint window)
    {
        if (Win32.IsIconic(window)) return;
        var maximized = Win32.IsZoomed(window);
        if (_wasMaximized != maximized) Capture(window);
        _wasMaximized = maximized;
    }

    public void Capture(nint window)
    {
        if (!_restored || _restoring || window == 0 || !Win32.IsWindowVisible(window) || Win32.IsIconic(window)) return;
        try
        {
            var placement = new WindowPlacement { Length = (uint)sizeof(WindowPlacement) };
            if (!Win32.GetWindowPlacement(window, &placement)) return;
            var rect = placement.Normal;
            if (rect.Right <= rect.Left || rect.Bottom <= rect.Top) return;
            var geometry = new WindowGeometry(rect.Left, rect.Top, rect.Right - rect.Left,
                rect.Bottom - rect.Top, Win32.IsZoomed(window));
            var json = JsonSerializer.Serialize(geometry, WindowGeometryContext.Default.WindowGeometry);
            if (_saved == json) return;
            store.WriteValue(UserStateStore.UiFile, key, json);
            _saved = json;
        }
        catch (Exception ex) { Log.Error($"preferences: {key} could not be saved", ex); }
    }
}
