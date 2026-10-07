using System.Text.Json;
using System.Text.Json.Serialization;
using Xlide.Vbe.Core.Hosting;

namespace Xlide.Vbe.Core.Editor;

/// <summary>Shared XLIDE normal window bounds in Win32 workspace coordinates. Minimized windows
/// never replace them; the native reader and writer both use WINDOWPLACEMENT.</summary>
public sealed record WindowGeometry(int Left, int Top, int Width, int Height, bool Maximized = false)
{
    public WindowGeometry Fit(PixelRect workArea)
    {
        var width = Math.Clamp(Width, Math.Min(400, workArea.Width), workArea.Width);
        var height = Math.Clamp(Height, Math.Min(300, workArea.Height), workArea.Height);
        return this with
        {
            Width = width,
            Height = height,
            Left = Math.Clamp(Left, workArea.Left, workArea.Left + workArea.Width - width),
            Top = Math.Clamp(Top, workArea.Top, workArea.Top + workArea.Height - height),
        };
    }
}

[JsonSerializable(typeof(WindowGeometry))]
public sealed partial class WindowGeometryContext : JsonSerializerContext;
