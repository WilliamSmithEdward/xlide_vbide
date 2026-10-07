using System.Runtime.InteropServices;

namespace Xlide.Vbe.Shim.Interop;

[StructLayout(LayoutKind.Sequential)]
internal struct WindowPlacement
{
    public uint Length;
    public uint Flags;
    public uint ShowCommand;
    public Point Minimum;
    public Point Maximum;
    public Rect Normal;
}

[StructLayout(LayoutKind.Sequential)]
internal struct MonitorInfo
{
    public uint Size;
    public Rect Monitor;
    public Rect Work;
    public uint Flags;
}

internal static unsafe partial class Win32
{
    public const uint WmExitSizeMove = 0x0232;

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool GetWindowPlacement(nint window, WindowPlacement* placement);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SetWindowPlacement(nint window, WindowPlacement* placement);

    [LibraryImport("user32.dll")]
    public static partial nint MonitorFromRect(Rect* rectangle, uint flags);

    [LibraryImport("user32.dll", EntryPoint = "GetMonitorInfoW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool GetMonitorInfo(nint monitor, MonitorInfo* info);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool IsIconic(nint window);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool IsZoomed(nint window);
}
