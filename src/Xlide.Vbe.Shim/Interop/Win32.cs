using System.Runtime.InteropServices;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>Win32 rectangle. Layout matches RECT exactly and is passed by pointer to the API.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct Rect
{
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}

/// <summary>Win32 POINT. Layout matches exactly and is passed by pointer to the API.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct Point
{
    public int X;
    public int Y;
}

/// <summary>Win32 SIZE, used for the scroll extent argument the in-place site passes by value.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct Size
{
    public int Width;
    public int Height;
}

/// <summary>
/// OLE SIZEL. An embedded object's extent is expressed in HIMETRIC units, which are hundredths of
/// a millimetre, not pixels.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
internal struct SizeL
{
    public int Width;
    public int Height;
}

/// <summary>Win32 window class registration block, wide character variant.</summary>
[StructLayout(LayoutKind.Sequential)]
internal unsafe struct WndClassExW
{
    public uint Size;
    public uint Style;
    public nint WindowProc;
    public int ClassExtra;
    public int WindowExtra;
    public nint Instance;
    public nint Icon;
    public nint Cursor;
    public nint Background;
    public char* MenuName;
    public char* ClassName;
    public nint SmallIcon;
}

/// <summary>
/// Creation parameters delivered with WM_NCCREATE. Only the trailing pointer matters here: it
/// carries the handle to the managed window object, which is the only chance to associate the two
/// before the first message arrives.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
internal unsafe struct CreateStructW
{
    public nint CreateParams;
    public nint Instance;
    public nint Menu;
    public nint Parent;
    public int Height;
    public int Width;
    public int Y;
    public int X;
    public int Style;
    public char* Name;
    public char* Class;
    public uint ExStyle;
}

/// <summary>Win32 message, as delivered to accelerator translation.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct Msg
{
    public nint Window;
    public uint Message;
    public nint WParam;
    public nint LParam;
    public uint Time;
    public int PointX;
    public int PointY;
}

/// <summary>
/// The Win32 surface the tool window needs. Declared with source-generated marshalling so nothing
/// is generated at run time, which is a requirement of ahead-of-time compilation.
/// </summary>
internal static unsafe partial class Win32
{
    public const int WsChild = unchecked((int)0x40000000);
    public const int WsVisible = 0x10000000;
    public const int WsClipChildren = 0x02000000;
    public const int WsClipSiblings = 0x04000000;

    public const uint CsHRedraw = 0x0002;
    public const uint CsVRedraw = 0x0001;
    public const uint CsDblClks = 0x0008;

    public const int GwlpUserData = -21;

    public const uint WmCreate = 0x0001;
    public const uint WmDestroy = 0x0002;
    public const uint WmSize = 0x0005;
    public const uint WmSetFocus = 0x0007;
    public const uint WmEraseBackground = 0x0014;
    public const uint WmNcCreate = 0x0081;
    public const uint WmNcDestroy = 0x0082;

    public const uint SwpNoSize = 0x0001;
    public const uint SwpNoMove = 0x0002;
    public const uint SwpNoZOrder = 0x0004;
    public const uint SwpNoActivate = 0x0010;

    /// <summary>COLOR_BTNFACE + 1, the value a class background brush takes.</summary>
    public const nint ColorButtonFace = 16;

    /// <summary>IDC_ARROW.</summary>
    public const int IdcArrow = 32512;

    public const int ErrorClassAlreadyExists = 1410;

    public const uint GetModuleHandleFromAddress = 0x00000004;
    public const uint GetModuleHandleUnchangedRefCount = 0x00000002;

    [LibraryImport("user32.dll", SetLastError = true, EntryPoint = "RegisterClassExW")]
    public static partial ushort RegisterClassEx(WndClassExW* windowClass);

    [LibraryImport("user32.dll", SetLastError = true, EntryPoint = "CreateWindowExW", StringMarshalling = StringMarshalling.Utf16)]
    public static partial nint CreateWindowEx(
        uint exStyle,
        string className,
        string? windowName,
        int style,
        int x,
        int y,
        int width,
        int height,
        nint parent,
        nint menu,
        nint instance,
        nint param);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool DestroyWindow(nint window);

    [LibraryImport("user32.dll", EntryPoint = "DefWindowProcW")]
    public static partial nint DefWindowProc(nint window, uint message, nint wParam, nint lParam);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SetWindowPos(
        nint window,
        nint insertAfter,
        int x,
        int y,
        int width,
        int height,
        uint flags);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool GetClientRect(nint window, Rect* rect);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool IsWindow(nint window);

    /// <summary>
    /// Converts a screen rectangle into a window's client coordinates, which is the space a child
    /// of that window is positioned in.
    ///
    /// Two points are mapped rather than one, because this also has to survive a right-to-left
    /// layout: mapping the corners lets the caller normalise, where mapping an origin and adding a
    /// size would place the rectangle on the wrong side.
    /// </summary>
    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial int MapWindowPoints(nint from, nint to, Point* points, uint count);

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial nint GetParent(nint window);

    /// <summary>
    /// State of a key at the point the message being processed was posted, which is what a key
    /// handler needs: asking for the state now would report where the modifier is by the time the
    /// message is dealt with, not where it was when the key was struck.
    /// </summary>
    [LibraryImport("user32.dll")]
    public static partial short GetKeyState(int key);

    /// <summary>
    /// A timer that fires on the thread that owns the window.
    ///
    /// This is why it is a window timer rather than a managed one. Everything on the editor's
    /// object model is apartment bound, so a callback on a pool thread cannot touch it, and
    /// marshalling back would need a pump this code does not own. A window timer arrives through
    /// the message loop the host is already running, on the only thread allowed to do the work.
    /// </summary>
    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial nuint SetTimer(nint window, nuint id, uint milliseconds, nint callback);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool KillTimer(nint window, nuint id);

    public const uint WmTimer = 0x0113;

    public const int VkShift = 0x10;
    public const int VkControl = 0x11;

    /// <summary>Mask for the high-order bit of a key state, which is set while the key is down.</summary>
    public const short KeyDownMask = unchecked((short)0x8000);

    [LibraryImport("user32.dll", EntryPoint = "FindWindowExW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    public static partial nint FindWindowEx(nint parent, nint childAfter, string? className, string? windowName);

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial uint GetWindowThreadProcessId(nint window, out uint processId);

    [LibraryImport("kernel32.dll")]
    public static partial uint GetCurrentProcessId();

    [LibraryImport("ole32.dll")]
    public static partial int CoCreateInstance(
        in Guid classId,
        nint outer,
        uint context,
        in Guid interfaceId,
        out nint instance);

    /// <summary>CLSCTX_INPROC_SERVER.</summary>
    public const uint ClassContextInProcessServer = 1;

    /// <summary>
    /// Asks a window for an object behind it. With the native object model identifier this is how
    /// a host application's own automation object is reached without the running object table.
    /// </summary>
    [LibraryImport("oleacc.dll")]
    public static partial int AccessibleObjectFromWindow(nint window, uint objectId, in Guid interfaceId, out nint result);

    [LibraryImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    public static partial nint SetWindowLongPtr(nint window, int index, nint value);

    [LibraryImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    public static partial nint GetWindowLongPtr(nint window, int index);

    [LibraryImport("user32.dll", EntryPoint = "LoadCursorW", SetLastError = true)]
    public static partial nint LoadCursor(nint instance, nint cursorName);

    [LibraryImport("kernel32.dll", EntryPoint = "GetModuleHandleExW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool GetModuleHandleEx(uint flags, nint address, nint* module);

    [LibraryImport("kernel32.dll", EntryPoint = "GetModuleFileNameW", SetLastError = true)]
    public static partial uint GetModuleFileName(nint module, char* fileName, uint size);
}
