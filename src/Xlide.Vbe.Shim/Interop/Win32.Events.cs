using System.Runtime.InteropServices;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>
/// The Win32 surface for observing the editor's own windows. The editor raises almost nothing
/// through its object model, so window events are how selection, layout, and lifetime become
/// visible to the add-in.
/// </summary>
internal static unsafe partial class Win32
{
    public const uint EventObjectCreate = 0x8000;
    public const uint EventObjectDestroy = 0x8001;
    public const uint EventObjectShow = 0x8002;
    public const uint EventObjectHide = 0x8003;
    public const uint EventObjectReorder = 0x8004;
    public const uint EventObjectFocus = 0x8005;
    public const uint EventObjectLocationChange = 0x800B;
    public const uint EventObjectNameChange = 0x800C;

    /// <summary>The event refers to the window itself rather than a child object inside it.</summary>
    public const int ObjIdWindow = 0;

    /// <summary>The event refers to the caret. Location changes with this id track the text cursor.</summary>
    public const int ObjIdCaret = -8;

    /// <summary>
    /// Deliver events asynchronously on the hooked thread's message loop. The alternative injects
    /// the callback into the thread that raised the event, which for a host we do not own is a
    /// stability risk with no benefit: we already live on the thread whose events matter.
    /// </summary>
    public const uint WinEventOutOfContext = 0x0000;

    public const uint GaRoot = 2;

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial nint SetWinEventHook(
        uint eventMin,
        uint eventMax,
        nint moduleHandle,
        nint callback,
        uint processId,
        uint threadId,
        uint flags);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool UnhookWinEvent(nint hook);

    [LibraryImport("user32.dll", EntryPoint = "GetClassNameW")]
    public static partial int GetClassName(nint window, char* buffer, int capacity);

    [LibraryImport("user32.dll")]
    public static partial nint GetAncestor(nint window, uint flags);

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial uint GetWindowThreadProcessId(nint window, uint* processId);

    [LibraryImport("user32.dll", EntryPoint = "GetWindowTextW")]
    public static partial int GetWindowText(nint window, char* buffer, int capacity);

    [LibraryImport("user32.dll", EntryPoint = "SetWindowTextW", StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SetWindowText(nint window, string text);

    /// <summary>IMAGE_ICON, and the flags for reading one out of a file at a given size.</summary>
    public const uint ImageIcon = 1;
    public const uint LrLoadFromFile = 0x0010;
    [LibraryImport("user32.dll", EntryPoint = "LoadImageW", StringMarshalling = StringMarshalling.Utf16)]
    public static partial nint LoadImage(nint instance, string name, uint type, int cx, int cy, uint load);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool DestroyIcon(nint icon);

    /// <summary>SM_CXICON / SM_CXSMICON: the sizes the shell wants for a window's two icons.</summary>
    public const int SmCxIcon = 11;
    public const int SmCxSmIcon = 49;

    [LibraryImport("user32.dll")]
    public static partial int GetSystemMetrics(int index);

    [LibraryImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool GetWindowRect(nint window, Rect* rect);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool IsWindowVisible(nint window);

    /// <summary>
    /// False while an app-modal dialog has disabled the window, which is how a dialog with no
    /// callback still announces itself: the windows it took input from say so.
    /// </summary>
    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool IsWindowEnabled(nint window);

    /// <summary>
    /// Enumerates every descendant of a window, not only its immediate children, which is why a
    /// single call over the editor's frame reaches the panes regardless of how deeply the editor
    /// nests them.
    /// </summary>
    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool EnumChildWindows(nint parent, nint callback, nint parameter);

    /// <summary>GW_HWNDNEXT. Steps to the next window at the same level in z order.</summary>
    public const uint GwHwndNext = 2;

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial nint GetTopWindow(nint parent);

    [LibraryImport("user32.dll", SetLastError = true)]
    public static partial nint GetWindow(nint window, uint command);

    /// <summary>HWND_TOP. Places a window above its siblings.</summary>
    public const nint HwndTop = 0;

    public const int SwHide = 0;

    /// <summary>SW_SHOWNOACTIVATE. Shows a window without taking focus from whatever has it.</summary>
    public const int SwShowNoActivate = 4;

    /// <summary>WM_GETICON / WM_SETICON, with the icon-size selectors they take.</summary>
    public const uint WmGetIcon = 0x007F;
    public const uint WmSetIcon = 0x0080;
    public const nint IconSmall = 0;
    public const nint IconBig = 1;
    public const nint IconSmall2 = 2;

    /// <summary>GCLP_HICON / GCLP_HICONSM: a window class's icons, the WM_GETICON fallback.</summary>
    public const int GclpHIcon = -14;
    public const int GclpHIconSm = -34;

    [LibraryImport("user32.dll", EntryPoint = "GetClassLongPtrW")]
    public static partial nint GetClassLongPtr(nint window, int index);

    /// <summary>SW_SHOW. Shows a window and activates it.</summary>
    public const int SwShow = 5;

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool ShowWindow(nint window, int command);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SetForegroundWindow(nint window);

    [LibraryImport("gdi32.dll")]
    public static partial nint CreateRectRgn(int left, int top, int right, int bottom);

    /// <summary>RGN_DIFF: the first region minus the second.</summary>
    public const int RgnDiff = 4;

    [LibraryImport("gdi32.dll")]
    public static partial int CombineRgn(nint destination, nint first, nint second, int mode);

    /// <summary>The window owns the region after a successful call; do not delete it.</summary>
    [LibraryImport("user32.dll")]
    public static partial int SetWindowRgn(nint window, nint region, [MarshalAs(UnmanagedType.Bool)] bool redraw);

    /// <summary>Region complexity answers: 0 none set, 1 empty, 2 simple, 3 complex.</summary>
    [LibraryImport("user32.dll")]
    public static partial int GetWindowRgn(nint window, nint region);

    [LibraryImport("comctl32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SetWindowSubclass(nint window, nint callback, nuint id, nuint reference);

    [LibraryImport("comctl32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool RemoveWindowSubclass(nint window, nint callback, nuint id);

    [LibraryImport("comctl32.dll")]
    public static partial nint DefSubclassProc(nint window, uint message, nint wParam, nint lParam);

    /// <summary>GWL_EXSTYLE.</summary>
    public const int GwlExStyle = -20;

    public const long WsExLayered = 0x80000;
    public const long WsExTransparent = 0x20;
    public const long WsExNoActivate = 0x8000000;

    /// <summary>LWA_ALPHA.</summary>
    public const uint LwaAlpha = 2;

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool SetLayeredWindowAttributes(nint window, uint colorKey, byte alpha, uint flags);

    /// <summary>Enumerates every top-level window on the desktop; the callback filters.</summary>
    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static partial bool EnumWindows(nint callback, nint parameter);

    [LibraryImport("user32.dll", EntryPoint = "SendMessageW")]
    public static partial nint SendMessage(nint window, uint message, nint wParam, nint lParam);
}