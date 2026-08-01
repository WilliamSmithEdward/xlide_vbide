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

    public const uint GaParent = 1;
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
}