using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>
/// The system's own "choose a folder" window.
///
/// The page cannot raise one for itself: a browser will offer a file, and only with a gesture it
/// trusts, and neither gets you a folder path. So the dialog asks the host for one and the host
/// shows the real thing, which is also the one a developer already knows how to drive — it has
/// their pinned places, their recent folders, and the address bar they can paste into.
///
/// This is the older SHBrowseForFolder rather than the Common Item Dialog, deliberately. The modern
/// one is a COM object with three interfaces to declare and an event sink to implement to preselect
/// a folder, all of it hand-written under ahead-of-time compilation where nothing can be generated
/// at run time. This is two calls and a callback, it honours a starting folder, and the window it
/// shows is the same shell control underneath.
/// </summary>
internal static partial class FolderPicker
{
    /// <summary>Show an edit box, and the newer visual style rather than the Windows 95 tree.</summary>
    private const uint BrowseNewDialogStyle = 0x0040;

    private const uint BrowseReturnOnlyFileSystemDirectories = 0x0001;

    /// <summary>Sent as the dialog comes up, which is the only chance to select a starting folder.</summary>
    private const uint OnInitialised = 1;

    /// <summary>Tells the dialog to select a folder, given as a string rather than an id list.</summary>
    private const uint SelectPathMessage = 0x0467;

    /// <summary>
    /// Asks the developer for a folder. Answers null when they close it without choosing one.
    /// </summary>
    /// <param name="owner">The window it belongs to, so it is modal to the editor rather than lost behind it.</param>
    /// <param name="title">The line above the tree, which is where the direction is said.</param>
    /// <param name="startAt">The folder to open on, when there is a sensible one.</param>
    public static string? Choose(nint owner, string title, string? startAt)
    {
        // The callback needs the starting folder and cannot be handed state, so it is held here for
        // the life of the call. The dialog is modal, so only one can be up at a time.
        _startAt = string.IsNullOrWhiteSpace(startAt) ? null : startAt;

        var info = new BrowseInfo
        {
            Owner = owner,
            Title = title,
            Flags = BrowseNewDialogStyle | BrowseReturnOnlyFileSystemDirectories,
            Callback = _startAt is null ? null : OnBrowseEvent,
        };

        var idList = nint.Zero;
        try
        {
            idList = SHBrowseForFolder(ref info);
            if (idList == nint.Zero)
            {
                return null;
            }

            var path = new char[260];
            return SHGetPathFromIDList(idList, path) ? new string(path).TrimEnd('\0') : null;
        }
        catch (Exception ex)
        {
            Log.Error("folder: the chooser could not be shown", ex);
            return null;
        }
        finally
        {
            if (idList != nint.Zero)
            {
                // The id list is the shell's allocation and is freed with the shell's allocator.
                CoTaskMemFree(idList);
            }

            _startAt = null;
        }
    }

    private static string? _startAt;

    private static int OnBrowseEvent(nint dialog, uint message, nint parameter, nint data)
    {
        if (message == OnInitialised && _startAt is not null)
        {
            SendMessage(dialog, SelectPathMessage, 1, _startAt);
        }

        return 0;
    }

    private delegate int BrowseCallback(nint dialog, uint message, nint parameter, nint data);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct BrowseInfo
    {
        public nint Owner;
        public nint Root;
        public nint DisplayNameBuffer;
        [MarshalAs(UnmanagedType.LPWStr)] public string Title;
        public uint Flags;
        [MarshalAs(UnmanagedType.FunctionPtr)] public BrowseCallback? Callback;
        public nint CallbackData;
        public int Image;
    }

    // Marshalled the old way rather than with LibraryImport: the structure carries a string and a
    // function pointer, which the source generator will not marshal for us.
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "SHBrowseForFolderW")]
    private static extern nint SHBrowseForFolder(ref BrowseInfo info);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "SHGetPathFromIDListW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SHGetPathFromIDList(nint idList, [Out] char[] path);

    [LibraryImport("ole32.dll")]
    private static partial void CoTaskMemFree(nint memory);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
    private static extern nint SendMessage(nint window, uint message, nint wParam, string lParam);
}
