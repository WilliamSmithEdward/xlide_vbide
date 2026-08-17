using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>
/// The system's own "choose a file" window, for the same reason the folder one exists next door:
/// a page can offer a file's BYTES, only behind a gesture the browser trusts, and never its path.
/// The host shows the real dialog instead - the one with the developer's pinned places, their
/// recent folders and an address bar - and hands back a path, which is what the picture loader
/// takes.
///
/// GetOpenFileName rather than the Common Item Dialog, the same call the folder picker's note
/// argues for: the modern one is three COM interfaces to hand-write under ahead-of-time
/// compilation, and this is one call with a structure. The window is the same shell control.
/// </summary>
internal static unsafe partial class FilePicker
{
    private const uint MustExist = 0x00001000 | 0x00000800; // FILEMUSTEXIST | PATHMUSTEXIST
    private const uint HideReadOnly = 0x00000004;
    private const uint ExplorerStyle = 0x00080000;

    /// <summary>
    /// Asks the developer for a file. Answers null when they close the dialog without choosing.
    /// </summary>
    /// <param name="owner">The window it belongs to, so it is modal to the editor.</param>
    /// <param name="title">The dialog's caption, which is where the purpose is said.</param>
    /// <param name="filter">
    /// The kinds offered, as pairs of description and pattern - `[("Pictures", "*.bmp;*.png")]`.
    /// Written on the wire as the double-null-terminated run the API wants, which is the one
    /// piece of this that has to be built by hand.
    /// </param>
    /// <param name="startAt">A file or folder to open on, when there is a sensible one.</param>
    public static string? Choose(
        nint owner, string title, IReadOnlyList<(string What, string Pattern)> filter, string? startAt)
    {
        var patterns = string.Concat(filter.Select(one => $"{one.What} ({one.Pattern})\0{one.Pattern}\0")) + "\0";

        // The dialog writes the chosen path into this buffer, so it is sized for the longest
        // path Windows will hand back rather than for the longest anyone expects.
        var chosen = new char[1024];

        var directory = startAt is { Length: > 0 }
            ? (Directory.Exists(startAt) ? startAt : Path.GetDirectoryName(startAt))
            : null;

        try
        {
            fixed (char* patternsAt = patterns)
            fixed (char* chosenAt = chosen)
            fixed (char* titleAt = title)
            fixed (char* directoryAt = directory)
            {
                var request = new OpenFileName
                {
                    Size = sizeof(OpenFileName),
                    Owner = owner,
                    Filter = (nint)patternsAt,
                    FilterIndex = 1,
                    File = (nint)chosenAt,
                    MaxFile = (uint)chosen.Length,
                    InitialDirectory = directory is null ? 0 : (nint)directoryAt,
                    Title = (nint)titleAt,
                    Flags = MustExist | HideReadOnly | ExplorerStyle,
                };

                if (!GetOpenFileNameW(ref request))
                {
                    return null;
                }
            }

            var path = new string(chosen).TrimEnd('\0');
            return path.Length == 0 ? null : path;
        }
        catch (Exception ex)
        {
            Log.Error("file: the chooser could not be shown", ex);
            return null;
        }
    }

    /// <summary>OPENFILENAMEW. Every pointer is passed as one and pinned by the caller, which is
    /// what keeps this a blittable structure the source generator will marshal.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct OpenFileName
    {
        public int Size;
        public nint Owner;
        public nint Instance;
        public nint Filter;
        public nint CustomFilter;
        public uint MaxCustomFilter;
        public uint FilterIndex;
        public nint File;
        public uint MaxFile;
        public nint FileTitle;
        public uint MaxFileTitle;
        public nint InitialDirectory;
        public nint Title;
        public uint Flags;
        public ushort FileOffset;
        public ushort FileExtension;
        public nint DefaultExtension;
        public nint CustomData;
        public nint Hook;
        public nint Template;
        public nint Reserved;
        public uint ReservedToo;
        public uint FlagsEx;
    }

    [LibraryImport("comdlg32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetOpenFileNameW(ref OpenFileName request);
}
