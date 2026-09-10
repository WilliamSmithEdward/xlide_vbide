namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// Word's save-time temp file, which a Word VBA project names as its own after every save.
///
/// Word saves a document by writing the new bytes beside it, renaming the original to
/// `~WRLnnnn.tmp`, renaming the new file into place and deleting the temp. The project's
/// FileName follows the RENAME of the original, so after a save it names a temp file that no
/// longer exists, and a different one after the next save - ~WRL0001.tmp, then ~WRL0003.tmp
/// (measured 2026-09-10, lessons.md finding 81). Read raw, that renamed the tree's project after
/// every save and restarted everything keyed by the identity under a name nothing could address.
///
/// Only the leaf is judged, and only Word's measured spelling of it. The owner file Word keeps
/// beside an open document - `~$` and the name less a letter or two - is not this, and neither
/// is any other `~WR?` temp Word writes for its own purposes, none of which has been seen in a
/// project's FileName.
/// </summary>
public static class WordSaveTemp
{
    /// <summary>True when the path's file name is `~WRL`, digits, `.tmp`, in any case.</summary>
    public static bool Matches(string? path)
    {
        if (string.IsNullOrEmpty(path))
        {
            return false;
        }

        var leaf = Path.GetFileName(path);

        // "~WRL" + at least one digit + ".tmp" is nine characters at its shortest.
        if (leaf.Length < 9
            || !leaf.StartsWith("~WRL", StringComparison.OrdinalIgnoreCase)
            || !leaf.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        foreach (var character in leaf.AsSpan(4, leaf.Length - 8))
        {
            if (!char.IsAsciiDigit(character))
            {
                return false;
            }
        }

        return true;
    }
}
