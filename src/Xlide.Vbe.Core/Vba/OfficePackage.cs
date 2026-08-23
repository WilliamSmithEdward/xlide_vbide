using System.IO.Compression;

namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// The VBA project out of a saved Office document, as bytes.
///
/// An `.xlsm` is a ZIP, and the whole VBA project - every module's source, its attribute header,
/// and every form's design - is one entry inside it. Reading it costs no COM, no export, no
/// running application and nothing written to disk, which is what makes it usable from a pass
/// that runs while the developer types.
///
/// WHAT THIS IS NOT is the live project. It is what was last SAVED, so a module added since is
/// absent from it, and a caller must treat absence as "no answer" rather than as a negative one.
/// </summary>
public static class OfficePackage
{
    /// <summary>Where each application keeps the project inside its package. Excel first, because
    /// that is the overwhelming case and the entry lookup is a dictionary hit either way.</summary>
    private static readonly string[] ProjectPaths =
    [
        "xl/vbaProject.bin",
        "word/vbaProject.bin",
        "ppt/vbaProject.bin",
        "visio/vbaProject.bin",
    ];

    /// <summary>A ceiling on the project stream, so a hostile or corrupt package cannot ask this
    /// to allocate the machine. Real projects are kilobytes; a form full of pictures is a few
    /// megabytes.</summary>
    private const int MostProjectBytes = 64 * 1024 * 1024;

    /// <summary>
    /// The `vbaProject.bin` inside a document, or null when there is none to have - a document
    /// with no VBA, a legacy `.xls` (which is a compound file rather than a package, and is not
    /// read here), a path that will not open, or bytes this cannot walk.
    ///
    /// Opened read-only and SHARED, because the application has the file open while we read it.
    /// </summary>
    public static byte[]? ProjectBytes(string documentPath)
    {
        using var file = new FileStream(
            documentPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var zip = new ZipArchive(file, ZipArchiveMode.Read);

        foreach (var path in ProjectPaths)
        {
            if (zip.GetEntry(path) is not { } entry)
            {
                continue;
            }

            if (entry.Length is <= 0 or > MostProjectBytes)
            {
                return null;
            }

            var bytes = new byte[entry.Length];
            using var stream = entry.Open();
            stream.ReadExactly(bytes);
            return bytes;
        }

        return null;
    }
}
