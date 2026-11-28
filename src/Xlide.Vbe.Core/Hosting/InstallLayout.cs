namespace Xlide.Vbe.Core.Hosting;

/// <summary>
/// Where an installation keeps the files of each architecture.
///
/// One installation serves 64-bit and 32-bit Office. A host loads only a library of its own
/// width, so there are two shims, each with the browser loader of the same width beside it. The
/// editor page, the engine and the icon are the same for both and are installed once.
///
/// The 64-bit shim sits in the installation folder, where it always has, with everything shared
/// around it. The 32-bit shim and its loader sit in one subfolder of that, and look one level up
/// for the rest. A development publish is not laid out this way: each architecture publishes to
/// its own folder with its own copy of the page, and nothing there is named like the subfolder.
/// </summary>
public static class InstallLayout
{
    /// <summary>Subfolder of an installation holding the 32-bit shim and its browser loader.</summary>
    public const string X86FolderName = "x86";

    /// <summary>
    /// The folder holding what both shims share, given the folder a shim was loaded from: that
    /// folder itself, or its parent when it is the 32-bit subfolder of an installation.
    /// </summary>
    public static string SharedDirectory(string shimDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shimDirectory);

        var folder = Path.TrimEndingDirectorySeparator(shimDirectory);
        if (!string.Equals(Path.GetFileName(folder), X86FolderName, StringComparison.OrdinalIgnoreCase))
        {
            return shimDirectory;
        }

        return Path.GetDirectoryName(folder) is { Length: > 0 } parent ? parent : shimDirectory;
    }

    /// <summary>Where an installation at <paramref name="installDirectory"/> keeps its 32-bit shim.</summary>
    public static string X86ShimPath(string installDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(installDirectory);
        return Path.Combine(installDirectory, X86FolderName, ProductIdentity.ShimFileName);
    }
}
