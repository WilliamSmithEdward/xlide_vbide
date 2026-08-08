namespace Xlide.Vbe.Core.Hosting;

/// <summary>
/// Where the browser host reads from and writes to.
///
/// Two of these are load-bearing and neither is obvious from the outside:
///
/// The user data folder must be writable. WebView2 creates a profile, a cache, and lock files
/// there on first use, so it can never be the install directory: that is under Program Files on a
/// real installation and a per-user process cannot write to it. Failure is not a permission error
/// the user sees, it is an environment that never completes creation.
///
/// The loader and the bundle sit next to the shim library, not next to the host executable.
/// The host is EXCEL.EXE and its directory belongs to Office, so the usual base-directory notion is
/// the wrong answer here. Callers pass the shim's own directory, which the shim resolves from its
/// module handle.
/// </summary>
public static class WebViewPaths
{
    /// <summary>Folder name under the product data directory holding the browser profile.</summary>
    public const string UserDataFolderName = "webview2";

    /// <summary>File name of the native WebView2 loader shipped beside the shim.</summary>
    public const string LoaderFileName = "WebView2Loader.dll";

    /// <summary>
    /// Virtual host name the editor bundle is served from.
    ///
    /// The name is mapped inside one browser instance and resolves nowhere else, so it must be one
    /// that can never become a real address. The .local suffix is reserved for link-local name
    /// resolution and is not registrable, which makes a collision with something the user's network
    /// or the public namespace owns impossible rather than unlikely.
    /// </summary>
    public const string EditorHostName = "xlide.local";

    /// <summary>Path of the editor bundle relative to the shim directory.</summary>
    /// <remarks>
    /// The bundle is a build output of the editor project, deployed beside the shim the same way
    /// the shell document is. It is absent from a working tree that has not built it, which is why
    /// every caller checks before mapping it.
    /// </remarks>
    public const string EditorContentRelativePath = @"ui\editor\dist";

    /// <summary>Document the editor bundle is entered through.</summary>
    public const string EditorEntryFileName = "index.html";

    /// <summary>
    /// Browser profile directory for the current user, PER PROCESS. Derived from the local
    /// application data path rather than read from the environment so it can be asserted in a
    /// test.
    ///
    /// ONE FOLDER FOR EVERY EXCEL WAS A SECOND EXCEL THAT NEVER STARTED. WebView2 takes a lock on
    /// its user data folder, and a second process pointed at the same one does not fail loudly:
    /// creating the environment simply never completes. The add-in loads, its door answers, the
    /// surface reports itself ready, and the page behind it never boots, so the loader spins for
    /// as long as the developer is willing to watch it. Reported twice before it was found
    /// (2026-08-08), because everything that can be asked says the session is healthy.
    ///
    /// The process id is what separates them, and it is the right key rather than a random one:
    /// it is stable for the session, and it lets a later session tell an abandoned folder from a
    /// live one. See <see cref="SweepAbandonedProfiles"/>.
    /// </summary>
    public static string UserDataFolder(string localApplicationDataPath, int processId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(localApplicationDataPath);

        return Path.Combine(
            localApplicationDataPath,
            ProductIdentity.DataFolderName,
            UserDataFolderName,
            processId.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Removes profile folders whose process is gone.
    ///
    /// A profile is tens of megabytes and one is made per Excel, so without this they accumulate
    /// for the life of the machine. Swept on start-up rather than on shutdown because a host that
    /// crashed never got to clean up, and the crash is exactly when a stale one is left behind.
    ///
    /// Anything that cannot be parsed, identified or deleted is left alone: a folder in use by a
    /// live process must survive, and being wrong about that is worse than a wasted megabyte.
    /// </summary>
    public static void SweepAbandonedProfiles(string localApplicationDataPath, Func<int, bool> isAlive)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(localApplicationDataPath);
        ArgumentNullException.ThrowIfNull(isAlive);

        var root = Path.Combine(localApplicationDataPath, ProductIdentity.DataFolderName, UserDataFolderName);
        if (!Directory.Exists(root))
        {
            return;
        }

        foreach (var folder in Directory.GetDirectories(root))
        {
            if (!int.TryParse(Path.GetFileName(folder), out var pid) || isAlive(pid))
            {
                continue;
            }

            try
            {
                Directory.Delete(folder, recursive: true);
            }
            catch (Exception)
            {
                // Held by something, or gone already. Either way the next session tries again.
            }
        }
    }

    /// <summary>Full path of the native loader that must be present beside the shim.</summary>
    public static string LoaderLibrary(string shimDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shimDirectory);

        return Path.Combine(shimDirectory, LoaderFileName);
    }

    /// <summary>
    /// Folder mapped to the virtual host name. This is a directory, not a file: the mapping serves
    /// everything under it, so it is the whole bundle rather than its entry document.
    /// </summary>
    public static string EditorContentRoot(string shimDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shimDirectory);

        return Path.Combine(shimDirectory, EditorContentRelativePath);
    }

    /// <summary>
    /// Full path of the editor's entry document. Its presence is what decides whether the bundle
    /// is worth mapping: mapping a folder that is not there succeeds and produces a blank pane on
    /// navigation instead of a diagnosable failure.
    /// </summary>
    public static string EditorEntryDocument(string shimDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shimDirectory);

        return Path.Combine(EditorContentRoot(shimDirectory), EditorEntryFileName);
    }

    /// <summary>
    /// Address the mapped bundle is entered through.
    ///
    /// The scheme is HTTPS rather than file, and that is the entire point of the mapping. A file
    /// URL is an opaque origin: module scripts, storage, and fetch are all refused there. A mapped
    /// host name gives a normal secure origin, so the bundle behaves the way it does in a browser.
    /// </summary>
    public static string EditorEntryUrl(string hostName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hostName);

        return $"https://{hostName}/{EditorEntryFileName}";
    }
}
