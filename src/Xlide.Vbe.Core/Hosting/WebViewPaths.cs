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
    /// Browser profile directory for the current user. Derived from the local application data
    /// path rather than read from the environment so it can be asserted in a test.
    /// </summary>
    public static string UserDataFolder(string localApplicationDataPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(localApplicationDataPath);

        return Path.Combine(localApplicationDataPath, ProductIdentity.DataFolderName, UserDataFolderName);
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
