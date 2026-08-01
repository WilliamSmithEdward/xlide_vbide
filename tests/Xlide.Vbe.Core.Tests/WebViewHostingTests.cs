using Xlide.Vbe.Core.Hosting;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The browser surface fails in ways that produce an empty pane rather than an error: an
/// unwritable profile directory, a loader that is not beside the shim, a document that never
/// shipped. These pin the decisions that avoid each of those.
/// </summary>
public class WebViewPathsTests
{
    private const string LocalAppData = @"C:\Users\someone\AppData\Local";
    private const string ShimDirectory = @"C:\Program Files\xlide";

    [Fact]
    public void TheProfileLivesUnderTheUsersOwnDataDirectory()
    {
        // Never the install directory. A per-user process cannot write there, and the browser
        // reports that as an environment that never finishes being created.
        var folder = WebViewPaths.UserDataFolder(LocalAppData);

        Assert.Equal(@"C:\Users\someone\AppData\Local\xlide_vbide\webview2", folder);
    }

    [Fact]
    public void TheProfileSharesTheProductDataFolderWithTheLogs()
    {
        var folder = WebViewPaths.UserDataFolder(LocalAppData);

        Assert.Contains(ProductIdentity.DataFolderName, folder, StringComparison.Ordinal);
    }

    [Fact]
    public void TheProfileIsNotUnderTheInstallDirectory()
    {
        var folder = WebViewPaths.UserDataFolder(LocalAppData);

        Assert.DoesNotContain(ShimDirectory, folder, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheLoaderIsExpectedBesideTheShim()
    {
        // Not on PATH and not in the host executable's directory: both belong to Office.
        Assert.Equal(@"C:\Program Files\xlide\WebView2Loader.dll", WebViewPaths.LoaderLibrary(ShimDirectory));
    }

    [Fact]
    public void TheContentRootIsExpectedBesideTheShim()
    {
        // Beside the shim, not beside the host executable: that directory is Office's.
        Assert.Equal(@"C:\Program Files\xlide\ui\editor\dist", WebViewPaths.EditorContentRoot(ShimDirectory));
    }

    [Fact]
    public void TheContentRootIsAFolderAndTheEntryDocumentSitsInsideIt()
    {
        // The mapping serves a directory. Handing it the entry document would map a path that is
        // not a folder, which succeeds and then serves nothing.
        var root = WebViewPaths.EditorContentRoot(ShimDirectory);
        var entry = WebViewPaths.EditorEntryDocument(ShimDirectory);

        Assert.Equal(Path.Combine(root, WebViewPaths.EditorEntryFileName), entry);
        Assert.Equal(root, Path.GetDirectoryName(entry));
    }

    [Fact]
    public void TheEntryDocumentIsTheFileWhosePresenceDecidesWhichSurfaceLoads()
    {
        Assert.Equal(@"C:\Program Files\xlide\ui\editor\dist\index.html", WebViewPaths.EditorEntryDocument(ShimDirectory));
    }

    [Fact]
    public void TheEntryUrlIsSecureHttpOnTheMappedHostName()
    {
        // Not file:. That scheme is an opaque origin, where module scripts, storage, and fetch are
        // all refused, which is the whole reason the mapping exists.
        Assert.Equal("https://xlide.local/index.html", WebViewPaths.EditorEntryUrl(WebViewPaths.EditorHostName));
    }

    [Fact]
    public void TheEntryUrlNamesTheHostItIsBuiltFrom()
    {
        var url = WebViewPaths.EditorEntryUrl("example.invalid");

        Assert.Equal("https://example.invalid/index.html", url);
        Assert.StartsWith("https://", url, StringComparison.Ordinal);
    }

    [Fact]
    public void TheEntryUrlAndTheEntryDocumentAgreeOnTheFileName()
    {
        // One constant, two consumers: the mapping serves the file the address asks for.
        Assert.EndsWith(
            "/" + WebViewPaths.EditorEntryFileName,
            WebViewPaths.EditorEntryUrl(WebViewPaths.EditorHostName),
            StringComparison.Ordinal);

        Assert.Equal(
            WebViewPaths.EditorEntryFileName,
            Path.GetFileName(WebViewPaths.EditorEntryDocument(ShimDirectory)));
    }

    [Fact]
    public void TheHostNameCannotCollideWithARealAddress()
    {
        // The name resolves only inside this browser. A registrable suffix would make the same name
        // resolvable elsewhere, so it uses the one reserved for link-local resolution.
        Assert.EndsWith(".local", WebViewPaths.EditorHostName, StringComparison.Ordinal);
        Assert.DoesNotContain("/", WebViewPaths.EditorHostName, StringComparison.Ordinal);
        Assert.DoesNotContain(":", WebViewPaths.EditorHostName, StringComparison.Ordinal);
    }
}
