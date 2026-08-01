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
    public void TheShellDocumentIsExpectedBesideTheShim()
    {
        Assert.Equal(@"C:\Program Files\xlide\ui\shell\index.html", WebViewPaths.ShellDocument(ShimDirectory));
    }

    [Fact]
    public void AnEmptyDirectoryIsRejectedRatherThanProducingARelativePath()
    {
        Assert.Throws<ArgumentException>(() => WebViewPaths.LoaderLibrary("  "));
        Assert.Throws<ArgumentException>(() => WebViewPaths.ShellDocument("  "));
        Assert.Throws<ArgumentException>(() => WebViewPaths.UserDataFolder("  "));
    }
}

/// <summary>The markup handed to the browser, and what it says when the real document is absent.</summary>
public class ShellDocumentTests
{
    [Fact]
    public void TheVersionTokenIsReplaced()
    {
        var result = ShellDocument.Compose($"<p>{ShellDocument.VersionToken}</p>", "150.0.4078.105");

        Assert.Equal("<p>150.0.4078.105</p>", result);
    }

    [Fact]
    public void EveryOccurrenceOfTheTokenIsReplaced()
    {
        var template = ShellDocument.VersionToken + "|" + ShellDocument.VersionToken;

        Assert.Equal("1.2|1.2", ShellDocument.Compose(template, "1.2"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void AnAbsentVersionBecomesAWordRatherThanAnEmptyGap(string? version)
    {
        Assert.Equal("unknown", ShellDocument.Compose(ShellDocument.VersionToken, version));
    }

    [Fact]
    public void ATemplateWithoutTheTokenIsUnchanged()
    {
        const string template = "<p>no token here</p>";

        Assert.Equal(template, ShellDocument.Compose(template, "1.2"));
    }

    [Fact]
    public void TheMissingDocumentNamesWhereItWasExpected()
    {
        // An empty pane is indistinguishable from a browser that never started, so the fallback
        // has to say which of the two happened.
        var markup = ShellDocument.Missing(@"C:\Program Files\xlide\ui\shell\index.html");

        Assert.Contains(@"C:\Program Files\xlide\ui\shell\index.html", markup, StringComparison.Ordinal);
        Assert.Contains("not installed", markup, StringComparison.Ordinal);
    }

    [Fact]
    public void TheMissingDocumentEscapesThePathItReports()
    {
        var markup = ShellDocument.Missing(@"C:\a<b>&c\index.html");

        Assert.Contains("&lt;b&gt;&amp;c", markup, StringComparison.Ordinal);
        Assert.DoesNotContain("<b>", markup, StringComparison.Ordinal);
    }

    [Fact]
    public void TheMissingDocumentKeepsItsOwnStyleRules()
    {
        // The fallback is composed through string interpolation over markup that contains braces.
        var markup = ShellDocument.Missing("x");

        Assert.Contains("color-scheme: light dark;", markup, StringComparison.Ordinal);
    }
}

/// <summary>
/// The tool window's identity in the editor. The position identifier is a layout key the editor
/// stores per user; changing it after release loses everyone's docking arrangement.
/// </summary>
public class ToolWindowIdentityTests
{
    [Fact]
    public void ThePositionIdentifierIsAWellFormedBracedIdentifier()
    {
        Assert.StartsWith("{", ProductIdentity.ToolWindowPositionGuid, StringComparison.Ordinal);
        Assert.EndsWith("}", ProductIdentity.ToolWindowPositionGuid, StringComparison.Ordinal);
        Assert.True(Guid.TryParse(ProductIdentity.ToolWindowPositionGuid, out _));
    }

    [Fact]
    public void ThePositionIdentifierIsNotTheClassIdentifier()
    {
        // They are different kinds of thing. The editor never resolves the position identifier to
        // a class, and reusing the class identifier for it invites the assumption that it does.
        Assert.NotEqual(
            Guid.Parse(ProductIdentity.ToolWindowHostClsid),
            Guid.Parse(ProductIdentity.ToolWindowPositionGuid));

        Assert.NotEqual(
            Guid.Parse(ProductIdentity.AddInClsid),
            Guid.Parse(ProductIdentity.ToolWindowPositionGuid));
    }

    [Fact]
    public void TheCaptionIsSet()
    {
        Assert.False(string.IsNullOrWhiteSpace(ProductIdentity.ToolWindowCaption));
    }
}
