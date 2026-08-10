using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// A settings file is something a developer grows, never something they need: every value has
/// a shipping default, every broken file reads as the defaults, and every value round-trips.
/// </summary>
public class ProductSettingsTests
{
    [Fact]
    public void TheDefaultsAreTheCompanionEditors()
    {
        var settings = ProductSettings.Default;

        Assert.Equal("comfy", settings.BlockLayout);
        Assert.True(settings.ContinueCommentOnNewline);
        Assert.True(settings.MirrorCommentSpacing);
        Assert.Equal(4, settings.FormatIndentSize);
    }

    /// <summary>
    /// A settings file written before the tabs option was removed still reads, and reads as
    /// spaces. The option could not be honoured: VBA's code store will not hold a tab, and
    /// expands any it is handed, so the page indented with tabs while the workbook held spaces.
    /// A developer who had it on gets the indent width they chose and nothing else changes.
    /// </summary>
    [Fact]
    public void AnOldFileAskingForTabsIsReadWithoutThem()
    {
        var settings = ProductSettings.Parse(
            "{ \"format.useTabs\": true, \"format.indentSize\": 2 }");

        Assert.Equal(2, settings.FormatIndentSize);
    }

    /// <summary>
    /// A settings file written before the canonical-keywords option was removed still reads.
    ///
    /// The same shape as the tabs option above, and removed for the same kind of reason: the
    /// switch could not be honoured. Two paths canonicalise keywords before the formatter is ever
    /// asked and neither consulted it, so a developer who turned it off watched their keywords be
    /// respelled anyway. Formatting still respells, always; the promise is what went.
    /// </summary>
    [Fact]
    public void AnOldFileAskingForPlainKeywordsIsReadWithoutThem()
    {
        var settings = ProductSettings.Parse(
            "{ \"format.canonicalKeywords\": false, \"format.indentSize\": 3 }");

        Assert.Equal(3, settings.FormatIndentSize);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(-3, 1)]
    [InlineData(9, 8)]
    [InlineData(100, 8)]
    [InlineData(2, 2)]
    public void TheIndentSizeIsClampedToItsLegalRange(int asked, int kept)
    {
        var settings = ProductSettings.Parse($"{{ \"format.indentSize\": {asked} }}");

        Assert.Equal(kept, settings.FormatIndentSize);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not json at all")]
    [InlineData("{ \"editor.blockLayout\": 17 }")]
    public void AnythingUnreadableReadsAsTheDefaults(string? json)
    {
        Assert.Equal(ProductSettings.Default, ProductSettings.Parse(json));
    }

    [Fact]
    public void ValuesRoundTripThroughTheFileText()
    {
        var settings = new ProductSettings
        {
            BlockLayout = "compact",
            ContinueCommentOnNewline = false,
            MirrorCommentSpacing = false,
        };

        Assert.Equal(settings, ProductSettings.Parse(settings.ToJson()));
    }

    [Fact]
    public void AnUnknownLayoutNormalisesToComfy()
    {
        var settings = ProductSettings.Parse("{ \"editor.blockLayout\": \"sideways\" }");

        Assert.Equal("comfy", settings.BlockLayout);
    }

    [Fact]
    public void UnknownKeysAreToleratedAndKnownOnesStillRead()
    {
        var settings = ProductSettings.Parse(
            "{ \"someday.newSetting\": true, \"editor.continueCommentOnNewline\": false }");

        Assert.False(settings.ContinueCommentOnNewline);
        Assert.Equal("comfy", settings.BlockLayout);
    }
}
