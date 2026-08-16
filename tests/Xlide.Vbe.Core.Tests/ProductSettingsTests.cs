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

        // On, and six points, which is what the editor's own Options dialog ships.
        Assert.True(settings.DesignerSnapToGrid);
        Assert.Equal(6, settings.DesignerGridSize);
    }

    [Theory]
    [InlineData(0, 2)]
    [InlineData(1, 2)]
    [InlineData(-4, 2)]
    [InlineData(25, 24)]
    [InlineData(8, 8)]
    public void TheGridSizeIsClampedToItsLegalRange(int asked, int kept)
    {
        var settings = ProductSettings.Parse($"{{ \"designer.gridSize\": {asked} }}");

        Assert.Equal(kept, settings.DesignerGridSize);
    }

    /// <summary>
    /// A KEY THE FILE DOES NOT MENTION READS AS ITS SHIPPING DEFAULT, not as its type's zero.
    ///
    /// Every developer with a settings file written before a setting existed is in this case,
    /// which is every developer the day any setting is added. It went unnoticed for as long as
    /// every key in the record happened to be in every file: the grid arrived on 2026-08-16 and
    /// read back off as a two-point spacing - `default(bool)` and `default(int)` clamped - from
    /// a file that named neither.
    /// </summary>
    [Fact]
    public void AKeyTheFileNeverMentionsKeepsItsShippingDefault()
    {
        var settings = ProductSettings.Parse("{ \"format.indentSize\": 2 }");

        Assert.Equal(2, settings.FormatIndentSize);
        Assert.True(settings.DesignerSnapToGrid);
        Assert.Equal(6, settings.DesignerGridSize);
        Assert.Equal("comfy", settings.BlockLayout);
        Assert.True(settings.ContinueCommentOnNewline);
        Assert.Equal("xlide", settings.SyncEngine);
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
