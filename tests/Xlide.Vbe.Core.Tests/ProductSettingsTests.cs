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
