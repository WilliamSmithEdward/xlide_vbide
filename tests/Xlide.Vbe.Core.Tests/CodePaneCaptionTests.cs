using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Caption reading decides which window an editor surface is placed over. Getting it wrong puts the
/// surface on the Immediate window, or on the wrong module, so the cases where it must decline to
/// answer matter as much as the ones where it answers.
/// </summary>
public class CodePaneCaptionTests
{
    [Theory]
    [InlineData("ProbeModule (Code)", "ProbeModule")]
    [InlineData("Sheet1 (Code)", "Sheet1")]
    [InlineData("  ThisWorkbook (Code)  ", "ThisWorkbook")]
    [InlineData("My Class (Code)", "My Class")]
    public void ReadsTheComponentNameFromAnEnglishCaption(string caption, string expected)
    {
        Assert.Equal(expected, CodePaneCaption.ComponentName(caption));
    }

    [Theory]
    [InlineData("Immediate")]
    [InlineData("Project - VBAProject")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void DeclinesToNameAComponentWhenTheCaptionIsNotACodePaneCaption(string? caption)
    {
        Assert.Null(CodePaneCaption.ComponentName(caption));
    }

    [Fact]
    public void RecognisesTheImmediateWindowWhichSharesTheCodePaneWindowClass()
    {
        Assert.True(CodePaneCaption.IsKnownNonCodePane("Immediate"));
        Assert.False(CodePaneCaption.IsKnownNonCodePane("ProbeModule (Code)"));
    }

    [Fact]
    public void MatchesAnOpenComponentByName()
    {
        var open = new[] { "Sheet1", "ThisWorkbook", "ProbeModule" };
        Assert.Equal("ProbeModule", CodePaneCaption.MatchComponent("ProbeModule (Code)", open));
    }

    [Fact]
    public void MatchesWhenTheSuffixIsInAnotherLanguage()
    {
        // The suffix is localised and the component name is not, because the user chose it. A
        // caption we cannot fully parse is still usable.
        var open = new[] { "Sheet1", "ProbeModule" };
        Assert.Equal("ProbeModule", CodePaneCaption.MatchComponent("ProbeModule (Code-Fenster)", open));
    }

    [Fact]
    public void PrefersTheLongestMatchingComponentName()
    {
        // Sheet1 is a prefix of Sheet10. Taking the first match would put the surface on the wrong
        // module for every workbook with more than nine sheets.
        var open = new[] { "Sheet1", "Sheet10" };
        Assert.Equal("Sheet10", CodePaneCaption.MatchComponent("Sheet10 (Code)", open));
    }

    [Fact]
    public void ReturnsNothingWhenNoOpenComponentExplainsTheCaption()
    {
        var open = new[] { "Sheet1", "ThisWorkbook" };
        Assert.Null(CodePaneCaption.MatchComponent("Immediate", open));
    }

    [Fact]
    public void ReturnsNothingWhenNoComponentsAreOpen()
    {
        Assert.Null(CodePaneCaption.MatchComponent("ProbeModule (Code)", []));
    }

    [Fact]
    public void DoesNotMatchAComponentThatMerelyAppearsLaterInTheCaption()
    {
        // A component name has to lead the caption. Otherwise "Copy of Sheet1 (Code)" would be
        // attributed to Sheet1.
        var open = new[] { "Sheet1" };
        Assert.Null(CodePaneCaption.MatchComponent("Copy of Sheet1 (Code)", open));
    }
}
