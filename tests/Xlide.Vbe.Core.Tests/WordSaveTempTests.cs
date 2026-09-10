using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The name Word's VBA project answers for its file after a save, told apart from every other
/// name a project can carry: the document itself, Word's owner file beside it, and temp names
/// that have never been seen in a project's FileName.
/// </summary>
public sealed class WordSaveTempTests
{
    [Theory]
    [InlineData("~WRL0001.tmp")]
    [InlineData("~WRL0003.tmp")]
    [InlineData(@"F:\GitHub\xlide\xlide_vbide\artifacts\fixtures\~wrl0003.TMP")]
    [InlineData("C:/Users/someone/Documents/~WRL1234.tmp")]
    [InlineData("~WRL7.tmp")]
    public void WordsSaveTempIsRecognisedWhereverItSits(string path)
    {
        Assert.True(WordSaveTemp.Matches(path));
    }

    [Theory]
    [InlineData("WordFixture.docm")]
    [InlineData(@"F:\GitHub\xlide\xlide_vbide\artifacts\fixtures\WordFixture.docm")]
    [InlineData("~$rdFixture.docm")]
    [InlineData("~WRL.tmp")]
    [InlineData("~WRLabcd.tmp")]
    [InlineData("~WRL0001.docm")]
    [InlineData("~WRD0001.tmp")]
    [InlineData("WRL0001.tmp")]
    [InlineData("")]
    [InlineData(null)]
    public void EveryOtherNameIsNot(string? path)
    {
        Assert.False(WordSaveTemp.Matches(path));
    }
}
