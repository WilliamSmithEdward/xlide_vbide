using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Where git.exe is looked for and in what order. The order is the whole contract: PATH wins,
/// because that is the git the developer's own shell runs, and the three Git for Windows
/// locations follow for a machine where PATH was never updated.
/// </summary>
public sealed class GitLocatorTests
{
    [Fact]
    public void PathEntriesComeFirstThenTheThreeInstallLocations()
    {
        var candidates = GitLocator.Candidates(
            @"C:\Tools;C:\Users\me\bin",
            @"C:\Program Files",
            @"C:\Program Files (x86)",
            @"C:\Users\me\AppData\Local");

        Assert.Equal(
            [
                @"C:\Tools\git.exe",
                @"C:\Users\me\bin\git.exe",
                @"C:\Program Files\Git\cmd\git.exe",
                @"C:\Program Files (x86)\Git\cmd\git.exe",
                @"C:\Users\me\AppData\Local\Programs\Git\cmd\git.exe",
            ],
            candidates);
    }

    [Fact]
    public void AFolderNamedTwiceIsProbedOnceAtItsFirstPosition()
    {
        // The usual machine: Git for Windows put itself on PATH, so Program Files would be tried
        // twice, and a second File.Exists cannot answer differently from the first.
        var candidates = GitLocator.Candidates(
            @"C:\Windows;c:\program files\git\cmd", @"C:\Program Files", null, null);

        Assert.Equal([@"C:\Windows\git.exe", @"c:\program files\git\cmd\git.exe"], candidates);
    }

    [Fact]
    public void QuotedBlankAndTrailingSlashEntriesAreReadAsCmdWouldReadThem()
    {
        var candidates = GitLocator.Candidates("\"C:\\Quoted Dir\";;  ;C:\\Trailing\\", null, null, null);

        Assert.Equal([@"C:\Quoted Dir\git.exe", @"C:\Trailing\git.exe"], candidates);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(";;")]
    public void NothingToLookInIsNoCandidates(string? pathVariable)
    {
        Assert.Empty(GitLocator.Candidates(pathVariable, null, string.Empty, "  "));
    }

    [Fact]
    public void TheFileNameIsTheProductsOwnConstant()
    {
        Assert.Equal("git.exe", ProductIdentity.GitFileName);
        var only = Assert.Single(GitLocator.Candidates(null, @"D:\PF", null, null));
        Assert.EndsWith(@"\git.exe", only, StringComparison.Ordinal);
    }
}
