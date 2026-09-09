using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The remote list parser against `git remote -v` output from git 2.55, tabs spelled as tokens
/// and put back by <see cref="Captured"/>.
/// </summary>
public sealed class GitRemotesTests
{
    /// <summary>xlide_vbide: origin, fetched and pushed over https.</summary>
    private const string Origin =
        "origin<TAB>https://github.com/WilliamSmithEdward/xlide_vbide.git (fetch)\n"
        + "origin<TAB>https://github.com/WilliamSmithEdward/xlide_vbide.git (push)\n";

    /// <summary>A fork: origin and upstream, git listing them by name.</summary>
    private const string Fork =
        "origin<TAB>git@github.com:ada/book.git (fetch)\n"
        + "origin<TAB>git@github.com:ada/book.git (push)\n"
        + "upstream<TAB>https://github.com/acme/book.git (fetch)\n"
        + "upstream<TAB>https://github.com/acme/book.git (push)\n";

    /// <summary>The suite's own remote: a bare repository beside the folder, by Windows path.</summary>
    private const string Local =
        "backup<TAB>D:\\Backups\\book.git (fetch)\n"
        + "backup<TAB>D:\\Backups\\book.git (push)\n"
        + "origin<TAB>C:\\Users\\Ada\\AppData\\Local\\Temp\\xlide-scm-51104-remote (fetch)\n"
        + "origin<TAB>C:\\Users\\Ada\\AppData\\Local\\Temp\\xlide-scm-51104-remote (push)\n";

    private static string Captured(string text) => text.Replace("<TAB>", "\t", StringComparison.Ordinal);

    [Fact]
    public void OneRemoteIsReadOnceByItsFetchUrl()
    {
        var remote = Assert.Single(GitRemotes.Parse(Captured(Origin)));

        Assert.Equal("origin", remote.Name);
        Assert.Equal("https://github.com/WilliamSmithEdward/xlide_vbide.git", remote.Url);
    }

    [Fact]
    public void APushUrlSetApartDoesNotReplaceTheFetchUrl()
    {
        var output = "origin<TAB>https://github.com/ada/book.git (fetch)\n"
            + "origin<TAB>git@github.com:ada/book.git (push)\n";

        var remote = Assert.Single(GitRemotes.Parse(Captured(output)));

        Assert.Equal("https://github.com/ada/book.git", remote.Url);
    }

    [Fact]
    public void RemotesKeepGitsOrderAndTheirNames()
    {
        var remotes = GitRemotes.Parse(Captured(Fork));

        Assert.Equal(["origin", "upstream"], remotes.Select(one => one.Name));
        Assert.Equal("https://github.com/acme/book.git", remotes[1].Url);
    }

    [Fact]
    public void ThePrimaryIsOriginWhereverGitListsIt()
    {
        var remotes = GitRemotes.Parse(Captured(Local));

        Assert.Equal(2, remotes.Count);
        Assert.Equal("backup", remotes[0].Name);

        var primary = GitRemotes.Primary(remotes);
        Assert.NotNull(primary);
        Assert.Equal("origin", primary.Name);
        Assert.Equal(@"C:\Users\Ada\AppData\Local\Temp\xlide-scm-51104-remote", primary.Url);
    }

    [Fact]
    public void WithoutOriginThePrimaryIsTheFirstAndWithNoneItIsNull()
    {
        var output = "mirror<TAB>https://example.com/book.git (fetch)\nmirror<TAB>https://example.com/book.git (push)\n";

        Assert.Equal("mirror", GitRemotes.Primary(GitRemotes.Parse(Captured(output)))?.Name);
        Assert.Null(GitRemotes.Primary(GitRemotes.Parse(string.Empty)));
        Assert.Empty(GitRemotes.Parse("\r\n\n"));
    }
}
