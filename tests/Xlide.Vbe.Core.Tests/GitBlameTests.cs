using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The blame parser against `git blame --porcelain` from git 2.55, and the mapping from the
/// committed file's lines to the editor's. The tab that opens every content line is spelled as
/// a token and put back by <see cref="Captured"/>.
/// </summary>
public sealed class GitBlameTests
{
    /// <summary>
    /// Books.bas at HEAD in a repository of three commits: the rename commit owns the header
    /// line, the root commit (a boundary) owns the skeleton, the middle commit the body. The last
    /// group names a commit whose facts were printed earlier and are not repeated.
    /// </summary>
    private const string BooksAtHead = """
        26b56c1d18902dafcc582ce7fde57d8c71d4635f 1 1 1
        author Ada Lovelace
        author-mail <ada@example.test>
        author-time 1788427800
        author-tz +0200
        committer Ada Lovelace
        committer-mail <ada@example.test>
        committer-time 1788427800
        committer-tz +0200
        summary Ledger is Books now
        previous 92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 Ledger.bas
        filename Books.bas
        <TAB>Attribute VB_Name = "Books"
        01571f56966ec45e571e2062667336765c0fdbf0 2 2 3
        author Ada Lovelace
        author-mail <ada@example.test>
        author-time 1788250500
        author-tz +0100
        committer Ada Lovelace
        committer-mail <ada@example.test>
        committer-time 1788250500
        committer-tz +0100
        summary Ledger and Account arrive
        boundary
        filename Ledger.bas
        <TAB>Option Explicit
        01571f56966ec45e571e2062667336765c0fdbf0 3 3
        <TAB>
        01571f56966ec45e571e2062667336765c0fdbf0 4 4
        <TAB>Public Sub Post()
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 5 5 4
        author Ada Lovelace
        author-mail <ada@example.test>
        author-time 1788339600
        author-tz +0100
        committer Ada Lovelace
        committer-mail <ada@example.test>
        committer-time 1788339600
        committer-tz +0100
        summary Ledger posts and closes
        previous 01571f56966ec45e571e2062667336765c0fdbf0 Ledger.bas
        filename Ledger.bas
        <TAB>    Debug.Print "Ledger posts"
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 6 6
        <TAB>End Sub
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 7 7
        <TAB>
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 8 8
        <TAB>Public Sub Close()
        01571f56966ec45e571e2062667336765c0fdbf0 6 9 1
        <TAB>End Sub

        """;

    /// <summary>The same file over `HEAD~1..HEAD`: everything older is charged to the range's
    /// edge commit and marked boundary.</summary>
    private const string BooksSinceParent = """
        26b56c1d18902dafcc582ce7fde57d8c71d4635f 1 1 1
        author Ada Lovelace
        author-mail <ada@example.test>
        author-time 1788427800
        author-tz +0200
        committer Ada Lovelace
        committer-mail <ada@example.test>
        committer-time 1788427800
        committer-tz +0200
        summary Ledger is Books now
        previous 92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 Ledger.bas
        filename Books.bas
        <TAB>Attribute VB_Name = "Books"
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 2 2 8
        author Ada Lovelace
        author-mail <ada@example.test>
        author-time 1788339600
        author-tz +0100
        committer Ada Lovelace
        committer-mail <ada@example.test>
        committer-time 1788339600
        committer-tz +0100
        summary Ledger posts and closes
        boundary
        filename Ledger.bas
        <TAB>Option Explicit
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 3 3
        <TAB>
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 4 4
        <TAB>Public Sub Post()
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 5 5
        <TAB>    Debug.Print "Ledger posts"
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 6 6
        <TAB>End Sub
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 7 7
        <TAB>
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 8 8
        <TAB>Public Sub Close()
        92eb6bb31205a76e71dacb0a49eb8aefe2d9c319 9 9
        <TAB>End Sub

        """;

    /// <summary>xlide_vbide's global.json at HEAD, one root commit, a negative offset.</summary>
    private const string GlobalJson = """
        1e5698150ce8dc7f4a524bf1066a1c3e648292bf 1 1 6
        author William Smith
        author-mail <williamsmithe@icloud.com>
        author-time 1785606574
        author-tz -0700
        committer William Smith
        committer-mail <williamsmithe@icloud.com>
        committer-time 1785606574
        committer-tz -0700
        summary Load a native add-in inside the Visual Basic Editor
        boundary
        filename global.json
        <TAB>{
        1e5698150ce8dc7f4a524bf1066a1c3e648292bf 2 2
        <TAB>  "sdk": {
        1e5698150ce8dc7f4a524bf1066a1c3e648292bf 3 3
        <TAB>    "version": "10.0.302",
        1e5698150ce8dc7f4a524bf1066a1c3e648292bf 4 4
        <TAB>    "rollForward": "latestFeature"
        1e5698150ce8dc7f4a524bf1066a1c3e648292bf 5 5
        <TAB>  }
        1e5698150ce8dc7f4a524bf1066a1c3e648292bf 6 6
        <TAB>}

        """;

    /// <summary>Books.bas as `git cat-file` answers it: the header line, then the code.</summary>
    private const string BooksFile = """
        Attribute VB_Name = "Books"
        Option Explicit

        Public Sub Post()
            Debug.Print "Ledger posts"
        End Sub

        Public Sub Close()
        End Sub

        """;

    private static string Captured(string text) => text
        .Replace("\r\n", "\n", StringComparison.Ordinal)
        .Replace("<TAB>", "\t", StringComparison.Ordinal);

    [Fact]
    public void EveryLineIsChargedToItsCommitInTheFilesOwnNumbering()
    {
        var lines = GitBlame.ParsePorcelain(Captured(BooksAtHead));

        Assert.Equal(Enumerable.Range(1, 9), lines.Select(one => one.Line));
        Assert.Equal(
            ["26b56c1", "01571f5", "01571f5", "01571f5", "92eb6bb", "92eb6bb", "92eb6bb", "92eb6bb", "01571f5"],
            lines.Select(one => one.ShortHash));

        var header = lines[0];
        Assert.Equal("26b56c1d18902dafcc582ce7fde57d8c71d4635f", header.Hash);
        Assert.Equal("Ada Lovelace", header.Author);
        Assert.Equal("Ledger is Books now", header.Summary);
    }

    [Fact]
    public void WhenIsTheAuthorsOwnClockAsIso8601()
    {
        var lines = GitBlame.ParsePorcelain(Captured(BooksAtHead));

        // author-time 1788427800 is 09:30 UTC on 2026-09-03; author-tz +0200 puts it at 11:30.
        Assert.Equal("2026-09-03T11:30:00+02:00", lines[0].When);
        Assert.Equal("2026-09-01T09:15:00+01:00", lines[1].When);
        Assert.Equal("2026-09-02T10:00:00+01:00", lines[4].When);
    }

    [Fact]
    public void AGroupWhoseCommitWasDescribedEarlierStillKnowsItsAuthor()
    {
        // The last group of the file is one line of the root commit, printed as a bare header
        // and its text: git said who that commit was eight groups ago and does not say it again.
        var last = GitBlame.ParsePorcelain(Captured(BooksAtHead))[8];

        Assert.Equal(9, last.Line);
        Assert.Equal("01571f5", last.ShortHash);
        Assert.Equal("Ada Lovelace", last.Author);
        Assert.Equal("Ledger and Account arrive", last.Summary);
        Assert.Equal("2026-09-01T09:15:00+01:00", last.When);
    }

    [Fact]
    public void ABoundaryCommitIsReadLikeAnyOther()
    {
        var lines = GitBlame.ParsePorcelain(Captured(BooksSinceParent));

        Assert.Equal(9, lines.Count);
        Assert.Equal("26b56c1", lines[0].ShortHash);
        Assert.All(lines.Skip(1), one => Assert.Equal("92eb6bb", one.ShortHash));
        Assert.All(lines.Skip(1), one => Assert.Equal("Ledger posts and closes", one.Summary));
    }

    [Fact]
    public void ANegativeOffsetIsKeptAsTheAuthorsOwn()
    {
        var lines = GitBlame.ParsePorcelain(Captured(GlobalJson));

        Assert.Equal(6, lines.Count);
        Assert.All(lines, one => Assert.Equal("2026-08-01T10:49:34-07:00", one.When));
        Assert.All(lines, one => Assert.Equal("William Smith", one.Author));
        Assert.Equal("Load a native add-in inside the Visual Basic Editor", lines[5].Summary);
    }

    [Fact]
    public void NoOutputIsNoLines()
    {
        Assert.Empty(GitBlame.ParsePorcelain(string.Empty));
    }

    [Fact]
    public void TheHeaderLineMapsNowhereAndALocalEditIsUncommitted()
    {
        // The editor holds the module without its attribute header, with CRLF endings and no
        // trailing newline, and with one line typed since the commit.
        var live = "Option Explicit\r\n\r\nPublic Sub Post()\r\n    Debug.Print \"Ledger posts\"\r\nEnd Sub\r\n"
            + "\r\nPublic Sub Close()\r\n    ' not yet\r\nEnd Sub";

        var map = GitBlame.MapToLive(
            GitBlame.ParsePorcelain(Captured(BooksAtHead)), Captured(BooksFile), live);

        Assert.Equal([8], map.Uncommitted);
        Assert.Equal([1, 2, 3, 4, 5, 6, 7, 9], map.Lines.Select(one => one.Line));

        // Live line 1 is the file's line 2, and carries that line's commit, not the header's.
        Assert.Equal("01571f5", map.Lines[0].ShortHash);
        Assert.Equal("92eb6bb", map.Lines.Single(one => one.Line == 4).ShortHash);
        Assert.Equal("01571f5", map.Lines.Single(one => one.Line == 9).ShortHash);
    }

    [Fact]
    public void AModuleWithNothingCommittedIsAllUncommitted()
    {
        var map = GitBlame.MapToLive([], string.Empty, "Option Explicit\nSub A()\nEnd Sub");

        Assert.Empty(map.Lines);
        Assert.Equal([1, 2, 3], map.Uncommitted);
    }

    [Fact]
    public void ARewrittenLineIsUncommittedAndTheRestKeepTheirBlame()
    {
        var live = "Option Explicit\n\nPublic Sub Post()\n    Debug.Print \"Books posts\"\nEnd Sub\n"
            + "\nPublic Sub Close()\nEnd Sub";

        var map = GitBlame.MapToLive(
            GitBlame.ParsePorcelain(Captured(BooksAtHead)), Captured(BooksFile), live);

        Assert.Equal([4], map.Uncommitted);
        Assert.Equal(7, map.Lines.Count);
    }

    [Fact]
    public void AFilesTrailingNewlineDoesNotUnmapAModulePastTheComparisonsCap()
    {
        // 2,100 lines is past the comparison's alignable middle. The exported file ends with a
        // newline and the module's text does not; left in, that phantom last line kept the
        // common-tail trim from engaging, the whole module fell into the over-cap path, and every
        // line of an untouched module read as uncommitted (found in review, 2026-09-08).
        var code = string.Join(
            "\r\n", Enumerable.Range(1, 2100).Select(at => $"    Debug.Print {at}"));
        var file = "Attribute VB_Name = \"Big\"\r\n" + code + "\r\n";
        var blamed = Enumerable.Range(1, 2101)
            .Select(line => new BlameLine(
                line, "01571f56966ec45e571e2062667336765c0fdbf0", "01571f5", "Ada",
                "2026-09-01T09:15:00+01:00", "big"))
            .ToList();

        var map = GitBlame.MapToLive(blamed, file, code);

        Assert.Empty(map.Uncommitted);
        Assert.Equal(2100, map.Lines.Count);
        Assert.Equal(1, map.Lines[0].Line);
        Assert.Equal(2100, map.Lines[^1].Line);
    }

    [Fact]
    public void AZeroOffsetIsSpelledZ()
    {
        // As git spells %aI for the same instant, so a commit's date reads the same from the log
        // and from blame.
        var porcelain = "01571f56966ec45e571e2062667336765c0fdbf0 1 1 1\n"
            + "author Ada\nauthor-mail <ada@example.test>\n"
            + "author-time 1788523200\nauthor-tz +0000\n"
            + "committer Ada\ncommitter-mail <ada@example.test>\n"
            + "committer-time 1788523200\ncommitter-tz +0000\n"
            + "summary utc\nfilename Ledger.bas\n\tOption Explicit\n";

        var line = Assert.Single(GitBlame.ParsePorcelain(porcelain));
        Assert.Equal("2026-09-04T12:00:00Z", line.When);
    }
}
