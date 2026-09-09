using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The history parser, against what git 2.55 actually prints for the pane's own log format.
///
/// The samples are verbatim captures. The control characters the format relies on - NUL between
/// fields, 0x1E before a record, 0x1F after the body, and the tab in a name-status row - are
/// spelled as tokens so the constants stay readable ASCII, and put back by <see cref="Captured"/>.
/// </summary>
public sealed class GitLogTests
{
    /// <summary>
    /// A repository built for the capture: a rename detected at 77% similarity, a body with two
    /// paragraphs and a trailing colon, and a root commit adding two files.
    /// </summary>
    private const string ThreeCommits = """
        <RS>26b56c1d18902dafcc582ce7fde57d8c71d4635f<NUL>26b56c1<NUL>Ada Lovelace<NUL>ada@example.test<NUL>2026-09-03T11:30:00+02:00<NUL>Ledger is Books now<NUL><US>

        R077<TAB>Ledger.bas<TAB>Books.bas
        <RS>92eb6bb31205a76e71dacb0a49eb8aefe2d9c319<NUL>92eb6bb<NUL>Ada Lovelace<NUL>ada@example.test<NUL>2026-09-02T10:00:00+01:00<NUL>Ledger posts and closes<NUL>The body has two paragraphs.

        This is the second, with a trailing colon:
        <US>

        M<TAB>Ledger.bas
        <RS>01571f56966ec45e571e2062667336765c0fdbf0<NUL>01571f5<NUL>Ada Lovelace<NUL>ada@example.test<NUL>2026-09-01T09:15:00+01:00<NUL>Ledger and Account arrive<NUL><US>

        A<TAB>Account.cls
        A<TAB>Ledger.bas

        """;

    /// <summary>`git log -3` over xlide_vbide itself on 2026-09-08: a long body with a blank line
    /// and backticks in it, a release commit with no body, and paths with directories.</summary>
    private const string ProductHistory = """
        <RS>5bbbd77cafad268c3e2e1cc25f39859b4db2960c<NUL>5bbbd77<NUL>William Smith<NUL>williamsmithe@icloud.com<NUL>2026-09-06T14:10:05-07:00<NUL>The record-cost bound measures the shape, not the runner<NUL>CI went red on `Version 0.14.2` over a change-log path nothing had touched in
        weeks: recording a ceiling-sized module answered 26.11 ms against an absolute
        25 ms bound, for work that takes under three here. A wall-clock ceiling on a
        shared runner is a coin toss, and this one had finally come up tails.

        The bound is now taken from the same machine, which is what the sibling test in
        this file already did: a 118 KB reference is measured first, and the 3.8 MB case
        has to come in under ninety-six times it. Linear in the text is about
        thirty-two times, so that allows three times over for a reference measurement
        that was mostly fixed overhead, while anything quadratic lands near a thousand
        times over and still fails. The floor that decides on a machine too fast to
        measure the reference is set where a shared runner cannot trip it, since the
        regression worth failing over is this landing on the per-keystroke path, which
        is hundreds of milliseconds rather than tens.

        No product code changed, so v0.14.2 as released is unaffected.
        <US>

        M<TAB>tests/Xlide.Vbe.Core.Tests/RecordCostTests.cs
        <RS>a3a4f516222e4b94a922baae5ed54079b469ae37<NUL>a3a4f51<NUL>William Smith<NUL>williamsmithe@icloud.com<NUL>2026-09-06T13:58:43-07:00<NUL>Version 0.14.2<NUL><US>

        M<TAB>Directory.Build.props
        M<TAB>docs/status.md
        <RS>35866d03747204b0cc784aff72a62ff19f787d63<NUL>35866d0<NUL>William Smith<NUL>williamsmithe@icloud.com<NUL>2026-09-06T13:58:28-07:00<NUL>The build instructions say what a first build actually needs<NUL>The section said .NET, C++ build tools and Node in one sentence, put the gate
        at twenty seconds when it is ninety, and never mentioned the thing that stops
        a first build dead: the engine compiles the analyzer out of the xlide_vscode
        checkout, so that repository has to be cloned as a sibling directory. It also
        never said that the two node projects need their own installs, that dev.ps1
        builds Release while the harness and every live suite need a Debug build for
        the api door, or that the fixtures the live tiers drive are build output with
        a generator each.
        <US>

        M<TAB>README.md

        """;

    private static string Captured(string text) => text
        .Replace("\r\n", "\n", StringComparison.Ordinal)
        .Replace("<NUL>", "\0", StringComparison.Ordinal)
        .Replace("<RS>", "\u001e", StringComparison.Ordinal)
        .Replace("<US>", "\u001f", StringComparison.Ordinal)
        .Replace("<TAB>", "\t", StringComparison.Ordinal);

    [Fact]
    public void TheFormatAndThePrefixAreWhatTheParserWasWrittenAgainst()
    {
        Assert.Equal("%x1e%H%x00%h%x00%an%x00%ae%x00%aI%x00%s%x00%b%x1f", GitArguments.LogFormat);
        Assert.Equal(
            ["--no-pager", "-c", "core.quotepath=false", "-c", "color.ui=never"],
            GitArguments.Prefix);
    }

    [Fact]
    public void EveryFieldOfACommitIsReadAndTheCommitsComeNewestFirst()
    {
        var commits = GitLog.Parse(Captured(ThreeCommits));

        Assert.Equal(3, commits.Count);
        var newest = commits[0];
        Assert.Equal("26b56c1d18902dafcc582ce7fde57d8c71d4635f", newest.Hash);
        Assert.Equal("26b56c1", newest.ShortHash);
        Assert.Equal("Ada Lovelace", newest.Author);
        Assert.Equal("ada@example.test", newest.Email);
        Assert.Equal("2026-09-03T11:30:00+02:00", newest.When);
        Assert.Equal("Ledger is Books now", newest.Subject);
        Assert.Equal(string.Empty, newest.Body);
        Assert.Equal(["01571f5"], commits.Skip(2).Select(one => one.ShortHash));
    }

    [Fact]
    public void ARenameIsOneChangeThatKnowsBothNames()
    {
        var change = Assert.Single(GitLog.Parse(Captured(ThreeCommits))[0].Files);

        // git prints the similarity as R077; the row's status is the letter alone.
        Assert.Equal("R", change.Status);
        Assert.Equal("Books.bas", change.Path);
        Assert.Equal("Ledger.bas", change.FromPath);
    }

    [Fact]
    public void AMultiParagraphBodyKeepsItsBlankLineAndLosesOnlyTheTrailingNewline()
    {
        var commit = GitLog.Parse(Captured(ThreeCommits))[1];

        Assert.Equal(
            "The body has two paragraphs.\n\nThis is the second, with a trailing colon:",
            commit.Body);
        var file = Assert.Single(commit.Files);
        Assert.Equal("M", file.Status);
        Assert.Equal("Ledger.bas", file.Path);
        Assert.Null(file.FromPath);
    }

    [Fact]
    public void ARootCommitListsEveryFileItAdded()
    {
        var root = GitLog.Parse(Captured(ThreeCommits))[2];

        Assert.Equal("Ledger and Account arrive", root.Subject);
        Assert.Equal(
            ["A Account.cls", "A Ledger.bas"],
            root.Files.Select(one => $"{one.Status} {one.Path}"));
    }

    [Fact]
    public void TheProductsOwnHistoryReadsWithDirectoriesInPathsAndBackticksInBodies()
    {
        var commits = GitLog.Parse(Captured(ProductHistory));

        Assert.Equal(["5bbbd77", "a3a4f51", "35866d0"], commits.Select(one => one.ShortHash));

        var bound = commits[0];
        Assert.StartsWith("CI went red on `Version 0.14.2`", bound.Body, StringComparison.Ordinal);
        Assert.EndsWith("as released is unaffected.", bound.Body, StringComparison.Ordinal);
        Assert.Contains("\n\nThe bound is now taken", bound.Body, StringComparison.Ordinal);
        Assert.Equal("tests/Xlide.Vbe.Core.Tests/RecordCostTests.cs", Assert.Single(bound.Files).Path);

        var release = commits[1];
        Assert.Equal("Version 0.14.2", release.Subject);
        Assert.Equal(string.Empty, release.Body);
        Assert.Equal(["Directory.Build.props", "docs/status.md"], release.Files.Select(one => one.Path));
        Assert.Equal("2026-09-06T13:58:43-07:00", release.When);
    }

    [Theory]
    [InlineData("")]
    [InlineData("\n")]
    [InlineData("   \n\n")]
    public void NothingIsNoCommits(string output)
    {
        Assert.Empty(GitLog.Parse(output));
    }

    [Fact]
    public void ARecordWithoutABodyFieldStillReads()
    {
        // A format that stops at the subject, or a body git never printed: the commit is still
        // there, with an empty body and its files.
        var output = Captured(
            "<RS>0123456789abcdef0123456789abcdef01234567<NUL>0123456<NUL>Ada<NUL>ada@example.test"
            + "<NUL>2026-09-01T09:15:00+01:00<NUL>Only a subject<US>\n\nD<TAB>Gone.bas\n");

        var commit = Assert.Single(GitLog.Parse(output));
        Assert.Equal("Only a subject", commit.Subject);
        Assert.Equal(string.Empty, commit.Body);
        Assert.Equal("D", Assert.Single(commit.Files).Status);
    }

    [Fact]
    public void AControlCharacterInsideABodyStaysInTheBody()
    {
        // git stores what was typed, 0x1E and 0x1F included. Neither may begin a record or end
        // the body early, or the text after it becomes a commit with no author that takes the
        // real commit's files (found in review, 2026-09-08).
        var output = Captured(
            "<RS>0123456789abcdef0123456789abcdef01234567<NUL>0123456<NUL>Ada<NUL>ada@example.test"
            + "<NUL>2026-09-01T09:15:00+01:00<NUL>Control characters<NUL>before-us<US>after-us\n"
            + "before-rs<RS>after-rs\n<US>\n\nM<TAB>Ledger.bas\n");

        var commit = Assert.Single(GitLog.Parse(output));
        Assert.Equal("Control characters", commit.Subject);
        Assert.Equal(Captured("before-us<US>after-us\nbefore-rs<RS>after-rs"), commit.Body);
        Assert.Equal("M", Assert.Single(commit.Files).Status);
    }
}
