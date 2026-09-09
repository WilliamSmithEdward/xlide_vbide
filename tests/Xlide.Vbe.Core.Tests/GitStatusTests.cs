using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The branch-state parser against `git status --porcelain=v2 --branch -z` as git 2.55 prints
/// it: a tracked branch, an unborn one, a conflict, and a detached head with a staged rename.
/// The NUL that ends every entry is spelled as a token and put back by <see cref="Captured"/>.
/// </summary>
public sealed class GitStatusTests
{
    /// <summary>xlide_vbide on main, 2026-09-08, with two untracked files in the tree.</summary>
    private const string OnMain =
        "# branch.oid 5bbbd77cafad268c3e2e1cc25f39859b4db2960c<NUL># branch.head main<NUL>"
        + "# branch.upstream origin/main<NUL># branch.ab +0 -0<NUL>? docs/source-control.md<NUL>"
        + "? ui/editor/src/blameformat.ts<NUL>";

    /// <summary>A repository straight after `git init -b main`.</summary>
    private const string Unborn = "# branch.oid (initial)<NUL># branch.head main<NUL>";

    /// <summary>A merge of `topic` into `main` that stopped on Account.cls.</summary>
    private const string Conflicted =
        "# branch.oid 7e2ae45f52f563ca097f4f3b61c522bc1e743ebd<NUL># branch.head main<NUL>"
        + "u UU N... 100644 100644 100644 100644 62789feb1e1db826f9f20533c0e96ab6fc65f225 "
        + "bc4ec8ac24799ba087c8debfcb2537b877575ef3 b3b62e5cec220972cbad7cfebcf5bbdd141b017c "
        + "Account.cls<NUL>? Books.bas.tmp<NUL>? Scratch.bas<NUL>";

    /// <summary>`git checkout --detach HEAD~1` then `git mv Account.cls Acct.cls`.</summary>
    private const string DetachedWithRename =
        "# branch.oid 26b56c1d18902dafcc582ce7fde57d8c71d4635f<NUL># branch.head (detached)<NUL>"
        + "2 R. N... 100644 100644 100644 62789feb1e1db826f9f20533c0e96ab6fc65f225 "
        + "62789feb1e1db826f9f20533c0e96ab6fc65f225 R100 Acct.cls<NUL>Account.cls<NUL>"
        + "? Books.bas.tmp<NUL>? Scratch.bas<NUL>";

    private static string Captured(string text) => text.Replace("<NUL>", "\0", StringComparison.Ordinal);

    [Fact]
    public void ATrackedBranchReportsItsCommitUpstreamAndCounts()
    {
        var state = GitStatus.Parse(Captured(OnMain));

        Assert.Equal("5bbbd77cafad268c3e2e1cc25f39859b4db2960c", state.Oid);
        Assert.Equal("main", state.Head);
        Assert.Equal("origin/main", state.Upstream);
        Assert.Equal(0, state.Ahead);
        Assert.Equal(0, state.Behind);
        Assert.False(state.UpstreamGone);
        Assert.Empty(state.Conflicts);

        // Untracked files are not changes to anything committed, so they are not rows.
        Assert.Empty(state.Changed);
    }

    [Fact]
    public void AnUpstreamWhoseRemoteBranchIsGoneHasNoCountsAndSaysSo()
    {
        // `git push origin --delete feature` elsewhere, then `git fetch --prune` here: the
        // upstream line stays, the counts line goes. Zero ahead here must not read as "on the
        // upstream", or the head's Undo would be refused as already pushed.
        var state = GitStatus.Parse(Captured(
            "# branch.oid 5bbbd77cafad268c3e2e1cc25f39859b4db2960c<NUL># branch.head feature<NUL>"
            + "# branch.upstream origin/feature<NUL>"));

        Assert.Equal("origin/feature", state.Upstream);
        Assert.True(state.UpstreamGone);
        Assert.Equal(0, state.Ahead);
    }

    [Fact]
    public void AnUnbornBranchHasANameAndNoCommit()
    {
        var state = GitStatus.Parse(Captured(Unborn));

        Assert.Null(state.Oid);
        Assert.Equal("main", state.Head);
        Assert.Null(state.Upstream);
    }

    [Fact]
    public void AnUnmergedFileIsAConflictAndNotAChange()
    {
        var state = GitStatus.Parse(Captured(Conflicted));

        Assert.Equal(["Account.cls"], state.Conflicts);
        Assert.Empty(state.Changed);
        Assert.Null(state.Upstream);
    }

    [Fact]
    public void ARenamedEntrysSecondTokenIsItsOldNameNotAnotherEntry()
    {
        // With -z the original path follows the entry as a NUL-terminated token of its own; read
        // as an entry it would have been a change called "Account.cls" that no longer exists.
        var state = GitStatus.Parse(Captured(DetachedWithRename));

        Assert.Equal("(detached)", state.Head);
        Assert.Equal(["Acct.cls"], state.Changed);
        Assert.Empty(state.Conflicts);
    }

    [Fact]
    public void AheadAndBehindAreReadAndAPathKeepsItsSpaces()
    {
        // The entry shape is the captured one with a two-word file name, which a module named
        // with a space would produce.
        var output = Captured(
            "# branch.oid 5bbbd77cafad268c3e2e1cc25f39859b4db2960c<NUL># branch.head feature<NUL>"
            + "# branch.upstream origin/feature<NUL># branch.ab +3 -1<NUL>"
            + "1 .M N... 100644 100644 100644 62789feb1e1db826f9f20533c0e96ab6fc65f225 "
            + "62789feb1e1db826f9f20533c0e96ab6fc65f225 My Module.bas<NUL>");

        var state = GitStatus.Parse(output);

        Assert.Equal(3, state.Ahead);
        Assert.Equal(1, state.Behind);
        Assert.Equal(["My Module.bas"], state.Changed);
    }

    [Fact]
    public void NoOutputIsNoBranch()
    {
        var state = GitStatus.Parse(string.Empty);

        Assert.Null(state.Oid);
        Assert.Equal(string.Empty, state.Head);
        Assert.Empty(state.Changed);
        Assert.Empty(state.Conflicts);
    }
}
