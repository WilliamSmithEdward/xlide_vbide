using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The branch list parser against `git for-each-ref` output from git 2.55 over refs/heads and
/// refs/remotes, tabs spelled as tokens and put back by <see cref="Captured"/>.
/// </summary>
public sealed class GitBranchesTests
{
    /// <summary>xlide_vbide: one branch, checked out, tracking origin.</summary>
    private const string OneTracked = "refs/heads/main<TAB>main<TAB>*<TAB>origin/main\n";

    /// <summary>A repository with `topic` tracking `main` and nothing tracking a remote.</summary>
    private const string TwoLocal =
        "refs/heads/main<TAB>main<TAB>*<TAB>\n"
        + "refs/heads/topic<TAB>topic<TAB> <TAB>main\n";

    /// <summary>
    /// After a fetch from two remotes: origin's HEAD pointer, origin's copy of main, a branch only
    /// origin has, and a branch with a slash in its own name on a second remote.
    /// </summary>
    private const string WithRemotes =
        "refs/heads/main<TAB>main<TAB>*<TAB>origin/main\n"
        + "refs/heads/topic<TAB>topic<TAB> <TAB>\n"
        + "refs/remotes/origin/HEAD<TAB>origin<TAB> <TAB>\n"
        + "refs/remotes/origin/elsewhere<TAB>origin/elsewhere<TAB> <TAB>\n"
        + "refs/remotes/origin/main<TAB>origin/main<TAB> <TAB>\n"
        + "refs/remotes/upstream/deep/name<TAB>upstream/deep/name<TAB> <TAB>\n";

    private static string Captured(string text) => text.Replace("<TAB>", "\t", StringComparison.Ordinal);

    [Fact]
    public void TheCheckedOutBranchAndItsUpstreamAreRead()
    {
        var branch = Assert.Single(GitBranches.Parse(Captured(OneTracked)));

        Assert.Equal("main", branch.Name);
        Assert.True(branch.Current);
        Assert.Equal("origin/main", branch.Upstream);
        Assert.Null(branch.Remote);
        Assert.Equal("main", branch.Ref);
    }

    [Fact]
    public void ABranchWithNoUpstreamHasNullNotAnEmptyName()
    {
        var branches = GitBranches.Parse(Captured(TwoLocal));

        Assert.Equal(2, branches.Count);
        Assert.Equal("main", branches[0].Name);
        Assert.True(branches[0].Current);
        Assert.Null(branches[0].Upstream);

        // A local branch can track another local branch; the column is whatever git says.
        Assert.Equal("topic", branches[1].Name);
        Assert.False(branches[1].Current);
        Assert.Equal("main", branches[1].Upstream);
    }

    [Fact]
    public void NoOutputIsNoBranches()
    {
        Assert.Empty(GitBranches.Parse(string.Empty));
        Assert.Empty(GitBranches.Parse("\n\n"));
    }

    [Fact]
    public void ARemotesBranchIsNamedWithoutTheRemoteAndItsHeadPointerIsNotOne()
    {
        var branches = GitBranches.Parse(Captured(WithRemotes));

        Assert.Equal(
            ["main", "topic", "origin/elsewhere", "origin/main", "upstream/deep/name"],
            branches.Select(one => one.Ref).ToArray());

        var elsewhere = branches[2];
        Assert.Equal("elsewhere", elsewhere.Name);
        Assert.Equal("origin", elsewhere.Remote);
        Assert.False(elsewhere.Current);
        Assert.Null(elsewhere.Upstream);

        // The remote is the first segment; the branch keeps every slash of its own name.
        var deep = branches[4];
        Assert.Equal("deep/name", deep.Name);
        Assert.Equal("upstream", deep.Remote);

        Assert.DoesNotContain(branches, one => one.Name == "HEAD");
    }

    [Fact]
    public void ChoicesDropARemotesBranchALocalOneAnswersToAndKeepLocalsFirst()
    {
        var choices = GitBranches.Choices(GitBranches.Parse(Captured(WithRemotes)));

        Assert.Equal(
            ["main", "topic", "origin/elsewhere", "upstream/deep/name"],
            choices.Select(one => one.Ref).ToArray());
    }
}
