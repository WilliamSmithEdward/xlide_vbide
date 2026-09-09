using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The branch list parser against `git for-each-ref` output from git 2.55, tabs spelled as
/// tokens and put back by <see cref="Captured"/>.
/// </summary>
public sealed class GitBranchesTests
{
    /// <summary>xlide_vbide: one branch, checked out, tracking origin.</summary>
    private const string OneTracked = "main<TAB>*<TAB>origin/main\n";

    /// <summary>A repository with `topic` tracking `main` and nothing tracking a remote.</summary>
    private const string TwoLocal = "main<TAB>*<TAB>\ntopic<TAB> <TAB>main\n";

    private static string Captured(string text) => text.Replace("<TAB>", "\t", StringComparison.Ordinal);

    [Fact]
    public void TheCheckedOutBranchAndItsUpstreamAreRead()
    {
        var branch = Assert.Single(GitBranches.Parse(Captured(OneTracked)));

        Assert.Equal("main", branch.Name);
        Assert.True(branch.Current);
        Assert.Equal("origin/main", branch.Upstream);
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
}
