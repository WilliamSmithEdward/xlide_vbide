namespace Xlide.Vbe.Core.Scm;

/// <summary>A branch as the select offers it: a local one, or one only a remote has.</summary>
/// <param name="Name">The branch's own name - for a remote's branch, without the remote in front.</param>
/// <param name="Current">Whether HEAD is on it; never true for a remote's branch.</param>
/// <param name="Upstream">What a local branch tracks, or null.</param>
/// <param name="Remote">The remote a remote-only branch belongs to; null for a local branch.</param>
public sealed record GitBranch(string Name, bool Current, string? Upstream, string? Remote = null)
{
    /// <summary>How git and the select name it: the remote in front for a remote's branch.</summary>
    public string Ref => Remote is null ? Name : $"{Remote}/{Name}";
}

public static class GitBranches
{
    private const string HeadsPrefix = "refs/heads/";
    private const string RemotesPrefix = "refs/remotes/";

    /// <summary>The format <see cref="Parse"/> reads, for `git for-each-ref` over refs/heads and refs/remotes.</summary>
    public const string Format = "%(refname)%09%(refname:short)%09%(HEAD)%09%(upstream:short)";

    /// <summary>
    /// Parses `git for-each-ref --format=<see cref="Format"/> refs/heads refs/remotes`: one ref
    /// per line, tab separated - the full ref name, its short name, a star for the checked-out
    /// branch, and the upstream a local branch tracks. A local branch is named by what follows
    /// refs/heads/, never by the short column, which git disambiguates to `heads/x` when a tag
    /// shares the name. A remote's branch is named without the remote, which is the first path
    /// segment under refs/remotes/ - a remote name holds no slash, the reading VS Code's git
    /// extension makes as well - and a remote's HEAD pointer is not a branch and is left out.
    /// </summary>
    public static IReadOnlyList<GitBranch> Parse(string output)
    {
        ArgumentNullException.ThrowIfNull(output);

        var branches = new List<GitBranch>();
        foreach (var raw in output.Split('\n'))
        {
            var line = raw.TrimEnd('\r');
            if (line.Length == 0)
            {
                continue;
            }

            var columns = line.Split('\t');
            var refName = columns[0].Trim();
            var current = columns.Length > 2 && columns[2].Trim() == "*";
            var tracked = columns.Length > 3 ? columns[3].Trim() : string.Empty;
            var upstream = tracked.Length > 0 ? tracked : null;

            if (refName.StartsWith(HeadsPrefix, StringComparison.Ordinal))
            {
                var name = refName[HeadsPrefix.Length..];
                if (name.Length > 0)
                {
                    branches.Add(new GitBranch(name, current, upstream));
                }
            }
            else if (refName.StartsWith(RemotesPrefix, StringComparison.Ordinal))
            {
                var rest = refName[RemotesPrefix.Length..];
                var slash = rest.IndexOf('/', StringComparison.Ordinal);
                if (slash <= 0 || slash == rest.Length - 1)
                {
                    continue;
                }

                var remote = rest[..slash];
                var name = rest[(slash + 1)..];
                if (name != "HEAD")
                {
                    branches.Add(new GitBranch(name, false, null, remote));
                }
            }
        }

        return branches;
    }

    /// <summary>
    /// The branches a select offers: every local branch, then the ones only a remote has. A
    /// remote's branch that a local branch already answers to by name is dropped, because the
    /// local one is what a checkout of that name reaches, and listing both would offer one branch
    /// twice.
    /// </summary>
    public static IReadOnlyList<GitBranch> Choices(IReadOnlyList<GitBranch> branches)
    {
        ArgumentNullException.ThrowIfNull(branches);

        var local = new HashSet<string>(
            branches.Where(one => one.Remote is null).Select(one => one.Name), StringComparer.Ordinal);
        var choices = new List<GitBranch>(branches.Where(one => one.Remote is null));
        choices.AddRange(branches.Where(one => one.Remote is not null && !local.Contains(one.Name)));
        return choices;
    }
}
