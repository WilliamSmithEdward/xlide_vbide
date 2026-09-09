namespace Xlide.Vbe.Core.Scm;

/// <summary>One local branch.</summary>
/// <param name="Current">True for the branch HEAD is on.</param>
/// <param name="Upstream">The remote branch it tracks, short form, or null.</param>
public sealed record GitBranch(string Name, bool Current, string? Upstream);

/// <summary>Reads `git for-each-ref` output over refs/heads.</summary>
public static class GitBranches
{
    /// <summary>
    /// Parses `git for-each-ref` over refs/heads with the format
    /// `%(refname:short)%09%(HEAD)%09%(upstream:short)`: one branch per line, tab-separated, the
    /// HEAD column a star or a space, the upstream column empty when the branch tracks nothing.
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
            var name = columns[0].Trim();
            if (name.Length == 0)
            {
                continue;
            }

            var current = columns.Length > 1 && columns[1].Trim() == "*";
            var tracked = columns.Length > 2 ? columns[2].Trim() : string.Empty;
            var upstream = tracked.Length > 0 ? tracked : null;
            branches.Add(new GitBranch(name, current, upstream));
        }

        return branches;
    }
}
