using System.Globalization;

namespace Xlide.Vbe.Core.Scm;

/// <summary>What `git status` says about the branch and the working tree.</summary>
/// <param name="Oid">The commit the branch is at, or null on an unborn branch.</param>
/// <param name="Head">The branch name, or "(detached)".</param>
/// <param name="Upstream">The tracked remote branch, or null when there is none.</param>
/// <param name="Ahead">Commits the branch has that its upstream lacks.</param>
/// <param name="Behind">Commits the upstream has that the branch lacks.</param>
/// <param name="Conflicts">Paths still unmerged.</param>
/// <param name="Changed">
/// Paths of ordinary and renamed changes, for the outside section's cross-check.
/// </param>
public sealed record GitBranchState(
    string? Oid, string Head, string? Upstream, int Ahead, int Behind,
    IReadOnlyList<string> Conflicts, IReadOnlyList<string> Changed);

/// <summary>Reads `git status --porcelain=v2 --branch -z` output.</summary>
public static class GitStatus
{
    private const string OidHeader = "# branch.oid ";
    private const string HeadHeader = "# branch.head ";
    private const string UpstreamHeader = "# branch.upstream ";
    private const string AheadBehindHeader = "# branch.ab ";

    /// <summary>
    /// Parses the status. With `-z` every entry ends in NUL and paths are never quoted, which is
    /// the only form in which a path holding a space or a quote comes back as itself. A renamed
    /// entry ("2") is followed by a second NUL-terminated token, its original path, which is why
    /// the entries cannot simply be split and read one by one.
    /// </summary>
    public static GitBranchState Parse(string output)
    {
        ArgumentNullException.ThrowIfNull(output);

        string? oid = null;
        var head = string.Empty;
        string? upstream = null;
        var ahead = 0;
        var behind = 0;
        var conflicts = new List<string>();
        var changed = new List<string>();

        var tokens = output.Split('\0');
        for (var at = 0; at < tokens.Length; at++)
        {
            var entry = tokens[at].TrimEnd('\n', '\r');
            if (entry.Length == 0)
            {
                continue;
            }

            if (entry.StartsWith(OidHeader, StringComparison.Ordinal))
            {
                var value = entry[OidHeader.Length..];
                oid = string.Equals(value, "(initial)", StringComparison.Ordinal) ? null : value;
            }
            else if (entry.StartsWith(HeadHeader, StringComparison.Ordinal))
            {
                head = entry[HeadHeader.Length..];
            }
            else if (entry.StartsWith(UpstreamHeader, StringComparison.Ordinal))
            {
                upstream = entry[UpstreamHeader.Length..];
            }
            else if (entry.StartsWith(AheadBehindHeader, StringComparison.Ordinal))
            {
                (ahead, behind) = AheadBehind(entry[AheadBehindHeader.Length..]);
            }
            else if (entry[0] == '1')
            {
                // "1 XY sub mH mI mW hH hI path": eight fields, then the path with any spaces.
                changed.Add(AfterSpaces(entry, 8));
            }
            else if (entry[0] == '2')
            {
                // "2 XY sub mH mI mW hH hI Xscore path" then the original path as its own token.
                changed.Add(AfterSpaces(entry, 9));
                at++;
            }
            else if (entry[0] == 'u')
            {
                // "u XY sub m1 m2 m3 mW h1 h2 h3 path".
                conflicts.Add(AfterSpaces(entry, 10));
            }
        }

        return new GitBranchState(oid, head, upstream, ahead, behind, conflicts, changed);
    }

    /// <summary>"+A -B" as git prints it.</summary>
    private static (int Ahead, int Behind) AheadBehind(string text)
    {
        var culture = CultureInfo.InvariantCulture;
        var ahead = 0;
        var behind = 0;
        foreach (var part in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (part.Length < 2
                || !int.TryParse(part[1..], NumberStyles.None, culture, out var count))
            {
                continue;
            }

            if (part[0] == '+')
            {
                ahead = count;
            }
            else if (part[0] == '-')
            {
                behind = count;
            }
        }

        return (ahead, behind);
    }

    /// <summary>
    /// The text after the n-th space, which is where a path with spaces of its own starts.
    /// </summary>
    private static string AfterSpaces(string entry, int spaces)
    {
        var at = -1;
        for (var found = 0; found < spaces; found++)
        {
            at = entry.IndexOf(' ', at + 1);
            if (at < 0)
            {
                return string.Empty;
            }
        }

        return entry[(at + 1)..];
    }
}
