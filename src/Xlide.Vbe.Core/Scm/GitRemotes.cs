namespace Xlide.Vbe.Core.Scm;

/// <summary>One remote, by the URL it is fetched from.</summary>
public sealed record GitRemote(string Name, string Url);

/// <summary>Reads `git remote -v` output.</summary>
public static class GitRemotes
{
    private const string FetchSuffix = " (fetch)";

    /// <summary>
    /// Parses `git remote -v`: one line per remote and direction, `name`, a tab, the URL, a
    /// space and `(fetch)` or `(push)`. Only the fetch lines are kept, one per remote, in git's
    /// order, because the push URL is the fetch URL unless somebody set it apart.
    /// </summary>
    public static IReadOnlyList<GitRemote> Parse(string output)
    {
        ArgumentNullException.ThrowIfNull(output);

        var remotes = new List<GitRemote>();
        foreach (var raw in output.Split('\n'))
        {
            var line = raw.TrimEnd('\r');
            if (!line.EndsWith(FetchSuffix, StringComparison.Ordinal))
            {
                continue;
            }

            var tab = line.IndexOf('\t', StringComparison.Ordinal);
            if (tab <= 0)
            {
                continue;
            }

            var name = line[..tab];
            var url = line[(tab + 1)..^FetchSuffix.Length].Trim();
            if (url.Length == 0 || remotes.Exists(one => string.Equals(one.Name, name, StringComparison.Ordinal)))
            {
                continue;
            }

            remotes.Add(new GitRemote(name, url));
        }

        return remotes;
    }

    /// <summary>
    /// The remote a push goes to and a remote URL is read from: origin when there is one, else
    /// the first there is, else null. The same choice a first push makes for its upstream.
    /// </summary>
    public static GitRemote? Primary(IReadOnlyList<GitRemote> remotes)
    {
        ArgumentNullException.ThrowIfNull(remotes);

        foreach (var remote in remotes)
        {
            if (remote.Name == "origin")
            {
                return remote;
            }
        }

        return remotes.Count > 0 ? remotes[0] : null;
    }
}
