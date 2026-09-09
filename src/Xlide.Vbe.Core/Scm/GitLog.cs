namespace Xlide.Vbe.Core.Scm;

/// <summary>One file a commit touched.</summary>
/// <param name="Status">
/// A added, M modified, D deleted, R renamed with <paramref name="FromPath"/> set; anything else
/// is git's own letter.
/// </param>
/// <param name="Path">The path after the commit, from the repository root, forward slashes.</param>
/// <param name="FromPath">The path before the commit, for a rename or a copy.</param>
public sealed record GitFileChange(string Status, string Path, string? FromPath);

/// <summary>One commit as the history lists it.</summary>
/// <param name="ShortHash">The abbreviated hash; the wire calls it `short`.</param>
/// <param name="When">The author date, ISO 8601 with its offset, as git prints %aI.</param>
/// <param name="Body">
/// The message below the subject, without its trailing newline; empty when there is none.
/// </param>
public sealed record GitCommit(
    string Hash,
    string ShortHash,
    string Author, string Email, string When, string Subject,
    string Body, IReadOnlyList<GitFileChange> Files);

/// <summary>Reads `git log --format=&lt;LogFormat&gt; --name-status -M` output.</summary>
public static class GitLog
{
    private const char RecordStart = '\x1e';
    private const char BodyEnd = '\x1f';

    /// <summary>
    /// Parses the log. Records begin with 0x1E, fields split on NUL, the body ends at 0x1F, and the
    /// name-status lines follow until the next record. Not `-z`: with it git would end the
    /// name-status rows with NUL as well and the two delimiters would collide inside the body.
    /// Tolerates an empty output, a record without a body, and a trailing newline.
    /// </summary>
    public static IReadOnlyList<GitCommit> Parse(string output)
    {
        ArgumentNullException.ThrowIfNull(output);

        var commits = new List<GitCommit>();
        foreach (var record in Records(output))
        {
            if (string.IsNullOrWhiteSpace(record))
            {
                continue;
            }

            // The LAST body end, not the first: git stores a control character typed into a
            // message verbatim, and the name-status trailer can never hold one, so the last 0x1F
            // is always the one the format printed.
            var end = record.LastIndexOf(BodyEnd);
            var header = end >= 0 ? record[..end] : record;
            var trailer = end >= 0 ? record[(end + 1)..] : string.Empty;

            var fields = header.Split('\0');
            var hash = At(fields, 0).Trim();
            if (hash.Length == 0)
            {
                continue;
            }

            commits.Add(new GitCommit(
                hash,
                At(fields, 1).Trim(),
                At(fields, 2),
                At(fields, 3),
                At(fields, 4).Trim(),
                At(fields, 5),
                At(fields, 6).TrimEnd('\n', '\r'),
                ParseNameStatus(trailer)));
        }

        return commits;
    }

    private static string At(string[] fields, int index) =>
        index < fields.Length ? fields[index] : string.Empty;

    /// <summary>
    /// The records: each begins at a 0x1E that a hash and a NUL follow, which is how the format
    /// prints one. A 0x1E git stored inside a message begins nothing, so a control character in a
    /// body cannot manufacture a commit that takes the real one's files.
    /// </summary>
    private static IEnumerable<string> Records(string output)
    {
        var start = -1;
        for (var at = 0; at < output.Length; at++)
        {
            if (output[at] != RecordStart || !HashFollows(output, at + 1))
            {
                continue;
            }

            if (start >= 0)
            {
                yield return output[(start + 1)..at];
            }

            start = at;
        }

        if (start >= 0)
        {
            yield return output[(start + 1)..];
        }
    }

    private static bool HashFollows(string output, int from)
    {
        var length = 0;
        while (from + length < output.Length && char.IsAsciiHexDigitLower(output[from + length]))
        {
            length++;
        }

        return length >= 40 && from + length < output.Length && output[from + length] == '\0';
    }

    /// <summary>
    /// The rows after a commit: a status, a tab and the path; or for a rename or a copy the
    /// status with git's similarity score, then the old and the new path, tab-separated. The
    /// letter alone is the status.
    /// </summary>
    private static List<GitFileChange> ParseNameStatus(string trailer)
    {
        var files = new List<GitFileChange>();
        foreach (var raw in trailer.Split('\n'))
        {
            var line = raw.TrimEnd('\r');
            var parts = line.Split('\t');
            if (parts.Length < 2 || parts[0].Length == 0)
            {
                continue;
            }

            var status = parts[0][..1];
            if (status is "R" or "C" && parts.Length >= 3)
            {
                files.Add(new GitFileChange(status, parts[2], parts[1]));
            }
            else
            {
                files.Add(new GitFileChange(status, parts[^1], null));
            }
        }

        return files;
    }
}
