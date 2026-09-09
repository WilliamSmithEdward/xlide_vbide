using System.Globalization;
using Xlide.Vbe.Core.Sync;

namespace Xlide.Vbe.Core.Scm;

/// <summary>Who last touched one line, in the numbering of the text it was read from.</summary>
/// <param name="ShortHash">The abbreviated hash; the wire calls it `short`.</param>
/// <param name="When">The author date, ISO 8601 with its offset.</param>
/// <param name="Summary">The commit's subject.</param>
public sealed record BlameLine(
    int Line,
    string Hash,
    string ShortHash,
    string Author,
    string When,
    string Summary);

/// <summary>
/// Blame in the editor's numbering: committed lines, and the live lines no commit holds.
/// </summary>
public sealed record BlameMap(IReadOnlyList<BlameLine> Lines, IReadOnlyList<int> Uncommitted);

/// <summary>Reads `git blame --porcelain` output and maps it onto the live module.</summary>
public static class GitBlame
{
    /// <summary>
    /// Parses the porcelain into one entry per FINAL line, the file's own numbering. A group
    /// opens with a header of hash, original line and final line (and a count), carries the
    /// commit's facts the first time git mentions it (and often again), and ends with the line's
    /// text behind a tab. `boundary`, `previous` and `filename` are read past. ShortHash is the
    /// first seven characters of the hash.
    /// </summary>
    public static IReadOnlyList<BlameLine> ParsePorcelain(string output)
    {
        ArgumentNullException.ThrowIfNull(output);

        var lines = new List<BlameLine>();
        var facts = new Dictionary<string, CommitFacts>(StringComparer.Ordinal);
        string? hash = null;
        var finalLine = 0;

        foreach (var raw in output.Split('\n'))
        {
            var line = raw.TrimEnd('\r');
            if (line.StartsWith('\t'))
            {
                if (hash is not null)
                {
                    var known = facts[hash];
                    lines.Add(new BlameLine(
                        finalLine, hash, hash.Length > 7 ? hash[..7] : hash,
                        known.Author, known.When(), known.Summary));
                    hash = null;
                }

                continue;
            }

            if (line.Length == 0)
            {
                continue;
            }

            if (Header(line) is var (headerHash, headerLine))
            {
                hash = headerHash;
                finalLine = headerLine;
                if (!facts.ContainsKey(hash))
                {
                    facts[hash] = new CommitFacts();
                }

                continue;
            }

            if (hash is null)
            {
                continue;
            }

            var space = line.IndexOf(' ', StringComparison.Ordinal);
            var key = space < 0 ? line : line[..space];
            var value = space < 0 ? string.Empty : line[(space + 1)..];
            var current = facts[hash];
            switch (key)
            {
                case "author":
                    current.Author = value;
                    break;
                case "author-time":
                    current.Seconds = Seconds(value);
                    break;
                case "author-tz":
                    current.Zone = value;
                    break;
                case "summary":
                    current.Summary = value;
                    break;
                default:
                    break;
            }
        }

        return lines;
    }

    /// <summary>
    /// Maps the file's lines onto the live text with the comparison the rows are drawn with: an
    /// Equal pair carries the file line's blame to the live line, and every other live line -
    /// added, rewritten, or a pair that was never committed - is uncommitted. One mechanism
    /// accounts for the attribute header the file carries and the module does not, for an
    /// attribute line inside a procedure, and for local edits. Lines come back in live numbering,
    /// ascending.
    /// </summary>
    public static BlameMap MapToLive(
        IReadOnlyList<BlameLine> fileLines, string fileText, string liveText)
    {
        ArgumentNullException.ThrowIfNull(fileLines);
        ArgumentNullException.ThrowIfNull(fileText);
        ArgumentNullException.ThrowIfNull(liveText);

        var byFileLine = new Dictionary<int, BlameLine>();
        foreach (var line in fileLines)
        {
            byFileLine[line.Line] = line;
        }

        var lines = new List<BlameLine>();
        var uncommitted = new List<int>();

        // TRAILING NEWLINES OFF BOTH SIDES, as the rows' comparison takes them: the exported file
        // ends with one and the module's text does not, and left in, that phantom last line
        // stopped the comparison's common-tail trim from engaging, so a module past 2,000 lines
        // fell into the diff's over-cap path and read as uncommitted from its first line
        // (found in review, 2026-09-08). Only lines after the last content line go, so the
        // blame numbering is untouched.
        var comparison = ModuleSync.Diff(
            ModuleSync.NormaliseEol(fileText).TrimEnd('\n'),
            ModuleSync.NormaliseEol(liveText).TrimEnd('\n'));
        foreach (var pair in comparison)
        {
            if (pair.RightNumber is not int live)
            {
                continue;
            }

            if (pair.Kind == DiffKind.Equal
                && pair.LeftNumber is int file
                && byFileLine.TryGetValue(file, out var blame))
            {
                lines.Add(blame with { Line = live });
            }
            else
            {
                uncommitted.Add(live);
            }
        }

        lines.Sort((left, right) => left.Line.CompareTo(right.Line));
        uncommitted.Sort();
        return new BlameMap(lines, uncommitted);
    }

    /// <summary>
    /// author-time as git prints it: seconds since the epoch, or zero when unreadable.
    /// </summary>
    private static long Seconds(string value) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var seconds)
            ? seconds
            : 0;

    /// <summary>The hash and final line of a group header, or null for any other line.</summary>
    private static (string Hash, int FinalLine)? Header(string line)
    {
        var parts = line.Split(' ');
        if (parts.Length is not (3 or 4) || parts[0].Length < 40)
        {
            return null;
        }

        foreach (var letter in parts[0])
        {
            if (!char.IsAsciiHexDigitLower(letter) && !char.IsAsciiDigit(letter))
            {
                return null;
            }
        }

        var culture = CultureInfo.InvariantCulture;
        return int.TryParse(parts[1], NumberStyles.None, culture, out _)
            && int.TryParse(parts[2], NumberStyles.None, culture, out var finalLine)
            ? (parts[0], finalLine)
            : null;
    }

    /// <summary>What the porcelain says about one commit, filled in as its lines arrive.</summary>
    private sealed class CommitFacts
    {
        public string Author { get; set; } = string.Empty;

        public long Seconds { get; set; }

        public string Zone { get; set; } = "+0000";

        public string Summary { get; set; } = string.Empty;

        /// <summary>
        /// author-time and author-tz as one ISO 8601 stamp in the author's own offset.
        /// </summary>
        public string When()
        {
            var culture = CultureInfo.InvariantCulture;
            var offset = TimeSpan.Zero;
            if (Zone.Length == 5
                && int.TryParse(Zone[1..3], NumberStyles.None, culture, out var hours)
                && int.TryParse(Zone[3..5], NumberStyles.None, culture, out var minutes))
            {
                offset = new TimeSpan(hours, minutes, 0);
                if (Zone[0] == '-')
                {
                    offset = offset.Negate();
                }
            }

            // Z for a zero offset, as git's own %aI spells it, so a commit's date reads the same
            // from the log and from blame.
            return DateTimeOffset.FromUnixTimeSeconds(Seconds)
                .ToOffset(offset)
                .ToString(
                    offset == TimeSpan.Zero ? "yyyy-MM-dd'T'HH:mm:ss'Z'" : "yyyy-MM-dd'T'HH:mm:ssK",
                    CultureInfo.InvariantCulture);
        }
    }
}
