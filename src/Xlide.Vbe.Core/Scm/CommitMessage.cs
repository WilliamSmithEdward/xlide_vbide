namespace Xlide.Vbe.Core.Scm;

/// <summary>What the pane needs of one change-log round to suggest a commit message.</summary>
/// <param name="Label">The round's label, or null when nobody labelled it.</param>
/// <param name="Entries">
/// Each module the round touched: what happened, its name, and its earlier name for a rename.
/// </param>
public sealed record RoundSummary(
    string Author,
    string? Label,
    IReadOnlyList<(string Kind, string Module, string? From)> Entries);

/// <summary>
/// A commit message from the change log, so an agent's work becomes a reviewable commit.
/// </summary>
public static class CommitMessage
{
    /// <summary>The most a subject line should carry; git shows more, nobody reads more.</summary>
    private const int MostCharacters = 200;

    /// <summary>
    /// A suggestion from the rounds since the last commit, newest first. When any round is
    /// labelled, the distinct labels joined by "; " - a label is what the author said the work
    /// was, which beats anything derived. Otherwise per author, "author: Written Ledger, Added
    /// Foo (was Bar)", each module once per author with the strongest word for what happened to
    /// it. Empty when there are no rounds. One line, never longer than 200 characters.
    /// </summary>
    public static string Suggest(IReadOnlyList<RoundSummary> rounds)
    {
        ArgumentNullException.ThrowIfNull(rounds);

        var labels = new List<string>();
        foreach (var round in rounds)
        {
            var label = OneLine(round.Label);
            if (label.Length > 0 && !labels.Contains(label, StringComparer.OrdinalIgnoreCase))
            {
                labels.Add(label);
            }
        }

        var line = labels.Count > 0
            ? string.Join("; ", labels)
            : string.Join("; ", ByAuthor(rounds));

        if (line.Length <= MostCharacters)
        {
            return line;
        }

        return line[..(MostCharacters - 3)].TrimEnd() + "...";
    }

    private static IEnumerable<string> ByAuthor(IReadOnlyList<RoundSummary> rounds)
    {
        var authors = new List<(string Author, List<Touched> Modules)>();
        foreach (var round in rounds)
        {
            var author = OneLine(round.Author);
            var at = authors.FindIndex(one => Same(one.Author, author));
            if (at < 0)
            {
                at = authors.Count;
                authors.Add((author, []));
            }

            foreach (var (kind, module, from) in round.Entries)
            {
                var modules = authors[at].Modules;
                var already = modules.FindIndex(one => Same(one.Module, module));
                if (already < 0)
                {
                    modules.Add(new Touched(module, kind, from));
                    continue;
                }

                // The same module in two rounds is one item: a module added and then written is
                // still an addition to the repository, and a removal is a removal whatever came
                // before it.
                var held = modules[already];
                modules[already] = new Touched(
                    held.Module,
                    Rank(kind) > Rank(held.Kind) ? kind : held.Kind,
                    held.From ?? from);
            }
        }

        foreach (var (author, modules) in authors)
        {
            if (modules.Count == 0)
            {
                continue;
            }

            yield return $"{author}: {string.Join(", ", modules.Select(Describe))}";
        }
    }

    private static string Describe(Touched touched)
    {
        // "written" on the wire, "Written" from the enum: one spelling on the way out.
        var kind = OneLine(touched.Kind);
        var word = kind.Length == 0
            ? string.Empty
            : char.ToUpperInvariant(kind[0]) + kind[1..].ToLowerInvariant() + " ";
        var from = touched.From is { Length: > 0 } was ? $" (was {OneLine(was)})" : string.Empty;
        return $"{word}{OneLine(touched.Module)}{from}";
    }

    private static int Rank(string kind) => OneLine(kind).ToUpperInvariant() switch
    {
        "REMOVED" => 3,
        "ADDED" => 2,
        "RENAMED" => 1,
        _ => 0,
    };

    private static bool Same(string left, string right) =>
        string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

    private static string OneLine(string? text) =>
        (text ?? string.Empty)
            .Replace("\r\n", " ", StringComparison.Ordinal)
            .Replace('\n', ' ')
            .Replace('\r', ' ')
            .Trim();

    private sealed record Touched(string Module, string Kind, string? From);
}
