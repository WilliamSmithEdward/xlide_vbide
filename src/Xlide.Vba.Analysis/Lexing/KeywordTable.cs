using System.Collections.Frozen;

namespace Xlide.Vba.Analysis.Lexing;

/// <summary>
/// Recognises VBA keywords and supplies the capitalization the editor displays.
///
/// Lookups happen once per identifier-shaped token, which on a large module is tens of thousands of
/// times per pass, so the tables are frozen at start-up and the hot path avoids allocating: a token
/// is matched from its span rather than from a lowercased copy of it.
/// </summary>
public static partial class KeywordTable
{
    /// <summary>
    /// Holds the built tables in a nested type so they cannot be built before the generated word
    /// lists exist.
    ///
    /// The lists are declared in the generated half of this partial class. Static field
    /// initializers run in declaration order within a type, and the order across the files of a
    /// partial type is not defined, so a table built from a field initializer here can and did read
    /// the lists while they were still null. Touching a nested type instead forces this type to
    /// finish initializing first, which is guaranteed rather than incidental.
    /// </summary>
    private static class Tables
    {
        internal static readonly FrozenDictionary<string, string> CanonicalByLower = BuildCanonical();
        internal static readonly FrozenSet<string> ReservedLower = BuildReserved();
    }

    private static FrozenDictionary<string, string> BuildCanonical()
    {
        var map = new Dictionary<string, string>(
            ReservedWords.Length + ContextualWords.Length,
            StringComparer.OrdinalIgnoreCase);

        foreach (var word in ReservedWords)
        {
            map[word] = word;
        }

        foreach (var word in ContextualWords)
        {
            map.TryAdd(word, word);
        }

        return map.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
    }

    private static FrozenSet<string> BuildReserved() =>
        ReservedWords.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>Canonical capitalization for a word, or null when it is not a keyword.</summary>
    public static string? Canonical(ReadOnlySpan<char> word) =>
        Tables.CanonicalByLower.GetAlternateLookup<ReadOnlySpan<char>>().TryGetValue(word, out var canonical)
            ? canonical
            : null;

    /// <summary>True when the word is a keyword of any kind.</summary>
    public static bool IsKeyword(ReadOnlySpan<char> word) =>
        Tables.CanonicalByLower.GetAlternateLookup<ReadOnlySpan<char>>().ContainsKey(word);

    /// <summary>
    /// True when the word can never be a user-defined name. Contextual keywords are capitalized by
    /// the editor but remain usable as names, so they are keywords without being reserved.
    /// </summary>
    public static bool IsReserved(ReadOnlySpan<char> word) =>
        Tables.ReservedLower.GetAlternateLookup<ReadOnlySpan<char>>().Contains(word);

    /// <summary>Number of reserved identifiers. Exposed so tests can assert the table was generated.</summary>
    public static int ReservedCount => Tables.ReservedLower.Count;
}
