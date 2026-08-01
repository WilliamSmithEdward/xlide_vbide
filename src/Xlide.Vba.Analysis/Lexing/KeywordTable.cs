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
            CasedWords.Length + ContextualWords.Length,
            StringComparer.OrdinalIgnoreCase);

        foreach (var word in CasedWords)
        {
            map[word] = word;
        }

        // Cased words win a spelling conflict, because they are the closed set defined by the
        // specification and a contextual word is only meaningful in one statement.
        foreach (var word in ContextualWords)
        {
            map.TryAdd(word, word);
        }

        return map.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Everything that can never be a user-defined name: the cased keywords plus the words reserved
    /// for the implementation. The second group is deliberately absent from the casing table, since
    /// those words lex as identifiers even though they cannot be declared.
    /// </summary>
    private static FrozenSet<string> BuildReserved() =>
        CasedWords.Concat(ImplementationWords).ToFrozenSet(StringComparer.OrdinalIgnoreCase);

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
