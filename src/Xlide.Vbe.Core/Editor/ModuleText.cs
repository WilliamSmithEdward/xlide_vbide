using System.Globalization;
using System.Text;

namespace Xlide.Vbe.Core.Editor;

/// <summary>A character the host would not keep: the text of it, its code point, and its line.</summary>
public readonly record struct LostCharacter(string Text, int CodePoint, int Line)
{
    /// <summary>How to name it to a developer: the character where that helps, always the number.</summary>
    public string Describe() => Rune.TryCreate(CodePoint, out var rune)
        && !Rune.IsControl(rune) && !Rune.IsWhiteSpace(rune)
        && CharUnicodeInfo.GetUnicodeCategory(Text, 0) != UnicodeCategory.NonSpacingMark
            ? $"'{Text}' (U+{CodePoint:X4})"
            : $"U+{CodePoint:X4}";
}

/// <summary>
/// What the host's editor will and will not hold in a module's text.
///
/// Its limits are not documented anywhere this product can cite, and the ones that are widely
/// repeated are the wrong ones: the 65,534-line ceiling is real and is not what a write runs into
/// first. So they are measured, and the measurement lives here with the number.
/// </summary>
public static class ModuleText
{
    /// <summary>
    /// The most the editor will hold in one line.
    ///
    /// Measured 2026-08-09 by writing a line of each width and reading it back: 1,020 and 1,022
    /// come back identical, 1,023 and above come back as two lines. The break carries NO line
    /// continuation, so a statement is cut in half and a string literal is left unterminated - a
    /// 2,018 character Debug.Print became a 1,023 character fragment followed by a 995 character
    /// one, and nothing anywhere said so.
    /// </summary>
    public const int LongestLine = 1022;

    /// <summary>
    /// The first character the host was handed and did not keep, or null when it kept them all.
    ///
    /// VBA stores module text in the system ANSI code page, not in Unicode, so a character outside
    /// that page never reaches this product: Excel converts it on the way in. What it converts to
    /// is not always obvious. Measured on code page 1252, 2026-08-09:
    ///
    ///   Cyrillic, Greek, Thai, CJK, emoji   become '?', which at least looks wrong
    ///   "cafe" + U+0301 (decomposed cafe)   becomes "cafe" + U+00B4, a DIFFERENT REAL CHARACTER
    ///
    /// The second is the dangerous one and it is the one a Western European developer meets: the
    /// precomposed spelling of the same word survives untouched, so the identical text is kept or
    /// silently altered according to its normalisation, and a repository authored on macOS is
    /// routinely decomposed. Nothing on screen distinguishes them.
    ///
    /// Asked as PRESENCE rather than equality, deliberately. The editor legitimately rewrites what
    /// it is given - completing a procedure's parentheses, respelling keywords - and all of that is
    /// ASCII, so comparing whole texts would refuse writes over the editor's own tidying. A
    /// non-ASCII character that went in and is nowhere in what came back was converted, and nothing
    /// else does that.
    ///
    /// By RUNE, so an astral character is reported as itself rather than as half of a surrogate
    /// pair that means nothing to anybody.
    /// </summary>
    public static LostCharacter? FirstCharacterLost(string intended, string kept)
    {
        var missing = new HashSet<int>();
        foreach (var rune in intended.EnumerateRunes())
        {
            if (rune.Value > 0x7F)
            {
                missing.Add(rune.Value);
            }
        }

        if (missing.Count == 0)
        {
            return null;
        }

        foreach (var rune in kept.EnumerateRunes())
        {
            missing.Remove(rune.Value);
        }

        if (missing.Count == 0)
        {
            return null;
        }

        var line = 1;
        foreach (var rune in intended.EnumerateRunes())
        {
            if (missing.Contains(rune.Value))
            {
                return new LostCharacter(rune.ToString(), rune.Value, line);
            }

            if (rune.Value == '\n')
            {
                line++;
            }
        }

        return null;
    }

    /// <summary>
    /// The first line the editor would break, as a 1-based line number and its length, or null
    /// when every line fits.
    ///
    /// 1-based because that is the number the developer's editor shows them, and a complaint that
    /// names a different line than the one on screen is worse than one that names none.
    ///
    /// Counts the line as the editor would: the line break is not part of it, and a CRLF counts as
    /// one break rather than a trailing character on the line before.
    /// </summary>
    public static (int At, int Length)? FirstLineTooLong(string text)
    {
        var at = 1;
        var from = 0;

        while (true)
        {
            var breakAt = text.IndexOf('\n', from);
            var end = breakAt < 0 ? text.Length : breakAt;
            var length = end - from;

            if (length > 0 && text[end - 1] == '\r')
            {
                length--;
            }

            if (length > LongestLine)
            {
                return (at, length);
            }

            if (breakAt < 0)
            {
                return null;
            }

            from = breakAt + 1;
            at++;
        }
    }
}
