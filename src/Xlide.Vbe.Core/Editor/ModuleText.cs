namespace Xlide.Vbe.Core.Editor;

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
