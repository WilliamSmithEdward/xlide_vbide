using System;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// What a line typed into the Immediate panel asks for, as far as break mode can answer it.
///
/// While execution is stopped the panel cannot evaluate: it evaluates by adding a procedure to
/// the project, and that would reset the debugger (#21). But the commonest line typed at a
/// breakpoint is <c>? name</c>, and the value of a local is already on screen in the Locals
/// panel, read off the editor's own Locals window. So a line that prints exactly one bare name is
/// recognised here and answered from there; anything with an operator, a member, a call or a
/// literal in it is still an evaluation, and still declined with a reason.
/// </summary>
public static class ImmediateLine
{
    // "? name", "?name", "Print name", "Debug.Print name" - the print forms the editor's own
    // window takes - followed by one identifier of any script, and nothing else.
    private static readonly Regex PrintedName = new(
        @"^\s*(?:\?|(?:Debug\s*\.\s*)?Print\b)\s*(?<name>\p{L}[\p{L}\p{N}_]*)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>The one name the line prints, or null when the line is anything else.</summary>
    public static string? NameToPrint(string? line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }

        var match = PrintedName.Match(line);
        return match.Success ? match.Groups["name"].Value : null;
    }

    /// <summary>
    /// A Locals value the way the Immediate window would print it: a String loses the quotes the
    /// Locals window draws around it, with its doubled inner quotes undoubled; everything else is
    /// printed as shown. An object row shows no value in Locals, so its type stands in.
    /// </summary>
    public static string AsPrinted(string value, string type)
    {
        ArgumentNullException.ThrowIfNull(value);
        ArgumentNullException.ThrowIfNull(type);

        if (string.Equals(type, "String", StringComparison.OrdinalIgnoreCase)
            && value.Length >= 2 && value[0] == '"' && value[^1] == '"')
        {
            return value[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
        }

        return value.Length == 0 ? $"<{type}>" : value;
    }
}
