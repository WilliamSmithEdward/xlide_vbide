using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// An event handler in a module's text, found or written the way the native code window's
/// Object and Procedure lists write one (#101).
///
/// The placement is the VBE's, measured 2026-10-12 against its own CodeModule.CreateEventProc on
/// a disposable workbook, whole module text compared: a new handler goes before the first event
/// handler in the module - of any of its objects, ordinary procedures not counted - whose name
/// sorts after its own, at the first line of that handler's leading comments, so the handlers
/// stay in name order however they were written; with none after it, at the end of the module.
/// The declaration is the VBE's too: <c>Private Sub Object_Event(params)</c>, one blank line,
/// <c>End Sub</c>.
/// </summary>
public static class EventHandlers
{
    /// <summary>The line (0-based) of a handler's declaration in any spelling, or -1 when absent.</summary>
    public static int Find(string source, string objectName, string eventName)
    {
        var declaration = Declaration($"{Regex.Escape(objectName)}_{Regex.Escape(eventName)}");
        return Array.FindIndex(Lines(source), line => declaration.IsMatch(line));
    }

    /// <summary>
    /// The module's text with a new handler placed where the native editor places it, and the
    /// 1-based line of its declaration. <paramref name="handlers"/> names every handler the
    /// module's objects can have (<c>Object_Event</c>), which tells a handler from an ordinary
    /// procedure. The caller has checked that the handler is not there.
    /// </summary>
    public static (string Text, int Header) Place(
        string source, string objectName, string eventName, string parameters, IReadOnlySet<string> handlers)
    {
        var name = $"{objectName}_{eventName}";
        var stub = new[] { $"Private Sub {name}({parameters})", string.Empty, "End Sub" };
        var lines = Lines(source);
        var anySub = Declaration(@"(\w+)");
        var next = Array.FindIndex(lines, line =>
            anySub.Match(line) is { Success: true } match
            && handlers.Contains(match.Groups[1].Value)
            && string.Compare(match.Groups[1].Value, name, StringComparison.InvariantCultureIgnoreCase) > 0);

        if (next < 0)
        {
            var bare = source.TrimEnd('\r', '\n');
            if (bare.Length == 0)
            {
                return (string.Join("\r\n", stub) + "\r\n", 1);
            }
            return ($"{bare}\r\n\r\n{string.Join("\r\n", stub)}\r\n", Lines(bare).Length + 2);
        }

        // Up over the handler's leading comments and the blank lines among them, then back down
        // past the blank lines that close whatever stands above.
        var at = next;
        while (at > 0 && (lines[at - 1].Trim().Length == 0 || IsComment(lines[at - 1])))
        {
            at--;
        }
        while (at < next && lines[at].Trim().Length == 0)
        {
            at++;
        }

        var placed = lines.Take(at).Concat(stub).Append(string.Empty).Concat(lines.Skip(at));
        return (string.Join("\r\n", placed), at + 1);
    }

    private static Regex Declaration(string name) => new(
        $@"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+{name}\s*\(",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static string[] Lines(string source) => source.Replace("\r\n", "\n").Split('\n');

    private static bool IsComment(string line)
    {
        var trimmed = line.TrimStart();
        return trimmed.StartsWith('\'')
            || (trimmed.StartsWith("Rem", StringComparison.OrdinalIgnoreCase)
                && (trimmed.Length == 3 || char.IsWhiteSpace(trimmed[3])));
    }
}
