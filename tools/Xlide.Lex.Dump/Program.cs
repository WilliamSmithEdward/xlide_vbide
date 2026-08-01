using System.Text;
using System.Text.Json;
using Xlide.Vba.Analysis.Lexing;

namespace Xlide.Lex.Dump;

/// <summary>
/// Prints one line of JSON per token for each file given, so the ported lexer's output can be
/// compared with the reference implementation's over the same corpus.
///
/// The comparison is what makes the port safe. Reading the ported code and judging it correct scales
/// badly and misses exactly the cases nobody thought to look at; running both over thousands of real
/// modules and diffing does not.
/// </summary>
internal static class Program
{
    /// <summary>
    /// Escapes only what the format requires. The default encoder also escapes characters that
    /// would be unsafe in a web page, which is correct for a web page and wrong here: VBA source is
    /// full of quotes and ampersands, and escaping them differently from the reference turns every
    /// string literal in the corpus into a spurious difference.
    /// </summary>
    private static readonly JsonSerializerOptions Options = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: xlide-lexdump <file> [file...]");
            return 2;
        }

        var output = Console.Out;

        foreach (var path in args)
        {
            string source;

            try
            {
                source = File.ReadAllText(path);
            }
            catch (IOException ex)
            {
                Console.Error.WriteLine($"{path}: {ex.Message}");
                return 1;
            }

            var module = Lexer.Tokenize(source);

            // The end token is an artefact of how this lexer keeps trailing trivia attached and has
            // no counterpart in the reference, so it is excluded from both the count and the stream.
            var counted = module.Tokens.Count(t => t.Kind != TokenKind.EndOfFile);

            output.WriteLine(JsonSerializer.Serialize(new { file = Path.GetFileName(path), tokens = counted }, Options));

            foreach (var token in module.Tokens)
            {
                // The end token exists only so trailing trivia has somewhere to live. The reference
                // has no equivalent, so it is not part of the comparison.
                if (token.Kind == TokenKind.EndOfFile)
                {
                    continue;
                }

                output.WriteLine(JsonSerializer.Serialize(new TokenLine(
                    Kind(token.Kind),
                    token.Start,
                    token.End,
                    token.Line,
                    token.Character,
                    module.TextOf(token).ToString(),
                    module.CanonicalTextOf(token)), Options));
            }
        }

        return 0;
    }

    /// <summary>
    /// Names match the reference implementation's, because the comparison is textual and a
    /// different spelling for the same concept would read as thousands of differences.
    /// </summary>
    private static string Kind(TokenKind kind) => kind switch
    {
        TokenKind.Newline => "newline",
        TokenKind.Comment => "comment",
        TokenKind.Keyword => "keyword",
        TokenKind.Identifier => "identifier",
        TokenKind.BracketedIdentifier => "bracketedIdentifier",
        TokenKind.IntegerLiteral => "integerLiteral",
        TokenKind.FloatLiteral => "floatLiteral",
        TokenKind.DateLiteral => "dateLiteral",
        TokenKind.StringLiteral => "stringLiteral",
        TokenKind.Operator => "operator",
        TokenKind.Punctuation => "punctuation",
        TokenKind.Colon => "colon",
        TokenKind.Directive => "directive",
        TokenKind.EndOfFile => "endOfFile",
        _ => "unknown",
    };

    private sealed record TokenLine(
        string kind,
        int start,
        int end,
        int line,
        int character,
        string text,
        string? canonical);
}
