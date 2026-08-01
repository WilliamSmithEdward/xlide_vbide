using System.Text;
using Xlide.Vba.Analysis.Lexing;
using Xunit;

namespace Xlide.Vba.Analysis.Tests;

/// <summary>
/// The lexer is the floor everything else stands on. If it loses a character or misplaces an offset,
/// every layer above reports the wrong location for the rest of the file, so the first thing proven
/// here is that nothing is lost.
/// </summary>
public class LexerTests
{
    /// <summary>
    /// Rebuilds the source from the token stream. Any difference means a character was dropped,
    /// duplicated, or reordered.
    /// </summary>
    private static string RoundTrip(string source)
    {
        var module = Lexer.Tokenize(source);
        var builder = new StringBuilder(source.Length);

        foreach (var token in module.Tokens)
        {
            foreach (var trivia in token.LeadingTrivia(module.Trivia))
            {
                builder.Append(trivia.Text(source));
            }

            builder.Append(token.Text(source));
        }

        return builder.ToString();
    }

    private static TokenKind[] KindsOf(string source) =>
        Lexer.Tokenize(source).Tokens
            .Where(t => t.Kind != TokenKind.EndOfFile)
            .Select(t => t.Kind)
            .ToArray();

    [Theory]
    [InlineData("")]
    [InlineData("Sub Main()\r\nEnd Sub\r\n")]
    [InlineData("    Dim n As Long   \r\n")]
    [InlineData("x = 1 ' trailing comment\r\n")]
    [InlineData("Rem an old style comment\r\n")]
    [InlineData("s = \"a \"\" quoted \"\" thing\"\r\n")]
    [InlineData("d = #2026-08-01#\r\n")]
    [InlineData("v = &HFF& + &O17\r\n")]
    [InlineData("Call Foo(a, b) : Call Bar(c)\r\n")]
    [InlineData("Total = _\r\n    1 + 2\r\n")]
    [InlineData("#If Win64 Then\r\n#End If\r\n")]
    [InlineData("[My Odd Name] = 1\r\n")]
    [InlineData("\r\n\r\n\r\n")]
    [InlineData("   ")]
    [InlineData("no trailing newline")]
    [InlineData("mixed\rline\nendings\r\n")]
    public void ReproducesTheSourceExactly(string source)
    {
        Assert.Equal(source, RoundTrip(source));
    }

    [Fact]
    public void ReproducesARealisticModuleExactly()
    {
        var source = string.Join("\r\n",
            "Option Explicit",
            "",
            "' Computes a running total.",
            "Public Function Total(ByVal values As Variant) As Double",
            "    Dim i As Long",
            "    Dim sum As Double",
            "",
            "    For i = LBound(values) To UBound(values)",
            "        sum = sum + values(i)   ' accumulate",
            "    Next i",
            "",
            "    Total = sum",
            "End Function",
            "");

        Assert.Equal(source, RoundTrip(source));
    }

    [Fact]
    public void EndsEveryStreamWithAnEndTokenThatSpansNothing()
    {
        var module = Lexer.Tokenize("x = 1");
        var last = module.Tokens[^1];

        Assert.Equal(TokenKind.EndOfFile, last.Kind);
        Assert.Equal(0, last.Length);
        Assert.Equal(module.Source.Length, last.Start);
    }

    [Fact]
    public void CarriesTrailingWhitespaceOnTheEndToken()
    {
        var module = Lexer.Tokenize("x = 1   ");
        var last = module.Tokens[^1];

        Assert.Equal(TokenKind.EndOfFile, last.Kind);
        Assert.Equal(1, last.TriviaCount);
        Assert.Equal("   ", last.LeadingTrivia(module.Trivia)[0].Text("x = 1   ").ToString());
    }

    [Fact]
    public void TreatsAContinuedLineAsOneLogicalLine()
    {
        // The physical line ends, the logical line does not, so no newline token is produced.
        var kinds = KindsOf("a = 1 + _\r\n    2\r\n");

        Assert.Equal(
            [TokenKind.Identifier, TokenKind.Operator, TokenKind.IntegerLiteral, TokenKind.Operator,
             TokenKind.IntegerLiteral, TokenKind.Newline],
            kinds);
    }

    [Fact]
    public void RecordsAContinuationAsTrivia()
    {
        var module = Lexer.Tokenize("a = 1 + _\r\n    2");
        var all = module.Trivia;

        Assert.Contains(all, t => t.Kind == TriviaKind.LineContinuation);
    }

    [Theory]
    [InlineData("1", TokenKind.IntegerLiteral)]
    [InlineData("42&", TokenKind.IntegerLiteral)]
    [InlineData("&HFF", TokenKind.IntegerLiteral)]
    [InlineData("&O777", TokenKind.IntegerLiteral)]
    [InlineData("1.5", TokenKind.FloatLiteral)]
    [InlineData(".5", TokenKind.FloatLiteral)]
    [InlineData("1E10", TokenKind.FloatLiteral)]
    [InlineData("1.5E-3", TokenKind.FloatLiteral)]
    [InlineData("1#", TokenKind.FloatLiteral)]
    [InlineData("1!", TokenKind.FloatLiteral)]
    [InlineData("\"text\"", TokenKind.StringLiteral)]
    [InlineData("#2026-01-01#", TokenKind.DateLiteral)]
    [InlineData("[Odd Name]", TokenKind.BracketedIdentifier)]
    public void ClassifiesLiterals(string source, TokenKind expected)
    {
        Assert.Equal(expected, KindsOf(source)[0]);
    }

    [Fact]
    public void TreatsAMemberAccessDotAsPunctuationNotAFloat()
    {
        // "a.b" must not lex as a float, or every member access becomes a number.
        Assert.Equal(
            [TokenKind.Identifier, TokenKind.Punctuation, TokenKind.Identifier],
            KindsOf("a.b"));
    }

    [Fact]
    public void SeparatesTheStatementSeparatorFromTheNamedArgumentMarker()
    {
        Assert.Equal(TokenKind.Colon, KindsOf("a : b")[1]);
        Assert.Equal(TokenKind.Operator, KindsOf("Foo x:=1")[2]);
    }

    [Theory]
    [InlineData("Sub")]
    [InlineData("sub")]
    [InlineData("SUB")]
    [InlineData("SuB")]
    public void RecognisesKeywordsRegardlessOfCase(string word)
    {
        Assert.Equal(TokenKind.Keyword, KindsOf(word)[0]);
    }

    [Fact]
    public void ReportsTheCapitalizationTheEditorShows()
    {
        var module = Lexer.Tokenize("dim x as long");

        Assert.Equal("Dim", module.CanonicalTextOf(module.Tokens[0]));
        Assert.Equal("As", module.CanonicalTextOf(module.Tokens[2]));
        Assert.Equal("Long", module.CanonicalTextOf(module.Tokens[3]));
    }

    [Fact]
    public void HasNoCanonicalFormForAnIdentifier()
    {
        var module = Lexer.Tokenize("myVariable");
        Assert.Null(module.CanonicalTextOf(module.Tokens[0]));
    }

    [Fact]
    public void TreatsRemAsAComment()
    {
        Assert.Equal([TokenKind.Comment, TokenKind.Newline], KindsOf("Rem this is ignored\r\n"));
    }

    [Fact]
    public void StopsAnUnterminatedStringAtTheLineEnd()
    {
        // Consuming the terminator would swallow the newline and make the rest of the file a string.
        Assert.Equal(
            [TokenKind.Identifier, TokenKind.Operator, TokenKind.StringLiteral, TokenKind.Newline,
             TokenKind.Identifier],
            KindsOf("s = \"unterminated\r\nx"));
    }

    [Fact]
    public void DoesNotSwallowTheFileOnAnUnterminatedBracketedName()
    {
        var kinds = KindsOf("[unterminated\r\nx = 1\r\n");

        Assert.Equal(TokenKind.Unknown, kinds[0]);
        Assert.Contains(TokenKind.Newline, kinds);
    }

    [Fact]
    public void DistinguishesADirectiveFromADateLiteral()
    {
        Assert.Equal(TokenKind.Directive, KindsOf("#If Win64 Then\r\n")[0]);
        Assert.Equal(TokenKind.DateLiteral, KindsOf("#2026-01-01#")[0]);
    }

    [Fact]
    public void ReportsZeroBasedLineAndColumn()
    {
        var module = Lexer.Tokenize("a\r\nbb\r\n  c");
        var tokens = module.Tokens.Where(t => t.Kind == TokenKind.Identifier).ToArray();

        Assert.Equal((0, 0), (tokens[0].Line, tokens[0].Character));
        Assert.Equal((1, 0), (tokens[1].Line, tokens[1].Character));
        Assert.Equal((2, 2), (tokens[2].Line, tokens[2].Character));
    }

    [Fact]
    public void NeverThrowsOnArbitraryInput()
    {
        // A module being typed into is malformed most of the time. Giving up is not an option.
        string[] hostile =
        [
            "\0\0\0", "\"", "[", "#", "&H", "_", "'", ":=", "1.2.3.4", "\r", "\n",
            "Sub", "End Sub Without Start", "\"\"\"\"\"", "&&&", "###",
        ];

        foreach (var source in hostile)
        {
            var module = Lexer.Tokenize(source);
            Assert.Equal(source, RoundTrip(source));
            Assert.Equal(TokenKind.EndOfFile, module.Tokens[^1].Kind);
        }
    }
}
