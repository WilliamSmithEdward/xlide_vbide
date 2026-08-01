namespace Xlide.Vba.Analysis.Lexing;

/// <summary>
/// Category of a lexical token, per MS-VBAL section 3.3.
/// </summary>
public enum TokenKind
{
    /// <summary>End of a logical line.</summary>
    Newline,

    /// <summary>An apostrophe comment or a Rem comment. Not syntactic, but preserved.</summary>
    Comment,

    /// <summary>A reserved identifier or contextual keyword.</summary>
    Keyword,

    /// <summary>An identifier that is not a reserved identifier.</summary>
    Identifier,

    /// <summary>A foreign name, written in brackets.</summary>
    BracketedIdentifier,

    /// <summary>An integer literal, including any type suffix.</summary>
    IntegerLiteral,

    /// <summary>A floating point literal, including any type suffix.</summary>
    FloatLiteral,

    /// <summary>A date literal, delimited by hashes.</summary>
    DateLiteral,

    /// <summary>A string literal.</summary>
    StringLiteral,

    /// <summary>An arithmetic, comparison, logical, or concatenation symbol.</summary>
    Operator,

    /// <summary>Structural punctuation: comma, dot, parentheses, semicolon.</summary>
    Punctuation,

    /// <summary>The statement separator. A named-argument marker is an operator, not this.</summary>
    Colon,

    /// <summary>A hash that begins a conditional compilation directive.</summary>
    Directive,

    /// <summary>Text matching no lexical rule. Never an error by itself.</summary>
    Unknown,

    /// <summary>
    /// Always the last token, spanning nothing. It exists so trivia at the end of a file has a
    /// token to lead, which keeps the stream lossless without inventing a token that claims to
    /// cover source it does not.
    /// </summary>
    EndOfFile,
}

/// <summary>Insignificant text attached to the token that follows it.</summary>
public enum TriviaKind
{
    Whitespace,

    /// <summary>Whitespace, an underscore, then a line terminator: one logical line continues.</summary>
    LineContinuation,
}

/// <summary>
/// A run of insignificant text. Stored as offsets rather than a string: trivia is the majority of a
/// source file by count, and materialising each piece would allocate more than the tokens do.
/// </summary>
public readonly record struct Trivia(TriviaKind Kind, int Start, int End)
{
    public int Length => End - Start;

    public ReadOnlySpan<char> Text(string source) => source.AsSpan(Start, Length);
}

/// <summary>
/// One lexical token.
///
/// A token carries offsets, not text. The source string is already in memory and every token's text
/// is a slice of it, so slicing on demand costs nothing and holding a copy would roughly double the
/// memory a tokenized module occupies. On a per-keystroke path over a module of tens of thousands of
/// lines, that difference is the difference between comfortable and not.
/// </summary>
public readonly record struct Token(
    TokenKind Kind,
    int Start,
    int End,
    int Line,
    int Character,
    int TriviaStart,
    int TriviaCount)
{
    public int Length => End - Start;

    /// <summary>The token's exact source text.</summary>
    public ReadOnlySpan<char> Text(string source) => source.AsSpan(Start, Length);

    /// <summary>Leading trivia, as a slice of the tokenized module's trivia.</summary>
    public ReadOnlySpan<Trivia> LeadingTrivia(ReadOnlySpan<Trivia> all) => all.Slice(TriviaStart, TriviaCount);
}

/// <summary>
/// The complete token stream for one module.
///
/// The stream is lossless: every character of the source belongs to exactly one token or to exactly
/// one piece of trivia, so concatenating them reproduces the source byte for byte. That property is
/// what makes it safe to build editing features on top, and it is asserted by test rather than
/// assumed.
/// </summary>
public sealed class TokenizedModule
{
    internal TokenizedModule(string source, Token[] tokens, Trivia[] trivia)
    {
        Source = source;
        Tokens = tokens;
        Trivia = trivia;
    }

    public string Source { get; }

    public Token[] Tokens { get; }

    public Trivia[] Trivia { get; }

    /// <summary>Text of a token.</summary>
    public ReadOnlySpan<char> TextOf(in Token token) => token.Text(Source);

    /// <summary>
    /// Canonical capitalization for a keyword, or null for every other kind. Identifiers are case
    /// insensitive in this language, so what the user typed and what the editor shows can differ.
    /// </summary>
    public string? CanonicalTextOf(in Token token) =>
        token.Kind == TokenKind.Keyword ? KeywordTable.Canonical(token.Text(Source)) : null;
}
