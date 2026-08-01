namespace Xlide.Vba.Analysis.Lexing;

/// <summary>
/// Turns VBA source into a lossless token stream, per MS-VBAL section 3.3.
///
/// Lossless is the load-bearing property: every character ends up in exactly one token or one piece
/// of trivia, so the stream can be concatenated back into the original source unchanged. Editing
/// features are built on that guarantee, and it is cheaper to hold it from the start than to
/// retrofit it.
///
/// The lexer never throws and never rejects input. Text that matches no rule becomes an unknown
/// token and lexing continues, because a module being typed into is malformed most of the time and
/// an analyzer that gives up on the first bad character is useless in an editor.
/// </summary>
public static class Lexer
{
    public static TokenizedModule Tokenize(string source)
    {
        ArgumentNullException.ThrowIfNull(source);

        // A token roughly every six characters, and rather less trivia, on typical source. Sizing
        // up front avoids repeated growth on the per-keystroke path.
        var tokens = new List<Token>(Math.Max(16, source.Length / 6));
        var trivia = new List<Trivia>(Math.Max(16, source.Length / 12));

        var position = 0;
        var line = 0;
        var lineStart = 0;

        while (true)
        {
            var triviaStart = trivia.Count;
            ScanTrivia(source, trivia, ref position, ref line, ref lineStart);

            if (position >= source.Length)
            {
                // The end token spans nothing and carries whatever trivia trailed the last real
                // token, so the stream stays lossless without a token claiming source it does not
                // cover.
                tokens.Add(new Token(
                    TokenKind.EndOfFile, source.Length, source.Length, line, source.Length - lineStart,
                    triviaStart, trivia.Count - triviaStart));
                break;
            }

            var start = position;
            var character = start - lineStart;
            var kind = ScanToken(source, ref position);

            tokens.Add(new Token(kind, start, position, line, character, triviaStart, trivia.Count - triviaStart));

            if (kind == TokenKind.Newline)
            {
                line++;
                lineStart = position;
            }
        }

        return new TokenizedModule(source, tokens.ToArray(), trivia.ToArray());
    }

    /// <summary>
    /// Consumes whitespace and line continuations. A continuation is whitespace, an underscore, then
    /// a line terminator: the physical line ends but the logical line does not, so the terminator is
    /// trivia rather than a newline token.
    /// </summary>
    private static void ScanTrivia(string source, List<Trivia> trivia, ref int position, ref int line, ref int lineStart)
    {
        while (position < source.Length)
        {
            var start = position;

            while (position < source.Length && IsWhitespace(source[position]))
            {
                position++;
            }

            if (position < source.Length && source[position] == '_' && IsLineContinuation(source, position))
            {
                position++;

                while (position < source.Length && IsWhitespace(source[position]))
                {
                    position++;
                }

                position += ConsumeLineTerminator(source, position);

                trivia.Add(new Trivia(TriviaKind.LineContinuation, start, position));
                line++;
                lineStart = position;
                continue;
            }

            if (position > start)
            {
                trivia.Add(new Trivia(TriviaKind.Whitespace, start, position));
            }

            return;
        }
    }

    private static TokenKind ScanToken(string source, ref int position)
    {
        var c = source[position];

        if (IsLineTerminator(c))
        {
            position += ConsumeLineTerminator(source, position);
            return TokenKind.Newline;
        }

        if (c == '\'')
        {
            ScanToEndOfLine(source, ref position);
            return TokenKind.Comment;
        }

        if (c == '"')
        {
            ScanString(source, ref position);
            return TokenKind.StringLiteral;
        }

        if (c == '[')
        {
            return ScanBracketedIdentifier(source, ref position);
        }

        if (c == '#')
        {
            return ScanHash(source, ref position);
        }

        if (char.IsAsciiDigit(c))
        {
            return ScanNumber(source, ref position);
        }

        // A dot begins a float only when a digit follows. Otherwise it is member access.
        if (c == '.' && position + 1 < source.Length && char.IsAsciiDigit(source[position + 1]))
        {
            return ScanNumber(source, ref position);
        }

        if (c == '&' && position + 1 < source.Length && IsRadixMarker(source[position + 1]))
        {
            return ScanRadixInteger(source, ref position);
        }

        if (IsIdentifierStart(c))
        {
            return ScanWord(source, ref position);
        }

        return ScanSymbol(source, ref position);
    }

    private static TokenKind ScanWord(string source, ref int position)
    {
        var start = position;

        while (position < source.Length && IsIdentifierPart(source[position]))
        {
            position++;
        }

        var word = source.AsSpan(start, position - start);

        // Rem introduces a comment that runs to the end of the line, so the rest of the line is not
        // tokenized as code at all.
        if (word.Equals("Rem", StringComparison.OrdinalIgnoreCase))
        {
            ScanToEndOfLine(source, ref position);
            return TokenKind.Comment;
        }

        if (KeywordTable.IsKeyword(word))
        {
            return TokenKind.Keyword;
        }

        // A type suffix binds to the name it follows and is part of the token.
        if (position < source.Length && IsTypeSuffix(source[position]))
        {
            position++;
        }

        return TokenKind.Identifier;
    }

    private static TokenKind ScanNumber(string source, ref int position)
    {
        var isFloat = false;

        while (position < source.Length && char.IsAsciiDigit(source[position]))
        {
            position++;
        }

        if (position < source.Length && source[position] == '.')
        {
            // A second dot belongs to member access, not to the number.
            isFloat = true;
            position++;

            while (position < source.Length && char.IsAsciiDigit(source[position]))
            {
                position++;
            }
        }

        if (position < source.Length && (source[position] is 'e' or 'E' or 'd' or 'D'))
        {
            var exponent = position + 1;

            if (exponent < source.Length && (source[exponent] is '+' or '-'))
            {
                exponent++;
            }

            // Only an exponent with digits is part of the number. Otherwise the letter starts a name.
            if (exponent < source.Length && char.IsAsciiDigit(source[exponent]))
            {
                isFloat = true;
                position = exponent;

                while (position < source.Length && char.IsAsciiDigit(source[position]))
                {
                    position++;
                }
            }
        }

        if (position < source.Length && IsTypeSuffix(source[position]))
        {
            // The suffix decides the type regardless of how the digits looked.
            isFloat = source[position] is '!' or '#' or '@';
            position++;
        }

        return isFloat ? TokenKind.FloatLiteral : TokenKind.IntegerLiteral;
    }

    private static TokenKind ScanRadixInteger(string source, ref int position)
    {
        // Skip the ampersand and the radix marker.
        var isHex = source[position + 1] is 'h' or 'H';
        position += 2;

        while (position < source.Length && IsRadixDigit(source[position], isHex))
        {
            position++;
        }

        // A trailing ampersand is a Long suffix and belongs to the literal.
        if (position < source.Length && (source[position] is '&' or '%' or '^'))
        {
            position++;
        }

        return TokenKind.IntegerLiteral;
    }

    private static void ScanString(string source, ref int position)
    {
        position++;

        while (position < source.Length)
        {
            var c = source[position];

            if (IsLineTerminator(c))
            {
                // An unterminated string ends at the line end. Consuming the terminator would
                // swallow the newline and make every following line part of the string.
                return;
            }

            position++;

            if (c != '"')
            {
                continue;
            }

            // A doubled quote is an escaped quote and the string continues.
            if (position < source.Length && source[position] == '"')
            {
                position++;
                continue;
            }

            return;
        }
    }

    private static TokenKind ScanBracketedIdentifier(string source, ref int position)
    {
        var start = position;
        position++;

        while (position < source.Length && source[position] != ']')
        {
            if (IsLineTerminator(source[position]))
            {
                // Unterminated. Report the bracket alone rather than consuming the rest of the file.
                position = start + 1;
                return TokenKind.Unknown;
            }

            position++;
        }

        if (position >= source.Length)
        {
            position = start + 1;
            return TokenKind.Unknown;
        }

        position++;
        return TokenKind.BracketedIdentifier;
    }

    /// <summary>
    /// A hash begins either a date literal or a conditional compilation directive. They are told
    /// apart by what follows: a directive is a hash immediately followed by a keyword.
    /// </summary>
    private static TokenKind ScanHash(string source, ref int position)
    {
        var start = position;
        var next = position + 1;

        if (next < source.Length && IsIdentifierStart(source[next]))
        {
            var wordEnd = next;
            while (wordEnd < source.Length && IsIdentifierPart(source[wordEnd]))
            {
                wordEnd++;
            }

            var word = source.AsSpan(next, wordEnd - next);
            if (IsDirectiveWord(word))
            {
                position = next;
                return TokenKind.Directive;
            }
        }

        // A date literal runs to its closing hash on the same line.
        position++;

        while (position < source.Length && source[position] != '#')
        {
            if (IsLineTerminator(source[position]))
            {
                position = start + 1;
                return TokenKind.Unknown;
            }

            position++;
        }

        if (position >= source.Length)
        {
            position = start + 1;
            return TokenKind.Unknown;
        }

        position++;
        return TokenKind.DateLiteral;
    }

    private static bool IsDirectiveWord(ReadOnlySpan<char> word) =>
        word.Equals("If", StringComparison.OrdinalIgnoreCase)
        || word.Equals("Else", StringComparison.OrdinalIgnoreCase)
        || word.Equals("ElseIf", StringComparison.OrdinalIgnoreCase)
        || word.Equals("End", StringComparison.OrdinalIgnoreCase)
        || word.Equals("Const", StringComparison.OrdinalIgnoreCase);

    private static TokenKind ScanSymbol(string source, ref int position)
    {
        var c = source[position];
        position++;

        switch (c)
        {
            case ':':
                // A colon followed by equals is the named-argument marker, one operator token.
                if (position < source.Length && source[position] == '=')
                {
                    position++;
                    return TokenKind.Operator;
                }

                return TokenKind.Colon;

            case '<':
                if (position < source.Length && (source[position] is '=' or '>'))
                {
                    position++;
                }

                return TokenKind.Operator;

            case '>':
                if (position < source.Length && source[position] == '=')
                {
                    position++;
                }

                return TokenKind.Operator;

            case '=':
            case '+':
            case '-':
            case '*':
            case '/':
            case '\\':
            case '^':
            case '&':
            case '!':
                return TokenKind.Operator;

            case ',':
            case '.':
            case '(':
            case ')':
            case ';':
                return TokenKind.Punctuation;

            default:
                return TokenKind.Unknown;
        }
    }

    private static void ScanToEndOfLine(string source, ref int position)
    {
        while (position < source.Length && !IsLineTerminator(source[position]))
        {
            position++;
        }
    }

    /// <summary>Length of the terminator at this position, treating a carriage return and line feed pair as one.</summary>
    private static int ConsumeLineTerminator(string source, int position)
    {
        if (position >= source.Length)
        {
            return 0;
        }

        if (source[position] == '\r')
        {
            return position + 1 < source.Length && source[position + 1] == '\n' ? 2 : 1;
        }

        return source[position] == '\n' ? 1 : 0;
    }

    /// <summary>True when an underscore at this position continues the logical line.</summary>
    private static bool IsLineContinuation(string source, int position)
    {
        var probe = position + 1;

        while (probe < source.Length && IsWhitespace(source[probe]))
        {
            probe++;
        }

        return probe < source.Length && IsLineTerminator(source[probe]);
    }

    /// <summary>
    /// Whitespace per MS-VBAL 3.2.2: tab, space, the end-of-message character, the double-byte
    /// space, and other Unicode space separators. Line terminators are excluded, because they end
    /// a logical line and are tokens rather than trivia.
    /// </summary>
    private static bool IsWhitespace(char c)
    {
        // Ordered by frequency. Almost every whitespace character in real source is a space or a
        // tab, and this runs once per character of the file.
        if (c is ' ' or '\t')
        {
            return true;
        }

        if (c < 128)
        {
            return false;
        }

        return c is '\u0019' or '\u3000'
            || (char.IsWhiteSpace(c) && !IsLineTerminator(c) && c is not ('\v' or '\f'));
    }

    private static bool IsLineTerminator(char c) => c is '\r' or '\n';

    private static bool IsIdentifierStart(char c) => char.IsLetter(c) || c == '_';

    private static bool IsIdentifierPart(char c) => char.IsLetterOrDigit(c) || c == '_';

    private static bool IsTypeSuffix(char c) => c is '%' or '&' or '^' or '!' or '#' or '@' or '$';

    private static bool IsRadixMarker(char c) => c is 'h' or 'H' or 'o' or 'O';

    private static bool IsRadixDigit(char c, bool isHex) =>
        isHex ? char.IsAsciiHexDigit(c) : c is >= '0' and <= '7';
}
