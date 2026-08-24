namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// Which lines of a module can carry a breakpoint.
///
/// Only executable statements can. Asking the editor for one anywhere else puts a modal on screen
/// saying so - "Breakpoint not allowed on this line" - which is the host's answer to a question
/// the developer did not ask: they clicked a margin, and a dialog is not a reasonable reply to
/// that. So the line is judged here and the common refusals never reach the editor.
///
/// THE RULE USED TO BE LINE-LOCAL, AND THAT IS NOT ENOUGH. It excluded lines that BEGIN a
/// declaration - `Enum`, `Type`, `Dim`, `Option` and the rest - and had no way to know that a
/// line sits INSIDE one. So every member of every Enum and every field of every Type passed it:
///
///     Public Enum Corner        refused, correctly
///         TopLeft               offered, and the editor refuses it with a modal
///     End Enum                  refused, correctly
///
///     Public Type Point         refused, correctly
///         X As Double           offered, and the editor refuses it with a modal
///
/// Measured on LanguageFixture 2026-08-24: lines 5, 14 and 15 all offered, all refused by the
/// editor, and all RECORDED by this product as set - because the toggle records what it asked
/// for rather than what happened, and VBIDE exposes no way to ask. The developer is left with a
/// margin dot where no breakpoint exists, which is the state the owner photographed.
///
/// So the whole module is read, not one line of it. Blocks do not nest in VBA - a Type cannot
/// contain an Enum - so a single flag is the whole of the bookkeeping.
/// </summary>
public static class Breakpoints
{
    /// <summary>
    /// Whether VBA will accept a breakpoint on <paramref name="line"/>, counting from 1.
    ///
    /// Out of range is false: a line that is not there can carry nothing.
    /// </summary>
    public static bool CanCarry(IReadOnlyList<string> lines, int line)
    {
        ArgumentNullException.ThrowIfNull(lines);

        if (line < 1 || line > lines.Count)
        {
            return false;
        }

        // Walked from the top, because whether a line is inside a declaration block is not a
        // property of the line. Cheap: a module is thousands of lines at most and this runs on a
        // margin click, not per keystroke.
        var insideBlock = false;
        for (var at = 0; at < line; at++)
        {
            var code = lines[at].Trim();

            if (insideBlock)
            {
                if (StartsWithWord(code, "End"))
                {
                    var rest = code[3..].TrimStart();
                    if (StartsWithWord(rest, "Type", "Enum"))
                    {
                        insideBlock = false;
                    }
                }

                // Every line up to and including the End belongs to the block.
                if (at == line - 1)
                {
                    return false;
                }

                continue;
            }

            if (OpensDeclarationBlock(code))
            {
                insideBlock = true;
                if (at == line - 1)
                {
                    return false;
                }

                continue;
            }

            if (at == line - 1)
            {
                return CanCarryAlone(code);
            }
        }

        return false;
    }

    /// <summary>True for a line that begins a `Type` or `Enum` block, with or without modifiers.</summary>
    private static bool OpensDeclarationBlock(string code)
    {
        var rest = code;
        foreach (var modifier in (string[])["Public", "Private", "Friend", "Global"])
        {
            if (StartsWithWord(rest, modifier))
            {
                rest = rest[modifier.Length..].TrimStart();
                break;
            }
        }

        return StartsWithWord(rest, "Type", "Enum");
    }

    /// <summary>
    /// What the line says about itself, ignoring where it sits.
    ///
    /// Declarations are excluded, not modifiers. A procedure can start with the same words a
    /// module-level declaration does, and a breakpoint on the opening line of a procedure is
    /// perfectly legal, so it is what FOLLOWS the modifiers that decides.
    /// </summary>
    private static bool CanCarryAlone(string code)
    {
        if (code.Length == 0 || code.StartsWith('\'')
            || StartsWithWord(code, "Rem"))
        {
            return false;
        }

        if (StartsWithWord(code, "Option", "Attribute", "Declare", "Dim", "Const", "Type", "Enum"))
        {
            return false;
        }

        if (StartsWithWord(code, "End"))
        {
            var afterEnd = code[3..].TrimStart();
            if (StartsWithWord(afterEnd, "Type", "Enum"))
            {
                return false;
            }
        }

        foreach (var modifier in (string[])["Public", "Private", "Friend", "Static", "Global"])
        {
            if (StartsWithWord(code, modifier))
            {
                var rest = code[modifier.Length..].TrimStart();
                return StartsWithWord(rest, "Sub", "Function", "Property");
            }
        }

        return true;
    }

    /// <summary>True when the text opens with one of these words, as a WORD and not a prefix.</summary>
    private static bool StartsWithWord(string text, params string[] words)
    {
        foreach (var word in words)
        {
            if (!text.StartsWith(word, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (text.Length == word.Length || !char.IsLetterOrDigit(text[word.Length]) && text[word.Length] != '_')
            {
                return true;
            }
        }

        return false;
    }
}
