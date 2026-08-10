using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The editor breaks a line it cannot hold instead of refusing it, and the break carries no
/// continuation, so what lands is not code. These pin the boundary that decides whether a write
/// happens at all, and the boundary is exactly where an off-by-one hides: 1,022 is fine, 1,023 is
/// the first width the editor splits (measured 2026-08-09).
/// </summary>
public class ModuleTextTests
{
    [Fact]
    public void NothingIsTooLongInOrdinaryCode()
    {
        Assert.Null(ModuleText.FirstLineTooLong("Option Explicit\r\n\r\nPublic Sub A()\r\nEnd Sub\r\n"));
    }

    [Fact]
    public void EmptyTextHasNoLineAtAll()
    {
        Assert.Null(ModuleText.FirstLineTooLong(string.Empty));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(1021)]
    [InlineData(1022)]
    public void AWidthTheEditorHoldsIsNotReported(int width)
    {
        Assert.Null(ModuleText.FirstLineTooLong(new string('a', width)));
    }

    [Theory]
    [InlineData(1023)]
    [InlineData(2048)]
    public void TheFirstWidthTheEditorSplitsIsReported(int width)
    {
        var found = ModuleText.FirstLineTooLong(new string('a', width));

        Assert.NotNull(found);
        Assert.Equal(1, found.Value.At);
        Assert.Equal(width, found.Value.Length);
    }

    [Fact]
    public void TheLineNumberIsTheOneTheDeveloperSees()
    {
        var text = $"Option Explicit\r\n\r\nPublic Sub A()\r\n    Debug.Print \"{new string('a', 1200)}\"\r\nEnd Sub";

        var found = ModuleText.FirstLineTooLong(text);

        Assert.NotNull(found);
        Assert.Equal(4, found.Value.At);
    }

    [Fact]
    public void ACarriageReturnIsNotPartOfTheLineItEnds()
    {
        // 1,022 characters and a CRLF. Counting the carriage return would make this 1,023 and
        // refuse a line the editor holds perfectly well.
        Assert.Null(ModuleText.FirstLineTooLong($"{new string('a', 1022)}\r\nEnd Sub"));
    }

    [Fact]
    public void ALineFeedOnItsOwnEndsALineToo()
    {
        // The surface sends LF; the object model gives back CRLF. Both are line breaks here.
        var found = ModuleText.FirstLineTooLong($"Option Explicit\n{new string('a', 1023)}\nEnd Sub");

        Assert.NotNull(found);
        Assert.Equal(2, found.Value.At);
        Assert.Equal(1023, found.Value.Length);
    }

    [Fact]
    public void TheFirstOneIsTheOneReported()
    {
        var text = $"{new string('a', 1100)}\r\n{new string('b', 1500)}";

        var found = ModuleText.FirstLineTooLong(text);

        Assert.NotNull(found);
        Assert.Equal(1, found.Value.At);
        Assert.Equal(1100, found.Value.Length);
    }
}

/// <summary>
/// What the host converts on the way in, and how it is noticed.
///
/// VBA stores module text in the system ANSI code page, so a character outside it is converted by
/// Excel before this product sees it. These pin the noticing, which decides whether an import
/// writes a module the developer's file cannot be rebuilt from. The live behaviour is machine
/// dependent - it is a property of the code page - so what is fixed here is the RULE, against
/// conversions measured on code page 1252 (2026-08-09).
/// </summary>
public class ModuleTextConversionTests
{
    private const string Body = "Option Explicit\n\nPublic Sub A()\nEnd Sub\n";

    [Fact]
    public void AsciiCannotBeLost()
    {
        Assert.Null(ModuleText.FirstCharacterLost(Body, Body));
    }

    [Fact]
    public void NothingIsReportedWhenTheHostKeptItAll()
    {
        const string text = "' déjà vu € œuvre Straße";

        Assert.Null(ModuleText.FirstCharacterLost(text, text));
    }

    [Fact]
    public void TheEditorsOwnRewritingIsNotALoss()
    {
        // It completes parentheses and respells keywords, all of it ASCII. Comparing whole texts
        // would refuse a write over the editor's own tidying, which is why this asks about
        // presence rather than equality.
        Assert.Null(ModuleText.FirstCharacterLost(
            "public sub Grüße()\nEnd Sub",
            "Public Sub Grüße()\r\nEnd Sub\r\n"));
    }

    [Theory]
    [InlineData("Проверка")]
    [InlineData("Δοκιμή")]
    [InlineData("ทดสอบ")]
    [InlineData("中文测试")]
    public void AScriptTheCodePageCannotSpellIsReported(string text)
    {
        // Code page 1252 answers '?' to every one of these.
        var lost = ModuleText.FirstCharacterLost($"' {text}", $"' {new string('?', text.Length)}");

        Assert.NotNull(lost);
        Assert.Equal(text[0], lost.Value.Text[0]);
    }

    [Fact]
    public void ADecomposedAccentIsReportedThoughItsPrecomposedSpellingSurvives()
    {
        // THE CASE THAT HIDES. Code page 1252 turns "cafe" + U+0301 into "cafe" + U+00B4: not a
        // question mark, a different real character that renders almost the same. The precomposed
        // spelling of the identical word is kept untouched, so the same text survives or is
        // altered by its normalisation alone, and a macOS-authored repository is decomposed.
        Assert.Null(ModuleText.FirstCharacterLost("café", "café"));

        var lost = ModuleText.FirstCharacterLost("café", "cafe´");

        Assert.NotNull(lost);
        Assert.Equal(0x0301, lost.Value.CodePoint);
    }

    [Fact]
    public void AnAstralCharacterIsReportedWholeRatherThanAsHalfASurrogate()
    {
        var lost = ModuleText.FirstCharacterLost("' \U0001F600", "' ??");

        Assert.NotNull(lost);
        Assert.Equal(0x1F600, lost.Value.CodePoint);
        Assert.Equal("\U0001F600", lost.Value.Text);
    }

    [Fact]
    public void TheLineIsTheOneTheDeveloperWouldLookAt()
    {
        var lost = ModuleText.FirstCharacterLost(
            "Option Explicit\n\n' ok\n' Проверка\n",
            "Option Explicit\n\n' ok\n' ????????\n");

        Assert.NotNull(lost);
        Assert.Equal(4, lost.Value.Line);
    }

    [Fact]
    public void TheFirstLossIsTheOneReported()
    {
        var lost = ModuleText.FirstCharacterLost("' Δ then ́", "' ? then ´");

        Assert.NotNull(lost);
        Assert.Equal(0x0394, lost.Value.CodePoint);
    }

    [Fact]
    public void ACharacterKeptSomewhereElseIsNotALoss()
    {
        // Conservative on purpose: code page conversion is per character and deterministic, so a
        // character present anywhere in the answer was not converted. Being wrong in this
        // direction costs nothing; being wrong the other way refuses a write that was fine.
        Assert.Null(ModuleText.FirstCharacterLost("é at the end", "at the end é"));
    }

    [Fact]
    public void ItSaysTheCharacterWhenThatHelpsAndTheNumberAlways()
    {
        Assert.Equal("'Δ' (U+0394)", new LostCharacter("Δ", 0x0394, 1).Describe());

        // A combining mark on its own renders on whatever precedes it, including the quote that
        // was meant to contain it, so it is named by number alone.
        Assert.Equal("U+0301", new LostCharacter("́", 0x0301, 1).Describe());
    }
}
