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
