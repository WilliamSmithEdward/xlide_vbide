using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

public class ImmediateTranscriptTests
{
    [Theory]
    [InlineData("", "new", "new")]
    [InlineData("old\r\n", "old\r\nnew\r\n", "new\r\n")]
    [InlineData("old\r\nkept\r\n", "kept\r\nnew\r\n", "new\r\n")]
    [InlineData("one\r\nsame\r\nsame\r\n", "same\r\nsame\r\nnew\r\n", "new\r\n")]
    [InlineData("same", "same", "")]
    [InlineData("old output", "rewritten output", null)]
    [InlineData("prefixvalue", "value\r\nnew", null)]
    public void FindsNewOutputAcrossWholeLineHistoryRollover(string previous, string current, string? added)
    {
        Assert.Equal(added, ImmediateTranscript.Added(previous, current));
    }

    [Theory]
    [InlineData(" 2 \r\n", " 2 ")]
    [InlineData("  padded  \r\n", "  padded  ")]
    [InlineData("\r\n", "")]
    [InlineData("", "")]
    [InlineData("first\r\n\r\nlast\r\n", "first\r\n\r\nlast")]
    [InlineData("tail without newline", "tail without newline")]
    public void PreservesOutputAndRemovesOnlyNativeFurniture(string output, string expected)
    {
        Assert.Equal(expected, ImmediateTranscript.Output("old output\r\n? expression 'tag\r\n"
            + output + "\r\n\0", "'tag"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("result without the command\r\n\0")]
    [InlineData("? expression 'tag\r\npartial")]
    public void MissingOrIncompleteOutputIsNotReportedAsSuccess(string? text)
    {
        Assert.Null(ImmediateTranscript.Output(text, "'tag"));
    }
}
