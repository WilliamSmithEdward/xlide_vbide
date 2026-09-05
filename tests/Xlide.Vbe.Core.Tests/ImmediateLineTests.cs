using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The one shape of Immediate line that break mode answers: a print of a single bare name (#21).
/// Everything else stays an evaluation, which break mode declines.
/// </summary>
public class ImmediateLineTests
{
    [Theory]
    [InlineData("? counter", "counter")]
    [InlineData("?counter", "counter")]
    [InlineData("  ?   counter  ", "counter")]
    [InlineData("Print counter", "counter")]
    [InlineData("print counter", "counter")]
    [InlineData("Debug.Print counter", "counter")]
    [InlineData("? größe", "größe")]
    [InlineData("? total_2", "total_2")]
    public void APrintOfOneNameIsRecognised(string line, string name)
    {
        Assert.Equal(name, ImmediateLine.NameToPrint(line));
    }

    [Theory]
    [InlineData("? counter + 1")]
    [InlineData("? obj.Name")]
    [InlineData("? Left$(label, 2)")]
    [InlineData("? \"text\"")]
    [InlineData("? 42")]
    [InlineData("?")]
    [InlineData("counter = 1")]
    [InlineData("? counter, label")]
    [InlineData("Printer")]
    [InlineData("? _leading")]
    [InlineData("")]
    [InlineData(null)]
    public void AnythingElseIsAnEvaluation(string? line)
    {
        Assert.Null(ImmediateLine.NameToPrint(line));
    }

    [Theory]
    [InlineData("\"start\"", "String", "start")]
    [InlineData("\"he said \"\"hi\"\"\"", "String", "he said \"hi\"")]
    [InlineData("\"\"", "String", "")]
    [InlineData("42", "Long", "42")]
    [InlineData("0.5", "Double", "0.5")]
    [InlineData("True", "Boolean", "True")]
    [InlineData("", "Worksheet", "<Worksheet>")]
    public void ALocalsValuePrintsTheWayTheImmediateWindowWould(string value, string type, string printed)
    {
        Assert.Equal(printed, ImmediateLine.AsPrinted(value, type));
    }
}
