using Xlide.Vbe.Core.Sync;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The allocation-free code comparison and the one-allocation header strip, against the
/// split-and-join they replaced, kept here as the oracle. Every input the two could read
/// differently is put to both: the three line endings and their mixtures, headers, attribute
/// lines inside a procedure, blank and whitespace lines above and below the code, and texts
/// with nothing in them.
/// </summary>
public sealed class CodeComparisonTests
{
    /// <summary>What CodeWithoutHeader was until 2026-09-09, line for line.</summary>
    private static string Reference(string source)
    {
        var lines = source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        var kept = new List<string>();
        var stillInHeader = true;
        foreach (var line in lines)
        {
            if (stillInHeader && (ModuleSync.IsAttributeLine(line) || ModuleSync.IsHeaderPreamble(line)))
            {
                continue;
            }

            if (ModuleSync.IsAttributeLine(line))
            {
                continue;
            }

            if (stillInHeader && line.Trim().Length == 0)
            {
                continue;
            }

            stillInHeader = false;
            kept.Add(line);
        }

        return string.Join("\n", kept);
    }

    private static IEnumerable<string> Samples()
    {
        yield return "";
        yield return "\n";
        yield return "\r\n";
        yield return "\r";
        yield return "Attribute VB_Name = \"M\"\r\n";
        yield return "Attribute VB_Name = \"M\"\r\nOption Explicit\r\n\r\nSub A()\r\nEnd Sub\r\n";
        yield return "Attribute VB_Name = \"M\"\nOption Explicit\n\nSub A()\nEnd Sub\n";
        yield return "Attribute VB_Name = \"M\"\rOption Explicit\r\rSub A()\rEnd Sub\r";
        yield return "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\n"
            + "Attribute VB_Exposed = False\r\n\r\n\r\nOption Explicit\r\n";
        yield return "Option Explicit\r\n   \r\n\t\r\nSub A()\r\n    Attribute A.VB_Description = \"x\"\r\nEnd Sub";
        yield return "\r\n\r\n  \r\nOption Explicit";
        yield return "Option Explicit\r\n\r\n\r\n";
        yield return "Option Explicit\r\n   \r\n";
        yield return "Option Explicit\r\nSub A()\r\n\r\n\r\nEnd Sub\r\n\r\n";
        yield return "Sub A()\r\n    x = 1\r\n\r\nEnd Sub\r\nAttribute VB_Description = \"tail\"";
        yield return "no newline at all";
        yield return "mixed\nendings\r\nhere\rtoo\n";
    }

    public static IEnumerable<object[]> Texts() => Samples().Select(one => new object[] { one });

    [Theory]
    [MemberData(nameof(Texts))]
    public void CodeWithoutHeaderAgreesWithTheSplitAndJoinItReplaced(string text)
    {
        Assert.Equal(Reference(text), ModuleSync.CodeWithoutHeader(text));
    }

    [Fact]
    public void SameCodeAgreesWithComparingTheNormalisedTextsAcrossEveryPair()
    {
        var samples = Samples().ToList();
        foreach (var left in samples)
        {
            foreach (var right in samples)
            {
                var expected = string.Equals(
                    Reference(left).TrimEnd('\n'), Reference(right).TrimEnd('\n'), StringComparison.Ordinal);
                Assert.True(
                    expected == ModuleSync.SameCode(left, right),
                    $"{Shown(left)} against {Shown(right)}: expected {expected}");
            }
        }
    }

    [Fact]
    public void TheThreeEndingsAndTrailingBlankLinesAreTheSameCodeAndABlankInsideIsNot()
    {
        Assert.True(ModuleSync.SameCode("Attribute VB_Name = \"M\"\r\nSub A()\r\nEnd Sub\r\n", "Sub A()\nEnd Sub"));
        Assert.True(ModuleSync.SameCode("Sub A()\rEnd Sub\r\r\r", "Sub A()\r\nEnd Sub"));
        Assert.False(ModuleSync.SameCode("Sub A()\r\n    x = 1\r\nEnd Sub", "Sub A()\r\n    x = 2\r\nEnd Sub"));
        Assert.False(ModuleSync.SameCode("Sub A()\r\n\r\n    x = 1\r\nEnd Sub", "Sub A()\r\n    x = 1\r\nEnd Sub"));
    }

    private static string Shown(string text) =>
        text.Replace("\r", "\\r", StringComparison.Ordinal).Replace("\n", "\\n", StringComparison.Ordinal);
}
