using System.Text.Json;
using Xlide.Vbe.Core.Engine;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// A finding arrives as an offset AND as a line and column, and the second one is the one to draw
/// with.
///
/// An offset means nothing without the text it was counted in. A diagnostics request may leave the
/// source out, and then the engine picks between its live copy of the module and the copy the
/// project was seeded with, a choice this side cannot see. Converting the offsets here against
/// whatever text this side happened to hold was right only while the two agreed; formatting a
/// module made them disagree, and a finding that belonged at 6:11 was drawn at 6:6, six columns
/// left of the word it was about, and stayed there.
///
/// So the engine now sends the position it measured in the text it actually analysed. Under
/// ahead-of-time compilation a field missing from the serializer context does not fail the build,
/// it silently reads as null at run time, which would put the old arithmetic straight back. These
/// tests are that guard: they run the real engine payload through the real context.
/// </summary>
public class EngineDiagnosticTests
{
    private static readonly JsonSerializerOptions Options = EngineJsonContext.Default.Options;

    [Fact]
    public void ThePositionSurvivesTheWire()
    {
        const string payload = """
        {
          "diagnostics": [
            {
              "code": "unreachable-argument",
              "message": "Close accepts no argument here.",
              "severity": "error",
              "span": { "start": 74, "end": 79 },
              "at": { "startLine": 6, "startColumn": 12, "endLine": 6, "endColumn": 17 }
            }
          ],
          "mode": "full"
        }
        """;

        var read = JsonSerializer.Deserialize<EngineDiagnostics>(payload, Options);

        Assert.NotNull(read);
        var finding = Assert.Single(read.Diagnostics);

        Assert.Equal(74, finding.Span.Start);
        Assert.Equal(79, finding.Span.End);

        Assert.NotNull(finding.At);
        Assert.Equal(6, finding.At.StartLine);
        Assert.Equal(12, finding.At.StartColumn);
        Assert.Equal(6, finding.At.EndLine);
        Assert.Equal(17, finding.At.EndColumn);
    }

    /// <summary>
    /// An engine that predates the field is still readable, and says so by leaving the position
    /// null rather than by reading as line zero. The caller falls back to its own arithmetic
    /// there, which is what it did everywhere before.
    /// </summary>
    [Fact]
    public void AnEngineThatSendsNoPositionReadsAsNoPosition()
    {
        const string payload = """
        {
          "diagnostics": [
            {
              "message": "Something older.",
              "severity": "warning",
              "span": { "start": 3, "end": 8 }
            }
          ]
        }
        """;

        var read = JsonSerializer.Deserialize<EngineDiagnostics>(payload, Options);

        Assert.NotNull(read);
        var finding = Assert.Single(read.Diagnostics);

        Assert.Equal(3, finding.Span.Start);
        Assert.Null(finding.At);
    }

    /// <summary>
    /// The arithmetic the fallback uses, against the two texts of the defect. Both are correct
    /// readings of the offset; they differ because they are readings of different texts, which is
    /// the whole reason the position now travels with the finding.
    /// </summary>
    [Fact]
    public void TheSameOffsetIsADifferentPositionInADifferentText()
    {
        var before = string.Join(
            "\r\n",
            "Option Explicit", "", "Public Sub Probe()", "    Dim n As Long", "    n = 1",
            "Workbooks.Close n", "    Debug.Print n", "End Sub", "");

        var after = string.Join(
            "\r\n",
            "Option Explicit", "", "Public Sub Probe()", "\tDim n As Long", "\tn = 1",
            "\tWorkbooks.Close n", "\tDebug.Print n", "End Sub", "");

        var offset = after.IndexOf("Close", StringComparison.Ordinal);

        Assert.Equal((6, 12), TextPositions.ToLineColumn(TextPositions.LineStarts(after), offset));
        Assert.Equal((6, 6), TextPositions.ToLineColumn(TextPositions.LineStarts(before), offset));
    }
}
