using System.Text.Json;
using Xlide.Vbe.Core.Engine;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The lightbulb's candidates as the shim reads them off a page message before relaying them to
/// the engine: every kind with the fields it needs, and all or nothing, because the verdicts come
/// back by position and one dropped candidate would pair every later verdict with the wrong entry.
/// </summary>
public sealed class EngineRefactorCandidateTests
{
    private static bool Read(string json, out EngineRefactorCandidate[] candidates)
    {
        using var document = JsonDocument.Parse(json);
        return EngineRefactorCandidate.TryReadAll(document.RootElement, out candidates);
    }

    [Fact]
    public void EveryKindIsReadWithTheFieldsItNeeds()
    {
        var ok = Read("""
            [
              {"kind":"inlineVariable","offset":12},
              {"kind":"introduceParameter","offset":12},
              {"kind":"moveToModule","offset":3},
              {"kind":"extractVariable","startOffset":40,"endOffset":49},
              {"kind":"extractMethod","startLine":9,"endLine":14},
              {"kind":"encapsulateField","fieldName":"Label"},
              {"kind":"implementInterface","interfaceName":"IStore"}
            ]
            """, out var candidates);

        Assert.True(ok);
        Assert.Equal(
            [
                new EngineRefactorCandidate("inlineVariable", Offset: 12),
                new EngineRefactorCandidate("introduceParameter", Offset: 12),
                new EngineRefactorCandidate("moveToModule", Offset: 3),
                new EngineRefactorCandidate("extractVariable", StartOffset: 40, EndOffset: 49),
                new EngineRefactorCandidate("extractMethod", StartLine: 9, EndLine: 14),
                new EngineRefactorCandidate("encapsulateField", FieldName: "Label"),
                new EngineRefactorCandidate("implementInterface", InterfaceName: "IStore"),
            ],
            candidates);
    }

    [Theory]
    [InlineData("""{"kind":"inlineVariable","offset":1}""")]
    [InlineData("""[{"kind":"renameEverything","offset":1}]""")]
    [InlineData("""[{"kind":"inlineVariable"}]""")]
    [InlineData("""[{"kind":"inlineVariable","offset":-1}]""")]
    [InlineData("""[{"kind":"inlineVariable","offset":"12"}]""")]
    [InlineData("""[{"kind":"inlineVariable","offset":1.5}]""")]
    [InlineData("""[{"kind":"extractVariable","startOffset":9,"endOffset":4}]""")]
    [InlineData("""[{"kind":"extractMethod","startLine":0,"endLine":2}]""")]
    [InlineData("""[{"kind":"extractMethod","startLine":5,"endLine":4}]""")]
    [InlineData("""[{"kind":"encapsulateField","fieldName":"   "}]""")]
    [InlineData("""[{"kind":"implementInterface"}]""")]
    [InlineData("""[{"offset":1}]""")]
    [InlineData("""[7]""")]
    public void AnythingMalformedRefusesTheWholeRequest(string json)
    {
        Assert.False(Read(json, out var candidates));
        Assert.Empty(candidates);
    }

    [Fact]
    public void OneBadCandidateAmongGoodOnesRefusesThemAll()
    {
        Assert.False(Read("""
            [{"kind":"inlineVariable","offset":12},{"kind":"inlineVariable","offset":-3}]
            """, out var candidates));
        Assert.Empty(candidates);
    }

    [Fact]
    public void NoMoreThanOneCaretRaisesIsRelayed()
    {
        static string Many(int count) =>
            "[" + string.Join(",", Enumerable.Repeat("""{"kind":"moveToModule","offset":1}""", count)) + "]";

        Assert.True(Read(Many(EngineRefactorCandidate.Most), out var most));
        Assert.Equal(EngineRefactorCandidate.Most, most.Length);
        Assert.False(Read(Many(EngineRefactorCandidate.Most + 1), out _));
    }

    [Fact]
    public void AnEmptyListIsAQuestionWithNothingInIt()
    {
        Assert.True(Read("[]", out var candidates));
        Assert.Empty(candidates);
    }

    [Fact]
    public void OnlyTheFieldsAKindUsesTravelToTheEngine()
    {
        var wire = JsonSerializer.Serialize(
            new EngineRefactorCandidate("extractVariable", StartOffset: 40, EndOffset: 49),
            EngineJsonContext.Default.EngineRefactorCandidate);

        Assert.Equal("""{"kind":"extractVariable","startOffset":40,"endOffset":49}""", wire);
    }

    [Fact]
    public void AVerdictReadsNullRefusedAsGoingThrough()
    {
        var answer = JsonSerializer.Deserialize(
            """{"verdicts":[{"kind":"inlineVariable","refused":null},{"kind":"moveToModule","refused":"no"}]}""",
            EngineJsonContext.Default.EngineRefactorings);

        Assert.NotNull(answer?.Verdicts);
        Assert.Null(answer!.Verdicts![0].Refused);
        Assert.Equal("no", answer.Verdicts[1].Refused);
    }
}
