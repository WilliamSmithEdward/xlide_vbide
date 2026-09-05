using System.Linq;
using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>Where annotations and attributes disagree, in both directions, and where neither can be asked.</summary>
public class AttributeDriftTests
{
    private static string Lines(params string[] lines) => string.Join("\r\n", lines) + "\r\n";

    private static readonly string Code = Lines(
        "'@PredeclaredId",
        "Option Explicit",
        "'@Description(\"One.\")",
        "Public Function Item() As Long",
        "End Function",
        "Public Function NewEnum() As IUnknown",
        "End Function");

    private static AttributeSet Actual(bool predeclared, string? itemDescription, int? newEnumId, string? moduleDescription = null) =>
        ModuleAttributes.Read(Lines(
            "Attribute VB_Name = \"Bag\"",
            $"Attribute VB_PredeclaredId = {predeclared}",
            "Attribute VB_Exposed = False",
            moduleDescription is null ? string.Empty : $"Attribute VB_Description = \"{moduleDescription}\"",
            "Public Function Item() As Long",
            itemDescription is null ? string.Empty : $"Attribute Item.VB_Description = \"{itemDescription}\"",
            "End Function",
            "Public Function NewEnum() As IUnknown",
            newEnumId is null ? string.Empty : $"Attribute NewEnum.VB_UserMemId = {newEnumId}",
            "End Function"));

    [Fact]
    public void InSyncIsSilent()
    {
        var drift = AttributeDrift.Between(Code, AttributeAnnotations.Read(Code), Actual(true, "One.", null), "class", "Bag");

        Assert.Empty(drift);
    }

    [Fact]
    public void AnAnnotationWithoutItsAttributeIsNotApplied()
    {
        var drift = AttributeDrift.Between(Code, AttributeAnnotations.Read(Code), Actual(false, "Other.", null), "class", "Bag");

        Assert.Collection(drift,
            one => { Assert.Equal(DriftKind.AnnotationNotApplied, one.Kind); Assert.Equal(1, one.Line); Assert.Contains("VB_PredeclaredId is False", one.Message); Assert.Equal("annotation-not-applied", one.Code); Assert.Equal("warning", one.Severity); },
            one => { Assert.Equal(DriftKind.AnnotationNotApplied, one.Kind); Assert.Equal(3, one.Line); Assert.Contains("Bag.Item's VB_Description is \"Other.\"", one.Message); Assert.Equal("Item", one.Target); });
    }

    [Fact]
    public void AnAttributeWithoutItsAnnotationIsReportedOnTheMemberItBelongsTo()
    {
        var drift = AttributeDrift.Between(Code, AttributeAnnotations.Read(Code), Actual(true, "One.", -4, "A bag."), "class", "Bag");

        Assert.Collection(drift,
            one => { Assert.Equal(DriftKind.AttributeNotAnnotated, one.Kind); Assert.Equal(1, one.Line); Assert.Contains("VB_Description is \"A bag.\"", one.Message); Assert.Equal("info", one.Severity); },
            one => { Assert.Equal(DriftKind.AttributeNotAnnotated, one.Kind); Assert.Equal(6, one.Line); Assert.Contains("enumerator", one.Message); Assert.Equal(AnnotationKind.Enumerator, one.Annotation); Assert.Equal("NewEnum", one.Target); });
    }

    [Fact]
    public void AnUnknownActualReportsEveryAnnotationAsUnconfirmedAndAsksNothingBackwards()
    {
        var drift = AttributeDrift.Between(Code, AttributeAnnotations.Read(Code), null, "class", "Bag");

        Assert.Equal(2, drift.Count);
        Assert.All(drift, one => Assert.Equal(DriftKind.AnnotationNotApplied, one.Kind));
        Assert.All(drift, one => Assert.Contains("does not carry Bag's attributes yet", one.Message));
    }

    [Theory]
    [InlineData("document")]
    [InlineData("userform")]
    public void AModuleThatCannotBeImportedCannotTakeAnnotations(string kind)
    {
        var drift = AttributeDrift.Between(Code, AttributeAnnotations.Read(Code), null, kind, "Sheet1");

        Assert.Equal(2, drift.Count);
        Assert.All(drift, one => Assert.Equal(DriftKind.AnnotationNotApplicable, one.Kind));
        Assert.Contains("cannot be imported", drift[0].Message);
    }

    [Fact]
    public void PredeclaredIdOnAStandardModuleIsNotApplicableWhileADescriptionIs()
    {
        var drift = AttributeDrift.Between(Code, AttributeAnnotations.Read(Code), null, "standard", "Macros");

        Assert.Equal(DriftKind.AnnotationNotApplicable, drift[0].Kind);
        Assert.Contains("only on a class module", drift[0].Message);
        Assert.Equal(DriftKind.AnnotationNotApplied, drift[1].Kind);
    }

    [Fact]
    public void ParserProblemsRideAlong()
    {
        var code = Lines("Public Sub Go()", "End Sub", "'@PredeclaredId");
        var drift = AttributeDrift.Between(code, AttributeAnnotations.Read(code), Actual(false, null, null), "class", "Bag");

        var one = Assert.Single(drift);
        Assert.Equal(DriftKind.AnnotationProblem, one.Kind);
        Assert.Equal("annotation-problem", one.Code);
        Assert.Equal(3, one.Line);
    }
}
