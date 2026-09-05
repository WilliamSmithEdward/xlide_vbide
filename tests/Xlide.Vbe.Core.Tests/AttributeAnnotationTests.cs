using System.Linq;
using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The annotations that stand for hidden attributes: what binds to what, and what is refused.
/// A misbound annotation would write an attribute onto the wrong member, so the binding rules
/// are the contract here.
/// </summary>
public class AttributeAnnotationTests
{
    private const string Crlf = "\r\n";

    private static string Lines(params string[] lines) => string.Join(Crlf, lines) + Crlf;

    [Fact]
    public void ModuleAnnotationsAreReadFromTheDeclarations()
    {
        var read = AttributeAnnotations.Read(Lines(
            "'@ModuleDescription(\"A bag of things.\")",
            "'@PredeclaredId",
            "Option Explicit",
            "'@Exposed",
            "Private items As Collection",
            "Public Sub Go()",
            "End Sub"));

        Assert.Empty(read.Problems);
        Assert.Collection(read.Annotations,
            one => { Assert.Equal(AnnotationKind.ModuleDescription, one.Kind); Assert.Equal("A bag of things.", one.Argument); Assert.Equal(1, one.Line); Assert.Null(one.Target); },
            one => { Assert.Equal(AnnotationKind.PredeclaredId, one.Kind); Assert.Null(one.Argument); Assert.Equal(2, one.Line); },
            one => { Assert.Equal(AnnotationKind.Exposed, one.Kind); Assert.Equal(4, one.Line); });
    }

    [Theory]
    [InlineData("'@Description(\"Adds one.\")", "Adds one.")]
    [InlineData("'@Description \"Adds one.\"", "Adds one.")]
    [InlineData("'@description(Adds one.)", "Adds one.")]
    [InlineData("'  @Description  (  \"Adds one.\"  )", "Adds one.")]
    [InlineData("'@Description(\"He said \"\"hi\"\".\")", "He said \"hi\".")]
    public void MemberAnnotationsBindToTheProcedureBelow(string line, string argument)
    {
        var read = AttributeAnnotations.Read(Lines(
            "Option Explicit",
            "",
            line,
            "' more words about it",
            "",
            "Public Sub Add()",
            "End Sub"));

        Assert.Empty(read.Problems);
        var one = Assert.Single(read.Annotations);
        Assert.Equal(AnnotationKind.Description, one.Kind);
        Assert.Equal(argument, one.Argument);
        Assert.Equal("Add", one.Target);
        Assert.Equal(6, one.TargetLine);
        Assert.Equal(3, one.Line);
    }

    [Fact]
    public void DefaultMemberEnumeratorAndHotkeyBindTooAndAPropertyLegIsItsOwnTarget()
    {
        var read = AttributeAnnotations.Read(Lines(
            "Option Explicit",
            "'@DefaultMember",
            "Public Property Get Item(ByVal i As Long) As Variant",
            "End Property",
            "'@Description(\"Sets one.\")",
            "Public Property Let Item(ByVal i As Long, ByVal v As Variant)",
            "End Property",
            "'@Enumerator",
            "Public Function NewEnum() As IUnknown",
            "End Function",
            "'@ExcelHotkey(\"D\")",
            "Public Sub Run()",
            "End Sub"));

        Assert.Empty(read.Problems);
        Assert.Equal(
            new[] { ("DefaultMember", "Item", 3), ("Description", "Item", 6), ("Enumerator", "NewEnum", 9), ("ExcelHotkey", "Run", 12) },
            read.Annotations.Select(one => (one.Kind.ToString(), one.Target!, one.TargetLine!.Value)).ToArray());
    }

    [Fact]
    public void AVariableDescriptionBindsToTheVariableBelow()
    {
        var read = AttributeAnnotations.Read(Lines(
            "Option Explicit",
            "'@VariableDescription(\"How many.\")",
            "Public Count As Long",
            "Public Sub Go()",
            "End Sub"));

        Assert.Empty(read.Problems);
        var one = Assert.Single(read.Annotations);
        Assert.Equal(AnnotationKind.VariableDescription, one.Kind);
        Assert.Equal("Count", one.Target);
        Assert.Equal(3, one.TargetLine);
    }

    [Fact]
    public void AModuleAnnotationBelowTheFirstProcedureIsAProblem()
    {
        var read = AttributeAnnotations.Read(Lines("Public Sub Go()", "End Sub", "'@PredeclaredId"));

        Assert.Empty(read.Annotations);
        var problem = Assert.Single(read.Problems);
        Assert.Equal(3, problem.Line);
        Assert.Contains("declarations section", problem.Message);
    }

    [Fact]
    public void AMemberAnnotationAboveAVariableAndAVariableOneAboveAProcedureAreProblems()
    {
        var read = AttributeAnnotations.Read(Lines(
            "'@Description(\"x\")",
            "Private total As Long",
            "'@VariableDescription(\"y\")",
            "Public Sub Go()",
            "End Sub"));

        Assert.Empty(read.Annotations);
        Assert.Equal(2, read.Problems.Count);
        Assert.Contains("is a variable", read.Problems[0].Message);
        Assert.Contains("is a procedure", read.Problems[1].Message);
    }

    [Theory]
    [InlineData("'@ExcelHotkey(\"DD\")")]
    [InlineData("'@ExcelHotkey(\"1\")")]
    [InlineData("'@ExcelHotkey")]
    [InlineData("'@Description")]
    public void ABadArgumentIsAProblemNotAnAnnotation(string line)
    {
        var read = AttributeAnnotations.Read(Lines("Option Explicit", line, "Public Sub Go()", "End Sub"));

        Assert.Empty(read.Annotations);
        Assert.Equal(2, Assert.Single(read.Problems).Line);
    }

    [Fact]
    public void ASecondDefaultMemberIsAProblem()
    {
        var read = AttributeAnnotations.Read(Lines(
            "'@DefaultMember", "Public Sub A()", "End Sub",
            "'@DefaultMember", "Public Sub B()", "End Sub"));

        Assert.Single(read.Annotations);
        Assert.Contains("one default member", Assert.Single(read.Problems).Message);
    }

    [Fact]
    public void AnAnnotationAboveNothingOrAboveAStatementIsAProblem()
    {
        var read = AttributeAnnotations.Read(Lines(
            "Public Sub Go()",
            "    '@Description(\"inside\")",
            "    Debug.Print 1",
            "End Sub",
            "'@Description(\"trailing\")"));

        Assert.Empty(read.Annotations);
        Assert.Equal(2, read.Problems.Count);
        Assert.Contains("not a procedure", read.Problems[0].Message);
        Assert.Contains("above nothing", read.Problems[1].Message);
    }

    [Fact]
    public void SomebodyElsesAnnotationsAreLeftAlone()
    {
        var read = AttributeAnnotations.Read(Lines(
            "'@Folder(\"Accounts\")",
            "'@IgnoreModule",
            "Option Explicit",
            "'@TestMethod",
            "Public Sub Go()",
            "End Sub"));

        Assert.Empty(read.Annotations);
        Assert.Empty(read.Problems);
    }

    [Fact]
    public void TheCanonicalSpellingIsTheDocumentedOne()
    {
        Assert.Equal("'@Description(\"x\")", new Annotation(AnnotationKind.Description, 1, "x", "A", 2).Canonical);
        Assert.Equal("'@PredeclaredId", new Annotation(AnnotationKind.PredeclaredId, 1, null, null, null).Canonical);
    }
}
