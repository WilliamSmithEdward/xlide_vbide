using System.Linq;
using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Reading attributes out of an export and writing them back in. The rewrite is about to be
/// imported over the developer's module, so what it leaves alone matters as much as what it
/// sets: every case here checks the whole text, not the one line.
/// </summary>
public class AttributeRewriteTests
{
    private const string Crlf = "\r\n";

    private static string Lines(params string[] lines) => string.Join(Crlf, lines);

    private static readonly string ClassExport = Lines(
        "VERSION 1.0 CLASS",
        "BEGIN",
        "  MultiUse = -1  'True",
        "END",
        "Attribute VB_Name = \"Bag\"",
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = False",
        "Attribute VB_Exposed = False",
        "Option Explicit",
        "",
        "Private items As Collection",
        "",
        "Public Function Item(ByVal index As Long) As Variant",
        "Attribute Item.VB_UserMemId = 0",
        "    Item = items(index)",
        "End Function",
        "",
        "Public Function NewEnum() As IUnknown",
        "Attribute NewEnum.VB_UserMemId = -4",
        "Attribute NewEnum.VB_Description = \"Walks the items.\"",
        "    Set NewEnum = items.[_NewEnum]",
        "End Function",
        "");

    [Fact]
    public void TheAttributesOfAnExportAreRead()
    {
        var read = ModuleAttributes.Read(ClassExport);

        Assert.Equal("Bag", read.Name);
        Assert.Null(read.Description);
        Assert.False(read.PredeclaredId);
        Assert.False(read.Exposed);
        Assert.Equal(0, read.Member("Item").UserMemId);
        Assert.Equal(-4, read.Member("NewEnum").UserMemId);
        Assert.Equal("Walks the items.", read.Member("NewEnum").Description);
        Assert.True(read.Member("Nothing").IsEmpty);
    }

    [Fact]
    public void AHotkeyAndADoubledQuoteAndAVariableDescriptionAreRead()
    {
        var read = ModuleAttributes.Read(Lines(
            "Attribute VB_Name = \"Macros\"",
            "Attribute VB_Description = \"He said \"\"hi\"\"\"",
            "Public Count As Long",
            "Attribute Count.VB_VarDescription = \"How many.\"",
            "Public Sub Run()",
            "Attribute Run.VB_ProcData.VB_Invoke_Func = \"D\\n14\"",
            "End Sub"));

        Assert.Equal("He said \"hi\"", read.Description);
        Assert.Equal("D", read.Member("Run").Hotkey);
        Assert.Equal("How many.", read.VariableDescriptions["Count"]);
        Assert.Null(read.PredeclaredId);
    }

    [Fact]
    public void ApplyingWritesEveryAnnotatedAttributeAndNothingElse()
    {
        var code = Lines(
            "'@ModuleDescription(\"A bag.\")",
            "'@PredeclaredId",
            "Option Explicit",
            "",
            "'@VariableDescription(\"The items.\")",
            "Private items As Collection",
            "",
            "'@DefaultMember",
            "'@Description(\"One item.\")",
            "Public Function Item(ByVal index As Long) As Variant",
            "    Item = items(index)",
            "End Function",
            "",
            "'@Enumerator",
            "Public Function NewEnum() As IUnknown",
            "    Set NewEnum = items.[_NewEnum]",
            "End Function",
            "");
        var export = Lines(
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"Bag\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_Exposed = False",
            "'@ModuleDescription(\"A bag.\")",
            "'@PredeclaredId",
            "Option Explicit",
            "",
            "'@VariableDescription(\"The items.\")",
            "Private items As Collection",
            "",
            "'@DefaultMember",
            "'@Description(\"One item.\")",
            "Public Function Item(ByVal index As Long) As Variant",
            "    Item = items(index)",
            "End Function",
            "",
            "'@Enumerator",
            "Public Function NewEnum() As IUnknown",
            "Attribute NewEnum.VB_UserMemId = -4",
            "    Set NewEnum = items.[_NewEnum]",
            "End Function",
            "");

        var result = AttributeRewriter.Apply(export, AttributeAnnotations.Read(code));

        Assert.Empty(result.Skipped);
        Assert.Equal(Lines(
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"Bag\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = False",
            "Attribute VB_Description = \"A bag.\"",
            "'@ModuleDescription(\"A bag.\")",
            "'@PredeclaredId",
            "Option Explicit",
            "",
            "'@VariableDescription(\"The items.\")",
            "Private items As Collection",
            "Attribute items.VB_VarDescription = \"The items.\"",
            "",
            "'@DefaultMember",
            "'@Description(\"One item.\")",
            "Public Function Item(ByVal index As Long) As Variant",
            "Attribute Item.VB_Description = \"One item.\"",
            "Attribute Item.VB_UserMemId = 0",
            "    Item = items(index)",
            "End Function",
            "",
            "'@Enumerator",
            "Public Function NewEnum() As IUnknown",
            "Attribute NewEnum.VB_UserMemId = -4",
            "    Set NewEnum = items.[_NewEnum]",
            "End Function",
            ""), result.Text);
        Assert.Equal(4, result.Changes.Count(change => change.From is null));
        Assert.Contains(result.Changes, change => change.Attribute == "VB_PredeclaredId" && change.From == "False" && change.To == "True");

        // Applying again changes nothing: the second run is the proof the first was complete.
        var again = AttributeRewriter.Apply(result.Text, AttributeAnnotations.Read(code));
        Assert.Empty(again.Changes);
        Assert.Equal(result.Text, again.Text);
    }

    [Fact]
    public void AnAnnotationAboveASpecificPropertyLegLandsOnThatLeg()
    {
        var code = Lines(
            "Option Explicit",
            "Public Property Get Total() As Long",
            "End Property",
            "'@Description(\"Sets it.\")",
            "Public Property Let Total(ByVal v As Long)",
            "End Property");
        var export = Lines(
            "Attribute VB_Name = \"M\"",
            "Option Explicit",
            "Public Property Get Total() As Long",
            "End Property",
            "'@Description(\"Sets it.\")",
            "Public Property Let Total(ByVal v As Long)",
            "End Property");

        var result = AttributeRewriter.Apply(export, AttributeAnnotations.Read(code));

        Assert.Equal(Lines(
            "Attribute VB_Name = \"M\"",
            "Option Explicit",
            "Public Property Get Total() As Long",
            "End Property",
            "'@Description(\"Sets it.\")",
            "Public Property Let Total(ByVal v As Long)",
            "Attribute Total.VB_Description = \"Sets it.\"",
            "End Property"), result.Text);
    }

    [Fact]
    public void AContinuedHeaderTakesItsAttributeAfterItsLastLine()
    {
        var code = Lines("'@Description(\"Long.\")", "Public Sub Go(ByVal a As Long, _", "    ByVal b As Long)", "End Sub");
        var export = Lines("Attribute VB_Name = \"M\"", "'@Description(\"Long.\")", "Public Sub Go(ByVal a As Long, _", "    ByVal b As Long)", "End Sub");

        var result = AttributeRewriter.Apply(export, AttributeAnnotations.Read(code));

        Assert.Equal(Lines("Attribute VB_Name = \"M\"", "'@Description(\"Long.\")", "Public Sub Go(ByVal a As Long, _", "    ByVal b As Long)", "Attribute Go.VB_Description = \"Long.\"", "End Sub"), result.Text);
    }

    [Fact]
    public void AStandardModuleCannotTakePredeclaredIdAndSaysSo()
    {
        var code = Lines("'@PredeclaredId", "Option Explicit");
        var export = Lines("Attribute VB_Name = \"M\"", "'@PredeclaredId", "Option Explicit");

        var result = AttributeRewriter.Apply(export, AttributeAnnotations.Read(code));

        Assert.Equal(export, result.Text);
        Assert.Empty(result.Changes);
        Assert.Contains("VB_PredeclaredId", Assert.Single(result.Skipped));
    }

    [Fact]
    public void RemovingTakesOneAttributeAwayAndLeavesTheRest()
    {
        var predeclared = ClassExport.Replace("Attribute VB_PredeclaredId = False", "Attribute VB_PredeclaredId = True");

        var offPredeclared = AttributeRewriter.Remove(predeclared, AnnotationKind.PredeclaredId, null);
        Assert.Equal(ClassExport, offPredeclared.Text);

        var noEnumerator = AttributeRewriter.Remove(ClassExport, AnnotationKind.Enumerator, "NewEnum");
        Assert.DoesNotContain("NewEnum.VB_UserMemId", noEnumerator.Text);
        Assert.Contains("NewEnum.VB_Description", noEnumerator.Text);
        Assert.Contains("Item.VB_UserMemId = 0", noEnumerator.Text);

        var noDescription = AttributeRewriter.Remove(ClassExport, AnnotationKind.Description, "NewEnum");
        Assert.DoesNotContain("VB_Description", noDescription.Text);
        Assert.Equal("NewEnum: VB_Description removed", noDescription.Changes.Single().ToString());
    }

    [Fact]
    public void ALineEndingIsKeptAsItCame()
    {
        var export = "Attribute VB_Name = \"M\"\n'@Description(\"x\")\nPublic Sub Go()\nEnd Sub\n";
        var result = AttributeRewriter.Apply(export, AttributeAnnotations.Read("'@Description(\"x\")\nPublic Sub Go()\nEnd Sub\n"));

        Assert.Equal("Attribute VB_Name = \"M\"\n'@Description(\"x\")\nPublic Sub Go()\nAttribute Go.VB_Description = \"x\"\nEnd Sub\n", result.Text);
    }
}
