using Xlide.Vbe.Core.Forms;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The form markup, both directions.
///
/// The language is a projection over a real design, so the properties that matter are the
/// projection's: printing is canonical and deterministic, parse(print(x)) gives back x, and a
/// document with an error parses to NOTHING rather than to most of a form - the apply side
/// promises that a bad document changes nothing, and this is where that promise is earned.
/// </summary>
public class FormMarkupTests
{
    private static PropertySpec Text(string path, string value) => new(path, value, PropertyValueKind.Text);

    private static PropertySpec Number(string path, string value) => new(path, value, PropertyValueKind.Number);

    private static PropertySpec Flag(string path, bool value) =>
        new(path, value ? "True" : "False", PropertyValueKind.Flag);

    private static ControlSpec Control(
        string type, string name, string? caption = null, string? parent = null,
        double? left = null, double? top = null, double? width = null, double? height = null,
        params PropertySpec[] properties) =>
        new(type, name, caption, left, top, width, height, parent, properties);

    private static FormSpec Sample() => new(
        "EntryForm", "Quarter Entry", 360, 320,
        [Number("BackColor", "-2147483633")],
        [
            Control("Label", "NameLabel", "Customer", null, 12, 14, 66, 16),
            Control("TextBox", "NameBox", null, null, 84, 12, 120, 20,
                Flag("Enabled", true), Number("Font.Size", "12")),
            Control("Frame", "Options", "Freight", null, 12, 112, 92, 66),
            Control("OptionButton", "PickGround", "Ground", "Options", 8, 14, 76, 16),
            Control("OptionButton", "PickAir", "Air", "Options", 8, 34, 76, 16),
            Control("MultiPage", "Wizard", null, null, 12, 188, 192, 86),
            Control("Page", "Page1", "Page1", "Wizard"),
            Control("CheckBox", "Agree", "Agreed", "Page1", 8, 8, 100, 16),
            Control("Page", "Page2", "Page2", "Wizard"),
            Control("CommandButton", "OkButton", "Start", null, 262, 250, 72, 24,
                Text("ControlTipText", "Runs the build")),
        ]);

    [Fact]
    public void PrintingIsStableThroughAParse()
    {
        var printed = FormMarkup.Print(Sample());
        var reparsed = FormMarkup.Parse(printed);
        Assert.Equal(printed, FormMarkup.Print(reparsed));
    }

    [Fact]
    public void AParseCarriesEverythingThePrintSaid()
    {
        var parsed = FormMarkup.Parse(FormMarkup.Print(Sample()));

        Assert.Equal("EntryForm", parsed.Name);
        Assert.Equal("Quarter Entry", parsed.Caption);
        Assert.Equal(360, parsed.Width);
        Assert.Equal(320, parsed.Height);
        Assert.Equal("-2147483633", Assert.Single(parsed.Properties).Value);

        Assert.Equal(10, parsed.Controls.Count);

        var pick = parsed.Controls.Single(control => control.Name == "PickAir");
        Assert.Equal("Options", pick.Parent);
        Assert.Equal("Air", pick.Caption);
        Assert.Equal(34, pick.Top);

        var agree = parsed.Controls.Single(control => control.Name == "Agree");
        Assert.Equal("Page1", agree.Parent);

        var page = parsed.Controls.Single(control => control.Name == "Page2");
        Assert.Equal("Wizard", page.Parent);
        Assert.Null(page.Left);

        var box = parsed.Controls.Single(control => control.Name == "NameBox");
        Assert.Equal(2, box.Properties.Count);
        Assert.Equal(PropertyValueKind.Flag, box.Properties[0].Kind);
        Assert.Equal("Font.Size", box.Properties[1].Path);
    }

    [Fact]
    public void TheDialectIsVbaFlavoured()
    {
        var parsed = FormMarkup.Parse(""""
            ' A comment above the form, and one after a header.
            Form Entry "It's ""quoted""" size 100x80   ' the caption holds an apostrophe AND quotes
                Label Hint "say ' nothing" at 1,2 size 3x4
                    ForeColor = &H8000000F&
                    Visible = False
            """");

        Assert.Equal("It's \"quoted\"", parsed.Caption);
        var hint = Assert.Single(parsed.Controls);
        Assert.Equal("say ' nothing", hint.Caption);
        Assert.Equal(unchecked((int)0x8000000F).ToString(), hint.Properties[0].Value);
        Assert.Equal(PropertyValueKind.Number, hint.Properties[0].Kind);
        Assert.Equal("False", hint.Properties[1].Value);
    }

    [Theory]
    [InlineData("", 1, "empty")]
    [InlineData("Label Hint\n", 1, "Form line")]
    [InlineData("Form F\nLabel Hint\n", 2, "unindented")]
    [InlineData("Form F\n  Label Hint\n", 2, "4 spaces")]
    [InlineData("Form F\n\tLabel Hint\n", 2, "tab")]
    [InlineData("Form F\n        Label Hint\n", 2, "under nothing")]
    [InlineData("Form F\n    Label Hint trailing junk here\n", 2, "rest of the line")]
    [InlineData("Form F\n    Label Hint \"never closed\n", 2, "never closes")]
    [InlineData("Form F\n    Label Hint at 1\n", 2, "two numbers")]
    [InlineData("Form F\n    Label Hint\n        Caption = maybe\n", 3, "not a value")]
    [InlineData("Form F\n    Label Hint\n        Caption =\n", 3, "missing")]
    [InlineData("Form F\n    MultiPage M\n        Label Hint\n", 3, "holds Pages")]
    [InlineData("Form F\n    Page P\n", 2, "under a MultiPage")]
    [InlineData("Form F\n    Label Hint\n        TextBox Inner\n", 3, "holds no controls")]
    public void ABadDocumentIsRefusedWithItsLine(string document, int line, string names)
    {
        var refused = Assert.Throws<FormMarkupException>(() => FormMarkup.Parse(document));
        Assert.Equal(line, refused.Line);
        Assert.Contains(names, refused.Message);
    }

    [Fact]
    public void SteppingBackOutOfAContainerLandsOnTheForm()
    {
        var parsed = FormMarkup.Parse("""
            Form F
                Frame Box at 1,1 size 50x50
                    Label Inner at 2,2 size 10x10
                Label Outer at 60,1 size 10x10
            """);

        Assert.Equal("Box", parsed.Controls.Single(control => control.Name == "Inner").Parent);
        Assert.Null(parsed.Controls.Single(control => control.Name == "Outer").Parent);
    }

    [Fact]
    public void GeometryHalvesMayBeAbsent()
    {
        // The walk answers null geometry for Pages, and a hand-written line may say only where
        // or only how big. Printing carries whichever halves exist as pairs.
        var parsed = FormMarkup.Parse("Form F\n    Label Hint size 10x20\n");
        var hint = Assert.Single(parsed.Controls);
        Assert.Null(hint.Left);
        Assert.Equal(10, hint.Width);

        Assert.Contains("size 10x20", FormMarkup.Print(parsed));
        Assert.DoesNotContain(" at ", FormMarkup.Print(parsed));
    }

    [Fact]
    public void CaseIsAsForgivingAsTheHost()
    {
        var parsed = FormMarkup.Parse("form F \"hi\" SIZE 10x10\n    LABEL Hint AT 1,2 Size 3x4\n");
        Assert.Equal("hi", parsed.Caption);
        Assert.Equal(1, Assert.Single(parsed.Controls).Left);
    }

    [Fact]
    public void LintOfACleanDocumentFindsNothing()
    {
        Assert.Empty(FormMarkup.Lint("Form F \"hi\" size 100x100\n    Label Hint at 1,2 size 3x4\n"));
    }

    [Fact]
    public void LintCollectsEveryRefusalWhereParseStopsAtTheFirst()
    {
        // Two bad lines. Parse dies on line 2; the lint reports both, each at its own line,
        // because each round retires exactly the line the parser refused.
        var findings = FormMarkup.Lint(
            "Form F size 100x100\n    Label A at banana\n    Label B at 1,2 size 3x4\n    Label C size pear\n");

        Assert.Equal(2, findings.Count);
        Assert.Equal(2, findings[0].Line);
        Assert.Equal(4, findings[1].Line);
        Assert.All(findings, finding => Assert.Equal(FormMarkupSeverity.Error, finding.Severity));
    }

    [Fact]
    public void LintAnchorsADuplicateNameAtItsSecondMention()
    {
        var findings = FormMarkup.Lint(
            "Form F size 100x100\n    Label Twin at 1,1 size 2x2\n    TextBox Twin at 9,9 size 2x2\n");

        var duplicate = Assert.Single(findings);
        Assert.Equal(3, duplicate.Line);
        Assert.Equal(FormMarkupSeverity.Error, duplicate.Severity);
        Assert.Contains("already taken", duplicate.Message);
    }

    [Fact]
    public void LintWarnsWhatAnApplyWouldSkip()
    {
        // A foreign type with no ProgId line is the apply's documented skip-with-a-note case,
        // so the lint says so BEFORE the apply. A stray Page beside it arrives as an ERROR -
        // the parser refuses one structurally, and the lint never re-judges the grammar.
        var findings = FormMarkup.Lint(
            "Form F size 100x100\n    Gadget Widget at 1,1 size 2x2\n    Page Stray \"P\"\n");

        Assert.Equal(2, findings.Count);
        var widget = findings.Single(finding => finding.Line == 2);
        Assert.Equal(FormMarkupSeverity.Warning, widget.Severity);
        Assert.Contains("ProgId", widget.Message);
        var stray = findings.Single(finding => finding.Line == 3);
        Assert.Equal(FormMarkupSeverity.Error, stray.Severity);
        Assert.Contains("MultiPage", stray.Message);
    }

    [Fact]
    public void LintForgivesAForeignTypeThatNamesItsProgId()
    {
        Assert.Empty(FormMarkup.Lint(
            "Form F size 100x100\n    Gadget Widget at 1,1 size 2x2\n        ProgId = \"Vendor.Widget.1\"\n"));
    }

    [Fact]
    public void LintReportsOrphansOfARetiredContainerAsTheirOwnFindings()
    {
        // The Frame line is bad; once retired, its child is indented under nothing - which is
        // true of the document as it stands, and both lines deserve their squiggle.
        var findings = FormMarkup.Lint(
            "Form F size 100x100\n    Frame Box at banana\n        Label Inner at 1,1 size 2x2\n");

        Assert.Equal(2, findings.Count);
        Assert.Equal(2, findings[0].Line);
        Assert.Equal(3, findings[1].Line);
    }
}
