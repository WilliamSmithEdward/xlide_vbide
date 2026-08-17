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

    /// <summary>
    /// A colour is spelled two ways because it IS two things. A literal is `#rrggbb`; a system
    /// colour is a question rather than a value - what does this machine call a button face - and
    /// it keeps its NAME so the question survives a round trip. Spelling a system colour as
    /// today's answer would freeze it and the control would stop following the theme.
    /// </summary>
    [Fact]
    public void ASystemColourKeepsItsNameAndALiteralItsHex()
    {
        Assert.Equal("ButtonFace", FormMarkup.SpellColour(unchecked((int)0x8000000F)));
        Assert.Equal("Highlight", FormMarkup.SpellColour(unchecked((int)0x8000000D)));
        Assert.Equal("#c0c0c0", FormMarkup.SpellColour(0x00C0C0C0));

        // Spaces, hyphens and case are forgiven, so the name the PANEL shows and the name the
        // document writes are one colour read by one function.
        Assert.Equal(unchecked((int)0x8000000F), FormMarkup.ReadColourName("ButtonFace"));
        Assert.Equal(unchecked((int)0x8000000F), FormMarkup.ReadColourName("Button Face"));
        Assert.Equal(unchecked((int)0x8000000F), FormMarkup.ReadColourName("buttonface"));
        Assert.Equal(unchecked((int)0x80000008), FormMarkup.ReadColourName("WindowText"));
        Assert.Equal(unchecked((int)0x8000001A), FormMarkup.ReadColourName("HotTrackedItem"));

        // A word that names no colour is not one, which is what lets a bare word be read as a
        // colour at all without ever swallowing something else.
        Assert.Null(FormMarkup.ReadColourName("Chartreuse"));
        Assert.Null(FormMarkup.ReadColourName(""));

        // A system index Win32 gave no name to keeps VBA's own hex rather than inventing one.
        Assert.Equal("&H80000019&", FormMarkup.SpellColour(unchecked((int)0x80000019)));
    }

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
    public void TheDialectKeepsVbasStringsAndFlags()
    {
        var parsed = FormMarkup.Parse(""""
            <!-- A comment above the form, and one inside it. -->
            <Form Name="Entry" Caption="It's ""quoted""" Width=100 Height=80>
                <!-- the caption above holds an apostrophe AND doubled quotes -->
                <Label Name="Hint" Caption="say ' nothing" Left=1 Top=2 Width=3 Height=4
                       ForeColor=ButtonFace Visible=False />
            </Form>
            """");

        Assert.Equal("It's \"quoted\"", parsed.Caption);
        var hint = Assert.Single(parsed.Controls);
        Assert.Equal("say ' nothing", hint.Caption);
        Assert.Equal(unchecked((int)0x8000000F).ToString(), hint.Properties[0].Value);
        Assert.Equal(PropertyValueKind.Colour, hint.Properties[0].Kind);
        Assert.Equal("False", hint.Properties[1].Value);
        Assert.Equal(PropertyValueKind.Flag, hint.Properties[1].Kind);
    }

    /// <summary>
    /// An element may wrap across lines, because an attribute list gets long once a control
    /// carries what the saved baseline says it changed. The scanner reads TAGS rather than lines,
    /// so the only thing a newline inside a tag changes is which line an error is reported on.
    /// </summary>
    [Fact]
    public void AnElementMayWrapAcrossLines()
    {
        var parsed = FormMarkup.Parse("""
            <Form Name="F">
                <TextBox Name="Box"
                         Left=1 Top=2 Width=3 Height=4
                         MaxLength=12
                         ControlSource="Sheet1!A1" />
            </Form>
            """);

        var box = Assert.Single(parsed.Controls);
        Assert.Equal(1, box.Left);
        Assert.Equal(2, box.Properties.Count);
        Assert.Equal("Sheet1!A1", box.Properties[1].Value);
        Assert.Equal(PropertyValueKind.Text, box.Properties[1].Kind);
    }

    /// <summary>
    /// A colour is read from every spelling the dialect takes: `#rrggbb`, the shorthand, a system
    /// colour's NAME, and VBA's own hex from a hand-written document.
    /// </summary>
    [Theory]
    [InlineData("#c0dcc0", 12639424)]
    [InlineData("#ff8000", 33023)]
    [InlineData("#000000", 0)]
    [InlineData("#fff", 16777215)]
    [InlineData("&H8000000F&", unchecked((int)0x8000000F))]
    [InlineData("ButtonFace", unchecked((int)0x8000000F))]
    [InlineData("Highlight", unchecked((int)0x8000000D))]
    public void AColourIsReadFromEverySpelling(string spelled, int stored)
    {
        var parsed = FormMarkup.Parse($"<Form Name=\"F\">\n    <Label Name=\"L\" BackColor={spelled} />\n</Form>\n");

        var colour = Assert.Single(Assert.Single(parsed.Controls).Properties);
        Assert.Equal(stored.ToString(), colour.Value);
        Assert.Equal(PropertyValueKind.Colour, colour.Kind);
    }

    [Fact]
    public void ALiteralPrintsAsHexAndASystemColourAsItsName()
    {
        var printed = FormMarkup.Print(new FormSpec(
            "F", null, null, null,
            [new PropertySpec("BackColor", "12639424", PropertyValueKind.Colour)],
            [Control("Label", "L", null, null, null, null, null, null,
                new PropertySpec("ForeColor", unchecked((int)0x8000000F).ToString(), PropertyValueKind.Colour))]));

        Assert.Contains("BackColor=\"#c0dcc0\"", printed);
        Assert.Contains("ForeColor=\"ButtonFace\"", printed);

        // And it round trips: what the printer wrote, the parser reads back to the same numbers.
        var back = FormMarkup.Parse(printed);
        Assert.Equal("12639424", back.Properties[0].Value);
        Assert.Equal(unchecked((int)0x8000000F).ToString(), back.Controls[0].Properties[0].Value);
    }

    [Fact]
    public void AColourThatIsNotOneIsRefusedByItsLine()
    {
        var refused = Assert.Throws<FormMarkupException>(() =>
            FormMarkup.Parse("<Form Name=\"F\">\n    <Label Name=\"L\" BackColor=#ggg />\n</Form>\n"));

        Assert.Equal(2, refused.Line);
        Assert.Contains("#rrggbb", refused.Reason);
    }

    /// <summary>
    /// Everything the grammar refuses, each with the line that earned it. A document that loses
    /// something QUIETLY is worse than one that will not parse, which is why element content and a
    /// repeated attribute are errors rather than shrugs.
    /// </summary>
    [Theory]
    [InlineData("", 1, "empty")]
    [InlineData("<Label Name=\"Hint\" />\n", 1, "<Form>")]
    [InlineData("<Form Name=\"F\"></Form>\n<Label Name=\"H\" />\n", 2, "one Form per document")]
    [InlineData("<Form Name=\"F\">\n    <Form Name=\"G\" />\n</Form>\n", 2, "holds no Form")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" />\n", 2, "never closed")]
    [InlineData("<Form Name=\"F\">\n    <Frame Name=\"B\">\n    </Label>\n</Form>\n", 3, "tags nest")]
    // Inside a Form the stack is never empty, so a stray close is a MISMATCH rather than an
    // orphan - it is closing the Form, just not by that name.
    [InlineData("<Form Name=\"F\">\n    </Label>\n</Form>\n", 2, "tags nest")]
    [InlineData("</Label>\n", 1, "closes nothing")]
    [InlineData("<Form Name=\"F\">\n    <Label />\n</Form>\n", 2, "needs a Name")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Left=1 />\n</Form>\n", 2, "come together")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Width=1 />\n</Form>\n", 2, "come together")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Left=x Top=1 />\n</Form>\n", 2, "takes a number")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Caption=maybe />\n</Form>\n", 2, "not a value")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Caption= />\n</Form>\n", 2, "missing its value")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Caption />\n</Form>\n", 2, "missing its value")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Top=1 Top=2 />\n</Form>\n", 2, "set twice")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\" Caption=\"never closed />\n</Form>\n", 2, "never closes")]
    [InlineData("<Form Name=\"F\">\n    stray words\n</Form>\n", 2, "text between tags")]
    [InlineData("<Form Name=\"F\">\n    <!-- never closed\n</Form>\n", 2, "comment opens")]
    [InlineData("<Form Name=\"F\">\n    <MultiPage Name=\"M\">\n        <Label Name=\"H\" />\n", 3, "holds Pages")]
    [InlineData("<Form Name=\"F\">\n    <Page Name=\"P\" />\n</Form>\n", 2, "under a MultiPage")]
    [InlineData("<Form Name=\"F\">\n    <Tab Name=\"T\" />\n</Form>\n", 2, "under a TabStrip")]
    [InlineData("<Form Name=\"F\">\n    <TabStrip Name=\"S\">\n        <Label Name=\"H\" />\n", 3, "holds Tabs")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"H\">\n        <TextBox Name=\"I\" />\n", 3, "holds no controls")]
    [InlineData("<Form Name=\"F\">\n    <Label Name=\"9Lives\" />\n</Form>\n", 2, "not a name")]
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
            <Form Name="F">
                <Frame Name="Box" Left=1 Top=1 Width=50 Height=50>
                    <Label Name="Inner" Left=2 Top=2 Width=10 Height=10 />
                </Frame>
                <Label Name="Outer" Left=60 Top=1 Width=10 Height=10 />
            </Form>
            """);

        Assert.Equal("Box", parsed.Controls.Single(control => control.Name == "Inner").Parent);
        Assert.Null(parsed.Controls.Single(control => control.Name == "Outer").Parent);
    }

    [Fact]
    public void GeometryHalvesMayBeAbsent()
    {
        // The walk answers null geometry for Pages, and a hand-written element may say only where
        // or only how big. Printing carries whichever halves exist as pairs.
        var parsed = FormMarkup.Parse("<Form Name=\"F\">\n    <Label Name=\"Hint\" Width=10 Height=20 />\n</Form>\n");
        var hint = Assert.Single(parsed.Controls);
        Assert.Null(hint.Left);
        Assert.Equal(10, hint.Width);

        Assert.Contains("Width=\"10\" Height=\"20\"", FormMarkup.Print(parsed));
        Assert.DoesNotContain("Left=", FormMarkup.Print(parsed));
    }

    [Fact]
    public void CaseIsAsForgivingAsTheHost()
    {
        var parsed = FormMarkup.Parse(
            "<form name=\"F\" CAPTION=\"hi\" width=10 height=10>\n    <LABEL Name=\"Hint\" LEFT=1 top=2 />\n</FORM>\n");
        Assert.Equal("hi", parsed.Caption);
        Assert.Equal(1, Assert.Single(parsed.Controls).Left);
    }

    [Fact]
    public void LintOfACleanDocumentFindsNothing()
    {
        Assert.Empty(FormMarkup.Lint(
            "<Form Name=\"F\" Caption=\"hi\" Width=100 Height=100>\n"
            + "    <Label Name=\"Hint\" Left=1 Top=2 Width=3 Height=4 />\n</Form>\n"));
    }

    [Fact]
    public void LintCollectsEveryRefusalWhereParseStopsAtTheFirst()
    {
        // Two bad elements. Parse dies on the first; the lint reports both, each at its own line,
        // because each round retires exactly the line the parser refused.
        var findings = FormMarkup.Lint(
            "<Form Name=\"F\" Width=100 Height=100>\n"
            + "    <Label Name=\"A\" Left=banana Top=1 />\n"
            + "    <Label Name=\"B\" Left=1 Top=2 Width=3 Height=4 />\n"
            + "    <Label Name=\"C\" Width=pear Height=1 />\n</Form>\n");

        Assert.Equal(2, findings.Count);
        Assert.Equal(2, findings[0].Line);
        Assert.Equal(4, findings[1].Line);
        Assert.All(findings, finding => Assert.Equal(FormMarkupSeverity.Error, finding.Severity));
    }

    [Fact]
    public void LintAnchorsADuplicateNameAtItsSecondMention()
    {
        var findings = FormMarkup.Lint(
            "<Form Name=\"F\" Width=100 Height=100>\n"
            + "    <Label Name=\"Twin\" Left=1 Top=1 Width=2 Height=2 />\n"
            + "    <TextBox Name=\"Twin\" Left=9 Top=9 Width=2 Height=2 />\n</Form>\n");

        var duplicate = Assert.Single(findings);
        Assert.Equal(3, duplicate.Line);
        Assert.Equal(FormMarkupSeverity.Error, duplicate.Severity);
        Assert.Contains("already taken", duplicate.Message);
    }

    [Fact]
    public void LintWarnsWhatAnApplyWouldSkip()
    {
        // A foreign type with no ProgId is the apply's documented skip-with-a-note case, so the
        // lint says so BEFORE the apply. A stray Page beside it arrives as an ERROR - the parser
        // refuses one structurally, and the lint never re-judges the grammar.
        var findings = FormMarkup.Lint(
            "<Form Name=\"F\" Width=100 Height=100>\n"
            + "    <Gadget Name=\"Widget\" Left=1 Top=1 Width=2 Height=2 />\n"
            + "    <Page Name=\"Stray\" Caption=\"P\" />\n</Form>\n");

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
            "<Form Name=\"F\" Width=100 Height=100>\n"
            + "    <Gadget Name=\"Widget\" Left=1 Top=1 Width=2 Height=2 ProgId=\"Vendor.Widget.1\" />\n</Form>\n"));
    }

    [Fact]
    public void LintRefusesANameMsFormsWouldNotTake()
    {
        // The model's own answer for one of these is `error 800a9c6c` and nothing else, so the
        // squiggle is the difference between a typo and a mystery.
        var findings = FormMarkup.Lint(
            "<Form Name=\"F\" Width=100 Height=100>\n"
            + "    <Label Name=\"_Leading\" Left=1 Top=1 Width=2 Height=2 />\n"
            + "    <Label Name=\"9Lives\" Left=3 Top=3 Width=2 Height=2 />\n</Form>\n");

        Assert.Equal(2, findings.Count);
        Assert.All(findings, finding => Assert.Equal(FormMarkupSeverity.Error, finding.Severity));
    }

    [Fact]
    public void AnEmptyDocumentSaysSoOnce()
    {
        // The tolerant pass retires a refused line and parses again. Retiring a line that is
        // already empty makes no progress, so an empty document refused, blanked nothing, and
        // refused again - two identical squiggles on line 1 (found in the 2026-08-16 hunt).
        var finding = Assert.Single(FormMarkup.Lint(string.Empty));
        Assert.Equal(1, finding.Line);
        Assert.Contains("empty", finding.Message);
    }

    [Fact]
    public void AnOrdinaryNameIsNotSquiggled()
    {
        Assert.Empty(FormMarkup.Lint(
            "<Form Name=\"F\" Width=100 Height=100>\n"
            + "    <Label Name=\"Name_2\" Left=1 Top=1 Width=2 Height=2 />\n"
            + "    <TextBox Name=\"b\" Left=3 Top=3 Width=2 Height=2 />\n</Form>\n"));
    }
}