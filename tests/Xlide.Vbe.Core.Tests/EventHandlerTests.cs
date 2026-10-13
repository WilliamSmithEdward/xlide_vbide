using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Where a handler picked in the code window's lists lands (#101). Each case is one the VBE's own
/// CreateEventProc was run on, 2026-10-12, with the text it produced: the stub sits before the
/// module's next event handler in name order, of any object, ahead of that handler's comments,
/// and at the end when there is none; ordinary procedures do not count.
/// </summary>
public class EventHandlerTests
{
    private static string Text(params string[] lines) => string.Join("\r\n", lines);

    /// <summary>The handlers a sheet, a workbook and a class with an Application variable can have, in part.</summary>
    private static readonly HashSet<string> Handlers = new(StringComparer.OrdinalIgnoreCase)
    {
        "Workbook_Activate", "Workbook_Open",
        "Worksheet_Activate", "Worksheet_Change", "Worksheet_SelectionChange", "Worksheet_TableUpdate",
        "Class_Initialize", "Class_Terminate", "app_AfterCalculate", "app_WorkbookOpen",
    };

    [Fact]
    public void TheFirstHandlerGoesAtTheEnd()
    {
        var source = Text("Public Sub Beta()", "End Sub", "", "Public Sub Yak()", "End Sub", "");

        var (text, header) = EventHandlers.Place(source, "Workbook", "Open", "", Handlers);

        Assert.Equal(Text("Public Sub Beta()", "End Sub", "", "Public Sub Yak()", "End Sub", "",
            "Private Sub Workbook_Open()", "", "End Sub", ""), text);
        Assert.Equal(7, header);
    }

    [Fact]
    public void AnEarlierNameGoesBeforeTheNextHandler()
    {
        var source = Text("Public Sub Beta()", "End Sub", "", "Public Sub Yak()", "End Sub", "",
            "Private Sub Workbook_Open()", "", "End Sub", "");

        var (text, header) = EventHandlers.Place(source, "Workbook", "Activate", "", Handlers);

        Assert.Equal(Text("Public Sub Beta()", "End Sub", "", "Public Sub Yak()", "End Sub", "",
            "Private Sub Workbook_Activate()", "", "End Sub", "",
            "Private Sub Workbook_Open()", "", "End Sub", ""), text);
        Assert.Equal(7, header);
    }

    [Fact]
    public void TheNewHandlerGoesAheadOfTheNextHandlersComments()
    {
        var source = Text("Option Explicit", "", "' Empty on purpose.", "",
            "Private Sub Worksheet_SelectionChange(ByVal Target As Range)", "", "End Sub", "");

        var (text, header) = EventHandlers.Place(source, "Worksheet", "Activate", "", Handlers);

        Assert.Equal(Text("Option Explicit", "", "Private Sub Worksheet_Activate()", "", "End Sub", "",
            "' Empty on purpose.", "",
            "Private Sub Worksheet_SelectionChange(ByVal Target As Range)", "", "End Sub", ""), text);
        Assert.Equal(3, header);
    }

    [Fact]
    public void OrdinaryProceduresDoNotPlaceIt()
    {
        // Zeta sorts after Worksheet_TableUpdate and Alpha stands last; neither is a handler, so
        // the new one goes past both, to the end.
        var source = Text("Option Explicit", "Public Sub Zeta()", "End Sub", "",
            "Private Sub Worksheet_SelectionChange(ByVal Target As Range)", "", "End Sub",
            "Public Sub Alpha()", "End Sub");

        var (text, header) = EventHandlers.Place(source, "Worksheet", "TableUpdate", "ByVal Target As TableObject", Handlers);

        Assert.EndsWith(Text("Public Sub Alpha()", "End Sub", "",
            "Private Sub Worksheet_TableUpdate(ByVal Target As TableObject)", "", "End Sub", ""), text);
        Assert.Equal(11, header);
    }

    [Fact]
    public void EveryObjectsHandlersAreOrderedByName()
    {
        // The class comparison's case: app_AfterCalculate goes before Class_Initialize, because
        // the VBE orders a new handler among all of them, not only its own object's.
        var source = Text("Private Sub Class_Initialize()", "", "End Sub", "");

        var (text, header) = EventHandlers.Place(source, "app", "AfterCalculate", "", Handlers);

        Assert.Equal(Text("Private Sub app_AfterCalculate()", "", "End Sub", "",
            "Private Sub Class_Initialize()", "", "End Sub", ""), text);
        Assert.Equal(1, header);
    }

    [Fact]
    public void ALaterNameThanEveryHandlerGoesAfterThem()
    {
        var source = Text("Private Sub app_WorkbookOpen(ByVal Wb As Workbook)", "", "End Sub", "");

        var (_, header) = EventHandlers.Place(source, "Class", "Initialize", "", Handlers);

        Assert.Equal(5, header);
    }

    [Fact]
    public void AnEmptyModuleTakesTheHandlerAlone()
    {
        Assert.Equal((Text("Private Sub Class_Initialize()", "", "End Sub", ""), 1),
            EventHandlers.Place(string.Empty, "Class", "Initialize", "", Handlers));
    }

    [Theory]
    [InlineData("Private Sub Worksheet_Change(ByVal Target As Range)")]
    [InlineData("Sub worksheet_change(ByVal Target As Range)")]
    [InlineData("  Public Static Sub Worksheet_Change (ByVal Target As Range)")]
    public void AStandingHandlerIsFoundInAnySpelling(string declaration)
    {
        Assert.Equal(1, EventHandlers.Find(Text("Option Explicit", declaration, "End Sub"), "Worksheet", "Change"));
    }

    [Fact]
    public void AHandlerOfALongerEventNameIsNotTheOne()
    {
        Assert.Equal(-1, EventHandlers.Find(Text("Private Sub Worksheet_ChangeSync()", "End Sub"), "Worksheet", "Change"));
    }
}
