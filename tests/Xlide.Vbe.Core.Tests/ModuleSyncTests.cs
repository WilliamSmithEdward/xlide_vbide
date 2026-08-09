using Xlide.Vbe.Core.Sync;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The half of import/export that needs no Excel: what a file is called, what kind of module it
/// holds, what changed, and what a plan would therefore do.
/// </summary>
public class ModuleSyncTests
{
    private const string DocumentHeader = """
        VERSION 1.0 CLASS
        BEGIN
          MultiUse = -1  'True
        END
        Attribute VB_Name = "Sheet1"
        Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
        Attribute VB_GlobalNameSpace = False
        Attribute VB_Creatable = False
        Attribute VB_PredeclaredId = True
        Attribute VB_Exposed = True
        """;

    private const string ClassHeader = """
        VERSION 1.0 CLASS
        BEGIN
          MultiUse = -1  'True
        END
        Attribute VB_Name = "Account"
        Attribute VB_GlobalNameSpace = False
        Attribute VB_Creatable = False
        Attribute VB_PredeclaredId = False
        Attribute VB_Exposed = False
        """;

    // A UserForm's VB_Base carries the type library AND the instance, which is the only reliable
    // way to tell one from a class without opening the designer.
    private const string UserFormHeader = """
        VERSION 5.00
        Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain
           Caption         =   "Main"
        End
        Attribute VB_Name = "FrmMain"
        Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}{9A2B7C31-0000-0000-C000-000000000046}"
        Attribute VB_Exposed = False
        """;

    [Fact]
    public void AStandardModuleIsABasAndEverythingElseIsACls()
    {
        Assert.Equal("Helpers.bas", ModuleSync.FileNameFor("Helpers", "standard"));
        Assert.Equal("Account.cls", ModuleSync.FileNameFor("Account", "class"));
        Assert.Equal("Sheet1.cls", ModuleSync.FileNameFor("Sheet1", "document"));
        Assert.Equal("FrmMain.cls", ModuleSync.FileNameFor("FrmMain", "userform"));
    }

    [Fact]
    public void ANameAFileCannotCarryIsMadeIntoOneThatItCan()
    {
        Assert.Equal("a_b.bas", ModuleSync.FileNameFor("a/b", "standard"));
        Assert.Equal("q_.bas", ModuleSync.FileNameFor("q?", "standard"));
    }

    [Fact]
    public void ATrailingDotOrSpaceIsDroppedBecauseWindowsWouldDropItSilently()
    {
        // Left on, the file would be written as "Odd" and looked for as "Odd. ", so an export would
        // report success and the next plan would say the file is missing, for ever.
        Assert.Equal("Odd.bas", ModuleSync.FileNameFor("Odd. ", "standard"));
        Assert.Equal("Odd.cls", ModuleSync.FileNameFor("Odd.", "class"));
    }

    [Fact]
    public void ABasIsAlwaysStandardWhateverItsHeaderSays()
    {
        Assert.Equal("standard", ModuleSync.ClassifyFile("Helpers.bas", DocumentHeader));
    }

    [Fact]
    public void AClsWithADocumentClassIdIsADocument()
    {
        Assert.Equal("document", ModuleSync.ClassifyFile("Sheet1.cls", DocumentHeader));
    }

    [Fact]
    public void AClsWithTwoGuidsIsAUserForm()
    {
        Assert.Equal("userform", ModuleSync.ClassifyFile("FrmMain.cls", UserFormHeader));
    }

    [Fact]
    public void AnOrdinaryClsIsAClass()
    {
        Assert.Equal("class", ModuleSync.ClassifyFile("Account.cls", ClassHeader));
    }

    [Fact]
    public void APredeclaredClassIsStillAClass()
    {
        // VB_PredeclaredId is set by document modules AND by any class written in the singleton
        // style. Reading it as proof of a document would refuse to create those, which is how a
        // whole family of library classes stops importing.
        var singleton = ClassHeader.Replace(
            "Attribute VB_PredeclaredId = False",
            "Attribute VB_PredeclaredId = True",
            StringComparison.Ordinal);
        Assert.Equal("class", ModuleSync.ClassifyFile("StdArray.cls", singleton));
    }

    [Theory]
    [InlineData("ThisWorkbook")]
    [InlineData("Sheet3")]
    [InlineData("Feuil2")]
    [InlineData("Tabelle1")]
    [InlineData("Hoja1")]
    public void AWellKnownDocumentNameIsADocumentWhenTheHeaderIsNoHelp(string moduleName)
    {
        Assert.Equal("document", ModuleSync.ClassifyFile($"{moduleName}.cls", "Option Explicit"));
    }

    [Fact]
    public void TheShownComparisonHasNoHeaderInIt()
    {
        var source = $"{DocumentHeader}\r\nOption Explicit\r\n\r\nSub Go()\r\nEnd Sub\r\n";
        var code = ModuleSync.CodeWithoutHeader(source);

        Assert.StartsWith("Option Explicit", code, StringComparison.Ordinal);
        Assert.DoesNotContain("Attribute VB_Name", code, StringComparison.Ordinal);
        Assert.DoesNotContain("VERSION", code, StringComparison.Ordinal);
        Assert.Contains("Sub Go()", code, StringComparison.Ordinal);
    }

    [Fact]
    public void AProcedureAttributeIsHiddenWhereverItSits()
    {
        var source = "Sub Go()\r\nAttribute Go.VB_Description = \"does it\"\r\n    Beep\r\nEnd Sub\r\n";
        var code = ModuleSync.CodeWithoutHeader(source);

        Assert.DoesNotContain("VB_Description", code, StringComparison.Ordinal);
        Assert.Contains("    Beep", code, StringComparison.Ordinal);
    }

    [Fact]
    public void LineEndingsAreNotADifference()
    {
        Assert.True(ModuleSync.SameText("a\r\nb", "a\nb"));
        Assert.False(ModuleSync.SameText("a\r\nb", "a\r\nc"));
    }

    [Fact]
    public void IdenticalTextComparesAsAllEqual()
    {
        var diff = ModuleSync.Diff("one\ntwo\nthree", "one\ntwo\nthree");
        Assert.All(diff, line => Assert.Equal(DiffKind.Equal, line.Kind));
        Assert.Equal(3, diff.Count);
    }

    [Fact]
    public void ARewrittenLineReadsAsChangedRatherThanAsTwoEdits()
    {
        var diff = ModuleSync.Diff("one\ntwo\nthree", "one\nTWO\nthree");
        var middle = Assert.Single(diff, line => line.Kind != DiffKind.Equal);

        Assert.Equal(DiffKind.Changed, middle.Kind);
        Assert.Equal("two", middle.Left);
        Assert.Equal("TWO", middle.Right);
        Assert.Equal(2, middle.LeftNumber);
        Assert.Equal(2, middle.RightNumber);
    }

    [Fact]
    public void ALineOnlyOnOneSideCarriesOnlyThatSidesNumber()
    {
        var added = ModuleSync.Diff("one\nthree", "one\ntwo\nthree")
            .Single(line => line.Kind == DiffKind.Added);
        Assert.Null(added.LeftNumber);
        Assert.Equal(2, added.RightNumber);
        Assert.Equal("two", added.Right);

        var removed = ModuleSync.Diff("one\ntwo\nthree", "one\nthree")
            .Single(line => line.Kind == DiffKind.Removed);
        Assert.Equal(2, removed.LeftNumber);
        Assert.Null(removed.RightNumber);
        Assert.Equal("two", removed.Left);
    }

    [Fact]
    public void ExportSeesTheThreeCasesApart()
    {
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [
                new LiveModule("Same", "standard", "Option Explicit"),
                new LiveModule("Different", "standard", "Option Explicit\r\nSub A()\r\nEnd Sub"),
                new LiveModule("Fresh", "class", "Option Explicit"),
            ],
            [
                new RepoFile("Same.bas", "Option Explicit"),
                new RepoFile("Different.bas", "Option Explicit"),
            ],
            ExportMode.ExportAll);

        Assert.Equal(SyncStatus.Unchanged, plan.Items.Single(i => i.ModuleName == "Same").Status);
        Assert.Equal(SyncStatus.WillWrite, plan.Items.Single(i => i.ModuleName == "Different").Status);
        Assert.Equal(SyncStatus.WillCreate, plan.Items.Single(i => i.ModuleName == "Fresh").Status);
        Assert.Equal("Fresh.cls", plan.Items.Single(i => i.ModuleName == "Fresh").FileName);
    }

    [Fact]
    public void AnUnchangedRowIsOfferedButNotTicked()
    {
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [new LiveModule("Same", "standard", "Option Explicit")],
            [new RepoFile("Same.bas", "Option Explicit")],
            ExportMode.ExportAll);

        var item = Assert.Single(plan.Items);
        Assert.False(item.Checked);
    }

    [Fact]
    public void ExportOnlyOffersToDeleteWhenAskedToMakeTheFolderMatch()
    {
        LiveModule[] modules = [new LiveModule("Kept", "standard", "Option Explicit")];
        RepoFile[] files = [new RepoFile("Kept.bas", "Option Explicit"), new RepoFile("Gone.bas", "Option Explicit")];

        var leaveAlone = ModuleSync.PlanExport("p", "Book.xlsm", @"C:\out", modules, files, ExportMode.ExportAll);
        Assert.DoesNotContain(leaveAlone.Items, i => i.Status == SyncStatus.WillRemove);

        var trueUp = ModuleSync.PlanExport("p", "Book.xlsm", @"C:\out", modules, files, ExportMode.TrueUp);
        var stale = Assert.Single(trueUp.Items, i => i.Status == SyncStatus.WillRemove);
        Assert.Equal("Gone.bas", stale.FileName);
        Assert.NotNull(stale.Warning);
    }

    [Fact]
    public void AFileThatIsNotAModuleIsNeverStale()
    {
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [],
            [new RepoFile("README.md", "notes"), new RepoFile("Gone.bas", "Option Explicit")],
            ExportMode.TrueUp);

        var stale = Assert.Single(plan.Items);
        Assert.Equal("Gone.bas", stale.FileName);
    }

    [Fact]
    public void ImportUpdatesWhatExistsAndCreatesWhatDoesNot()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [new LiveModule("Existing", "standard", "Option Explicit")],
            [
                new RepoFile("Existing.bas", "Option Explicit\r\nSub A()\r\nEnd Sub"),
                new RepoFile("Brand.bas", "Option Explicit"),
            ],
            ImportMode.UpdateOnly);

        Assert.Equal(SyncStatus.WillUpdate, plan.Items.Single(i => i.ModuleName == "Existing").Status);
        Assert.Equal(SyncStatus.WillCreate, plan.Items.Single(i => i.ModuleName == "Brand").Status);
    }

    [Fact]
    public void ADocumentThatIsNotInTheProjectIsSkippedRatherThanAttempted()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [],
            [new RepoFile("Sheet1.cls", DocumentHeader)],
            ImportMode.UpdateOnly);

        var item = Assert.Single(plan.Items);
        Assert.Equal(SyncStatus.SkippingImport, item.Status);
        Assert.True(item.CannotBeCreated);
        Assert.False(item.Checked);
        Assert.NotNull(item.Warning);
        Assert.Single(plan.Warnings);
    }

    [Fact]
    public void ADocumentThatIsInTheProjectIsUpdatedNormally()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [new LiveModule("Sheet1", "document", "Option Explicit")],
            [new RepoFile("Sheet1.cls", $"{DocumentHeader}\r\nOption Explicit\r\nSub A()\r\nEnd Sub")],
            ImportMode.UpdateOnly);

        var item = Assert.Single(plan.Items);
        Assert.Equal(SyncStatus.WillUpdate, item.Status);
        Assert.False(item.CannotBeCreated);
        Assert.True(item.Checked);
    }

    [Fact]
    public void ImportTrueUpNeverOffersToDeleteADocumentOrAForm()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [
                new LiveModule("Orphan", "standard", "Option Explicit"),
                new LiveModule("ThisWorkbook", "document", "Option Explicit"),
                new LiveModule("FrmMain", "userform", "Option Explicit"),
            ],
            [],
            ImportMode.TrueUpStandardClass);

        var removals = plan.Items.Where(i => i.Status == SyncStatus.WillRemove).ToList();
        var only = Assert.Single(removals);
        Assert.Equal("Orphan", only.ModuleName);
    }

    [Fact]
    public void AFileThatWouldNotReadIsReportedRatherThanTreatedAsEmpty()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [new LiveModule("Locked", "standard", "Option Explicit")],
            [new RepoFile("Locked.bas", string.Empty, "the file is in use by another program")],
            ImportMode.UpdateOnly);

        var item = Assert.Single(plan.Items);
        Assert.Equal(SyncStatus.ReadError, item.Status);
        Assert.False(item.Checked);
        Assert.Equal("the file is in use by another program", item.Warning);
    }

    [Fact]
    public void RowsThatDoSomethingComeFirst()
    {
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [
                new LiveModule("Zebra", "standard", "changed"),
                new LiveModule("Alpha", "standard", "same"),
            ],
            [new RepoFile("Alpha.bas", "same"), new RepoFile("Zebra.bas", "was")],
            ExportMode.ExportAll);

        Assert.Equal("Zebra", plan.Items[0].ModuleName);
        Assert.Equal("Alpha", plan.Items[1].ModuleName);
    }

    [Fact]
    public void ModesAreReadLenientlyAndDefaultToTheSafeOne()
    {
        Assert.Equal(ExportMode.TrueUp, ModuleSync.ExportModeFrom("trueUp"));
        Assert.Equal(ExportMode.TrueUp, ModuleSync.ExportModeFrom("TRUE-UP"));
        Assert.Equal(ExportMode.ExportAll, ModuleSync.ExportModeFrom("exportAll"));
        Assert.Equal(ExportMode.ExportAll, ModuleSync.ExportModeFrom(null));
        Assert.Equal(ExportMode.ExportAll, ModuleSync.ExportModeFrom("nonsense"));

        Assert.Equal(ImportMode.TrueUpStandardClass, ModuleSync.ImportModeFrom("trueUpStandardClass"));
        Assert.Equal(ImportMode.UpdateOnly, ModuleSync.ImportModeFrom(null));
        Assert.Equal(ImportMode.UpdateOnly, ModuleSync.ImportModeFrom("nonsense"));
    }

    [Fact]
    public void APayloadIsTheRawTextSoWhatIsWrittenRoundTrips()
    {
        var full = $"{ClassHeader}\r\nOption Explicit\r\n";
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [new LiveModule("Account", "class", full)],
            [],
            ExportMode.ExportAll);

        // The shown comparison hides the header; the payload must not, or the exported file would
        // lose the attributes and come back as a different kind of module.
        Assert.Equal(full, Assert.Single(plan.Items).PayloadSource);
    }
}
