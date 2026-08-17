using Xlide.Vbe.Core.Sync;
using System.Linq;
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
    public void AStandardModuleIsABasAFormIsAFrmAndEverythingElseIsACls()
    {
        Assert.Equal("Helpers.bas", ModuleSync.FileNameFor("Helpers", "standard"));
        Assert.Equal("Account.cls", ModuleSync.FileNameFor("Account", "class"));
        Assert.Equal("Sheet1.cls", ModuleSync.FileNameFor("Sheet1", "document"));

        // A form is a .frm from 2026-08-16: the VBE's own exporter writes one, and only that file
        // names the binary sidecar its controls live in. As a .cls it named a temporary path.
        Assert.Equal("FrmMain.frm", ModuleSync.FileNameFor("FrmMain", "userform"));
        Assert.Equal("FrmMain.form", ModuleSync.DesignFileNameFor("FrmMain"));
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
    public void AFormsDesignerBlockIsNotShownAsCode()
    {
        // A UserForm's file opens with its designer: a BEGIN block of Caption, ClientHeight,
        // StartUpPosition and the rest. None of it is code and none of it is editable here, and it
        // used to be drawn as the first screenful of a form's comparison, because the predicate
        // that hides it and the predicate the shim reads headers with were written separately and
        // only one of them knew about designer properties.
        const string source = """
            VERSION 5.00
            Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain
               Caption         =   "Main"
               ClientHeight    =   3000
               StartUpPosition =   1
            End
            Attribute VB_Name = "FrmMain"
            Option Explicit

            Private Sub UserForm_Initialize()
            End Sub
            """;

        var code = ModuleSync.CodeWithoutHeader(source);

        Assert.StartsWith("Option Explicit", code, StringComparison.Ordinal);
        Assert.DoesNotContain("Caption", code, StringComparison.Ordinal);
        Assert.DoesNotContain("StartUpPosition", code, StringComparison.Ordinal);
        Assert.Contains("UserForm_Initialize", code, StringComparison.Ordinal);
    }

    [Fact]
    public void CodeThatMerelyBeginsLikeADesignerLineIsStillCode()
    {
        // The designer words mean something only before the first line of code. A procedure that
        // assigns a Caption is a procedure.
        const string source = """
            Option Explicit

            Sub Go()
                Caption = "hello"
            End Sub
            """;

        var code = ModuleSync.CodeWithoutHeader(source);

        Assert.Contains("    Caption = \"hello\"", code, StringComparison.Ordinal);
    }

    [Fact]
    public void TwoIdenticalFilesCondenseToASingleGap()
    {
        const string four = """
            one
            two
            three
            four
            """;

        var only = Assert.Single(ModuleSync.Condense(ModuleSync.Diff(four, four)));

        Assert.Equal(DiffKind.Gap, only.Kind);
        Assert.Contains("4 identical lines", only.Left, StringComparison.Ordinal);
    }

    [Fact]
    public void AChangeKeepsItsSurroundingsAndDropsTheRest()
    {
        // Two hundred lines differing in one of them: a comparison of the change and its
        // neighbours, not of the file.
        var left = string.Join("\n", Enumerable.Range(1, 200).Select(n => $"line {n}"));
        var right = left.Replace("line 100", "LINE 100", StringComparison.Ordinal);

        var condensed = ModuleSync.Condense(ModuleSync.Diff(left, right), context: 3);

        Assert.Single(condensed, line => line.Kind == DiffKind.Changed);
        Assert.Contains(condensed, line => line.Left == "line 97");
        Assert.Contains(condensed, line => line.Left == "line 103");
        Assert.DoesNotContain(condensed, line => line.Left == "line 50");
        Assert.Equal(2, condensed.Count(line => line.Kind == DiffKind.Gap));
        Assert.True(condensed.Count < 20, $"condensed to {condensed.Count} lines, which is not condensed");
    }

    [Fact]
    public void ASmallComparisonIsLeftWhole()
    {
        // Nothing to gain by breaking up a comparison shorter than the context either side.
        var condensed = ModuleSync.Condense(
            ModuleSync.Diff("one\ntwo", "one\nTWO"),
            context: 3);

        Assert.DoesNotContain(condensed, line => line.Kind == DiffKind.Gap);
        Assert.Equal(2, condensed.Count);
    }

    [Fact]
    public void AComparisonWithNothingInCommonIsCappedAndSaysSo()
    {
        // The FIRST export: every module against a file that is not there, so every line is a
        // change and there is no agreement to condense. This is the case that made a plan 15MB.
        var whole = string.Join("\n", Enumerable.Range(1, 5000).Select(n => $"line {n}"));

        var condensed = ModuleSync.Condense(ModuleSync.Diff(whole, string.Empty), most: 400);

        Assert.Equal(401, condensed.Count);
        var last = condensed[^1];
        Assert.Equal(DiffKind.Gap, last.Kind);
        Assert.Contains("4,600 not shown lines", last.Left, StringComparison.Ordinal);
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

    /*
     * A FORM'S DESIGN IS A FILE OF ITS OWN.
     *
     * A UserForm exported as code alone cannot be put back - Excel writes its controls into a
     * binary .frx - so a form in source control was half a form. The markup is the design as text,
     * and it rides as a second row so the developer can see it, tick it and diff it.
     */
    private const string FormMarkup = "Form Entry \"Quarter\" size 360x320\r\n"
        + "    CommandButton OkButton \"Start\" at 262,250 size 72x24\r\n";

    private const string FormCode = "Option Explicit\r\n";

    [Fact]
    public void AFormExportsItsDesignBesideItsCode()
    {
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [new LiveModule("Entry", "userform", $"{ClassHeader}\r\n{FormCode}", FormMarkup)],
            [],
            ExportMode.ExportAll);

        var code = Assert.Single(plan.Items, item => item.FileName == "Entry.frm");
        var design = Assert.Single(plan.Items, item => item.FileName == "Entry.form");

        Assert.False(code.IsDesign);
        Assert.True(design.IsDesign);
        Assert.Equal(SyncStatus.WillCreate, design.Status);
        Assert.Equal(FormMarkup, design.PayloadSource);
    }

    [Fact]
    public void ADesignAlreadyOnDiskIsUnchangedRatherThanRewritten()
    {
        var plan = ModuleSync.PlanExport(
            "p", "Book.xlsm", @"C:\out",
            [new LiveModule("Entry", "userform", $"{ClassHeader}\r\n{FormCode}", FormMarkup)],
            [new RepoFile("Entry.form", FormMarkup)],
            ExportMode.ExportAll);

        Assert.Equal(SyncStatus.Unchanged, Assert.Single(plan.Items, one => one.IsDesign).Status);
    }

    [Fact]
    public void ADesignFileImportsAsAnApplyToTheFormItNames()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\out",
            [new LiveModule("Entry", "userform", $"{ClassHeader}\r\n{FormCode}", "Form Entry\r\n")],
            [new RepoFile("Entry.form", FormMarkup)],
            ImportMode.UpdateOnly);

        var design = Assert.Single(plan.Items, one => one.IsDesign);
        Assert.Equal(SyncStatus.WillUpdate, design.Status);
        Assert.Equal("Entry", design.ModuleName);
        Assert.Equal(FormMarkup, design.PayloadSource);
    }

    [Fact]
    public void ADesignWhoseFormIsMissingIsRefusedRatherThanGuessedAt()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\out",
            [],
            [new RepoFile("Entry.form", FormMarkup)],
            ImportMode.UpdateOnly);

        var design = Assert.Single(plan.Items, one => one.IsDesign);
        Assert.Equal(SyncStatus.SkippingImport, design.Status);
        Assert.True(design.CannotBeCreated);
        Assert.Contains("Add the form first", design.Warning);
    }

    /*
     * IMPORT CREATES A FORM, 2026-08-16. The pair the VBE's own exporter writes is what its
     * importer reads, so the refusal that stood here for months - "a UserForm's designer is not
     * in this file" - was only ever true of the code file alone.
     */
    private const string ExportedForm = """
        VERSION 5.00
        Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain
           Caption         =   "Main"
           OleObjectBlob   =   "FrmMain.frx":0000
        End
        Attribute VB_Name = "FrmMain"
        Attribute VB_GlobalNameSpace = False
        Option Explicit
        """;

    [Fact]
    public void AFormArrivesWithItsSidecarAndIsCreatedRatherThanSkipped()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [],
            [new RepoFile("FrmMain.frm", ExportedForm), new RepoFile("FrmMain.frx", string.Empty)],
            ImportMode.UpdateOnly);

        var item = Assert.Single(plan.Items);
        Assert.Equal(SyncStatus.WillCreate, item.Status);
        Assert.False(item.CannotBeCreated);
        Assert.True(item.Checked);
        Assert.Contains("binary sidecar", item.Warning);
        Assert.Empty(plan.Warnings);
    }

    [Fact]
    public void TheSharedPlannersClsNamingStillReadsAsAForm()
    {
        // The companion editor writes a form's code as `.cls` (xlide_vscode#21) where this product
        // writes `.frm`. Same bytes, two names - so the TEXT decides, not the extension.
        Assert.Equal("userform", ModuleSync.ClassifyFile("FrmMain.cls", ExportedForm));

        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [],
            [new RepoFile("FrmMain.cls", ExportedForm), new RepoFile("FrmMain.frx", string.Empty)],
            ImportMode.UpdateOnly);

        Assert.Equal(SyncStatus.WillCreate, Assert.Single(plan.Items).Status);
    }

    [Fact]
    public void AFormWithoutItsSidecarIsRefusedByTheNameOfTheFileItWants()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [],
            [new RepoFile("FrmMain.frm", ExportedForm)],
            ImportMode.UpdateOnly);

        var item = Assert.Single(plan.Items);
        Assert.Equal(SyncStatus.SkippingImport, item.Status);
        Assert.True(item.CannotBeCreated);
        Assert.Contains("FrmMain.frx", item.Warning);
        Assert.Contains("FrmMain.frx", Assert.Single(plan.Warnings));
    }

    [Fact]
    public void ADesignBesideAnArrivingFormIsOfferedUntickedBecauseTheBinaryIsAuthoritative()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [],
            [
                new RepoFile("FrmMain.frm", ExportedForm),
                new RepoFile("FrmMain.frx", string.Empty),
                new RepoFile("FrmMain.form", FormMarkup),
            ],
            ImportMode.UpdateOnly);

        var design = Assert.Single(plan.Items, one => one.IsDesign);
        Assert.Equal(SyncStatus.WillUpdate, design.Status);

        // The guard: the sidecar carries every control and the markup's list is TOTAL, so
        // applying the text on top can only take things away when the two disagree.
        Assert.False(design.Checked);
        Assert.Contains("would be removed", design.Warning);
    }

    [Fact]
    public void ASidecarIsNeverARowOfItsOwn()
    {
        var plan = ModuleSync.PlanImport(
            "p", "Book.xlsm", @"C:\in",
            [],
            [new RepoFile("FrmMain.frm", ExportedForm), new RepoFile("FrmMain.frx", string.Empty)],
            ImportMode.UpdateOnly);

        Assert.DoesNotContain(plan.Items, one => one.FileName.EndsWith(".frx"));
    }
}
