using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The pane's rows: the live project against the branch head, and against the folder. Only code
/// is compared; a file's attribute header, its line endings and its trailing newline are the
/// export's business, and none of them may make an untouched module read as changed.
/// </summary>
public sealed class ScmRowsTests
{
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

    private const string FormHeader = """
        VERSION 5.00
        Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm
           Caption         =   "Entry"
           ClientHeight    =   3015
           OleObjectBlob   =   "EntryForm.frx":0000
        End
        Attribute VB_Name = "EntryForm"
        Attribute VB_GlobalNameSpace = False
        Attribute VB_Creatable = False
        Attribute VB_PredeclaredId = True
        Attribute VB_Exposed = False
        """;

    private const string LedgerCode =
        "Option Explicit\r\n\r\nPublic Sub Post()\r\n    Debug.Print \"Ledger\"\r\nEnd Sub";

    /// <summary>What the export writes for Ledger: header, body, CRLF, a final newline.</summary>
    private const string LedgerFile =
        "Attribute VB_Name = \"Ledger\"\r\nOption Explicit\r\n\r\nPublic Sub Post()\r\n"
        + "    Debug.Print \"Ledger\"\r\nEnd Sub\r\n";

    private static Dictionary<string, string> Files(params (string Name, string Text)[] files)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (name, text) in files)
        {
            result[name] = text;
        }

        return result;
    }

    [Fact]
    public void AModuleWhoseCodeMatchesItsFileIsNotARow()
    {
        var rows = ScmRows.Compute(
            [new ScmLiveModule("Ledger", "standard", LedgerCode)],
            Files(("Ledger.bas", LedgerFile)),
            []);

        Assert.Empty(rows);
    }

    [Fact]
    public void LineEndingsAndTheFilesTrailingNewlineChangeNothing()
    {
        // The file as another editor may have saved it: LF, and a blank line the code store
        // would not keep at the top.
        var file = "Attribute VB_Name = \"Ledger\"\n\nOption Explicit\n\nPublic Sub Post()\n"
            + "    Debug.Print \"Ledger\"\nEnd Sub\n";

        var rows = ScmRows.Compute(
            [new ScmLiveModule("Ledger", "standard", LedgerCode)],
            Files(("Ledger.bas", file)),
            []);

        Assert.Empty(rows);
    }

    [Fact]
    public void AClassWithAHeaderAtHeadIsComparedByItsCodeAlone()
    {
        var code = "Option Explicit\r\n\r\nPublic Balance As Currency";

        Assert.Empty(ScmRows.Compute(
            [new ScmLiveModule("Account", "class", code)],
            Files(("Account.cls", $"{ClassHeader}\r\n{code}\r\n")),
            []));

        var modified = Assert.Single(ScmRows.Compute(
            [new ScmLiveModule("Account", "class", code.Replace("Currency", "Double", StringComparison.Ordinal))],
            Files(("Account.cls", $"{ClassHeader}\r\n{code}\r\n")),
            []));
        Assert.Equal(new ScmRow("Account", "class", "Account.cls", "modified", null), modified);
    }

    [Fact]
    public void AFormsDesignerPreambleIsNotCode()
    {
        var code = "Option Explicit\r\n\r\nPrivate Sub OK_Click()\r\n    Me.Hide\r\nEnd Sub";

        Assert.Empty(ScmRows.Compute(
            [new ScmLiveModule("EntryForm", "userform", code)],
            Files(
                ("EntryForm.frm", $"{FormHeader}\r\n{code}\r\n"),
                ("EntryForm.frx", "\0\0binary"),
                ("EntryForm.form", "<form/>")),
            []));
    }

    [Fact]
    public void ANewModuleIsAddedAndAFileWithNoModuleIsDeleted()
    {
        var rows = ScmRows.Compute(
            [new ScmLiveModule("Reports", "standard", "Option Explicit")],
            Files(("Account.cls", $"{ClassHeader}\r\nOption Explicit\r\n"), ("Ledger.bas", LedgerFile)),
            []);

        Assert.Equal(
            [
                new ScmRow("Account", "class", "Account.cls", "deleted", null),
                new ScmRow("Ledger", "standard", "Ledger.bas", "deleted", null),
                new ScmRow("Reports", "standard", "Reports.bas", "added", null),
            ],
            rows);
    }

    [Fact]
    public void ARenameTheLogKnowsIsOneRowNotADeletionAndAnAddition()
    {
        var rows = ScmRows.Compute(
            [new ScmLiveModule("Books", "standard", LedgerCode)],
            Files(("Ledger.bas", LedgerFile)),
            [("Ledger", "Books")]);

        Assert.Equal([new ScmRow("Books", "standard", "Books.bas", "renamed", "Ledger")], rows);
    }

    [Fact]
    public void ARenameTheHeadDoesNotBearOutIsLeftAsItIs()
    {
        // The log says Ledger became Books, but the head has both files: nothing to pair.
        var rows = ScmRows.Compute(
            [
                new ScmLiveModule("Books", "standard", LedgerCode),
                new ScmLiveModule("Ledger", "standard", LedgerCode),
            ],
            Files(
                ("Ledger.bas", LedgerFile),
                ("Books.bas", LedgerFile.Replace("Ledger", "Books", StringComparison.Ordinal))),
            [("Ledger", "Books")]);

        Assert.Equal(["Books modified"], rows.Select(row => $"{row.Module} {row.Status}"));
    }

    [Fact]
    public void RowsAreSortedByModuleNameIgnoringCaseAndKeyedFilesMayCarryADirectory()
    {
        var rows = ScmRows.Compute(
            [
                new ScmLiveModule("zeta", "standard", "a"),
                new ScmLiveModule("Alpha", "standard", "a"),
                new ScmLiveModule("beta", "class", "a"),
            ],
            Files(("src/Alpha.bas", "Attribute VB_Name = \"Alpha\"\r\nb\r\n")),
            []);

        Assert.Equal(["Alpha", "beta", "zeta"], rows.Select(row => row.Module));
        Assert.Equal(["modified", "added", "added"], rows.Select(row => row.Status));
    }

    [Fact]
    public void TheFolderSectionNamesWhatDiffersWhatIsMissingAndWhatIsExtra()
    {
        var rows = ScmRows.Outside(
            [
                new ScmLiveModule("Ledger", "standard", LedgerCode),
                new ScmLiveModule("Reports", "standard", "Option Explicit"),
                new ScmLiveModule("Account", "class", "Option Explicit"),
            ],
            Files(
                ("Ledger.bas", LedgerFile.Replace(
                    "Debug.Print \"Ledger\"", "Debug.Print \"Ledger posts\"", StringComparison.Ordinal)),
                ("Account.cls", $"{ClassHeader}\r\nOption Explicit\r\n"),
                ("Stray.cls", $"{ClassHeader}\r\nOption Explicit\r\n"),
                ("Notes.txt", "not a module")));

        Assert.Equal(
            [
                new ScmRow("Ledger", "standard", "Ledger.bas", "folderNewer", null),
                new ScmRow("Reports", "standard", "Reports.bas", "missingInFolder", null),
                new ScmRow("Stray", "class", "Stray.cls", "missingInProject", null),
            ],
            rows);
    }

    [Fact]
    public void AFolderThatMatchesTheProjectHasNoOutsideRows()
    {
        Assert.Empty(ScmRows.Outside(
            [new ScmLiveModule("Ledger", "standard", LedgerCode)],
            Files(("Ledger.bas", LedgerFile), (".xlide-sync.lock", "xlide, process 1"))));
    }

    /// <summary>The companion editor writes a form as .cls (xlide_vscode#21); this side writes .frm.</summary>
    private const string FormAsClass = """
        VERSION 5.00
        Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EntryForm
           Caption         =   "Entry"
           OleObjectBlob   =   "EntryForm.frx":0000
        End
        Attribute VB_Name = "EntryForm"
        Attribute VB_GlobalNameSpace = False
        Option Explicit
        Private Sub OkButton_Click()
        End Sub

        """;

    [Fact]
    public void AFormTheCompanionWroteAsAClassFileIsItsOwnFile()
    {
        // A folder shared with the companion editor holds EntryForm.cls for the form. Read by
        // the name this side writes alone, that was the form "added" as EntryForm.frm beside its
        // own file "deleted" - two rows over a form nothing had touched.
        var live = new ScmLiveModule(
            "EntryForm", "userform", "Option Explicit\r\nPrivate Sub OkButton_Click()\r\nEnd Sub");

        Assert.Empty(ScmRows.Compute([live], Files(("EntryForm.cls", FormAsClass)), []));
        Assert.Empty(ScmRows.Outside([live], Files(("EntryForm.cls", FormAsClass))));

        // And a change is a change of THAT file.
        var edited = live with { Code = live.Code + "\r\n' typed since" };
        Assert.Equal(
            [new ScmRow("EntryForm", "userform", "EntryForm.cls", "modified", null)],
            ScmRows.Compute([edited], Files(("EntryForm.cls", FormAsClass)), []));
    }

    [Fact]
    public void AClassFileThatIsNotAFormDoesNotStandInForOne()
    {
        var live = new ScmLiveModule("EntryForm", "userform", "Option Explicit");

        // Two rows for one name: sorted by file within the module.
        Assert.Equal(
            [
                new ScmRow("EntryForm", "class", "EntryForm.cls", "deleted", null),
                new ScmRow("EntryForm", "userform", "EntryForm.frm", "added", null),
            ],
            ScmRows.Compute([live], Files(("EntryForm.cls", $"{ClassHeader}\r\nOption Explicit\r\n")), []));
    }
}
