using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The @Folder comment as Rubberduck documents it, read leniently and written one way. The
/// folder view is only ever a drawing of what these answer, so an answer that moved would move
/// a module across the tree.
/// </summary>
public class FolderAnnotationTests
{
    [Fact]
    public void TheDocumentedFormIsRead()
    {
        Assert.Equal("Accounts.Ledger", FolderAnnotation.Of("'@Folder(\"Accounts.Ledger\")\r\nOption Explicit\r\n"));
    }

    [Theory]
    [InlineData("'@Folder \"Accounts.Ledger\"")]
    [InlineData("'@Folder(Accounts.Ledger)")]
    [InlineData("'@Folder Accounts.Ledger")]
    [InlineData("'@folder(\"Accounts.Ledger\")")]
    [InlineData("'   @FOLDER   (  \"Accounts.Ledger\"  )")]
    [InlineData("    '@Folder(\"Accounts.Ledger\")")]
    [InlineData("'@Folder Accounts.Ledger this part is prose")]
    public void TheLooserSpellingsReadTheSame(string line)
    {
        Assert.Equal("Accounts.Ledger", FolderAnnotation.Of(line + "\r\nOption Explicit\r\n"));
    }

    [Fact]
    public void ItNeedNotBeTheFirstLine()
    {
        var source = "Option Explicit\r\n\r\n' The ledger.\r\n'@Folder(\"Accounts\")\r\n\r\nPrivate total As Long\r\n";

        Assert.Equal("Accounts", FolderAnnotation.Of(source));
    }

    [Fact]
    public void AnAnnotationBelowTheFirstProcedureIsAComment()
    {
        var source = "Option Explicit\r\n\r\nPublic Sub Go()\r\nEnd Sub\r\n\r\n'@Folder(\"Late\")\r\n";

        Assert.Null(FolderAnnotation.Of(source));
    }

    [Theory]
    [InlineData("Sub Go()")]
    [InlineData("Private Function F() As Long")]
    [InlineData("Public Static Sub Go()")]
    [InlineData("Friend Property Get P() As Long")]
    public void EveryProcedureHeaderEndsTheDeclarations(string header)
    {
        Assert.Null(FolderAnnotation.Of(header + "\r\n'@Folder(\"Late\")\r\n"));
    }

    [Fact]
    public void ADeclareIsNotAProcedure()
    {
        var source = "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\r\n'@Folder(\"Api\")\r\n";

        Assert.Equal("Api", FolderAnnotation.Of(source));
    }

    [Fact]
    public void TheFirstOfTwoWins()
    {
        var source = "'@Folder(\"First\")\r\n'@Folder(\"Second\")\r\n";

        Assert.Equal("First", FolderAnnotation.Of(source));
    }

    [Theory]
    [InlineData("'@Folder(\"\")")]
    [InlineData("'@Folder")]
    [InlineData("'@Folder(\" . . \")")]
    [InlineData("' @Folder-ish prose")]
    [InlineData("'@Folders(\"A\")")]
    [InlineData("Dim x As Long ' @Folder(\"A\")")]
    public void NothingNamedIsNoFolder(string line)
    {
        Assert.Null(FolderAnnotation.Of(line + "\r\n"));
    }

    [Theory]
    [InlineData(" Accounts . Ledger ", "Accounts.Ledger")]
    [InlineData("Accounts..Ledger", "Accounts.Ledger")]
    [InlineData(".Accounts.", "Accounts")]
    [InlineData("Accounts Payable.Q1", "Accounts Payable.Q1")]
    public void ThePathIsNormalizedSegmentBySegment(string raw, string kept)
    {
        Assert.Equal(kept, FolderAnnotation.Of($"'@Folder(\"{raw}\")\r\n"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Option Explicit")]
    public void NoTextIsNoFolder(string? source)
    {
        Assert.Null(FolderAnnotation.Of(source));
    }

    [Fact]
    public void EveryLineEndingIsALine()
    {
        Assert.Equal("A", FolderAnnotation.Of("Option Explicit\n'@Folder(\"A\")\nSub Go()\nEnd Sub"));
        Assert.Equal("A", FolderAnnotation.Of("Option Explicit\r'@Folder(\"A\")\rSub Go()"));
    }

    [Fact]
    public void TheCanonicalFormIsTheDocumentedOne()
    {
        Assert.Equal("'@Folder(\"Accounts.Ledger\")", FolderAnnotation.Canonical("Accounts.Ledger"));
        Assert.Collection(
            FolderAnnotation.Segments("Accounts.Ledger"),
            first => Assert.Equal("Accounts", first),
            second => Assert.Equal("Ledger", second));
    }
}
