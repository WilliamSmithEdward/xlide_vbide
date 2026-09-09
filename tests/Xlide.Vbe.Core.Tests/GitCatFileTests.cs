using System.Text;
using Xlide.Vbe.Core.Scm;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The `git cat-file --batch` parser. The header of a found object names the OBJECT and not
/// what was asked for, which is why the parser pairs answers with requests by position, and why
/// a missing name in the middle of a batch is the case worth pinning.
/// </summary>
public sealed class GitCatFileTests
{
    /// <summary>
    /// `HEAD:Books.bas`, `HEAD:Ledger.bas` (renamed away) and `HEAD:Account.cls` on stdin, git
    /// 2.55. The sizes are byte counts and the texts are LF, as the repository stores them.
    /// </summary>
    private const string Batch = """
        6b6453edc66dca662eb7dc3e28f596383664c0be blob 130
        Attribute VB_Name = "Books"
        Option Explicit

        Public Sub Post()
            Debug.Print "Ledger posts"
        End Sub

        Public Sub Close()
        End Sub

        HEAD:Ledger.bas missing
        62789feb1e1db826f9f20533c0e96ab6fc65f225 blob 257
        VERSION 1.0 CLASS
        BEGIN
          MultiUse = -1  'True
        END
        Attribute VB_Name = "Account"
        Attribute VB_GlobalNameSpace = False
        Attribute VB_Creatable = False
        Attribute VB_PredeclaredId = False
        Attribute VB_Exposed = False
        Option Explicit

        Public Balance As Currency


        """;

    private static readonly string[] Requested = ["HEAD:Books.bas", "HEAD:Ledger.bas", "HEAD:Account.cls"];

    private static byte[] Bytes(string text) =>
        Encoding.UTF8.GetBytes(text.Replace("\r\n", "\n", StringComparison.Ordinal));

    [Fact]
    public void EachAnswerLandsUnderTheNameThatAskedForIt()
    {
        var files = GitCatFile.ParseBatch(Bytes(Batch), Requested);

        Assert.Equal(2, files.Count);
        Assert.StartsWith(
            "Attribute VB_Name = \"Books\"\nOption Explicit\n", files["HEAD:Books.bas"], StringComparison.Ordinal);
        Assert.EndsWith("Public Sub Close()\nEnd Sub\n", files["HEAD:Books.bas"], StringComparison.Ordinal);
        Assert.Equal(130, Encoding.UTF8.GetByteCount(files["HEAD:Books.bas"]));

        // The missing name took its slot and nothing else: the third request still got the third
        // object, whole, to its declared size.
        Assert.False(files.ContainsKey("HEAD:Ledger.bas"));
        Assert.StartsWith("VERSION 1.0 CLASS\n", files["HEAD:Account.cls"], StringComparison.Ordinal);
        Assert.EndsWith("Public Balance As Currency\n", files["HEAD:Account.cls"], StringComparison.Ordinal);
        Assert.Equal(257, Encoding.UTF8.GetByteCount(files["HEAD:Account.cls"]));
    }

    [Fact]
    public void AByteOrderMarkIsNotPartOfTheText()
    {
        // Eight bytes declared: the mark and five letters. A file saved by an editor that writes
        // the mark compares equal to the module it came from.
        byte[] output =
        [
            .. Encoding.UTF8.GetBytes("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa blob 8\n"),
            0xEF, 0xBB, 0xBF,
            .. Encoding.UTF8.GetBytes("hello\n"),
        ];

        var files = GitCatFile.ParseBatch(output, ["HEAD:Marked.bas"]);

        Assert.Equal("hello", files["HEAD:Marked.bas"]);
    }

    [Fact]
    public void ANonAsciiModuleCountsBytesNotCharacters()
    {
        // "Größe" is seven bytes in UTF-8 and five characters; the size line says seven.
        var output = Encoding.UTF8.GetBytes("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb blob 7\nGröße\n");

        var files = GitCatFile.ParseBatch(output, ["HEAD:Größe.bas"]);

        Assert.Equal("Größe", files["HEAD:Größe.bas"]);
    }

    [Fact]
    public void AStreamCutShortStillYieldsWhatArrived()
    {
        var output = Encoding.UTF8.GetBytes("cccccccccccccccccccccccccccccccccccccccc blob 130\nOption Exp");

        var files = GitCatFile.ParseBatch(output, ["HEAD:Cut.bas"]);

        Assert.Equal("Option Exp", files["HEAD:Cut.bas"]);
    }

    [Fact]
    public void NothingAskedOrNothingAnsweredIsEmpty()
    {
        Assert.Empty(GitCatFile.ParseBatch([], Requested));
        Assert.Empty(GitCatFile.ParseBatch(Bytes(Batch), []));
    }
}
