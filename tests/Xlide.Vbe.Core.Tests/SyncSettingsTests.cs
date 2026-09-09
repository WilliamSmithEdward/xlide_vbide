using Xlide.Vbe.Core.Sync;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// sync.json: a file written before the Repository field existed must read as "not under source
/// control", never as null, and the field must survive every way the settings are copied.
/// </summary>
public sealed class SyncSettingsTests
{
    private const string ThreeKeyFile = """
        {
          "Projects": {
            "C:\\Books\\Ledger.xlsm": {
              "Folder": "C:\\Books\\src",
              "ExportMode": "TrueUp",
              "ImportMode": "updateOnly"
            }
          }
        }
        """;

    [Fact]
    public void AFileWrittenBeforeTheFieldExistedReadsRepositoryAsEmpty()
    {
        var choice = SyncSettings.Parse(ThreeKeyFile).For(@"C:\Books\Ledger.xlsm");

        Assert.Equal(string.Empty, choice.Repository);
        Assert.Equal(@"C:\Books\src", choice.Folder);
        Assert.Equal("TrueUp", choice.ExportMode);
        Assert.Equal("updateOnly", choice.ImportMode);
    }

    [Fact]
    public void AProjectTheFileDoesNotKnowIsNotUnderSourceControl()
    {
        Assert.Equal(string.Empty, SyncSettings.Parse(ThreeKeyFile).For(@"C:\Other.xlsm").Repository);
        Assert.Equal(string.Empty, SyncSettings.Empty.For("anything").Repository);
    }

    [Fact]
    public void WithAndForRoundTripTheRepository()
    {
        var was = SyncSettings.Parse(ThreeKeyFile);
        var remembered = was.For(@"C:\Books\Ledger.xlsm");

        var now = was.With(@"C:\Books\Ledger.xlsm", remembered with { Repository = @"C:\Books\repo" });

        var choice = now.For(@"C:\Books\Ledger.xlsm");
        Assert.Equal(@"C:\Books\repo", choice.Repository);
        Assert.Equal(@"C:\Books\src", choice.Folder);

        // The original is untouched: With answers new settings.
        Assert.Equal(string.Empty, was.For(@"C:\Books\Ledger.xlsm").Repository);
    }

    [Fact]
    public void TheRepositorySurvivesTheFileText()
    {
        var written = SyncSettings.Empty
            .With("p", new SyncChoice { Folder = @"C:\f", Repository = @"C:\r" })
            .ToJson();

        var read = SyncSettings.Parse(written).For("p");
        Assert.Equal(@"C:\r", read.Repository);
        Assert.Equal(@"C:\f", read.Folder);
        Assert.Contains("\"Repository\"", written, StringComparison.Ordinal);
    }
}
