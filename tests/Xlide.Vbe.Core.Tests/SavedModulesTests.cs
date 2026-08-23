using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The attribute header of every module, read out of the saved document.
///
/// This is the only way the product can learn `VB_PredeclaredId`, which decides whether a class
/// module's own name is a value or a type - and therefore whether `Ticket.ChangeTest` is correct
/// code or `Variable not defined` (xlide_vscode#47). A code pane never shows the attribute and
/// the object model has no property for it.
///
/// Two halves, like <see cref="SavedDesignTests"/>. The TOLERANCE half needs nothing and always
/// runs, because the promise made to the analyzer is that an unreadable document answers UNKNOWN
/// rather than throwing or guessing: a wrong `false` here is a red squiggle under working code,
/// on every use of every predeclared singleton in the project. The AGREEMENT half runs against
/// the generated fixture and pins what the reader was built against, which was measured first in
/// the harness (tools\harness\vba-storage.mjs) against the same file.
/// </summary>
public class SavedModulesTests
{
    /// <summary>The fixture, or null on a checkout that has not generated one.</summary>
    private static string? Fixture()
    {
        var here = AppContext.BaseDirectory;
        for (var up = 0; up < 8 && here is not null; up++)
        {
            if (File.Exists(Path.Combine(here, "xlide_vbide.slnx")))
            {
                var fixture = Path.Combine(here, "artifacts", "fixtures", "LanguageFixture.xlsm");
                return File.Exists(fixture) ? fixture : null;
            }

            here = Path.GetDirectoryName(here);
        }

        return null;
    }

    /* ---- tolerance: never throws, never guesses -------------------------------------------- */

    [Fact]
    public void ADocumentThatIsNotThereSaysNothing()
    {
        Assert.Null(SavedModules.For(null));
        Assert.Null(SavedModules.For(string.Empty));
        Assert.Null(SavedModules.For("   "));
        Assert.Null(SavedModules.For(Path.Combine(Path.GetTempPath(), "no-such-workbook.xlsm")));
    }

    [Fact]
    public void AFileThatIsNotAPackageSaysNothing()
    {
        var text = Path.Combine(Path.GetTempPath(), $"xlide-not-a-workbook-{Guid.NewGuid():N}.xlsm");
        File.WriteAllText(text, "This is not a package, and saying so must not throw.");
        try
        {
            Assert.Null(SavedModules.For(text));
        }
        finally
        {
            File.Delete(text);
        }
    }

    [Fact]
    public void BytesThatAreNotAContainerDecompressToNothing()
    {
        Assert.Null(VbaCompression.Decompress([], 0, 4096));
        Assert.Null(VbaCompression.Decompress([0x00, 0x01, 0x02], 0, 4096));
        Assert.Null(VbaCompression.Decompress([0x01], 5, 4096));

        // A caller asking for nothing gets nothing rather than an empty walk.
        Assert.Null(VbaCompression.Decompress([0x01, 0x00, 0xB0], 0, 0));
    }

    /* ---- agreement: the fixture's own modules ---------------------------------------------- */

    [Fact]
    public void APlainClassModuleIsNotPredeclared()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        var saved = SavedModules.For(fixture);
        Assert.NotNull(saved);

        // Gadget is a class module the fixture builder adds through the object model, so it is
        // born with the attribute set to False - and False is the ONLY state that reports.
        Assert.False(saved!.PredeclaredIdOf("Gadget"));
        Assert.True(saved.Knows("Gadget"));
    }

    [Fact]
    public void ADocumentModuleIsPredeclared()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        var saved = SavedModules.For(fixture);
        Assert.NotNull(saved);

        // Sheet and workbook modules always carry it, which is what makes their bare names usable
        // as values. The analyzer knows that from the module KIND, so this is not what the flag is
        // read for - but it is proof the reader is reading the real attribute and not a default.
        Assert.True(saved!.PredeclaredIdOf("ThisWorkbook"));
        Assert.True(saved.PredeclaredIdOf("Sheet1"));

        // A form has one too.
        Assert.True(saved.PredeclaredIdOf("EntryForm"));
    }

    [Fact]
    public void AStandardModuleCarriesNoSuchAttributeAndSaysUnknown()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        var saved = SavedModules.For(fixture);
        Assert.NotNull(saved);

        // A .bas header is one line, `Attribute VB_Name`. The module is KNOWN and the flag is not
        // there, and those are two different facts: unknown, not false.
        Assert.True(saved!.Knows("Shapes"));
        Assert.Null(saved.PredeclaredIdOf("Shapes"));
    }

    [Fact]
    public void AModuleTheSavedFileDoesNotCarryIsUnknown()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        var saved = SavedModules.For(fixture);
        Assert.NotNull(saved);

        // Which is every module added since the last save. Silence is the only safe answer.
        Assert.False(saved!.Knows("AddedSinceTheLastSave"));
        Assert.Null(saved.PredeclaredIdOf("AddedSinceTheLastSave"));
    }

    [Fact]
    public void AModuleTheImportReplacedStopsBeingAnswered()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        // A COPY, so the doubt raised here is keyed to a path no other test reads. Doubt lives
        // until the document is saved again, which is exactly the point of it.
        var copy = Path.Combine(Path.GetTempPath(), $"xlide-doubt-{Guid.NewGuid():N}.xlsm");
        File.Copy(fixture!, copy);
        try
        {
            var saved = SavedModules.For(copy);
            Assert.NotNull(saved);
            Assert.False(saved!.PredeclaredIdOf("Gadget"));

            // The import removes a class and puts another of the same name in its place. The file
            // still describes the OLD one, and the flag is the one thing about a class that the
            // editor will not let a developer change - so the file's answer is now about a module
            // that is gone, and answering it would be the one way this can be actively wrong.
            SavedModules.Doubt(copy, "Gadget");

            Assert.Null(saved.PredeclaredIdOf("Gadget"));
            Assert.False(saved.Knows("Gadget"));

            // And only that module. Doubt is per name, not a blanket over the document.
            Assert.True(saved.PredeclaredIdOf("ThisWorkbook"));
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void SavingTheDocumentSettlesTheDoubt()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        var copy = Path.Combine(Path.GetTempPath(), $"xlide-settle-{Guid.NewGuid():N}.xlsm");
        File.Copy(fixture!, copy);
        try
        {
            Assert.False(SavedModules.For(copy)!.PredeclaredIdOf("Gadget"));
            SavedModules.Doubt(copy, "Gadget");
            Assert.Null(SavedModules.For(copy)!.PredeclaredIdOf("Gadget"));

            // A SAVE is what makes the file describe the project again, and the file's own write
            // time is how that is noticed - nothing has to be told. Without this the doubt would
            // be permanent for the session, and a class would go unanswered for ever after one
            // import.
            File.SetLastWriteTimeUtc(copy, File.GetLastWriteTimeUtc(copy).AddSeconds(5));

            Assert.False(SavedModules.For(copy)!.PredeclaredIdOf("Gadget"));
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void NamesAreMatchedTheWayTheLanguageMatchesThem()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        var saved = SavedModules.For(fixture);
        Assert.NotNull(saved);

        // VBA resolves names case-insensitively, and the host reports whatever spelling the
        // editor holds. A lookup that missed on case would answer unknown for a module it has.
        Assert.False(saved!.PredeclaredIdOf("gadget"));
        Assert.False(saved.PredeclaredIdOf("GADGET"));
    }
}
