using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The project's own conditional compilation arguments, read out of the saved document.
///
/// WHY THEY MATTER. The analyzer knows the compiler's constants - VBA7, Win64, Mac - and drops
/// the arms they decide. A project's OWN constants live in the VBE's Project Properties box, the
/// object model has no property for them, and without them every `#If MY_FLAG Then` is
/// undecidable: both arms are analyzed, both contribute declarations, and a finding can be
/// reported from an arm the compiler never compiles.
///
/// Two halves, the way <see cref="SavedModulesTests"/> has them. TOLERANCE always runs, because
/// the promise is that an unreadable document answers null rather than throwing - conditional
/// constants are an optimisation of correctness, and a reader that can fail is worse than one
/// that cannot answer. AGREEMENT runs against a workbook VBA ITSELF wrote, which is the only way
/// to know the record was read and not merely parsed: this product cannot write the record, so a
/// fixture it generated would prove nothing about the format.
/// </summary>
public class ProjectConstantsTests
{
    /// <summary>The repository root, walking up from wherever the tests were built to.</summary>
    private static string? RepoRoot()
    {
        var here = AppContext.BaseDirectory;
        for (var up = 0; up < 8 && here is not null; up++)
        {
            if (File.Exists(Path.Combine(here, "xlide_vbide.slnx")))
            {
                return here;
            }

            here = Path.GetDirectoryName(here);
        }

        return null;
    }

    /// <summary>
    /// A workbook the VBE saved with a conditional compilation argument set.
    ///
    /// It lives in the neighbouring analyzer checkout's Excel corpus rather than in this
    /// repository, because this product has no way to CREATE one: the VBE writes the record when
    /// the project is saved and offers no programmatic way to set the property, so a fixture
    /// generated here would carry no record and the agreement half would pass over nothing.
    /// Absent on a clone without the neighbour, and skipped rather than failed.
    /// </summary>
    private static string? Corpus()
    {
        if (RepoRoot() is not { } root)
        {
            return null;
        }

        var neighbour = Path.GetDirectoryName(root);
        if (neighbour is null)
        {
            return null;
        }

        var book = Path.Combine(neighbour, "xlide_vscode", "excel_test_workbook", "fullBuild.xlsm");
        return File.Exists(book) ? book : null;
    }

    [Fact]
    public void ADocumentThatIsNotAPackageAnswersNothing()
    {
        var scratch = Path.Combine(Path.GetTempPath(), $"xlide-constants-{Guid.NewGuid():N}.xlsm");
        File.WriteAllText(scratch, "this is not a compound file");

        try
        {
            Assert.Null(SavedModules.For(scratch)?.ConditionalConstants);
        }
        finally
        {
            File.Delete(scratch);
        }
    }

    [Fact]
    public void ADocumentThatIsNotThereAnswersNothing()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"xlide-absent-{Guid.NewGuid():N}.xlsm");
        Assert.Null(SavedModules.For(missing)?.ConditionalConstants);
    }

    [Fact]
    public void AProjectWithNoConstantsAnswersNothing()
    {
        if (RepoRoot() is not { } root)
        {
            return;
        }

        var fixture = Path.Combine(root, "artifacts", "fixtures", "DebugFixture.xlsm");
        if (!File.Exists(fixture))
        {
            return;
        }

        // NOT AN EMPTY STRING. A project that declares none has no record at all, and the
        // difference matters downstream: null is "nothing to tell the analyzer", where an empty
        // string would be an environment that says every flag is undefined.
        Assert.Null(SavedModules.For(fixture)?.ConditionalConstants);
    }

    [Fact]
    public void AProjectDeclaringAConstantAnswersIt()
    {
        if (Corpus() is not { } book)
        {
            return;
        }

        var raw = SavedModules.For(book)?.ConditionalConstants;

        Assert.NotNull(raw);
        Assert.Contains("=", raw);

        // The VBE writes `Name = Value`, pairs separated by colons. Read against the file rather
        // than against a remembered string: what is asserted is that a real record was found and
        // says something of that shape, not that this particular corpus workbook still sets this
        // particular flag.
        foreach (var entry in raw!.Split(':', StringSplitOptions.RemoveEmptyEntries))
        {
            var equals = entry.IndexOf('=', StringComparison.Ordinal);
            Assert.True(equals > 0, $"'{entry}' is not a Name = Value pair");
            Assert.False(string.IsNullOrWhiteSpace(entry[..equals]), $"'{entry}' names nothing");
        }
    }
}
