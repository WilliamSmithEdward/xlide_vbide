using System;
using System.Diagnostics;
using System.IO;
using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// What reading the saved attribute headers costs, on the path that runs while somebody types.
///
/// THIS IS THE WHOLE REASON THE READER WORKS THE WAY IT DOES. The other way to learn
/// `VB_PredeclaredId` is to Export each class module to a temporary file and read the header
/// back, which is a file written and deleted per class per pass - and the pass that would do it,
/// <c>ProjectReader.ReadAll</c>, runs on the host's user interface thread on every analysis, not
/// only when a project opens. Reading the package instead costs one PARSE per save, and a stat
/// per pass after that.
///
/// So the two numbers below are the design, measured on LanguageFixture.xlsm (2026-08-23):
///
///   cold   11ms   the whole package walked: unzip, compound file, decompress the module table,
///                 and decompress the head of every one of its eight module streams. Paid when
///                 the document is saved, and not again until it is saved again.
///   warm    4us   what every pass after that actually pays, which is one file timestamp.
///                 1000 passes came to 4ms.
///
/// Against the export it replaces: a temporary file written, read and deleted PER CLASS MODULE
/// PER PASS. At even a few milliseconds each, a project carrying a dozen classes - which the
/// stdVBA library alone is - would have spent longer than this whole walk, on the user interface
/// thread, every time somebody stopped typing.
///
/// The bounds are loose enough not to fail on a slow or busy machine and tight enough to notice
/// the shape changing, which is what they are for.
/// </summary>
public sealed class SavedModulesCostTests
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

    [Fact]
    public void TheFirstReadOfADocumentIsCheapEnoughToDoOnASave()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        // A COPY per read, because the cache is keyed by path and write time: reading the same
        // path twice would measure the dictionary, which is the next test's job.
        var copy = Path.Combine(Path.GetTempPath(), $"xlide-cost-{Guid.NewGuid():N}.xlsm");
        File.Copy(fixture!, copy);
        try
        {
            // Warm the framework's zip and file machinery, so the figure is the walk rather than
            // the first call through .NET.
            Assert.NotNull(SavedModules.For(copy));

            var again = Path.Combine(Path.GetTempPath(), $"xlide-cost-{Guid.NewGuid():N}.xlsm");
            File.Copy(fixture!, again);
            try
            {
                var clock = Stopwatch.StartNew();
                var read = SavedModules.For(again);
                clock.Stop();

                Assert.NotNull(read);
                Assert.False(read!.PredeclaredIdOf("Gadget"));
                Assert.True(clock.ElapsedMilliseconds < 250,
                    $"walking a saved package took {clock.ElapsedMilliseconds}ms, which is no longer "
                    + "the cost of a save");
            }
            finally
            {
                File.Delete(again);
            }
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void EveryReadAfterThatIsAStat()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No LanguageFixture.xlsm; run tools\\New-Fixture.ps1.");

        Assert.NotNull(SavedModules.For(fixture));

        // A thousand passes, which is far more than a typing session produces. The point is that
        // the package is not walked again while the file has not moved.
        var clock = Stopwatch.StartNew();
        for (var pass = 0; pass < 1000; pass++)
        {
            _ = SavedModules.For(fixture);
        }

        clock.Stop();

        Assert.True(clock.ElapsedMilliseconds < 500,
            $"1000 passes cost {clock.ElapsedMilliseconds}ms, so something is re-walking the "
            + "package rather than answering from the write time");
    }
}
