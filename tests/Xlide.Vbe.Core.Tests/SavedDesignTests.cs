using Xlide.Vbe.Core.Forms;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The saved baseline: which properties a workbook's own storage says are not the file format
/// default.
///
/// Two halves, deliberately. The TOLERANCE half needs nothing and always runs, because the
/// promise this makes to the projection is that it never throws and never lies - a workbook it
/// cannot read is a form with no baseline, which is a real answer. The AGREEMENT half runs
/// against the generated fixture and pins the numbers the reader was built against; the fixture
/// is machine-made rather than committed, so it skips where there is none, and the same
/// comparison is driven live from designer-features.mjs against the running editor.
/// </summary>
public class SavedDesignTests
{
    /// <summary>The fixture, or null on a checkout that has not generated one.</summary>
    private static string? Fixture()
    {
        var here = AppContext.BaseDirectory;
        for (var up = 0; up < 8 && here is not null; up++)
        {
            if (File.Exists(Path.Combine(here, "xlide_vbide.slnx")))
            {
                var fixture = Path.Combine(here, "artifacts", "fixtures", "FormFixture.xlsm");
                return File.Exists(fixture) ? fixture : null;
            }

            here = Path.GetDirectoryName(here);
        }

        return null;
    }

    [Fact]
    public void AWorkbookThatIsNotThereHasNoBaseline()
    {
        Assert.Null(SavedDesign.For(null));
        Assert.Null(SavedDesign.For(string.Empty));
        Assert.Null(SavedDesign.For("   "));
        Assert.Null(SavedDesign.For(Path.Combine(Path.GetTempPath(), "no-such-workbook.xlsm")));
    }

    [Fact]
    public void AFileThatIsNotAWorkbookHasNoBaseline()
    {
        var text = Path.Combine(Path.GetTempPath(), $"xlide-not-a-workbook-{Guid.NewGuid():N}.xlsm");
        File.WriteAllText(text, "This is not a package, and saying so must not throw.");
        try
        {
            Assert.Null(SavedDesign.For(text));
        }
        finally
        {
            File.Delete(text);
        }
    }

    [Fact]
    public void BytesThatAreNotACompoundFileAreRefusedRatherThanParsed()
    {
        Assert.Null(CompoundFile.TryRead([]));
        Assert.Null(CompoundFile.TryRead(new byte[511]));
        Assert.Null(CompoundFile.TryRead(new byte[4096]));

        // The right length and the wrong signature is the case that would otherwise walk a FAT
        // made of zeroes.
        var wrong = new byte[4096];
        wrong[0] = 0xD0;
        Assert.Null(CompoundFile.TryRead(wrong));
    }

    [Fact]
    public void TheFixtureReportsWhatItWasBuiltWith()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No FormFixture.xlsm; run tools\\New-FormFixture.ps1.");

        var design = SavedDesign.For(fixture);
        Assert.NotNull(design);
        Assert.True(design.Knows("EntryForm"));
        Assert.False(design.Knows("NoSuchForm"));

        // form-plan.mjs gives the OK button a caption and a picture with a placement, and the
        // Image nothing but a picture. Both come back as exactly that.
        Assert.Contains("Caption", design.ChangedOn("EntryForm", "OkButton"));
        Assert.Contains("Picture", design.ChangedOn("EntryForm", "OkButton"));
        Assert.Contains("PicturePosition", design.ChangedOn("EntryForm", "OkButton"));
        Assert.Contains("Picture", design.ChangedOn("EntryForm", "Badge"));
        Assert.Contains("PictureSizeMode", design.ChangedOn("EntryForm", "Badge"));
        Assert.DoesNotContain("Caption", design.ChangedOn("EntryForm", "Badge"));

        // A ScrollBar and a SpinButton differ from the file's default in one thing each.
        Assert.Contains("Orientation", design.ChangedOn("EntryForm", "Amount"));
        Assert.Contains("Orientation", design.ChangedOn("EntryForm", "Steps"));

        // Containment: one storage deep for the Frame, two for a control on a MultiPage's page.
        Assert.Contains("Caption", design.ChangedOn("EntryForm", "Options.PickGround"));
        Assert.Contains("Caption", design.ChangedOn("EntryForm", "Options.PickAir"));
        Assert.Contains("Caption", design.ChangedOn("EntryForm", "Wizard.Page1.Agree"));

        // A control nobody asked about, and a control that is not there.
        Assert.Empty(design.ChangedOn("EntryForm", "NoSuchControl"));
        Assert.Empty(design.ChangedOn("NoSuchForm", "OkButton"));
    }

    [Fact]
    public void NothingStructuralReachesTheAnswer()
    {
        var fixture = Fixture();
        Assert.SkipWhen(fixture is null, "No FormFixture.xlsm; run tools\\New-FormFixture.ps1.");

        var design = SavedDesign.For(fixture);
        Assert.NotNull(design);

        var controls = design.Controls("EntryForm").ToList();
        Assert.NotEmpty(controls);

        // A MultiPage carries its own TabStrip as a nameless site. Nothing can ask about a
        // control with no name, so it must not be in the answer under any spelling.
        Assert.DoesNotContain(controls, one => one.Contains('?', StringComparison.Ordinal));

        foreach (var control in controls)
        {
            var changed = design.ChangedOn("EntryForm", control);

            // Size is MUST-be-1 in every mask and TabIndex is set on all but the first control,
            // so either one appearing means the structural filter stopped working - and a list
            // that always says the same thing says nothing.
            Assert.DoesNotContain("Size", changed);
            Assert.DoesNotContain("TabIndex", changed);
            Assert.DoesNotContain("ObjectStreamSize", changed);
            Assert.DoesNotContain("ClsidCacheIndex", changed);

            // The list is sorted and free of repeats, because a control described twice - once as
            // a site, once by its own storage - merges rather than appending.
            Assert.Equal(changed.Distinct().Order(StringComparer.Ordinal), changed);
        }
    }
}
