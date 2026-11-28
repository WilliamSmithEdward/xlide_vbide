using System.Globalization;
using Xlide.Vbe.Core.Registration;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The registration plan decides whether the add-in loads at all. A wrong key produces silence:
/// the editor simply does not call us, with no error anywhere. These tests pin the shape.
/// </summary>
public class RegistrationPlanTests
{
    private const string ShimPath = @"C:\Program Files\xlide\Xlide.Vbe.Shim.dll";

    private static IReadOnlyList<RegistryEntry> Plan(HostBitness bitness = HostBitness.X64) =>
        RegistrationPlan.Build(ShimPath, bitness, RegistryScope.CurrentUser);

    [Fact]
    public void SixtyFourBitOfficeLooksUnderAddins64()
    {
        Assert.Equal(@"Software\Microsoft\VBA\VBE\6.0\Addins64", RegistrationPlan.AddInsKeyPath(HostBitness.X64));
    }

    [Fact]
    public void ThirtyTwoBitOfficeLooksUnderAddins()
    {
        Assert.Equal(@"Software\Microsoft\VBA\VBE\6.0\Addins", RegistrationPlan.AddInsKeyPath(HostBitness.X86));
    }

    [Fact]
    public void AddInKeyIsNamedForTheProgIdBecauseThatIsHowTheEditorFindsIt()
    {
        var addInKey = $@"{RegistrationPlan.AddInsKeyPath(HostBitness.X64)}\{ProductIdentity.AddInProgId}";
        Assert.Contains(Plan(), e => e.Path == addInKey);
    }

    [Fact]
    public void LoadBehaviourRequestsLoadAtStartup()
    {
        var entry = Plan().Single(e => e.Name == "LoadBehavior");

        Assert.True(entry.IsDword);
        Assert.Equal("3", entry.Value);
    }

    [Fact]
    public void ProgIdResolvesToTheClassIdentifier()
    {
        var entry = Plan().Single(e =>
            e.Path == $@"Software\Classes\{ProductIdentity.AddInProgId}\CLSID" && e.Name is null);

        Assert.Equal($"{{{ProductIdentity.AddInClsid}}}", entry.Value);
    }

    [Fact]
    public void ClassIdentifierResolvesToTheNativeServerDirectly()
    {
        // A native server is registered as itself. If this ever points at a runtime host, the
        // ahead-of-time build has silently regressed into a runtime-loading one.
        var entry = Plan().Single(e =>
            e.Path == $@"Software\Classes\CLSID\{{{ProductIdentity.AddInClsid}}}\InprocServer32" && e.Name is null);

        Assert.Equal(ShimPath, entry.Value);
    }

    [Fact]
    public void ServerIsApartmentThreadedBecauseTheEditorIsSingleThreaded()
    {
        var entries = Plan().Where(e => e.Name == "ThreadingModel").ToList();

        Assert.NotEmpty(entries);
        Assert.All(entries, e => Assert.Equal("Apartment", e.Value));
    }

    [Fact]
    public void EverythingIsWrittenUnderTheUserHiveSoInstallNeedsNoAdministrator()
    {
        Assert.All(Plan(), e =>
            Assert.True(
                e.Path.StartsWith(@"Software\Classes", StringComparison.Ordinal) ||
                e.Path.StartsWith(@"Software\Microsoft\VBA", StringComparison.Ordinal),
                $"unexpected registry location: {e.Path}"));
    }

    [Fact]
    public void ThirtyTwoBitClassResolvesToItsServerInTheViewAThirtyTwoBitHostReads()
    {
        var entry = Plan(HostBitness.X86).Single(e =>
            e.Path == $@"Software\Classes\WOW6432Node\CLSID\{{{ProductIdentity.AddInClsid}}}\InprocServer32"
            && e.Name is null);

        Assert.Equal(ShimPath, entry.Value);
    }

    [Fact]
    public void ThirtyTwoBitRegistrationLeavesTheSixtyFourBitClassAlone()
    {
        // The plain key is the one a 64-bit host loads from. A 32-bit library's path written
        // there is a library that host cannot load, and it would replace the one it can.
        Assert.DoesNotContain(Plan(HostBitness.X86), e =>
            e.Path.StartsWith(@"Software\Classes\CLSID\", StringComparison.Ordinal));
    }

    [Fact]
    public void SixtyFourBitRegistrationDoesNotUseTheCompatibilityNode()
    {
        Assert.DoesNotContain(Plan(), e => e.Path.Contains("WOW6432Node", StringComparison.Ordinal));
    }

    [Fact]
    public void BothWidthsCanBeRegisteredAtOnce()
    {
        // One installation registers both, each with its own library. Wherever the two plans
        // write the same value they must write the same thing, or whichever is applied second
        // undoes the first.
        var x64 = RegistrationPlan.Build(@"C:\xlide\Xlide.Vbe.Shim.dll", HostBitness.X64, RegistryScope.CurrentUser);
        var x86 = RegistrationPlan.Build(@"C:\xlide\x86\Xlide.Vbe.Shim.dll", HostBitness.X86, RegistryScope.CurrentUser);

        var shared = x64
            .Join(x86, left => (left.Path, left.Name), right => (right.Path, right.Name), (left, right) => (left, right))
            .ToList();

        // The ProgID, which is the one thing both widths read from the same key.
        Assert.Equal(2, shared.Count);
        Assert.All(shared, pair => Assert.Equal(pair.left.Value, pair.right.Value));
    }

    [Theory]
    [InlineData(HostBitness.X64)]
    [InlineData(HostBitness.X86)]
    public void EachWidthRegistersItsOwnAddInKeyAndItsOwnClass(HostBitness bitness)
    {
        var plan = Plan(bitness);
        var classKey = RegistrationPlan.ClassKeyPath(bitness, RegistryScope.CurrentUser);
        var addInKey = $@"{RegistrationPlan.AddInsKeyPath(bitness)}\{ProductIdentity.AddInProgId}";

        Assert.Contains(plan, e => e.Path == $@"{classKey}\InprocServer32" && e.Name is null && e.Value == ShimPath);
        Assert.Contains(plan, e => e.Path == addInKey && e.Name == "LoadBehavior");
    }

    [Fact]
    public void AnEmptyServerPathIsRejectedRatherThanRegisteringSomethingUnloadable()
    {
        Assert.Throws<ArgumentException>(() =>
            RegistrationPlan.Build("  ", HostBitness.X64, RegistryScope.CurrentUser));
    }

    [Fact]
    public void OverlayDetectionLooksAtTheVbaNamespaceOfficeActuallyOwns()
    {
        // Click-to-Run keeps its machine-level VBA values only inside this overlay, so its
        // presence is what identifies an Office whose registry views go through it.
        Assert.Equal(
            @"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\SOFTWARE\Microsoft\VBA",
            RegistrationPlan.OverlayVbaKey);
    }

    [Fact]
    public void OverlayRebasePreservesTheEntryShapeUnderTheClickToRunBranch()
    {
        // The overlay's MACHINE branch mirrors an ordinary hive, so the same plan entries must
        // land beneath it unchanged; a transformation here would fork the registration layout.
        var entry = Plan()[0];

        Assert.Equal(
            $@"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\{entry.Path}",
            RegistrationPlan.OverlayPath(entry.Path));
    }
}
