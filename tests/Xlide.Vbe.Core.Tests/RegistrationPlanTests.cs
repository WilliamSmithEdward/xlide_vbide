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
    public void ThirtyTwoBitRegistrationAddsTheCompatibilityNode()
    {
        var wow = Plan(HostBitness.X86)
            .Where(e => e.Path.Contains("WOW6432Node", StringComparison.Ordinal))
            .ToList();

        Assert.NotEmpty(wow);
    }

    [Fact]
    public void SixtyFourBitRegistrationDoesNotUseTheCompatibilityNode()
    {
        Assert.DoesNotContain(Plan(), e => e.Path.Contains("WOW6432Node", StringComparison.Ordinal));
    }

    [Fact]
    public void AnEmptyServerPathIsRejectedRatherThanRegisteringSomethingUnloadable()
    {
        Assert.Throws<ArgumentException>(() =>
            RegistrationPlan.Build("  ", HostBitness.X64, RegistryScope.CurrentUser));
    }
}
