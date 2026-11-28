using Xlide.Vbe.Core.Hosting;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Where each shim looks for the files both of them share. A wrong answer is an editor that
/// loads and then shows the native pane, or analyses nothing, in one width of Office only.
/// </summary>
public class InstallLayoutTests
{
    private const string Install = @"C:\Users\someone\AppData\Local\Programs\xlide";

    [Fact]
    public void TheSixtyFourBitShimSharesTheFolderItIsIn()
    {
        Assert.Equal(Install, InstallLayout.SharedDirectory(Install));
    }

    [Fact]
    public void TheThirtyTwoBitShimLooksOneFolderUp()
    {
        Assert.Equal(Install, InstallLayout.SharedDirectory(Install + @"\x86"));
    }

    [Theory]
    [InlineData(@"\x86\")]
    [InlineData(@"\X86")]
    public void HoweverItsFolderIsSpelled(string tail)
    {
        Assert.Equal(Install, InstallLayout.SharedDirectory(Install + tail));
    }

    [Theory]
    [InlineData(@"F:\repo\artifacts\publish\Xlide.Vbe.Shim\debug_win-x86")]
    [InlineData(@"F:\repo\artifacts\publish\Xlide.Vbe.Shim\release_win-x64")]
    public void ADevelopmentPublishKeepsItsOwnCopyOfEverything(string publishFolder)
    {
        // Each architecture publishes whole, page and all, into a folder named for its runtime.
        Assert.Equal(publishFolder, InstallLayout.SharedDirectory(publishFolder));
    }

    [Fact]
    public void TheInstallerPutsTheThirtyTwoBitShimWhereThatShimLooksUpFrom()
    {
        var shim = InstallLayout.X86ShimPath(Install);

        Assert.Equal(Install + @"\x86\Xlide.Vbe.Shim.dll", shim);
        Assert.Equal(Install, InstallLayout.SharedDirectory(Path.GetDirectoryName(shim)!));
    }

    [Fact]
    public void AFolderWithNoParentIsItsOwn()
    {
        Assert.Equal("x86", InstallLayout.SharedDirectory("x86"));
    }
}
