using System.Text.RegularExpressions;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Updates;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The version the updater compares against is read from the assembly, which takes it from
/// Directory.Build.props. If those two ever disagreed the updater would offer the build that is
/// already running, or stay quiet about a newer one, so the props file is the oracle here.
/// </summary>
public partial class ProductVersionTests
{
    [Fact]
    public void TheCurrentVersionIsTheOneInDirectoryBuildProps()
    {
        var props = FindUpward("Directory.Build.props");
        var declared = VersionElement().Match(File.ReadAllText(props)).Groups[1].Value;

        Assert.Matches(@"^\d+\.\d+\.\d+$", declared);
        Assert.Equal(declared, ProductVersion.Current);
    }

    [Fact]
    public void TheRepositoryIdentityIsTheOneInDirectoryBuildProps()
    {
        var props = FindUpward("Directory.Build.props");
        var declared = RepositoryUrlElement().Match(File.ReadAllText(props)).Groups[1].Value;

        Assert.Equal(
            $"https://github.com/{ProductIdentity.RepositoryOwner}/{ProductIdentity.RepositoryName}",
            declared);
    }

    [Theory]
    [InlineData("1.2.3", 1, 2, 3)]
    [InlineData("v1.2.3", 1, 2, 3)]
    [InlineData("V0.20.21", 0, 20, 21)]
    [InlineData(" v10.0.0 ", 10, 0, 0)]
    public void ReadsAThreePartTagWithOrWithoutItsV(string text, int major, int minor, int build)
    {
        Assert.True(ProductVersion.TryParse(text, out var version));
        Assert.Equal(new Version(major, minor, build), version);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("latest")]
    [InlineData("1.2")]
    [InlineData("1.2.3.4")]
    [InlineData("1.2.3-beta")]
    [InlineData("v1.2.3+build")]
    [InlineData("1..3")]
    [InlineData("-1.2.3")]
    [InlineData("vv1.2.3")]
    public void RefusesAnythingThatIsNotAThreePartVersion(string? text)
    {
        Assert.False(ProductVersion.TryParse(text, out _));
    }

    [Theory]
    [InlineData("0.20.21", "0.20.22", -1)]
    [InlineData("0.20.22", "0.20.21", 1)]
    [InlineData("v0.20.21", "0.20.21", 0)]
    [InlineData("0.9.9", "0.10.0", -1)]
    [InlineData("1.0.0", "0.99.99", 1)]
    public void OrdersVersionsNumerically(string left, string right, int expectedSign)
    {
        Assert.Equal(expectedSign, Math.Sign(ProductVersion.Compare(left, right)));
    }

    [Fact]
    public void TextThatIsNotAVersionIsOlderThanAnyVersion()
    {
        Assert.True(ProductVersion.Compare("latest", "0.0.1") < 0);
        Assert.True(ProductVersion.Compare("0.0.1", null) > 0);
        Assert.Equal(0, ProductVersion.Compare("garbage", "also garbage"));
    }

    private static string FindUpward(string fileName)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, fileName);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new FileNotFoundException($"No {fileName} above {AppContext.BaseDirectory}.");
    }

    [GeneratedRegex("<Version>([^<]+)</Version>")]
    private static partial Regex VersionElement();

    [GeneratedRegex("<RepositoryUrl>([^<]+)</RepositoryUrl>")]
    private static partial Regex RepositoryUrlElement();
}
