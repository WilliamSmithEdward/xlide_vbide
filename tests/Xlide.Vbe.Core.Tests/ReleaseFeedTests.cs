using Xlide.Vbe.Core.Updates;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The release feed is the one thing that stands between "GitHub said so" and "run this
/// executable", so every way its answer could be short of a release is a refusal here, and
/// the happy path is the real shape of the API's answer for v0.20.21.
/// </summary>
public class ReleaseFeedTests
{
    private const string Digest = "sha256:9817edc6bc781c8ac66c4653879a0e44ee3d37e3aa97d9c855741f21bcc97e1e";

    /// <summary>The live answer's shape for v0.20.21, cut to the fields the parser reads plus noise.</summary>
    private static string Latest(
        string tag = "v0.20.21",
        string assetName = "xlide-setup.exe",
        string digest = Digest,
        string download = "https://github.com/WilliamSmithEdward/xlide_vbide/releases/download/v0.20.21/xlide-setup.exe",
        string page = "https://github.com/WilliamSmithEdward/xlide_vbide/releases/tag/v0.20.21",
        long size = 33608192,
        bool prerelease = false,
        bool draft = false,
        string body = "Patch release for the Rename Symbol box (#128).\r\n\r\n- A name typed into Rename Symbol is readable.") =>
        $$"""
        {
          "url": "https://api.github.com/repos/WilliamSmithEdward/xlide_vbide/releases/270112",
          "html_url": "{{page}}",
          "tag_name": "{{tag}}",
          "name": "{{tag}}",
          "draft": {{(draft ? "true" : "false")}},
          "prerelease": {{(prerelease ? "true" : "false")}},
          "published_at": "2026-10-28T17:19:39Z",
          "author": { "login": "WilliamSmithEdward" },
          "assets": [
            { "name": "malware-report.json", "size": 2233, "digest": "sha256:304cfe6c1b463d56a39183fbf7594ebf8e1730bbffcd638f6d62ed12da3d1930",
              "browser_download_url": "https://github.com/WilliamSmithEdward/xlide_vbide/releases/download/v0.20.21/malware-report.json" },
            { "name": "{{assetName}}", "size": {{size}}, "digest": {{(digest is null ? "null" : $"\"{digest}\"")}},
              "content_type": "application/x-msdos-program",
              "browser_download_url": "{{download}}" }
          ],
          "body": {{System.Text.Json.JsonSerializer.Serialize(body)}}
        }
        """;

    [Fact]
    public void ReadsTheLiveAnswerShape()
    {
        var release = ReleaseFeed.Parse(Latest(), UpdateSource.GitHub, out var refusal);

        Assert.Null(refusal);
        Assert.NotNull(release);
        Assert.Equal("0.20.21", release.Version);
        Assert.Equal("v0.20.21", release.Tag);
        Assert.Equal("https://github.com/WilliamSmithEdward/xlide_vbide/releases/tag/v0.20.21", release.PageUrl.ToString());
        Assert.Equal("2026-10-28T17:19:39Z", release.PublishedAt);
        Assert.Equal("Patch release for the Rename Symbol box (#128).\n\n- A name typed into Rename Symbol is readable.", release.Notes);
        Assert.Equal("xlide-setup.exe", release.Installer.Name);
        Assert.Equal(33608192, release.Installer.Size);
        Assert.Equal("9817edc6bc781c8ac66c4653879a0e44ee3d37e3aa97d9c855741f21bcc97e1e", release.Installer.Sha256);
        Assert.Equal(
            "https://github.com/WilliamSmithEdward/xlide_vbide/releases/download/v0.20.21/xlide-setup.exe",
            release.Installer.DownloadUrl.ToString());
    }

    [Fact]
    public void TheDigestIsLowerCasedHoweverItArrives()
    {
        var release = ReleaseFeed.Parse(Latest(digest: Digest.ToUpperInvariant().Replace("SHA256", "sha256", StringComparison.Ordinal)), UpdateSource.GitHub, out _);
        Assert.Equal("9817edc6bc781c8ac66c4653879a0e44ee3d37e3aa97d9c855741f21bcc97e1e", release!.Installer.Sha256);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not json")]
    [InlineData("[]")]
    [InlineData("\"a string\"")]
    public void RefusesAnAnswerThatIsNotARelease(string json)
    {
        Assert.Null(ReleaseFeed.Parse(json, UpdateSource.GitHub, out var refusal));
        Assert.NotNull(refusal);
    }

    [Fact]
    public void RefusesAPreReleaseAndADraft()
    {
        Assert.Null(ReleaseFeed.Parse(Latest(prerelease: true), UpdateSource.GitHub, out var refusal));
        Assert.Contains("pre-release", refusal, StringComparison.Ordinal);
        Assert.Null(ReleaseFeed.Parse(Latest(draft: true), UpdateSource.GitHub, out _));
    }

    [Theory]
    [InlineData("latest")]
    [InlineData("v1.2")]
    [InlineData("v1.2.3-rc1")]
    public void RefusesATagThatIsNotAVersion(string tag)
    {
        Assert.Null(ReleaseFeed.Parse(Latest(tag: tag), UpdateSource.GitHub, out var refusal));
        Assert.Contains("not a version", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void RefusesAReleaseWithoutTheInstaller()
    {
        Assert.Null(ReleaseFeed.Parse(Latest(assetName: "xlide-setup.zip"), UpdateSource.GitHub, out var refusal));
        Assert.Equal("the release has no xlide-setup.exe", refusal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("md5:9817edc6bc781c8ac66c4653879a0e44ee3d37e3aa97d9c855741f21bcc97e1e")]
    [InlineData("sha256:9817edc6")]
    [InlineData("sha256:zz17edc6bc781c8ac66c4653879a0e44ee3d37e3aa97d9c855741f21bcc97e1e")]
    public void RefusesAnInstallerWithoutASha256Digest(string? digest)
    {
        Assert.Null(ReleaseFeed.Parse(Latest(digest: digest!), UpdateSource.GitHub, out var refusal));
        Assert.Contains("sha256", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void RefusesAnInstallerWithNoSize()
    {
        Assert.Null(ReleaseFeed.Parse(Latest(size: 0), UpdateSource.GitHub, out var refusal));
        Assert.Contains("no size", refusal, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("https://example.com/xlide-setup.exe")]
    [InlineData("http://github.com/WilliamSmithEdward/xlide_vbide/releases/download/v0.20.21/xlide-setup.exe")]
    [InlineData("https://github.com/SomeoneElse/xlide_vbide/releases/download/v0.20.21/xlide-setup.exe")]
    [InlineData("https://github.com/WilliamSmithEdward/xlide_vbide/raw/main/xlide-setup.exe")]
    [InlineData("https://objects.githubusercontent.com/github-production-release-asset/xlide-setup.exe")]
    [InlineData("not a url")]
    public void RefusesADownloadOffTheRepositorysOwnPath(string download)
    {
        Assert.Null(ReleaseFeed.Parse(Latest(download: download), UpdateSource.GitHub, out var refusal));
        Assert.Contains("download path", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void RefusesAReleasePageOffTheRepository()
    {
        Assert.Null(ReleaseFeed.Parse(Latest(page: "https://example.com/releases"), UpdateSource.GitHub, out var refusal));
        Assert.Contains("release page", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void CutsLongNotesForTheDialog()
    {
        var release = ReleaseFeed.Parse(Latest(body: new string('x', ReleaseFeed.NotesLimit + 500)), UpdateSource.GitHub, out _);
        Assert.Equal(ReleaseFeed.NotesLimit + "\n...".Length, release!.Notes.Length);
        Assert.EndsWith("\n...", release.Notes, StringComparison.Ordinal);
    }

    [Fact]
    public void ALoopbackSourceAcceptsOnlyItsOwnOrigin()
    {
        var source = UpdateSource.Loopback(new Uri("http://127.0.0.1:43210/latest"));
        var local = Latest(
            download: "http://127.0.0.1:43210/xlide-setup.exe",
            page: "http://127.0.0.1:43210/release");

        var release = ReleaseFeed.Parse(local, source, out var refusal);
        Assert.Null(refusal);
        Assert.Equal("http://127.0.0.1:43210/xlide-setup.exe", release!.Installer.DownloadUrl.ToString());

        // The real GitHub address is refused by a loopback source, and the loopback one by GitHub's:
        // neither can be turned toward the other's origin.
        Assert.Null(ReleaseFeed.Parse(Latest(page: "http://127.0.0.1:43210/release"), source, out _));
        Assert.Null(ReleaseFeed.Parse(local, UpdateSource.GitHub, out _));
        Assert.Null(ReleaseFeed.Parse(
            Latest(download: "http://127.0.0.1:9/xlide-setup.exe", page: "http://127.0.0.1:43210/release"), source, out _));
    }

    [Theory]
    [InlineData("https://127.0.0.1:43210/latest")]
    [InlineData("http://localhost:43210/latest")]
    [InlineData("http://192.168.1.2:43210/latest")]
    [InlineData("http://api.github.com/latest")]
    public void ALoopbackSourceIsPlainHttpOn127001AndNothingElse(string feed)
    {
        Assert.Throws<ArgumentException>(() => UpdateSource.Loopback(new Uri(feed)));
    }

    [Fact]
    public void TheGitHubSourceAsksTheRepositorysLatestRelease()
    {
        Assert.Equal(
            "https://api.github.com/repos/WilliamSmithEdward/xlide_vbide/releases/latest",
            UpdateSource.GitHub.FeedUrl.ToString());
        Assert.Equal("https://github.com/WilliamSmithEdward/xlide_vbide/releases", ReleaseFeed.ReleasesPageUrl);
        Assert.Equal("xlide-setup-0.20.22.exe", ReleaseFeed.DownloadFileName("0.20.22"));
    }

    [Theory]
    [InlineData("0.20.21", "0.20.22", null, UpdateVerdict.Available)]
    [InlineData("0.20.21", "0.20.22", "0.20.22", UpdateVerdict.Skipped)]
    [InlineData("0.20.21", "0.20.22", "0.20.20", UpdateVerdict.Available)]
    [InlineData("0.20.22", "0.20.22", null, UpdateVerdict.UpToDate)]
    [InlineData("0.21.0", "0.20.22", null, UpdateVerdict.UpToDate)]
    [InlineData("0.21.0", "0.20.22", "0.20.22", UpdateVerdict.UpToDate)]
    public void JudgesTheReleaseAgainstTheRunningBuild(string installed, string offered, string? skipped, UpdateVerdict expected)
    {
        var release = ReleaseFeed.Parse(
            Latest(tag: "v" + offered, download: $"https://github.com/WilliamSmithEdward/xlide_vbide/releases/download/v{offered}/xlide-setup.exe"),
            UpdateSource.GitHub, out _);

        Assert.Equal(expected, ReleaseFeed.Judge(release!, installed, skipped));
    }
}
