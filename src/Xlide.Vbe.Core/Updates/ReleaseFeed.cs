using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Updates;

/// <summary>The installer a release carries, with the facts that let a download be checked.</summary>
/// <param name="Name">The asset's file name, always <see cref="ReleaseFeed.InstallerAssetName"/>.</param>
/// <param name="Size">Its length in bytes, as the release records it.</param>
/// <param name="Sha256">Its SHA-256 as GitHub computed it on upload, lower-case hex.</param>
/// <param name="DownloadUrl">Where the bytes are, on the repository's own download path.</param>
public sealed record ReleaseAsset(string Name, long Size, string Sha256, Uri DownloadUrl);

/// <summary>One published release, as the updater needs it.</summary>
/// <param name="Version">The three-part version the tag names, without its "v".</param>
/// <param name="Tag">The tag as published, "v0.20.21".</param>
/// <param name="PageUrl">The release's page, for a developer who wants to read before installing.</param>
/// <param name="Notes">The release notes as written, markdown, capped for a dialog.</param>
/// <param name="PublishedAt">When it was published, as the API wrote it.</param>
/// <param name="Installer">The installer asset.</param>
public sealed record Release(
    string Version,
    string Tag,
    Uri PageUrl,
    string Notes,
    string PublishedAt,
    ReleaseAsset Installer);

/// <summary>What a release means for the build that is running.</summary>
public enum UpdateVerdict
{
    /// <summary>The release is newer than this build and has not been declined.</summary>
    Available,

    /// <summary>The release is newer, and the developer chose to skip this one.</summary>
    Skipped,

    /// <summary>This build is the release, or newer than it (a development build usually is).</summary>
    UpToDate,
}

/// <summary>
/// Where releases come from, and which addresses an answer from there may point at.
///
/// The feed is the one address the updater asks. Everything else - the installer's download
/// address, the release page - arrives IN the feed's answer, and an answer that pointed anywhere
/// else would be the way a tampered reply turned into a downloaded executable. So the source
/// decides, before anything is fetched, whether an address belongs to it: the product's GitHub
/// repository and nothing else in a shipped build, and a loopback origin for the harness, which
/// serves a release of its own to drive the whole path without touching GitHub.
/// </summary>
public sealed class UpdateSource
{
    private readonly bool _loopback;

    private UpdateSource(Uri feedUrl, bool loopback)
    {
        FeedUrl = feedUrl;
        _loopback = loopback;
    }

    /// <summary>The address the newest release is read from.</summary>
    public Uri FeedUrl { get; }

    /// <summary>The product's own repository on GitHub: the only source a shipped build has.</summary>
    public static UpdateSource GitHub { get; } = new(new Uri(ReleaseFeed.LatestReleaseApiUrl), loopback: false);

    /// <summary>
    /// A feed served from this machine over plain HTTP on 127.0.0.1, for the harness. Refused
    /// for any other address, so this can never be turned toward the network.
    /// </summary>
    public static UpdateSource Loopback(Uri feedUrl)
    {
        ArgumentNullException.ThrowIfNull(feedUrl);
        if (!feedUrl.IsAbsoluteUri || feedUrl.Scheme != Uri.UriSchemeHttp || !feedUrl.IsLoopback
            || !string.Equals(feedUrl.Host, "127.0.0.1", StringComparison.Ordinal))
        {
            throw new ArgumentException("A loopback feed is http://127.0.0.1:<port>/... and nothing else.", nameof(feedUrl));
        }

        return new UpdateSource(feedUrl, loopback: true);
    }

    /// <summary>Whether an installer may be downloaded from an address this feed named.</summary>
    public bool AllowsDownload(Uri url)
    {
        ArgumentNullException.ThrowIfNull(url);
        if (_loopback)
        {
            return SameOrigin(url);
        }

        return url.IsAbsoluteUri
            && url.Scheme == Uri.UriSchemeHttps
            && string.Equals(url.Host, "github.com", StringComparison.OrdinalIgnoreCase)
            && url.AbsolutePath.StartsWith(ReleaseFeed.DownloadPathPrefix, StringComparison.Ordinal);
    }

    /// <summary>Whether a release page may be opened in the developer's browser.</summary>
    public bool AllowsPage(Uri url)
    {
        ArgumentNullException.ThrowIfNull(url);
        if (_loopback)
        {
            return SameOrigin(url);
        }

        return url.IsAbsoluteUri
            && url.Scheme == Uri.UriSchemeHttps
            && string.Equals(url.Host, "github.com", StringComparison.OrdinalIgnoreCase)
            && url.AbsolutePath.StartsWith(ReleaseFeed.RepositoryPathPrefix, StringComparison.Ordinal);
    }

    private bool SameOrigin(Uri url) =>
        url.IsAbsoluteUri
        && url.Scheme == FeedUrl.Scheme
        && string.Equals(url.Host, FeedUrl.Host, StringComparison.OrdinalIgnoreCase)
        && url.Port == FeedUrl.Port;
}

/// <summary>
/// Reads the GitHub releases API's "latest release" answer into a <see cref="Release"/>, and
/// judges it against the build that is running.
///
/// The parse is strict on purpose. The updater will download and run what this names, so an
/// answer with no installer, no digest, a tag that is not a version, or an address off the
/// repository's own download path is refused with a sentence rather than guessed at. The digest
/// is GitHub's own SHA-256 of the uploaded asset, served over TLS from the API; checking the
/// download against it is what stands in for a code signature.
/// </summary>
public static partial class ReleaseFeed
{
    /// <summary>The installer's file name in every release.</summary>
    public const string InstallerAssetName = "xlide-setup.exe";

    /// <summary>Release notes longer than this are cut for the dialog; the page has the rest.</summary>
    public const int NotesLimit = 6000;

    /// <summary>The API address of the newest release.</summary>
    public static string LatestReleaseApiUrl { get; } = string.Create(CultureInfo.InvariantCulture,
        $"https://api.github.com/repos/{ProductIdentity.RepositoryOwner}/{ProductIdentity.RepositoryName}/releases/latest");

    /// <summary>The releases page a developer reads.</summary>
    public static string ReleasesPageUrl { get; } = string.Create(CultureInfo.InvariantCulture,
        $"https://github.com/{ProductIdentity.RepositoryOwner}/{ProductIdentity.RepositoryName}/releases");

    /// <summary>Every path under the repository on github.com starts with this.</summary>
    internal static string RepositoryPathPrefix { get; } = string.Create(CultureInfo.InvariantCulture,
        $"/{ProductIdentity.RepositoryOwner}/{ProductIdentity.RepositoryName}/");

    /// <summary>Every release asset's download path starts with this.</summary>
    internal static string DownloadPathPrefix { get; } = RepositoryPathPrefix + "releases/download/";

    /// <summary>The file a downloaded installer is kept as, named for the version it installs.</summary>
    public static string DownloadFileName(string version)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(version);
        return string.Create(CultureInfo.InvariantCulture, $"xlide-setup-{version}.exe");
    }

    /// <summary>
    /// Reads one release from the API's JSON. Null with a reason when the answer is not a release
    /// this product can install from.
    /// </summary>
    public static Release? Parse(string json, UpdateSource source, out string? refusal)
    {
        ArgumentNullException.ThrowIfNull(source);
        refusal = null;

        if (string.IsNullOrWhiteSpace(json))
        {
            refusal = "the release feed answered nothing";
            return null;
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json);
        }
        catch (JsonException)
        {
            refusal = "the release feed did not answer JSON";
            return null;
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                refusal = "the release feed did not answer a release";
                return null;
            }

            if (Flag(root, "draft") || Flag(root, "prerelease"))
            {
                refusal = "the newest release is a draft or a pre-release";
                return null;
            }

            var tag = Text(root, "tag_name");
            if (!ProductVersion.TryParse(tag, out var version))
            {
                refusal = tag is null ? "the release has no tag" : $"the release tag '{tag}' is not a version";
                return null;
            }

            if (!Uri.TryCreate(Text(root, "html_url"), UriKind.Absolute, out var page) || !source.AllowsPage(page))
            {
                refusal = "the release page is not on the product's repository";
                return null;
            }

            if (!root.TryGetProperty("assets", out var assets) || assets.ValueKind != JsonValueKind.Array)
            {
                refusal = "the release lists no files";
                return null;
            }

            ReleaseAsset? installer = null;
            foreach (var asset in assets.EnumerateArray())
            {
                if (asset.ValueKind != JsonValueKind.Object
                    || !string.Equals(Text(asset, "name"), InstallerAssetName, StringComparison.Ordinal))
                {
                    continue;
                }

                var size = asset.TryGetProperty("size", out var sizeValue)
                    && sizeValue.ValueKind == JsonValueKind.Number
                    && sizeValue.TryGetInt64(out var bytes)
                    ? bytes
                    : 0;
                if (size <= 0)
                {
                    refusal = $"the release's {InstallerAssetName} has no size";
                    return null;
                }

                var digest = Text(asset, "digest");
                var hash = digest is not null ? Sha256Digest().Match(digest) : null;
                if (hash is not { Success: true })
                {
                    refusal = $"the release's {InstallerAssetName} carries no sha256 digest";
                    return null;
                }

                if (!Uri.TryCreate(Text(asset, "browser_download_url"), UriKind.Absolute, out var download)
                    || !source.AllowsDownload(download))
                {
                    refusal = $"the release's {InstallerAssetName} is not on the repository's download path";
                    return null;
                }

                installer = new ReleaseAsset(InstallerAssetName, size, hash.Groups["hex"].Value.ToLowerInvariant(), download);
                break;
            }

            if (installer is null)
            {
                refusal = $"the release has no {InstallerAssetName}";
                return null;
            }

            var notes = (Text(root, "body") ?? string.Empty).Replace("\r\n", "\n", StringComparison.Ordinal).Trim();
            if (notes.Length > NotesLimit)
            {
                notes = notes[..NotesLimit].TrimEnd() + "\n...";
            }

            return new Release(
                ProductVersion.Format(version),
                tag!,
                page,
                notes,
                Text(root, "published_at") ?? string.Empty,
                installer);
        }
    }

    /// <summary>What a release means for the build that is running.</summary>
    public static UpdateVerdict Judge(Release release, string installedVersion, string? skippedVersion)
    {
        ArgumentNullException.ThrowIfNull(release);

        if (ProductVersion.Compare(release.Version, installedVersion) <= 0)
        {
            return UpdateVerdict.UpToDate;
        }

        return skippedVersion is not null && ProductVersion.Compare(release.Version, skippedVersion) == 0
            ? UpdateVerdict.Skipped
            : UpdateVerdict.Available;
    }

    private static string? Text(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    private static bool Flag(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;

    [GeneratedRegex("^sha256:(?<hex>[0-9a-fA-F]{64})$")]
    private static partial Regex Sha256Digest();
}
