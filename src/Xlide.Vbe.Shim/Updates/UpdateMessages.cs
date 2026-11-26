using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Updates;

/// <summary>
/// Where the self-updater stands, as the page's dialog and the api's `update` route both read
/// it. One shape for both readers, so a suite that drives the dialog and one that drives the
/// route are looking at the same facts.
/// </summary>
/// <param name="Phase">idle, checking, upToDate, available, skipped, downloading, verified,
/// launched, exited or failed.</param>
/// <param name="Installed">The version this build is.</param>
/// <param name="DevelopmentBuild">True when the library is loaded from somewhere other than the
/// installer's folder, which is when installing would register a different copy.</param>
/// <param name="LoadedFrom">Where the library is loaded from.</param>
/// <param name="Host">The Office application this session is in, for the sentence that asks the
/// developer to close it.</param>
/// <param name="Feed">The address the newest release is read from.</param>
/// <param name="CheckedAt">When the last check finished, or null before one has.</param>
/// <param name="Manual">Whether the developer asked for the last check by hand.</param>
/// <param name="Version">The newest release's version, once a check has answered.</param>
/// <param name="Tag">Its tag.</param>
/// <param name="PublishedAt">When it was published.</param>
/// <param name="Notes">Its release notes, markdown as written.</param>
/// <param name="Page">Its page on the repository.</param>
/// <param name="Size">The installer's size in bytes.</param>
/// <param name="Received">How many bytes of it have arrived.</param>
/// <param name="Path">The verified installer on disk, once downloaded.</param>
/// <param name="Pid">The started installer's process id, once launched.</param>
/// <param name="ExitCode">The installer's exit code, when it ended while this session still ran -
/// which means it did not install.</param>
/// <param name="Error">Why the last step failed, in a sentence.</param>
/// <param name="SkippedVersion">The version the developer declined, if any.</param>
public sealed record UpdateStateReply(
    [property: JsonPropertyName("phase")] string Phase,
    [property: JsonPropertyName("installed")] string Installed,
    [property: JsonPropertyName("developmentBuild")] bool DevelopmentBuild,
    [property: JsonPropertyName("loadedFrom")] string? LoadedFrom,
    [property: JsonPropertyName("host")] string Host,
    [property: JsonPropertyName("feed")] string Feed,
    [property: JsonPropertyName("checkedAt")] string? CheckedAt,
    [property: JsonPropertyName("manual")] bool Manual,
    [property: JsonPropertyName("version")] string? Version,
    [property: JsonPropertyName("tag")] string? Tag,
    [property: JsonPropertyName("publishedAt")] string? PublishedAt,
    [property: JsonPropertyName("notes")] string? Notes,
    [property: JsonPropertyName("page")] string? Page,
    [property: JsonPropertyName("size")] long Size,
    [property: JsonPropertyName("received")] long Received,
    [property: JsonPropertyName("path")] string? Path,
    [property: JsonPropertyName("pid")] int Pid,
    [property: JsonPropertyName("exitCode")] int? ExitCode,
    [property: JsonPropertyName("failure")] string? Failure,
    [property: JsonPropertyName("skippedVersion")] string? SkippedVersion);

/// <summary>
/// The updater's state for the page. `open` asks the page to put its dialog up now: an
/// automatic check that found a newer release, any outcome of a check the developer asked for,
/// or a step that failed. Progress while the dialog already stands rides with it false.
/// </summary>
public sealed record SetUpdateMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("open")] bool Open,
    [property: JsonPropertyName("state")] UpdateStateReply State);
