using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Changes;

/*
 * The change log's answers, in the shapes the pane and the xlide api both read.
 *
 * NOT IN THE DEBUG SERVER, though the route reads them too. Everything in Diagnostics/ApiServer.cs
 * sits inside `#if DEBUG`, and the Changes PANE is a shipped feature: putting these beside the
 * route's other replies built in Debug and broke Release, where the types simply are not there
 * (2026-08-22, caught by CI within the minute). Sync answers the dialog and the door from one set
 * of records for exactly this reason, and this follows it.
 */

/// <summary>One module's before and after within a round.</summary>
public sealed record ChangeEntryReply(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("added")] int Added,
    [property: JsonPropertyName("removed")] int Removed,
    [property: JsonPropertyName("from")] string? From,
    [property: JsonPropertyName("held")] bool Held);

/// <summary>
/// One round: a stretch of writes by one hand, ended by a snapshot, by the hand changing, by a
/// save, or by a long enough silence.
/// </summary>
public sealed record ChangeRoundReply(
    [property: JsonPropertyName("round")] int Round,
    [property: JsonPropertyName("started")] string Started,
    [property: JsonPropertyName("ended")] string Ended,
    [property: JsonPropertyName("by")] string By,
    [property: JsonPropertyName("label")] string? Label,
    [property: JsonPropertyName("open")] bool Open,
    [property: JsonPropertyName("accepted")] bool Accepted,
    [property: JsonPropertyName("entries")] ChangeEntryReply[] Entries);

/// <summary>
/// The change log's answer: which project it is about, where it is kept, the newest round marked
/// reviewed, and the rounds themselves - newest first, the one still running at the front.
///
/// `covers` says in words what the log records, because a log that is silently partial is worse
/// than no log at all.
/// </summary>
public sealed record ChangeLogReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("project")] string Project,
    [property: JsonPropertyName("acceptedAt")] int AcceptedAt,
    [property: JsonPropertyName("covers")] string Covers,
    [property: JsonPropertyName("rounds")] ChangeRoundReply[] Rounds,
    /// <summary>
    /// How many rounds the log holds, which is not always how many are in `rounds`.
    ///
    /// A view that stops at its limit and says nothing makes "the newest two hundred" and "all of
    /// them" the same answer. That matters to a reader deciding whether it has seen the whole
    /// history - a developer looking for an edit they remember, or an agent told to review what it
    /// changed - and this log's stance everywhere else is to say what it cannot show.
    /// </summary>
    [property: JsonPropertyName("total")] int Total = 0);

/// <summary>One text the change log kept, asked for by round and module.</summary>
public sealed record ChangeTextReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("round")] int Round,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("which")] string Which,
    [property: JsonPropertyName("held")] bool Held,
    [property: JsonPropertyName("text")] string? Text);

/// <summary>
/// One module's change, lined up side by side - the same rows the import/export dialog draws,
/// from the same bounded comparison, so the two surfaces cannot disagree about what changed.
/// </summary>
public sealed record ChangeDiffReply(
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("round")] int Round,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("held")] bool Held,
    [property: JsonPropertyName("rows")] Sync.SyncDiffRow[] Rows);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(ChangeEntryReply))]
[JsonSerializable(typeof(ChangeRoundReply))]
[JsonSerializable(typeof(ChangeLogReply))]
[JsonSerializable(typeof(ChangeTextReply))]
[JsonSerializable(typeof(ChangeDiffReply))]
internal sealed partial class ChangeJsonContext : JsonSerializerContext;
