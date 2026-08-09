using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Sync;

/// <summary>One line of a side-by-side comparison. A number is absent where that side has no line.</summary>
public sealed record SyncDiffRow(
    [property: JsonPropertyName("leftNumber")] int? LeftNumber,
    [property: JsonPropertyName("rightNumber")] int? RightNumber,
    [property: JsonPropertyName("left")] string Left,
    [property: JsonPropertyName("right")] string Right,
    [property: JsonPropertyName("kind")] string Kind);

/// <summary>One row of an import or export plan, as the dialog draws it and the api reads it.</summary>
public sealed record SyncItemRow(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("file")] string File,
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("checked")] bool Checked,
    [property: JsonPropertyName("detail")] string Detail,
    [property: JsonPropertyName("warning")] string? Warning,
    [property: JsonPropertyName("inProject")] bool InProject,
    [property: JsonPropertyName("inFolder")] bool InFolder,
    [property: JsonPropertyName("cannotCreate")] bool CannotCreate,
    [property: JsonPropertyName("leftTitle")] string LeftTitle,
    [property: JsonPropertyName("rightTitle")] string RightTitle,
    [property: JsonPropertyName("diff")] SyncDiffRow[] Diff,
    [property: JsonPropertyName("diffWithHeaders")] SyncDiffRow[] DiffWithHeaders);

/// <summary>What an import or export would do, without doing any of it.</summary>
public sealed record SyncPlanReply(
    [property: JsonPropertyName("direction")] string Direction,
    [property: JsonPropertyName("project")] string Project,
    [property: JsonPropertyName("projectId")] string ProjectId,
    [property: JsonPropertyName("folder")] string Folder,
    [property: JsonPropertyName("mode")] string Mode,
    [property: JsonPropertyName("items")] SyncItemRow[] Items,
    [property: JsonPropertyName("warnings")] string[] Warnings);

/// <summary>What applying it actually did.</summary>
public sealed record SyncApplyReply(
    [property: JsonPropertyName("summary")] string Summary,
    [property: JsonPropertyName("changed")] string[] Changed,
    [property: JsonPropertyName("skipped")] string[] Skipped,
    [property: JsonPropertyName("removed")] string[] Removed,
    [property: JsonPropertyName("failed")] string[] Failed);

/// <summary>The folder and modes a project remembers.</summary>
public sealed record SyncSettingsReply(
    [property: JsonPropertyName("projectId")] string ProjectId,
    [property: JsonPropertyName("folder")] string Folder,
    [property: JsonPropertyName("exportMode")] string ExportMode,
    [property: JsonPropertyName("importMode")] string ImportMode);

/// <summary>Why the request could not be answered.</summary>
public sealed record SyncErrorReply([property: JsonPropertyName("error")] string Error);

/// <summary>
/// These shapes are the product's, not the dev door's.
///
/// They started out beside the debug api's other replies and could not stay there: that file
/// compiles only in Debug, and the import/export dialog ships in the release build. A release
/// carries these types and no server to serve them, which is the right way round.
/// </summary>
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(SyncDiffRow))]
[JsonSerializable(typeof(SyncItemRow))]
[JsonSerializable(typeof(SyncPlanReply))]
[JsonSerializable(typeof(SyncApplyReply))]
[JsonSerializable(typeof(SyncSettingsReply))]
[JsonSerializable(typeof(SyncErrorReply))]
internal sealed partial class SyncJsonContext : JsonSerializerContext;
