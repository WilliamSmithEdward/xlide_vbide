using System.Text.Json;
using System.Text.Json.Serialization;

namespace Xlide.Vbe.Core.Sync;

/// <summary>What one project remembers about importing and exporting.</summary>
public sealed record SyncChoice
{
    /// <summary>The folder last used, or empty when the project has never been synced.</summary>
    public string Folder { get; init; } = string.Empty;

    /// <summary>exportAll or trueUp.</summary>
    public string ExportMode { get; init; } = "exportAll";

    /// <summary>updateOnly or trueUpStandardClass.</summary>
    public string ImportMode { get; init; } = "updateOnly";
}

/// <summary>
/// The folder and modes each project was last synced with, so the dialog opens where the developer
/// left it rather than asking again every time.
///
/// Kept beside the product's settings and separate from them: this grows a row per workbook a
/// developer has ever synced, and settings.json is meant to stay small enough to read by hand.
/// A project is keyed by the same id everything else uses, which for a saved workbook is its path
/// and for an unsaved one is its COM identity, so an unsaved workbook remembers nothing across a
/// restart, which is right, because it is a different project by then.
/// </summary>
public sealed record SyncSettings
{
    public Dictionary<string, SyncChoice> Projects { get; init; } = [];

    public static SyncSettings Empty { get; } = new();

    /// <summary>What this project remembers, or the defaults if it remembers nothing.</summary>
    public SyncChoice For(string projectId) =>
        Projects.TryGetValue(projectId, out var choice) ? choice : new SyncChoice();

    /// <summary>The same settings with one project's choice replaced.</summary>
    public SyncSettings With(string projectId, SyncChoice choice)
    {
        var projects = new Dictionary<string, SyncChoice>(Projects, StringComparer.OrdinalIgnoreCase)
        {
            [projectId] = choice,
        };
        return new SyncSettings { Projects = projects };
    }

    /// <summary>Reads the file. Anything unreadable answers empty, because a bad file is not a bad editor.</summary>
    public static SyncSettings Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return Empty;
        }

        try
        {
            return JsonSerializer.Deserialize(json, SyncSettingsContext.Default.SyncSettings) ?? Empty;
        }
        catch (JsonException)
        {
            return Empty;
        }
    }

    public string ToJson() => JsonSerializer.Serialize(this, SyncSettingsContext.Default.SyncSettings);
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(SyncSettings))]
public sealed partial class SyncSettingsContext : JsonSerializerContext;
