using System.Text.Json;
using System.Text.Json.Serialization;

namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// The developer's choices, named key for key after the companion editor's own settings so the
/// two products describe the same behaviour in the same words. Everything has a default worth
/// shipping; a settings file is something a developer grows, never something they need.
/// </summary>
public sealed record ProductSettings
{
    /// <summary>
    /// Smart Enter and block-snippet layout: "comfy" opens spacer lines around the editable
    /// body, "compact" places the body directly above the closer.
    /// </summary>
    [JsonPropertyName("editor.blockLayout")]
    public string BlockLayout { get; init; } = "comfy";

    /// <summary>Enter at the end of a whole-line comment continues the apostrophes.</summary>
    [JsonPropertyName("editor.continueCommentOnNewline")]
    public bool ContinueCommentOnNewline { get; init; } = true;

    /// <summary>A continued comment also mirrors the spaces after the apostrophes.</summary>
    [JsonPropertyName("editor.mirrorCommentSpacing")]
    public bool MirrorCommentSpacing { get; init; } = true;

    /// <summary>Format Module: spaces per indent level.</summary>
    [JsonPropertyName("format.indentSize")]
    public int FormatIndentSize { get; init; } = 4;

    /// <summary>Format Module: indent with tabs rather than spaces.</summary>
    [JsonPropertyName("format.useTabs")]
    public bool FormatUseTabs { get; init; }

    /// <summary>Format Module: respell keywords in their canonical case.</summary>
    [JsonPropertyName("format.canonicalKeywords")]
    public bool FormatCanonicalKeywords { get; init; } = true;

    public static ProductSettings Default { get; } = new();

    /// <summary>The settings with every value forced into its legal range.</summary>
    public ProductSettings Normalized() => this with
    {
        BlockLayout = string.Equals(BlockLayout, "compact", StringComparison.OrdinalIgnoreCase)
            ? "compact"
            : "comfy",
        FormatIndentSize = Math.Clamp(FormatIndentSize, 1, 8),
    };

    /// <summary>
    /// Reads settings from their file's text. Anything unreadable — missing, empty, malformed,
    /// or holding the wrong shapes — answers the defaults, because a broken settings file must
    /// never be a broken editor.
    /// </summary>
    public static ProductSettings Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return Default;
        }

        try
        {
            var parsed = JsonSerializer.Deserialize(json, ProductSettingsContext.Default.ProductSettings);
            return parsed?.Normalized() ?? Default;
        }
        catch (JsonException)
        {
            return Default;
        }
    }

    /// <summary>The file's text for these settings, stable and readable for hand edits.</summary>
    public string ToJson() =>
        JsonSerializer.Serialize(Normalized(), ProductSettingsContext.Default.ProductSettings);
}

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(ProductSettings))]
public sealed partial class ProductSettingsContext : JsonSerializerContext;
