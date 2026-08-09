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

    /// <summary>
    /// The explorer's tree follows the editor: the module being worked on unfolds its procedures,
    /// and everything folds away when the last tab closes. Off leaves the tree entirely to the
    /// hand that opened it.
    /// </summary>
    [JsonPropertyName("explorer.treeFollowsEditor")]
    public bool TreeFollowsEditor { get; init; } = true;

    /// <summary>
    /// One indent level, in spaces. Governs the EDITOR's own indentation, what smart Enter
    /// leaves behind, and Format Module — which used to be three different behaviours from one
    /// name (2026-08-08).
    /// </summary>
    [JsonPropertyName("format.indentSize")]
    public int FormatIndentSize { get; init; } = 4;

    // THERE IS NO "indent with tabs" SETTING, and there cannot be a working one.
    //
    // VBA's code store will not hold a tab character. The editor expands every one it is handed
    // to the next four-column stop, on both write paths this product uses and mid-line as well as
    // leading: "    Dim n\tAs Long" is read back as "    Dim n   As Long" (measured 2026-08-07).
    //
    // So the setting could only ever have been half honoured. The page indented with tabs, the
    // workbook held spaces, and the two disagreed for as long as the module stayed open, which is
    // the one thing a surface covering the host's own editor must never do. It was removed on the
    // developer's call the day it was measured; a `format.useTabs` left in an older settings file
    // is ignored rather than read.
    //
    // Indentation is spaces, FormatIndentSize of them, and Backspace in a line's leading
    // whitespace takes back a whole level of them.

    /// <summary>Format Module: respell keywords in their canonical case.</summary>
    [JsonPropertyName("format.canonicalKeywords")]
    public bool FormatCanonicalKeywords { get; init; } = true;

    /// <summary>
    /// Which code decides what an import or export will do: "xlide" or "builtIn".
    ///
    /// "xlide" runs the companion editor's own planner inside the engine, so both products make
    /// identical decisions about file names, module kinds and what counts as a change, and a fix
    /// to it fixes both. It is the default for that reason. "builtIn" works it out inside the
    /// add-in, which needs no engine and keeps import and export working when the engine is not
    /// running.
    /// </summary>
    [JsonPropertyName("sync.engine")]
    public string SyncEngine { get; init; } = "xlide";

    public static ProductSettings Default { get; } = new();

    /// <summary>The settings with every value forced into its legal range.</summary>
    public ProductSettings Normalized() => this with
    {
        BlockLayout = string.Equals(BlockLayout, "compact", StringComparison.OrdinalIgnoreCase)
            ? "compact"
            : "comfy",
        FormatIndentSize = Math.Clamp(FormatIndentSize, 1, 8),
        SyncEngine = string.Equals(SyncEngine, "builtIn", StringComparison.OrdinalIgnoreCase)
            ? "builtIn"
            : "xlide",
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
