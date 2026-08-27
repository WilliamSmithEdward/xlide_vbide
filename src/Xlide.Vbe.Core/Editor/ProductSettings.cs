using System.Text.Json;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Text.Json.Serialization;

namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// The developer's choices, named key for key after the companion editor's own settings so the
/// two products describe the same behaviour in the same words. Everything has a default worth
/// shipping; a settings file is something a developer grows, never something they need.
/// </summary>
public sealed record ProductSettings
{
    /*
     * THE PROPERTIES ARE `set`, NOT `init`, AND THAT IS LOAD-BEARING.
     *
     * An init-only property can only be assigned in an object initializer, so the JSON source
     * generator has to set every one of them at construction - passing `default` for each key
     * the document does not mention. That silently clobbers the initializers below: a file
     * naming only `format.indentSize` produced FALSE for every boolean in this record, because
     * `default(bool)` is false and the "= true" beside it never survived.
     *
     * Nobody saw it for as long as every key was in every file. It surfaced the moment a
     * setting was ADDED - the designer grid, 2026-08-16 - which read back as off, at a
     * two-point spacing, on a machine whose settings file predated it by an hour. Every
     * developer with an older file was in that case, for every boolean.
     *
     * With settable properties the generator constructs first and assigns only what the
     * document names, so an absent key keeps the value written here. A record with setters is
     * a small ugliness; a settings type that cannot be added to is a larger one.
     */

    /// <summary>
    /// Smart Enter and block-snippet layout: "comfy" opens spacer lines around the editable
    /// body, "compact" places the body directly above the closer.
    /// </summary>
    [JsonPropertyName("editor.blockLayout")]
    public string BlockLayout { get; set; } = "comfy";

    /// <summary>Enter at the end of a whole-line comment continues the apostrophes.</summary>
    [JsonPropertyName("editor.continueCommentOnNewline")]
    public bool ContinueCommentOnNewline { get; set; } = true;

    /// <summary>A continued comment also mirrors the spaces after the apostrophes.</summary>
    [JsonPropertyName("editor.mirrorCommentSpacing")]
    public bool MirrorCommentSpacing { get; set; } = true;

    /// <summary>
    /// The explorer's tree follows the editor: the module being worked on unfolds its procedures,
    /// and everything folds away when the last tab closes. Off leaves the tree entirely to the
    /// hand that opened it.
    /// </summary>
    [JsonPropertyName("explorer.treeFollowsEditor")]
    public bool TreeFollowsEditor { get; set; } = true;

    /// <summary>
    /// One indent level, in spaces. Governs the EDITOR's own indentation, what smart Enter
    /// leaves behind, and Format Module - which used to be three different behaviours from one
    /// name (2026-08-08).
    /// </summary>
    [JsonPropertyName("format.indentSize")]
    public int FormatIndentSize { get; set; } = 4;

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

    // THERE IS NO "canonical keywords" SETTING, and there cannot be a working one. Two paths
    // canonicalise keywords before the formatter is ever asked and neither consults a setting: the
    // HOST respells them as it takes a module, and the page recases every touched line 200ms after
    // it settles. Typing `public sub go()` with the switch OFF still produced `Public Sub go()`
    // (2026-08-09). Formatting still respells, always; the switch promised what it could not
    // deliver. A `format.canonicalKeywords` in an older settings file is ignored rather than read.

    /// <summary>
    /// What the form designer's canvas snaps POINTER gestures to: "grid", "objects" or "off".
    ///
    /// One or the other, never both. They collide when they disagree - a control near a
    /// neighbour's edge and near a grid line has two right answers and takes whichever the code
    /// asked first, which reads as a designer that cannot make up its mind (the owner,
    /// 2026-08-16: "looks like snap to grid mode, and snap to object mode collide... should we
    /// have it be either or (or neither)").
    ///
    /// "grid" is the editor's own Align Controls to Grid and the default for that reason.
    /// "objects" is the design-surface behaviour: edges and centres of the neighbours sharing a
    /// container, with a guide drawn where the gesture lines up.
    ///
    /// The KEYBOARD is never snapped, whichever this says. Arrows and Shift+arrows move by a
    /// single point, because the hand that reaches for them has already decided that neither the
    /// grid nor the neighbours are where they want the thing - and a nudge that jumps six points
    /// is not a nudge. Holding ALT escapes whichever mode is on, for one gesture.
    /// </summary>
    [JsonPropertyName("designer.snap")]
    public string DesignerSnap { get; set; } = "grid";

    /// <summary>The grid's spacing in POINTS, the designer's own unit. Six, as the editor's
    /// Options dialog ships it. Read only in "grid" mode.</summary>
    [JsonPropertyName("designer.gridSize")]
    public int DesignerGridSize { get; set; } = 6;

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
    public string SyncEngine { get; set; } = "xlide";

    /// <summary>
    /// Machine-wide analyzer rule severity overrides, keyed by the rule's stable diagnostic
    /// code - the same key and vocabulary as the companion editor's
    /// `xlide.analysis.ruleSeverityOverrides`, so a team can carry one convention across both.
    ///
    /// The values are GUARDED BY THE ANALYZER, not by this record: a warning or information
    /// rule may be set to "off"; a rule whose default severity is error may at most be
    /// downgraded to "warning", and only where the analyzer marks that safe. An entry the
    /// analyzer does not allow is carried in the file but changes nothing - the guard lives in
    /// one place, upstream, and this record does not restate it because a restatement is a
    /// second copy that drifts.
    ///
    /// Machine-wide means this user on this machine: the file lives in user space and every
    /// workbook and session reads it. Suppressing one finding at one line is a different tool -
    /// the inline `' @xlide-analysis-disable-*` comments, which travel with the code.
    /// </summary>
    [JsonPropertyName("analysis.ruleSeverityOverrides")]
    public Dictionary<string, string>? AnalysisRuleSeverityOverrides { get; set; }

    /// <summary>
    /// The xlide api: the local, token-gated door an agent drives this session through.
    ///
    /// THREE STATES, AND THE THIRD ONE IS THE POINT. Null means "whatever this build defaults
    /// to" - on in a dev build, where the harness IS the reason the build exists, and off in a
    /// shipped one, where a door nobody asked for should not be listening. True and false are a
    /// developer having answered the question, and their answer holds in either build.
    ///
    /// It is a capability rather than a preference, which is why it is not in the settings
    /// dialog: it is turned on from the agent card, next to the address it hands out.
    /// </summary>
    [JsonPropertyName("api.enabled")]
    public bool? ApiEnabled { get; set; }

    public static ProductSettings Default { get; } = new();

    /// <summary>The settings with every value forced into its legal range.</summary>
    public ProductSettings Normalized() => this with
    {
        // Lower-cased and sorted so the file is stable under hand edits and diffs read clean.
        // Unknown codes and values survive normalization on purpose: validation against the
        // catalog is the analyzer's, and a file written by a newer build must not be trimmed by
        // an older one reading it.
        AnalysisRuleSeverityOverrides = AnalysisRuleSeverityOverrides is { Count: > 0 } overrides
            ? overrides
                .Where(entry => !string.IsNullOrWhiteSpace(entry.Key)
                    && !string.IsNullOrWhiteSpace(entry.Value))
                .GroupBy(entry => entry.Key.Trim().ToLowerInvariant(), StringComparer.Ordinal)
                .OrderBy(group => group.Key, StringComparer.Ordinal)
                .ToDictionary(
                    group => group.Key,
                    group => group.Last().Value.Trim().ToLowerInvariant(),
                    StringComparer.Ordinal)
            : null,
        BlockLayout = string.Equals(BlockLayout, "compact", StringComparison.OrdinalIgnoreCase)
            ? "compact"
            : "comfy",
        FormatIndentSize = Math.Clamp(FormatIndentSize, 1, 8),
        DesignerGridSize = Math.Clamp(DesignerGridSize, 2, 24),
        DesignerSnap = DesignerSnap?.ToLowerInvariant() switch
        {
            "objects" => "objects",
            "off" => "off",
            _ => "grid",
        },
        SyncEngine = string.Equals(SyncEngine, "builtIn", StringComparison.OrdinalIgnoreCase)
            ? "builtIn"
            : "xlide",
    };

    /// <summary>
    /// Reads settings from their file's text. Anything unreadable - missing, empty, malformed,
    /// or holding the wrong shapes - answers the defaults, because a broken settings file must
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
