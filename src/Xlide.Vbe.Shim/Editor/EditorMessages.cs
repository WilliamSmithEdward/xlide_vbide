using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>Tells the surface which module to show and what its text is.</summary>
public sealed record LoadDocumentMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("text")] string Text);

/// <summary>One squiggle. Positions are one-based lines and columns, as the surface expects.</summary>
public sealed record EditorMarker(
    [property: JsonPropertyName("startLine")] int StartLine,
    [property: JsonPropertyName("startColumn")] int StartColumn,
    [property: JsonPropertyName("endLine")] int EndLine,
    [property: JsonPropertyName("endColumn")] int EndColumn,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("code")] string? Code);

/// <summary>Replaces every squiggle shown on the module.</summary>
public sealed record SetDiagnosticsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("markers")] EditorMarker[] Markers);

/// <summary>Chooses the surface's theme.</summary>
public sealed record SetThemeMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("theme")] string Theme);

/// <summary>Scrolls a line into view without moving the caret.</summary>
public sealed record RevealLineMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("line")] int Line);

/// <summary>
/// Serialisation for surface messages. Source generated, because ahead-of-time compilation has no
/// reflection to fall back on and a message type that is not registered here fails at run time
/// rather than at build time.
/// </summary>
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(LoadDocumentMessage))]
[JsonSerializable(typeof(SetDiagnosticsMessage))]
[JsonSerializable(typeof(SetThemeMessage))]
[JsonSerializable(typeof(RevealLineMessage))]
[JsonSerializable(typeof(EditorMarker))]
public sealed partial class EditorMessageContext : JsonSerializerContext;
