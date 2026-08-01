using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.UI;

/// <summary>One finding as the panel renders it.</summary>
public sealed record FindingPayload(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column);

/// <summary>A batch of findings, replacing whatever the panel was showing.</summary>
public sealed record FindingsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("findings")] FindingPayload[] Findings);

/// <summary>A request from the panel to put the caret somewhere.</summary>
public sealed record NavigateMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column);

/// <summary>
/// Serialisation for panel messages. Source generated, because ahead-of-time compilation has no
/// reflection to fall back on.
/// </summary>
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(FindingsMessage))]
[JsonSerializable(typeof(FindingPayload))]
[JsonSerializable(typeof(NavigateMessage))]
public sealed partial class PanelJsonContext : JsonSerializerContext;
