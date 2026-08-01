using System.Text.Json;
using System.Text.Json.Serialization;

namespace Xlide.Vbe.Core.Engine;

/// <summary>One module as the add-in reads it out of the editor.</summary>
public sealed record EngineModule(
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("source")] string Source,
    [property: JsonPropertyName("type")] string Type);

/// <summary>A finding, with a span in character offsets into the module source.</summary>
public sealed record EngineDiagnostic(
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("span")] EngineSpan Span);

public sealed record EngineSpan(
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End);

public sealed record EngineDiagnostics(
    [property: JsonPropertyName("diagnostics")] EngineDiagnostic[] Diagnostics,
    [property: JsonPropertyName("mode")] string? Mode);

/// <summary>
/// Serialisation for the engine protocol.
///
/// Every type is registered on a source-generated context because the add-in is compiled ahead of
/// time, where reflection-based serialisation cannot be used at all. Adding a type to the protocol
/// means adding it here, and forgetting to shows up as a run-time failure rather than a build one.
/// </summary>
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(EngineModule))]
[JsonSerializable(typeof(EngineModule[]))]
[JsonSerializable(typeof(EngineDiagnostic))]
[JsonSerializable(typeof(EngineDiagnostics))]
[JsonSerializable(typeof(EngineSpan))]
[JsonSerializable(typeof(JsonElement))]
[JsonSerializable(typeof(Dictionary<string, object>))]
public sealed partial class EngineJsonContext : JsonSerializerContext;

/// <summary>
/// Converts between the editor's positions and the engine's.
///
/// The editor counts lines from one and gives no column at all for a whole line; the engine works
/// in character offsets into the module text. Converting in one place keeps the arithmetic from
/// being repeated, and repeated slightly differently, at every call site.
/// </summary>
public static class TextPositions
{
    /// <summary>
    /// Offsets at which each line begins. Built once per module version and reused, because
    /// converting a diagnostic without it costs a scan of the whole module per diagnostic.
    /// </summary>
    public static int[] LineStarts(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        var starts = new List<int>(Math.Max(16, text.Length / 40)) { 0 };

        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == '\n')
            {
                starts.Add(i + 1);
            }
        }

        return [.. starts];
    }

    /// <summary>
    /// Converts a character offset to a one-based line and column, which is what both the editor
    /// and the editing surface use.
    /// </summary>
    public static (int Line, int Column) ToLineColumn(int[] lineStarts, int offset)
    {
        ArgumentNullException.ThrowIfNull(lineStarts);

        if (lineStarts.Length == 0)
        {
            return (1, 1);
        }

        var low = 0;
        var high = lineStarts.Length - 1;

        while (low < high)
        {
            var middle = (low + high + 1) / 2;
            if (lineStarts[middle] <= offset)
            {
                low = middle;
            }
            else
            {
                high = middle - 1;
            }
        }

        return (low + 1, offset - lineStarts[low] + 1);
    }

    /// <summary>Converts a one-based line and column back to a character offset.</summary>
    public static int ToOffset(int[] lineStarts, int line, int column)
    {
        ArgumentNullException.ThrowIfNull(lineStarts);

        if (lineStarts.Length == 0)
        {
            return 0;
        }

        var index = Math.Clamp(line - 1, 0, lineStarts.Length - 1);
        return lineStarts[index] + Math.Max(0, column - 1);
    }
}
