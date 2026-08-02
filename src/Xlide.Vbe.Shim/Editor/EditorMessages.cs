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

/// <summary>The modules the editor has open, and which one is showing.</summary>
public sealed record SetModulesMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("modules")] string[] Modules,
    [property: JsonPropertyName("active")] string? Active);

/// <summary>One finding as the surface's panel wants it.</summary>
public sealed record SurfaceFinding(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column);

/// <summary>Everything the panel lists, for every module.</summary>
public sealed record SetFindingsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("findings")] SurfaceFinding[] Findings);

/// <summary>One component in a project, with the kind the editor reports for it.</summary>
public sealed record SurfaceComponent(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] int Kind);

/// <summary>One project and everything in it.</summary>
public sealed record SurfaceProject(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("components")] SurfaceComponent[] Components);

/// <summary>The whole project tree, for the explorer.</summary>
public sealed record SetProjectsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("projects")] SurfaceProject[] Projects);

/// <summary>
/// The module's text as the editor now holds it, to be adopted without disturbing the developer.
///
/// Distinct from loading a document: loading replaces the model and resets the undo stack and the
/// caret, which is right when the developer opens a different module and wrong when the text they
/// are in the middle of typing has merely been normalised underneath them.
/// </summary>
public sealed record SyncDocumentMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("text")] string Text);

/// <summary>Asks the surface to run one of the editor's own commands.</summary>
public sealed record EditorCommandMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] string Id);

/// <summary>Something the developer should be told, shown briefly and not dwelt on.</summary>
public sealed record NoticeMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("text")] string Text);

/// <summary>One line of output for the Immediate panel.</summary>
public sealed record ImmediateResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("failed")] bool Failed);

/// <summary>The line execution is stopped on, or null when nothing is stopped.</summary>
public sealed record SetCurrentLineMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("line")] int? Line);

/// <summary>Every line in the shown module that carries a breakpoint.</summary>
public sealed record SetBreakpointsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("lines")] int[] Lines);

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
[JsonSerializable(typeof(SetModulesMessage))]
[JsonSerializable(typeof(SetFindingsMessage))]
[JsonSerializable(typeof(SetProjectsMessage))]
[JsonSerializable(typeof(SyncDocumentMessage))]
[JsonSerializable(typeof(EditorCommandMessage))]
[JsonSerializable(typeof(NoticeMessage))]
[JsonSerializable(typeof(ImmediateResultMessage))]
[JsonSerializable(typeof(SetCurrentLineMessage))]
[JsonSerializable(typeof(SetBreakpointsMessage))]
[JsonSerializable(typeof(SurfaceProject))]
[JsonSerializable(typeof(SurfaceComponent))]
[JsonSerializable(typeof(SurfaceFinding))]
[JsonSerializable(typeof(EditorMarker))]
public sealed partial class EditorMessageContext : JsonSerializerContext;
