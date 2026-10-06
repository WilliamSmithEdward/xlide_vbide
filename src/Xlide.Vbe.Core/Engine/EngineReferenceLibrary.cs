using System.Text.Json.Serialization;

namespace Xlide.Vbe.Core.Engine;

/// <summary>Plain, bounded facts from a referenced COM type library.</summary>
public sealed record EngineReferenceMember(string Name, string Kind, string Signature);
public sealed record EngineReferenceType(string Name, string Kind, EngineReferenceMember[] Members);
public sealed record EngineReferenceLibrary(string Name, [property: JsonPropertyName("guid")] string LibraryId, EngineReferenceType[] Types);
