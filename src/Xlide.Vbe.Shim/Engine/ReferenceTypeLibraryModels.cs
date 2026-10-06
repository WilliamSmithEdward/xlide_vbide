using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;

namespace Xlide.Vbe.Shim.Engine;

/// <summary>
/// Read referenced libraries that the analyzer does not already model. The Object Browser's COM
/// reader is the source of truth: no second hardcoded list of Scripting or RegExp members can drift
/// from the library actually installed on this machine. Cache only plain data; COM handles are
/// released after the first read and are never passed to the engine thread.
/// </summary>
internal static class ReferenceTypeLibraryModels
{
    private static readonly Dictionary<string, EngineReferenceLibrary> Cache = new(StringComparer.OrdinalIgnoreCase);

    private static readonly HashSet<string> ModelledGuids = new(StringComparer.OrdinalIgnoreCase)
    {
        "{00020813-0000-0000-C000-000000000046}", // Excel
        "{00020905-0000-0000-C000-000000000046}", // Word
        "{91493440-5A91-11CF-8700-00AA0060263B}", // PowerPoint
        "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", // Access
    };

    private static readonly HashSet<string> BuiltInNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "VBA", "stdole", "Office", "MSForms",
    };

    public static EngineReferenceLibrary[] For(ProjectReference[] references)
    {
        var found = new List<EngineReferenceLibrary>();
        foreach (var reference in references)
        {
            if (reference.Broken || reference.LibraryId.Length == 0 || reference.FullPath.Length == 0
                || ModelledGuids.Contains(reference.LibraryId) || BuiltInNames.Contains(reference.Name))
            {
                continue;
            }

            var key = $"{reference.LibraryId}\0{reference.Name}\0{reference.FullPath}";
            if (!Cache.TryGetValue(key, out var model))
            {
                try
                {
                    model = Read(reference);
                }
                catch (Exception ex)
                {
                    // One damaged third-party typelib must not prevent the project's modules
                    // and its other references from reaching the analyzer.
                    Log.Info($"typelib: '{reference.Name}' could not be read ({ex.GetType().Name})");
                    continue;
                }
                if (model is not null)
                {
                    Cache[key] = model;
                }
            }

            if (model is not null)
            {
                found.Add(model);
            }
        }

        return [.. found];
    }

    private static EngineReferenceLibrary? Read(ProjectReference reference)
    {
        using var catalog = new TypeLibraryCatalog();
        catalog.AddLibrary(reference.Name, reference.Name, reference.FullPath);
        var types = catalog.TypesOf(reference.Name);
        if (types is null || types.Count == 0)
        {
            return null;
        }

        // A malformed or enormous third-party library must not make every edit wait for a huge
        // project/open payload. Refuse the whole model rather than asserting absence from a cut.
        if (types.Count > 512)
        {
            Log.Info($"typelib: '{reference.Name}' exceeds the 512-type analysis limit");
            return null;
        }

        var rows = new List<EngineReferenceType>(types.Count);
        var totalMembers = 0;
        foreach (var type in types)
        {
            var members = catalog.MembersOf(reference.Name, type.Name);
            if (members is null || members.Count > 512 || (totalMembers += members.Count) > 20_000)
            {
                Log.Info($"typelib: '{reference.Name}' exceeds the analysis member limit");
                return null;
            }

            rows.Add(new EngineReferenceType(type.Name, type.Kind,
                [.. members.Select(member => new EngineReferenceMember(
                    member.Name, member.Kind, member.Signature))]));
        }

        Log.Info($"typelib: '{reference.Name}' supplied {rows.Count} types and {totalMembers} members to analysis");
        return new EngineReferenceLibrary(reference.Name, reference.LibraryId, [.. rows]);
    }
}
