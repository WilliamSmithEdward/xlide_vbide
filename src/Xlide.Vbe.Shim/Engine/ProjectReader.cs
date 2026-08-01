using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Engine;

/// <summary>Everything the engine needs to know about one project at one moment.</summary>
/// <param name="ProjectId">Stable identity of the project within this session.</param>
/// <param name="Generation">Increases whenever any module's text changes.</param>
/// <param name="Modules">Every module, with its current text.</param>
internal sealed record ProjectSnapshot(string ProjectId, int Generation, EngineModule[] Modules);

/// <summary>
/// Reads module sources out of the editor.
///
/// Every call here crosses into the host, so this runs on the host user interface thread and is
/// called deliberately rather than often. Reading a large project is measured in the tens of
/// milliseconds, which is fine when a project opens and not fine per keystroke: keystroke updates
/// send one module's text, not a snapshot.
/// </summary>
internal static class ProjectReader
{
    /// <summary>
    /// Component type as the analyzer names it. The editor reports a numeric kind, and the
    /// distinction matters: a document module has an implicit instance and different rules apply to
    /// it than to a class the user news up.
    /// </summary>
    private static string TypeName(int componentType) => componentType switch
    {
        1 => "standard",
        2 => "class",
        3 => "userform",
        11 or 100 => "document",
        _ => "standard",
    };

    /// <summary>Reads every module of every project the editor has open.</summary>
    public static List<ProjectSnapshot> ReadAll(DispatchObject editor, int generation)
    {
        var snapshots = new List<ProjectSnapshot>();

        using var projects = editor.GetObject("VBProjects");
        if (projects is null)
        {
            return snapshots;
        }

        var projectCount = projects.GetInt32("Count");

        for (var i = 1; i <= projectCount; i++)
        {
            using var project = projects.GetItem(i);
            if (project is null)
            {
                continue;
            }

            var snapshot = Read(project, generation);
            if (snapshot is not null)
            {
                snapshots.Add(snapshot);
            }
        }

        return snapshots;
    }

    /// <summary>Reads one project.</summary>
    public static ProjectSnapshot? Read(DispatchObject project, int generation)
    {
        try
        {
            var name = project.GetString("Name");
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            using var components = project.GetObject("VBComponents");
            if (components is null)
            {
                return null;
            }

            var count = components.GetInt32("Count");
            var modules = new List<EngineModule>(count);

            for (var i = 1; i <= count; i++)
            {
                using var component = components.GetItem(i);
                if (component is null)
                {
                    continue;
                }

                var moduleName = component.GetString("Name");
                if (string.IsNullOrEmpty(moduleName))
                {
                    continue;
                }

                var source = ReadSource(component);
                if (source is null)
                {
                    continue;
                }

                modules.Add(new EngineModule(moduleName, source, TypeName(component.GetInt32("Type"))));
            }

            return new ProjectSnapshot(name, generation, [.. modules]);
        }
        catch (Exception ex)
        {
            Log.Error("project: could not be read", ex);
            return null;
        }
    }

    /// <summary>
    /// Reads a component's whole source. An empty module has no lines at all, and asking such a
    /// module for line one raises rather than returning nothing.
    /// </summary>
    public static string? ReadSource(DispatchObject component)
    {
        try
        {
            using var code = component.GetObject("CodeModule");
            if (code is null)
            {
                return null;
            }

            var lines = code.GetInt32("CountOfLines");
            return lines <= 0 ? string.Empty : code.GetStringIndexed("Lines", 1, lines) ?? string.Empty;
        }
        catch (Exception ex)
        {
            Log.Error("project: a module's source could not be read", ex);
            return null;
        }
    }
}
