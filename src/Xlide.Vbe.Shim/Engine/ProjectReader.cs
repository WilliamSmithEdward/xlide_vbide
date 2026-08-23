using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Core.Vba;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Engine;

/// <summary>Everything the engine needs to know about one project at one moment.</summary>
/// <param name="ProjectId">Stable identity of the project within this session.</param>
/// <param name="DisplayName">What the developer calls it: the workbook's file name.</param>
/// <param name="Generation">Increases whenever any module's text changes.</param>
/// <param name="Modules">Every module, with its current text.</param>
internal sealed record ProjectSnapshot(string ProjectId, string DisplayName, int Generation, EngineModule[] Modules);

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
    public static string TypeName(int componentType) => componentType switch
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

    /// <summary>
    /// The identity a project is addressed by, and the name the developer knows it by.
    ///
    /// The project's own Name is "VBAProject" for nearly every workbook, so addressing by Name
    /// made two open workbooks the same project everywhere downstream: the engine merged their
    /// modules, findings crossed workbooks, and a write aimed at whichever came first. A saved
    /// workbook's full file path is unique among the open ones and stable for the session, so
    /// that is the identity.
    ///
    /// TWO UNSAVED WORKBOOKS HAVE NEITHER. Both answer "VBAProject" to Name and both raise when
    /// asked for a FileName, so both used to land on the same id, and everything keyed by it
    /// collided: `liveKey(projectId, moduleName)` made one workbook's Sheet1 the other's, the
    /// engine seeded one over the other, and the skip compared the wrong workbook's sources.
    /// Reported from a live session with two new workbooks open (2026-08-08).
    ///
    /// So an unsaved project is identified by its own COM identity, which is what COM means by
    /// identity: the canonical IUnknown for the object. It is unique among the projects alive at
    /// any moment and stable for as long as the project is, which is exactly the lifetime the id
    /// has to cover. It changes if the workbook is closed and reopened, and that is correct:
    /// that is a different project.
    /// </summary>
    public static (string Id, string DisplayName) Identity(DispatchObject project)
    {
        var name = project.GetString("Name") ?? "VBAProject";

        try
        {
            if (project.GetString("FileName") is { Length: > 0 } fileName)
            {
                return (fileName.ToLowerInvariant(), Path.GetFileName(fileName));
            }
        }
        catch (Exception)
        {
            // Unsaved: the property raises rather than answering empty.
        }

        // Unsaved. The name alone is not an identity, so it is qualified by the object's own.
        var identity = ComIdentityOf(project);
        return identity == 0
            ? (name, name)
            : ($"{name.ToLowerInvariant()}#{identity:x}", name);
    }

    /// <summary>
    /// The class modules each project held last pass, so one that APPEARS can be spotted.
    ///
    /// A class that has just arrived may not be the class the saved file describes under that
    /// name, and the saved file is where its default-instance flag is read from. The sync's own
    /// import says so directly (<see cref="Core.Vba.SavedModules.Doubt"/>), but the editor's own
    /// File / Import File goes nowhere near this product - so the appearance is noticed here
    /// instead, where every component is already walked on every pass, and every route in is
    /// covered by one rule rather than by remembering to call one thing from each of them.
    /// </summary>
    private static readonly Dictionary<string, HashSet<string>> ClassesLastPass =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Notes the class modules this pass found, and doubts the ones that were not here before.
    ///
    /// The FIRST pass for a project doubts nothing: everything is new to it, and doubting the
    /// whole project would mean the flag never answered at all. What it is watching for is a name
    /// arriving into a project already being watched.
    /// </summary>
    private static void DoubtClassesThatJustArrived(string? savedPath, HashSet<string> classes)
    {
        if (savedPath is null)
        {
            // Never saved, so there is no saved answer to be wrong, and nothing to doubt.
            return;
        }

        if (ClassesLastPass.TryGetValue(savedPath, out var before))
        {
            foreach (var name in classes)
            {
                if (!before.Contains(name))
                {
                    Log.Info($"seed: class {name} was not here last pass, so the saved file's "
                        + "header is not trusted for it until the document is saved again");
                    Core.Vba.SavedModules.Doubt(savedPath, name);
                }
            }
        }

        ClassesLastPass[savedPath] = classes;
    }

    /// <summary>
    /// Puts each class module's default-instance flag onto the modules the walk built.
    ///
    /// FOR CLASS MODULES AND NOTHING ELSE. The analyzer consults the flag only for `class`,
    /// because a document module and a form have a default instance by their kind and a standard
    /// module cannot have one at all - so answering for the others would only invite reading it
    /// as a claim about them.
    ///
    /// The answer comes from the SAVED document, because the editor has no property for it and
    /// would only give it up through an export (xlide_vscode#50). So a class created since the
    /// last save answers null, and null is the state that keeps the analyzer quiet - which is the
    /// right way round: the loud state is a false `Variable not defined` on correct code.
    /// </summary>
    private static void FillInPredeclaredIds(
        DispatchObject project, List<EngineModule> modules, HashSet<string> classNames)
    {
        var savedPath = SavedPathOf(project);

        // Recorded even when the set is empty, so that removing a project's last class and
        // importing it back is still seen as an arrival rather than as nothing having happened.
        DoubtClassesThatJustArrived(savedPath, classNames);

        if (classNames.Count == 0)
        {
            // No class module, so nothing asks the question, and a project without one does not
            // pay to have it answered.
            return;
        }

        var saved = SavedModules.For(savedPath);

        for (var at = 0; at < modules.Count; at++)
        {
            var module = modules[at];
            if (module.Type != "class")
            {
                continue;
            }

            var predeclaredId = saved?.PredeclaredIdOf(module.ModuleName);
            Log.Verbose($"seed: class {module.ModuleName} is "
                + predeclaredId switch
                {
                    true => "predeclared (its bare name is a value)",
                    false => "not predeclared (its bare name is a type)",
                    null => "unanswered - the saved file cannot vouch for it",
                });

            modules[at] = module with { PredeclaredId = predeclaredId };
        }
    }

    /// <summary>
    /// Where the project's document sits on disk, or null when it has never been saved.
    ///
    /// Not <see cref="Identity"/>'s id, which is lowercased for comparison and falls back to a
    /// name plus a COM pointer when there is no file. This wants the real path or nothing.
    /// </summary>
    public static string? SavedPathOf(DispatchObject project)
    {
        try
        {
            return project.GetString("FileName") is { Length: > 0 } fileName ? fileName : null;
        }
        catch (Exception)
        {
            // Unsaved: the property raises rather than answering empty.
            return null;
        }
    }

    /// <summary>
    /// The canonical IUnknown pointer for an object, which is COM's own definition of identity:
    /// two references to the same object always answer the same value, and two different objects
    /// never do. Zero when it cannot be asked, in which case the caller falls back to the name.
    ///
    /// Queried rather than read off the dispatch pointer we hold, because a QueryInterface for a
    /// different interface may hand back a different pointer for the same object. Only IUnknown
    /// is promised to be canonical.
    /// </summary>
    private static nint ComIdentityOf(DispatchObject project)
    {
        var pointer = project.Pointer;
        if (pointer == 0)
        {
            return 0;
        }

        var iid = new Guid("00000000-0000-0000-C000-000000000046");
        if (System.Runtime.InteropServices.Marshal.QueryInterface(pointer, in iid, out var unknown) < 0)
        {
            return 0;
        }

        // Released at once: the pointer VALUE is the identity, and holding a reference to read a
        // number would be a reference nothing gives back.
        System.Runtime.InteropServices.Marshal.Release(unknown);
        return unknown;
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

            // Which components are classes, gathered as the walk goes. The default-instance flag
            // cannot be filled in DURING the walk: whether the saved file may be trusted for a
            // class depends on whether that class was here last pass, and that is only knowable
            // once every component has been seen. So the walk collects, and the flags go on after.
            var classNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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

                var componentType = component.GetInt32("Type");

                // A form's controls ride the seed as implicit members: declared by the designer
                // where the analyzer cannot see, supplied by the one side that can. This is the
                // read the analyzer's #17 fix exists to receive, and it costs a designer walk
                // only for form components, on the host thread this reader already owns.
                var members = componentType == 3
                    ? Editor.FormDesignService.ControlMembers(component)
                    : null;

                if (componentType == 3)
                {
                    // Null and empty are DIFFERENT seeds - unreadable vs vouched-for-empty
                    // (xlide_vscode#26) - and nothing else says which one a pass sent, which
                    // is exactly what the 2026-08-19 ghost-control hunt needed to see.
                    Diagnostics.Log.Verbose($"seed: {moduleName} carries "
                        + (members is null ? "no designer answer" : $"{members.Length} control(s)"
                            + (members.Length == 0 ? string.Empty : $": {string.Join(", ", members.Select(m => m.Name))}")));
                }

                if (componentType == 2)
                {
                    classNames.Add(moduleName);
                }

                modules.Add(new EngineModule(moduleName, source, TypeName(componentType), members));
            }

            FillInPredeclaredIds(project, modules, classNames);

            var (id, displayName) = Identity(project);
            return new ProjectSnapshot(id, displayName, generation, [.. modules]);
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
