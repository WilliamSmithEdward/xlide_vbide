using System.Text;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Engine;

namespace Xlide.Vbe.Shim.Sync;

/// <summary>What applying a plan actually did.</summary>
/// <param name="Changed">Files written, or modules created and replaced.</param>
/// <param name="Skipped">Rows that were selected and did nothing, with the reason.</param>
/// <param name="Removed">Files deleted, or modules deleted.</param>
/// <param name="Failed">Rows that were meant to do something and could not, with the reason.</param>
internal sealed record SyncApplyResult(
    IReadOnlyList<string> Changed,
    IReadOnlyList<string> Skipped,
    IReadOnlyList<string> Removed,
    IReadOnlyList<string> Failed);

/// <summary>
/// The half of import/export that touches the editor and the disk.
///
/// The planning is in <see cref="ModuleSync"/> and has no idea either exists, which is what lets
/// the dialog and the debug api ask the same question and get the same answer. This half reads what
/// the project holds, reads what the folder holds, and carries out the rows a caller selected.
///
/// ENCODING is the trap in all of this, and the reason the source is assembled rather than copied.
/// The editor's own Export writes the file in the machine's ANSI code page, so a module holding
/// Japanese or Cyrillic comes back as nonsense on a machine whose code page cannot spell it, and
/// the companion editor writes UTF-8. Rather than convert between them and hope:
///
///   - The HEADER (VERSION, BEGIN/END, the Attribute lines) comes from a temporary Export. It is
///     the only way to learn VB_PredeclaredId, VB_Exposed and VB_Base, and it is ASCII, so it is
///     read as Latin-1, which cannot fail and leaves every byte alone.
///   - The BODY comes from the code module over COM, as a string. COM carries text as UTF-16, so
///     there is no code page in the path at all and nothing to lose.
///   - The two are spliced, and the name in the header is rewritten from the name the editor gave
///     us, so a module named in a script the code page cannot spell still names itself correctly.
///
/// The same reasoning runs backwards on import: a new module is created by importing a header-only
/// file, which gets the kind and the attributes exactly right, and the body is then written through
/// COM. Nothing non-ASCII is ever handed to the editor through a file.
/// </summary>
internal static class ModuleSyncService
{
    /// <summary>How the shim writes a module's text, which is the path that also tells the surface
    /// and the analyzer about it. Import uses it so an imported module lands everywhere at once.
    ///
    /// Null when the module took it. Anything else is what the editor said, and a row that gets one
    /// is a row that FAILED, however tidily it was asked for.</summary>
    internal delegate string? WriteModuleText(string component, string text, string? projectId);

    /// <summary>Files are written as UTF-8 with no mark, which is what the companion editor writes.</summary>
    private static readonly UTF8Encoding FileEncoding = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>Reads every module of a project, with the full text a file needs.</summary>
    public static List<LiveModule> ReadLiveModules(DispatchObject project)
    {
        var modules = new List<LiveModule>();

        using var components = project.GetObject("VBComponents");
        if (components is null)
        {
            return modules;
        }

        var count = components.GetInt32("Count");
        for (var i = 1; i <= count; i++)
        {
            using var component = components.GetItem(i);
            if (component is null)
            {
                continue;
            }

            var name = component.GetString("Name");
            if (string.IsNullOrEmpty(name))
            {
                continue;
            }

            var kind = ProjectReader.TypeName(component.GetInt32("Type"));
            var body = ProjectReader.ReadSource(component) ?? string.Empty;
            modules.Add(new LiveModule(name, kind, FullSourceFor(component, name, kind, body)));
        }

        return modules;
    }

    /// <summary>Reads the module files a folder holds. A file that will not read is reported, not dropped.</summary>
    public static List<RepoFile> ReadFolder(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return [];
        }

        // READ AT ONCE, in a fixed order.
        //
        // A folder holds one file per module and each read waits on the disk rather than on the
        // processor, so reading them one after another spends the whole wait N times. Ordered so
        // the plan reads the same twice running - the rows are sorted by status afterwards, and
        // the order within a status is this one.
        //
        // A file that will not read becomes a row saying so rather than an exception: one locked
        // file must not cost the developer the other twenty.
        return Directory.EnumerateFiles(folder)
            .Select(Path.GetFileName)
            .Where(name => name is not null && ModuleSync.IsModuleFileName(name))
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .AsParallel()
            .AsOrdered()
            .Select(name =>
            {
                try
                {
                    return new RepoFile(name!, File.ReadAllText(Path.Combine(folder, name!), FileEncoding));
                }
                catch (Exception ex)
                {
                    return new RepoFile(name!, string.Empty, ex.Message);
                }
            })
            .ToList();
    }

    /// <summary>Carries out the rows a caller selected, and says what happened to each.</summary>
    public static SyncApplyResult Apply(
        DispatchObject project,
        SyncPlan plan,
        IReadOnlySet<string> selected,
        WriteModuleText write)
    {
        var changed = new List<string>();
        var skipped = new List<string>();
        var removed = new List<string>();
        var failed = new List<string>();

        foreach (var item in plan.Items.Where(item => selected.Contains(item.Id)))
        {
            try
            {
                switch (item.Status)
                {
                    case SyncStatus.Unchanged:
                        skipped.Add($"{item.FileName} (already the same)");
                        break;

                    case SyncStatus.SkippingImport:
                        skipped.Add($"{item.FileName} (a {item.ModuleKind} cannot be created from a file)");
                        break;

                    case SyncStatus.ReadError:
                        skipped.Add($"{item.FileName} ({item.Warning})");
                        break;

                    case SyncStatus.WillRemove when plan.Direction == SyncDirection.Export:
                        if (DeleteFile(plan.Folder, item.FileName))
                        {
                            removed.Add(item.FileName);
                        }
                        else
                        {
                            skipped.Add($"{item.FileName} (already gone)");
                        }

                        break;

                    case SyncStatus.WillRemove:
                        RemoveComponent(project, item.ModuleName);
                        removed.Add(item.ModuleName);
                        break;

                    // IMPORT COUNTS A ROW CHANGED ONLY IF THE MODULE TOOK IT.
                    //
                    // The write's complaint used to be discarded here, so a module the editor
                    // refused was reported to the developer as imported. That is the worst place in
                    // the product for a false success: they came here to move code between a
                    // repository and a workbook, and the whole point is knowing which end has what.
                    case SyncStatus.WillCreate when plan.Direction == SyncDirection.Import:
                        if (CreateComponent(project, item, write, plan.ProjectId) is { } createRefused)
                        {
                            failed.Add($"{item.ModuleName} ({createRefused})");
                        }
                        else
                        {
                            changed.Add(item.ModuleName);
                        }

                        break;

                    case SyncStatus.WillUpdate:
                        if (write(item.ModuleName, ModuleSync.CodeWithoutHeader(item.PayloadSource), plan.ProjectId)
                            is { } updateRefused)
                        {
                            failed.Add($"{item.ModuleName} ({updateRefused})");
                        }
                        else
                        {
                            changed.Add(item.ModuleName);
                        }

                        break;

                    case SyncStatus.WillCreate:
                    case SyncStatus.WillWrite:
                        WriteFile(plan.Folder, item.FileName, item.PayloadSource);
                        changed.Add(item.FileName);
                        break;

                    default:
                        skipped.Add(item.FileName);
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Error($"sync: {item.Id} could not be applied", ex);
                failed.Add($"{item.FileName} ({ex.Message})");
            }
        }

        return new SyncApplyResult(changed, skipped, removed, failed);
    }

    /// <summary>
    /// A path is only allowed to name a file directly inside the folder the plan is about.
    ///
    /// Everything a row deletes is named by that row, and a row's file name comes from a module
    /// name that a developer chose, so a name carrying separators or dots would otherwise be able
    /// to walk out of the folder and delete something else. Checked immediately before the delete
    /// rather than when the plan was built, because a plan can be applied later than it was made.
    /// </summary>
    internal static bool IsInsideFolder(string folder, string fileName)
    {
        if (fileName.Contains('/', StringComparison.Ordinal)
            || fileName.Contains('\\', StringComparison.Ordinal)
            || Path.IsPathRooted(fileName))
        {
            return false;
        }

        var root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(folder));
        var target = Path.GetFullPath(Path.Combine(root, fileName));
        return string.Equals(Path.GetDirectoryName(target), root, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The header the editor would write for this module, spliced onto the body COM gave us.
    ///
    /// A standard module's header is one attribute and is written here rather than exported, which
    /// saves a temporary file per module on the common case. Everything else is exported once, for
    /// its VB_Base, VB_PredeclaredId and VB_Exposed, none of which the object model will tell us.
    /// </summary>
    private static string FullSourceFor(DispatchObject component, string name, string kind, string body)
    {
        var header = kind == "standard"
            ? $"Attribute VB_Name = \"{name}\""
            : ExportedHeader(component, name);

        var text = new StringBuilder(header.Length + body.Length + 4);
        text.Append(header);
        if (!header.EndsWith('\n'))
        {
            text.Append("\r\n");
        }

        text.Append(body);
        if (body.Length > 0 && !body.EndsWith('\n'))
        {
            text.Append("\r\n");
        }

        return text.ToString();
    }

    /// <summary>
    /// The header lines of a temporary export: everything before the first line that is code.
    ///
    /// Read as Latin-1 so no code page can refuse a byte, and the name is rewritten from the one
    /// the editor gave us, which is the only field in there that can carry a character the export's
    /// code page could not spell.
    /// </summary>
    private static string ExportedHeader(DispatchObject component, string name)
    {
        var temporary = Path.Combine(Path.GetTempPath(), $"xlide-sync-{Guid.NewGuid():N}.cls");
        try
        {
            component.Invoke("Export", temporary);
            var lines = File.ReadAllText(temporary, Encoding.Latin1)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Split('\n');

            var header = new List<string>();
            foreach (var line in lines)
            {
                if (ModuleSync.IsAttributeLine(line) || ModuleSync.IsHeaderPreamble(line))
                {
                    header.Add(RewriteName(line, name));
                    continue;
                }

                if (line.Trim().Length == 0 && header.Count > 0)
                {
                    // A blank line inside the header block is part of it; one after the last
                    // attribute is where the code starts.
                    break;
                }

                if (header.Count > 0)
                {
                    break;
                }
            }

            return string.Join("\r\n", header);
        }
        catch (Exception ex)
        {
            Log.Warn($"sync: {name} would not export its header ({ex.Message}); writing a plain one");
            return $"Attribute VB_Name = \"{name}\"";
        }
        finally
        {
            TryDelete(temporary);

            // A UserForm exports its designer beside its code, whatever extension the code was
            // asked for, and that sidecar is as temporary as the file it belongs to.
            TryDelete(Path.ChangeExtension(temporary, ".frx"));
        }
    }

    private static string RewriteName(string line, string name) =>
        line.TrimStart().StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase)
            ? $"Attribute VB_Name = \"{name}\""
            : line;

    private static void WriteFile(string folder, string fileName, string source)
    {
        if (!IsInsideFolder(folder, fileName))
        {
            throw new InvalidOperationException($"'{fileName}' does not name a file in this folder.");
        }

        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, fileName), source, FileEncoding);
    }

    private static bool DeleteFile(string folder, string fileName)
    {
        if (!IsInsideFolder(folder, fileName))
        {
            throw new InvalidOperationException($"'{fileName}' does not name a file in this folder.");
        }

        var path = Path.Combine(folder, fileName);
        if (!File.Exists(path))
        {
            return false;
        }

        File.Delete(path);
        return true;
    }

    /// <summary>
    /// Creates a module from a file and then writes its body.
    ///
    /// Only the header is imported. Importing the whole file would be simpler and would be wrong:
    /// the editor reads an imported file in the machine's code page, so any character the code page
    /// cannot spell would arrive mangled, and the file this product writes is UTF-8. The header is
    /// ASCII, so importing it is safe, and it is what carries the kind and the attributes.
    ///
    /// Null when the module is there with its body in it. Anything else is the editor's complaint
    /// about the body, and the module exists but is empty - which the caller reports rather than
    /// counting as an import.
    /// </summary>
    private static string? CreateComponent(
        DispatchObject project,
        SyncItem item,
        WriteModuleText write,
        string projectId)
    {
        using var components = project.GetObject("VBComponents")
            ?? throw new InvalidOperationException("the project would not list its components");

        var header = HeaderOf(item.PayloadSource, item.ModuleName);
        var temporary = Path.Combine(
            Path.GetTempPath(),
            $"xlide-sync-{Guid.NewGuid():N}{Path.GetExtension(item.FileName)}");

        try
        {
            File.WriteAllText(temporary, $"{header}\r\n", Encoding.Latin1);
            components.Invoke("Import", temporary);
        }
        finally
        {
            TryDelete(temporary);
        }

        // Import names the module from its VB_Name attribute, and appends a number when something
        // already answers to that name. Nothing should, because a row only creates what the plan
        // found missing, but a rename between the plan and the apply is possible and silently
        // ending up with Module1 is not an outcome worth allowing.
        using var made = FindComponent(components, item.ModuleName);
        if (made is null)
        {
            throw new InvalidOperationException(
                $"the editor imported {item.FileName} but no module named {item.ModuleName} appeared");
        }

        var body = ModuleSync.CodeWithoutHeader(item.PayloadSource);
        return body.Length > 0 ? write(item.ModuleName, body, projectId) : null;
    }

    private static void RemoveComponent(DispatchObject project, string moduleName)
    {
        using var components = project.GetObject("VBComponents")
            ?? throw new InvalidOperationException("the project would not list its components");
        using var component = FindComponent(components, moduleName)
            ?? throw new InvalidOperationException($"no module named {moduleName} is in the project");

        components.InvokeWithObject("Remove", component);
    }

    /// <summary>
    /// The header lines of a file, which is everything up to the first line of code.
    /// </summary>
    private static string HeaderOf(string source, string moduleName)
    {
        var header = new List<string>();
        foreach (var line in ModuleSync.NormaliseEol(source).Split('\n'))
        {
            if (ModuleSync.IsAttributeLine(line) || ModuleSync.IsHeaderPreamble(line))
            {
                header.Add(line);
                continue;
            }

            if (header.Count > 0)
            {
                break;
            }
        }

        // A hand-written file need not carry a header at all, and the editor needs at least a name
        // to import it under.
        if (!header.Any(line => line.TrimStart().StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase)))
        {
            header.Add($"Attribute VB_Name = \"{moduleName}\"");
        }

        return string.Join("\r\n", header);
    }

    private static DispatchObject? FindComponent(DispatchObject components, string name)
    {
        var count = components.GetInt32("Count");
        for (var i = 1; i <= count; i++)
        {
            var component = components.GetItem(i);
            if (component is null)
            {
                continue;
            }

            if (string.Equals(component.GetString("Name"), name, StringComparison.OrdinalIgnoreCase))
            {
                return component;
            }

            component.Dispose();
        }

        return null;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex)
        {
            Log.Warn($"sync: a temporary file was left behind at {path} ({ex.Message})");
        }
    }
}
