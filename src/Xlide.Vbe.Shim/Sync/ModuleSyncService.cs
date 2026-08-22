using System.Text;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
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
/// the dialog and the xlide api ask the same question and get the same answer. This half reads what
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

    /// <summary>Reads every module of a project, with the full text a file needs - and for a
    /// UserForm, its DESIGN as markup, which is the half of a form that code alone cannot carry.
    /// The inventory is optional: without it the walk is the code-only one it always was.</summary>
    public static List<LiveModule> ReadLiveModules(DispatchObject project, ControlDefaults? defaults = null)
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

            // The design rides beside the code for a form, and only for a form. The projection is
            // the product's own - the same text the designer tab holds - so a file written here
            // and a document edited there are the same thing.
            string? design = null;
            if (string.Equals(kind, "userform", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    design = FormDesignService.MarkupOf(component, name, out _, defaults);
                }
                catch (Exception ex)
                {
                    Log.Info($"sync: {name}'s design could not be projected ({ex.GetType().Name})");
                }

                Log.Info($"sync: {name} is a form; its design is "
                    + (design is null ? "not available" : $"{design.Length} char(s)"));
            }

            modules.Add(new LiveModule(name, kind, FullSourceFor(component, name, kind, body), design));
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
            .Where(name => name is not null
                && (ModuleSync.IsModuleFileName(name)
                    || ModuleSync.IsDesignFileName(name)
                    || ModuleSync.IsSidecarFileName(name)))
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .AsParallel()
            .AsOrdered()
            .Select(name =>
            {
                // A SIDECAR IS LISTED AND NEVER READ. It is a form's controls in binary, and the
                // planner's only question about it is whether it is there - a form's pair has to
                // be whole before an import can create the form. Reading it as text would hand
                // the plan a megabyte of mojibake to diff.
                if (ModuleSync.IsSidecarFileName(name!))
                {
                    return new RepoFile(name!, string.Empty);
                }

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

    /// <summary>
    /// Carries out the rows a caller selected, and says what happened to each.
    ///
    /// HELD AGAINST THE FOLDER FOR THE WHOLE APPLY. Each file is written atomically, so nobody has
    /// ever read half a module, and that is a different guarantee from the one needed here: two
    /// Excels exporting the same project to one folder each work out a plan, and each plan's
    /// deletions are computed from the modules THAT instance holds. Interleave them and the second
    /// instance deletes files the first has just written, every step of it individually correct.
    ///
    /// The lock is a file in the target folder rather than anything machine-wide, because the
    /// folder is what is being contended: two instances exporting different projects to different
    /// folders have nothing to say to each other and should not wait on one another.
    /// </summary>
    public static SyncApplyResult Apply(
        DispatchObject project,
        SyncPlan plan,
        IReadOnlySet<string> selected,
        WriteModuleText write)
    {
        using var folderLock = FolderLock.Take(plan.Folder);
        if (folderLock.HeldByAnother)
        {
            return new SyncApplyResult([], [], [], [folderLock.Complaint!]);
        }

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
                    // The complaint is added as it stands: the writer already names the module it
                    // is about, and wrapping it in the name again reads as two different things
                    // went wrong. Only the exception path below adds a name, because an exception
                    // message has none.
                    case SyncStatus.WillCreate when plan.Direction == SyncDirection.Import:
                        if (CreateComponent(project, item, write, plan.ProjectId, plan.Folder) is { } createRefused)
                        {
                            failed.Add(createRefused);
                        }
                        else
                        {
                            changed.Add(item.ModuleName);
                        }

                        break;

                    // A DESIGN goes through the markup apply, not through a module write: the file
                    // is xlide's own text for the form's controls, and the apply that lands it is
                    // the same name-keyed diff the designer tab's Ctrl+S makes. A document that
                    // does not parse changes nothing and says which line stopped it.
                    case SyncStatus.WillUpdate when item.IsDesign:
                    {
                        using var designComponents = project.GetObject("VBComponents");
                        using var form = designComponents is null
                            ? null
                            : FindComponent(designComponents, item.ModuleName);
                        if (form is null)
                        {
                            failed.Add($"{item.FileName} (no form named {item.ModuleName})");
                            break;
                        }

                        var outcome = FormDesignService.ApplyMarkup(form, item.ModuleName, item.PayloadSource);
                        if (outcome.Ok)
                        {
                            changed.Add($"{item.ModuleName} (design: +{outcome.Added.Count} "
                                + $"-{outcome.Removed.Count}, {outcome.Set} set)");
                        }
                        else
                        {
                            failed.Add($"{item.FileName} ({outcome.Refused ?? "the design was refused"})");
                        }

                        break;
                    }

                    case SyncStatus.WillUpdate:
                        if (write(item.ModuleName, ModuleSync.CodeWithoutHeader(item.PayloadSource), plan.ProjectId)
                            is { } updateRefused)
                        {
                            failed.Add(updateRefused);
                        }
                        else
                        {
                            changed.Add(item.ModuleName);
                        }

                        break;

                    // A FORM IS EXPORTED BY THE VBE, not assembled here. Its controls live in a
                    // binary sidecar that only the exporter writes, and the .frm names that file -
                    // so a form written from spliced text names a sidecar that does not exist.
                    // The encoding compromise the rest of this class exists to avoid is accepted
                    // for a form deliberately: the .frm and its .frx must agree byte for byte, and
                    // only the exporter can promise that.
                    case SyncStatus.WillCreate when IsFormCode(item):
                    case SyncStatus.WillWrite when IsFormCode(item):
                    {
                        using var formComponents = project.GetObject("VBComponents");
                        using var form = formComponents is null
                            ? null
                            : FindComponent(formComponents, item.ModuleName);
                        if (form is null)
                        {
                            failed.Add($"{item.FileName} (no form named {item.ModuleName})");
                            break;
                        }

                        var destination = Path.Combine(plan.Folder, item.FileName);
                        form.Invoke("Export", destination);
                        changed.Add(item.FileName);

                        // The sidecar is written beside it by the exporter, under the same name.
                        var sidecar = Path.ChangeExtension(item.FileName, ".frx");
                        if (File.Exists(Path.Combine(plan.Folder, sidecar)))
                        {
                            changed.Add(sidecar);
                        }

                        break;
                    }

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

    /// <summary>
    /// Writes an exported file so it is either the old one or the new one, never half of either.
    ///
    /// WriteAllText truncates and then fills, so between those two there is a file on disk with
    /// the developer's name on it and nothing in it. Anything reading the folder in that window
    /// reads a truncated module: the companion editor watching the folder, a build, a git status,
    /// another Excel importing from the same folder - which is the case this side has no lock for.
    /// A crash or a full disk leaves the same wreckage permanently.
    ///
    /// So the content goes to a temporary file beside it and is then moved over the top, which the
    /// file system does as one step.
    ///
    /// Move rather than Replace, though Replace would keep the destination's attributes. Replace
    /// is the fussier call - it fails outright on some network and non-NTFS volumes - and an
    /// export folder is exactly the kind of place that is sometimes a share. A moved-in file
    /// inherits the folder's permissions, which is what every file this writes for the first time
    /// already gets.
    ///
    /// The temporary file is named after the one it will become, so a stray left by a process that
    /// died mid-write is obviously ours and obviously junk. It sits in the SAME folder on purpose:
    /// an atomic move needs one volume, and the system temp folder is regularly a different one.
    /// Nothing reads it as a module either way, because it is neither a .bas nor a .cls.
    /// </summary>
    private static void WriteFile(string folder, string fileName, string source)
    {
        if (!IsInsideFolder(folder, fileName))
        {
            throw new InvalidOperationException($"'{fileName}' does not name a file in this folder.");
        }

        Directory.CreateDirectory(folder);

        var path = Path.Combine(folder, fileName);
        var partial = Path.Combine(folder, $".{fileName}.xlide-partial");

        try
        {
            File.WriteAllText(partial, source, FileEncoding);
            File.Move(partial, path, overwrite: true);
        }
        catch
        {
            TryDelete(partial);
            throw;
        }
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
        string projectId,
        string folder)
    {
        using var components = project.GetObject("VBComponents")
            ?? throw new InvalidOperationException("the project would not list its components");

        // A FORM IS CREATED FROM ITS PAIR, whole, and it is the one create that imports the file
        // rather than a header made from it.
        if (ModuleSync.IsFormText(item.PayloadSource))
        {
            return CreateFormFromPair(components, item, folder);
        }

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
        if (body.Length == 0)
        {
            return null;
        }

        if (write(item.ModuleName, body, projectId) is not { } refused)
        {
            return null;
        }

        // The module exists and is empty, because only the header went in. Take it away again: a
        // row reported failed should leave the project as it found it, and an empty module named
        // after a file the developer still has is worse than no module at all.
        try
        {
            components.InvokeWithObject("Remove", made);
        }
        catch (Exception ex)
        {
            Log.Warn($"sync: {item.ModuleName} was created, its body was refused, and the empty"
                + $" module could not be taken away again ({ex.Message})");
            return $"{refused} An empty {item.ModuleName} was left in the project.";
        }

        return refused;
    }

    /// <summary>
    /// Creates a FORM from the pair the VBE's own exporter wrote: the header text, and the binary
    /// sidecar it names. This is the whole reason import can make a form at all - the sidecar
    /// carries the controls, and nothing this side can build one from text.
    ///
    /// THROUGH A TEMPORARY COPY, for two reasons that both matter. The importer decides what to
    /// make from the file's EXTENSION, and the shared planner writes a form's code as `.cls`
    /// (xlide_vscode#21) - imported under that name it becomes a class module holding a form's
    /// header, which is not a form and cannot be undone into one. And the sidecar has to sit
    /// beside it under exactly the name the `OleObjectBlob` line spells, which a copy can promise
    /// and a developer's folder cannot.
    ///
    /// The encoding compromise the rest of this class avoids is accepted here for the reason the
    /// EXPORT accepts it: the pair must agree byte for byte, so the bytes are copied rather than
    /// the text.
    /// </summary>
    private static string? CreateFormFromPair(DispatchObject components, SyncItem item, string folder)
    {
        var source = Path.Combine(folder, item.FileName);
        if (!File.Exists(source))
        {
            return $"{item.FileName} is not in the folder any more";
        }

        var staging = Path.Combine(Path.GetTempPath(), $"xlide-import-{Guid.NewGuid():N}");
        Directory.CreateDirectory(staging);
        try
        {
            var imported = Path.Combine(staging, $"{item.ModuleName}.frm");
            File.Copy(source, imported);

            if (ModuleSync.SidecarNamedBy(item.PayloadSource) is { Length: > 0 } named)
            {
                var sidecar = Path.Combine(folder, named);
                if (!File.Exists(sidecar))
                {
                    return $"{item.FileName} names {named}, which is not in the folder";
                }

                File.Copy(sidecar, Path.Combine(staging, Path.GetFileName(named)));
            }

            components.Invoke("Import", imported);
        }
        catch (Exception why)
        {
            // THE IMPORTER'S OWN LOG IS THE ONLY EXPLANATION IT GIVES. "Errors during load" names
            // a `.log` file it writes beside the module - and that file is in the staging folder
            // this method is about to delete, so the reason would die with it (2026-08-16, where
            // it took a second run to find out that a pair whose `Begin` line and `VB_Name`
            // disagree is refused). Whatever it says is carried out with the refusal.
            var log = Path.Combine(staging, $"{item.ModuleName}.log");
            var said = File.Exists(log) ? ReadLogQuietly(log) : null;
            return said is { Length: > 0 }
                ? $"{item.FileName} ({why.Message.Trim()}) - the editor's log says: {said}"
                : $"{item.FileName} ({why.Message.Trim()})";
        }
        finally
        {
            TryDeleteFolder(staging);
        }

        // Import takes the name from VB_Name and appends a number when something already answers
        // to it. Nothing should - a create row is a row the plan found missing - but a form named
        // UserForm1 arriving where EntryForm was asked for is a silent wrong outcome, not an
        // import, so it is reported and taken back out.
        using var made = FindComponent(components, item.ModuleName);
        if (made is null)
        {
            return $"the editor imported {item.FileName} but no form named {item.ModuleName} appeared";
        }

        FormDesignService.KeepDesignerDown(made);
        return null;
    }

    /// <summary>The importer's log as one line, short enough to stand in a refusal. Its own
    /// failure is nothing: a missing explanation is what this method exists to improve on.</summary>
    private static string? ReadLogQuietly(string path)
    {
        try
        {
            var lines = File.ReadAllLines(path, FileEncoding)
                .Select(line => line.Trim())
                .Where(line => line.Length > 0)
                .Take(4);
            return string.Join(" / ", lines);
        }
        catch (Exception ex)
        {
            Log.Warn($"sync: the editor's import log could not be read ({ex.Message})");
            return null;
        }
    }

    private static void TryDeleteFolder(string folder)
    {
        try
        {
            Directory.Delete(folder, recursive: true);
        }
        catch (Exception ex)
        {
            Log.Warn($"sync: the import staging folder could not be removed ({ex.Message})");
        }
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

    /// <summary>A form's CODE row - the .frm - as against its design row or any other module.</summary>
    private static bool IsFormCode(SyncItem item) =>
        !item.IsDesign
        && string.Equals(item.ModuleKind, "userform", StringComparison.OrdinalIgnoreCase);

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
