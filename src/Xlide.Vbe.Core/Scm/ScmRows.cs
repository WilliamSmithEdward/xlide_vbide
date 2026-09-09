using Xlide.Vbe.Core.Sync;

namespace Xlide.Vbe.Core.Scm;

/// <summary>A module as the editor holds it.</summary>
/// <param name="Kind">standard, class, document or userform.</param>
/// <param name="Code">The module's own text, no attribute header.</param>
public sealed record ScmLiveModule(string Name, string Kind, string Code);

/// <summary>One row of the pane.</summary>
/// <param name="File">The module's file name in the folder.</param>
/// <param name="Status">
/// Against the branch head: modified, added, deleted, renamed. Against the folder: folderNewer,
/// missingInFolder, missingInProject.
/// </param>
/// <param name="From">The module's earlier name, for a renamed row.</param>
public sealed record ScmRow(string Module, string Kind, string File, string Status, string? From);

/// <summary>The rows: the live project against the branch head, and against the folder.</summary>
public static class ScmRows
{
    /// <summary>
    /// The live project against the branch head. Only code is compared - the attribute header a
    /// file carries and a form's design are the export's business at save and commit - so a file's
    /// header is taken off and both sides go through the same normalisation, which is what keeps
    /// an untouched module clean whatever the file's line endings or a leading blank line. A
    /// deleted+added pair the change log knows as a rename becomes one renamed row (Module = to,
    /// From = from). Unchanged modules are not rows. Sorted by module name, ordinal ignore case.
    /// </summary>
    /// <param name="headFiles">
    /// File name (no directory) to the file's text, header included, for every module file at
    /// HEAD under the folder.
    /// </param>
    /// <param name="renames">(from, to) pairs the change log knows.</param>
    public static IReadOnlyList<ScmRow> Compute(
        IReadOnlyList<ScmLiveModule> live,
        IReadOnlyDictionary<string, string> headFiles,
        IReadOnlyList<(string From, string To)> renames)
    {
        ArgumentNullException.ThrowIfNull(live);
        ArgumentNullException.ThrowIfNull(headFiles);
        ArgumentNullException.ThrowIfNull(renames);

        var head = ByFileName(headFiles);
        var rows = new List<ScmRow>();
        var claimed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var module in live)
        {
            if (FileOf(module, head) is { } paired)
            {
                claimed.Add(paired.File);
                if (!ModuleSync.SameCode(paired.Text, module.Code))
                {
                    rows.Add(new ScmRow(module.Name, module.Kind, paired.File, "modified", null));
                }
            }
            else
            {
                rows.Add(new ScmRow(
                    module.Name, module.Kind, ModuleSync.FileNameFor(module.Name, module.Kind), "added", null));
            }
        }

        foreach (var (file, text) in head)
        {
            if (!claimed.Contains(file) && ModuleSync.IsModuleFileName(file))
            {
                rows.Add(new ScmRow(
                    ModuleSync.ModuleNameFromFileName(file),
                    ModuleSync.ClassifyFile(file, text),
                    file,
                    "deleted",
                    null));
            }
        }

        foreach (var (from, to) in renames)
        {
            var gone = rows.FindIndex(row => row.Status == "deleted" && Same(row.Module, from));
            var came = rows.FindIndex(row => row.Status == "added" && Same(row.Module, to));
            if (gone < 0 || came < 0)
            {
                continue;
            }

            var arrived = rows[came];
            rows[came] = arrived with { Status = "renamed", From = rows[gone].Module };
            rows.RemoveAt(gone);
        }

        return Sorted(rows);
    }

    /// <summary>
    /// The folder against the live project: folderNewer when both exist and the code differs,
    /// missingInFolder for a module with no file, missingInProject for a file with no module (File
    /// names it, Module is the name without its extension, Kind from the file). This is the
    /// outside-change signal - a checkout, a pull, an edit in another editor - so nothing here
    /// decides anything; the pane offers Import and Export and the developer chooses.
    /// </summary>
    public static IReadOnlyList<ScmRow> Outside(
        IReadOnlyList<ScmLiveModule> live, IReadOnlyDictionary<string, string> folderFiles)
    {
        ArgumentNullException.ThrowIfNull(live);
        ArgumentNullException.ThrowIfNull(folderFiles);

        var folder = ByFileName(folderFiles);
        var rows = new List<ScmRow>();
        var claimed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var module in live)
        {
            if (FileOf(module, folder) is { } paired)
            {
                claimed.Add(paired.File);
                if (!ModuleSync.SameCode(paired.Text, module.Code))
                {
                    rows.Add(new ScmRow(module.Name, module.Kind, paired.File, "folderNewer", null));
                }
            }
            else
            {
                rows.Add(new ScmRow(
                    module.Name, module.Kind, ModuleSync.FileNameFor(module.Name, module.Kind), "missingInFolder", null));
            }
        }

        foreach (var (file, text) in folder)
        {
            if (!claimed.Contains(file) && ModuleSync.IsModuleFileName(file))
            {
                rows.Add(new ScmRow(
                    ModuleSync.ModuleNameFromFileName(file),
                    ModuleSync.ClassifyFile(file, text),
                    file,
                    "missingInProject",
                    null));
            }
        }

        return Sorted(rows);
    }

    /// <summary>
    /// The file a live module has on the other side: the name this product writes, or, for a
    /// form, the .cls the companion editor writes a form as (xlide_vscode#21) when no .frm is
    /// there. A folder the two products share holds a form either way, and read by name alone it
    /// was two false rows - the form "added" under the spelling this side writes and its own file
    /// "deleted" - where the file is the form's, whatever it is called.
    /// </summary>
    private static (string File, string Text)? FileOf(ScmLiveModule module, Dictionary<string, string> files)
    {
        var file = ModuleSync.FileNameFor(module.Name, module.Kind);
        if (files.TryGetValue(file, out var text))
        {
            return (file, text);
        }

        if (string.Equals(module.Kind, "userform", StringComparison.OrdinalIgnoreCase))
        {
            var asClass = ModuleSync.FileNameFor(module.Name, "class");
            if (files.TryGetValue(asClass, out var classText) && ModuleSync.IsFormText(classText))
            {
                return (asClass, classText);
            }
        }

        return null;
    }

    private static Dictionary<string, string> ByFileName(IReadOnlyDictionary<string, string> files)
    {
        var byName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (file, text) in files)
        {
            byName[Path.GetFileName(file)] = text;
        }

        return byName;
    }

    private static bool Same(string left, string right) =>
        string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

    private static List<ScmRow> Sorted(List<ScmRow> rows) =>
        [.. rows
            .OrderBy(row => row.Module, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.File, StringComparer.OrdinalIgnoreCase)];
}
