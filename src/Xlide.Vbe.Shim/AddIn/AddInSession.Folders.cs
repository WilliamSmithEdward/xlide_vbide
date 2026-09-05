using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Core.Vba;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The folder each module says it is in, for the explorer's folder view (#23).
///
/// A folder is nothing but a comment at the top of a module - <c>'@Folder("Parent.Child")</c>,
/// the Rubberduck convention - so the tree needs every module's declarations read, opened or
/// not. Two sources feed one cache. The analysis pass already reads every module of a project
/// whose text moved and hands the snapshot here, so a folder typed into an open module reaches
/// the tree on the pass that follows the keystroke, at a text scan and no read of its own. And a
/// module the pass has not described yet - the beat between the tree's first publish and the
/// first pass - is read once through the object model, declarations only, so the first tree a
/// developer sees is already in folders rather than flat for a second.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>Folder by project identity and module name, as last read. Null is "none", kept so a read is not repeated.</summary>
    private readonly Dictionary<string, string?> _folders = new(StringComparer.OrdinalIgnoreCase);

    private static string FolderKey(string projectId, string module) => $"{projectId}\0{module}";

    /// <summary>
    /// A project's folders, from the snapshot an analysis pass just read. Runs on the pass's
    /// worker thread; republishes the tree through a hop to the host thread, and only when a
    /// folder actually changed - a keystroke inside a procedure must not redraw the explorer.
    /// </summary>
    private void RememberFolders(string projectId, IReadOnlyList<EngineModule> modules)
    {
        var changed = false;
        lock (_folders)
        {
            foreach (var module in modules)
            {
                var folder = FolderAnnotation.Of(module.Source);
                var key = FolderKey(projectId, module.ModuleName);
                if (!_folders.TryGetValue(key, out var was) || !string.Equals(was, folder, StringComparison.Ordinal))
                {
                    _folders[key] = folder;
                    changed = true;
                }
            }
        }

        if (changed)
        {
            Log.Verbose($"folders: {System.IO.Path.GetFileName(projectId)} moved a module between folders, republishing the tree");
            _editorSurface?.RunOnHostThread(PublishProjects);
        }
    }

    /// <summary>
    /// The folder as the shown module's TYPED text names it, on the live-analysis cadence.
    ///
    /// Typing feeds the engine's live copy and the full pass that fills the cache above is
    /// deferred until things go quiet, so an annotation typed into a module reached the tree
    /// only with the next full pass - which, with the caret parked on the line, was never (the
    /// owner, 2026-09-05, watching Late stay at the root under its new annotation). The live
    /// handler runs once per typing pause with the text in hand, which is the cadence the
    /// squiggles follow, so the tree follows it too. Host thread; republishes only on a change.
    /// </summary>
    private void NoteTypedFolder(string? projectId, string module, string source)
    {
        if (projectId is null)
        {
            return;
        }

        var folder = FolderAnnotation.Of(source);
        var key = FolderKey(projectId, module);
        lock (_folders)
        {
            if (_folders.TryGetValue(key, out var was) && string.Equals(was, folder, StringComparison.Ordinal))
            {
                return;
            }

            _folders[key] = folder;
        }

        Log.Verbose($"folders: {module} typed its way to {folder ?? "the root"}, republishing the tree");
        PublishProjects();
    }

    /// <summary>
    /// The folder of one component, for the tree: from the cache when a pass has described it,
    /// read from its declarations otherwise. Host thread, inside the tree's own walk.
    /// </summary>
    private string? FolderOf(string projectId, DispatchObject component, string name)
    {
        var key = FolderKey(projectId, name);
        lock (_folders)
        {
            if (_folders.TryGetValue(key, out var known))
            {
                return known;
            }
        }

        string? folder = null;
        try
        {
            // The declarations section only: an annotation below the first procedure does not
            // count, and a 65,000-line module's body is not worth reading to learn that.
            using var code = component.GetObject("CodeModule");
            var declarations = code?.GetInt32("CountOfDeclarationLines") ?? 0;
            if (declarations > 0)
            {
                folder = FolderAnnotation.Of(code!.GetStringIndexed("Lines", 1, declarations));
            }
        }
        catch (Exception ex)
        {
            // A module that will not answer is a module in no folder, and the pass that follows
            // will say otherwise if it is. Not cached, so that pass's answer is the first one kept.
            Log.Verbose($"folders: {name} could not be read ({ex.GetType().Name})");
            return null;
        }

        lock (_folders)
        {
            _folders[key] = folder;
        }

        return folder;
    }

    /// <summary>
    /// Forgets the modules a project no longer holds, so a module removed and re-added under the
    /// same name is read afresh rather than drawn where its predecessor sat.
    /// </summary>
    private void PruneFolders(string projectId, IEnumerable<string> live)
    {
        var keep = new HashSet<string>(live.Select(name => FolderKey(projectId, name)), StringComparer.OrdinalIgnoreCase);
        var prefix = FolderKey(projectId, string.Empty);
        lock (_folders)
        {
            foreach (var stale in _folders.Keys
                .Where(key => key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && !keep.Contains(key))
                .ToList())
            {
                _folders.Remove(stale);
            }
        }
    }

    /// <summary>
    /// The procedure the PAGE's caret is in, by the editor's own ProcOfLine on the shown module,
    /// or null in the declarations section or with nothing shown. The status bar's current
    /// procedure is computed on the page from the text; this is the native answer to the same
    /// question, which is what a parity check compares it to.
    /// </summary>
    private string? ProcedureAtPageCaret()
    {
        var surface = _editorSurface;
        if (surface?.Module is not { Length: > 0 } shown)
        {
            return null;
        }

        try
        {
            using var component = FindComponent(shown, _shownProject, out _);
            using var module = component?.GetObject("CodeModule");

            // vbext_pk_Proc, the procedure itself rather than a property's leg. Outside every
            // procedure the call raises, and the raise is the answer.
            var found = module?.CallToString("ProcOfLine", surface.CaretLine, 0);
            return string.IsNullOrWhiteSpace(found) ? null : found;
        }
        catch (Exception)
        {
            return null;
        }
    }
}
