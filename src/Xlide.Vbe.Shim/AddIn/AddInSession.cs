using System.Text.Json;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.Sync;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// One connected lifetime of the add-in inside one editor instance.
///
/// The session owns every resource that must be released before the host tears down: automation
/// references, window hooks, tool windows, and the engine connection. It is stopped from
/// OnBeginShutdown, which is the last moment at which touching the object model is safe.
/// </summary>
internal sealed partial class AddInSession : IDisposable
{
    private readonly DispatchObject _editor;

    /// <summary>
    /// The editor's own object for this add-in. Held because it is what the editor would want back
    /// to create a tool window, and released at shutdown like everything else.
    ///
    /// No tool window is created. The editor will not size one in any state: setting a width or a
    /// height throws whether the window floats or is docked, docking one produces a band six pixels
    /// high with a negative client area, and its contents do not follow when the user resizes it.
    /// A panel in one is either invisible or a stub floating over the code. The product's panels
    /// live in the editing surface, which owns its own layout completely.
    /// </summary>
    private readonly DispatchObject? _addIn;

    private CodePaneTracker? _codePanes;
    private AnalysisService? _analysis;
    private ImmediateEvaluator? _immediate;
    private ImmediateReader? _immediateReader;
    private bool _windowsHidden;
    private EditorSurface? _editorSurface;
    private HostChrome? _hostChrome;

    /// <summary>
    /// The most recent findings for every module, kept so a module can be decorated the moment it
    /// is shown. Analysis runs per project and the surface shows one module at a time, so without
    /// this a module opened between two passes carries no squiggles until the next one.
    /// </summary>
    private IReadOnlyList<Finding> _findings = [];

    /// <summary>
    /// The VBE validates a line when the caret leaves it. While the developer is typing on a
    /// line, verdicts touching it are held out of every publish; the caret settling elsewhere
    /// releases the hold and republishes from the unfiltered findings above - no re-analysis.
    /// </summary>
    private readonly ActiveLineHold _activeLineHold = new();

    /// <summary>
    /// What each module read back as the last time this add-in wrote it, keyed by
    /// <see cref="WrittenKey"/> - one baseline per (workbook, module), never per bare name.
    ///
    /// This is the baseline a later comparison is made against, and it is deliberately not the
    /// surface's text. The editor rewrites what it is given as it takes a module in: it completes
    /// a procedure's parentheses, inserts the blank body of one, and respells keywords. Comparing
    /// the module against the surface would see all of that as a change and pull it back into the
    /// document, on top of somebody who is still typing. Comparing it against this sees only what
    /// changed the module after we wrote it, which is what "something else changed it" means.
    ///
    /// The workbook is part of the key because the baseline feeds the line-diff write: with two
    /// workbooks' Module1 both live (decision 12), a name-keyed baseline would diff one module
    /// against the other's text and write the resulting merge over real code.
    /// </summary>
    private readonly Dictionary<string, string> _writtenModules = new(StringComparer.Ordinal);

    /// <summary>The written-baseline key: a module of one workbook, by lowercased display name.</summary>
    private static string WrittenKey(string component, string? projectDisplay) =>
        $"{(projectDisplay ?? string.Empty).ToLowerInvariant()}\0{component.ToLowerInvariant()}";

    /// <summary>
    /// The last thing said about a module's write, so it is said once rather than every debounce.
    ///
    /// The write-back fires on a pause in typing, so a module carrying a line the editor will not
    /// hold is refused again on every pause. The complaint is worth one notice, not one per pause:
    /// the second is the same sentence and the tenth is in the way of the work. Cleared when the
    /// module writes, so the next real problem is announced.
    /// </summary>
    private readonly Dictionary<string, string> _saidAboutWrite = new(StringComparer.Ordinal);

    /// <summary>
    /// The last thing said about a module losing a character to the host's code page.
    ///
    /// Its own record rather than a row in the one above, because it is not cleared by the write
    /// succeeding: that write DID succeed, with a character converted, and the surface goes on
    /// holding the original, so every later write-back finds the same loss. Saying it again each
    /// time would be exactly the repetition the other record exists to prevent.
    /// </summary>
    private readonly Dictionary<string, string> _saidAboutConversion = new(StringComparer.Ordinal);

    /// <summary>
    /// Identity of the project whose module the surface is showing, or null when it could not
    /// be told. This is the tie-break for every bare module name that reaches the session while
    /// the page's protocol still speaks in names alone: the module being edited outranks a
    /// same-named module in another workbook.
    /// </summary>
    private string? _shownProject;

    /// <summary>When an unknown project last triggered a full pass, so two quick shows cost one.</summary>
    private long _lastUnknownProjectPass;

    /// <summary>The developer's settings, loaded once per session and written on every change.</summary>
    private ProductSettings _settings = ProductSettings.Default;

    /// <summary>Where the settings live: beside the logs, hand-editable, growable.</summary>
    private static string SettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "xlide_vbide",
        "settings.json");

    /// <summary>
    /// Import and export, for whoever is asking.
    ///
    /// The dialog and the xlide api both land here, which is the whole design: an api that wrote
    /// files by itself would drift from the button, and the first anyone would know of it is a
    /// harness that passes against a product that is broken. One implementation, two doors.
    /// </summary>
    /// <summary>
    /// Reads the editor for everything a plan needs. MUST run on the host thread: every line of it
    /// is the object model, which is apartment bound.
    ///
    /// Answers false with a reason when there is nothing to plan against, so the caller can send
    /// that reason rather than an empty plan.
    /// </summary>
    private bool TryGatherSyncInputs(
        IReadOnlyDictionary<string, string> query,
        out SyncInputs? inputs,
        out string? refusal)
    {
        inputs = null;
        refusal = null;

        try
        {
            query.TryGetValue("direction", out var direction);
            query.TryGetValue("project", out var project);
            query.TryGetValue("folder", out var asked);
            query.TryGetValue("mode", out var mode);

            var projectId = ProjectIdFromDisplay(project) ?? _shownProject;
            if (string.IsNullOrEmpty(projectId))
            {
                refusal = Refuse("no project is shown, and none was named");
                return false;
            }

            var remembered = LoadSyncSettings().For(projectId);
            var folder = string.IsNullOrEmpty(asked) ? remembered.Folder : asked;
            if (string.IsNullOrEmpty(folder))
            {
                refusal = Refuse("no folder was given and this project remembers none");
                return false;
            }

            // THE PROJECT THAT WAS RESOLVED, not whichever one the editor happens to call active.
            //
            // These two lines used to fall back differently: the identity to the SHOWN project and
            // the modules to ActiveVBProject. Open two workbooks and they are routinely different -
            // measured 2026-08-21, with nothing contrived, a session two seconds after opening
            // DebugFixture and TwinFixture side by side answered a plan titled DebugFixture.xlsm
            // whose rows were TwinFixture's six modules.
            //
            // Applying that plan would export one workbook's modules into the other's folder, or
            // import that folder over the other workbook's code. The lookup takes an identity as
            // well as a display name, so it is given the id that was resolved, and a project that
            // cannot be found is refused rather than swapped for a different one. It is the same
            // defect this helper's own comment records being closed at nine other call sites.
            using var target = FindProjectByDisplayName(projectId);
            if (target is null)
            {
                refusal = Refuse($"the project could not be reached: nothing open is {projectId}");
                return false;
            }

            var importing = string.Equals(direction, "import", StringComparison.OrdinalIgnoreCase);

            // Timed because the plan's cost decides what is worth moving off this thread, and a
            // guess about which half is slow is how the wrong half gets optimised.
            var readingAt = System.Diagnostics.Stopwatch.StartNew();
            var live = ModuleSyncService.ReadLiveModules(target, _controlDefaults);
            Log.Verbose($"sync: read {live.Count} module(s) from the editor in {readingAt.ElapsedMilliseconds}ms");

            inputs = new SyncInputs(
                projectId,
                DisplayFromProjectId(projectId) ?? projectId,
                folder,
                live,
                mode ?? (importing ? remembered.ImportMode : remembered.ExportMode),
                importing);
            return true;
        }
        catch (Exception ex)
        {
            Log.Error("sync: the project could not be read for a plan", ex);
            refusal = Refuse(ex.Message.Trim());
            return false;
        }
    }

    private static string Refuse(string why) => System.Text.Json.JsonSerializer.Serialize(
        new SyncErrorReply(why), SyncJsonContext.Default.SyncErrorReply);

    /// <summary>Everything a plan needs that had to be read from the editor.</summary>
    private sealed record SyncInputs(
        string ProjectId,
        string DisplayName,
        string Folder,
        List<LiveModule> Live,
        string? Mode,
        bool Importing);

    /// <summary>
    /// Works out the plan, and says which planner did.
    ///
    /// NO COM IN HERE, deliberately: everything it touches is the modules already read, the
    /// folder, and the engine. That is what lets the dialog run it off the host thread, and the
    /// reason it matters is measured: on a project of 81,795 lines the shared planner takes
    /// 2,167ms, and every one of those milliseconds used to be Excel frozen (2026-08-09).
    ///
    /// WHICH PLANNER, and only the planner. Both answer the same shape and both are carried out
    /// by the same apply, so the choice changes who DECIDES what an import would do, never who
    /// does it. The companion editor's is the default because the two products write into the
    /// same folders and a developer moves between them: one implementation of file naming,
    /// module classification and staleness cannot disagree with itself. The built-in one is not
    /// a fallback nobody meant to use, because it needs no engine, so import and export keep working
    /// when the engine is down.
    /// </summary>
    private (SyncPlan Plan, string Planner) BuildPlan(SyncInputs inputs)
    {
        var folderAt = System.Diagnostics.Stopwatch.StartNew();
        var onDisk = ModuleSyncService.ReadFolder(inputs.Folder);
        Log.Verbose($"sync: read {onDisk.Count} file(s) from the folder in {folderAt.ElapsedMilliseconds}ms");

        var planningAt = System.Diagnostics.Stopwatch.StartNew();
        var wantsShared = !string.Equals(_settings.SyncEngine, "builtIn", StringComparison.OrdinalIgnoreCase);
        var shared = wantsShared
            ? SharedPlan(inputs.ProjectId, inputs.DisplayName, inputs.Folder, inputs.Live, inputs.Mode, inputs.Importing)
            : null;

        var plan = shared ?? (inputs.Importing
            ? ModuleSync.PlanImport(
                inputs.ProjectId, inputs.DisplayName, inputs.Folder, inputs.Live, onDisk,
                ModuleSync.ImportModeFrom(inputs.Mode))
            : ModuleSync.PlanExport(
                inputs.ProjectId, inputs.DisplayName, inputs.Folder, inputs.Live, onDisk,
                ModuleSync.ExportModeFrom(inputs.Mode)));

        // A FORM'S DESIGN is this product's own row, and the SHARED planner cannot know about it:
        // the markup belongs to xlide_vbide and the companion editor has never heard of it. So the
        // rows are added to whichever plan came back, from the one implementation Core holds.
        if (shared is not null)
        {
            var designs = inputs.Importing
                ? ModuleSync.DesignRowsForImport(inputs.Live, onDisk)
                : ModuleSync.DesignRowsForExport(
                    inputs.Live, onDisk.ToDictionary(one => one.FileName, StringComparer.OrdinalIgnoreCase));

            if (designs.Count > 0)
            {
                plan = new SyncPlan
                {
                    Direction = plan.Direction,
                    ProjectId = plan.ProjectId,
                    ProjectName = plan.ProjectName,
                    Folder = plan.Folder,
                    ExportMode = plan.ExportMode,
                    ImportMode = plan.ImportMode,
                    Items = [.. plan.Items, .. designs],
                    Warnings = plan.Warnings,
                };
            }
        }

        Log.Verbose($"sync: worked out {plan.Items.Count} row(s) in {planningAt.ElapsedMilliseconds}ms"
            + $" ({(shared is null ? "built in" : "shared")})");

        return (plan, shared is null ? "builtIn" : "xlide");
    }

    private string HandleSync(IReadOnlyDictionary<string, string> query, string body)
    {
        // Unwritten CODE edits go with the sync, the same way they go with a run: the plan reads
        // the live project over COM, and a module the developer has not finished typing must be
        // read as typed, not as last written back. The designer documents' half of the same rule
        // is flushed by the callers (the dialog page-side, the api route's pool pre-step),
        // because applying a designer document needs the page to pump and this method runs on
        // the host thread.
        //
        // THE DIRTY SET IS TAKEN FIRST, because this flush is also what blinded the import
        // guard: writing the typing so the plan reads the text the developer sees clears the
        // very flag the guard refuses by, and an import then replaced the freshly written
        // typing with the file - the exact loss the guard exists to stop, one save later
        // (verify.ps1 -Deep, 2026-08-19). A module that was mid-typing when the gesture
        // started stays refused for THIS apply, whatever the flush did for the plan.
        var unwrittenBefore = new HashSet<string>(
            (_editorSurface?.DocumentTable ?? [])
                .Where(doc => doc.Unwritten)
                .Select(doc =>
                    $"{(doc.Project ?? string.Empty).ToLowerInvariant()}\0{doc.Module.ToLowerInvariant()}"),
            StringComparer.Ordinal);
        _editorSurface?.FlushEdits();

            // Import and export, the same way the dialog does it.
            //
            // This route does NOT have its own idea of what an import means: it calls the
            // service the dialog calls, so a plan read here is the plan drawn there and an
            // apply leaves the project in the state the button would have left it. That is the
            // whole point of routing both through one service rather than teaching the api to
            // write files by itself.
            query.TryGetValue("action", out var syncAction);
            query.TryGetValue("direction", out var syncDirection);
            query.TryGetValue("project", out var syncProject);
            query.TryGetValue("folder", out var syncFolder);
            query.TryGetValue("mode", out var syncMode);
            query.TryGetValue("select", out var syncSelect);

            var syncProjectId = ProjectIdFromDisplay(syncProject) ?? _shownProject;
            if (string.IsNullOrEmpty(syncProjectId))
            {
                return System.Text.Json.JsonSerializer.Serialize(
                    new SyncErrorReply("no project is shown, and none was named"),
                    SyncJsonContext.Default.SyncErrorReply);
            }

            try
            {
                if (syncAction == "settings")
                {
                    var stored = LoadSyncSettings();
                    if (syncFolder is not null || query.ContainsKey("exportMode") || query.ContainsKey("importMode"))
                    {
                        var was = stored.For(syncProjectId);
                        var choice = new SyncChoice
                        {
                            Folder = syncFolder ?? was.Folder,
                            ExportMode = query.TryGetValue("exportMode", out var em) ? em : was.ExportMode,
                            ImportMode = query.TryGetValue("importMode", out var im) ? im : was.ImportMode,
                        };
                        stored = stored.With(syncProjectId, choice);
                        SaveSyncSettings(stored);
                    }

                    var now = stored.For(syncProjectId);
                    return System.Text.Json.JsonSerializer.Serialize(
                        new SyncSettingsReply(syncProjectId, now.Folder, now.ExportMode, now.ImportMode),
                        SyncJsonContext.Default.SyncSettingsReply);
                }

                var remembered = LoadSyncSettings().For(syncProjectId);
                var folder = string.IsNullOrEmpty(syncFolder) ? remembered.Folder : syncFolder;
                if (string.IsNullOrEmpty(folder))
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new SyncErrorReply("no folder was given and this project remembers none"),
                        SyncJsonContext.Default.SyncErrorReply);
                }

                var importing = string.Equals(syncDirection, "import", StringComparison.OrdinalIgnoreCase);
                // The resolved identity, for the reasons under the other one of these.
                using var syncTarget = FindProjectByDisplayName(syncProjectId);
                if (syncTarget is null)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new SyncErrorReply($"the project could not be reached: nothing open is {syncProjectId}"),
                        SyncJsonContext.Default.SyncErrorReply);
                }

                var live = ModuleSyncService.ReadLiveModules(syncTarget, _controlDefaults);
                var displayName = DisplayFromProjectId(syncProjectId) ?? syncProjectId;

                // WHICH PLANNER, and only the planner.
                //
                // Both answer the same shape and both are carried out by the same apply below, so
                // the choice changes who DECIDES what an import would do, never who does it. The
                // companion editor's planner is the default because the two products write into
                // the same folders and a developer moves between them: one implementation of file
                // naming, module classification and staleness cannot disagree with itself.
                //
                // The built-in one is not a fallback nobody meant to use, because it needs no engine, so
                // import and export keep working when the engine is down, and it is what a
                // developer chooses if they would rather the add-in never depended on it.
                var mode = syncMode ?? (importing ? remembered.ImportMode : remembered.ExportMode);
                var (plan, planner) = BuildPlan(
                    new SyncInputs(syncProjectId, displayName, folder, live, mode, importing));

                if (syncAction != "apply")
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        SyncPlanReplyFor(plan, planner),
                        SyncJsonContext.Default.SyncPlanReply);
                }

                // Which rows to carry out. A body names them one per line, which is what the
                // dialog sends after the developer has ticked and unticked; `select=checked`
                // takes the plan's own ticks, which is what a caller driving this from a script
                // almost always means; `select=all` takes everything the plan offered.
                var chosen = !string.IsNullOrWhiteSpace(body)
                    ? new HashSet<string>(
                        body.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
                        StringComparer.Ordinal)
                    : syncSelect == "all"
                        ? [.. plan.Items.Select(item => item.Id)]
                        : new HashSet<string>(
                            plan.Items.Where(item => item.Checked).Select(item => item.Id),
                            StringComparer.Ordinal);

                // ALL OR NOTHING, because the good copy is the file and it is safe on disk.
                //
                // A file carrying a character this machine's VBA cannot store would import as
                // something else - question marks, or a decomposed accent turned into a spacing
                // one - and the module would then differ from the file, so the next export would
                // be ticked and would overwrite the developer's source with it. Measured
                // 2026-08-09: one import and one export destroyed a Cyrillic file byte for byte,
                // reporting "1 changed, 0 failed" at both ends.
                var applied = ModuleSyncService.Apply(
                    syncTarget, plan, chosen,
                    (component, text, owner) =>
                    {
                        /*
                         * A MODULE WITH EDITS THE DEVELOPER HAS NOT WRITTEN IS NOT IMPORTED OVER.
                         *
                         * Import writes the module and then syncs the surface, which replaces the
                         * document on screen: measured 2026-08-09, typing into a module and then
                         * importing over it took the typing away with no question, no notice, and
                         * "1 changed, 0 failed". Closing a tab in that state raises a whole
                         * save/discard/cancel gate, because throwing a developer's work away
                         * silently is not something this product does anywhere else.
                         *
                         * Refused rather than asked, and the difference is worth stating: an
                         * import touching twelve modules would ask twelve questions, and the
                         * developer already has somewhere to see this - the row says so, beside
                         * every other row that could not be applied. They write or discard, and
                         * import again.
                         */
                        var guardKey = $"{(DisplayFromProjectId(owner) ?? string.Empty).ToLowerInvariant()}"
                            + $"\0{component.ToLowerInvariant()}";
                        if (unwrittenBefore.Contains(guardKey)
                            || _editorSurface?.HasUnwritten(component, DisplayFromProjectId(owner)) == true)
                        {
                            return $"{component} was not imported: it held edits you had not written "
                                + "yet, and importing would have replaced them. They are written to "
                                + "the module now - review them, then import again.";
                        }

                        return WriteModule(component, text, owner, hostRewrite: true, keepEveryCharacter: true);
                    });

                // The folder and mode that were just used become the ones this project
                // remembers, exactly as pressing Apply in the dialog does.
                SaveSyncSettings(LoadSyncSettings().With(syncProjectId, new SyncChoice
                {
                    Folder = folder,
                    ExportMode = importing ? remembered.ExportMode : ModuleSync.ExportModeFrom(syncMode ?? remembered.ExportMode).ToString(),
                    ImportMode = importing ? ModuleSync.ImportModeFrom(syncMode ?? remembered.ImportMode).ToString() : remembered.ImportMode,
                }));

                if (importing && (applied.Changed.Count > 0 || applied.Removed.Count > 0))
                {
                    // An import changes what the project contains, so the tree and the engine
                    // are told, the way they are told when a component is added by hand.
                    PublishProjects();
                    _analysis?.Reanalyse();
                }

                Log.Info($"sync: {(importing ? "import" : "export")} applied: "
                    + $"{applied.Changed.Count} changed, {applied.Skipped.Count} skipped, "
                    + $"{applied.Removed.Count} removed, {applied.Failed.Count} failed");

                return System.Text.Json.JsonSerializer.Serialize(
                    new SyncApplyReply(
                        $"{applied.Changed.Count} changed, {applied.Skipped.Count} skipped, "
                            + $"{applied.Removed.Count} removed, {applied.Failed.Count} failed",
                        [.. applied.Changed],
                        [.. applied.Skipped],
                        [.. applied.Removed],
                        [.. applied.Failed]),
                    SyncJsonContext.Default.SyncApplyReply);
            }
            catch (Exception ex)
            {
                Log.Error("sync: the request failed", ex);
                return System.Text.Json.JsonSerializer.Serialize(
                    new SyncErrorReply($"sync failed: {ex.Message.Trim()}"),
                    SyncJsonContext.Default.SyncErrorReply);
            }
    }

    /// <summary>
    /// The import/export dialog, asking. It goes through the same call the xlide api's `sync` route
    /// goes through, so the dialog cannot be shown a plan the api would not answer, and Apply
    /// cannot leave the project in a state the api would not have left it in.
    /// </summary>
    private void OnSyncRequested(int requestId, IReadOnlyDictionary<string, string> arguments, string body)
    {
        // A PLAN LEAVES THIS THREAD; EVERYTHING ELSE STAYS ON IT.
        //
        // Working out a plan reads the folder and, with the shared planner, waits on the engine:
        // 2,167ms measured on a project of 81,795 lines, every millisecond of it with Excel
        // frozen, because this handler runs on the host user interface thread. The modules have
        // to be read here, because the object model is apartment bound, but nothing after that does.
        //
        // Apply is deliberately left alone: it writes modules through COM, so it belongs on this
        // thread, and a developer who has just pressed Apply is expecting it to work.
        arguments.TryGetValue("action", out var action);
        SyncInputs? inputs = null;
        string? refusal = null;
        var planning = action is null or "plan" && TryGatherSyncInputs(arguments, out inputs, out refusal);

        if (planning)
        {
            var surface = _editorSurface;
            _ = Task.Run(() =>
            {
                string answer;
                try
                {
                    var (plan, planner) = BuildPlan(inputs!);
                    answer = System.Text.Json.JsonSerializer.Serialize(
                        SyncPlanReplyFor(plan, planner),
                        SyncJsonContext.Default.SyncPlanReply);
                }
                catch (Exception ex)
                {
                    Log.Error("sync: the plan could not be worked out", ex);
                    answer = System.Text.Json.JsonSerializer.Serialize(
                        new SyncErrorReply(ex.Message.Trim()),
                        SyncJsonContext.Default.SyncErrorReply);
                }

                // Back to the host thread to answer: posting to the page is its business.
                surface?.RunOnHostThread(() => surface.ShowSyncResult(requestId, answer));
            });

            return;
        }

        string json;
        try
        {
            json = refusal
                ?? (action == "browse" ? ChooseSyncFolder(arguments) : HandleSync(arguments, body));
        }
        catch (Exception ex)
        {
            Log.Error("sync: the dialog's request failed", ex);
            json = System.Text.Json.JsonSerializer.Serialize(
                new SyncErrorReply(ex.Message.Trim()),
                SyncJsonContext.Default.SyncErrorReply);
        }

        _editorSurface?.ShowSyncResult(requestId, json);
    }

    /// <summary>
    /// Raises the system's folder chooser, because a page cannot.
    ///
    /// Modal to the editor's own window, so it cannot end up behind it, and the answer is written
    /// straight into what the project remembers, which is the same state the api's settings route
    /// writes, so choosing a folder either way leaves the product in one place.
    /// </summary>
    private string ChooseSyncFolder(IReadOnlyDictionary<string, string> arguments)
    {
        arguments.TryGetValue("project", out var display);
        arguments.TryGetValue("direction", out var direction);
        var projectId = ProjectIdFromDisplay(display) ?? _shownProject ?? string.Empty;
        var remembered = LoadSyncSettings().For(projectId);

        var chosen = FolderPicker.Choose(
            // The editor's own window, so the chooser is modal to what the developer is looking at
            // rather than to the desktop. Not the page's overlay: that handle is a debug-only
            // affordance for cropping screenshots, and this ships.
            CodePaneTracker.MainWindow(),
            direction == "import" ? "Import modules from this folder" : "Export modules to this folder",
            arguments.TryGetValue("folder", out var startAt) && startAt.Length > 0 ? startAt : remembered.Folder);

        if (chosen is null)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new SyncSettingsReply(projectId, remembered.Folder, remembered.ExportMode, remembered.ImportMode),
                SyncJsonContext.Default.SyncSettingsReply);
        }

        var settings = LoadSyncSettings().With(projectId, remembered with { Folder = chosen });
        SaveSyncSettings(settings);
        var now = settings.For(projectId);
        return System.Text.Json.JsonSerializer.Serialize(
            new SyncSettingsReply(projectId, now.Folder, now.ExportMode, now.ImportMode),
            SyncJsonContext.Default.SyncSettingsReply);
    }

    /// <summary>
    /// The plan as the COMPANION EDITOR'S planner works it out, running in the engine.
    ///
    /// Answers null when the engine cannot be reached or will not answer, and the caller falls
    /// back to the built-in planner. That fallback is deliberate and it is silent by design in one
    /// direction only: a developer pressing Export while the engine is starting should get their
    /// export, not a dialog about which planner is in charge. It is written to the log every time,
    /// so a session that quietly ran on the other planner can still be told apart afterwards.
    /// </summary>
    private SyncPlan? SharedPlan(
        string projectId,
        string displayName,
        string folder,
        IReadOnlyList<LiveModule> live,
        string? mode,
        bool importing)
    {
        if (_analysis?.Engine is not { } engine)
        {
            Log.Info("sync: the engine is not up, so the built-in planner is answering this one");
            return null;
        }

        try
        {
            var modules = live
                .Select(module => new Dictionary<string, object>
                {
                    ["name"] = module.Name,
                    ["type"] = module.Kind,
                    ["source"] = module.Source,
                })
                .ToList();

            using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var answer = engine
                .SyncPlanAsync(
                    importing ? "import" : "export",
                    displayName,
                    folder,
                    mode,
                    modules,
                    cancellation.Token)
                .GetAwaiter()
                .GetResult();

            if (answer is not { } json)
            {
                Log.Warn("sync: the engine answered nothing; the built-in planner is answering this one");
                return null;
            }

            var read = SharedPlanFrom(json, projectId, displayName, folder, mode, importing);
            if (read is null)
            {
                // The one null path that used to say nothing, which is how a green suite ran on
                // the built-in planner while asking for the shared one.
                Log.Warn("sync: the engine's plan could not be read"
                    + $" (it answered {json.ValueKind}, {json.GetRawText().Length} chars);"
                    + " the built-in planner is answering");
            }

            return read;
        }
        catch (Exception ex)
        {
            Log.Warn($"sync: the engine could not plan this ({ex.Message}); the built-in planner is answering");
            return null;
        }
    }

    /// <summary>
    /// Reads their plan into ours.
    ///
    /// The two shapes are nearly the same, which is not a coincidence: the built-in planner was
    /// written from theirs. The mapping is spelled out rather than assumed, because the one field
    /// that is easy to get wrong is the payload. LEFT is the source side in BOTH directions:
    /// exporting, the left is the module and the right is the file; importing, the left is the
    /// file and the right is the module, so the raw left is what an apply writes either way.
    /// </summary>
    private static SyncPlan? SharedPlanFrom(
        JsonElement json,
        string projectId,
        string displayName,
        string folder,
        string? mode,
        bool importing)
    {
        if (!json.TryGetProperty("items", out var items) || items.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        var rows = new List<SyncItem>();
        foreach (var item in items.EnumerateArray())
        {
            var status = StatusFrom(Text(item, "status"));
            var payload = Text(item, "leftRawCode");

            // AN UNCHANGED ROW ARRIVES WITHOUT ITS COMPARISON, AND IS DRAWN FROM ITS OWN TEXT.
            //
            // The engine stops sending line-by-line agreement for a row whose two sides are equal,
            // because that is the whole cost of a plan: 163,000 comparison entries for a project of
            // 81,795 lines, 1,417ms of a 1,710ms plan in the pipe and the JSON either side of it.
            //
            // What is written here is the line that would have survived the condensing anyway, from
            // the same method the built-in planner uses, so an unchanged row is the SAME object
            // whichever planner answered. Anything with a real difference still arrives drawn by
            // their planner, because that picture is of a decision and the decisions are theirs.
            var unchanged = status == SyncStatus.Unchanged;

            rows.Add(new SyncItem
            {
                Id = Text(item, "id"),
                ModuleName = Text(item, "moduleName"),
                ModuleKind = Text(item, "moduleType"),
                FileName = Text(item, "relativeName"),
                Status = status,
                Checked = Flag(item, "checked"),
                Detail = Text(item, "detail", ModuleSync.DetailFor(status)),
                Warning = item.TryGetProperty("warning", out var warning)
                    && warning.ValueKind == JsonValueKind.String
                        ? warning.GetString()
                        : null,
                ExistsInProject = Flag(item, "existsInWorkbook"),
                ExistsInFolder = Flag(item, "existsInRepo"),
                CannotBeCreated = Flag(item, "unsupportedDirectCreation"),
                LeftTitle = Text(item, "leftTitle"),
                RightTitle = Text(item, "rightTitle"),
                Diff = unchanged
                    ? [.. ModuleSync.Identical(ModuleSync.CodeWithoutHeader(payload))]
                    : DiffFrom(item, "diff"),
                DiffWithHeaders = unchanged
                    ? [.. ModuleSync.Identical(payload)]
                    : DiffFrom(item, "diffWithHeaders"),
                PayloadSource = payload,
            });
        }

        var warnings = new List<string>();
        if (json.TryGetProperty("warnings", out var said) && said.ValueKind == JsonValueKind.Array)
        {
            foreach (var line in said.EnumerateArray())
            {
                if (line.ValueKind == JsonValueKind.String && line.GetString() is { Length: > 0 } text)
                {
                    warnings.Add(text);
                }
            }
        }

        return new SyncPlan
        {
            Direction = importing ? SyncDirection.Import : SyncDirection.Export,
            ProjectId = projectId,
            ProjectName = displayName,
            Folder = folder,
            ExportMode = ModuleSync.ExportModeFrom(mode),
            ImportMode = ModuleSync.ImportModeFrom(mode),
            Items = rows,
            Warnings = warnings,
        };
    }

    private static List<SyncDiffLine> DiffFrom(JsonElement item, string name)
    {
        if (!item.TryGetProperty(name, out var lines) || lines.ValueKind != JsonValueKind.Array)
        {
            return [];
        }

        var read = new List<SyncDiffLine>();
        foreach (var line in lines.EnumerateArray())
        {
            read.Add(new SyncDiffLine(
                Number(line, "leftNumber"),
                Number(line, "rightNumber"),
                Text(line, "left"),
                Text(line, "right"),
                Text(line, "kind") switch
                {
                    "changed" => DiffKind.Changed,
                    "added" => DiffKind.Added,
                    "removed" => DiffKind.Removed,
                    _ => DiffKind.Equal,
                }));
        }

        return read;
    }

    /// <summary>
    /// Their status word as ours. A word we do not know falls to Unchanged, and SAYS SO.
    ///
    /// Their vocabulary is seven words and this maps all seven, so the fallback is unreachable
    /// today. It is reachable the moment they add an eighth, and the failure would be the quietest
    /// kind there is: a row that means to do something reads as "already the same", is not ticked,
    /// and the developer is told their project and their folder agree. The planner is imported from
    /// a checkout that moves on its own, which is exactly the situation this cannot be silent in.
    /// </summary>
    private static SyncStatus StatusFrom(string status) => status switch
    {
        "will-create" => SyncStatus.WillCreate,
        "will-write" => SyncStatus.WillWrite,
        "will-update" => SyncStatus.WillUpdate,
        "will-remove" => SyncStatus.WillRemove,
        "skipping-import" => SyncStatus.SkippingImport,
        "read-error" => SyncStatus.ReadError,
        "unchanged" => SyncStatus.Unchanged,
        _ => UnknownStatus(status),
    };

    private static SyncStatus UnknownStatus(string status)
    {
        Log.Warn($"sync: the shared planner said '{status}', which this does not know."
            + " The row is being read as unchanged, which may not be what it means.");
        return SyncStatus.Unchanged;
    }

    private static string Text(JsonElement holder, string name, string fallback = "") =>
        holder.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? fallback
            : fallback;

    private static bool Flag(JsonElement holder, string name) =>
        holder.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;

    private static int? Number(JsonElement holder, string name) =>
        holder.TryGetProperty(name, out var value) && value.TryGetInt32(out var read) ? read : null;

    /// <summary>Where each project's import/export folder is remembered, beside the settings.</summary>
    private static string SyncSettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "xlide_vbide",
        "sync.json");

    private static SyncSettings LoadSyncSettings()
    {
        try
        {
            var path = SyncSettingsPath;
            return File.Exists(path) ? SyncSettings.Parse(File.ReadAllText(path)) : SyncSettings.Empty;
        }
        catch (Exception ex)
        {
            Log.Info($"sync: the remembered folders could not be read ({ex.GetType().Name})");
            return SyncSettings.Empty;
        }
    }

    private static void SaveSyncSettings(SyncSettings settings)
    {
        try
        {
            var path = SyncSettingsPath;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, settings.ToJson());
        }
        catch (Exception ex)
        {
            Log.Error("sync: the folder could not be remembered; it holds for this session only", ex);
        }
    }

    /// <summary>
    /// The plan as it goes over the wire. One mapping, so the dialog and the api are looking at the
    /// same rows in the same order with the same words on them.
    /// </summary>
    private static SyncPlanReply SyncPlanReplyFor(SyncPlan plan, string planner) => new(
        plan.Direction == SyncDirection.Import ? "import" : "export",
        plan.ProjectName,
        plan.ProjectId,
        plan.Folder,
        plan.Direction == SyncDirection.Import
            ? plan.ImportMode == ImportMode.TrueUpStandardClass ? "trueUpStandardClass" : "updateOnly"
            : plan.ExportMode == ExportMode.TrueUp ? "trueUp" : "exportAll",
        planner,
        [.. plan.Items.Select(item => new SyncItemRow(
            item.Id,
            item.ModuleName,
            item.ModuleKind,
            item.FileName,
            ModuleSync.NameOf(item.Status),
            item.Checked,
            item.Detail,
            item.Warning,
            item.ExistsInProject,
            item.ExistsInFolder,
            item.CannotBeCreated,
            item.LeftTitle,
            item.RightTitle,
            // CONDENSED ON THE WAY OUT, not in the plan itself: what a row DOES is decided from
            // the whole comparison, and only what is SHOWN is shortened. Whole, a project of
            // 81,795 lines answered 15MB of which every byte was comparison lines, for a dialog
            // that shows one row at a time (2026-08-09).
            [.. ModuleSync.Condense(item.Diff).Select(SyncDiffRowFor)],
            [.. ModuleSync.Condense(item.DiffWithHeaders).Select(SyncDiffRowFor)]))],
        [.. plan.Warnings]);

    private static SyncDiffRow SyncDiffRowFor(SyncDiffLine line) =>
        new(line.LeftNumber, line.RightNumber, line.Left, line.Right, ModuleSync.NameOf(line.Kind));

    private static ProductSettings LoadSettings()
    {
        try
        {
            var path = SettingsPath;
            return File.Exists(path) ? ProductSettings.Parse(File.ReadAllText(path)) : ProductSettings.Default;
        }
        catch (Exception ex)
        {
            Log.Info($"settings: could not be read, using defaults ({ex.GetType().Name})");
            return ProductSettings.Default;
        }
    }

    /// <summary>
    /// A change from the page's dialog: adopted, written through, and echoed back - the echo is
    /// the page's confirmation that the choice will survive the session.
    /// </summary>
    private void OnSettingsChanged(ProductSettings updated)
    {
        _settings = updated.Normalized();

        // The analyzer's copy of the machine-wide rule choices follows the record wherever the
        // change came from - the dialog, the door, or a hand edit adopted at the next change.
        _analysis?.UseSeverityOverrides(_settings.AnalysisRuleSeverityOverrides);

        SaveSettings();
        Log.Info($"settings: saved (blockLayout {_settings.BlockLayout}"
            + $", continueComment {_settings.ContinueCommentOnNewline}"
            + $", mirrorSpacing {_settings.MirrorCommentSpacing})");

        _editorSurface?.ShowSettings(_settings);
    }

    private bool _stopped;

    public AddInSession(DispatchObject editor, DispatchObject? addIn)
    {
        _editor = editor;
        _addIn = addIn;

        // The inventory reads each bare control's type library through the panel's own reader, so
        // a class is walked once for both. A field initializer cannot say this - it may not name
        // another field - which is the only reason it is here.
        _controlDefaults = new ControlDefaults(_propertyTypes);
    }

    /// <summary>Automation object for the editor itself.</summary>
    public DispatchObject Editor => _editor;

    /// <summary>
    /// Whether a component is the evaluator's scratch module, which briefly exists during every
    /// Immediate evaluation and must never reach a tab, the explorer, the properties panel, or
    /// the editor: it surfacing is what made every evaluation flash the screen.
    /// </summary>
    private static bool IsScratchComponent(string? name) =>
        string.Equals(name, ImmediateEvaluator.ScratchModule, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Walks a project's components, handing each real one to `take` with its name. Two rules,
    /// each stated once because they are correctness properties, not conveniences: the
    /// evaluator's scratch module is skipped - a fixture that counts it counts wrong, and so
    /// does a tree that lists it - and an entry that cannot be read costs that entry alone,
    /// never the walk. The walk was hand-rolled three times (twice behind the xlide api, once
    /// in the publish path) with the filter spelled positively in one and negatively in
    /// another, and the projects route's copy had the containment wrong: its per-component
    /// read sat inside the try wrapping the whole project, so one unreadable component dropped
    /// the entire project from the reply (the audit's B18).
    /// </summary>
    private static void ForEachRealComponent(DispatchObject project, Action<DispatchObject, string> take)
    {
        using var components = project.GetObject("VBComponents");
        var count = components?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            try
            {
                using var component = components!.GetItem(i);
                if (component?.GetString("Name") is { Length: > 0 } name && !IsScratchComponent(name))
                {
                    take(component, name);
                }
            }
            catch (Exception ex)
            {
                Log.Verbose($"components: entry {i} could not be read ({ex.GetType().Name})");
            }
        }
    }

    /// <summary>
    /// Whether the editor is stopped inside the module this product evaluates lines in.
    ///
    /// The one state that is safe to end without asking. A developer stopped at their own
    /// breakpoint is stopped in one of THEIR modules; only this product ever puts code in the
    /// scratch module, so a session stopped there is one nobody is looking at and one that will
    /// otherwise refuse every evaluation that follows.
    /// </summary>
    private bool StoppedInScratchModule()
    {
        try
        {
            using var project = _editor.GetObject("ActiveVBProject");
            if ((project?.GetInt32("Mode") ?? DesignMode) == DesignMode)
            {
                return false;
            }

            using var pane = _editor.GetObject("ActiveCodePane");
            using var code = pane?.GetObject("CodeModule");
            using var component = code?.GetObject("Parent");

            return IsScratchComponent(component?.GetString("Name"));
        }
        catch (Exception ex)
        {
            // Unreadable answers no. A reset that cannot be justified is worse than one skipped.
            Log.Info($"immediate: could not tell where the editor is stopped ({ex.GetType().Name})");
            return false;
        }
    }

    /// <summary>
    /// Whether the editor is stopped inside this product's scratch module, asked from ANY thread.
    ///
    /// Reading the editor is COM, so the question hops to the host thread and comes straight back.
    /// It is deliberately a hop per ask rather than a cached flag: the whole point of asking
    /// repeatedly is that the answer changes underneath, and a cached one would be the reason the
    /// loop never ends.
    /// </summary>
    private bool ScratchBreakStanding()
    {
        if (_editorSurface is not { } surface)
        {
            return false;
        }

        var answered = new ManualResetEventSlim(false);
        var standing = false;

        surface.RunOnHostThread(() =>
        {
            try
            {
                standing = StoppedInScratchModule();
            }
            finally
            {
                answered.Set();
            }
        });

        // A host thread that cannot answer in two seconds is busy with something that is not this,
        // and reporting "not stuck" lets the evaluation try and report its own failure honestly.
        var got = answered.Wait(2000);
        answered.Dispose();
        return got && standing;
    }

    /// <summary>Takes the scratch module away, which a reset on its own does not.</summary>
    private void RemoveScratchModule()
    {
        try
        {
            using var project = _editor.GetObject("ActiveVBProject");
            using var components = project?.GetObject("VBComponents");
            if (components is null)
            {
                return;
            }

            // By index, the way the evaluator's own clean-up does it: `Item` by name throws when
            // the module is absent, which is the ordinary case and not worth an exception.
            var count = components.GetInt32("Count");
            for (var i = count; i >= 1; i--)
            {
                using var candidate = components.GetItem(i);
                if (IsScratchComponent(candidate?.GetString("Name")) && candidate is not null)
                {
                    components.InvokeWithObject("Remove", candidate);
                    Log.Info("immediate: the scratch module has been taken away");
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"immediate: the scratch module could not be taken away ({ex.GetType().Name})");
        }
    }

    public void Start()
    {
        Log.Info("session starting");
        ReportEnvironment();
        Log.Info("session started");
    }

    /// <summary>Called once the host has finished its own startup and the object model is settled.</summary>
    public void HostStartupComplete()
    {
        // THE ENGINE GOES FIRST, because it is the longest pole in the session by two orders of
        // magnitude and everything else here is microseconds against it.
        //
        // Measured 2026-08-11 on a real ALT+F11: the page is up and usable at 574ms, and the engine
        // is not connected until 3,370ms - so for 2.8 seconds the editor is on screen with no
        // diagnostics, no completions and no hover, which is what "xlide is slow to load" actually
        // describes. The cost is reading a 90 MB Node image off disk while Excel, the VBE and a
        // WebView2 browser process are all starting at once; standalone the same launch is 1.27s
        // cold and 190ms warm.
        //
        // Nothing below depends on it: Start only spawns the process and returns, and the first
        // pass runs from the connect callback seconds later, by which time both calls under this
        // have long finished. So starting it before two COM enumerations rather than after them
        // costs nothing and takes their duration off a wait the developer is sitting through.
        StartAnalysis();

        ReportOpenProjects();
        TrackCodePanes();

        // An editor with no panes at all never fires the pane events the surface normally
        // arrives on, and a developer opening a fresh workbook's editor would meet the native
        // gray shell. The surface goes up now, showing the empty workspace and the explorer.
        TryShowEmptyWorkspace();

        // The local door: state, windows, commands by name. Shipped in every build and OPENED in
        // none of them unless someone said so - see ApiWanted for which way each build leans.
        //
        // The lean is LOGGED, and said as one of two whole phrases rather than an interpolated
        // word, because `ApiOpenUnlessTold` is a const: the compiler folds the ternary and only
        // the taken phrase survives into the binary. So `verify.ps1` can read a published shim
        // and prove which way it leans, instead of trusting that the `#if` was written the right
        // way round - which is the one thing about this change worth proving mechanically.
        Log.Info(ApiOpenUnlessTold
            ? "xlide api: this build opens the door unless told otherwise"
            : "xlide api: this build keeps the door shut unless told otherwise");

        OpenOrCloseApi(ApiWanted(_settings));
    }

    /// <summary>
    /// Whether the api door should be listening, from the setting and this build's own lean.
    ///
    /// A dev build leans OPEN because the harness is the reason it exists; a shipped build leans
    /// SHUT because a door nobody asked for should not be listening. An explicit true or false in
    /// the settings file outranks the lean in both.
    /// </summary>
    /// <summary>
    /// Which way this build leans when the settings file has not said.
    ///
    /// THE ONE PLACE THE TWO BUILDS DIFFER, and it is a named constant rather than a bare `#if`
    /// around the expression so that it reads as a value and cannot be flattened by accident: a
    /// pass over this file that strips DEBUG directives and keeps the DEBUG branch turns a bare
    /// conditional into `true` and ships the door open, which is exactly the failure this whole
    /// change exists to prevent. It happened, to me, within a minute of writing it (2026-08-22).
    /// `verify.ps1` asserts the shipped default rather than trusting this line.
    /// </summary>
    private const bool ApiOpenUnlessTold =
#if DEBUG
        true;
#else
        false;
#endif

    private static bool ApiWanted(ProductSettings settings) => settings.ApiEnabled ?? ApiOpenUnlessTold;

    /// <summary>
    /// Opens the api door, or shuts it. Idempotent, and safe to call before the session is up.
    ///
    /// BOTH DOORS MOVE TOGETHER. The HTTP door and the in-process one (GetObject(, "Xlide.Api"))
    /// are one capability wearing two coats - the same routes, the same authority over the same
    /// project - so a switch that moved only one of them would be a switch that lies.
    /// </summary>
    private void OpenOrCloseApi(bool open)
    {
        if (open == (_apiServer is not null))
        {
            return;
        }

        if (open)
        {
            _apiServer = ApiServer.Start(AnswerApiRequest);
            OfferInsideDoor();
            return;
        }

        RetireInsideDoor();
        _apiServer?.Dispose();
        _apiServer = null;
        Log.Info("xlide api: the door is shut");
    }

    /// <summary>
    /// Creates the surface over a frame if it is not already there, and wires it. A frame change
    /// means a new surface rather than a move: the surface belongs to one parent.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.MemberNotNullWhen(true, nameof(_editorSurface))]
    private bool EnsureSurfaceForFrame(nint host)
    {
        if (_editorSurface is not null && _editorSurface.Host != host)
        {
            Log.Info("editor surface: the document area changed, rebuilding");
            _editorSurface.Dispose();
            _editorSurface = null;

            _frameSubclass?.Dispose();
            _frameSubclass = null;
        }

        if (_editorSurface is not null)
        {
            return true;
        }

        _editorSurface = EditorSurface.Create(host, default);
        if (_editorSurface is null)
        {
            return false;
        }

        // In the frame's message chain, so a resize re-places the surface synchronously -
        // before the native layout paints - instead of a posted event later. The event route
        // stays as the correcting pass.
        _frameSubclass ??= FrameSubclass.Install(host, PlaceSurfaceFast);

        _editorSurface.KeyPressed = OnSurfaceKey;
        // The page asked for a module that is gone: say so on the surface rather than leaving the
        // developer looking at the module they were already on, wondering whether the click landed.
        _editorSurface.ModuleRequested = (component, project, face) =>
        {
            // A designer tab is this product's own state, not a mirror of a native pane: the
            // native designer window stays down (the Toolbox trap), so activating one changes
            // nothing in the object model and everything in the published strip.
            if (face == "design")
            {
                OpenDesignerTab(component, project);
                return;
            }

            // The developer asked for the CODE tab, which must take the active slot back even
            // when the native pane underneath was already this module and nothing else moves.
            _activeDesignerTab = null;
            if (ShowModule(component, project) is { } missing)
            {
                _editorSurface?.Notify(missing);
            }
        };
        // The page's own navigation shows its failure on screen, so it discards the complaint the
        // door reads. (A method group cannot bind here now that GoTo answers one.)
        _editorSurface.NavigateRequested = (component, line, column, project) =>
            GoTo(component, line, column, project);
        _editorSurface.CommandRequested = RunCommand;
        // The document names its workbook by display name; the write path resolves that to the
        // project identity so a same-named module in another workbook is never the one written.
        //
        // A REFUSED WRITE IS SAID OUT LOUD HERE. This is the developer's own typing on its way to
        // the module, and a write the editor will not take means the text on screen is not the text
        // that will run, compile or save. Silence let them go on typing into a surface that had
        // stopped being the truth.
        _editorSurface.TextChanged = (component, project, text) =>
        {
            var refused = WriteModule(component, text, ProjectIdFromDisplay(project));
            var key = WrittenKey(component, project);

            if (refused is null)
            {
                _saidAboutWrite.Remove(key);
                return;
            }

            if (_saidAboutWrite.TryGetValue(key, out var already) && already == refused)
            {
                return;
            }

            _saidAboutWrite[key] = refused;
            _editorSurface?.Notify(refused);
        };
        _editorSurface.BreakpointToggleRequested = ToggleBreakpoint;
        _editorSurface.LinesShifted = OnLinesShifted;
        _editorSurface.SearchRequested = OnSearchRequested;
        _editorSurface.ReplaceAllRequested = OnReplaceAllRequested;
        _editorSurface.Polled = PollDebugState;
        _editorSurface.PlacementSettled = RefreshSurfacePlacement;
        _editorSurface.EvaluateRequested = line => EvaluateImmediate(line);
        _editorSurface.ExternalOpenRequested = OpenExternal;
        _editorSurface.RenameUndoRequested = id => _editorSurface?.RunOnHostThread(() => UndoRename(id));
        _editorSurface.DocumentRequested = PublishDocument;
        _editorSurface.FormMarkupRequested = PublishFormMarkup;
        _editorSurface.FormMarkupApplyRequested = ApplyFormMarkup;
        _editorSurface.DesignerEventStubRequested = OnDesignerEventStub;
        _editorSurface.DesignerZOrderRequested = OnDesignerZOrder;
        _editorSurface.TestsActionRequested = (action, target, file, tags, outcomes) =>
            _editorSurface?.RunOnHostThread(() => OnTestsAction(action, target, file, tags, outcomes));

        // "Size to Fit" on the canvas. The measurement is the host's because MSForms' AutoSize is
        // the only thing that knows what a control's natural size is - see MeasureAutoSize - and
        // the ANSWER goes back to the page, which writes it into the document. So the form is
        // touched to measure and never to change: the size reaches it at Ctrl+S like every other
        // gesture, and a Ctrl+Z before that leaves nothing behind.
        _editorSurface.DesignerAutoSizeRequested = (id, module, project, control) =>
            _editorSurface?.RunOnHostThread(() =>
            {
                using var component = FindComponent(module, ProjectIdFromDisplay(project), out _);
                if (component is null)
                {
                    _editorSurface?.ShowDesignerAutoSize(id, null, null);
                    return;
                }

                var (width, height, refused) = FormDesignService.MeasureAutoSize(component, module, control);
                if (refused is not null)
                {
                    Log.Info($"designer: autosize for {module}.{control} answered nothing ({refused})");
                }

                _editorSurface?.ShowDesignerAutoSize(id, width, height);
            });
        _editorSurface.DesignerSetPropertyRequested = OnDesignerSetProperty;
        _editorSurface.DesignerSelectionRequested = (module, project, control) =>
            _editorSurface?.RunOnHostThread(() =>
            {
                // The canvas said what it selected; the panel follows. An empty control is
                // the form's ground - the component itself, the row set that always existed.
                _propertiesTarget = module;
                _propertiesControl = string.IsNullOrEmpty(control) ? null : control;
                PublishProperties();
            });
        // Pure text both ways: the lint is Core's tolerant parse, no designer is touched,
        // and the answer goes straight back - the one language, saying early what the
        // apply would say late.
        _editorSurface.FormMarkupLintRequested = (module, project, markup) =>
        {
            // The tolerant pass collects the squiggles; the strict parse - the ONE grammar,
            // the apply's own - yields the draft the canvas previews as the developer types.
            // A text the apply would refuse has no draft, and the canvas holds its last
            // picture rather than blanking under a half-typed line.
            Core.Forms.FormSpec? draft = null;
            try
            {
                draft = Core.Forms.FormMarkup.Parse(markup);
            }
            catch (Core.Forms.FormMarkupException)
            {
                // The findings already carry the refusal, line and all.
            }

            _editorSurface?.PublishFormMarkupLint(module, project, Core.Forms.FormMarkup.Lint(markup), draft);
        };
        _editorSurface.FormMarkupVocabularyRequested = (module, project) =>
            _editorSurface?.RunOnHostThread(() => PublishFormMarkupVocabulary(module, project));
        _editorSurface.PanelChanged = OnPanelChanged;
        _editorSurface.MenuRequested = OnMenuRequested;
        _editorSurface.MenuExecuteRequested = OnMenuExecuteRequested;
        _editorSurface.PropertyEditRequested = OnPropertyEdit;
        _editorSurface.PicturePickRequested = OnPicturePick;
        _editorSurface.ComponentSelected = OnComponentSelected;
        // The tab's X does not read the outcome: it sees its tab go, or sees the question. The
        // xlide api's caller has neither, so the method answers and this discards.
        _editorSurface.ModuleCloseRequested = (component, project, action, face) =>
        {
            // A designer tab's unapplied edits live in the PAGE, so nothing here can know to
            // hold the close - and the page asks its own question before sending one, which it
            // did not until 2026-08-16 (this comment claimed it did; the hunt found the claim
            // false and a Ctrl+W losing three moves with nothing said). By the time a close for
            // a designer tab arrives here it has been answered, so it is unconditional.
            if (face == "design")
            {
                CloseDesignerTab(component, project);
                return;
            }

            OnModuleCloseRequested(component, project, action);
        };
        _editorSurface.ComponentInsertRequested = InsertComponent;
        _editorSurface.HostActionRequested = OnHostActionRequested;
        _editorSurface.ComponentRemoveRequested = (component, project) =>
        {
            if (RemoveComponent(component, project) is { } refused)
            {
                _editorSurface?.Notify(refused);
            }
        };
        _editorSurface.CompletionRequested = OnCompletionRequested;
        _editorSurface.HoverRequested = OnHoverRequested;
        _editorSurface.SignatureHelpRequested = OnSignatureHelpRequested;
        _editorSurface.SmartEnterRequested = OnSmartEnterRequested;
        _editorSurface.CanonicalCaseRequested = OnCanonicalCaseRequested;
        _editorSurface.LoopSyncRequested = OnLoopSyncRequested;
        _editorSurface.CodeActionsRequested = OnCodeActionsRequested;
        _editorSurface.AnalysisRulesRequested = OnAnalysisRulesRequested;
        _editorSurface.SuppressFindingRequested = OnSuppressFindingRequested;
        _editorSurface.RuleSeverityChangeRequested = (code, severity) =>
            _ = Task.Run(async () =>
            {
                // SAID ON SCREEN, whatever happened. A page-initiated change has no reply
                // channel of its own, and the difference between "off everywhere" and "cannot
                // be off, the analyzer allows warning" is exactly what the person clicking a
                // menu item needs to hear - silence on the refusal was how the first build made
                // an always-on rule's Turn Off look like a button that did nothing.
                var said = await ApplyRuleSeverityAsync(code, severity).ConfigureAwait(false);

                // On the browser's own thread, like every message to the page: posted from this
                // pool thread the notice answers UI_E_WRONG_THREAD and never arrives - measured,
                // the refusal for an always-on rule was sent and nothing showed.
                var surface = _editorSurface;
                surface?.RunOnHostThread(() => surface.Notify(said));
            });
        _editorSurface.NavigationRequested = OnNavigationRequested;
        _editorSurface.RenameRequested = OnRenameRequested;
        _editorSurface.ModuleRenameRequested = OnModuleRenameRequested;
        _editorSurface.ExtractMethodRequested = OnExtractMethodRequested;
        _editorSurface.ImplementInterfaceRequested = OnImplementInterfaceRequested;
        _editorSurface.EncapsulateFieldRequested = OnEncapsulateFieldRequested;
        _editorSurface.ExtractVariableRequested = OnExtractVariableRequested;
        _editorSurface.InlineVariableRequested = OnInlineVariableRequested;
        _editorSurface.MoveToModuleRequested = OnMoveToModuleRequested;
        _editorSurface.IntroduceParameterRequested = OnIntroduceParameterRequested;
        _editorSurface.OutlineRequested = OnOutlineRequested;
        _editorSurface.SyncRequested = OnSyncRequested;
        _editorSurface.ChangesRequested = OnChangesRequested;
        _editorSurface.ApiRequested = OnApiRequested;
        _editorSurface.SemanticTokensRequested = OnSemanticTokensRequested;
        _editorSurface.LiveAnalysisDue = OnLiveAnalysisDue;
        _editorSurface.LiveTextPushed = (module, full, edits) => _analysis?.NotifyLiveText(module, full, edits);

        // The active-line hold: typing on a line hides the verdicts about it, and the caret
        // settling anywhere else brings them back. Both handlers run on the host thread, and
        // both republish only when the hold actually changed - a keystroke on an already-held
        // line and a caret resting where it was cost nothing. When it did change, the markers
        // that can differ belong to the module released and the module now held, so only those
        // are re-sent; the previously held name is read before the mutation because that is the
        // one Begin overwrites and Release erases.
        _editorSurface.LineTyped = line =>
        {
            var previouslyHeld = _activeLineHold.Module;
            if (_editorSurface?.Module is { } typedModule && _activeLineHold.Begin(typedModule, line))
            {
                PublishMarkersToSurface(previouslyHeld, typedModule);
                PublishFindingsToSurface();
            }
        };
        _editorSurface.CaretLineSettled = line =>
        {
            var previouslyHeld = _activeLineHold.Module;
            if (_activeLineHold.Release(_editorSurface?.Module, line))
            {
                PublishMarkersToSurface(previouslyHeld);
                PublishFindingsToSurface();
            }
        };

        // The moment the page is up is the moment the menu bar can be covered, and it is not a
        // window event, so nothing else would recompute the bounds. The settings ride the same
        // moment: the page's typing behaviour starts from what the developer chose last time.
        _settings = LoadSettings();
        _editorSurface.SettingsChangeRequested = OnSettingsChanged;
        _editorSurface.Ready = () =>
        {
            RefreshSurfacePlacement();
            _editorSurface?.ShowSettings(_settings);

            // A ready can be a RELOADED page, not only the first boot. The surface just
            // re-opened every live document from its own table; everything else the page
            // draws is re-said here, because the first boot got it from the held-message
            // replay and a second ready has nothing held (found live 2026-08-06: a reload
            // came back with models and no tabs, and later without its properties pane).
            //
            // The what-was-last-sent caches reset FIRST: they exist to spare a page that
            // already has the picture, and a page that just booted has nothing. Without
            // this, PublishModules compared against the pre-reload list, matched, and sent
            // nothing - the tabs stayed gone (the tap showed every republish EXCEPT
            // setModules, 2026-08-06).
            _lastModulesKey = null;
            _lastLanguageFactsKey = null;
            _editorSurface?.ShowInstallPath(Interop.ShimModule.Directory);
            _editorSurface?.ShowSystemColours();
            _hostChrome ??= HostChrome.Install(CodePaneTracker.MainWindow(), Interop.ShimModule.Directory);
            PublishModules();
            PublishProjects();
            PublishFindingsToSurface();
            PublishMarkersToSurface();
            PublishBreakpoints();
            PublishProperties();
            UpdateDebugState();

            // SAY THAT THE LANGUAGE IS STILL COMING. The surface is usable about 2.8 seconds
            // before the engine answers anything (lesson 64), and an editor that looks finished
            // and returns no completions reads as broken. Only when it is genuinely not up yet:
            // a reload of the page mid-session finds the engine already running and says nothing.
            if (_analysis is { EngineIsUp: false })
            {
                _editorSurface?.Hold("Starting analysis. Problems, completions and hover will "
                    + "arrive in a moment.");
            }

            InstallConsoleRing();
        };

        // While the loader shows, placement is re-asserted on its heartbeat: the editor is still
        // arranging itself - restoring its size, raising its own bands - and with no pane open
        // there is no window event to notice any of it. Without this, the loader keeps covering
        // the window as it was at the first placement, and a band of native chrome outlives it.
        _editorSurface.LoadingPulse = RefreshSurfacePlacement;

        // THE LOADER GOES UP BEFORE THE SLOW PART, which is everything below this line.
        //
        // The surface exists by now and its overlay knows how to paint the loading screen, but
        // nothing had put it on screen yet: placement was first asked for much later, from the
        // empty-workspace path or from the first pane event. Between here and there the host
        // thread does the work that costs the visible half-second - hiding the editor's own
        // windows, floating and ghosting both debug palettes, starting their reader - and it does
        // all of it without returning to the message loop. So the developer pressed ALT+F11 and
        // watched the editor's frozen native chrome, which reads as the host having hung.
        //
        // Placed and painted here instead, so that time is spent looking at xlide starting.
        // Placement is re-asserted constantly afterwards - the loader's own pulse, the frame
        // subclass, every pane event - so a first placement made a moment before the windows
        // below move is corrected within a tick rather than being load-bearing.
        RefreshSurfacePlacement();

        // Now rather than at start-up. The editor answers that these windows are visible before
        // it has created them, so hiding one then closes something with no window behind it and
        // there is nothing to identify afterwards.
        HideReplacedWindows();
        HideNativeToolbars();
        PrepareLocalsGhost();
        PrepareWatchGhost();

        // Both ghosts are read from one dedicated thread; the host thread only asks and looks.
        // Reading from here re-enters the editor's own accessibility provider and dies in
        // native faults - GhostReaderThread carries the story.
        _ghostReaders = GhostReaderThread.Start(_localsPalette, _watchPalette);

        DarkenTitleBar(host);

        return true;
    }

    /// <summary>True while the surface is up over an editor that has no panes anywhere.</summary>
    private bool _watchingEmpty;

    /// <summary>
    /// Whether the editor's own window is on screen, as placement last saw it. Defaults true so a
    /// session that has not placed anything yet behaves as it always did.
    ///
    /// The polling tiers consult it: an editor whose window has been closed needs nothing watched,
    /// and `_watchingEmpty` alone does not notice, because it is a fact about the workspace and
    /// survives the window that was showing it.
    /// </summary>
    private bool _frameVisible = true;

    /// <summary>The language facts last pushed to the page, so unchanged ones are not re-sent.</summary>
    private string? _lastLanguageFactsKey;

    /// <summary>When the full pass last ran, and whether one is owed from a skipped turn.</summary>
    private long _lastFullAnalysis;
    private bool _fullAnalysisDeferred;
    private const long FullAnalysisQuietMilliseconds = 3000;

    /// <summary>Which panes existed and showed when they were last logged. See TrackCodePanes.</summary>
    private string? _lastPaneComposition;

    /// <summary>In the frame's message chain while the surface lives. See EnsureSurfaceForFrame.</summary>
    private FrameSubclass? _frameSubclass;

    /// <summary>
    /// Puts the surface over an editor that has no panes at all, which is what a fresh
    /// workbook's editor is. The surface normally arrives with the first pane; without this, an
    /// editor with no module yet stays native gray, and the explorer shown here is how the first
    /// module gets opened or inserted at all.
    /// </summary>
    /// <summary>
    /// Leaves the empty workspace when the object model says modules are open after all.
    ///
    /// The empty view and the tab strip answer to different authorities: the strip is the
    /// object model's open list, while the shown document waits for the PANE TRACKER to
    /// follow a pane. The tracker only ever holds the pane windows it can match by caption
    /// (lesson 28), and when it can match none - every pass "saw 0" while CodePanes held two
    /// - nothing ever showed a module, so the surface drew tabs above its own "No module is
    /// open" (the developer, 2026-08-06). The object model is the authority on whether
    /// anything is open at all, so the empty view defers to it: the active pane's module goes
    /// up, and the workspace stops calling itself empty.
    /// </summary>
    private void AdoptOpenModuleIfEmpty()
    {
        if (!_watchingEmpty || _editorSurface is null || _editorSurface.Module is not null)
        {
            return;
        }

        if (ReadOpenModules() is not { Count: > 0 } open)
        {
            return;
        }

        // The active pane names the module the developer would expect to be looking at; with
        // no active pane the first open one is as good an answer as the editor has.
        var component = ActiveComponentName() ?? open[0].Name;

        // ReadOpenModules carries the workbook by the DISPLAY name the tab strip shows, and
        // everything downstream of here addresses projects by ID. Handing the display name
        // straight on found no component and returned in silence, leaving the very empty view
        // this method exists to clear (2026-08-06).
        var display = open.FirstOrDefault(entry =>
            string.Equals(entry.Name, component, StringComparison.OrdinalIgnoreCase)).Project;
        var project = ProjectIdFromDisplay(display) ?? ActivePaneOwner(component);

        Log.Info($"editor surface: {open.Count} module(s) are open after all, showing {component}");

        _watchingEmpty = false;
        ShowModuleInSurface(component, project);
        PublishModules();
        UpdatePolling();
    }

    /// <summary>The active code pane's component name, or null when there is no active pane.</summary>
    private string? ActiveComponentName()
    {
        try
        {
            using var active = _editor.GetObject("ActiveCodePane");
            using var module = active?.GetObject("CodeModule");
            using var component = module?.GetObject("Parent");
            return component?.GetString("Name");
        }
        catch (Exception ex)
        {
            Log.Verbose($"editor surface: no active pane to adopt ({ex.GetType().Name})");
            return null;
        }
    }

    private void TryShowEmptyWorkspace()
    {
        if (_editorSurface is not null || _stopped)
        {
            return;
        }

        var frame = CodePaneTracker.FindFrame();
        if (frame == 0)
        {
            return;
        }

        var documentArea = Win32.FindWindowEx(frame, 0, "MDIClient", null);
        if (documentArea == 0)
        {
            return;
        }

        _frame = frame;
        _documentArea = documentArea;

        if (!EnsureSurfaceForFrame(frame))
        {
            return;
        }

        _surfaceShown = true;
        _watchingEmpty = true;
        _editorSurface!.Clear();
        _editorSurface.ShowModules([], [], null, null);
        PublishProjects();
        RefreshSurfacePlacement();
        UpdatePolling();

        Log.Info("editor surface: opened onto an editor with no panes");
    }

    /// <summary>
    /// Takes the user to a finding: the native pane is selected and the caret placed on it, and the
    /// surface over that pane scrolls to match.
    ///
    /// The native pane is moved as well as the surface, because it stays the text of record and
    /// what the debugger drives. Leaving it where it was would put the two out of step the first
    /// time the user pressed F8.
    /// </summary>
    /// <returns>
    /// Null when the caret is where it was asked to be, otherwise what stopped it. A navigation
    /// that finds no pane leaves the caret in the module that was already shown, and a caller that
    /// goes on to set a breakpoint or press Run then acts on the wrong module. The developer sees
    /// that happen; a script told nothing does not.
    /// </returns>
    private string? GoTo(string component, int line, int column, string? projectDisplay = null)
    {
        try
        {
            // A NAMED workbook that resolves to nothing is refused here rather than dropped to
            // null, because null means "whichever holds this module" - so a caret or a breakpoint
            // aimed at a misspelt workbook would land in another one that happens to have a module
            // of the same name, which is the case this argument exists to prevent.
            var projectId = ProjectIdFromDisplay(projectDisplay);
            if (projectId is null && projectDisplay is { Length: > 0 })
            {
                Log.Info($"navigate: no workbook answers to '{projectDisplay}'");
                return $"no open workbook answers to '{projectDisplay}'";
            }

            using var pane = FindCodePane(component, projectId);
            if (pane is null)
            {
                Log.Info($"navigate: no pane for {component}");
                return $"no code pane for '{component}'"
                    + (projectDisplay is { Length: > 0 } ? $" in {projectDisplay}" : string.Empty);
            }

            pane.Invoke("Show");
            pane.Invoke("SetSelection", line, column, line, column);

            // The surface shows one (module, WORKBOOK) pair, and a module name is not unique
            // across workbooks: two open workbooks can each hold a Helpers. Comparing the name
            // alone made a navigation to the other workbook's Helpers a no-op on the surface -
            // the native pane moved and the surface stayed where it was, showing a different
            // workbook's module of the same name (2026-08-07). A navigation that names no
            // project keeps the old meaning: whichever one is shown.
            // AND A DESIGNER TAB IS NOT ITS FORM'S CODE. "The shown module" carries a name and a
            // workbook and no FACE, so a form whose designer tab holds the active slot reported
            // that module as already showing: the branch below revealed a line in a document
            // nobody was looking at, and clicking a Sub under a form in the tree did nothing
            // visible at all (the owner, 2026-08-18). The code face has to take the slot first,
            // which is exactly what a click on the code tab does.
            var designerHoldsTheSlot = _activeDesignerTab is not null;

            var alreadyShowing =
                !designerHoldsTheSlot
                && string.Equals(_editorSurface?.Module, component, StringComparison.OrdinalIgnoreCase)
                && (projectId is null
                    || string.Equals(_shownProject, projectId, StringComparison.OrdinalIgnoreCase));

            if (alreadyShowing)
            {
                _editorSurface!.Reveal(line);
            }
            else if (_editorSurface is not null)
            {
                // TAKING THE SLOT BACK FROM WHICHEVER DESIGNER TAB HELD IT, not only this form's.
                // PublishModules clears this when the NATIVE active pane moves, and here it does
                // not move - the form's code pane was already the active one underneath, which is
                // why the tab never changed. It also decides where F5 and Ctrl+S are aimed, so a
                // navigation that leaves it standing leaves Run pointed at a form the developer
                // has navigated away from: with FormA's tab up, going to Module1 already put the
                // code tab on screen and left Run applying FormA.
                _activeDesignerTab = null;

                // A navigation can target a module the surface has never shown. The Show
                // above opens it in the object model, but from the empty workspace that
                // open arrives with no window event the tracker can hear, and the surface
                // sat on the empty view while the object model insisted the pane was
                // active (2026-08-05, caught by the semantic probe on its first run). The
                // module goes onto the surface the way the explorer puts one there.
                ShowModuleInSurface(component, projectId);
                PublishModules();
                _editorSurface.Reveal(line);
            }

            // Activating the native pane moved keyboard focus onto it, and the developer is not
            // looking at it. Without this, every key after a navigation went to the covered
            // pane: Ctrl+W closed nothing, typing typed nowhere visible.
            _editorSurface?.Focus();

            Log.Info($"navigate: {component}({line},{column})");
            return null;
        }
        catch (Exception ex)
        {
            Log.Error($"navigate: could not go to {component}({line},{column})", ex);
            return $"the navigation to '{component}' raised {ex.GetType().Name}";
        }
    }

    /// <summary>
    /// Brings up the analysis engine and reports what it finds.
    ///
    /// Started, not awaited. The host is still finishing its own start-up at this point and nothing
    /// here is worth delaying that for; findings arrive when they arrive.
    /// </summary>
    private void StartAnalysis()
    {
        try
        {
            _analysis = new AnalysisService(_editor);
            _analysis.UseSeverityOverrides(_settings.AnalysisRuleSeverityOverrides);

            // The Tests pane rediscovers whenever a pass finds the text moved, from the very
            // snapshot the pass just read - no read of its own, and silent unless the
            // discovered shape changed.
            _analysis.SnapshotObserved = OnAnalysisSnapshot;

            // And the tree - with the two panes' file lists behind it - follows files opening
            // and closing in the host, which nothing in here would otherwise notice.
            _analysis.ProjectsObserved = OnProjectsObserved;

            // The engine dying is not a crash and is not recoverable in this session: it is a
            // separate process and it is not restarted. Everything else keeps working, which is
            // precisely why it has to be said - the Problems panel reads "0 Errors" either way.
            _analysis.EngineStopped = () => _editorSurface?.Notify(
                "Analysis has stopped. The problems listed are from the last pass and will not "
                + "change. Close and reopen the editor to start it again.");

            // AND THE WAIT BEFORE IT EVER STARTS, which is the one the developer actually meets.
            //
            // Measured on a real ALT+F11 (lesson 64): the surface is up and usable at 574ms and
            // the engine is not connected until 3,370ms. For those 2.8 seconds the editor looks
            // finished and has no diagnostics, no completions and no hover - so a developer types
            // into what appears to be a working editor and gets nothing back, which reads as
            // broken rather than as loading. Nothing can make a 90 MB image load faster here; what
            // it can do is stop pretending.
            //
            // Held rather than timed, because the length of the wait is not known in advance, and
            // taken away by the engine itself rather than by a clock.
            _analysis.EngineReady = () => _editorSurface?.RunOnHostThread(() =>
                _editorSurface?.Hold(string.Empty));

            // The read of the projects belongs to the thread that owns them. The door is the
            // overlay's action timer, and it answers false while there is no surface to carry
            // it - the service retries rather than reading from the wrong thread.
            _analysis.HostMarshal = action =>
            {
                var surface = _editorSurface;
                if (surface is null)
                {
                    return false;
                }

                surface.RunOnHostThread(action);
                return true;
            };

            _analysis.LanguageFactsReady += (types, procedures) =>
            {
                // Unchanged words are not sent: the page rebuilds its tokenizer on arrival,
                // which re-tokenizes the whole module, and the lists rarely change.
                var factsKey = string.Join('\n', types) + "\0" + string.Join('\n', procedures);
                if (factsKey == _lastLanguageFactsKey)
                {
                    return;
                }

                _lastLanguageFactsKey = factsKey;
                Log.Info($"analysis: language facts, {types.Count} type(s), {procedures.Count} procedure(s)");
                _editorSurface?.RunOnHostThread(() =>
                    _editorSurface?.SetLanguageFacts([.. types], [.. procedures]));
            };

            _analysis.FindingsReady += findings =>
            {
                _findings = findings;
                Log.Info($"analysis: {findings.Count} finding(s)");

                // The log keeps a bounded record for support. A project with thousands of findings
                // would otherwise write a novel on every pass.
                foreach (var finding in findings.Take(20))
                {
                    Log.Info($"  {finding.Module}({finding.StartLine},{finding.StartColumn}) " +
                             $"{finding.Severity} {finding.Code}: {finding.Message}");
                }

                if (findings.Count > 20)
                {
                    Log.Info($"  and {findings.Count - 20} more");
                }

                // This callback arrives on the engine's reader thread, and the browser refuses
                // any other thread than its own. Without the hop, every mid-typing refresh dies
                // with UI_E_WRONG_THREAD and the panel goes stale until a module switch.
                _editorSurface?.RunOnHostThread(() =>
                {
                    PublishMarkersToSurface();
                    PublishFindingsToSurface();
                });
            };

            _analysis.Start();
        }
        catch (Exception ex)
        {
            Log.Error("analysis: could not be started", ex);
        }
    }

    /// <summary>
    /// Keeps the editing surface over whichever pane is being edited.
    ///
    /// Created on first use rather than at start-up, because until a pane exists there is nothing to
    /// cover and no rectangle to use. When no pane is visible the surface is hidden rather than
    /// destroyed: rebuilding a browser costs far more than leaving one parked off screen.
    /// </summary>
    private void FollowActivePane(IReadOnlyList<CodePane> panes)
    {
        try
        {
            var pane = panes.FirstOrDefault(p => p.IsVisible && !IsScratchComponent(p.Component));

            Log.Verbose($"follow: {panes.Count} pane(s), covering '{pane.Component ?? "(none)"}'");

            if (pane.Window == 0)
            {
                // A visible pane that was filtered out is the evaluator's scratch module mid
                // evaluation. Everything stays as it is; the pane is gone a moment later.
                if (panes.Any(p => p.IsVisible))
                {
                    Log.Verbose("follow: only the evaluator's scratch pane is visible, staying put");
                    return;
                }

                OnNoVisiblePane();
                return;
            }

            // The surface is a peer of the document area, not of the documents inside it.
            //
            // Put among the panes, it was a sibling of them, and the editor raises a pane whenever
            // it activates one. That happens before anything can react, so switching module showed
            // the pane being activated, scrollbars and all, until the surface was raised again. It
            // is a race that cannot be won from the outside: the editor is always first.
            //
            // A child of the frame is not in that fight at all. Activating a pane reorders the
            // document area's children and leaves the frame's children alone, so nothing ever comes
            // between the surface and the panes it covers. It is positioned on the document area's
            // rectangle, so it still covers exactly that and nothing else.
            var documentArea = Win32.GetParent(pane.Window);
            var host = Win32.GetAncestor(pane.Window, Win32.GaRoot);
            if (documentArea == 0 || host == 0)
            {
                return;
            }

            // Remembered so that placement can be recomputed at moments that are not window
            // events: a menu item opening a native window, or the page announcing it is ready.
            _frame = host;
            _documentArea = documentArea;

            if (!EnsureSurfaceForFrame(host))
            {
                return;
            }

            // The surface covers the whole document area, not the rectangle of one pane. Switching
            // module is then a message to a surface that never moved and was never uncovered.
            //
            // The native panes keep running underneath, unchanged and never seen. They remain the
            // text of record, the compile target, and what the debugger drives.
            _surfaceShown = true;
            _watchingEmpty = false;
            UpdatePolling();

            // A drag is the same answer per mouse move with only rectangles changed, and the
            // frame routes already follow rectangles synchronously. Everything below - the
            // full placement with policing and bands, the module and project publishes, debug
            // state, resync - is for events that changed WHAT is on screen, not where its
            // edges sit this instant; running it per move tick was the drag latency of
            // 2026-08-05. A geometry-only event takes the fast path and arms the settle,
            // whose full pass re-derives placement once the events pause. Debug freshness
            // rides the poll bursts the step commands arm, as it already does.
            var substance = string.Join("|",
                    panes.Select(p => p.Component + (p.IsVisible ? string.Empty : "~")))
                + "\0" + pane.Window + "\0" + pane.Component + "\0" + (pane.Project ?? string.Empty);
            if (substance == _lastFollowSubstance)
            {
                PlaceSurfaceFast();
                return;
            }

            _lastFollowSubstance = substance;

            // One placement path for every trigger, so policing, chrome, and bounds cannot
            // drift apart between the window-event route and the recomputed one.
            RefreshSurfacePlacement();

            // The compare carries the workbook when the tracker could name one: two workbooks'
            // Module1 are two documents, and a name-only compare would never switch between
            // them. A tracker that cannot say (its project goes null when two open panes share
            // a caption) keeps the name-only behaviour.
            var followedDisplay = DisplayFromProjectId(pane.Project);
            if (pane.Component is not null
                && (pane.Component != _editorSurface.Module
                    || (followedDisplay is not null && _editorSurface.Project is not null
                        && !string.Equals(followedDisplay, _editorSurface.Project, StringComparison.OrdinalIgnoreCase))))
            {
                // Before the active document changes: activating a pane is the moment the host
                // may read the module - F5 is one keystroke away - so the text is made true.
                _editorSurface.FlushEdits();

                // The object model's active pane resolves the owner a caption could not.
                ShowModuleInSurface(pane.Component, pane.Project ?? ActivePaneOwner(pane.Component));
            }

            PublishModules();
            PublishProjects();

            // The editor moves and activates panes as it steps, so this is also a signal that
            // execution may have moved on, and that the module may have been changed by something
            // other than the developer.
            UpdateDebugState();
            ResyncFromModule();
        }
        catch (Exception ex)
        {
            Log.Error("editor surface: could not follow the active pane", ex);
        }
    }

    /// <summary>
    /// Handles a key the editor owns, pressed while the surface has focus.
    ///
    /// The surface covers the pane the editor would have received these through, so without this
    /// they stop working: F5 no longer runs anything, and the browser underneath treats it as a
    /// request to reload the page, which throws away the document the developer is editing.
    ///
    /// A recognised key is always claimed, whether or not the command it names could run. Passing
    /// an unavailable F5 on to the document would reload it, which is a worse answer than nothing
    /// happening.
    /// </summary>
    private bool OnSurfaceKey(uint virtualKey)
    {
        var shift = (Win32.GetKeyState(Win32.VkShift) & Win32.KeyDownMask) != 0;
        var control = (Win32.GetKeyState(Win32.VkControl) & Win32.KeyDownMask) != 0;

        // Keys the surface owns are claimed here, before the host is asked about them. F1 opens the
        // host's help, and a key that reaches the host is gone: the browser's hook is the only place
        // it can be taken, and taking it means the command has to be asked for rather than left to
        // the document's own key handling.
        if (VbeCommands.SurfaceCommandForKey(virtualKey, shift, control) is { } surfaceCommand)
        {
            Log.Info($"key: 0x{virtualKey:X2} -> surface {surfaceCommand}");
            _editorSurface?.RunEditorCommand(surfaceCommand);
            return true;
        }

        // The closing keys are handled above, as SURFACE commands, because the page is what has
        // tabs. They are still claimed here whatever happens to them, because unclaimed the
        // browser treats Ctrl+W as an instruction to close ITS window, which takes the whole
        // surface with it.
        //
        // They used to be answered here instead, by closing whatever module the host believed was
        // shown. With two workbooks open that belief drifts: the host held a null module and a
        // project belonging to the OTHER workbook while the page had TwinCaller active, so the key
        // was claimed and nothing closed at all (the developer, 2026-08-07).

        var command = VbeCommands.ForKey(virtualKey, shift, control);
        Log.Info($"key: 0x{virtualKey:X2}{(shift ? " shift" : string.Empty)}{(control ? " ctrl" : string.Empty)}"
                 + $" -> {(command == 0 ? "not ours" : command.ToString(System.Globalization.CultureInfo.InvariantCulture))}");

        if (command == 0)
        {
            return false;
        }

        ExecuteEditorCommand(command);
        return true;
    }

    /// <summary>
    /// Decides what no visible pane means, because it means two different things.
    ///
    /// During a switch there is a moment with no pane window at all: the editor keeps a window
    /// only for the active pane, and the old one is gone before the new one exists. The open list
    /// still counts both, so nothing is done and the next event resolves it. When the open list
    /// itself is empty, the developer closed every tab, and the surface stays exactly where it is
    /// showing its empty workspace: hiding it would hand the whole frame back to the native
    /// editor, and a developer with no tabs open still owns an editor.
    /// </summary>
    private void OnNoVisiblePane()
    {
        // The next pane to appear is a change of substance whatever its identity says.
        _lastFollowSubstance = null;

        if (_editorSurface is null || _frame == 0 || _documentArea == 0)
        {
            // No surface yet and no pane to hang one on: the empty-workspace path can still
            // stand one up over the bare frame.
            TryShowEmptyWorkspace();

            if (_editorSurface is null)
            {
                _surfaceShown = false;
            }

            return;
        }

        var open = ReadOpenModules();
        if (open is not { Count: 0 })
        {
            // A window gap mid-churn, or a pane list that is refusing: either way, not an empty
            // editor. Everything stays as it is.
            return;
        }

        Log.Info("editor surface: no panes open, showing the empty workspace");

        _surfaceShown = true;
        _watchingEmpty = true;
        _editorSurface.Clear();
        _editorSurface.ShowModules([], [], null, null);
        RefreshSurfacePlacement();
        UpdatePolling();
    }

    /// <summary>
    /// Runs one of the editor's commands, whichever way the developer asked for it.
    ///
    /// Every route goes through here: the key, the toolbar button, and the glyph margin. Having
    /// two of these is exactly how the toolbar's toggle came to set a breakpoint that was never
    /// drawn: the bookkeeping was on the key path and the button went straight at the command.
    /// </summary>
    private VbeCommands.CommandRun ExecuteEditorCommand(int command, bool skipDesignerApply = false)
    {
        if (command == 0)
        {
            return VbeCommands.CommandRun.No("no such command");
        }

        // The Object Browser is ours (developer, 2026-08-05): a floating themed window of our
        // own over the typelib catalog, outside the canvas. The native window - which cannot
        // float, cannot be adopted, and paints only docked (lesson 32) - is never opened at all.
        if (command == VbeCommands.Command.ObjectBrowser)
        {
            OpenBrowserPalette();
            return VbeCommands.CommandRun.Ok("the palette was opened");
        }

        // The editor runs what the module holds, and acts on its own caret. Both are brought up to
        // date here, at the one moment it matters: running code the developer has not finished
        // typing is worse than a short pause before it starts. A toolbar button also takes focus
        // off the surface, which is when the two carets are furthest apart.
        _editorSurface?.FlushEdits();
        SyncCaretToPane();

        // Run with a designer tab holding the active slot runs THE FORM - the editor's own F5
        // with a designer window selected. The flush above still matters: the form's
        // code-behind may hold unwritten page edits, and running stale code is the same bug
        // running a stale module was.
        //
        // AND THE DOCUMENT IS FLUSHED THE SAME WAY, through the tab's own apply. The designer's
        // document is the transaction log and the form only catches up on an apply, so a run
        // that skipped it launched the last one: undo a control move, press F5, and the form on
        // screen still holds the move while the canvas beside it does not (the owner,
        // 2026-08-16, with a screenshot of a Hold button in two places). The run rides the
        // apply's callback, so a refused apply launches nothing and says why at the document.
        // The FILE is not written on the way: Run never saves in the native editor, and the
        // save that once rode here is the 2026-08-27 story at the "runOnly" callback.
        if (command == VbeCommands.Command.Run && _activeDesignerTab is { } designTab)
        {
            // A second launch aimed while the project is ALREADY running writes a design over a
            // live instance: the apply lands mid-run, the new form never stands, and every
            // designer open after it answers "has no designer" until the project resets (the
            // owner, 2026-08-19: "trying to launch a form, when a form is already active").
            // What the developer wants is the form that is standing, so it comes forward
            // instead. BREAK mode falls through past this fork on purpose: F5 stopped at a
            // breakpoint is CONTINUE, which is exactly the plain command below.
            var mode = ProjectModeNow();
            if (mode == RunMode)
            {
                var standing = FindStandingFormWindow();
                if (standing != 0)
                {
                    Win32.SetForegroundWindow(standing);
                    Log.Info("run form: a form is already standing; brought forward instead of launched again");
                    return VbeCommands.CommandRun.Ok(
                        "a form is already running - its window was brought forward; close it before launching again");
                }

                return VbeCommands.CommandRun.No(
                    "the project is already running - close the running form, or Reset, before launching another");
            }

            if (mode == DesignMode)
            {
                if (!skipDesignerApply)
                {
                    _editorSurface?.RequestDesignerApplySave(
                        designTab.Module, DisplayFromProjectId(designTab.ProjectId), run: true);
                    return VbeCommands.CommandRun.Ok("the designer tab applies, then runs");
                }

                return RunFormFromDesigner(designTab.Module, designTab.ProjectId);
            }
        }

        // Save with a designer tab holding the active slot is the designer's FlushEdits:
        // the tab's document lives in the PAGE, so the page applies it to the form and
        // calls back for the raw save ("saveOnly"). Ctrl+S is a HOST accelerator - the
        // page never sees the key at all, which is why a page-side binding alone read as
        // "CTRL+S still not working" from both halves (the owner, 2026-08-15).
        if (command == VbeCommands.Command.Save && !skipDesignerApply && _activeDesignerTab is { } saveTab)
        {
            _editorSurface?.RequestDesignerApplySave(saveTab.Module, DisplayFromProjectId(saveTab.ProjectId));
            return VbeCommands.CommandRun.Ok("the designer tab applies, then saves");
        }

        // Toggling a breakpoint is bookkeeping as well as a command. The editor cannot report which
        // lines carry one, so the record kept here is the only thing the surface can draw from, and
        // a route that skips it sets a breakpoint that is real and invisible.
        if (command == VbeCommands.Command.ToggleBreakpoint)
        {
            ToggleBreakpoint(_editorSurface?.CaretLine ?? 0);
            return VbeCommands.CommandRun.Ok("the breakpoint was toggled");
        }

        /*
         * RUN RUNS THE PROCEDURE THE CURSOR IS IN, and it used to say "executed" when the cursor
         * was not in one.
         *
         * Measured with a witness only running can write - a module-level String the procedure
         * assigns - because "ran: true" is the claim under test and cannot also be the evidence:
         *
         *   caret inside the Sub          ran=true executed   witness "" -> "YES"
         *   caret on a declaration        ran=true executed   witness "" -> ""
         *   caret on Option Explicit      ran=true executed   witness "" -> ""
         *
         * Same answer, opposite outcomes, and no dialog to hint otherwise. That is the same false
         * success as the toolbar press that reported Reset while the command was greyed, and as
         * the Design Mode button that reported a toggle Excel was holding down.
         *
         * DESIGN MODE ONLY. In break, Run is CONTINUE - it resumes the stopped procedure and has
         * nothing to do with where the caret is sitting.
         */
        if (command == VbeCommands.Command.Run && ProjectModeNow() == DesignMode
            && ProcedureAtCaret() is null)
        {
            Log.Info("run: the caret is not inside a procedure, so there is nothing to run");
            _editorSurface?.Notify("Run runs the procedure the cursor is in. "
                + "Put the cursor inside one, or pick a test or a form to run.");
            return VbeCommands.CommandRun.No(
                "the cursor is not inside a procedure, so Run has nothing to run - "
                + "put it inside the Sub or Function you mean");
        }

        // PRESSING DESIGN MODE THROWS THE DEVELOPER'S RUN AWAY. Read before, said after.
        var stoppedBeforeThis = command == VbeCommands.Command.DesignMode
            && ProjectModeNow() == BreakMode;

        // And the toggle as it stands, so the poll can tell a press that MOVED it from one that
        // did nothing. Excel holds design mode down while a workbook's macros are disabled, and
        // the button reports success either way.
        if (command == VbeCommands.Command.DesignMode)
        {
            _designModeBefore = VbeCommands.HostIsInDesignMode(_editor);
        }

        var outcome = VbeCommands.Execute(_editor, command);
        var ran = outcome.Ran;

        /*
         * EXCEL'S DESIGN MODE IS A STATE, AND IT USED TO BE AN INVISIBLE ONE.
         *
         * The toggle is the host's, not the editor's: while it is on, nothing runs at all, and
         * Reset, Break and Step Out are greyed on every bar that carries them - measured 0 of 6
         * Reset controls enabled, with the project reading design mode, no dialog standing, no
         * form showing and Excel's own surface still interactive. Pressing it while stopped also
         * discards the break: `published break, live 1, marker 6` becomes `published design,
         * live 2, marker none` on the next tick, silently.
         *
         * That is a developer's run gone and every debug command refusing, with a title bar
         * reading exactly as it would if the run had simply finished, and the way out is the same
         * button that greyed everything else. This says so, both ways.
         */
        //
        // ANSWERED FROM THE NEXT POLL, not from here. The toggle's own state does not change
        // under `Execute` - Office refreshes command bar state on its own schedule - so reading
        // it on the line after the press returns what it was BEFORE, and the first version of
        // this said "the host is now out of design mode" at the moment it went in. Caught in the
        // log by pressing it while stopped. The poll owns the host thread on an ordinary frame
        // and reads it a tick later, by which time it is true.
        if (command == VbeCommands.Command.DesignMode && ran)
        {
            _designModePressedWhileStopped = stoppedBeforeThis;
            _designModePresses = 3;

            // A press we ALREADY know does nothing is not reported as having done something. The
            // first one cannot know - the toggle is read back a tick later - but the second
            // onwards can, and a button that answers "executed" while nothing moves is the
            // false success this door has been cleaning out all week.
            if (_designModeStuck)
            {
                outcome = VbeCommands.CommandRun.No(
                    "Excel will not leave design mode, which usually means this workbook's macros "
                    + "are disabled; close it and reopen it with macros enabled");
                ran = false;
            }
        }

        // AND A COMMAND GREYED BY IT SAYS WHICH STATE GREYED IT. A refusal that names only the
        // control is a dead end; this one names the way out. Asked only when a command was
        // actually turned away as greyed, so the extra read costs nothing on the usual path.
        if (!ran
            && outcome.Detail.StartsWith("currently disabled", StringComparison.Ordinal)
            && command != VbeCommands.Command.DesignMode
            && VbeCommands.HostIsInDesignMode(_editor))
        {
            // TWO DESIGN MODES, AND THE ADVICE IS DIFFERENT. One the developer pressed, which the
            // toggle undoes; one EXCEL is holding down because the workbook's macros are disabled,
            // which it will not. Telling the second to press Design Mode is advice that cannot
            // work - measured: the toggle stays -1 and Reset stays 0 of 6 however many times it
            // is pressed, while the button answers "executed" every time.
            Log.Info($"command: {command} is greyed because Excel is in design mode"
                + (_designModeStuck ? ", which it will not leave" : string.Empty));
            outcome = VbeCommands.CommandRun.No($"{outcome.Detail} - Excel is in DESIGN MODE, so "
                + (_designModeStuck
                    ? "nothing will run, and it will not leave it - which usually means this "
                        + "workbook's macros are disabled; close it and reopen with macros enabled"
                    : "nothing will run; press Design Mode to leave it"));
            _editorSurface?.Notify(_designModeStuck
                ? "Excel is in design mode and will not leave it, so that command is greyed. That "
                    + "usually means this workbook's macros are disabled - close it and reopen it "
                    + "with macros enabled."
                : "Excel is in design mode, so that command is greyed. Press Design Mode to leave it.");
        }

        // A RESET THE EDITOR GREYS WHILE SOMETHING IS STOPPED HAS ONE MORE THING TO TRY.
        //
        // Issue #9: a break in which every debug command answered "currently disabled", nothing
        // the product offered could leave it, and the only remedy was restarting Excel - which
        // costs the developer every unsaved workbook in the host, not just their debugging
        // session. Run, Break, Reset, Design Mode, Step Out and Compile are all aimed at the
        // ACTIVE project, and this surface covers the editor's own windows, so the editor can be
        // left attending to a project that is not the stopped one.
        //
        // What a person does then is click the stopped project in the Project Explorer. So does
        // this: it brings a code pane of the stopped project forward and asks again, once. If the
        // refusal was real the second answer is the same one and nothing has been lost.
        if (!ran && command == VbeCommands.Command.Reset && AttendToStoppedProject() is { } woken)
        {
            outcome = VbeCommands.Execute(_editor, command);
            ran = outcome.Ran;
            Log.Info($"reset: refused, so {woken} was brought forward and it was asked again - "
                + (ran ? "it ran" : $"still refused ({outcome.Detail})"));
            outcome = ran
                ? VbeCommands.CommandRun.Ok($"executed after bringing {woken} forward")
                : VbeCommands.CommandRun.No(
                    $"{outcome.Detail}; {woken} is stopped and was brought forward first");
        }

        // A refused Call Stack looked broken rather than declined ("won't appear again",
        // 2026-08-05: the break had ended and the native command was disabled). The button
        // greys with the debug mode now, but a race between click and mode change still
        // deserves words instead of silence.
        if (!ran && command == VbeCommands.Command.CallStack)
        {
            _editorSurface?.Notify("Call Stack shows the stopped procedure chain; it needs break mode.");
        }

        // The Watch panel's own buttons are the only route to these now (the Debug menu's
        // items are retired), so a declined one has to say why rather than do nothing. The
        // editor disables Edit Watch until a watch is selected in the window the ghost
        // palette IS, and Quick Watch until an expression sits under the caret.
        if (!ran && command == VbeCommands.Command.EditWatch)
        {
            _editorSurface?.Notify("Edit Watch works on a selected watch; add one first.");
        }

        if (!ran && command == VbeCommands.Command.QuickWatch)
        {
            _editorSurface?.Notify("Quick Watch reads the expression at the cursor; put the caret on one first.");
        }

        // Clearing all breakpoints clears the editor's; the drawn record must follow, whichever
        // route asked for it.
        if (command == VbeCommands.Command.ClearAllBreakpoints)
        {
            ForgetBreakpoints();
        }

        // A save just cleaned the workbook - but the flag flips a beat AFTER the command
        // returns, so one immediate republish read the old value and the dot lingered. The
        // next few polls re-derive it; the change-key keeps the repeats free.
        if (command == VbeCommands.Command.Save)
        {
            PublishModules();
            _resyncPanePolls = Math.Max(_resyncPanePolls, 3);
        }

        WatchDebugState();

        // A command can change the native window landscape - the Object Browser above all -
        // and no event the tracker recognises announces it. The menu route always re-derived
        // placement after executing; this route learned the same manners (2026-08-05, the
        // Browser opening invisible under the surface).
        RefreshSurfacePlacement();

        return outcome;
    }

    /// <summary>
    /// A designer window left standing for a form the Run command was aimed at. Measured
    /// 2026-08-14, one layer per trace: the command POSTS the editor's own action, so a
    /// synchronous put-down un-aims it and the run degrades to the Macros dialog; and a
    /// put-down on the tick that sees the mode LEAVE design lands between run-start and the
    /// form window appearing, and the launching form dies with its designer. Natively the
    /// designer stands visible behind the running form, so that is what this does: down on
    /// the tick that sees the run OVER (design again after not-design), or on the deadline
    /// when the run never took hold.
    /// </summary>
    private (string Module, string ProjectId, long Deadline, bool SawRun)? _designerRunStanding;

    /// <summary>
    /// Runs a form the way the editor's own Run does with its designer selected. The Run
    /// command aims at the editor's active window, so the form's native designer window is
    /// made visible and focused - this product otherwise keeps designer windows down - and
    /// goes back down through <see cref="_designerRunStanding"/> once the run has taken
    /// hold, because the command's action is posted and putting the window down here would
    /// un-aim it before the action reads its target.
    /// </summary>
    private VbeCommands.CommandRun RunFormFromDesigner(string moduleName, string projectId)
    {
        try
        {
            using var component = FindComponent(moduleName, projectId, out _);
            if (component is null)
            {
                return VbeCommands.CommandRun.No($"no component named {moduleName}");
            }

            using var window = component.CallObject("DesignerWindow");
            if (window is null)
            {
                return VbeCommands.CommandRun.No($"{moduleName}'s designer window would not open");
            }

            // ARMED BEFORE THE WINDOW IS SHOWN. Showing it raises a window event, the event runs
            // the designer checks, and those put every designer tab back down - so arming after
            // the command left a gap wide enough for the aimed window to be hidden inside it,
            // which is what happened (2026-08-17). The deadline starts here for the same reason:
            // it measures the run, and the run is about to start.
            _designerRunStanding = (moduleName, projectId, Environment.TickCount64 + 3000, false);
            FormDesignService.StandingForRun = moduleName;

            window.SetBool("Visible", true);
            window.Invoke("SetFocus");

            Log.Info($"run form: {moduleName} through the editor's own Run");
            var outcome = VbeCommands.Execute(_editor, VbeCommands.Command.Run);

            WatchDebugState();

            return outcome.Ran ? VbeCommands.CommandRun.Ok($"ran {moduleName}") : outcome;
        }
        catch (Exception ex)
        {
            // The exception is over if the run never started, and leaving it armed would keep a
            // designer window standing for the three seconds until the deadline collected it.
            _designerRunStanding = null;
            FormDesignService.StandingForRun = null;
            Log.Error($"run form: {moduleName} would not run", ex);
            return VbeCommands.CommandRun.No($"{moduleName}'s designer would not take focus");
        }
    }

    /// <summary>
    /// Writes what the developer typed back into the module.
    ///
    /// The module is the text of record. Everything else in the host reads it and nothing reads the
    /// surface: the compiler, the debugger, the file the workbook saves, and the analyzer all go to
    /// the module, so an edit that has not reached it has not happened. Before this existed, typing
    /// in the surface changed nothing at all: the code would not run, would not save, and the
    /// analyzer went on reporting defects in text the developer had already fixed.
    ///
    /// The whole module is replaced rather than the changed range applied. The host's own line
    /// operations are one call per line and its line numbers shift under each other as they are
    /// applied, so replacing once is faster and shifts nothing under anything.
    ///
    /// A REPLACE IS TWO CALLS, AND THE EDITOR CAN REFUSE THE SECOND. This used to claim its failure
    /// mode was a module unchanged. It is not, and the difference is a developer's code: the delete
    /// lands, the add is refused partway, and what is left is neither body. Measured 2026-08-09 - a
    /// module of 2,002 working lines was asked to take a body the editor would not have, and came
    /// back holding 31,956 lines of the new one. The route replied ok. So the old text is kept in
    /// hand and put back if the write is refused, on both paths, and the complaint is RETURNED
    /// rather than only logged.
    ///
    /// Null means it was written. Anything else is what went wrong, in words, for the caller to
    /// show and for the door to answer with.
    ///
    /// Writing resets the project, which discards any running state. That is what the host's own
    /// editor does when a module is edited, so it is parity rather than a regression, and it is
    /// why this is debounced rather than done per keystroke.
    /// </summary>
    /// <param name="keepEveryCharacter">
    /// Refuse the write, and put back what the module held, if the host converts a character on
    /// the way in rather than storing it.
    ///
    /// TRUE FOR IMPORT ONLY, and the asymmetry is the whole point. The host stores module text in
    /// the system ANSI code page, so a character outside it is converted by Excel before this
    /// product sees it. Where the developer's good copy lives decides what to do about that:
    ///
    ///   IMPORT: the good copy is the FILE, and it is safe on disk. Refusing costs nothing they
    ///   still have. Accepting costs them the file, because the module now differs from it, so the
    ///   next export is ticked "will write" and overwrites their source with the converted text.
    ///   Measured 2026-08-09: a repository file carrying Cyrillic came back byte for byte
    ///   destroyed after one import and one export, with "1 changed, 0 failed" reported at both.
    ///
    ///   TYPING: the good copy is on the SURFACE and nowhere else. Refusing leaves the page and
    ///   the module disagreeing for as long as the tab stays open, which is the state the tabs
    ///   setting was removed for. So typing is told once and not refused; it cannot destroy a file
    ///   either way, because the module never held the good version for an export to write.
    /// </param>
    private string? WriteModule(
        string component,
        string text,
        string? ownerProject = null,
        bool hostRewrite = false,
        bool keepEveryCharacter = false)
    {
        try
        {
            // A STOPPED PROJECT CANNOT BE EDITED, and asking anyway is not a harmless failure.
            //
            // The editor answers an edit made in break mode by asking "This action will reset
            // your project, proceed anyway?" - a real question, which anything answering dialogs
            // on our behalf declines, so the write comes back as a bare COM error naming nothing
            // that happened. And it keeps being asked: a random walk that wrote while stopped
            // raised and cancelled that dialog over and over, and the editor eventually faulted
            // in VBE7.DLL with 0xc0000005 and took Excel with it (issue #6, chaos.mjs).
            //
            // Refusing is also the only answer here that cannot lose anything. Proceeding means
            // agreeing to reset the developer's debugging session - their call stack, their
            // locals, the run they stopped on purpose - and that is theirs to decide.
            //
            // Read fresh rather than taken from _inBreak, which is as old as the last poll. The
            // evaluator learned the same lesson: acting on the cached flag put work into a
            // stopped project "in ways that have nothing to do with what the developer typed".
            if (ProjectModeNow() != DesignMode)
            {
                Log.Warn($"write: {component} not written - the project is stopped in the debugger");

                // NAMES THE WAY OUT, and names it for both kinds of caller. A developer presses
                // Reset; a caller driving the door has no hands, and a message that only says
                // "press Reset" leaves it wedged with no move to make. The route exists, so it
                // is said here rather than left to be discovered.
                return $"{component} was not written: the project is stopped in the debugger. "
                    + "Editing now would reset it and lose the run, so it was not attempted. "
                    + "Press Reset in the editor, or POST command?name=reset, and write again.";
            }

            // A write is normally about the module on the surface, so it goes to the shown
            // project's component - never to a same-named module in another workbook that
            // happened to enumerate first. A caller who knows better says so: the close
            // path reverts background tabs, whose owner is the tab's workbook, not the
            // shown one.
            var owner = ownerProject
                ?? (string.Equals(component, _editorSurface?.Module, StringComparison.OrdinalIgnoreCase)
                    ? _shownProject
                    : null);

            using var found = FindComponent(component, owner, out var foundOwner);
            using var module = found?.GetObject("CodeModule");
            if (found is null || module is null)
            {
                Log.Warn($"write: {component} has no code module");
                return $"{component} has no code module to write to";
            }

            // A LINE THE EDITOR CANNOT HOLD IS NOT WRITTEN AT ALL.
            //
            // It does not refuse one. It takes it and BREAKS it, at 1,023 characters, with no
            // continuation character: the statement is cut in half and a string literal is left
            // unterminated, so what lands is not code. Measured 2026-08-09 - 1,022 characters come
            // back identical, 1,023 come back as two lines, and a 2,018-character Debug.Print
            // became a 1,023-character fragment followed by a 995-character one.
            //
            // Refused here rather than repaired. A continuation would have to be inserted INSIDE
            // the developer's expression, and inside a string literal that means splitting it and
            // concatenating - a rewrite of their code to make it fit, decided by us. Better to say
            // which line and let them break it where it belongs.
            if (Core.Editor.ModuleText.FirstLineTooLong(text) is { } tooLong)
            {
                Log.Warn($"write: {component} line {tooLong.At} is {tooLong.Length} characters; nothing written");
                return $"{component} was not written: line {tooLong.At} is {tooLong.Length} characters. "
                    + $"The editor holds {Core.Editor.ModuleText.LongestLine} in a line and breaks "
                    + "anything longer in half without a continuation, which would not be valid "
                    + "code. Break the line yourself and it will write.";
            }

            // What the module holds before any of this, when a caller has said the write is all or
            // nothing. One read, on the import path only, which is already the slow one.
            var wasHoldingBefore = keepEveryCharacter ? ProjectReader.ReadSource(found) : null;

            // The baseline belongs to the workbook actually found: a line diff computed against
            // another workbook's same-named module would write a merge of the two.
            var writtenKey = WrittenKey(component, DisplayFromProjectId(foundOwner ?? owner));

            // The changed lines alone, when we know where they are. Replacing a whole large
            // module makes the host reparse every line of it - seconds, on the thread the
            // keystrokes live on - where typing only ever touches a few. The whole replace below
            // remains for a rewrite too large to call an edit.
            //
            // NO BASELINE IS NOT A REASON TO REPLACE THE MODULE. It used to be: the baseline is
            // remembered from what this session has written or shown, so the FIRST write to a
            // module nobody has opened had none and replaced all of it. That is not a rare path -
            // it is a rename reaching modules that are not open, an import, a replace-all, any
            // api write - and on a 64,802-line module it takes Excel off the air long enough for
            // Windows to mark the window Not Responding (caught doing exactly that at 05:17 on
            // 2026-08-21, "took 0ms to compare and is replacing the module").
            //
            // So the module is asked what it holds instead: one read, about 40ms at that size,
            // against seconds of reparse. The import path has already read it for its
            // all-or-nothing check, and that copy is used rather than read again.
            var baseline = _writtenModules.TryGetValue(writtenKey, out var known)
                ? known
                : wasHoldingBefore ?? ProjectReader.ReadSource(found);

            // Declared out here because the call below is short-circuited: with no baseline there
            // is no diff to attempt and no window to hear about.
            LineWindow? window = null;

            // TIMED, because this is the host thread and the numbers were guesses until they were
            // not. The two halves answer different questions: `found` is ours - comparing the
            // texts and deciding what to write - and `took` is the editor's, taking it. A write
            // that is slow in the first is a bug here; one slow in the second is what the editor
            // charges for the module's size.
            var diffing = System.Diagnostics.Stopwatch.StartNew();
            var wroteDiff = baseline is not null && TryWriteLineDiff(module, baseline, text, out window);
            diffing.Stop();

            if (!wroteDiff)
            {
                // WHAT THE MODULE HOLDS NOW, BEFORE ANY OF IT IS DELETED.
                //
                // The copy costs a read of the same text the write reads back anyway, and it is the
                // only thing standing between a refused write and a developer's lost work. Measured
                // against the write it protects: 3ms of a 1,037ms write at 1,002 lines, 66ms of a
                // 12,594ms write at 40,002. Half a percent, at the size where losing it hurts most.
                //
                // Only on this path. The diff path already knows the lines it is about to remove
                // and puts them back itself, so it pays nothing.
                //
                // And the import path has already read it, at the top, for the all-or-nothing
                // check - the same component, with nothing written in between, so the second read
                // could only ever return the same text. Import is the slow path and this is the
                // largest single read in it.
                var wasHolding = wasHoldingBefore ?? ProjectReader.ReadSource(found);

                var existing = module.GetInt32("CountOfLines");
                if (existing > 0)
                {
                    module.Invoke("DeleteLines", 1, existing);
                }

                // A module with nothing in it is a legitimate state, and asking the host to add
                // an empty string to one is not.
                if (text.Length > 0)
                {
                    try
                    {
                        FillEmptyModule(module, text);
                    }
                    catch (Exception refusal)
                    {
                        var restored = PutModuleBack(module, wasHolding);
                        Log.Error($"write: {component} was refused{(restored ? " and its previous text was put back" : string.Empty)}", refusal);

                        // The words the editor used, which name the real limit. "Out of memory" is
                        // what it says about a module past its identifier budget, and that is far
                        // more useful to a developer than anything this could invent.
                        return restored
                            ? $"{component} was not written: {refusal.Message}. What it held before is back."
                            : $"{component} was not written: {refusal.Message}. Its previous text could NOT be restored.";
                    }
                }
            }

            // READ BACK ONLY WHAT WAS WRITTEN, when what was written can be checked on its own.
            //
            // The read below is the editor's own copy, and it is here because the editor rewrites
            // what it is given (see the comment under it). A whole-module read to check a
            // one-line edit costs the same as reading the module: on a 64,802-line module a write
            // of identical text - which touches the editor not at all - still cost about 100ms of
            // host thread, most of it this read and the comparison after it, on every pause in
            // typing (measured 2026-08-21).
            //
            // So: nothing written means nothing to check. A window written means the window is
            // read back and compared against exactly what went in, with the module's line count
            // as the guard - if the editor inserted or removed a line anywhere, or changed one
            // outside the window, the count moves and the whole module is read as before. Any
            // mismatch at all falls back to the full read, so the checks below see the same text
            // they always did.
            var checkedWindow = wroteDiff && (window is null || WindowCameBackIntact(module, window));

            // Read straight back and remembered, but not pushed into the surface.
            //
            // The editor rewrites what it is given, and its rewrites are the kind a developer is in
            // the middle of doing for themselves: it completes the parentheses on a procedure and
            // inserts a blank body for one. Sending that back mid-keystroke duplicated what had
            // just been typed and inserted lines nobody asked for. What it holds is remembered as
            // the baseline instead, so a later comparison sees changes made by something else and
            // not the editor's own tidying of our own write.
            var stored = checkedWindow ? text : ProjectReader.ReadSource(found);

            // A CHARACTER THE HOST CONVERTED RATHER THAN STORED.
            //
            // Nothing failed: the write was accepted and Excel quietly substituted, because module
            // text lives in the system ANSI code page. On import that is the developer's file about
            // to be lost, so the module goes back to what it held and the row is reported failed.
            // On every other path it is said once and allowed, for the reasons on the parameter.
            // Skipped when the read-back above already proved the module holds exactly this text:
            // a comparison of a string with itself can only answer "nothing was lost", and on a
            // 1.49MB module it walks both copies to say so.
            if (!ReferenceEquals(text, stored)
                && Core.Editor.ModuleText.FirstCharacterLost(text, stored ?? string.Empty) is { } lost)
            {
                var said = $"{component}: this machine's VBA cannot store {lost.Describe()}, on line "
                    + $"{lost.Line}. Excel converts it on the way in, so the module would not hold "
                    + "what was written.";

                if (keepEveryCharacter)
                {
                    var restored = PutModuleBack(module, wasHoldingBefore);
                    Log.Warn($"write: {said} Nothing was written"
                        + (restored ? " and the module is as it was." : "; the module could NOT be put back."));

                    return restored
                        ? $"{said} Nothing was written and the file on disk is untouched."
                        : $"{said} Nothing was written, and the module could NOT be put back.";
                }

                // Once per module, not once per pause in typing. The surface goes on holding the
                // character we cannot store, so every later write-back finds the same loss, and
                // the tenth copy of the sentence is in the way of the work. Kept apart from the
                // refusal notices because this one is not cleared by the write succeeding: the
                // write DID succeed, and saying it again next time would be the same repetition.
                if (!_saidAboutConversion.TryGetValue(writtenKey, out var already) || already != said)
                {
                    _saidAboutConversion[writtenKey] = said;
                    Log.Warn($"write: {said} It was written anyway; this path has nowhere else to keep it.");
                    _editorSurface?.Notify(said);
                }
            }

            // WHAT THIS WRITE DID, for the change log. The texts are both already in hand - the
            // baseline is what the diff was computed against and `stored` is what the module holds
            // now - so recording costs a hash and, the first time a round touches a module, one
            // copy. Nothing is read, nothing is written to the workbook, and a log that cannot
            // keep up never fails the write: see RecordChange.
            RecordChange(
                component, foundOwner ?? owner, Core.Changes.ChangeKind.Written,
                baseline, stored ?? text,
                componentKind: ComponentKind(found.GetInt32("Type")));

            _writtenModules[writtenKey] = stored ?? text;

            if (hostRewrite)
            {
                // The engine's live copy of a module OUTRANKS its seeded copy everywhere -
                // search, completion, diagnosis - which is what keeps answers exact
                // mid-keystroke, and the typing path maintains it by streaming edits ahead of
                // the write. A host rewrite bypasses the page, so the live copy is corrected
                // here, or the engine keeps diagnosing text that no longer exists - the
                // problems of a discarded edit survived the close and the reopen (2026-08-05).
                _analysis?.NotifyLiveText(component, stored ?? text, null, owner);

                // AND THE SURFACE, which this used to leave behind entirely.
                //
                // "A host rewrite bypasses the page" was the reason given for correcting the
                // engine here, and the page was then never corrected at all: a module written
                // from outside while its pane is open kept whatever text the surface last had.
                // Found 2026-08-08 with a freshly built fixture - the workbook held 42 lines,
                // the native pane held them too, and the editor showed an EMPTY document, so
                // every breakpoint on it was refused as "not an executable statement" because
                // the line being asked about did not exist on the surface.
                //
                // Sync only touches a document the surface actually holds, and leaves an
                // unwritten one alone, so this cannot overwrite an edit in flight.
                _editorSurface?.Sync(component, DisplayFromProjectId(foundOwner ?? owner), stored ?? text);

                // AND THE FOLDER, from the same text (#23). Typing moves a module on the typing
                // pause, from text that may not have reached the host yet; a rewrite that puts
                // the host's own text back then finds nothing to write, so no pass follows and
                // nothing else would read the annotation again - the tree kept the typed folder
                // over a document that no longer carried it (2026-09-05).
                NoteTypedFolder(foundOwner ?? owner, component, stored ?? text);
            }

            Log.Verbose($"write: {component} took {diffing.ElapsedMilliseconds}ms to compare"
                        + (window is { } touched
                            ? $" and wrote {touched.Count} line(s) at {touched.At} of {touched.TotalLines}"
                            : wroteDiff ? " and had nothing to write" : " and is replacing the module"));

            Log.Info($"write: {component}, {text.Length} character(s){(wroteDiff ? " as a line diff" : string.Empty)}"
                     + (stored is not null && stored != text ? " (the editor reformatted it)" : string.Empty)
                     + (hostRewrite ? " (host rewrite; the engine's live copy follows)" : string.Empty));

            // The write just made the workbook dirty; the tab dots follow. The page skips the
            // rebuild when nothing it draws has changed, so a write that changed no flag is free.
            PublishModules();

            // The full pass re-reads every module, reseeds the engine, and diagnoses the whole
            // project - work worth doing, but not per pause: the live pass keeps the shown
            // module honest between full passes, and a full pass running is what a completion
            // queues behind. It runs when the write was structural, or when its turn has come.
            // A host rewrite never waits its turn: the developer just watched the text change,
            // and the Problems pane must follow it now, not a quiet-period later.
            var now = Environment.TickCount64;
            if (hostRewrite || !wroteDiff || now - _lastFullAnalysis > FullAnalysisQuietMilliseconds)
            {
                _lastFullAnalysis = now;
                _fullAnalysisDeferred = false;
                _analysis?.Reanalyse();
            }
            else
            {
                _fullAnalysisDeferred = true;
            }

            return null;
        }
        catch (Exception ex)
        {
            Log.Error($"write: {component} could not be updated", ex);
            return $"{component} could not be written: {ex.Message}";
        }
    }

    /// <summary>
    /// Puts a text into an EMPTY module, and only that text.
    ///
    /// NOT AddFromString, WHICH CORRUPTS SOME MODULES. Hand it a module holding a `Declare`
    /// broken over a line continuation and the editor appends a line reading `()` to the end of
    /// the module - not a rewrite of the statement, a new line of nonsense at the bottom, which
    /// does not compile. Reproduced on 2026-08-21 outside this product entirely, in a throwaway
    /// workbook with no add-in loaded:
    ///
    ///   Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    ///       (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
    ///
    ///   AddFromString: 1,124 lines in, 1,125 back, every line identical plus a trailing `()`
    ///   InsertLines:   1,124 lines in, 1,125 back, every line identical plus a trailing blank
    ///
    /// Put the same declaration on one line and AddFromString is clean, which is why it took a
    /// real module to find: it is the continuation the editor mishandles. It cost the owner a
    /// module that would not compile, pasted from a file that was fine (VBA-JSON, whose Mac
    /// branch declares four of these).
    ///
    /// So the lines go in through InsertLines, which reproduces the text exactly, and the blank
    /// line it leaves after them is removed. Anything past the line count that was handed over
    /// was added by the editor, and none of it was asked for.
    /// </summary>
    private static void FillEmptyModule(DispatchObject module, string text)
    {
        module.Invoke("InsertLines", 1, text);

        var wanted = Core.Editor.LineDiff.CountLines(text);
        var present = module.GetInt32("CountOfLines");

        if (present > wanted)
        {
            module.Invoke("DeleteLines", wanted + 1, present - wanted);
        }
    }

    /// <summary>
    /// Puts a module's previous text back after a write the editor refused.
    ///
    /// True when the module is holding it again. False is the case that matters: the developer's
    /// text is gone and nothing here can bring it back, so the caller says so in those words rather
    /// than reporting a tidy failure. It happens - a module pushed past the editor's identifier
    /// budget leaves the whole VBE unable to take anything, including what it had a moment ago,
    /// and it stays that way until Excel is restarted (measured 2026-08-09).
    /// </summary>
    private static bool PutModuleBack(DispatchObject module, string? previous)
    {
        try
        {
            var present = module.GetInt32("CountOfLines");
            if (present > 0)
            {
                module.Invoke("DeleteLines", 1, present);
            }

            if (!string.IsNullOrEmpty(previous))
            {
                FillEmptyModule(module, previous);
            }

            return true;
        }
        catch (Exception ex)
        {
            Log.Error("write: the module's previous text could not be put back", ex);
            return false;
        }
    }

    /// <summary>
    /// Writes only the lines that changed between the baseline and the new text: the common
    /// prefix and suffix are found, and the window between them is deleted and re-inserted at
    /// one anchor, so nothing shifts under anything. False when the change is too large to be
    /// typing - a paste of a module's worth of text is a whole replace, honestly.
    ///
    /// The window it removes is kept until the replacement is in, so an insert the editor refuses
    /// gives the removed lines back instead of costing them. Free here, unlike on the whole-replace
    /// path: these lines are already in hand.
    /// </summary>
    /// <summary>
    /// What a diff write put into the module: where it started, how many lines it inserted, how
    /// many the module should hold afterwards, and the exact text of the window. Enough to check
    /// the write by reading back THAT WINDOW rather than the whole module - see the read-back in
    /// <see cref="WriteModule"/>. Null when nothing was written at all.
    /// </summary>
    private sealed record LineWindow(int At, int Count, int TotalLines, string Text);

    private static bool TryWriteLineDiff(DispatchObject module, string baseline, string text, out LineWindow? wrote)
    {
        const int LargestDiffLines = 400;

        wrote = null;

        var diff = Core.Editor.LineDiff.Between(baseline, text, LargestDiffLines);
        if (diff.Change == Core.Editor.LineChange.Wholesale)
        {
            return false;
        }

        if (diff.Change == Core.Editor.LineChange.Identical)
        {
            // The module already holds exactly this. Nothing is written, so there is nothing to
            // read back and nothing that could have been converted on the way in.
            return true;
        }

        // Never ask to delete lines that are not there.
        //
        // An EMPTY module has CountOfLines 0, but an empty text is one empty line - so the window
        // says "delete 1 from line 1" and the editor refuses the whole write with "Invalid
        // procedure call or argument". Nothing is written, and the only place it is said is the
        // log: the write route's reply looks like every other success. It took the code out of
        // the first module of every fixture built through the door, leaving a workbook that
        // looked right and no longer exercised what it existed for (2026-08-07).
        var present = module.GetInt32("CountOfLines");
        var removing = diff.Removing;
        if (diff.At - 1 + removing > present)
        {
            removing = Math.Max(0, present - (diff.At - 1));
        }

        if (removing > 0)
        {
            module.Invoke("DeleteLines", diff.At, removing);
        }

        // ON THE COUNT, not on whether the text is empty: one empty line is a line, and reading
        // it as nothing to insert drops the blank line at the end of a text that ends with an
        // ending - which is how a file on disk ends, so an import lost it.
        if (diff.Inserting > 0)
        {
            try
            {
                module.Invoke("InsertLines", diff.At, diff.Text);
            }
            catch
            {
                // The lines just removed, back where they were. They are at most
                // LargestDiffLines, and the window is already in hand, so this costs one call
                // and no read at all.
                if (removing > 0 && diff.Removed.Length > 0)
                {
                    try
                    {
                        module.Invoke("InsertLines", diff.At, diff.Removed);
                    }
                    catch (Exception putBack)
                    {
                        Log.Error("write: a removed window could not be put back", putBack);
                    }
                }

                throw;
            }
        }

        wrote = new LineWindow(diff.At, diff.Inserting, diff.TotalLines, diff.Text);
        return true;
    }

    /// <summary>
    /// Whether the module holds exactly the window that was just written to it, and nothing else
    /// moved: the line count is what the write expected, and the window reads back character for
    /// character. False for anything else - the editor tidied the text, converted a character it
    /// cannot store, or added a line somewhere - and the caller then reads the whole module,
    /// which is what it used to do unconditionally.
    /// </summary>
    private static bool WindowCameBackIntact(DispatchObject module, LineWindow wrote)
    {
        try
        {
            if (module.GetInt32("CountOfLines") != wrote.TotalLines)
            {
                return false;
            }

            if (wrote.Count == 0)
            {
                // A pure deletion: the count above is the whole check.
                return true;
            }

            return string.Equals(
                module.GetStringIndexed("Lines", wrote.At, wrote.Count),
                wrote.Text,
                StringComparison.Ordinal);
        }
        catch (Exception ex)
        {
            // A module that will not answer about its own lines is one to read whole.
            Log.Info($"write: a written window could not be read back ({ex.GetType().Name})");
            return false;
        }
    }

    /// <summary>
    /// Breakpoints the developer has set, by module.
    ///
    /// Kept here because the editor does not expose them. It has a command that toggles the one at
    /// its caret and no way at all to ask which lines carry one, so the only way to draw them is to
    /// remember every toggle that went through us. The surface is the only way to set one now that
    /// the native panes are covered, so this stays in step in practice; a breakpoint set some other
    /// way would be real and undrawn, which is why this is a record of what we did rather than a
    /// claim about what the editor holds.
    /// </summary>
    ///
    /// KEYED BY WORKBOOK AND MODULE, through WrittenKey, not by module name. Two open workbooks
    /// can each hold a `Helpers`, and keyed by name alone a breakpoint set in one was reported
    /// against the other: the dot drew on the wrong module and a run that should have stopped did
    /// not. Fifth defect of that shape in this codebase (2026-08-07), and the last of the ones
    /// that were known.
    private readonly Dictionary<string, BreakpointRecord> _breakpoints = new(StringComparer.Ordinal);

    /// <summary>
    /// One module's breakpoints, with the names AS SPELLED beside them.
    ///
    /// The key is lowercased, because two workbooks holding `Helpers` and `helpers` are holding
    /// the same module as far as VBA is concerned. Reporting the key back is what a first version
    /// did, and the `breakpoints` route then answered `helpers @ renamefixture.xlsm` - so a
    /// caller comparing against the name the product shows found nothing, which is a door that
    /// mangles its own answers. The spellings are kept so the route can hand back what a
    /// developer would recognise (2026-08-08).
    /// </summary>
    private sealed record BreakpointRecord(string Module, string? Project, SortedSet<int> Lines);

    /// <summary>The breakpoint key for a module of a workbook. Null project means the shown one.</summary>
    private string BreakpointKey(string module, string? projectDisplay = null) =>
        WrittenKey(module, projectDisplay ?? DisplayFromProjectId(_shownProject));

    /// <summary>The lines recorded for a module of a workbook, empty when it has none.</summary>
    private SortedSet<int> BreakpointsFor(string module, string? projectDisplay = null) =>
        _breakpoints.TryGetValue(BreakpointKey(module, projectDisplay), out var found) ? found.Lines : [];

    /// <summary>Whether execution was stopped last time it was looked at.</summary>
    private bool _inBreak;

    /// <summary>
    /// Whether an evaluation of ours raised a compile error, which drops the project out of design
    /// mode a moment after the evaluation has returned.
    ///
    /// Set on the way OUT, because that is where the compile error is seen, and acted on the way
    /// IN to the next evaluation, because that is where nothing is in flight to race.
    /// </summary>
    private volatile bool _immediateLeftItStopped;

    /// <summary>
    /// Whether the immediate window has already asked for a Reset during THIS stopped session.
    ///
    /// The pair below is how "did the reset work" gets answered without guessing at a duration.
    /// Asking is recorded here; arriving back a whole evaluation later, still stopped, is what
    /// proves it did not take.
    /// </summary>
    private bool _immediateResetAsked;

    /// <summary>
    /// Whether that Reset was refused, so it should stop asking.
    ///
    /// Both are cleared the moment the project is seen out of break mode, whoever got it there:
    /// the refusal is a fact about ONE stopped session, not about the session that follows it.
    /// </summary>
    private bool _immediateResetRefused;

    /// <summary>
    /// The stop the surface caret was last moved to, as "module:line". The debug poll runs the
    /// stop path every 150ms for as long as the break lasts, and the caret may only follow a
    /// stop ONCE - a caret re-set per tick would take the caret back from the developer the
    /// moment they clicked anywhere else mid-break.
    /// </summary>
    private string? _lastStopFollowed;

    /// <summary>Project modes, as the editor numbers them.</summary>
    private const int RunMode = 0;
    private const int BreakMode = 1;
    private const int DesignMode = 2;

    /// <summary>
    /// Brings a code pane of a STOPPED project forward, so the editor's project-scoped commands
    /// have it to act on, and answers which project that was.
    ///
    /// <see cref="ProjectModeNow"/> reads the ACTIVE project, which is the right reading for
    /// "may this be edited" and the wrong one for "is anything stopped": with two workbooks open
    /// the stopped one need not be the one the editor is attending to, and every command that
    /// would leave the break is aimed at the one it is. This looks at all of them.
    ///
    /// Null when nothing is stopped, or when nothing in the stopped project would come forward -
    /// both of which mean the caller has nothing further to try.
    /// </summary>
    private string? AttendToStoppedProject()
    {
        try
        {
            using var projects = _editor.GetObject("VBProjects");
            var count = projects?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var project = projects!.GetItem(i);
                if (project is null)
                {
                    continue;
                }

                int mode;
                try
                {
                    mode = project.GetInt32("Mode");
                }
                catch (Exception)
                {
                    // A project that will not say is not one to reach into.
                    continue;
                }

                if (mode == DesignMode)
                {
                    continue;
                }

                var name = project.GetString("Name") ?? $"project {i}";
                using var components = project.GetObject("VBComponents");
                var componentCount = components?.GetInt32("Count") ?? 0;

                for (var c = 1; c <= componentCount; c++)
                {
                    using var component = components!.GetItem(c);
                    using var module = component?.GetObject("CodeModule");

                    // Reading CodePane opens one where none was open, which is what makes this
                    // work at all on a project whose windows the developer never touched.
                    using var pane = module?.GetObject("CodePane");
                    if (pane is null)
                    {
                        continue;
                    }

                    pane.Invoke("Show");
                    try
                    {
                        _editor.SetObject("ActiveCodePane", pane);
                    }
                    catch (Exception ex)
                    {
                        Log.Info($"stopped project: {name} would not take the active pane "
                            + $"({ex.GetType().Name})");
                    }

                    return name;
                }

                Log.Info($"stopped project: {name} is stopped and has no code pane to bring forward");
                return null;
            }
        }
        catch (Exception ex)
        {
            Log.Warn($"stopped project: the projects would not be walked ({ex.GetType().Name})");
        }

        return null;
    }

    /// <summary>
    /// The procedure the native caret is sitting in, or null when it is not in one.
    ///
    /// Asked of the EDITOR, through the code module's own ProcOfLine, rather than of the
    /// analyzer: this is the same question the Run command answers for itself, so the same
    /// source gives the same answer, and it cannot be stale the way a parsed outline can be
    /// between a keystroke and the next analysis pass.
    /// </summary>
    private string? ProcedureAtCaret()
    {
        try
        {
            using var pane = _editor.GetObject("ActiveCodePane");
            if (pane is null)
            {
                return null;
            }

            Span<int> selection = stackalloc int[4];
            pane.InvokeInt32s("GetSelection", selection);
            var line = selection[0];
            if (line < 1)
            {
                return null;
            }

            using var module = pane.GetObject("CodeModule");

            // vbext_pk_Proc: the procedure itself rather than a Property Get or Let leg, which is
            // what Run acts on. A line outside every procedure raises rather than answering, and
            // that raise IS the answer.
            var found = module?.CallToString("ProcOfLine", line, 0);
            return string.IsNullOrWhiteSpace(found) ? null : found;
        }
        catch (Exception)
        {
            // The declarations section answers by raising. So does a pane mid-teardown, and both
            // mean the same thing here: there is nothing under the caret to run.
            return null;
        }
    }

    /// <summary>The active project's mode right now, read fresh; design when nothing answers.</summary>
    private int ProjectModeNow()
    {
        try
        {
            using var project = _editor.GetObject("ActiveVBProject");
            return project?.GetInt32("Mode") ?? DesignMode;
        }
        catch
        {
            return DesignMode;
        }
    }

    /// <summary>
    /// The window of a UserForm this process is showing - ThunderDFrame is the runtime form's
    /// own class - or 0 when none stands. Break mode never shows one; run mode with a form up
    /// always does.
    /// </summary>
    private static nint FindStandingFormWindow()
    {
        var pid = (uint)Environment.ProcessId;
        nint frame = 0;
        while ((frame = Win32.FindWindowEx(0, frame, "ThunderDFrame", null)) != 0)
        {
            Win32.GetWindowThreadProcessId(frame, out var owner);
            if (owner == pid)
            {
                return frame;
            }
        }

        return 0;
    }

    /// <summary>How often the execution state is looked at while anything might be running.</summary>
    private const uint DebugPollMilliseconds = 150;

    /// <summary>How often the editor's Immediate window is read while it is being looked at.</summary>
    private const uint ImmediatePollMilliseconds = 300;

    /// <summary>
    /// Polls left before watching stops.
    ///
    /// Running a procedure does not block the call that started it: the command returns and the
    /// code runs afterwards, so the state at the moment the command was issued is always "not
    /// running yet". Checking once found nothing every time, and the stopped line never appeared.
    /// Watching for a while after is the only way to see the transition, and it stops on its own
    /// so that a host sitting idle is not polled forever.
    /// </summary>
    private int _pollsRemaining;

    /// <summary>
    /// Whether VBA will accept a breakpoint on a line.
    ///
    /// Only executable statements can carry one. Asking the editor to set one anywhere else puts a
    /// modal dialog on screen saying so, which is the host's answer to a question the developer did
    /// not ask: they clicked a margin, and a dialog is not a reasonable reply to that. The line is
    /// checked here so the common refusals never reach it.
    ///
    /// Declarations are excluded, not modifiers. A procedure can start with the same words a
    /// module-level declaration does, and a breakpoint on the opening line of a procedure is
    /// perfectly legal, so it is what follows the modifiers that decides.
    /// </summary>
    /// <summary>
    /// The shown module's lines, as the surface holds them - which is what the developer is
    /// clicking in, and so what the breakpoint rule has to be judged against.
    ///
    /// Empty when there is no document, which the rule reads as "no line can carry one": a
    /// margin click on nothing is not a breakpoint.
    /// </summary>
    private string[] SurfaceLines() =>
        _editorSurface?.Text is { } text
            ? text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n')
            : [];

    /// <summary>
    /// Answers the Search panel: the engine searches the modules it holds - live text where a
    /// module is being edited - and the hits come back with workbook display names. The edits
    /// the developer has not paused long enough to write are flushed first, so the search
    /// describes what they see.
    /// </summary>
    /// <summary>The referenced type libraries, loaded on the browser's first question.</summary>
    private TypeLibraryCatalog? _typeLibraries;

    /// <summary>
    /// The catalog of every library the open projects reference, built once. References are
    /// read through the same object model the rest of the session lives on; a reference whose
    /// file is missing or refuses to load is skipped with a log line, not a failure.
    /// </summary>
    private TypeLibraryCatalog TypeLibraries()
    {
        if (_typeLibraries is not null)
        {
            return _typeLibraries;
        }

        var catalog = new TypeLibraryCatalog();

        try
        {
            using var projects = _editor.GetObject("VBProjects");
            var projectCount = projects?.GetInt32("Count") ?? 0;

            for (var p = 1; p <= projectCount; p++)
            {
                using var project = projects!.GetItem(p);
                using var references = project?.GetObject("References");
                var referenceCount = references?.GetInt32("Count") ?? 0;

                for (var r = 1; r <= referenceCount; r++)
                {
                    try
                    {
                        using var reference = references!.GetItem(r);
                        if (reference is null || reference.GetBool("IsBroken"))
                        {
                            continue;
                        }

                        var name = reference.GetString("Name");
                        var description = reference.GetString("Description");
                        var path = reference.GetString("FullPath");
                        if (name is { Length: > 0 } && path is { Length: > 0 })
                        {
                            catalog.AddLibrary(name, description ?? name, path);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Verbose($"typelib: a reference could not be read ({ex.GetType().Name})");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"typelib: the references could not be enumerated ({ex.GetType().Name})");
        }

        _typeLibraries = catalog;
        return catalog;
    }

    /// <summary>The floating Object Browser window, kept across closes once opened.</summary>
    private BrowserPalette? _browserPalette;

    /// <summary>
    /// Summons the Object Browser: a floating themed window of our own (developer,
    /// 2026-08-05), browsing the open projects and every referenced type library. Opened
    /// once and hidden on close, so a second summons brings the same view straight back.
    /// </summary>
    private void OpenBrowserPalette()
    {
        if (_browserPalette is not null)
        {
            _browserPalette.Present();
            return;
        }

        var palette = BrowserPalette.Open(_frame);
        if (palette is null)
        {
            _editorSurface?.Notify("The Object Browser window could not be opened.");
            return;
        }

        palette.LibrariesRequested = BrowseLibraries;
        palette.TypesRequested = BrowseTypes;
        palette.MembersRequested = BrowseMembers;
        palette.NavigateRequested = (module, line, project) =>
        {
            // The editor's window comes forward first: the palette holds the foreground,
            // and a navigation nobody can see is a navigation that did not happen.
            if (_frame != 0)
            {
                Win32.SetForegroundWindow(_frame);
            }

            GoTo(module, line, 1, project);
        };

        DarkenTitleBar(palette.Handle);
        _browserPalette = palette;
    }

    /// <summary>
    /// What the Browser lists at its top level: the open projects first - the developer's
    /// own code is what they browse most - then every referenced type library.
    /// </summary>
    private ObLibraryRow[] BrowseLibraries()
    {
        var rows = new List<ObLibraryRow>();

        try
        {
            using var projects = _editor.GetObject("VBProjects");
            var count = projects?.GetInt32("Count") ?? 0;

            for (var p = 1; p <= count; p++)
            {
                using var project = projects!.GetItem(p);
                if (project is null)
                {
                    continue;
                }

                var identity = ProjectReader.Identity(project);
                rows.Add(new ObLibraryRow(
                    identity.DisplayName,
                    project.GetString("Name") ?? "VBA project",
                    "project"));
            }
        }
        catch (Exception ex)
        {
            Log.Verbose($"object browser: the projects could not be listed ({ex.GetType().Name})");
        }

        foreach (var row in TypeLibraries().Libraries())
        {
            rows.Add(new ObLibraryRow(row.Name, row.Description, "library"));
        }

        Log.Verbose($"object browser: {rows.Count} librarie(s)");
        return [.. rows];
    }

    /// <summary>A library's types: a project answers with its modules, a typelib with its catalog.</summary>
    private ObTypeRow[] BrowseTypes(string library)
    {
        using var project = FindProjectByDisplayName(library);
        if (project is not null)
        {
            var rows = new List<ObTypeRow>();

            try
            {
                using var components = project.GetObject("VBComponents");
                var count = components?.GetInt32("Count") ?? 0;

                for (var i = 1; i <= count; i++)
                {
                    using var component = components!.GetItem(i);
                    var name = component?.GetString("Name");
                    if (name is { Length: > 0 })
                    {
                        rows.Add(new ObTypeRow(name, ComponentKind(component!.GetInt32("Type"))));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Verbose($"object browser: {library}'s modules could not be listed ({ex.GetType().Name})");
            }

            Log.Verbose($"object browser: {library} -> {rows.Count} module(s)");
            return [.. rows];
        }

        var types = TypeLibraries().TypesOf(library) ?? [];
        Log.Verbose($"object browser: {library} -> {types.Count} type(s)");
        return [.. types.Select(row => new ObTypeRow(row.Name, row.Kind))];
    }

    /// <summary>
    /// A type's members. A project module's are read from its own code, line numbers and
    /// all, which is what makes them navigable; a typelib's come from the catalog.
    /// </summary>
    private void BrowseMembers(string library, string type, Action<ObMemberRow[]> reply)
    {
        var projectId = ProjectIdFromDisplay(library);
        if (projectId is not null)
        {
            var members = ScanProjectMembers(type, projectId);
            Log.Verbose($"object browser: {library}.{type} -> {members.Length} member(s) from code");
            reply(members);
            return;
        }

        var rows = TypeLibraries().MembersOf(library, type) ?? [];
        Log.Verbose($"object browser: {library}.{type} -> {rows.Count} member(s)");
        reply([.. rows.Select(row => new ObMemberRow(row.Name, row.Kind, row.Signature, row.Description, 0))]);
    }

    /// <summary>A module's members, read from the module itself.</summary>
    private ObMemberRow[] ScanProjectMembers(string module, string projectId)
    {
        try
        {
            using var component = FindComponent(module, projectId, out _);
            var source = component is null ? null : ProjectReader.ReadSource(component);
            return source is null ? [] : ScanModuleMembers(source);
        }
        catch (Exception ex)
        {
            Log.Verbose($"object browser: {module} could not be read ({ex.GetType().Name})");
            return [];
        }
    }

    /// <summary>The Browser's word for a component type number.</summary>
    private static string ComponentKind(int componentType) => componentType switch
    {
        2 => "class",
        3 => "form",
        100 => "document",
        _ => "module",
    };

    /// <summary>
    /// Finds a module's members by reading its declarations: procedures, module-level
    /// variables and constants, events, API declares, and Enum and Type blocks. A line
    /// scan, not a parse - it reads the declaration lines the developer wrote and skips
    /// procedure bodies, which is exactly the list the Browser shows.
    /// </summary>
    private static ObMemberRow[] ScanModuleMembers(string source)
    {
        var members = new List<ObMemberRow>();
        var lines = source.Replace("\r\n", "\n").Split('\n');
        var inProcedure = false;
        var inBlock = false;

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (line.Length == 0 || line.StartsWith('\''))
            {
                continue;
            }

            if (inBlock)
            {
                if (StartsWithKeyword(line, "End Enum") || StartsWithKeyword(line, "End Type"))
                {
                    inBlock = false;
                }

                continue;
            }

            if (inProcedure)
            {
                if (StartsWithKeyword(line, "End Sub") || StartsWithKeyword(line, "End Function")
                    || StartsWithKeyword(line, "End Property"))
                {
                    inProcedure = false;
                }

                continue;
            }

            // The access modifiers say nothing about what follows; strip them and look.
            var rest = line;
            var hadModifier = false;
            while (TakeKeyword(ref rest, "Public") || TakeKeyword(ref rest, "Private")
                   || TakeKeyword(ref rest, "Friend") || TakeKeyword(ref rest, "Global")
                   || TakeKeyword(ref rest, "Static"))
            {
                hadModifier = true;
            }

            var signature = line.TrimEnd('_').TrimEnd();
            var lineNumber = i + 1;

            if (TakeKeyword(ref rest, "Declare"))
            {
                _ = TakeKeyword(ref rest, "PtrSafe");
                var kind = TakeKeyword(ref rest, "Function") ? "Function"
                    : TakeKeyword(ref rest, "Sub") ? "Sub"
                    : null;
                if (kind is not null && LeadingIdentifier(rest) is { Length: > 0 } declared)
                {
                    members.Add(new ObMemberRow(declared, kind, signature, string.Empty, lineNumber));
                }
            }
            else if (TakeKeyword(ref rest, "Sub"))
            {
                AddNamed(members, rest, "Sub", signature, lineNumber);
                inProcedure = true;
            }
            else if (TakeKeyword(ref rest, "Function"))
            {
                AddNamed(members, rest, "Function", signature, lineNumber);
                inProcedure = true;
            }
            else if (TakeKeyword(ref rest, "Property"))
            {
                _ = TakeKeyword(ref rest, "Get") || TakeKeyword(ref rest, "Let") || TakeKeyword(ref rest, "Set");

                // Get, Let, and Set share the one name; the Browser lists it once.
                var name = LeadingIdentifier(rest);
                if (name.Length > 0
                    && !members.Exists(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase)))
                {
                    members.Add(new ObMemberRow(name, "Property", signature, string.Empty, lineNumber));
                }

                inProcedure = true;
            }
            else if (TakeKeyword(ref rest, "Enum"))
            {
                AddNamed(members, rest, "Enum", signature, lineNumber);
                inBlock = true;
            }
            else if (TakeKeyword(ref rest, "Type"))
            {
                AddNamed(members, rest, "Type", signature, lineNumber);
                inBlock = true;
            }
            else if (TakeKeyword(ref rest, "Const"))
            {
                AddNamed(members, rest, "Const", signature, lineNumber);
            }
            else if (TakeKeyword(ref rest, "Event"))
            {
                AddNamed(members, rest, "Event", signature, lineNumber);
            }
            else if (TakeKeyword(ref rest, "WithEvents"))
            {
                AddNamed(members, rest, "Field", signature, lineNumber);
            }
            else if (hadModifier || TakeKeyword(ref rest, "Dim"))
            {
                // A declaration line can carry several names: Dim a As Long, b As String.
                foreach (var piece in rest.Split(','))
                {
                    var name = LeadingIdentifier(piece.Trim());
                    if (name.Length > 0)
                    {
                        members.Add(new ObMemberRow(name, "Field", piece.Trim(), string.Empty, lineNumber));
                    }
                }
            }
        }

        return [.. members];
    }

    private static void AddNamed(List<ObMemberRow> members, string rest, string kind, string signature, int line)
    {
        var name = LeadingIdentifier(rest);
        if (name.Length > 0)
        {
            members.Add(new ObMemberRow(name, kind, signature, string.Empty, line));
        }
    }

    /// <summary>Whether the line begins with the keyword as a whole word.</summary>
    private static bool StartsWithKeyword(string text, string keyword) =>
        text.StartsWith(keyword, StringComparison.OrdinalIgnoreCase)
        && (text.Length == keyword.Length || !IsIdentifierChar(text[keyword.Length]));

    /// <summary>Consumes a leading keyword and the space after it, reporting whether it was there.</summary>
    private static bool TakeKeyword(ref string text, string keyword)
    {
        if (!StartsWithKeyword(text, keyword))
        {
            return false;
        }

        text = text[keyword.Length..].TrimStart();
        return true;
    }

    /// <summary>The identifier the text starts with, or empty.</summary>
    private static string LeadingIdentifier(string text)
    {
        var end = 0;
        while (end < text.Length && IsIdentifierChar(text[end]))
        {
            end++;
        }

        return text[..end];
    }

    private static bool IsIdentifierChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    private void OnSearchRequested(int id, string query, bool matchCase, bool wholeWord, string scope)
    {
        _editorSurface?.FlushEdits();

        var projectId = scope is "module" or "project" ? _shownProject : null;
        var module = scope is "module" ? _editorSurface?.Module : null;

        _ = RunSearchAsync(id, query, matchCase, wholeWord, scope, projectId, module, replacement: null);
    }

    /// <summary>
    /// Replace across a scope: the same search, then every hit's line rewritten through the
    /// module the hit lives in - ReplaceLine, surgical, no module reset. The shown module's
    /// resync adopts the change the way it adopts any edit made outside the surface.
    /// </summary>
    private void OnReplaceAllRequested(int id, string query, bool matchCase, bool wholeWord, string scope, string replacement)
    {
        _editorSurface?.FlushEdits();

        var projectId = scope is "module" or "project" ? _shownProject : null;
        var module = scope is "module" ? _editorSurface?.Module : null;

        _ = RunSearchAsync(id, query, matchCase, wholeWord, scope, projectId, module, replacement);
    }

    private async Task RunSearchAsync(
        int id, string query, bool matchCase, bool wholeWord, string scope,
        string? projectId, string? module, string? replacement)
    {
        try
        {
            var result = _analysis is null
                ? null
                : await _analysis.SearchAsync(scope, projectId, module, query, matchCase, wholeWord, CancellationToken.None)
                    .ConfigureAwait(false);

            // Back to the host thread: replacement drives the object model, and even a plain
            // answer touches the surface.
            _editorSurface?.RunOnHostThread(() =>
            {
                var matches = result?.Matches ?? [];
                var replaced = 0;

                if (replacement is not null && matches.Length > 0)
                {
                    replaced = ReplaceMatches(matches, query, matchCase, wholeWord, replacement);

                    // The modules changed under everything that mirrors them. The shown one is
                    // resynced NOW - waiting for a pane event left the editor showing the old
                    // text while the panel claimed the replacement had happened ("replace is
                    // not working", 2026-08-04) - and the full pass re-reads the rest.
                    ResyncFromModule();
                    _resyncPanePolls = Math.Max(_resyncPanePolls, 2);
                    _analysis?.Reanalyse();
                }

                // A replace answers with its count and an empty list rather than the matches
                // it just destroyed: re-listing the old previews read as "still matching".
                _editorSurface?.ShowSearchResults(
                    id,
                    replacement is not null
                        ? []
                        : [.. matches.Select(m => new SurfaceSearchMatch(
                            DisplayFromProjectId(m.ProjectId), m.Module, m.Line, m.Column, m.Length, m.Preview))],
                    result?.Truncated ?? false,
                    replaced);

                Log.Info($"search: '{query}' in {scope} -> {matches.Length} match(es)"
                         + (replacement is null ? string.Empty : $", {replaced} replaced"));
            });
        }
        catch (Exception ex)
        {
            Log.Error($"search: '{query}' failed", ex);
            _editorSurface?.RunOnHostThread(() => _editorSurface?.ShowSearchResults(id, [], truncated: false));
        }
    }

    /// <summary>
    /// Rewrites every matched line through its own module. Grouped per module and applied
    /// bottom-up per line so earlier replacements cannot move later ones; each line is re-read
    /// and re-matched at the moment of writing, so a stale hit rewrites nothing.
    /// </summary>
    private int ReplaceMatches(EngineSearchMatch[] matches, string query, bool matchCase, bool wholeWord, string replacement)
    {
        var replaced = 0;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        foreach (var group in matches.GroupBy(m => (m.ProjectId, m.Module)))
        {
            try
            {
                using var component = FindComponent(group.Key.Module, group.Key.ProjectId, out _);
                using var code = component?.GetObject("CodeModule");
                if (code is null)
                {
                    continue;
                }

                var inModule = 0;
                foreach (var line in group.Select(m => m.Line).Distinct().OrderByDescending(l => l))
                {
                    var text = code.GetStringIndexed("Lines", line, 1);
                    if (text is null)
                    {
                        continue;
                    }

                    var rewritten = ReplaceInLine(text, query, replacement, comparison, wholeWord, out var count);
                    if (count > 0)
                    {
                        code.Invoke("ReplaceLine", line, rewritten);
                        inModule += count;
                    }
                }

                replaced += inModule;

                // A replace is a host rewrite: the engine's live copy of this module - which
                // outranks its seeded copy in every answer - still holds the pre-replace text
                // for any module the developer has typed in, and the dirty-dot comparison
                // still holds the pre-replace text as current. Both adopt the truth here.
                if (inModule > 0 && component is not null
                    && ProjectReader.ReadSource(component) is { } adopted)
                {
                    var display = DisplayFromProjectId(group.Key.ProjectId);
                    _writtenModules[WrittenKey(group.Key.Module, display)] = adopted;
                    _analysis?.NotifyLiveText(group.Key.Module, adopted, null, group.Key.ProjectId);

                    // AND THE SURFACE, which is the copy the developer is looking at.
                    //
                    // It cannot be left to the ResyncFromModule call that follows this, and the
                    // reason is the line above. That resync asks "has the module changed since we
                    // last wrote it?" by comparing the module against `_writtenModules` - so
                    // updating the baseline here, which the dirty dot needs, tells the resync that
                    // nothing changed outside the surface and it skips every module this just
                    // rewrote. The two lines disabled each other: the workbook held the
                    // replacement, the editor went on showing the text from before it, and the
                    // next keystroke would have written the stale text back over the replacement.
                    //
                    // Measured 2026-08-11 by the new search suite. This is the second time this
                    // exact symptom has been fixed - "replace is not working", 2026-08-04, which
                    // is what added the resync - and it came back because the fix depended on a
                    // heuristic that a later, unrelated, correct change made blind.
                    _editorSurface?.Sync(group.Key.Module, display, adopted);
                }
            }
            catch (Exception ex)
            {
                Log.Info($"search: could not replace in {group.Key.Module} ({ex.GetType().Name})");
            }
        }

        return replaced;
    }

    /// <summary>Every occurrence in one line, with the same matching the engine searched by.</summary>
    private static string ReplaceInLine(
        string line, string query, string replacement, StringComparison comparison, bool wholeWord, out int count)
    {
        count = 0;
        var result = new System.Text.StringBuilder(line.Length);
        var from = 0;

        while (from < line.Length)
        {
            var at = line.IndexOf(query, from, comparison);
            if (at < 0)
            {
                break;
            }

            var boundaryBefore = at == 0 || !IsWordCharacter(line[at - 1]);
            var boundaryAfter = at + query.Length >= line.Length || !IsWordCharacter(line[at + query.Length]);

            if (wholeWord && (!boundaryBefore || !boundaryAfter))
            {
                result.Append(line, from, at + 1 - from);
                from = at + 1;
                continue;
            }

            result.Append(line, from, at - from);
            result.Append(replacement);
            from = at + query.Length;
            count++;
        }

        result.Append(line, from, line.Length - from);
        return result.ToString();
    }

    private static bool IsWordCharacter(char c) => char.IsLetterOrDigit(c) || c == '_';

    /// <summary>Moves line-anchored breakpoints with the text they were set on.</summary>
    private void OnLinesShifted(string module, int afterLine, int delta)
    {
        var key = BreakpointKey(module);
        if (!_breakpoints.TryGetValue(key, out var record) || record.Lines.Count == 0)
        {
            return;
        }

        var lines = record.Lines;

        var moved = new SortedSet<int>();
        foreach (var line in lines)
        {
            moved.Add(line > afterLine ? Math.Max(afterLine, line + delta) : line);
        }

        if (!moved.SetEquals(lines))
        {
            _breakpoints[key] = record with { Lines = moved };
            PublishBreakpoints();
            Log.Verbose($"breakpoint: shifted with the text, now [{string.Join(",", moved)}]");
        }
    }

    /// <summary>Toggles a breakpoint on a line of the module currently shown.</summary>
    private void ToggleBreakpoint(int line)
    {
        var module = _editorSurface?.Module;
        if (module is null || line < 1)
        {
            return;
        }

        // Clearing is never validity-gated. A recorded dot must always answer a click,
        // whatever its line has since become: gating the clear on the line still being
        // executable left a drifted dot that five clicks could not remove (2026-08-04).
        var key = BreakpointKey(module);
        var clearing = _breakpoints.TryGetValue(key, out var recorded) && recorded.Lines.Contains(line);

        // THE WHOLE MODULE, not the one line. Whether a line can carry a breakpoint depends on
        // whether it sits inside a Type or an Enum, which the line itself cannot say - see
        // Core.Editor.Breakpoints, and the margin dots on Enum members that came of asking it
        // one line at a time.
        if (!clearing && !Core.Editor.Breakpoints.CanCarry(SurfaceLines(), line))
        {
            // Refused silently on screen, by design (the developer, 2026-08-04): the hover
            // preview already showed an orange cross where no breakpoint can go, and a click
            // there draws nothing - nothing may ever appear that looks like a breakpoint the
            // developer did not get. The page mirrors Core.Editor.Breakpoints for that preview.
            Log.Info($"breakpoint: {module}({line}) is not an executable statement");
            return;
        }

        try
        {
            // The command acts on the caret, so the caret is put on the line first. Everything the
            // developer typed goes with it: a breakpoint is set by line number, and writing the
            // module afterwards would move it.
            _editorSurface?.FlushEdits();

            // The SHOWN project's copy, not the first module of that name. Without it the
            // command lands in whichever workbook answered first, which is how a breakpoint got
            // set on the twin.
            using var pane = FindCodePane(module, _shownProject);
            if (pane is null)
            {
                return;
            }

            pane.Invoke("SetSelection", line, 1, line, 1);

            if (!VbeCommands.Execute(_editor, VbeCommands.Command.ToggleBreakpoint).Ran)
            {
                return;
            }

            if (!_breakpoints.TryGetValue(key, out var held))
            {
                held = new BreakpointRecord(module, DisplayFromProjectId(_shownProject), []);
                _breakpoints[key] = held;
            }

            var lines = held.Lines;
            if (!lines.Remove(line))
            {
                lines.Add(line);
            }

            _editorSurface?.ShowBreakpoints([.. lines]);
            Log.Info($"breakpoint: {module}({line}) {(lines.Contains(line) ? "set" : "cleared")}");
        }
        catch (Exception ex)
        {
            Log.Error($"breakpoint: {module}({line}) could not be toggled", ex);
        }
    }

    /// <summary>Sends the surface the breakpoints belonging to the module it is showing.</summary>
    private void PublishBreakpoints()
    {
        var module = _editorSurface?.Module;
        if (module is null)
        {
            return;
        }

        _editorSurface?.ShowBreakpoints([.. BreakpointsFor(module)]);
    }

    /// <summary>
    /// The component whose properties the panel shows: the explorer's selection, or the shown
    /// module when nothing has been selected.
    /// </summary>
    private string? _propertiesTarget;

    /*
     * The document-component properties that are safe to read, which are the ones the editor's own
     * Properties window shows. That window filters to the properties the type library marks
     * browsable; until the type library is read directly (the IntelliSense track), these lists ARE
     * that filter. They are not an aesthetic choice: reading a property runs its getter, and some
     * of the unlisted getters do real work. Reading a workbook's mail properties starts the mail
     * system's profile wizard on a machine with none, which is how this was learned.
     */

    private static readonly string[] WorksheetProperties =
    [
        "Name", "DisplayPageBreaks", "DisplayRightToLeft", "EnableAutoFilter", "EnableCalculation",
        "EnableFormatConditionsCalculation", "EnableOutlining", "EnablePivotTable",
        "EnableSelection", "ScrollArea", "StandardWidth", "Visible",
    ];

    private static readonly string[] WorkbookProperties =
    [
        "AccuracyVersion", "AutoUpdateFrequency", "AutoUpdateSaveChanges",
        "ChangeHistoryDuration", "ConflictResolution", "Date1904", "DisplayDrawingObjects",
        "DisplayInkComments", "EnableAutoRecover", "EncryptionProvider", "EnvelopeVisible",
        "Final", "ForceFullCalculation", "HighlightChangesOnScreen", "InactiveListBorderVisible",
        "IsAddin", "KeepChangeHistory", "ListChangesOnNewSheet", "Password",
        "PrecisionAsDisplayed", "ReadOnlyRecommended", "RemovePersonalInformation", "Saved",
        "SaveLinkValues", "ShowConflictHistory", "ShowPivotChartActiveFields",
        "ShowPivotTableFieldList", "TemplateRemoveExtData", "UpdateLinks",
        "UpdateRemoteReferences",
    ];

    /// <summary>Component types, as the editor numbers them.</summary>
    private const int DocumentComponent = 100;

    /// <summary>
    /// What a document component is, and which of its properties may be read, told from the names
    /// its collection carries. Names are safe to enumerate; it is values that run getters. A
    /// document kind this does not recognise shows only its names, which loses detail and starts
    /// nothing.
    /// </summary>
    private static (string Kind, string[]? Allowed) ClassifyDocument(HashSet<string> names)
    {
        if (names.Contains("StandardWidth"))
        {
            return ("Worksheet", WorksheetProperties);
        }

        if (names.Contains("Date1904"))
        {
            return ("Workbook", WorkbookProperties);
        }

        return ("Document", ["Name"]);
    }

    /// <summary>
    /// Sends the surface the properties of the selected component, shaped the way the editor's own
    /// Properties window shapes them: an object header naming the component and its class, the
    /// code name as "(Name)" sorted first, and for a document component the host object's
    /// browsable properties alongside it.
    /// </summary>
    /// <summary>
    /// The control the Properties panel is aimed at, on the form `_propertiesTarget` names -
    /// the canvas selection's half of M4. Null means the panel shows the component itself,
    /// which is everything that existed before controls could be selected. Cleared wherever
    /// the target moves to another component.
    /// </summary>
    private string? _propertiesControl;

    /// <summary>
    /// What the panel's numbers MEAN, read from each object's type library once per class: enum
    /// members by name, colours as the hex the VBE spells. Lives for the session because a
    /// loaded type does not change, and the panel republishes on every selection.
    /// </summary>
    private readonly PropertyTypes _propertyTypes = new();

    /// <summary>
    /// What each control KIND holds untouched, measured once per kind from a bare instance of
    /// its coclass, and what each of those properties means. The inventory the markup projection
    /// will compare against to print only what a developer changed, and the vocabulary the markup
    /// tab's completions and hovers answer from; readable through `defaults?type=` and
    /// `vocabulary`, which is how the claims get checked before anything depends on them.
    /// </summary>
    private readonly ControlDefaults _controlDefaults;

    /// <summary>
    /// The markup language's vocabulary, sent to the page for its completions and hovers. Measured
    /// from coclasses and type libraries, so the only per-form part is the Form entry itself - the
    /// module names the live form that describes it.
    /// </summary>
    private void PublishFormMarkupVocabulary(string module, string? projectDisplay)
    {
        try
        {
            using var component = FindComponent(
                module, ProjectIdFromDisplay(projectDisplay) ?? _shownProject, out _);
            _editorSurface?.PublishFormMarkupVocabulary(
                FormMarkupVocabulary.Of(_controlDefaults, _propertyTypes, component));
        }
        catch (Exception ex)
        {
            // A vocabulary that cannot be read leaves the tab's completions empty, which is the
            // surface this product had yesterday - not a reason to take the tab down.
            Log.Info($"markup vocabulary: {module} could not be described ({ex.GetType().Name}: {ex.Message.Trim()})");
        }
    }

    private void PublishProperties()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        var target = _propertiesTarget ?? surface.Module;
        if (target is null)
        {
            ShowNoProperties();
            return;
        }

        if (_propertiesControl is { Length: > 0 } selectedControl)
        {
            PublishControlProperties(target, selectedControl);
            return;
        }

        try
        {
            using var found = FindComponent(target, null, out var foundIn);
            using var properties = found?.GetObject("Properties");
            if (found is null || properties is null)
            {
                // The component the panel is aimed at is not there any more. Republishing is what
                // a removal does; noticing that the target went with it is THIS method's job,
                // because the alternative is every removal path remembering to clear the aim and
                // one of them forgetting - which is what happened, and the panel went on
                // describing a form the tree no longer listed (reported 2026-08-15).
                if (ForgetMissingTarget(target, surface))
                {
                    PublishProperties();
                }

                return;
            }

            var componentType = found.GetInt32("Type");
            var count = properties.GetInt32("Count");

            // What the values MEAN, from the type library of the thing that owns them - and what
            // the library says is not the developer's business, which is how the panel knows to
            // leave a row out rather than draw it as "[object]".
            using var typed = TypeSourceOf(found, target, foundIn);
            var described = typed is null ? null : _propertyTypes.Describe(typed);

            // Names first, values second. Enumerating names runs nothing; it is the value reads
            // that must be limited to what is known to be safe.
            var names = new List<string>(count);
            for (var i = 1; i <= count; i++)
            {
                using var property = properties.GetItem(i);

                try
                {
                    if (property?.GetString("Name") is { Length: > 0 } name)
                    {
                        names.Add(name);
                    }
                }
                catch (Exception)
                {
                    // A property that will not even say its name has nothing to offer the panel.
                }
            }

            string kind;
            HashSet<string>? allowed;

            if (componentType == DocumentComponent)
            {
                var (documentKind, list) = ClassifyDocument(new HashSet<string>(names, StringComparer.OrdinalIgnoreCase));
                kind = documentKind;
                allowed = list is null ? null : new HashSet<string>(list, StringComparer.OrdinalIgnoreCase);
            }
            else
            {
                kind = componentType switch
                {
                    1 => "Module",
                    2 => "Class Module",
                    3 => "UserForm",
                    11 => "ActiveX Designer",
                    _ => "Component",
                };
                allowed = null;
            }

            var entries = new List<SurfacePropertyEntry>(names.Count + 1);

            // The code name. A document component's collection does not carry it (its Name is the
            // host object's), so it is added here; everywhere else the collection's Name IS the
            // code name and is shown under the same spelling the editor uses.
            if (componentType == DocumentComponent)
            {
                entries.Add(new SurfacePropertyEntry("(Name)", target, true, false));
            }

            foreach (var name in names)
            {
                if (allowed is not null && !allowed.Contains(name))
                {
                    continue;
                }

                // A leading underscore is a library's own business, the same rule the Object
                // Browser and the defaults walk already keep - `_Font_Reserved` is nobody's
                // property. So is anything the library marks hidden or restricted.
                PropertyTypes.Described? meaning = null;
                described?.TryGetValue(name, out meaning);
                if (name.StartsWith('_') || meaning is { Hidden: true })
                {
                    continue;
                }

                using var property = properties.GetItem(name);
                if (property is null)
                {
                    continue;
                }

                var isCodeName = componentType != DocumentComponent
                    && string.Equals(name, "Name", StringComparison.OrdinalIgnoreCase);
                var shownName = isCodeName ? "(Name)" : name;

                // The code name is the VBE's, not the designer library's, so the library has no
                // say in whether it can be written: it is the rename gesture, and it stays.
                var row = DescribeProperty(shownName, property, isCodeName ? null : meaning);

                // AN OBJECT IS NOT A VALUE, and a row that says so helps nobody (the owner,
                // 2026-08-16: "what about properties that say [object]?"). A FONT is the one
                // object this panel serves as its parts, the way the control rows do, and a
                // PICTURE is the one it serves as itself. The rest go: `Controls` and
                // `ActiveControl` are runtime state the native panel does not show either.
                if (row.Value == ObjectValue)
                {
                    // The object comes from the DESIGNER, not from the VBE property wrapper
                    // around it: a Property's own `Value` hands back a font this side cannot
                    // read through, where the designer answers the same font the control rows
                    // read - and the same picture the canvas paints.
                    if (string.Equals(name, "Font", StringComparison.OrdinalIgnoreCase)
                        && typed is not null)
                    {
                        using var font = typed.GetDispId("Font") != DispId.Unknown
                            ? typed.GetObject("Font")
                            : null;
                        entries.AddRange(FontRows(font));
                    }
                    else if (FormDesignService.IsPictureSlot(name) && typed is not null)
                    {
                        entries.Add(PictureRow(typed, name));
                    }

                    continue;
                }

                // A ROW NOBODY CAN SET IS NOT A PROPERTY OF THE DESIGN (the owner, 2026-08-16:
                // "if theyre jot settable; dint suow them"). On a UserForm that is `CanPaste`,
                // `CanUndo` and `CanRedo` - questions about the editing session rather than
                // about the form, which read False on a form nobody has touched and cannot be
                // written at all - plus `InsideWidth` and `InsideHeight`, which the canvas
                // reads for its parity but which a developer cannot type into. A panel is for
                // setting properties, and everything left in it now can be set.
                if (!row.Writable)
                {
                    continue;
                }

                entries.Add(row);
            }

            // Alphabetical, which puts "(Name)" first: the parenthesis sorts before any letter,
            // and that is the point of the parenthesis.
            entries.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
            surface.ShowProperties(target, kind, [.. entries]);
            Log.Info($"properties: {target} ({kind}), showing {entries.Count} of {count}");
        }
        catch (Exception ex)
        {
            Log.Error($"properties: {target} could not be read", ex);
        }
    }

    /// <summary>
    /// The object whose type library describes what a component's property values MEAN.
    ///
    /// A form's design properties are the DESIGNER's, so that is what a form answers. A DOCUMENT
    /// component's are the HOST object's - the worksheet or the workbook itself, over in Excel's
    /// object model rather than the editor's. Until this, a document component had no designer
    /// and therefore no library, so a worksheet's rows read `Visible -1` and `EnableSelection 0`
    /// where the language says `xlSheetVisible` and `xlNoRestrictions` (the owner, 2026-08-15:
    /// "please update enums for workbook / worksheet properties").
    ///
    /// NOT through the property bag, which was the first attempt and looked right: a VBE
    /// `Property` answers `Object` with its own VALUE, not with the thing it belongs to, so
    /// asking the first property in a worksheet's collection returned Excel's Application and
    /// the panel learned the enums of the wrong class entirely.
    ///
    /// The workbook comes from the same trust-free application route the dirty dots use, and the
    /// sheet is matched by CODE NAME - which is what a document component's name IS, and which
    /// keeps a chart sheet reading a chart's library rather than a neighbouring worksheet's.
    ///
    /// A component with neither keeps the raw numbers, which is what the panel showed before any
    /// of this and is still honest.
    /// </summary>
    private DispatchObject? TypeSourceOf(DispatchObject component, string name, string? projectId)
    {
        try
        {
            if (component.GetInt32("Type") != DocumentComponent)
            {
                return component.GetObject("Designer");
            }

            if (DisplayFromProjectId(projectId) is not { Length: > 0 } display)
            {
                return null;
            }

            // Not a `using`: the workbook IS the answer for the workbook document, and a caller
            // that owns what it is given cannot be handed something already released.
            var workbook = FindWorkbookByDisplay(display);
            if (workbook is null)
            {
                return null;
            }

            // The workbook document is the one whose code name is the workbook's own.
            if (string.Equals(workbook.GetString("CodeName"), name, StringComparison.OrdinalIgnoreCase))
            {
                return workbook;
            }

            try
            {
                using var sheets = workbook.GetObject("Sheets");
                var count = sheets?.GetInt32("Count") ?? 0;
                for (var i = 1; i <= count; i++)
                {
                    var sheet = sheets!.GetItem(i);
                    if (sheet is not null
                        && string.Equals(sheet.GetString("CodeName"), name, StringComparison.OrdinalIgnoreCase))
                    {
                        return sheet;
                    }

                    sheet?.Dispose();
                }
            }
            finally
            {
                workbook.Dispose();
            }

            return null;
        }
        catch (Exception ex)
        {
            // Nothing to say the developer can act on: the rows fall back to their raw values,
            // which is a panel that reads worse rather than a panel that lies.
            Log.Info($"properties: no type library for {name} ({ex.GetType().Name})");
            return null;
        }
    }

    /// <summary>
    /// The panel with nothing in it: no object, no rows. What it looks like before anything has
    /// been selected, and what it must go back to when the selection stops existing - a panel
    /// that keeps its last picture is a panel describing something that is not there.
    /// </summary>
    private void ShowNoProperties() => _editorSurface?.ShowProperties(string.Empty, string.Empty, []);

    /// <summary>
    /// Drops an aim at a component that is no longer in the project. True when there is a
    /// different one to fall back to, so the caller may publish again; false when the panel has
    /// been emptied and there is nothing more to do.
    /// </summary>
    private bool ForgetMissingTarget(string missing, EditorSurface surface)
    {
        Log.Info($"properties: {missing} is gone, the panel lets it go");
        _propertiesTarget = null;
        _propertiesControl = null;

        // The shown module, unless the shown module IS the thing that went - which is the case
        // when the removed component was the open one, and is also the guard that keeps the
        // caller's re-entry to a single hop.
        if (surface.Module is { Length: > 0 } shown
            && !string.Equals(shown, missing, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        ShowNoProperties();
        return false;
    }

    /// <summary>
    /// A control row's edit: through the designer service, the same write an api `set`
    /// makes, so the panel, the routes and the markup apply stay one operation. A rename
    /// carries the selection with it; the code-behind's handlers stay put, which is the
    /// native designer's own behaviour.
    /// </summary>
    private void OnControlPropertyEdit(string formName, string controlName, string name, string value)
    {
        try
        {
            using var component = FindComponent(formName);
            using var designer = component?.GetObject("Designer");
            if (component is null || designer is null)
            {
                _propertiesControl = null;
                PublishProperties();
                return;
            }

            var property = string.Equals(name, "(Name)", StringComparison.Ordinal) ? "Name" : name;

            // The panel shows the developer's spelling, the model stores a number: the same
            // translation the component rows make, on the way in.
            var stored = value;
            using (var control = FormDesignService.FindControlNamed(designer, controlName, 0))
            {
                if (control is not null && _propertyTypes.Of(control).TryGetValue(property, out var shape))
                {
                    stored = PropertyTypes.Unspell(shape, value);
                }
            }

            var display = FormDesignService.SetControlProperty(component, designer, controlName, property, stored, null);
            Log.Info($"properties: {formName}.{controlName}.{property} = {display}");

            if (string.Equals(property, "Name", StringComparison.OrdinalIgnoreCase))
            {
                _propertiesControl = value;
            }

            FormDesignService.KeepDesignerDown(component);
            RefreshDesignerTabFor(formName);
            PublishProperties();
        }
        catch (Exception ex)
        {
            Log.Error($"properties: {formName}.{controlName}.{name} could not be set", ex);
            _editorSurface?.Notify($"{name}: {ex.Message}");
            PublishProperties();
        }
    }

    /// <summary>
    /// The panel aimed at a CONTROL: the curated rows the designer service speaks - identity,
    /// caption, geometry, state, colours, font - read tolerantly off the live control, and
    /// written back through the same service an api `set` uses. The native panel enumerates
    /// the control's whole typelib; this one shows what the model can honestly round-trip,
    /// and grows with the service.
    /// </summary>
    private void PublishControlProperties(string formName, string controlName)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        try
        {
            using var component = FindComponent(formName);
            using var designer = component?.GetObject("Designer");
            using var control = designer is null
                ? null
                : FormDesignService.FindControlNamed(designer, controlName, 0);
            if (component is null || control is null)
            {
                // The control is gone - removed, or the markup renamed it. The panel falls
                // back to the form rather than showing rows of nothing.
                _propertiesControl = null;
                PublishProperties();
                return;
            }

            var kind = FormDesignService.FriendlyTypeOf(control);
            var entries = new List<SurfacePropertyEntry>
            {
                new("(Name)", controlName, true, false),
            };

            void Number(string name)
            {
                if (FormDesignService.TryNumber(control, name) is { } value)
                {
                    entries.Add(new SurfacePropertyEntry(name, FormDesignService.FormatNumber(value), true, false));
                }
            }

            void Flag(string name)
            {
                if (FormDesignService.TryFlag(control, name) is { } value)
                {
                    entries.Add(new SurfacePropertyEntry(name, value ? "True" : "False", true, true));
                }
            }

            // The control's own type library, for the same reason the component's rows read one:
            // a colour is hex and an enum is a name, in both panels, because they are one panel
            // as far as the developer is concerned.
            var shapes = _propertyTypes.Of(control);

            void Colour(string name)
            {
                if (FormDesignService.TryInt(control, name) is { } value)
                {
                    shapes.TryGetValue(name, out var shape);
                    var raw = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    entries.Add(new SurfacePropertyEntry(name, PropertyTypes.Spell(shape, raw), true, false,
                        null, FormDesignService.OleColorToCss(value)));
                }
            }

            // A property whose values the library NAMES - a control's BorderStyle, SpecialEffect,
            // MousePointer, TextAlign - shown as the name, offered as the list, typed as either.
            void Named(string name)
            {
                if (!shapes.TryGetValue(name, out var shape) || shape.Members is not { Count: > 0 } members)
                {
                    return;
                }

                if (FormDesignService.TryInt(control, name) is { } value)
                {
                    var raw = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    entries.Add(new SurfacePropertyEntry(name, PropertyTypes.Spell(shape, raw), true, false,
                        [.. members.Select(member => member.Name)]));
                }
            }

            if (FormDesignService.TryText(control, "Caption") is { } caption)
            {
                entries.Add(new SurfacePropertyEntry("Caption", caption, true, false));
            }

            Number("Left");
            Number("Top");
            Number("Width");
            Number("Height");
            Flag("Enabled");
            Flag("Visible");
            if (FormDesignService.TryInt(control, "TabIndex") is { } tabIndex)
            {
                entries.Add(new SurfacePropertyEntry("TabIndex", tabIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), true, false));
            }

            // TABSTOP, which is how a control is taken OUT of the tab order without being taken
            // out of the sequence (the owner, 2026-08-18: "is there a way to take a control out
            // of the tab order?"). It was reachable from nothing: the native Properties window
            // offers it, this panel did not, and the dialect cannot spell it - MSForms packs it
            // into VariousPropertyBits, the one saved-mask field that names many properties with
            // one bit, so the projection deliberately never reads the document's word for it.
            // `Flag` adds nothing it cannot read, which is what keeps it off an Image.
            Flag("TabStop");

            Colour("BackColor");
            Colour("ForeColor");
            // Offered only where the control HAS them: a Label has no SpecialEffect worth
            // showing and a CommandButton no TextAlign, and `Named` adds nothing it cannot read.
            Named("BorderStyle");
            Named("SpecialEffect");
            Named("TextAlign");
            Named("MousePointer");

            try
            {
                using var font = control.GetDispId("Font") != DispId.Unknown ? control.GetObject("Font") : null;
                entries.AddRange(FontRows(font));
            }
            catch
            {
                // A control without a font shows none.
            }

            // The picture rows, where the control HAS them: an Image and a Frame carry a
            // surface picture, a button carries one beside its caption, a Label either - and a
            // TextBox carries none at all, so nothing is offered on one.
            foreach (var slot in new[] { "Picture", "MouseIcon" })
            {
                if (control.GetDispId(slot) != DispId.Unknown)
                {
                    entries.Add(PictureRow(control, slot));
                }
            }

            // How the picture SITS: a surface control takes a size mode and an alignment, a
            // caption control a position relative to its caption, and `Named` adds nothing to a
            // control that has neither. Offered whether or not a picture is loaded, because that
            // is what the FORM's panel does one method over and what the native window does -
            // they are settings about pictures, not about this picture.
            Named("PictureSizeMode");
            Named("PictureAlignment");
            Named("PicturePosition");
            Flag("PictureTiling");

            FormDesignService.KeepDesignerDown(component);

            entries.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
            surface.ShowProperties(controlName, kind, [.. entries]);
            Log.Info($"properties: {formName}.{controlName} ({kind}), showing {entries.Count} control row(s)");
        }
        catch (Exception ex)
        {
            Log.Error($"properties: {formName}.{controlName} could not be read", ex);
        }
    }

    /// <summary>
    /// One property, rendered for the panel. Whether it is offered for editing comes from the type
    /// it currently holds: values of simple types are editable, objects and the unreadable are not.
    /// The editor can still refuse an edit, and that refusal is reported when it happens.
    /// </summary>
    /// <summary>
    /// A font as the panel shows one: its parts, each a row a developer can edit, because a font
    /// is the one object-valued property this panel can serve. Both panels build them here - the
    /// control's, and now the component's - so a form's font rows and a button's are the same
    /// rows rather than two lists that agree today.
    /// </summary>
    /// <summary>
    /// The sizes a font list offers. This one is WRITTEN DOWN rather than measured, unlike the
    /// faces beside it and unlike everything else in this panel: a point size is a number, not a
    /// capability, so there is nothing to ask the machine. It is the ramp every office
    /// application offers, and the row still takes any number typed into it.
    /// </summary>
    private static readonly string[] FontSizes =
        ["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "28", "36", "48", "72"];

    private static List<SurfacePropertyEntry> FontRows(DispatchObject? font)
    {
        var rows = new List<SurfacePropertyEntry>(4);
        if (font is null)
        {
            return rows;
        }

        if (FormDesignService.TryText(font, "Name") is { } fontName)
        {
            // The machine's own faces, offered beside the field rather than instead of it (the
            // owner, 2026-08-16: "should font be a drop down selector for all properties?"). The
            // row keeps taking free text, because MSForms stores a face as a string and a form
            // written elsewhere may name a font this machine has never had.
            rows.Add(new SurfacePropertyEntry(
                "Font.Name", fontName, true, false, [.. InstalledFonts.All]));
        }

        if (FormDesignService.TryNumber(font, "Size") is { } fontSize)
        {
            rows.Add(new SurfacePropertyEntry(
                "Font.Size", FormDesignService.FormatNumber(fontSize), true, false, FontSizes));
        }

        if (FormDesignService.TryFlag(font, "Bold") is { } bold)
        {
            rows.Add(new SurfacePropertyEntry("Font.Bold", bold ? "True" : "False", true, true));
        }

        if (FormDesignService.TryFlag(font, "Italic") is { } italic)
        {
            rows.Add(new SurfacePropertyEntry("Font.Italic", italic ? "True" : "False", true, true));
        }

        return rows;
    }

    /// <summary>
    /// A picture as the panel shows one: what it HOLDS in the native designer's own words
    /// (`(None)`, `(Bitmap)`, `(Icon)`, `(Metafile)`) and the picture itself as a thumbnail
    /// where there are pixels to show. Writable, because the row's Browse is the write - a
    /// picture is chosen from a file rather than typed.
    /// </summary>
    private static SurfacePropertyEntry PictureRow(DispatchObject owner, string property)
    {
        using var picture = PictureBytes.PictureOn(owner, property);
        return new SurfacePropertyEntry(
            property, PictureBytes.Describe(picture), true, false, null, null,
            true, picture is null ? null : PictureBytes.DataUriOf(picture));
    }

    /// <summary>What a VBE Property whose value is an object reads as, which is the signal the
    /// panel drops the row on.</summary>
    private const string ObjectValue = "[object]";

    private static SurfacePropertyEntry DescribeProperty(
        string shownName, DispatchObject property, PropertyTypes.Described? meaning)
    {
        var shape = meaning is null || (meaning.Members is null && !meaning.Colour)
            ? null
            : new PropertyTypes.Shape(meaning.Members, meaning.Colour);

        try
        {
            var (kind, display) = property.ReadProperty("Value");

            // WRITABLE is a fact about the type, not about the value that came back. The variant
            // rule below was all this had, and it reads `CanPaste` as an editable Boolean - it is
            // a Boolean, and it has no setter (the owner, 2026-08-16, on the rows that say
            // "[object]"; this is the same defect one row over). The library's own PUT is the
            // answer where there is a library; the variant rule stays as the fallback for a
            // component whose type nothing here can reach.
            var takesAValue = kind is VarEnum.VT_BSTR or VarEnum.VT_BOOL
                or VarEnum.VT_I2 or VarEnum.VT_I4 or VarEnum.VT_INT
                or VarEnum.VT_R4 or VarEnum.VT_R8 or VarEnum.VT_EMPTY;
            var writable = takesAValue && (meaning is null || meaning.Settable);

            // The value as the developer WRITES it - fmCycleAllForms rather than 0, &H8000000F&
            // rather than -2147483633 - with the enum's members offered as the row's choices,
            // and a colour's CSS so the row can show the colour itself rather than its digits.
            var spelled = PropertyTypes.Spell(shape, display);
            var options = shape?.Members is { Count: > 0 } members
                ? members.Select(member => member.Name).ToArray()
                : null;
            var swatch = shape?.Colour == true
                && int.TryParse(display, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var ole)
                ? FormDesignService.OleColorToCss(ole)
                : null;

            return new SurfacePropertyEntry(
                shownName, spelled, writable, kind == VarEnum.VT_BOOL, options, swatch);
        }
        catch (Exception)
        {
            // Some values refuse to be read in some host states. The property still exists, and a
            // row that says so beats a property that silently vanishes.
            return new SurfacePropertyEntry(shownName, "(unavailable)", false, false);
        }
    }

    /// <summary>Follows the explorer's selection with the properties panel, and nothing else.</summary>
    private void OnComponentSelected(string component)
    {
        _propertiesTarget = component;
        _propertiesControl = null;
        PublishProperties();
    }

    /// <summary>
    /// Writes a property the developer edited in the panel, in the type the property currently
    /// holds. A refusal is reported in the editor's own words; a rename is adopted everywhere the
    /// old name was a key.
    /// </summary>
    private void OnPropertyEdit(string component, string name, string value)
    {
        // The panel aimed at a CONTROL: the edit goes through the designer service - the
        // same write an api `set` makes - and the open designer tab re-projects, liveness
        // and all. The component echo names the control while one is targeted.
        if (_propertiesControl is { Length: > 0 } editedControl
            && string.Equals(component, editedControl, StringComparison.OrdinalIgnoreCase)
            && _propertiesTarget is { Length: > 0 } owningForm)
        {
            OnControlPropertyEdit(owningForm, editedControl, name, value);
            return;
        }

        try
        {
            using var found = FindComponent(component, null, out var foundIn);
            if (found is null)
            {
                Log.Info($"properties: {component} no longer exists");
                PublishProperties();
                return;
            }

            // "(Name)" is the code name, which is the component's own rather than one of the
            // collection's. For a document component the collection's Name is the host object's.
            if (string.Equals(name, "(Name)", StringComparison.Ordinal))
            {
                found.SetString("Name", value);

                var actual = found.GetString("Name") ?? value;
                Log.Info($"properties: {component} code name = '{actual}'");

                if (!string.Equals(actual, component, StringComparison.OrdinalIgnoreCase))
                {
                    AdoptRename(component, actual);
                }
                else
                {
                    PublishProperties();
                }

                return;
            }

            // A PICTURE is a file rather than a value, so it goes through the designer service -
            // the one place that knows a path becomes an IPictureDisp, and the same call the
            // control rows and the api route make.
            if (FormDesignService.IsPictureSlot(name))
            {
                using var designer = found.GetObject("Designer");
                if (designer is null)
                {
                    _editorSurface?.Notify($"{component} has no designer to write a picture to");
                    return;
                }

                var held = FormDesignService.SetControlProperty(found, designer, null, name, value, null);
                FormDesignService.KeepDesignerDown(found);
                Log.Info($"properties: {component}.{name} is {held}");
                RefreshDesignerTabFor(component);
                PublishProperties();
                return;
            }

            using var properties = found.GetObject("Properties");
            using var property = properties?.GetItem(name);
            if (property is null)
            {
                Log.Info($"properties: {component}.{name} no longer exists");
                PublishProperties();
                return;
            }

            // The panel shows what the developer WRITES - fmCycleAllForms, xlSheetVeryHidden,
            // &H8000000F& - and the model stores a number, so the name goes back through the type
            // library on the way in. Anything the library does not name passes through untouched,
            // which is how a developer who types the number still gets the number.
            using var typed = TypeSourceOf(found, component, foundIn);
            PropertyTypes.Shape? shape = null;
            if (typed is not null)
            {
                _propertyTypes.Of(typed).TryGetValue(name, out shape);
            }

            var stored = PropertyTypes.Unspell(shape, value);

            if (!WriteProperty(property, stored, out var complaint))
            {
                _editorSurface?.Notify($"{name}: {complaint}");
                PublishProperties();
                return;
            }

            Log.Info($"properties: {component}.{name} = '{value}'"
                + (stored == value ? string.Empty : $" ({stored})"));

            // Renaming changes the key everything else holds: the write baseline, the breakpoint
            // record, the tabs, the explorer, and the name the surface files the document under.
            var renamed = string.Equals(name, "Name", StringComparison.OrdinalIgnoreCase)
                ? found.GetString("Name")
                : null;

            if (renamed is not null && !string.Equals(renamed, component, StringComparison.OrdinalIgnoreCase))
            {
                // A designer tab is keyed by the module's name; AdoptRename re-keys it with
                // everything else, so the tab follows the rename instead of closing.
                AdoptRename(component, renamed);
            }
            else
            {
                // The Properties panel joins the liveness funnel: a form-level edit made
                // there reaches an open designer tab the way the api routes' edits do.
                RefreshDesignerTabFor(component);
                PublishProperties();
            }
        }
        catch (Exception ex)
        {
            // The message is the editor's own thanks to the exception information the dispatch
            // layer captures, and it is the answer to why the edit was refused.
            Log.Error($"properties: {component}.{name} could not be set", ex);
            _editorSurface?.Notify($"{name}: {ex.Message}");
            PublishProperties();
        }
    }

    /// <summary>
    /// Browse, on a picture row: the machine's own file dialog, and then the ordinary property
    /// edit with the path it answered.
    ///
    /// The dialog is the HOST's because a page cannot raise one that hands back a path - the same
    /// reason the sync dialog's folder button asks this side. It is modal to the editor's own
    /// window, so it cannot end up behind what the developer is looking at, and closing it
    /// without choosing writes nothing at all.
    /// </summary>
    private void OnPicturePick(string component, string property)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        surface.RunOnHostThread(() =>
        {
            var chosen = FilePicker.Choose(
                CodePaneTracker.MainWindow(),
                $"Choose a picture for {property}",
                // The kinds OleLoadPicturePath reads, which is the set the native designer's own
                // dialog offers. Listed apart as well as together, because a developer looking
                // for one icon among four hundred bitmaps wants the list narrowed.
                [
                    ("Pictures", "*.bmp;*.dib;*.gif;*.jpg;*.jpeg;*.png;*.ico;*.cur;*.wmf;*.emf"),
                    ("Bitmaps", "*.bmp;*.dib"),
                    ("Icons", "*.ico;*.cur"),
                    ("Metafiles", "*.wmf;*.emf"),
                    ("All files", "*.*"),
                ],
                _lastPictureFolder);

            if (chosen is null)
            {
                return;
            }

            // Where they were last, so the next picture is chosen from the same folder rather
            // than from wherever the dialog would open on its own.
            _lastPictureFolder = Path.GetDirectoryName(chosen);
            OnPropertyEdit(component, property, chosen);
        });
    }

    /// <summary>The folder the last picture came from, so the dialog opens there again. Session
    /// state rather than a setting: it follows the work, and nothing should outlive the session.</summary>
    private string? _lastPictureFolder;

    /// <summary>
    /// Writes a value into a property in the type the property holds now. False with a complaint
    /// when the text cannot become that type; the write itself may still throw, and that is the
    /// editor refusing rather than the value failing to parse.
    /// </summary>
    private static bool WriteProperty(DispatchObject property, string value, out string complaint)
    {
        complaint = string.Empty;

        switch (property.GetVarType("Value"))
        {
            case VarEnum.VT_BOOL:
                if (bool.TryParse(value, out var flag))
                {
                    property.SetBool("Value", flag);
                    return true;
                }

                complaint = $"'{value}' is not True or False.";
                return false;

            case VarEnum.VT_I2 or VarEnum.VT_I4 or VarEnum.VT_INT:
                if (int.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var number))
                {
                    property.SetInt32("Value", number);
                    return true;
                }

                complaint = $"'{value}' is not a whole number.";
                return false;

            case VarEnum.VT_R4 or VarEnum.VT_R8:
                if (double.TryParse(value, System.Globalization.CultureInfo.InvariantCulture, out var real))
                {
                    property.SetDouble("Value", real);
                    return true;
                }

                complaint = $"'{value}' is not a number.";
                return false;

            case VarEnum.VT_BSTR or VarEnum.VT_EMPTY:
                property.SetString("Value", value);
                return true;

            default:
                complaint = "This property cannot be edited here.";
                return false;
        }
    }

    /// <summary>Moves every record keyed by a component's old name to its new one.</summary>
    private void AdoptRename(string oldName, string newName)
    {
        // EVERY RENAME PATH FUNNELS THROUGH HERE - the strip's, the panel's (Name) row, the
        // component route and the undo - which is why the change log is told here and nowhere
        // else. A rename moves the module's entry rather than starting another, so a module
        // renamed and then written reads as one module that changed its name.
        RecordChange(
            newName, _shownProject, Core.Changes.ChangeKind.Renamed, null, null, from: oldName);

        // The baseline key carries the workbook; a rename reaches here for the selected
        // component, whose workbook is the shown one. The bare-name fallback key migrates
        // too, for a baseline recorded before the workbook could be told.
        var display = DisplayFromProjectId(_shownProject);
        foreach (var owner in new[] { display, null })
        {
            if (_writtenModules.Remove(WrittenKey(oldName, owner), out var baseline))
            {
                _writtenModules[WrittenKey(newName, owner)] = baseline;
            }
        }

        // The same workbook the baselines used: a rename reaches here for the selected
        // component, and the selected component belongs to the shown workbook.
        if (_breakpoints.Remove(WrittenKey(oldName, display), out var moving))
        {
            _breakpoints[WrittenKey(newName, display)] = moving with { Module = newName };
        }

        if (string.Equals(_propertiesTarget, oldName, StringComparison.OrdinalIgnoreCase))
        {
            _propertiesTarget = newName;
        }

        // A designer tab is keyed by the module's name, and it FOLLOWS the rename the way the
        // native designer window would: re-keyed in place rather than closed, never left
        // standing as a corpse under a name that no longer answers. Every rename path - the
        // strip's, the panel's (Name) row, the component route and the undo - funnels through
        // here (proven live 2026-08-13), so this is the one place the tab is told.
        for (var index = 0; index < _designerTabs.Count; index++)
        {
            if (string.Equals(_designerTabs[index].Module, oldName, StringComparison.OrdinalIgnoreCase))
            {
                _designerTabs[index] = (newName, _designerTabs[index].ProjectId);
            }
        }

        if (_activeDesignerTab is { } shownTab
            && string.Equals(shownTab.Module, oldName, StringComparison.OrdinalIgnoreCase))
        {
            _activeDesignerTab = (newName, shownTab.ProjectId);
        }

        // The renamed module's code pane did not MOVE; its name changed. Left alone, the next
        // publish would read the new name as a native move and take the active slot away from
        // a designer tab that holds it.
        if (string.Equals(_lastNativeActive.Module, oldName, StringComparison.OrdinalIgnoreCase))
        {
            _lastNativeActive = (newName, _lastNativeActive.Project);
        }

        Log.Info($"properties: {oldName} renamed to {newName}");

        // Only a rename of the shown module reloads the editor; renaming anything else must not
        // take the developer away from what they were editing. The analyzer re-runs either way,
        // because its findings carry the old name until it does.
        if (string.Equals(_editorSurface?.Module, oldName, StringComparison.OrdinalIgnoreCase))
        {
            ShowModuleInSurface(newName);
        }
        else
        {
            PublishProperties();
        }

        ComponentsChanged();
    }

    /// <summary>
    /// The workbook's components changed: everything that has to follow, in one place.
    ///
    /// THE TAB STRIP AND THE EXPLORER TREE ARE SEPARATE PUBLISHES, and the analyzer is a third
    /// thing again. Nothing refreshes on its own, so every site that adds, removes or renames a
    /// component has to say so - and each one remembering a different subset is exactly how this
    /// went wrong twice in one day: inserting a component from the page refreshed neither the
    /// strip nor the tree, and the first version of the `component` route refreshed the strip
    /// only, leaving the explorer listing three components while the strip showed eight.
    ///
    /// So there is one method, and the question "did you remember to refresh the tree" has one
    /// answer (the developer, 2026-08-07: any touch to the modules in a workbook must refresh the
    /// tree and the surface, and any relevant panes).
    /// </summary>
    private void ComponentsChanged()
    {
        PublishModules();
        PublishProjects();

        // The Properties panel is one of those relevant panes, and it was the one left out: it
        // holds the name of whatever it was last aimed at, so a removal took a component out of
        // the tree and left the panel describing it (reported 2026-08-15). After the two publishes
        // above, because the first prunes the surface's documents and this reads what they left.
        PublishProperties();

        _analysis?.Reanalyse();
    }

    /// <summary>
    /// Works out whether execution is stopped, and marks the line it is stopped on.
    ///
    /// The project reports its own mode, which is the only reading of this that is neither
    /// localised nor inferred. The first attempt used whether the reset command was available, and
    /// that is enabled in design mode as well, so the marker appeared before anything had run.
    ///
    /// The line comes from the editor's own caret, which it moves onto the statement it stopped at.
    /// There is no property for the current statement; this is the only thing that reports it.
    /// </summary>
    /// <summary>What the Locals panel last got, so unchanged readings send nothing.</summary>
    private string? _lastLocalsKey;

    /// <summary>
    /// Reads the mirrored Locals window and forwards what changed to the panel.
    ///
    /// Runs on every debug-state read: entering a break, every step (each is a command, and
    /// commands start the watch), and the polls in between. The reading is diffed as a whole,
    /// so sitting in a break costs a read and no message.
    /// </summary>
    private void PublishLocals(bool stopped)
    {
        if (_editorSurface is null)
        {
            return;
        }

        if (!stopped)
        {
            if (_lastLocalsKey is not null)
            {
                _lastLocalsKey = null;
                _editorSurface.ShowLocals(stopped: false, null, []);
            }

            return;
        }

        // Asks for a fresh read and shows the latest one that has landed; the reading thread
        // answers within a tick or two. Until it does - or while reads are failing - an empty
        // snapshot stands in, so the panel honestly shows a stopped emptiness rather than
        // whatever it said before the break.
        _ghostReaders?.RequestRead();
        var snapshot = _ghostReaders?.Locals ?? new LocalsReader.LocalsSnapshot(null, []);

        var rows = new SurfaceLocalRow[snapshot.Rows.Count];
        for (var i = 0; i < rows.Length; i++)
        {
            var row = snapshot.Rows[i];
            rows[i] = new SurfaceLocalRow(row.Expression, row.Value, row.Type);
        }

        var key = $"{snapshot.Context}\u0001{string.Join('\u0001', rows.Select(r => $"{r.Expression}={r.Value}:{r.Kind}"))}";
        if (key == _lastLocalsKey)
        {
            return;
        }

        _lastLocalsKey = key;
        _editorSurface.ShowLocals(stopped: true, snapshot.Context, rows);
        Log.Info($"locals: {rows.Length} row(s) at {snapshot.Context ?? "(no context)"}");
    }

    /// <summary>The debug mode the page last heard, so an unchanged mode sends nothing.</summary>
    private string? _lastPublishedMode;

    /// <summary>
    /// Tells the page which debug mode the editor is in - "design", "run", or "break" - so
    /// controls that only mean something stopped (the Call Stack button) can grey honestly
    /// instead of clicking into silence.
    /// </summary>
    private void PublishDebugMode(int mode)
    {
        var name = mode == BreakMode ? "break" : mode == DesignMode ? "design" : "run";
        if (name == _lastPublishedMode)
        {
            return;
        }

        _lastPublishedMode = name;
        _editorSurface?.ShowDebugMode(name);
    }

    /// <summary>What the Watch panel was last sent, so an unchanged reading sends nothing.</summary>
    private string? _lastWatchesKey;

    /// <summary>
    /// Reads the mirrored Watches window and forwards what changed to the Watch panel, on the
    /// same cadence as the Locals mirror above.
    /// </summary>
    private void PublishWatches(bool stopped)
    {
        if (_editorSurface is null)
        {
            return;
        }

        if (!stopped)
        {
            if (_lastWatchesKey is not null)
            {
                _lastWatchesKey = null;
                _editorSurface.ShowWatches(stopped: false, []);
            }

            return;
        }

        // The same stand-in discipline as the Locals mirror above: an empty reading until the
        // reading thread's first answer lands.
        _ghostReaders?.RequestRead();
        var reading = _ghostReaders?.Watches ?? [];

        var rows = new SurfaceWatchRow[reading.Count];
        for (var i = 0; i < rows.Length; i++)
        {
            var row = reading[i];
            rows[i] = new SurfaceWatchRow(row.Expression, row.Value, row.Type, row.Context);
        }

        var key = string.Join("|", rows.Select(r => $"{r.Expression}={r.Value}:{r.Kind}@{r.Context}"));
        if (key == _lastWatchesKey)
        {
            return;
        }

        _lastWatchesKey = key;
        _editorSurface.ShowWatches(stopped: true, rows);
        Log.Info($"watch: {rows.Length} row(s)");
    }

    private void UpdateDebugState()
    {
        try
        {
            using var project = _editor.GetObject("ActiveVBProject");
            var mode = project?.GetInt32("Mode") ?? DesignMode;

            // Collapses to one line per state in the development log.
            Log.Verbose($"debug: mode {mode} ({(mode == BreakMode ? "break" : mode == DesignMode ? "design" : "run")})");

            PublishDebugMode(mode);

            // A designer window a Run was aimed at goes down when the run is OVER - design
            // mode again after the run was seen - or at the deadline when the run never took
            // hold. Not at the Run itself (the posted action still needs the aim), and not
            // when the run STARTS (hiding the designer between run-start and the form window
            // appearing killed the launching form; natively the designer stands behind it).
            if (_designerRunStanding is { } standing)
            {
                if (mode != DesignMode)
                {
                    if (!standing.SawRun)
                    {
                        _designerRunStanding = (standing.Module, standing.ProjectId, standing.Deadline, true);
                    }
                }
                else if (standing.SawRun || Environment.TickCount64 > standing.Deadline)
                {
                    _designerRunStanding = null;
                    FormDesignService.StandingForRun = null;
                    try
                    {
                        using var ranForm = FindComponent(standing.Module, standing.ProjectId, out _);
                        if (ranForm is not null)
                        {
                            // The service directly, not the guarded put-down: this IS the end of
                            // the exception, and the field it guards on was cleared a line ago.
                            FormDesignService.KeepDesignerDown(ranForm);
                            Log.Info($"run form: {standing.Module}'s designer put back down");
                        }
                    }
                    catch
                    {
                        // A component gone mid-run has no window left to put down.
                    }
                }
            }

            // THE TITLE BAR SAYS THE MODE, because it is the one piece of state a developer needs
            // at a glance from another window: whether the thing they alt-tabbed away from is
            // sitting at a breakpoint. Design is said too, rather than left blank: a mode you only
            // see sometimes is one you cannot rely on reading, and a window saying nothing is
            // indistinguishable from one that has stopped reporting.
            // Remembered rather than applied from here. The mode is known at this moment and the
            // workbook is cheap to read while the project object is already in hand; the MODULE
            // changes at a different moment entirely, so the caption is composed by
            // RefreshWindowTitle, which the tab strip's own publish also calls.
            //
            // The workbook as the SHELL spells it, not as the project id does. The id is a
            // lowercased path, so composing from it put "debugfixture.xlsm" on the title bar of a
            // file called DebugFixture.xlsm.
            _titleMode = mode == BreakMode ? "break" : mode == DesignMode ? "design" : "running";
            RefreshWindowTitle();

            // The read answered, so the busy episode, if there was one, is over.
            _debugReadFailureLogged = false;

            // What the Design Mode press actually did, read once Office has refreshed the toggle.
            if (_designModePresses > 0 && --_designModePresses == 0)
            {
                var nowInDesignMode = VbeCommands.HostIsInDesignMode(_editor);

                /*
                 * A PRESS THAT MOVED NOTHING IS THE INTERESTING ONE.
                 *
                 * Excel holds design mode down for a workbook whose macros are DISABLED, and it
                 * will not let go of it: measured on a one-module workbook opened with macros
                 * off, the toggle stays -1 and every Reset control stays greyed however many
                 * times the button is pressed, while the button answers "executed" each time.
                 * Nothing the product offers clears that, which is the whole of issue #9's
                 * "only restarting Excel works" - and the way out is not this button at all, it
                 * is opening the workbook again with macros enabled.
                 */
                _designModeStuck = nowInDesignMode && _designModeBefore == nowInDesignMode;

                Log.Info($"design mode: the host is now {(nowInDesignMode ? "IN" : "out of")} it"
                    + (_designModeStuck ? " and would not leave when asked" : string.Empty)
                    + (nowInDesignMode && !_designModeStuck && _designModePressedWhileStopped
                        ? ", and the break it was in has gone with it"
                        : string.Empty));

                _editorSurface?.Notify(_designModeStuck
                    ? "Excel would not leave design mode. That usually means this workbook's "
                        + "macros are disabled - nothing will run and Reset stays greyed until it "
                        + "is closed and reopened with macros enabled."
                    : nowInDesignMode
                        ? (_designModePressedWhileStopped
                            ? "Excel is now in design mode, which ended the run you had stopped. "
                                + "Nothing will run, and Reset stays greyed, until you press "
                                + "Design Mode again."
                            : "Excel is now in design mode. Nothing will run, and Reset, Break "
                                + "and Step Out stay greyed, until you press Design Mode again.")
                        : "Excel has left design mode. Code will run again.");
            }

            if (mode != BreakMode)
            {
                if (_inBreak)
                {
                    _inBreak = false;
                    _editorSurface?.ShowCurrentLine(null);
                    _lastStopFollowed = null;

                    // The project is running again, however it got there - a developer pressing
                    // Reset, the code finishing, anything. Whatever refused the last recovery is
                    // about a session that is over, so the next one may ask again.
                    _immediateResetAsked = false;
                    _immediateResetRefused = false;

                    // Forgotten at exit so the NEXT break starts empty: the readings outlive
                    // the break otherwise, and the previous break's variables are exactly
                    // stale enough to mislead.
                    _ghostReaders?.ClearReadings();

                    PublishLocals(stopped: false);
                    PublishWatches(stopped: false);
                    Log.Info($"debug: mode {mode}, not stopped");
                }

                return;
            }

            PublishLocals(stopped: true);
            PublishWatches(stopped: true);

            using var pane = _editor.GetObject("ActiveCodePane");
            if (pane is null)
            {
                return;
            }

            Span<int> selection = stackalloc int[4];
            pane.InvokeInt32s("GetSelection", selection);

            var line = selection[0];
            if (line < 1)
            {
                return;
            }

            using var module = pane.GetObject("CodeModule");
            using var component = module?.GetObject("Parent");
            var name = component?.GetString("Name");

            // Stopped inside the evaluator's scratch frame is not somewhere the developer can be
            // taken: the module is ours and about to vanish. The stop is real, so the marker
            // rules below still run against whatever module is showing; only the switch is
            // refused.
            if (name is not null && IsScratchComponent(name))
            {
                return;
            }

            // THE STOPPED MODULE IS NAMED WITH ITS WORKBOOK, and the marker is only drawn once
            // the surface is showing it.
            //
            // A bare name resolves shown-project-first, so with two workbooks each holding a
            // Shapes a stop in one of them switched the surface to the OTHER one's - and the
            // marker below went on regardless, at a line number belonging to a module that never
            // ran. That is the reading in issue #9's photograph: a current-statement marker on an
            // `Enum` member, a line no debugger can stop at, in a module the developer had not
            // launched. A marker on the wrong line is worse than none, because it is the one
            // thing on screen that says where execution is.
            string? stoppedIn = null;
            try
            {
                using var collection = component?.GetObject("Collection");
                using var owner = collection?.GetObject("Parent");
                stoppedIn = owner is null ? null : ProjectReader.Identity(owner).Id;
            }
            catch (Exception)
            {
                // The workbook simply stays unnamed, and the checks below take the safer branch.
            }

            if (name is not null && (name != _editorSurface?.Module || stoppedIn != _shownProject))
            {
                ShowModuleInSurface(name, stoppedIn);
            }

            // Nothing to point at when the stop could not be placed - the surface keeps whatever
            // it was showing, with no marker claiming to be the current statement. The rest of
            // the tick still runs: the recovery below is about the SESSION, not about what is on
            // screen, and skipping it here is how a break that clears itself stopped clearing.
            var placed = name is not null && name == _editorSurface?.Module;

            /*
             * AND A LINE THAT CANNOT CARRY A STATEMENT CANNOT BE THE CURRENT ONE.
             *
             * VBA compiles the whole project before it runs anything, so one module that does not
             * parse stops every procedure in every other module - and the editor answers by
             * stopping on the line the COMPILER complained about. For a declarations-section
             * error that is `Option Explicit`, an `Enum` member or a `Type` field: measured at
             * line 1 of a module whose only fault was a duplicate Enum member, with the active
             * project still reading design mode. Nothing is executing there and nothing ever
             * could, so a current-statement marker on it is the same false claim the breakpoint
             * margin used to make on those lines, and the rule that stopped that one answers this.
             *
             * Judged once per stop. The rule needs the WHOLE module - whether a line sits inside
             * a Type or an Enum is not something the line can say - and this runs on every poll.
             */
            var stopKey = $"{name}:{line}";
            if (placed && stopKey != _markerJudgedFor)
            {
                _markerJudgedFor = stopKey;
                _markerCarries = Core.Editor.Breakpoints.CanCarry(SurfaceLines(), line);
            }

            if (!placed || !_markerCarries)
            {
                if (!_debugStopUnplaced)
                {
                    _debugStopUnplaced = true;
                    Log.Info($"debug: stopped at line {line} of "
                        + $"{name ?? "a module that would not name itself"}, "
                        + (placed
                            ? "which is not an executable statement - the editor stops on the line "
                                + "a compile error names, and nothing is running there"
                            : "which is not what the surface is showing")
                        + " - no current-line marker was drawn");
                }

                _editorSurface?.ShowCurrentLine(null);
            }
            else
            {
                _debugStopUnplaced = false;
                _editorSurface?.ShowCurrentLine(line);

                // THE CARET FOLLOWS THE STOP, once per stop. The host put its own caret on the
                // stopped statement - the native pane reads line 18 while stopped there - and the
                // native editor, like every reference debugger, moves the caret to each stop. The
                // surface only scrolled, so its caret (and the status bar reading it) stayed
                // wherever the developer last clicked: a bar saying "Ln 9" over a stop at 18 is
                // misdirection about where a Step acts, found by the first suite that asked the
                // bar (2026-08-12). Once per stop and not per tick, so the developer's mid-break
                // clicks are not fought; a step to a NEW line is a new stop and follows again.
                if (stopKey != _lastStopFollowed)
                {
                    _lastStopFollowed = stopKey;
                    _editorSurface?.SetCaret(line, Math.Max(1, selection[1]));
                }
                else
                {
                    _editorSurface?.Reveal(line);
                }
            }

            if (!_inBreak)
            {
                Log.Info($"debug: stopped at {name}({line})");

                /*
                 * A BREAK THIS PRODUCT CAUSED IS CLEARED HERE, which is what closes issue #7.
                 *
                 * Evaluating anything in a project that will not compile drops it out of design
                 * mode a moment after the evaluation returns. Everything then refuses: writes,
                 * component changes, later evaluations - and the fix for the broken module IS a
                 * write, so the developer could not get out without knowing to reset first.
                 *
                 * Cleared from THE POLL and not from the routes, after two attempts that were not
                 * sound. Resetting on the way out of the evaluation raced the evaluation still
                 * running and turned "Compile error: Expected: identifier" into a bare COM error
                 * seventeen seconds late. Resetting inside the module write route exceeded that
                 * route's own three-second host crossing and turned a clear refusal into a
                 * timeout. This tick is on the host thread on an ORDINARY frame, inside nobody's
                 * budget and racing nothing - which is what both of those wanted and neither had.
                 *
                 * ONLY OUR BREAK. `_immediateLeftItStopped` is set when one of our evaluations
                 * did not RUN and the editor complained, so nothing of the developer's ever
                 * started and there is no session to lose. A developer at their own breakpoint
                 * never matches, and their stop is theirs.
                 */
                if (_immediateLeftItStopped)
                {
                    _immediateLeftItStopped = false;
                    Log.Info("debug: this stop was our own compile error, clearing it");
                    ClearOurOwnBreak();
                }

            }

            _inBreak = true;

            /*
             * A RUN STOPPED BY A COMPILE ERROR IS OURS, and this is the only frame that can
             * clear it.
             *
             * A project that will not compile answers every test with the editor's compile box;
             * dismissing it leaves VBA STOPPED, so the `Application.Run` that started the test
             * NEVER RETURNS - and the run's own finally, the flag it clears and the rows it
             * would settle all sit beneath that suspended frame. The pane was left with a row
             * reading "running" for ever and every later run refused as already in flight, for
             * the rest of the session (#10). The evaluator learned this in the same words.
             *
             * THE DISCRIMINATOR IS THE DIALOG, NOT THE MODULE. The first attempt asked whether
             * the stop was inside the generated runner, on the reasoning that nobody can be
             * looking at a module that exists only during a run. It never once matched: a
             * compile error stops at the line the COMPILER named, which is the developer's own
             * module, not the frame that is suspended. What separates this from a developer's
             * breakpoint inside a test - which must never be cleared, and which `tests?
             * action=debug` exists to sit in - is that the guard has just taken a "Compile
             * error" box off the screen.
             *
             * Asked once per run, for issue #6's reason: a reset that was refused will be
             * refused again, and asking every tick is what tore the editor down.
             */
            if (_testsRunning && !_testsBreakCleared && GuardClearedACompileError())
            {
                _testsBreakCleared = true;
                Log.Warn("tests: the run is stopped behind a compile error, which it cannot "
                    + "return from - clearing it so the run can finish and say why");
                ClearOurOwnBreak();
            }
        }
        catch (Exception ex)
        {
            // Refused while the editor runs the developer's code, and this polls at 150ms, so it
            // is said once per episode rather than seven times a second for the whole run.
            if (!_debugReadFailureLogged)
            {
                _debugReadFailureLogged = true;
                Log.Info($"debug: the execution state could not be read while the editor is busy ({ex.Message})");
            }
        }
    }

    /// <summary>Whether the current episode of unreadable execution state has been reported.</summary>
    private bool _debugReadFailureLogged;

    /// <summary>
    /// Whether this run's stop inside its own scaffolding has already been cleared once.
    ///
    /// Asked once per run, for the reason the evaluator's own recovery is: a reset that was
    /// refused will be refused again, and asking every tick is what turned one stopped project
    /// into a loop that issued Reset, raised its confirmation and tore down panes until the
    /// editor faulted in VBE7.DLL (issue #6).
    /// </summary>
    private bool _testsBreakCleared;

    /// <summary>
    /// Polls remaining before the Design Mode toggle is read back.
    ///
    /// Office refreshes command bar state on its own schedule, so the toggle still reads its old
    /// value on the line after the press. Counted down on the poll, which runs on the host thread
    /// on an ordinary frame and is where every other after-the-fact reading in here happens.
    /// </summary>
    private int _designModePresses;

    /// <summary>Whether that press caught the project stopped, which is a run it threw away.</summary>
    private bool _designModePressedWhileStopped;

    /// <summary>The toggle as it stood before the press, so a press that moved nothing is visible.</summary>
    private bool _designModeBefore;

    /// <summary>
    /// Whether Excel is holding design mode down and will not let go.
    ///
    /// Two states read as design mode and want opposite advice: one the developer pressed, which
    /// the toggle undoes, and one Excel is enforcing because the workbook's macros are disabled,
    /// which it will not. Telling the second to press Design Mode is advice that cannot work.
    /// </summary>
    private bool _designModeStuck;

    /// <summary>
    /// Whether the current stop is one the surface cannot point at, and has already said so.
    ///
    /// A stop the surface is not showing gets NO marker. Drawing one anyway put a current-line
    /// marker on an `Enum` member of a module that had never run (issue #9's photograph), because
    /// the module was named without its workbook and the switch landed on a same-named module in
    /// the other one.
    /// </summary>
    private bool _debugStopUnplaced;

    /// <summary>The stop whose line was last judged executable, so the whole module is split once.</summary>
    private string? _markerJudgedFor;

    /// <summary>Whether that line can carry a statement at all - see Core.Editor.Breakpoints.</summary>
    private bool _markerCarries;

    /// <summary>Whether the developer is looking at the Immediate panel.</summary>
    private bool _watchingImmediate;

    /// <summary>
    /// Tracks each panel's own visibility transitions. With two docks the page can show more
    /// than one panel at once, so a message about one panel says nothing about the others -
    /// and only the Immediate mirror costs anything to watch.
    /// </summary>
    private void OnPanelChanged(string name, bool open)
    {
        if (name == "immediate")
        {
            _watchingImmediate = open;
            UpdatePolling();
        }
    }

    /// <summary>Starts watching the execution state, for a while.</summary>
    private void WatchDebugState()
    {
        // Twenty seconds of watching. Long enough for a procedure that does some work before it
        // reaches a breakpoint, short enough that a run which never stops does not poll all day.
        _pollsRemaining = (int)(20_000 / DebugPollMilliseconds);
        UpdatePolling();
        UpdateDebugState();
    }

    /// <summary>
    /// Sets the poll rate from what is actually being watched, or stops polling.
    ///
    /// Two things want a timer and they want it at different rates. Stepping moves the stopped line
    /// on every keystroke and has to keep up; the Immediate window only has to look live. Neither
    /// runs when nothing is watching, so a host sitting idle is not polled at all.
    /// </summary>
    private void UpdatePolling()
    {
        /*
         * A CLOSE IS WATCHED CLOSELY, BRIEFLY.
         *
         * Closing a tab felt slow and all of the wait was a tick. The window is posted a close,
         * the editor tears it down in about five milliseconds, and the strip then sat until the
         * next poll re-derived the pane picture: the window died at .430 and the tab left at
         * .581, a gap of 151ms against a 150ms poll (measured 2026-08-08).
         *
         * The destroy events cannot be used to shortcut it. They name the pane's CHILD windows,
         * not the pane, so there is nothing in them to match a tracked pane against - which was
         * the first attempt, and it never fired once.
         *
         * So the resync polls that a close already asks for run FAST. It is a handful of ticks
         * over about a tenth of a second, only after a close, and then the interval goes back.
         * Nothing else polls harder and nothing polls harder for longer.
         */
        // THE EMPTY-WORKSPACE TIER NEEDS A WINDOW TO BE EMPTY IN.
        //
        // `_watchingEmpty` says the editor has no panes, which stays true after its window closes:
        // the frame-hide path hides the surface and returns, and the two places that clear the
        // flag both need a pane to do it. So opening Alt+F11 on a workbook with nothing open, then
        // closing the editor and going back to the grid, left the whole project-and-component walk
        // running once a second on Excel's own UI thread for the rest of the session, publishing a
        // tree to a page nobody can see.
        //
        // Gated here rather than by clearing the flag, because the flag is a fact about the
        // WORKSPACE and this is a fact about the WINDOW. Read from the record placement keeps
        // rather than asked of the window: this runs from several places and the answer only
        // changes when placement notices it change, which is where UpdatePolling is called from.
        // THE IDLE TIER WATCHES THE WORKSPACE, and it exists because the tier above it only
        // covers an EMPTY editor. A workbook opened or closed in the host produces no event in
        // here, and with a module open the poll stopped altogether - so the project-count check
        // in PollDebugState, which was written for exactly this in 2026-08-08, never ran while
        // the developer had a module open and was not stepping. Closing a workbook whose
        // modules were never opened left it in the tree, in both panes' file lists, and its
        // tests in the Tests pane, until something unrelated happened (the owner, 2026-08-20).
        //
        // Cheap on purpose: this tick reads one property - the project count - where the empty
        // tier walks every project's components. It needs the frame VISIBLE, so an editor
        // window that has gone away is not polled about a tree nobody can see.
        var interval = _resyncPanePolls > 0 ? ClosingPollMilliseconds
            : _pollsRemaining > 0 ? DebugPollMilliseconds
            : _watchingImmediate ? ImmediatePollMilliseconds
            : _watchingEmpty && _frameVisible ? EmptyWorkspacePollMilliseconds
            : _frameVisible ? WorkspaceWatchPollMilliseconds
            : 0;

        _editorSurface?.Poll(interval);
        PerfCounters.Poll(interval);
    }

    /// <summary>
    /// How often the project tree is refreshed while the editor has no panes. With no panes
    /// there are no window events, and the explorer is the only way the first module gets
    /// opened, so it cannot be allowed to sit stale.
    /// </summary>
    /// <summary>
    /// How often the pane picture is re-derived in the moments after a close, until the resync
    /// polls a close asks for are spent. Short because the thing being waited for is a window
    /// teardown that takes about five milliseconds, and the developer is watching a tab they
    /// just asked to go.
    /// </summary>
    private const uint ClosingPollMilliseconds = 16;

    private const uint EmptyWorkspacePollMilliseconds = 1000;

    /// <summary>
    /// How often an otherwise idle editor asks whether the set of open files has changed. One
    /// property read per tick, against the collection the tree and both panes' file lists are
    /// built from; a changed count is the only thing that provokes any further work.
    /// </summary>
    private const uint WorkspaceWatchPollMilliseconds = 1000;

    /// <summary>
    /// Checks that the surface still agrees with the module, and adopts the module when it does
    /// not.
    ///
    /// The module is the source of truth. It can change without the surface having asked: a macro
    /// can rewrite it, an import can replace it, and the editor itself rewrites parts of it. When
    /// that happens the surface is showing something that no longer exists, and every position it
    /// reports is against the wrong text.
    ///
    /// An edit the developer has not finished is never overwritten. Their work outranks a
    /// difference that has not been reconciled yet, and the write that is already scheduled will
    /// reconcile it a moment later.
    /// </summary>
    private void ResyncFromModule()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        // Every live document is compared, not just the active one: a macro can rewrite a
        // module whose tab sits in a background group, and that model must not drift. Project
        // identities resolve once per workbook rather than once per document - the resolve
        // walks the project collection, and this runs on every pane follow.
        Dictionary<string, string?>? identities = null;
        var adopted = false;

        foreach (var (module, project) in surface.OpenDocuments)
        {
            // An edit the developer has not finished is never overwritten. Per document: a
            // sibling tab's typing does not freeze this one's resync.
            if (surface.HasUnwritten(module, project))
            {
                continue;
            }

            try
            {
                var displayKey = (project ?? string.Empty).ToLowerInvariant();
                identities ??= new Dictionary<string, string?>(StringComparer.Ordinal);
                if (!identities.TryGetValue(displayKey, out var projectId))
                {
                    projectId = project is null ? _shownProject : ProjectIdFromDisplay(project);
                    identities[displayKey] = projectId;
                }

                using var found = FindComponent(module, projectId, out _);
                var stored = found is null ? null : ProjectReader.ReadSource(found);
                if (stored is null)
                {
                    continue;
                }

                // Against what the module said last time, not against the surface. The two
                // differ by the editor's own reformatting from the moment anything is written,
                // and that difference is not a change anybody made.
                var key = WrittenKey(module, project);
                if (_writtenModules.TryGetValue(key, out var baseline) && baseline == stored)
                {
                    continue;
                }

                Log.Info($"resync: {module} changed outside the surface, adopting the module");
                _writtenModules[key] = stored;
                surface.Sync(module, project, stored);
                adopted = true;
            }
            catch (Exception ex)
            {
                Log.Error($"resync: {module} could not be compared with the module", ex);
            }
        }

        if (adopted)
        {
            _analysis?.Reanalyse();
        }
    }

    /// <summary>
    /// How many projects the editor holds, or -1 when it will not say. Deliberately just the
    /// count: this runs on every tick, and anything heavier would be paid for constantly to
    /// answer a question whose answer is almost always "the same as last time".
    /// </summary>
    private int ProjectCount()
    {
        try
        {
            using var projects = _editor.GetObject("VBProjects");
            return projects?.GetInt32("Count") ?? -1;
        }
        catch (Exception)
        {
            // The editor refuses while it is busy. The next tick asks again.
            return -1;
        }
    }

    /// <summary>The project count the tree was last built for. See PollDebugState.</summary>
    private int _lastProjectCount = -1;

    /// <summary>One tick of the execution watch.</summary>
    private void PollDebugState()
    {
        // Stamped per TICK, not per configuration change: this is the pulse the xlide api
        // reports, and it should mean "the host thread ran our periodic work just now".
        // Note the honest limit - an idle editor stops polling altogether, so an old
        // heartbeat means blocked only while something should be watching (see doctor).
        PerfCounters.Beat();

        /*
         * THE IDLE TIER IS A WORKSPACE WATCH, NOT A DEBUG POLL, and everything below this is
         * the debug poll.
         *
         * The tick that only exists to notice a file opening or closing was landing in the
         * whole of PollDebugState: PublishModules, whose own comment says the pane walk is
         * several invokes per open pane plus a Workbooks find per workbook EVERY TICK, and
         * UpdateDebugState after it. Measured on an editor with nothing happening: 241 KB/s of
         * allocation on Excel's own UI thread, against 4 KB/s with the editor window shut -
         * a second-by-second cost this tier never meant to buy, and one the empty-workspace
         * tier above was always careful about (2026-08-20).
         *
         * So the idle tier asks its one question and leaves. The heartbeat is stamped first,
         * because the doctor reads it to mean "the host thread ran our periodic work".
         */
        if (_resyncPanePolls == 0 && _pollsRemaining == 0 && !_watchingImmediate && !_watchingEmpty)
        {
            WatchOpenProjects();
            return;
        }

        // While the editor has no panes, the tree is the only living thing on screen and the
        // only route to a first module, so it follows the project as it grows.
        if (_watchingEmpty)
        {
            PublishProjects();
            AdoptOpenModuleIfEmpty();
        }
        else
        {
            /*
             * A WORKBOOK THAT APPEARS WHILE A MODULE IS OPEN USED TO NEVER APPEAR.
             *
             * The republish above is gated on the editor having no panes, so the tree followed
             * the project set only until the first module was opened. After that a workbook
             * created or opened alongside was absent from the explorer entirely: no row, no
             * modules, no route to its code, and nothing on screen to say it existed. Adding two
             * workbooks and finding the tree still showing one is where this was seen
             * (2026-08-08).
             *
             * The pane tracker cannot cover it either. It watches code pane WINDOWS, and a
             * workbook nobody has opened a module in has none.
             *
             * So the count is checked instead, which is one property read per tick against the
             * collection the tree is built from. A change in it is the only thing that provokes
             * the rebuild, so the ordinary case costs a single call and publishes nothing.
             */
            WatchOpenProjects();
        }

        _immediateReader?.Poll();

        /*
         * THE TITLE BAR IS NOT RE-ASSERTED HERE, and that is deliberate.
         *
         * The caption is taken over at start-up and retaken on every rename of the frame window.
         * On 2026-08-07 it was seen back to "Microsoft Visual Basic for Applications -
         * DebugFixture.xlsm" mid-session, and a re-apply was briefly added to this poll to cover
         * it. It was removed again for two reasons, both worth keeping written down.
         *
         * It did not work: this poll STOPS WHEN THE EDITOR IS IDLE, by design, so a caption
         * overwritten while nothing else is happening stays overwritten. Proved by overwriting it
         * by hand on the very window the chrome owns, and finding the editor's own name still
         * there three seconds later. And a poll is the wrong shape for this regardless: the
         * caption changes at a known moment and belongs on that moment's event, not on a tick
         * that has to keep asking a question whose answer is almost always no.
         *
         * WHERE A NEXT ATTEMPT SHOULD START, since one afternoon was already lost to a wrong
         * turn here: it is NOT established whether the name-change event arrives. The tracker
         * handles renames and then returns BEFORE the line that logs window events, so the log
         * contains no rename lines whether or not any were delivered. Reading that emptiness as
         * "the event never comes" is the mistake to avoid; instrument the handler itself.
         */

        // A stale pane picture is retried here, because the editor that refused it also stopped
        // producing trustworthy window events. Success fires the tracker's Changed, which is what
        // catches the tab strip up with everything that happened while the editor was busy.
        //
        // A close ALSO republishes the module list outright, because the tracker cannot see
        // this change: it only ever holds the pane windows it can match - the active one, in
        // practice - so closing a HIDDEN pane leaves its picture identical, Changed never
        // fires, and the strip kept showing the closed module's tab ("the tab X doesn't close
        // it if it's not focused", 2026-08-04, the real mechanism at last). The strip's truth
        // is the object model's open list; after a close it is re-read and re-sent, and the
        // page skips the rebuild when nothing actually changed.
        if (_resyncPanePolls > 0)
        {
            _resyncPanePolls--;
            _codePanes?.Refresh();

            // The fast interval belongs to these polls and to nothing else, so it is given back
            // the moment the last one is spent rather than on the next unrelated event.
            if (_resyncPanePolls == 0)
            {
                UpdatePolling();
            }
        }
        else if (_codePanes is { Stale: true })
        {
            _codePanes.Refresh();
        }

        // Every tick, not only the resync ones: a save made on the host's side of the fence -
        // Excel's own Ctrl+S, an autosave - flips the dirty flags with no event we hear, and
        // the dots must follow. The change-key inside makes an unchanged strip send NO MESSAGE,
        // but the work before the key is not "a read": the pane walk is several invokes per
        // open pane and the dirty flags cost a Workbooks find per distinct workbook, every
        // tick, because the key contains those flags. perf().publishUs is the measured cost.
        PublishModules();

        UpdateDebugState();

        // Watching continues for as long as execution is stopped, because the developer is about
        // to step and every step moves the marker.
        if (_inBreak)
        {
            _pollsRemaining = (int)(20_000 / DebugPollMilliseconds);
            return;
        }

        if (--_pollsRemaining <= 0)
        {
            _pollsRemaining = 0;
            UpdatePolling();
        }
    }

    /// <summary>
    /// The one question the idle tier exists to ask: has the set of open files changed? One
    /// property read against the collection the tree and both panes' file lists are built
    /// from, and a change in it is the only thing that provokes any further work.
    ///
    /// A WORKBOOK THAT APPEARS WHILE A MODULE IS OPEN USED TO NEVER APPEAR. The republish for
    /// an empty editor is gated on having no panes, so the tree followed the project set only
    /// until the first module was opened; after that a workbook opened alongside was absent
    /// from the explorer entirely - no row, no modules, no route to its code (2026-08-08). The
    /// pane tracker cannot cover it either: it watches code pane WINDOWS, and a workbook
    /// nobody has opened a module in has none.
    /// </summary>
    private void WatchOpenProjects()
    {
        var projectCount = ProjectCount();
        if (projectCount >= 0 && projectCount != _lastProjectCount)
        {
            Log.Info($"explorer: the project count went {_lastProjectCount} to {projectCount}, republishing");
            _lastProjectCount = projectCount;
            PublishProjects();
        }
    }

    /// <summary>
    /// Starts reading the Immediate window, its handle already worked out - by caption match
    /// normally, by the visibility diff as the fallback. The window keeps its handle once
    /// hidden, which is what makes it readable afterwards.
    /// </summary>
    private void AttachImmediateReader(nint window)
    {
        if (window == 0)
        {
            return;
        }

        _immediateReader = ImmediateReader.Create(window);

        if (_immediateReader is null)
        {
            Log.Info("immediate: Debug.Print output cannot be read on this host");
            return;
        }

        // Whatever it already holds is from before this session and is not news.
        _immediateReader.Reset();
        _immediateReader.Appended = OnDebugOutput;

        Log.Info($"immediate: reading Debug.Print from window 0x{window:X}");
    }

    /// <summary>
    /// Shows what Debug.Print wrote.
    ///
    /// Split into lines rather than shown as one block, because the panel is a log and each line
    /// is one thing the developer's code said.
    /// </summary>
    private void OnDebugOutput(string text)
    {
        foreach (var line in text.Split('\n'))
        {
            var trimmed = line.TrimEnd('\r');
            if (trimmed.Length > 0)
            {
                // Logged as well as shown, because whether capture is working at all is the
                // question a support log has to be able to answer.
                Log.Info($"immediate: captured '{(trimmed.Length > 80 ? trimmed[..80] : trimmed)}'");
                _editorSurface?.ShowImmediateResult(trimmed, failed: false);
            }
        }
    }

    /// <summary>
    /// Runs a line the developer entered in the Immediate panel.
    ///
    /// Their edits go to the module first. Evaluating compiles the project, so a line that refers
    /// to something just typed has to be able to see it.
    /// </summary>
    /// <summary>
    /// The addresses the page is allowed to have opened. Three, exactly, spelled here.
    ///
    /// This is a list rather than a check on the scheme or the host because the page asks for the
    /// opening and the page is the part of this product most exposed to whatever it renders. An
    /// allowed HOST would let anything under it through; an allowed ADDRESS lets through only what
    /// is written on this line, so the worst a page that has been talked into asking can do is ask
    /// for one of the developer's own sponsorship pages.
    /// </summary>
    private static readonly string[] OpenableAddresses =
    [
        "https://github.com/sponsors/WilliamSmithEdward",
        "https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD",
        "https://cash.app/$williamesmithjcil",
    ];

    /// <summary>
    /// Opens an address in whatever browser the machine uses, if it is one of the three above.
    ///
    /// The surface is a page that is not allowed to navigate anywhere, so a link on it would look
    /// like a link and do nothing. The host has no such restriction, and this is the one thing it
    /// opens.
    /// </summary>
    private void OpenExternal(string url)
    {
        if (!Array.Exists(OpenableAddresses, allowed => string.Equals(allowed, url, StringComparison.Ordinal)))
        {
            Log.Info($"external: refused an address that is not one of ours ({url.Length} characters)");
            return;
        }

        var result = Win32.ShellExecute(0, "open", url, null, null, Win32.ShowNormal);
        if ((long)result <= Win32.ShellExecuteFailure)
        {
            Log.Error($"external: the shell would not open the address (result {(long)result})");
            _editorSurface?.Notify("That address could not be opened. It is in the About dialog to copy.");
            return;
        }

        Log.Info("external: opened a sponsorship address");
    }

    /*
     * The undeclared-variable filter that stood here (2026-08-13, one day) is gone on purpose:
     * the analyzer's worker carries host-supplied designer members now (xlide_vscode#18), so a
     * control resolves where the finding is MADE rather than being swallowed after. Worth
     * remembering why the filter could not stay: it also swallowed a genuinely undeclared name
     * that happened to match a control, which the analyzer now reports truthfully.
     */

    /// <summary>
    /// The designer tabs the editor has open, in opening order, keyed by PROJECT ID rather
    /// than display name so a workbook whose display changes (unsaved numbering, save-as)
    /// keeps its tab. This list is product state, not a mirror: the native editor holds no
    /// designer window for these - KeepDesignerDown is the standing contract, because a live
    /// designer window summons the Toolbox and a save while one exists restores it on open.
    /// The pane half of the published strip mirrors the host; this half IS the truth, which
    /// is the one designated exception to the mirror (docs/userform-designer.md).
    /// </summary>
    private readonly List<(string Module, string ProjectId)> _designerTabs = [];

    /// <summary>The designer tab holding the active slot, when one is.</summary>
    private (string Module, string ProjectId)? _activeDesignerTab;

    /// <summary>
    /// The native (module, project) pair the last publish saw. When it moves, the developer
    /// acted in the native editor or a code activation landed, and either way the code pane
    /// takes the active slot back from any designer tab that held it.
    /// </summary>
    private (string? Module, string? Project) _lastNativeActive;

    /// <summary>
    /// Opens (or re-activates) a form's designer tab. Validates against the object model
    /// first: a name that is not a component, or a component that is not a form, is refused
    /// out loud rather than opening a tab whose every markup request would apologise.
    /// </summary>
    private void OpenDesignerTab(string moduleName, string? projectDisplay)
    {
        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;

        try
        {
            using var component = FindComponent(moduleName, projectId, out var owner);
            if (component is null)
            {
                _editorSurface?.Notify($"no component named {moduleName}");
                return;
            }

            if (component.GetInt32("Type") != 3)
            {
                _editorSurface?.Notify($"{moduleName} is not a UserForm, so it has no designer");
                return;
            }

            // The component's own casing, so the strip and every later request agree with the
            // tree rather than with however the click spelled it.
            var cased = component.GetString("Name") ?? moduleName;
            var ownerId = owner ?? projectId ?? string.Empty;

            if (!_designerTabs.Any(tab => string.Equals(tab.Module, cased, StringComparison.OrdinalIgnoreCase)
                && string.Equals(tab.ProjectId, ownerId, StringComparison.OrdinalIgnoreCase)))
            {
                _designerTabs.Add((cased, ownerId));
                Log.Info($"designer tab: opened {cased}");
            }

            _activeDesignerTab = (cased, ownerId);

            // The native designer's click selects the form in the Properties window; the
            // designer tab does the same (queued 2026-08-13, when a probe showed the panel
            // holding a prior component over an active designer tab).
            _propertiesTarget = cased;
            _propertiesControl = null;

            PublishModules();
            PublishProperties();
        }
        catch (Exception ex)
        {
            Log.Info($"designer tab: {moduleName} would not open ({ex.GetType().Name})");
            _editorSurface?.Notify($"{moduleName}'s designer could not be read");
        }
    }

    /// <summary>
    /// An open designer tab follows every xlide-side change to its form: the designer routes
    /// funnel through here after a successful mutation and the tab re-projects, which is what
    /// keeps the markup and the visual current without anyone re-activating the tab. What comes
    /// from OUTSIDE the funnel is caught by the fingerprint below instead - and lands here too,
    /// so this is also where the ANALYSIS learns the form moved.
    /// </summary>
    private void RefreshDesignerTabFor(string moduleName)
    {
        // The engine's seed carries the form's control list, so a designer mutation is an
        // analysis change whether or not any tab is open: without this poke a removed control
        // kept resolving as a ghost - completion, diagnostics and paint all serving the stale
        // list - until an unrelated module write happened to run a pass (the 2026-08-19 hunt).
        // Cheap when nothing material moved: the pass's sameness gate compares the walked
        // control list (SeedOf), so a geometry-only drag costs one comparison, not an analysis.
        _analysis?.Reanalyse();

        var tab = _designerTabs.FirstOrDefault(t =>
            string.Equals(t.Module, moduleName, StringComparison.OrdinalIgnoreCase));
        if (tab != default)
        {
            // The tab now shows whatever the form holds, so the liveness key is stale by
            // definition. Forgotten rather than recomputed: the next check re-seeds it silently,
            // where a stale one would report this product's own edit as somebody else's.
            _designerFingerprints.Remove(moduleName);
            PublishFormMarkup(tab.Module, DisplayFromProjectId(tab.ProjectId));

            // AND THE PANEL IS THE OTHER HALF OF THE SAME PICTURE. A control removed through the
            // api left the panel describing it - name, geometry, font and all - because only the
            // TAB was re-projected (found in the 2026-08-16 hunt). The panel's own publish
            // already falls back to the form when the control it is aimed at is gone, so this
            // needs to ask it rather than to duplicate the check.
            if (_propertiesControl is { Length: > 0 }
                && string.Equals(_propertiesTarget, moduleName, StringComparison.OrdinalIgnoreCase))
            {
                PublishProperties();
            }
        }
    }

    /// <summary>
    /// The last cheap key each open designer tab's form answered, so an edit made OUTSIDE this
    /// product's funnel can be noticed. Keyed by module name, the same key the tabs are.
    /// </summary>
    private readonly Dictionary<string, string> _designerFingerprints = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>When the liveness check last ran, and the shortest gap between two of them.</summary>
    private long _lastDesignerCheck;

    private const int DesignerCheckFloorMilliseconds = 500;

    /// <summary>
    /// Asks whether a form changed behind the product's back, and re-projects the tab if it did.
    ///
    /// THE ONE EDIT THE FUNNEL CANNOT SEE. Every designer mutation xlide makes re-projects its
    /// tab; a change made in the native designer underneath goes round all of it, and no event in
    /// the object model announces one - a form has no revision counter the way a code pane does.
    /// So it is ASKED, and the question has to be cheap enough and rare enough to be free.
    ///
    /// RARE: this runs on a window appearing or going, which is what the pane tracker already
    /// refreshes on, and NOT on a move, which is the event a resize drag fires thousands of. A
    /// native designer session is bracketed by exactly those events - the window is shown, edited,
    /// and hidden or destroyed - so the closing bracket is where an outside edit is caught. What
    /// is deliberately NOT covered: an edit made while the xlide tab is on screen, which cannot
    /// happen, because the surface covers the frame the native designer would draw in.
    ///
    /// CHEAP: the key is each control's name and bounds, nothing else - see FingerprintOf - so a
    /// check is about half a projection and only pays for a re-projection when it differs. With
    /// no designer tab open it costs one dictionary lookup and returns.
    /// </summary>
    private void CheckDesignerTabsForOutsideEdits() => CheckDesignerTabsForOutsideEdits(false);

    /// <summary>
    /// The check, with the floor optionally stood down - which is what an api caller wants: a
    /// probe asking "is the tab current" must not be told "checked half a second ago".
    /// Answers the modules it found changed.
    /// </summary>
    internal List<string> CheckDesignerTabsForOutsideEdits(bool now)
    {
        var stale = new List<string>();
        if (_designerTabs.Count == 0 || _stopped)
        {
            return stale;
        }

        // A FLOOR BETWEEN CHECKS, because these events arrive in bursts: a tooltip appearing and
        // dying is two of them, and a menu opening is several. Half a second means a burst costs
        // one check, and it costs nothing that matters to the thing being watched - a native
        // designer session lasts seconds at least, and its closing bracket is what catches it.
        var at = Environment.TickCount64;
        if (!now && at - _lastDesignerCheck < DesignerCheckFloorMilliseconds)
        {
            return stale;
        }

        _lastDesignerCheck = at;

        foreach (var tab in _designerTabs.ToArray())
        {
            try
            {
                using var component = FindComponent(tab.Module, tab.ProjectId, out _);
                using var designer = component?.GetObject("Designer");
                if (component is null || designer is null)
                {
                    continue;
                }

                var shape = FormDesignService.FingerprintOf(designer);
                FormDesignService.KeepDesignerDown(component);
                if (shape is null)
                {
                    // A form that would not answer is not a form that changed. Treating a failed
                    // read as a difference would re-project for ever on a form mid-teardown.
                    continue;
                }

                if (_designerFingerprints.TryGetValue(tab.Module, out var before) && before == shape)
                {
                    continue;
                }

                _designerFingerprints[tab.Module] = shape;
                if (before is not null)
                {
                    Log.Info($"designer: {tab.Module} changed outside xlide, re-projecting its tab");
                    stale.Add(tab.Module);

                    // The key is put back AFTER the refresh, which forgets it: this check already
                    // knows the answer, and re-seeding it here saves the next check a walk.
                    RefreshDesignerTabFor(tab.Module);
                    _designerFingerprints[tab.Module] = shape;
                }
            }
            catch (Exception why)
            {
                Log.Verbose($"designer: {tab.Module} would not answer a liveness check ({why.Message.Trim()})");
            }
        }

        return stale;
    }

    /// <summary>
    /// The canvas's double-click: the control's default event handler in the code-behind,
    /// written when it is not there and shown either way - the native designer's own
    /// gesture. The form's own handlers answer to "UserForm" whatever the form is called,
    /// which is VBA's convention rather than ours.
    /// </summary>
    /// <summary>
    /// Bring to Front / Send to Back from the canvas: MSForms' own ZOrder, on the model.
    ///
    /// The one canvas gesture that does NOT write the document, because the dialect cannot say
    /// what it does: the Controls collection a projection walks is not in z-order and does not
    /// move when ZOrder is called (measured 2026-08-16, proved on the running form by two
    /// overlapping labels swapping which was on top). So this is a property-panel-shaped edit -
    /// straight at the model, effective at once, no Ctrl+S in between - and the canvas cannot
    /// draw the result, which is recorded rather than papered over.
    /// </summary>
    private void OnDesignerZOrder(string moduleName, string? projectDisplay, string controlName, bool toFront)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;
        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out _);
                using var designer = component?.GetObject("Designer");
                if (component is null || designer is null)
                {
                    surface.Notify($"{moduleName} has no designer to reorder");
                    return;
                }

                if (!FormDesignService.ZOrderControl(designer, controlName, toFront))
                {
                    surface.Notify($"no control named {controlName} on {moduleName}");
                    return;
                }

                FormDesignService.KeepDesignerDown(component);
                surface.Notify($"{controlName} moved to the {(toFront ? "front" : "back")}");
                Log.Info($"designer: {controlName} to the {(toFront ? "front" : "back")} on {moduleName}");
            }
            catch (Exception why)
            {
                surface.Notify($"{controlName} could not be reordered ({why.Message.Trim()})");
            }
        });
    }

    /// <summary>
    /// One property of one control, written from a designer surface - today the tab-order dialog,
    /// whose Move Up is a TabIndex write and nothing else. It goes through the same
    /// SetControlProperty the Properties panel and the api route use, so a reorder cannot be a
    /// third way of writing a control's property, and the open tab re-projects afterwards.
    /// </summary>
    private void OnDesignerSetProperty(
        string moduleName, string? projectDisplay, string controlName, string property, string value)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;
        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out _);
                using var designer = component?.GetObject("Designer");
                if (component is null || designer is null)
                {
                    surface.Notify($"{moduleName} has no designer to write to");
                    return;
                }

                var shown = FormDesignService.SetControlProperty(
                    component, designer, controlName, property, value, null);
                FormDesignService.KeepDesignerDown(component);
                Log.Info($"designer: {controlName}.{property} is {shown} on {moduleName}");
                RefreshDesignerTabFor(moduleName);
            }
            catch (Exception why)
            {
                surface.Notify($"{controlName}.{property} refused the write ({why.Message.Trim()})");
            }
        });
    }

    private void OnDesignerEventStub(string moduleName, string? projectDisplay, string? controlName)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;

        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out _);
                if (component is null)
                {
                    surface.Notify($"no component named {moduleName}");
                    return;
                }

                var owner = "UserForm";
                var defaultEvent = "Click";
                if (controlName is { Length: > 0 })
                {
                    using var designer = component.GetObject("Designer");
                    using var control = designer is null
                        ? null
                        : FormDesignService.FindControlNamed(designer, controlName, 0);
                    if (control is null)
                    {
                        surface.Notify($"no control named {controlName} on {moduleName}");
                        return;
                    }

                    owner = controlName;
                    defaultEvent = FormDesignService.DefaultEventFor(FormDesignService.FriendlyTypeOf(control));
                    FormDesignService.KeepDesignerDown(component);
                }

                var source = ProjectReader.ReadSource(component) ?? string.Empty;
                var header = $"Private Sub {owner}_{defaultEvent}(";
                var lines = source.Replace("\r\n", "\n").Split('\n');
                var at = Array.FindIndex(lines, line =>
                    line.TrimStart().StartsWith(header, StringComparison.OrdinalIgnoreCase));

                if (at >= 0)
                {
                    // Standing already: the gesture navigates, never duplicates.
                    ShowModuleInSurface(moduleName, projectId);
                    surface.SetCaret(at + 2, 1);
                    return;
                }

                var bare = source.TrimEnd('\r', '\n');
                var stub = $"Private Sub {owner}_{defaultEvent}()\r\n\r\nEnd Sub";
                var text = bare.Length == 0 ? $"{stub}\r\n" : $"{bare}\r\n\r\n{stub}\r\n";
                if (WriteModule(moduleName, text, projectId, hostRewrite: true) is { } refused)
                {
                    surface.Notify($"the event stub could not be written: {refused}");
                    return;
                }

                Log.Info($"designer: event stub {owner}_{defaultEvent} on {moduleName}");
                ShowModuleInSurface(moduleName, projectId);
                surface.SetCaret(text.Replace("\r\n", "\n").TrimEnd('\n').Split('\n').Length - 1, 1);
            }
            catch (Exception ex)
            {
                Log.Error($"designer: the event stub on {moduleName} failed", ex);
                surface.Notify("the event stub could not be written");
            }
        });
    }

    /// <summary>Closes a designer tab, by name; closing what is not open changes nothing.</summary>
    private void CloseDesignerTab(string moduleName, string? projectDisplay)
    {
        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject ?? string.Empty;

        var removed = _designerTabs.RemoveAll(tab =>
            string.Equals(tab.Module, moduleName, StringComparison.OrdinalIgnoreCase)
            && string.Equals(tab.ProjectId, projectId, StringComparison.OrdinalIgnoreCase));

        if (_activeDesignerTab is { } shown
            && string.Equals(shown.Module, moduleName, StringComparison.OrdinalIgnoreCase)
            && string.Equals(shown.ProjectId, projectId, StringComparison.OrdinalIgnoreCase))
        {
            _activeDesignerTab = null;
        }

        if (removed > 0)
        {
            Log.Info($"designer tab: closed {moduleName}");
            PublishModules();
        }
    }

    /// <summary>
    /// The markup tab's Ctrl+S: the document applied to the live form through the SAME service
    /// the api's applyMarkup route wraps, so the two paths are one operation by construction.
    /// The outcome goes back first, then a fresh projection - even when the apply stopped
    /// partway, because what LANDED is on the form and the tab must show the truth.
    /// </summary>
    private void ApplyFormMarkup(string moduleName, string? projectDisplay, string markup)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;

        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out var owner);
                if (component is null)
                {
                    surface.PublishFormMarkupApplied(moduleName, projectDisplay, false, [], [], 0,
                        $"no component named {moduleName}");
                    CloseDesignerTab(moduleName, projectDisplay);
                    return;
                }

                if (component.GetInt32("Type") != 3)
                {
                    surface.PublishFormMarkupApplied(moduleName, projectDisplay, false, [], [], 0,
                        $"{moduleName} is not a UserForm");
                    return;
                }

                var display = DisplayFromProjectId(owner ?? projectId);
                var outcome = FormDesignService.ApplyMarkup(component, moduleName, markup);
                Log.Info($"form markup: apply to {moduleName} "
                    + $"{(outcome.Ok ? "ok" : "stopped")}: +{outcome.Added.Count} -{outcome.Removed.Count} set {outcome.Set}"
                    + (outcome.Refused is null ? string.Empty : $" ({outcome.Refused})"));

                surface.PublishFormMarkupApplied(moduleName, display, outcome.Ok,
                    outcome.Added, outcome.Removed, outcome.Set, outcome.Refused);

                // The fresh projection, unless nothing was touched: a parse refusal changes
                // no state, and the republish would clobber the developer's document with
                // the pre-edit text while they are looking at the refusal it earned.
                if (!outcome.ParseFailed)
                {
                    PublishFormMarkup(moduleName, display);

                    // The funnel's other mouth: this path publishes directly rather than
                    // through RefreshDesignerTabFor, so it pokes the analysis itself - the
                    // seed carries the control list, and an apply may have changed it.
                    _analysis?.Reanalyse();
                }
            }
            catch (Exception ex)
            {
                Log.Warn($"form markup: apply to {moduleName} failed ({ex.GetType().Name})");
                surface.PublishFormMarkupApplied(moduleName, projectDisplay, false, [], [], 0,
                    "the apply failed before it could start");
            }
        });
    }

    /// <summary>
    /// The markup tab's text: the form walked into the markup layer's language, or the reason
    /// it could not be. Same conversion rule as PublishDocument below, same thread rule.
    /// </summary>
    private void PublishFormMarkup(string moduleName, string? projectDisplay)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;

        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out var owner);
                if (component is null)
                {
                    surface.PublishFormMarkup(moduleName, projectDisplay, null, $"no component named {moduleName}");
                    // A designer tab whose form is gone is a corpse; the request that found
                    // out is the moment it is collected.
                    CloseDesignerTab(moduleName, projectDisplay);
                    return;
                }

                // ONE walk feeds both halves of the designer tab: the text is Print of the
                // spec the visual renders, so they cannot describe two moments of the form.
                var spec = FormDesignService.SpecOf(component, moduleName, out var reason, _controlDefaults);

                /*
                 * A FORM MID-TEARDOWN ANSWERS "HAS NO DESIGNER", and the reason used to go
                 * straight onto the canvas and stay there (the owner, 2026-08-16: "sometimes I
                 * see an overlay that says entryform2 has no designer" - EntryForm2 being a
                 * suite's form, removed while its tab was open).
                 *
                 * The component is still in the collection for a moment after a Remove, so the
                 * corpse check above does not catch it; what the walk finds is a component whose
                 * Designer will not answer. That is indistinguishable HERE from the first-touch
                 * flake the suite has a row for, so it is asked once more before being believed:
                 * a form that is going answers the same way twice and its tab is collected, and
                 * a form that was merely slow answers properly the second time and nothing is
                 * said at all.
                 */
                if (spec is null)
                {
                    using var again = FindComponent(moduleName, projectId, out var stillOwned);
                    spec = again is null
                        ? null
                        : FormDesignService.SpecOf(again, moduleName, out reason, _controlDefaults);

                    if (spec is null)
                    {
                        /*
                         * STILL IN THE COLLECTION MEANS STILL ALIVE, and that is the difference
                         * the two asks above cannot see.
                         *
                         * Asking twice separates a slow answer from a settled one by COUNT, not by
                         * TIME: both calls land in the same instant, so a form mid-teardown - its
                         * designer torn down and about to come back, right after a Run closes -
                         * answers "no designer" both times and had its tab collected for good.
                         * That is #62, and it cost about half the suite runs on 2026-08-17.
                         *
                         * A REMOVED component leaves the collection shortly; a torn-down one does
                         * not. So the collection is what is asked, and a component that is still
                         * there keeps its tab: the next projection finds the designer back and
                         * publishes normally. A tab is only collected once the component itself
                         * has gone, which is the case the row "a removed form leaves no tab
                         * standing" is about.
                         */
                        if (again is not null)
                        {
                            Log.Info($"form markup: {moduleName} has no designer just now ({reason}), "
                                + "but the component is still there - keeping its tab");
                            surface.PublishFormMarkup(moduleName, projectDisplay, null, reason);
                            return;
                        }

                        Log.Info($"form markup: {moduleName} is gone ({reason}); collecting its tab");
                        surface.PublishFormMarkup(moduleName, projectDisplay, null, reason);
                        CloseDesignerTab(moduleName, projectDisplay);
                        return;
                    }

                    owner = stillOwned ?? owner;
                }

                var markup = Core.Forms.FormMarkup.Print(spec);
                Log.Info($"form markup: publishing {moduleName}, {(markup?.Length ?? 0)} char(s), "
                    + $"{spec?.Controls.Count ?? 0} control(s){(reason is null ? "" : $" ({reason})")}");
                surface.PublishFormMarkup(moduleName, DisplayFromProjectId(owner ?? projectId), markup, reason, spec);
            }
            catch (Exception ex)
            {
                Log.Verbose($"form markup: {moduleName} could not be published ({ex.GetType().Name})");
                surface.PublishFormMarkup(moduleName, projectDisplay, null, "the form could not be read");
            }
        });
    }

    /// <summary>
    /// Gives the page a module's text without activating it.
    ///
    /// The page holds a module's text once it has been ACTIVATED, so a workspace opened onto eight
    /// modules holds one - and anything that draws a module it is not showing (peeking a
    /// definition, previewing a reference) had nothing to draw. Answered without touching which
    /// pane is active, because being taken to what you asked to look at is the whole complaint.
    /// </summary>
    private void PublishDocument(string moduleName, string? projectDisplay)
    {
        if (_editorSurface is not { } surface)
        {
            return;
        }

        // The page names a workbook the way it is shown; the object model wants the project's own
        // identity. Every other route that takes a project from the page converts here, and this
        // one did not - so the component was never found and the answer never came (2026-08-07).
        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;

        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out var owner);
                var text = component is null ? null : ProjectReader.ReadSource(component);
                if (text is null)
                {
                    Log.Info($"document: {moduleName} could not be found to publish");
                    return;
                }

                Log.Info($"document: publishing {moduleName}, {text.Length} char(s), without activating it");
                surface.Publish(moduleName, DisplayFromProjectId(owner ?? projectId), text);
            }
            catch (Exception ex)
            {
                Log.Verbose($"document: {moduleName} could not be published ({ex.GetType().Name})");
            }
        });
    }

    /// <summary>
    /// Evaluates a line in the Immediate window and RETURNS its verdict.
    ///
    /// It used to return nothing, showing the result on the page and telling the caller only that
    /// it had run. So the debug route could report that an evaluation had been asked for and never
    /// what it came to, and the Immediate window ended up with a route nothing could assert on.
    /// </summary>
    private ImmediateEvaluator.Result EvaluateImmediate(string line)
    {
        Log.Info($"immediate: evaluate '{(line.Length > 80 ? line[..80] : line)}'");

        // An attachment that failed at start-up is retried the moment output is about to
        // matter. The window certainly exists by now - start-up toggled it visible - and the
        // caption identification does not care that it is hidden.
        if (_immediateReader is null && _immediateCaption is { Length: > 0 } caption)
        {
            AttachImmediateReader(CodePaneTracker.FindPaneByCaption(caption));
            if (_immediateReader is not null)
            {
                Log.Info("immediate: reader attached on retry");
            }
        }

        /*
         * STOPPED INSIDE OUR OWN SCRATCH PROCEDURE: clear it before doing anything else.
         *
         * A line that will not compile raises the editor's "Compile error" box, and dismissing it
         * leaves VBA STOPPED INSIDE the scratch module. `Application.Run` never returns, because
         * that call frame is suspended in the debugger, so nothing written after it can ever run:
         * the evaluator's own clean-up, its mode check, and the flag it sets for the next
         * evaluation all sit BENEATH the suspended frame. Three attempts at recovering after the
         * fact failed for that one reason, and the log said so by containing none of their output
         * (2026-08-07).
         *
         * The recovery therefore belongs HERE, on the way into a later evaluation, on a frame
         * that is not suspended. The discriminator is exact rather than a guess about modes: if
         * the editor is stopped and the pane it is stopped in is the scratch module, the session
         * being ended is one this product started and nobody else can be looking at it. A
         * developer stopped at their own breakpoint is in one of their own modules and is never
         * touched.
         */

        _editorSurface?.FlushEdits();

        var evaluator = _immediate;
        if (evaluator is null)
        {
            evaluator = _immediate = new ImmediateEvaluator(_editor);

            // "? name" in break mode is answered from the Locals ghost (#21): the row the panel
            // shows for that name, read the way the panel reads it. A fresh read is asked for so
            // a value that moved on the last step is the one printed, and the latest landed
            // reading answers - the reader is a tick or two behind at most, and the panel on
            // screen is the same reading.
            evaluator.LocalLookup = name =>
            {
                _ghostReaders?.RequestRead();
                var snapshot = _ghostReaders?.Locals;
                if (snapshot is null)
                {
                    return null;
                }

                foreach (var row in snapshot.Rows)
                {
                    if (string.Equals(row.Expression, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return (row.Value, row.Type);
                    }
                }

                return null;
            };

            // PUT THE PROJECT BACK when an evaluation leaves it stopped. A line that will not
            // compile raises the editor's own "Compile error" box, and behind it the project is
            // out of design mode: every evaluation after that answered "Not available while
            // execution is stopped", so one mistyped line made the Immediate window useless until
            // somebody thought to press Reset (measured 2026-08-07 with `?((`).
            //
            // Through the editor's own Reset command, which is what a developer would press, and
            // which this session already knows how to execute. The evaluator notices; it does not
            // reach for COM of its own to fix it.
            evaluator.StoppedUnexpectedly = () =>
            {
                // ASKED ONCE. A reset that was refused will be refused again for the same reason,
                // and asking on every evaluation is what turned one stopped project into a loop
                // that issued Reset, raised its confirmation and tore down panes until the editor
                // faulted in VBE7.DLL and took Excel with it (issue #6, found by chaos.mjs).
                // ASKED ONCE, AND JUDGED LATER.
                //
                // Whether a reset worked cannot be read the instant the command returns: the
                // editor's mode does not turn over that fast, and a first attempt at this checked
                // 9ms afterwards and called a reset refused that had simply not landed yet. The
                // evidence that it failed is being asked AGAIN - arriving here a second time
                // means a whole evaluation has passed and the project is still stopped, which is
                // time that really elapsed rather than a guess about how much would be enough.
                //
                // And nothing here sleeps, deliberately. This runs on the host thread, which is
                // the thread the confirmation's own message loop is on, so waiting for the answer
                // here would be waiting for a thread this call is standing on.
                if (_immediateResetAsked)
                {
                    if (!_immediateResetRefused)
                    {
                        _immediateResetRefused = true;
                        Log.Warn("immediate: the reset did not take - the project is still "
                            + "stopped an evaluation later. Something declined its confirmation, "
                            + "or the editor will not leave break mode. Press Reset in the editor.");
                    }

                    return;
                }

                Log.Info("immediate: the line left the project stopped, resetting");
                _immediateResetAsked = true;
                try
                {
                    // Reset asks "This action will reset your project, proceed anyway?", and
                    // whatever answers a dialog blocking the host thread would otherwise decline
                    // it - correctly, for a question nobody here asked, and fatally for this one.
                    using (Diagnostics.DialogWatch.ExpectingConfirmation())
                    {
                        ExecuteEditorCommand(VbeCommands.Command.Reset);
                    }
                }
                catch (Exception ex)
                {
                    Log.Info($"immediate: reset FAILED ({ex.GetType().Name}: {ex.Message})");
                }
            };
        }

        // The mode is read now rather than taken from the cached flag. The flag is as old as the
        // last poll, and evaluating during an unnoticed break added a module to a stopped
        // project, which fails in ways that have nothing to do with what the developer typed.
        var stopped = _inBreak;
        try
        {
            using var project = _editor.GetObject("ActiveVBProject");
            stopped = (project?.GetInt32("Mode") ?? DesignMode) != DesignMode;
        }
        catch (Exception)
        {
            // The cached answer stands.
        }

        var result = evaluator.Evaluate(line, stopped);

        Log.Info($"immediate: {(result.Failed ? "failed" : "ok")}"
                 + (result.Text.Length > 0 ? $" '{(result.Text.Length > 80 ? result.Text[..80] : result.Text)}'" : string.Empty));

        // A successful statement says nothing, which is the native window's own manner. What the
        // code printed arrives through the Debug.Print reader; the ceremony stays silent.
        if (result.Failed || result.Text.Length > 0)
        {
            _editorSurface?.ShowImmediateResult(result.Text, result.Failed);
        }

        // Evaluating adds and removes a module, which the analyzer would otherwise report on.
        _analysis?.Reanalyse();

        // Running the line took the host through pane churn that ends with a native pane active
        // and the keyboard on it. The developer is mid-conversation with the prompt.
        _editorSurface?.Focus();

        return result;
    }

    /// <summary>Answers the surface's menu bar with the items the editor holds right now.</summary>
    private void OnMenuRequested(int[] path)
    {
        try
        {
            var items = VbeMenus.Read(_editor, path);
            _editorSurface?.ShowMenu(path, items);
            Log.Info($"menu: [{string.Join(",", path)}] read, {items.Length} item(s)");
        }
        catch (Exception ex)
        {
            Log.Error($"menu: [{string.Join(",", path)}] could not be read", ex);

            // An empty menu renders as an empty menu, which at least answers the click.
            _editorSurface?.ShowMenu(path, []);
        }
    }

    /// <summary>
    /// Runs a menu item the developer chose from the surface's menu bar.
    ///
    /// Most items are executed exactly where they live, which is what keeps this menu complete: it
    /// can run anything the native menu can, dialogs included. The exceptions are the commands the
    /// session has its own path for, and the windows the surface replaces; those are routed to the
    /// replacement, because executing them natively would either skip the session's bookkeeping or
    /// put a native window on screen that the surface exists to replace.
    /// </summary>
    private void OnMenuExecuteRequested(int[] path)
    {
        try
        {
            using var control = VbeMenus.ControlAt(_editor, path);
            if (control is null)
            {
                Log.Info($"menu: [{string.Join(",", path)}] no longer exists");
                return;
            }

            var id = control.GetInt32("Id");
            if (RouteMenuCommand(id))
            {
                Log.Info($"menu: [{string.Join(",", path)}] routed as command {id}");
                return;
            }

            if (!control.GetBool("Enabled"))
            {
                // The page draws disabled items as disabled, but its picture is as old as the
                // moment the menu opened, and the editor moves on underneath it.
                _editorSurface?.Notify("That menu item is not available right now.");
                return;
            }

            // The item acts on the module and on the editor's own caret, so both are brought
            // current first: compiling, saving and exporting must see what was just typed.
            _editorSurface?.FlushEdits();
            SyncCaretToPane();

            control.Invoke("Execute");
            Log.Info($"menu: [{string.Join(",", path)}] executed ({id})");

            if (id == VbeCommands.Command.ClearAllBreakpoints)
            {
                ForgetBreakpoints();
            }

            // A menu item can start a run, and it can open or close native windows the surface
            // must make room for. Both are watched for rather than assumed.
            WatchDebugState();
            RefreshSurfacePlacement();
        }
        catch (Exception ex)
        {
            Log.Error($"menu: [{string.Join(",", path)}] could not be executed", ex);
            _editorSurface?.Notify("That menu item could not be run.");
        }
    }

    /// <summary>
    /// Runs a menu command through its surface-side owner instead of the native item, when it has
    /// one. True when the command was taken.
    /// </summary>
    private bool RouteMenuCommand(int id)
    {
        if (VbeCommands.RoutesThroughSession(id))
        {
            ExecuteEditorCommand(id);
            return true;
        }

        switch (id)
        {
            // Editing commands act on the text the developer sees, which is the surface's.
            // Executed natively they would act on the covered pane, and the native find dialog
            // is a modal over an editor nobody is looking at.
            case VbeCommands.Command.Undo:
                _editorSurface?.RunEditorCommand("undo");
                return true;

            case VbeCommands.Command.Redo:
                _editorSurface?.RunEditorCommand("redo");
                return true;

            case VbeCommands.Command.Find:
                _editorSurface?.RunEditorCommand("actions.find");
                return true;

            case VbeCommands.Command.Replace:
                _editorSurface?.RunEditorCommand("editor.action.startFindReplaceAction");
                return true;

            // Windows the surface has its own version of. The native ones were closed at start-up
            // and reopening one would put it behind the surface, which reads as nothing happening.
            case VbeCommands.Command.ImmediateWindow:
                _editorSurface?.RunEditorCommand("xlide.panel.immediate");
                return true;

            case VbeCommands.Command.LocalsWindow:
                _editorSurface?.RunEditorCommand("xlide.panel.locals");
                return true;

            case VbeCommands.Command.PropertiesWindow:
                _editorSurface?.RunEditorCommand("xlide.panel.properties");
                return true;

            case VbeCommands.Command.ProjectExplorer:
                _editorSurface?.Notify("The project explorer is part of the editor and always shown.");
                return true;

            // The native Options dialog is superseded: its Editor and Editor Format tabs
            // configure an editor nobody is looking at. Tools > Options opens the product's
            // own settings, which is where the choices that matter now live.
            case VbeCommands.Command.Options:
                _editorSurface?.RunEditorCommand("xlide.openSettings");
                return true;

            default:
                return false;
        }
    }

    /// <summary>
    /// Drops every breakpoint this add-in recorded, after the editor cleared them all. The record
    /// only mirrors the editor; when the editor forgets, remembering draws dots on lines that no
    /// longer stop anything.
    /// </summary>
    private void ForgetBreakpoints()
    {
        _breakpoints.Clear();
        _editorSurface?.ShowBreakpoints([]);
    }

    /// <summary>
    /// The one shape of an engine round trip made from a surface request handler: capture the
    /// module and its live text on the host thread, answer empty when the session cannot ask,
    /// resolve OFF the host thread under the standard three-second deadline - these ride on
    /// keystrokes, and the developer must never wait for one - log the outcome either way, and
    /// marshal the reply back, because the browser may only be spoken to from the thread that
    /// owns it. A request that fails answers empty rather than never: the editor is left
    /// waiting on nothing.
    ///
    /// Seven handlers were this block copied out with only the ask, the projection, the empty
    /// value and the log verb changing, and the shared policy had already drifted at the edges:
    /// code actions' guard skipped the source it does not use, and four of the seven logged
    /// only non-empty answers, so a feature answering empty was indistinguishable from one
    /// never asked (the audit's B14). Every request logs now, the verbose-dev rule.
    ///
    /// `ask` runs on a pool thread and answers the reply's whole value, empty included, so the
    /// helper holds no null convention of its own. `reply` always runs on the host thread, with
    /// `empty` when the ask threw. `locus` is the position part of the log line ("@120",
    /// "@3..9"), and `describe` names the answer for it. Closed generic instantiations only,
    /// so ahead-of-time compilation is unaffected.
    /// </summary>
    private void AnswerFromEngine<T>(
        string verb,
        string locus,
        T empty,
        Func<AnalysisService, string, string, CancellationToken, Task<T>> ask,
        Action<EditorSurface, T> reply,
        Func<T, string> describe)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            if (_editorSurface is { } idle)
            {
                reply(idle, empty);
            }

            return;
        }

        _ = Task.Run(async () =>
        {
            var value = empty;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                value = await ask(analysis, module, source, deadline.Token).ConfigureAwait(false);
                Log.Info($"{verb}: {module}{locus} -> {describe(value)}");
            }
            catch (Exception ex)
            {
                Log.Info($"{verb}: {module}{locus} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => reply(surface, value));
        });
    }

    /// <summary>Answers a completion request from the surface.</summary>
    private void OnCompletionRequested(int requestId, int offset) =>
        AnswerFromEngine<SurfaceCompletionItem[]>(
            "completion",
            $"@{offset}",
            [],
            async (analysis, module, source, token) =>
            {
                var answered = await analysis.CompleteAsync(module, source, offset, token)
                    .ConfigureAwait(false);
                return answered is null ? [] : [.. answered.Select(item => new SurfaceCompletionItem(
                    item.Label,
                    item.Kind,
                    item.Detail,
                    item.Documentation,
                    item.InsertText,
                    item.FilterText,
                    item.SortText))];
            },
            (surface, items) => surface.ShowCompletions(requestId, items),
            items => $"{items.Length} item(s)");

    /// <summary>Answers a hover request from the surface, the same way a completion is.</summary>
    private void OnHoverRequested(int requestId, int offset)
    {
        // An attribute annotation under the caret is this session's to explain, not the engine's:
        // what it writes, and what the module has now.
        if (_editorSurface is { Module: { } hoveredModule, Text: { } hoveredText } surface
            && AttributeHover(hoveredModule, _shownProject, hoveredText, offset) is { } ours)
        {
            surface.ShowHover(requestId, ours);
            return;
        }

        AnswerFromEngine<SurfaceHoverPayload?>(
            "hover",
            $"@{offset}",
            null,
            async (analysis, module, source, token) =>
            {
                var answered = await analysis.HoverAsync(module, source, offset, token)
                    .ConfigureAwait(false);
                return answered is null ? null : new SurfaceHoverPayload(
                    answered.Signature,
                    answered.Details,
                    answered.Documentation,
                    answered.Span.Start,
                    answered.Span.End);
            },
            (surface, payload) => surface.ShowHover(requestId, payload),
            payload => payload is null ? "nothing" : payload.Signature);
    }

    /// <summary>Answers a call-tip request from the surface, the same way a hover is answered.</summary>
    private void OnSignatureHelpRequested(int requestId, int offset) =>
        AnswerFromEngine<SurfaceSignatureInfo?>(
            "signature",
            $"@{offset}",
            null,
            async (analysis, module, source, token) =>
            {
                var answered = await analysis.SignatureHelpAsync(module, source, offset, token)
                    .ConfigureAwait(false);
                return answered is null ? null : new SurfaceSignatureInfo(
                    answered.Label,
                    [.. answered.Parameters.Select(parameter =>
                        new SurfaceSignatureParameter(parameter.Label, parameter.Documentation))],
                    answered.ActiveParameter,
                    answered.Documentation,
                    answered.Details);
            },
            (surface, payload) => surface.ShowSignatureHelp(requestId, payload),
            payload => payload is null ? "nothing" : payload.Label);

    /// <summary>
    /// Answers a Smart Enter request from the surface: what the Enter that just went in should
    /// leave behind. Answered the way a completion is: capture on the host thread, resolve off
    /// it, marshal the answer back. A request that fails answers empty rather than never.
    /// </summary>
    private void OnSmartEnterRequested(int requestId, int offset)
    {
        // Captured on the host thread, before the hop: the developer's typing choices decide
        // what Enter leaves behind, and reading them off the session from a background task
        // would race a settings change.
        var typing = _settings;

        AnswerFromEngine<(SurfaceTextEdit[] Edits, int? Caret)>(
            "smartEnter",
            $"@{offset}",
            ([], null),
            async (analysis, module, source, token) =>
            {
                var answered = await analysis.SmartEnterAsync(module, source, offset, typing, token)
                    .ConfigureAwait(false);
                return answered is null
                    ? ([], null)
                    : ([.. answered.Edits.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))],
                        answered.Caret);
            },
            (surface, value) => surface.ShowSmartEnter(requestId, value.Edits, value.Caret),
            value => $"{value.Edits.Length} edit(s)");
    }

    /// <summary>
    /// Answers a canonical-case request from the surface: the case corrections for a span,
    /// resolved from the same project facts completion uses.
    /// </summary>
    private void OnCanonicalCaseRequested(int requestId, int start, int end, bool single, bool completeHeader) =>
        AnswerFromEngine<SurfaceTextEdit[]>(
            "canonicalCase",
            $"@{start}..{end}",
            [],
            async (analysis, module, source, token) =>
            {
                var answered = await analysis
                    .CanonicalCaseAsync(module, source, start, end, single, completeHeader, token)
                    .ConfigureAwait(false);
                return answered is null
                    ? []
                    : [.. answered.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))];
            },
            (surface, edits) => surface.ShowCanonicalCase(requestId, edits),
            edits => $"{edits.Length} edit(s)");

    /// <summary>
    /// Answers a loop-sync request from the surface: the paired iterator rename, when the edit
    /// at the offset touched one side of a For/Next pair.
    /// </summary>
    private void OnLoopSyncRequested(int requestId, int offset) =>
        AnswerFromEngine<SurfaceTextEdit[]>(
            "loopSync",
            $"@{offset}",
            [],
            async (analysis, module, source, token) =>
            {
                var answered = await analysis.LoopSyncAsync(module, source, offset, token)
                    .ConfigureAwait(false);
                return answered is null
                    ? []
                    : [.. answered.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))];
            },
            (surface, edits) => surface.ShowLoopSync(requestId, edits),
            edits => $"{edits.Length} edit(s)");

    /// <summary>
    /// Answers a quick-fix request from the surface: what can be fixed over a span, and the edits
    /// that would fix it. Answered the same way a hover is, and empty on failure for the same
    /// reason - a lightbulb that does not appear is what the developer already sees when there is
    /// nothing to fix.
    /// </summary>
    private void OnCodeActionsRequested(int requestId, int start, int end)
    {
        // Read on the host thread, where the shown project is decided, for the fixes this
        // session offers itself alongside the engine's (#23's attributes).
        var shownProject = _shownProject;
        AnswerFromEngine<SurfaceCodeAction[]>(
            "codeAction",
            $"@{start}..{end}",
            [],
            async (analysis, module, source, token) =>
            {
                var answered = await analysis.CodeActionsAsync(module, start, end, token)
                    .ConfigureAwait(false);
                return [.. answered.Select(action => new SurfaceCodeAction(
                    action.Title,
                    action.IsPreferred ?? false,
                    action.Code,
                    action.Span.Start,
                    action.Span.End,
                    [.. action.Edits.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))])),
                    .. AttributeCodeActions(module, shownProject, source, start, end)];
            },
            (surface, actions) => surface.ShowCodeActions(requestId, actions),
            actions => $"{actions.Length} fix(es)");
    }

    /// <summary>
    /// The analyzer's rule catalog, fetched once per session.
    ///
    /// The catalog is a build-time table - it changes when the engine binary does and never
    /// otherwise - so the first asker pays the round trip and everyone after reads the cache.
    /// Null until the engine has answered once; a caller that gets null says "not yet", never
    /// "no rules".
    /// </summary>
    private EngineAnalysisRules? _analysisRuleCatalog;

    private async Task<EngineAnalysisRules?> AnalysisRuleCatalogAsync()
    {
        if (_analysisRuleCatalog is { } held)
        {
            return held;
        }

        if (_analysis is not { } analysis)
        {
            return null;
        }

        var fetched = await analysis.RulesAsync(CancellationToken.None).ConfigureAwait(false);
        if (fetched is not null)
        {
            _analysisRuleCatalog = fetched;
        }

        return fetched;
    }

    /// <summary>
    /// Changes one analyzer rule's severity FOR THIS MACHINE, from any entry point: the rules
    /// modal, the problems pane's menu, the lightbulb, and the xlide api all arrive here, so
    /// every one of them behaves identically and persists identically.
    ///
    /// Validated against the analyzer's own catalog before anything is written, because the
    /// engine IGNORES a disallowed override rather than failing - a caller who turned an
    /// error-severity rule "off" would see nothing change and nothing say why. The refusal
    /// happens here, in words, with the moves that ARE legal named.
    ///
    /// "default" clears the override. The change is saved to the settings file, handed to the
    /// analysis service, and a reanalysis is provoked so every open workbook's findings move.
    /// </summary>
    /// <returns>What happened, in words - for the route's reply and the page's notice.</returns>
    private async Task<string> ApplyRuleSeverityAsync(string code, string severity)
    {
        var ruleCode = (code ?? string.Empty).Trim().ToLowerInvariant();
        var wanted = (severity ?? string.Empty).Trim().ToLowerInvariant();

        var catalog = await AnalysisRuleCatalogAsync().ConfigureAwait(false);
        if (catalog is null)
        {
            return "the analysis engine is not up, so nothing was changed";
        }

        var rule = catalog.Rules.FirstOrDefault(one =>
            string.Equals(one.Code, ruleCode, StringComparison.OrdinalIgnoreCase));
        if (rule is null)
        {
            return $"'{ruleCode}' is not an analyzer rule; GET analysis lists them";
        }

        if (wanted is not ("default" or "off" or "warning" or "error" or "information"))
        {
            return $"'{wanted}' is not a severity; pass off, warning, error, information, "
                + "or default to clear the override";
        }

        if (wanted != "default" && !rule.Allowed.Contains(wanted, StringComparer.OrdinalIgnoreCase))
        {
            return rule.Allowed.Length == 0
                ? $"{rule.Code} cannot be changed: its default is {rule.DefaultSeverity} and the "
                    + "analyzer allows no override - most error rules mirror a VBE compile failure"
                : $"{rule.Code} cannot be '{wanted}'; the analyzer allows: "
                    + string.Join(", ", rule.Allowed);
        }

        var overrides = new Dictionary<string, string>(
            _settings.AnalysisRuleSeverityOverrides ?? [], StringComparer.Ordinal);
        var stood = overrides.TryGetValue(rule.Code, out var previous) ? previous : null;

        if (wanted == "default")
        {
            if (stood is null)
            {
                return $"{rule.Code} already stands at its default ({rule.DefaultSeverity})";
            }

            overrides.Remove(rule.Code);
        }
        else
        {
            if (string.Equals(stood, wanted, StringComparison.Ordinal))
            {
                return $"{rule.Code} is already {wanted} on this machine";
            }

            overrides[rule.Code] = wanted;
        }

        OnSettingsChanged(_settings with
        {
            AnalysisRuleSeverityOverrides = overrides.Count > 0 ? overrides : null,
        });

        // The findings on screen are about the OLD policy until a pass runs under the new one.
        _analysis?.Reanalyse();

        Log.Info($"analysis: {rule.Code} -> "
            + $"{(wanted == "default" ? $"default ({rule.DefaultSeverity})" : wanted)} for this machine");
        return wanted == "default"
            ? $"{rule.Code} is back to its default ({rule.DefaultSeverity}) on this machine"
            : $"{rule.Code} is {wanted} everywhere on this machine";
    }

    /// <summary>
    /// Writes an inline suppression above a finding's line, in any module of any open workbook.
    ///
    /// The directive's shape mirrors the analyzer's own suppression quick fix
    /// (diagnosticCodeActions.ts): the finding line's leading whitespace, then
    /// `' @xlide-analysis-disable-next-line <code>`. The lightbulb path applies the engine's
    /// edit; this one exists for the problems pane, whose findings live in modules that need not
    /// be open - so it goes through WriteModule, the hardened write, and inherits its refusals:
    /// a stopped project answers in words rather than raising the editor's reset dialog.
    /// </summary>
    private void OnSuppressFindingRequested(string module, string? projectDisplay, int line, string code)
    {
        if (SuppressFinding(module, projectDisplay, line, code) is { } refused)
        {
            _editorSurface?.Notify(refused);
        }
    }

    /// <summary>
    /// The mechanism behind it, shared with the analysis route so the api and the pane cannot
    /// drift. Null when the directive was written; the refusal in words otherwise.
    /// </summary>
    private string? SuppressFinding(string module, string? projectDisplay, int line, string code)
    {
        try
        {
            // The developer's unwritten keystrokes reach the module first, so the directive is
            // inserted into the text as typed rather than as last written back.
            _editorSurface?.FlushEdits();

            var projectId = ProjectIdFromDisplay(projectDisplay);
            using var found = FindComponent(module, projectId, out var owner);
            if (found is null)
            {
                return $"Nothing named {module} to suppress in.";
            }

            var source = ProjectReader.ReadSource(found);
            if (source is null)
            {
                return $"{module}'s text could not be read.";
            }

            var lines = source.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
            if (line > lines.Length)
            {
                return $"{module} has no line {line} to suppress at.";
            }

            var ruleCode = code.Trim().ToLowerInvariant();

            /*
             * THE DIRECTIVE IS CHOSEN BY THE RULE'S OWN SCOPE, read from the catalog.
             *
             * A module-scoped finding - option-explicit-missing is the owner's example - anchors
             * at (1,1), and a disable-next-line above it covers the line BELOW the directive,
             * which the re-anchored finding never sits on: the comment suppressed nothing, the
             * squiggle moved onto the directive itself, and the pane kept the row (the owner's
             * screenshot, 2026-08-27). Those rules take the FILE directive, at the very top,
             * which is the one place the analyzer accepts it. Everything line-anchored keeps the
             * next-line directive above the finding. The analyzer's own quick fix offers only
             * next-line today, module-scoped rules included - filed upstream.
             */
            var scopes = AnalysisRuleCatalogAsync().GetAwaiter().GetResult()?.Rules
                .FirstOrDefault(one => string.Equals(one.Code, ruleCode, StringComparison.OrdinalIgnoreCase))
                ?.SuppressionScopes ?? [];
            var wholeModule = scopes.Contains("module", StringComparer.OrdinalIgnoreCase)
                && !scopes.Contains("line", StringComparer.OrdinalIgnoreCase);

            string rewritten;
            if (wholeModule)
            {
                var directive = $"' @xlide-analysis-disable-file {ruleCode}";
                rewritten = string.Join("\r\n", [directive, .. lines]);
            }
            else
            {
                var at = lines[line - 1];
                var indent = at[..(at.Length - at.TrimStart().Length)];
                var directive = $"{indent}' @xlide-analysis-disable-next-line {ruleCode}";
                rewritten = string.Join("\r\n", [.. lines[..(line - 1)], directive, .. lines[(line - 1)..]]);
            }

            if (WriteModule(module, rewritten, owner ?? projectId, hostRewrite: true) is { } refused)
            {
                return refused;
            }

            Log.Info($"analysis: suppressed {ruleCode} at {module}({line}) with an inline "
                + $"{(wholeModule ? "disable-file" : "disable-next-line")} directive");
            return null;
        }
        catch (Exception ex)
        {
            Log.Error($"analysis: suppressing {code} at {module}({line}) failed", ex);
            return "The suppression could not be written; the log says why.";
        }
    }

    /// <summary>
    /// Answers the surface's request for the rule catalog and the standing overrides - the one
    /// payload the rules modal renders from.
    /// </summary>
    private void OnAnalysisRulesRequested(int requestId)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceAnalysisRule[] rules = [];
            var failed = true;
            try
            {
                var catalog = await AnalysisRuleCatalogAsync().ConfigureAwait(false);
                rules = catalog?.Rules
                    .Select(rule => new SurfaceAnalysisRule(
                        rule.Code, rule.Title, rule.Category, rule.DefaultSeverity, rule.Allowed))
                    .ToArray() ?? [];
                failed = catalog is null;
            }
            catch (Exception ex)
            {
                Log.Error("analysis: the rules request failed", ex);
            }

            // THE REPLY CROSSES ON THE BROWSER'S OWN THREAD, like every reply in this file.
            // Posted from this pool thread it never arrives at all: PostWebMessageAsString
            // answers UI_E_WRONG_THREAD, the error lands in the log, and the page's ask times
            // out into "no catalog answer" - which is how the first build of this shipped a
            // lightbulb entry and a modal that could never load (2026-08-27, and the reason
            // AnswerFromEngine ends with RunOnHostThread).
            surface.RunOnHostThread(() => surface.ShowAnalysisRules(
                requestId,
                rules,
                _settings.AnalysisRuleSeverityOverrides
                    ?? new Dictionary<string, string>(StringComparer.Ordinal),
                failed));
        });
    }

    /// <summary>
    /// Answers an outline request from the surface: a module's procedures for its tree node.
    /// The live text is captured when the request is about the module being edited; any other
    /// module answers from the engine's seeded copy, which is current as of the last write-back.
    /// </summary>
    private void OnOutlineRequested(int requestId, string moduleName, string? projectDisplay)
    {
        var surface = _editorSurface;

        if (surface is null || _analysis is not { } analysis)
        {
            // Not "this module is empty" - "nothing here can answer yet". The page keeps what
            // it already shows.
            _editorSurface?.ShowOutline(requestId, [], failed: true);
            return;
        }

        // The tree names the workbook its row belongs to, which is what makes a shared module
        // name unfold the right workbook's procedures.
        var projectId = ProjectIdFromDisplay(projectDisplay);

        // No source travels with the request: the engine's live copy is exact - didChange rides
        // the same FIFO pipe ahead of this - and serialising a 918KB module once a second to
        // tell the engine what it already holds was most of this request's cost.
        _ = Task.Run(async () =>
        {
            SurfaceOutlineProcedure[] procedures = [];
            var failed = false;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                var answered = await analysis.OutlineAsync(moduleName, projectId, source: null, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    procedures = [.. answered.Select(procedure =>
                        new SurfaceOutlineProcedure(procedure.Name, procedure.Kind, procedure.Line))];
                }
                else
                {
                    failed = true;
                }

                Log.Info($"outline: {moduleName} -> {procedures.Length} procedure(s)");
            }
            catch (Exception ex)
            {
                failed = true;
                Log.Info($"outline: {moduleName} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowOutline(requestId, procedures, failed));
        });
    }

    /// <summary>
    /// Renames a symbol everywhere it is used in the workbook, whether its module has a tab open
    /// or not (the developer, 2026-08-06).
    ///
    /// The write goes through the same writer every other module write uses, which reaches a
    /// module through the object model and so does not care whether anything is showing it. Open
    /// tabs are then given the new text, because the page's model is what the developer is
    /// looking at and it would otherwise still say the old name.
    ///
    /// All or nothing. The engine refuses rather than returning a partial answer, and a write
    /// that fails part-way is reported as what it is: a rename that stopped, naming the modules
    /// that did change. A rename that quietly reaches most of a project compiles until the module
    /// nobody renamed runs.
    /// </summary>
    private void OnRenameRequested(int requestId, int offset, string newName)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowRenamed(requestId, null, newName, [], 0, "There is nothing here to rename.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            string? oldName = null;
            string? refused = null;
            string? renamedComponent = null;
            string[] changed = [];
            var replaced = 0;
            Xlide.Vbe.Core.Engine.EngineRenamedModule[] modules = [];

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis.RenameAsync(module, offset, newName, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was renamed.";
                }
                else
                {
                    oldName = outcome.Answer.OldName;
                    refused = outcome.Answer.Refused;
                    modules = outcome.Answer.Modules;
                    replaced = modules.Sum(entry => entry.Replaced);

                    // The word was a MODULE's name, not a symbol in one. The engine says so
                    // rather than the page guessing, because only the engine knows every module
                    // of the workbook - and the page only knows the ones with a tab open. The
                    // component half is the add-in's either way, so it goes to the same route
                    // the explorer's Rename uses.
                    renamedComponent = outcome.Answer.Module;
                }
            }
            catch (Exception ex)
            {
                refused = "The rename could not be worked out, so nothing was renamed.";
                Log.Info($"rename: {module}@{offset} failed ({ex.GetType().Name})");
            }

            if (refused is not null)
            {
                surface.RunOnHostThread(() =>
                    surface.ShowRenamed(requestId, oldName, newName, [], 0, refused));
                return;
            }

            if (renamedComponent is { } component)
            {
                // A module rename with nothing else naming it still has a component to rename,
                // so it does NOT get the empty-modules early return below.
                surface.RunOnHostThread(() =>
                    ApplyModuleRename(requestId, component, _shownProject, newName, modules));
                return;
            }

            if (modules.Length == 0)
            {
                surface.RunOnHostThread(() =>
                    surface.ShowRenamed(requestId, oldName, newName, [], 0, "Nothing to rename."));
                return;
            }

            // The writes are the host's own object model, so they belong on the host thread -
            // the same thread every other module write happens on.
            surface.RunOnHostThread(() =>
            {
                // What each module holds now, so this can be put back. Read before the first
                // write, and read from the MODULES rather than reused from the engine's input,
                // because the editor rewrites what it is handed.
                _undoableRename = new RenameUndo(
                    oldName ?? string.Empty,
                    newName,
                    null,
                    _shownProject,
                    CaptureBefore(modules.Select(entry => entry.Module), _shownProject));

                var written = new List<string>(modules.Length);
                string? stopped = null;

                foreach (var entry in modules)
                {
                    try
                    {
                        // A refused write stops the rename here, exactly as a thrown one does. It
                        // used to be counted as a module renamed, which is the reading a developer
                        // can least afford in the middle of a rename: half the uses rewritten and a
                        // report saying all of them were.
                        if (WriteModule(entry.Module, entry.Source, _shownProject, hostRewrite: true) is { } refused)
                        {
                            stopped = $"'{entry.Module}' could not be written, so the rename stopped there. {refused}";
                            break;
                        }

                        written.Add(entry.Module);

                        // The page holds a model per OPEN module. Syncing one that is not open is
                        // harmless - the page has nothing by that name to sync - and syncing one
                        // that is open is the whole point.
                        surface.Sync(entry.Module, display, entry.Source);
                    }
                    catch (Exception ex)
                    {
                        stopped = $"'{entry.Module}' could not be written, so the rename stopped there.";
                        Log.Error($"rename: writing {entry.Module} failed", ex);
                        break;
                    }
                }

                Log.Info($"rename: {oldName} -> {newName}, {replaced} use(s) in {written.Count} module(s)");
                surface.ShowRenamed(requestId, oldName, newName, [.. written], replaced, stopped);

                // The findings describe the old names until something asks again.
                _analysis?.Reanalyse();
            });
        });
    }

    /// <summary>
    /// Lifts the selected lines into a Private procedure below the one they came from.
    ///
    /// ONE MODULE, ONE WRITE, and the same shape as a rename: the engine works out the whole new
    /// text and refuses anything it cannot make safely, so by the time a single character is
    /// written the transformation is already known to be possible. What is left here is the
    /// host's half - the write, the sync into the open tab, and the undo slot that makes Ctrl+Z
    /// put it back, because an edit the host made was never on the page's undo stack.
    /// </summary>
    /// <summary>
    /// Puts a property pair in front of a module variable and makes the variable private.
    ///
    /// The third of these, and the same shape as the two before it: the engine works out the whole
    /// new text and refuses what it cannot write, so the host's half is one write, one sync and
    /// the undo slot. Nothing outside the module changes, because the property keeps the
    /// variable's name.
    /// </summary>
    /// <summary>
    /// Replaces a local with what it was assigned, and takes its declaration away.
    ///
    /// Nothing is asked of the developer: what the name stands for is already in the code. The
    /// engine refuses anything it cannot do without deciding a question of VBA's - brackets around
    /// an argument change how it binds - so the host's half is the usual write, sync and undo.
    /// </summary>
    /// <summary>
    /// Moves a procedure into another module, with every qualified call site that named the old
    /// one repointed.
    ///
    /// SEVERAL MODULES, which makes this the rename's shape rather than the other refactorings':
    /// the engine works out every module's new text before a character is written, and the writes
    /// stop at the first refusal rather than carrying on - half a move is a project that does not
    /// compile, and reporting it as done is the reading a developer can least afford.
    /// </summary>
    /// <summary>
    /// Turns a local into a parameter, and gives every call site the value it used to be assigned.
    ///
    /// The other multi-module one, and the same bargain as the move: nothing is written until every
    /// module's new text is known, and the writes stop at the first refusal rather than leaving
    /// half the callers passing an argument the signature does not have.
    /// </summary>
    private void OnIntroduceParameterRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowParameterIntroduced(requestId, null, null, null, null, [], 0,
                "There is nothing open to work on.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineIntroduceParameter? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .IntroduceParameterAsync(module, offset, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing changed.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The parameter could not be worked out, so nothing changed.";
                Log.Info($"parameter: {module}@{offset} failed ({ex.GetType().Name})");
            }

            if (refused is null && (answer?.Modules is null || answer.Modules.Length == 0))
            {
                refused = "The engine did not say what the modules should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowParameterIntroduced(requestId, null, null, null, null, [], 0, why));
                return;
            }

            var made = answer!;

            surface.RunOnHostThread(() =>
            {
                var before = CaptureBefore(made.Modules!.Select(one => one.Module), _shownProject);
                var written = new List<string>(made.Modules!.Length);
                string? stopped = null;

                foreach (var entry in made.Modules!)
                {
                    try
                    {
                        if (WriteModule(entry.Module, entry.Source, _shownProject, hostRewrite: true) is { } why)
                        {
                            stopped = $"'{entry.Module}' could not be written, so it stopped there. {why}";
                            break;
                        }

                        written.Add(entry.Module);
                        surface.Sync(entry.Module, display, entry.Source);
                    }
                    catch (Exception ex)
                    {
                        stopped = $"'{entry.Module}' could not be written, so it stopped there.";
                        Log.Error($"parameter: writing {entry.Module} failed", ex);
                        break;
                    }
                }

                _undoableRename = new RenameUndo(
                    made.Parameter ?? string.Empty,
                    made.Procedure ?? module,
                    null,
                    _shownProject,
                    before);

                Log.Info($"parameter: {made.Parameter} As {made.Type} on {made.Procedure}, "
                    + $"{made.CallSites ?? 0} call site(s)");
                surface.ShowParameterIntroduced(
                    requestId, made.Parameter, made.Type, made.Value, made.Procedure,
                    [.. written], made.CallSites ?? 0, stopped);

                _analysis?.Reanalyse();
            });
        });
    }

    private void OnMoveToModuleRequested(int requestId, int offset, string targetModule)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowMoved(requestId, null, null, null, [], 0, "There is nothing open to move from.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineMoveToModule? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .MoveToModuleAsync(module, offset, targetModule, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was moved.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The move could not be worked out, so nothing changed.";
                Log.Info($"move: {module}@{offset} to {targetModule} failed ({ex.GetType().Name})");
            }

            if (refused is null && (answer?.Modules is null || answer.Modules.Length == 0))
            {
                refused = "The engine did not say what the modules should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowMoved(requestId, null, null, null, [], 0, why));
                return;
            }

            var made = answer!;

            surface.RunOnHostThread(() =>
            {
                var before = CaptureBefore(made.Modules!.Select(one => one.Module), _shownProject);
                var written = new List<string>(made.Modules!.Length);
                string? stopped = null;

                foreach (var entry in made.Modules!)
                {
                    try
                    {
                        if (WriteModule(entry.Module, entry.Source, _shownProject, hostRewrite: true) is { } why)
                        {
                            stopped = $"'{entry.Module}' could not be written, so the move stopped there. {why}";
                            break;
                        }

                        written.Add(entry.Module);
                        surface.Sync(entry.Module, display, entry.Source);
                    }
                    catch (Exception ex)
                    {
                        stopped = $"'{entry.Module}' could not be written, so the move stopped there.";
                        Log.Error($"move: writing {entry.Module} failed", ex);
                        break;
                    }
                }

                _undoableRename = new RenameUndo(
                    made.From ?? module,
                    made.To ?? targetModule,
                    null,
                    _shownProject,
                    before);

                Log.Info($"move: {made.Moved} from {made.From} to {made.To}, "
                    + $"{written.Count} module(s), {made.Requalified ?? 0} call site(s) repointed");
                surface.ShowMoved(
                    requestId, made.Moved, made.From, made.To, [.. written], made.Requalified ?? 0, stopped);

                _analysis?.Reanalyse();
            });
        });
    }

    private void OnInlineVariableRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowVariableInlined(requestId, null, null, 0, null,
                "There is nothing open to inline in.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineInlineVariable? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .InlineVariableAsync(module, offset, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was inlined.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The value could not be worked out, so nothing changed.";
                Log.Info($"inline: {module}@{offset} failed ({ex.GetType().Name})");
            }

            if (refused is null && answer?.Source is null)
            {
                refused = "The engine did not say what the module should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowVariableInlined(requestId, null, null, 0, null, why));
                return;
            }

            var made = answer!;

            surface.RunOnHostThread(() =>
            {
                var before = CaptureBefore([module], _shownProject);

                try
                {
                    if (WriteModule(module, made.Source!, _shownProject, hostRewrite: true) is { } stopped)
                    {
                        surface.ShowVariableInlined(requestId, null, null, 0, null,
                            $"'{module}' could not be written, so nothing was inlined. {stopped}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"inline: writing {module} failed", ex);
                    surface.ShowVariableInlined(requestId, null, null, 0, null,
                        $"'{module}' could not be written, so nothing was inlined.");
                    return;
                }

                _undoableRename = new RenameUndo(
                    made.Variable ?? string.Empty,
                    made.Value ?? string.Empty,
                    null,
                    _shownProject,
                    before);

                surface.Sync(module, display, made.Source!);

                Log.Info($"inline: {made.Variable} -> {made.Value} in {module}, {made.Replaced} use(s)");
                surface.ShowVariableInlined(
                    requestId, made.Variable, made.Value, made.Replaced ?? 0, module, null);

                _analysis?.Reanalyse();
            });
        });
    }

    /// <summary>
    /// Gives a selected expression a name: a declaration and its assignment above the statement it
    /// came from, and the selection replaced by the name.
    ///
    /// The fourth of these and the same shape as the three before it. What is different is where
    /// the TYPE comes from - the analyzer, through resolveExpressionType, rather than from
    /// anything decided here - and that it also answers whether the assignment needs Set, which no
    /// rule of thumb over a type name can settle.
    /// </summary>
    private void OnExtractVariableRequested(int requestId, int startOffset, int endOffset, string newName)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowVariableExtracted(requestId, null, null, false, null, null,
                "There is nothing open to extract from.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineExtractVariable? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .ExtractVariableAsync(module, startOffset, endOffset, newName, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was extracted.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The expression could not be worked out, so nothing changed.";
                Log.Info($"variable: {module} {startOffset}-{endOffset} failed ({ex.GetType().Name})");
            }

            if (refused is null && answer?.Source is null)
            {
                refused = "The engine did not say what the module should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowVariableExtracted(requestId, null, null, false, null, null, why));
                return;
            }

            var made = answer!;

            surface.RunOnHostThread(() =>
            {
                var before = CaptureBefore([module], _shownProject);

                try
                {
                    if (WriteModule(module, made.Source!, _shownProject, hostRewrite: true) is { } stopped)
                    {
                        surface.ShowVariableExtracted(requestId, null, null, false, null, null,
                            $"'{module}' could not be written, so nothing was extracted. {stopped}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"variable: writing {module} failed", ex);
                    surface.ShowVariableExtracted(requestId, null, null, false, null, null,
                        $"'{module}' could not be written, so nothing was extracted.");
                    return;
                }

                _undoableRename = new RenameUndo(
                    made.Expression ?? newName,
                    made.Variable ?? newName,
                    null,
                    _shownProject,
                    before);

                surface.Sync(module, display, made.Source!);

                Log.Info($"variable: {made.Variable} As {made.Type} = {made.Expression} in {module}");
                surface.ShowVariableExtracted(
                    requestId, made.Variable, made.Type, made.IsObject ?? false, made.Expression, module, null);

                _analysis?.Reanalyse();
            });
        });
    }

    private void OnEncapsulateFieldRequested(int requestId, string fieldName)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowEncapsulated(requestId, null, null, [], null, "There is nothing open to encapsulate in.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineEncapsulateField? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .EncapsulateFieldAsync(module, fieldName, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was changed.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The property could not be worked out, so nothing changed.";
                Log.Info($"encapsulate: {module}.{fieldName} failed ({ex.GetType().Name})");
            }

            if (refused is null && answer?.Source is null)
            {
                refused = "The engine did not say what the module should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowEncapsulated(requestId, null, null, [], null, why));
                return;
            }

            var made = answer!;

            surface.RunOnHostThread(() =>
            {
                var before = CaptureBefore([module], _shownProject);

                try
                {
                    if (WriteModule(module, made.Source!, _shownProject, hostRewrite: true) is { } stopped)
                    {
                        surface.ShowEncapsulated(requestId, null, null, [], null,
                            $"'{module}' could not be written, so nothing was encapsulated. {stopped}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"encapsulate: writing {module} failed", ex);
                    surface.ShowEncapsulated(requestId, null, null, [], null,
                        $"'{module}' could not be written, so nothing was encapsulated.");
                    return;
                }

                // The same slot the other three use, holding what the module said before an edit
                // the HOST made, which the page's undo stack never saw.
                _undoableRename = new RenameUndo(
                    made.Field ?? fieldName,
                    made.BackingField ?? fieldName,
                    null,
                    _shownProject,
                    before);

                surface.Sync(module, display, made.Source!);

                Log.Info($"encapsulate: {module}.{made.Field} is now a property over {made.BackingField}");
                surface.ShowEncapsulated(
                    requestId, made.Field, made.BackingField, made.Accessors ?? [], module, null);

                _analysis?.Reanalyse();
            });
        });
    }

    /// <summary>
    /// Writes the stubs a class owes the interfaces it declares.
    ///
    /// The same shape as an extraction, and for the same reason: the engine works out the whole
    /// new text and refuses anything it cannot write, so the host's half is one write, one sync
    /// and the undo slot. What differs is that nothing is asked of the developer first - the
    /// members, their names and their signatures are all the interface's, so there is nothing to
    /// name and no dialog to name it in.
    /// </summary>
    private void OnImplementInterfaceRequested(int requestId, string? interfaceName)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowImplemented(requestId, [], [], null, "There is nothing open to implement into.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineImplementInterface? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .ImplementInterfaceAsync(module, interfaceName, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was written.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The members could not be worked out, so nothing changed.";
                Log.Info($"implement: {module} failed ({ex.GetType().Name})");
            }

            if (refused is null && answer?.Source is null)
            {
                refused = "The engine did not say what the module should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowImplemented(requestId, [], [], null, why));
                return;
            }

            var written = answer!;

            surface.RunOnHostThread(() =>
            {
                var before = CaptureBefore([module], _shownProject);

                try
                {
                    if (WriteModule(module, written.Source!, _shownProject, hostRewrite: true) is { } stopped)
                    {
                        surface.ShowImplemented(requestId, [], [], null,
                            $"'{module}' could not be written, so nothing was implemented. {stopped}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"implement: writing {module} failed", ex);
                    surface.ShowImplemented(requestId, [], [], null,
                        $"'{module}' could not be written, so nothing was implemented.");
                    return;
                }

                var interfaces = written.Interfaces ?? [];
                var added = written.Added ?? [];

                // The same slot rename and extraction use, holding the same thing: what a module
                // said before an edit the HOST made, which the page's undo stack never saw.
                _undoableRename = new RenameUndo(
                    module,
                    interfaces.Length > 0 ? interfaces[0] : module,
                    null,
                    _shownProject,
                    before);

                surface.Sync(module, display, written.Source!);

                Log.Info($"implement: {module} gained {added.Length} member(s) of {string.Join(", ", interfaces)}");
                surface.ShowImplemented(requestId, interfaces, added, module, null);

                _analysis?.Reanalyse();
            });
        });
    }

    private void OnExtractMethodRequested(int requestId, int startLine, int endLine, string newName)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowExtracted(requestId, null, null, null, null, "There is nothing open to extract from.");
            return;
        }

        var display = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineExtractMethod? answer = null;
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .ExtractMethodAsync(module, startLine, endLine, newName, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was extracted.";
                }
                else
                {
                    answer = outcome.Answer;
                    refused = outcome.Answer.Refused;
                }
            }
            catch (Exception ex)
            {
                refused = "The extraction could not be worked out, so nothing changed.";
                Log.Info($"extract: {module} {startLine}-{endLine} failed ({ex.GetType().Name})");
            }

            if (refused is null && answer?.Source is null)
            {
                refused = "The engine did not say what the module should hold, so nothing changed.";
            }

            if (refused is not null)
            {
                var why = refused;
                surface.RunOnHostThread(() =>
                    surface.ShowExtracted(requestId, null, null, null, null, why));
                return;
            }

            var made = answer!;

            // The write is the host's own object model, so it belongs on the host thread - the
            // same thread every other module write happens on.
            surface.RunOnHostThread(() =>
            {
                // Read before the first write, and read from the MODULE rather than reused from
                // what the engine was given, because the editor rewrites what it is handed.
                var before = CaptureBefore([module], _shownProject);

                try
                {
                    if (WriteModule(module, made.Source!, _shownProject, hostRewrite: true) is { } stopped)
                    {
                        surface.ShowExtracted(requestId, null, null, null, null,
                            $"'{module}' could not be written, so nothing was extracted. {stopped}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"extract: writing {module} failed", ex);
                    surface.ShowExtracted(requestId, null, null, null, null,
                        $"'{module}' could not be written, so nothing was extracted.");
                    return;
                }

                // One slot, shared with rename, because it holds the same thing: what a module
                // said before an edit the HOST made, which the page's undo stack never saw.
                _undoableRename = new RenameUndo(
                    made.From ?? module,
                    made.Procedure ?? newName,
                    null,
                    _shownProject,
                    before);

                surface.Sync(module, display, made.Source!);

                Log.Info($"extract: {made.From} -> {made.Procedure} in {module} ({made.Signature})");
                surface.ShowExtracted(requestId, made.Procedure, made.From, made.Signature, module, null);

                // The findings describe the module as it was until something asks again.
                _analysis?.Reanalyse();
            });
        });
    }

    /// <summary>
    /// Renames a module and everything that names it, across the workbook, whether each module
    /// has a tab open or not.
    ///
    /// THE ORDER IS THE WHOLE PROBLEM. Rename the component first and every qualified call points
    /// at a module that no longer exists; rewrite the calls first and they point at one that does
    /// not exist yet. Either way the project does not compile in between, and a failure part-way
    /// leaves it there.
    ///
    /// So nothing is written until everything is known. The engine works out every module's new
    /// text and hands it back whole, having already refused a name that is taken, is not an
    /// identifier, or is a keyword. Only then is the component renamed - the one step that can
    /// still be refused by the host - and only if THAT succeeds is a single line of code written.
    /// The window where the project is inconsistent is the gap between two operations that have
    /// both already been proven possible.
    /// </summary>
    private void OnModuleRenameRequested(int requestId, string moduleName, string? projectDisplay, string newName)
    {
        var surface = _editorSurface;

        if (surface is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowRenamed(requestId, moduleName, newName, [], 0, "There is nothing here to rename.");
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay);
        var display = DisplayFromProjectId(projectId ?? _shownProject);

        _ = Task.Run(async () =>
        {
            Xlide.Vbe.Core.Engine.EngineRenamedModule[] modules = [];
            string? refused = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                var answered = await analysis
                    .RenameModuleAsync(moduleName, projectId, newName, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not { } outcome)
                {
                    refused = "The analyzer is not available, so nothing was renamed.";
                }
                else
                {
                    refused = outcome.Answer.Refused;
                    modules = outcome.Answer.Modules;
                }
            }
            catch (Exception ex)
            {
                refused = "The rename could not be worked out, so nothing was renamed.";
                Log.Info($"renameModule: {moduleName} failed ({ex.GetType().Name})");
            }

            if (refused is not null)
            {
                surface.RunOnHostThread(() =>
                    surface.ShowRenamed(requestId, moduleName, newName, [], 0, refused));
                return;
            }

            surface.RunOnHostThread(() =>
                ApplyModuleRename(requestId, moduleName, projectId ?? _shownProject, newName, modules));
        });
    }

    /// <summary>
    /// Renames the component, then writes what every module must say afterwards. The host thread
    /// owns both, and both entry points - the explorer's Rename and renaming a module's name in
    /// code - land here so there is one order rather than two.
    ///
    /// The component goes first because it is the only step the HOST can still refuse, and the
    /// only one whose failure leaves nothing changed. Its answer is read back rather than
    /// assumed: the editor normalises a name it dislikes, and rewriting every caller to a name it
    /// did not accept would break all of them at once.
    /// </summary>
    /// <summary>
    /// What one rename changed, kept so it can be put back.
    ///
    /// A rename is the one operation here that edits several modules at once, and the editor's
    /// undo stack is PER MODEL: Ctrl+Z in the module you are looking at reverses that module's
    /// share and leaves every other one renamed, which is a half-renamed project and worse than
    /// no undo at all. So the reversal is the add-in's, over the same modules the rename touched.
    ///
    /// One slot. A second rename replaces it - undo goes back one step, the way the operation is
    /// one step - and the texts are the ones read out of the modules immediately before writing,
    /// not the ones the engine was given, because the editor rewrites what it is handed.
    /// </summary>
    private sealed record RenameUndo(
        string OldName,
        string NewName,
        string? Component,
        string? ProjectId,
        (string Module, string Text)[] Before);

    private RenameUndo? _undoableRename;

    /// <summary>Reads what each module holds now, so a rename can be put back exactly.</summary>
    private (string Module, string Text)[] CaptureBefore(IEnumerable<string> modules, string? projectId)
    {
        var captured = new List<(string, string)>();

        foreach (var module in modules)
        {
            try
            {
                using var component = FindComponent(module, projectId, out _);
                var text = component is null ? null : ProjectReader.ReadSource(component);
                if (text is null)
                {
                    continue;
                }

                captured.Add((module, text));
            }
            catch (Exception ex)
            {
                Log.Verbose($"rename: {module} could not be captured for undo ({ex.GetType().Name})");
            }
        }

        return [.. captured];
    }

    /// <summary>
    /// Puts the last rename back: every module's text as it was, and the component's old name.
    ///
    /// The component goes LAST, mirroring the order the rename does it in. A rename renames the
    /// component first because that is the only step the host can still refuse; an undo rewrites
    /// the callers first, so the project is never pointing at a name that does not exist.
    /// </summary>
    /// <summary>
    /// What an undo came to, for a caller that is not the page. The page learns all of this from
    /// ShowRenamed; a script calling the same operation through the door used to be told `ran:
    /// true` whether the undo restored six modules, restored two and stopped, or found nothing to
    /// undo at all - and the message explaining which went to a page request id the door invented.
    /// </summary>
    internal readonly record struct UndoRenameOutcome(
        bool Undone,
        string? From,
        string? To,
        string[] Modules,
        string? Stopped);

    private UndoRenameOutcome UndoRename(int requestId)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return new UndoRenameOutcome(false, null, null, [], "the surface is not up");
        }

        if (_undoableRename is not { } undo)
        {
            surface.ShowRenamed(requestId, null, null, [], 0, "There is no rename to undo.");
            return new UndoRenameOutcome(false, null, null, [], "There is no rename to undo.");
        }

        var display = DisplayFromProjectId(undo.ProjectId);
        var restored = new List<string>(undo.Before.Length);
        string? stopped = null;

        foreach (var (module, text) in undo.Before)
        {
            // The module answers to its NEW name now, if it was the one renamed.
            var target = undo.Component is { } renamed
                && string.Equals(module, renamed, StringComparison.OrdinalIgnoreCase)
                ? undo.NewName
                : module;

            try
            {
                if (WriteModule(target, text, undo.ProjectId, hostRewrite: true) is { } refused)
                {
                    stopped = $"'{target}' could not be written, so the undo stopped there. {refused}";
                    break;
                }

                restored.Add(module);
                surface.Sync(target, display, text);
            }
            catch (Exception ex)
            {
                stopped = $"'{target}' could not be written, so the undo stopped there.";
                Log.Error($"undo rename: writing {target} failed", ex);
                break;
            }
        }

        if (stopped is null && undo.Component is { } component)
        {
            try
            {
                using var found = FindComponent(undo.NewName, undo.ProjectId, out _);
                found?.SetString("Name", component);
                AdoptRename(undo.NewName, component);
            }
            catch (Exception ex)
            {
                stopped = $"The editor would not rename '{undo.NewName}' back to '{component}'.";
                Log.Error($"undo rename: the component would not go back", ex);
            }
        }

        Log.Info($"undo rename: {undo.NewName} -> {undo.OldName}, {restored.Count} module(s)");

        // Spent either way. An undo that half-worked must not be offered again, because the
        // texts it holds no longer describe what is there.
        _undoableRename = null;

        ComponentsChanged();
        surface.ShowRenamed(requestId, undo.NewName, undo.OldName, [.. restored], restored.Count, stopped);

        return new UndoRenameOutcome(stopped is null, undo.NewName, undo.OldName, [.. restored], stopped);
    }

    private void ApplyModuleRename(
        int requestId,
        string moduleName,
        string? projectId,
        string newName,
        Xlide.Vbe.Core.Engine.EngineRenamedModule[] modules)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        var display = DisplayFromProjectId(projectId);

        // Captured BEFORE the component is renamed, while every module still answers to the name
        // it has now - including the one about to change.
        var before = CaptureBefore(modules.Select(entry => entry.Module), projectId);

        string actual;
        try
        {
            using var found = FindComponent(moduleName, projectId, out _);
            if (found is null)
            {
                surface.ShowRenamed(requestId, moduleName, newName, [], 0,
                    $"'{moduleName}' could not be found, so nothing was renamed.");
                return;
            }

            found.SetString("Name", newName);
            actual = found.GetString("Name") ?? newName;
        }
        catch (Exception ex)
        {
            Log.Error($"renameModule: the host refused to rename {moduleName}", ex);
            surface.ShowRenamed(requestId, moduleName, newName, [], 0,
                $"The editor would not rename '{moduleName}', so nothing was renamed.");
            return;
        }

        AdoptRename(moduleName, actual);

        _undoableRename = new RenameUndo(moduleName, actual, moduleName, projectId, before);

        var written = new List<string>(modules.Length);
        var replaced = 0;
        string? stopped = null;

        foreach (var entry in modules)
        {
            // A module that mentions the renamed one may BE the renamed one; it answers to its
            // new name now.
            var target = string.Equals(entry.Module, moduleName, StringComparison.OrdinalIgnoreCase)
                ? actual
                : entry.Module;

            try
            {
                if (WriteModule(target, entry.Source, projectId, hostRewrite: true) is { } refused)
                {
                    stopped = $"'{target}' could not be written, so the rename stopped there. {refused}";
                    break;
                }

                written.Add(target);
                replaced += entry.Replaced;
                surface.Sync(target, display, entry.Source);
            }
            catch (Exception ex)
            {
                stopped = $"'{target}' could not be written, so the rename stopped there.";
                Log.Error($"renameModule: writing {target} failed", ex);
                break;
            }
        }

        Log.Info($"renameModule: {moduleName} -> {actual}, {replaced} mention(s) in {written.Count} module(s)");
        surface.ShowRenamed(requestId, moduleName, actual, [.. written], replaced, stopped);

        PublishProjects();
        _analysis?.Reanalyse();
    }

    /// <summary>
    /// Answers a navigation request from the surface: where the symbol at the caret is declared,
    /// or everywhere in the workbook it is used.
    ///
    /// Every answer names a module of the SHOWN workbook, which is the invariant that makes one
    /// display name right for all of them: the engine resolves within one project and never
    /// across two, because two open workbooks can each hold a Module1 and a Recalculate.
    /// </summary>
    private void OnNavigationRequested(int requestId, int offset, bool references, bool includeDeclaration)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowLocations(requestId, []);
            return;
        }

        var workbook = DisplayFromProjectId(_shownProject);

        _ = Task.Run(async () =>
        {
            SurfaceLocation[] locations = [];

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                var answered = await analysis
                    .NavigateAsync(module, offset, references, includeDeclaration, deadline.Token)
                    .ConfigureAwait(false);

                locations = [.. answered.Select(location => new SurfaceLocation(
                    location.Module, workbook, location.Line, location.Column, location.Length,
                    location.Preview, location.Kind))];

                Log.Info($"{(references ? "references" : "definition")}: {module}@{offset} -> {locations.Length}");
            }
            catch (Exception ex)
            {
                Log.Info($"{(references ? "references" : "definition")}: {module}@{offset} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowLocations(requestId, locations));
        });
    }

    /// <summary>
    /// Answers a colouring request from the surface. Addressed by module name the way the outline
    /// is, not taken from the shown module: a split shows two at once and the editor colours both.
    /// </summary>
    private void OnSemanticTokensRequested(int requestId, string moduleName, string? projectDisplay)
    {
        var surface = _editorSurface;

        if (surface is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowSemanticTokens(requestId, [], failed: true);
            return;
        }

        var projectId = ProjectIdFromDisplay(projectDisplay);

        _ = Task.Run(async () =>
        {
            SurfaceSemanticToken[] tokens = [];
            var failed = false;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                var answered = await analysis.SemanticTokensAsync(moduleName, projectId, deadline.Token)
                    .ConfigureAwait(false);

                tokens = [.. answered.Select(token =>
                    new SurfaceSemanticToken(token.Start, token.End, token.Type, token.Modifiers))];
            }
            catch (Exception ex)
            {
                failed = true;
                Log.Info($"semanticTokens: {moduleName} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowSemanticTokens(requestId, tokens, failed));
        });
    }

    /// <summary>
    /// Analyses the shown module's text as typed, once the typing rests. This is what keeps the
    /// squiggles describing the text on screen rather than the text as of the last write-back:
    /// an error the developer deleted goes away on the next pause, not the next module write.
    /// The caret rides along so the engine holds back the transient complaints of the
    /// expression being typed. A result computed for text that has moved on is dropped - the
    /// keystroke that moved it has already scheduled the next pass.
    /// </summary>
    private void OnLiveAnalysisDue()
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            return;
        }

        // The folder the typed text names, on this same cadence (#23).
        NoteTypedFolder(_shownProject, module, source);

        // A full pass skipped during the typing runs once things go quiet.
        var now = Environment.TickCount64;
        if (_fullAnalysisDeferred && now - _lastFullAnalysis > FullAnalysisQuietMilliseconds)
        {
            _lastFullAnalysis = now;
            _fullAnalysisDeferred = false;
            analysis.Reanalyse();
        }

        var lineStarts = TextPositions.LineStarts(source);
        var caretOffset = TextPositions.ToOffset(lineStarts, surface.CaretLine, surface.CaretColumn);

        // Captured on the host thread now: the answer replaces this project's copy of the
        // module, never a same-named module in another workbook.
        var shownProject = _shownProject;
        bool SameHome(Finding finding) =>
            string.Equals(finding.Module, module, StringComparison.OrdinalIgnoreCase)
            && (finding.Project is null || shownProject is null
                || string.Equals(finding.Project, shownProject, StringComparison.OrdinalIgnoreCase));

        _ = Task.Run(async () =>
        {
            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                var findings = await analysis.DiagnoseLiveAsync(module, source, caretOffset, deadline.Token)
                    .ConfigureAwait(false);

                if (findings is null)
                {
                    return;
                }

                Log.Info($"live: {module} -> {findings.Count} finding(s)");

                surface.RunOnHostThread(() =>
                {
                    if (surface.Module != module || !ReferenceEquals(surface.Text, source) && surface.Text != source)
                    {
                        Log.Info($"live: {module} answer was for older text, dropped");
                        return;
                    }

                    // Unchanged findings are not republished: every publish redraws the panel
                    // and the tree, and most pauses change nothing.
                    var existing = _findings.Where(SameHome);
                    if (existing.SequenceEqual(findings))
                    {
                        return;
                    }

                    _findings = [.. _findings.Where(finding => !SameHome(finding)), .. findings];
                    PublishMarkersToSurface();
                    PublishFindingsToSurface();
                });
            }
            catch (Exception ex)
            {
                Log.Info($"live: {module} failed ({ex.GetType().Name})");
            }
        });
    }

    /// <summary>Runs a command the developer chose from the toolbar.</summary>
    private void RunCommand(string name, string? project)
    {
        /*
         * A DIALOG RAISED FROM A WORKBOOK'S ROW IS ABOUT THAT WORKBOOK.
         *
         * References and Project Properties are the editor's OWN dialogs and they act on the
         * ACTIVE project - so the tree's menu opened whichever workbook happened to be in
         * front, whatever row was right-clicked, with nothing on screen to say so (the owner,
         * 2026-08-20: "it seems it's loading for the workbook for the active tab, no matter
         * which workbook i right click on"). The page now says which workbook it meant, and
         * the project is made active first, which is exactly what clicking that project in
         * the editor's own explorer would do before opening the same dialog.
         */
        if (project is { Length: > 0 })
        {
            ActivateProject(project);
        }

        // The designer's save callback: the RAW File Save, skipping the designer branch -
        // that branch is what asked the page to apply in the first place, and letting the
        // callback loop back into it would apply forever.
        if (string.Equals(name, "saveOnly", StringComparison.Ordinal))
        {
            ExecuteEditorCommand(VbeCommands.Command.Save, skipDesignerApply: true);
            return;
        }

        // F5's callback: the launch alone, no save. The native editor never saves on Run, and
        // the save that used to sit here was worse than a parity break: on a never-saved
        // workbook it raised the Save As dialog, the posted Run launched the form OVER it, the
        // dialog waited out the form's modal loop and appeared the moment the form closed, and
        // cancelling it then unwound two interleaved modal loops, which the editor did not
        // always survive (the owner, 2026-08-27: "it brings up a save as dialog... and if i
        // hit cancel it crashes"). What F5 must guarantee - the window that opens is the
        // document - is the APPLY, which already happened on the way here.
        if (string.Equals(name, "runOnly", StringComparison.Ordinal))
        {
            ExecuteEditorCommand(VbeCommands.Command.Run, skipDesignerApply: true);
            return;
        }

        var command = VbeCommands.ForName(name);
        if (command == 0)
        {
            Log.Info($"command: '{name}' is not one of ours");
            return;
        }

        // A BUTTON THAT DID NOTHING SAYS SO.
        //
        // The editor disables its own commands by state - Reset unless something is stopped,
        // Break unless something is running - and this discarded that answer, so pressing a
        // button the editor would not honour looked exactly like pressing one it did. The strip
        // draws Reset enabled whatever the host thinks, so the developer clicks, nothing happens,
        // and nothing anywhere says why (the owner, 2026-08-23, pointing at Reset on a project
        // stopped in an Enum: "I think you'd need to trigger this" - it had been pressed, and it
        // had been refused, silently).
        //
        // Noticing costs one line and is the difference between a product that will not do
        // something and a product that appears broken.
        var outcome = ExecuteEditorCommand(command);
        if (!outcome.Ran)
        {
            Log.Info($"command: the strip's '{name}' was refused ({outcome.Detail})");
            _editorSurface?.Notify($"{name}: the editor refused it - {outcome.Detail}.");
        }
    }

    /// <summary>
    /// Issues Reset for a break this product's own evaluation caused, off the host thread.
    ///
    /// OFF IT, deliberately. The caller is a poll tick ON the host thread, and Reset raises a
    /// confirmation whose message loop is that same thread - so issuing it inline would park the
    /// tick inside a modal and there would be nothing left to answer it. The command is marshalled
    /// back instead, from a pool thread that stays free, which is the shape the immediate route's
    /// own long note arrived at: the door's thread issues, the host thread pumps, the frame
    /// unwinds.
    ///
    /// The confirmation is answered by the expectation's own watcher, which is why that had to
    /// start answering rather than only choosing.
    /// </summary>
    private void ClearOurOwnBreak()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        _ = System.Threading.Tasks.Task.Run(() =>
        {
            try
            {
                using (Diagnostics.DialogWatch.ExpectingConfirmation(6000))
                {
                    surface.RunOnHostThread(() => ExecuteEditorCommand(VbeCommands.Command.Reset));

                    // WAITED ON THE POLLED FLAG, not on ProjectModeNow. That reads the project
                    // over COM, and this is a pool thread: an apartment-bound object touched from
                    // the wrong thread is how this product has killed Excel before, and it does
                    // not fail politely. The poll owns the host thread and publishes what it saw;
                    // reading its bool from here is free and safe.
                    var backBy = Environment.TickCount64 + 5000;
                    while (Environment.TickCount64 < backBy && _inBreak)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }

                // SAID ON SCREEN, not only in the log. This clears a stop without being asked,
                // which is right when the stop is ours and nothing of the developer's is on the
                // stack - but a state that changes by itself and says nothing is the shape of
                // every silent thing found today. If the discrimination is ever wrong, the person
                // it is wrong for should be able to see what happened rather than wonder.
                if (!_inBreak)
                {
                    Log.Info("debug: the compile-error break is cleared; the project is back in design mode");
                    _editorSurface?.Notify("The compile error stopped the project, so it was reset. "
                        + "Nothing had run, so nothing was lost.");
                }
                else
                {
                    Log.Info("debug: the compile-error break would NOT clear; press Reset in the editor");
                }
            }
            catch (Exception ex)
            {
                Log.Info($"debug: clearing our own break failed ({ex.GetType().Name}: {ex.Message})");
            }
        });
    }

    /// <summary>
    /// Makes one open file's project the editor's active one, so a command that acts on the
    /// ACTIVE project acts on the file the developer named. Answers whether it moved; a name
    /// nothing answers to is logged and the command runs where it would have run anyway,
    /// because refusing to open a dialog at all would be the worse of the two wrongs.
    /// </summary>
    private bool ActivateProject(string display)
    {
        try
        {
            using var wanted = FindProjectByDisplayName(display);
            if (wanted is null)
            {
                Log.Info($"command: no open file called {display}; the active project stands");
                return false;
            }

            _editor.SetObject("ActiveVBProject", wanted);
            return true;
        }
        catch (Exception ex)
        {
            Log.Warn($"command: {display} could not be made active, {ex.Message}");
            return false;
        }
    }

    /// <summary>Puts the native pane's caret where the surface's caret is.</summary>
    private void SyncCaretToPane()
    {
        var surface = _editorSurface;
        if (surface?.Module is not { } module)
        {
            return;
        }

        try
        {
            using var pane = FindCodePane(module);
            pane?.Invoke("SetSelection", surface.CaretLine, surface.CaretColumn, surface.CaretLine, surface.CaretColumn);
        }
        catch (Exception ex)
        {
            Log.Error($"caret: could not be moved to {module}({surface.CaretLine},{surface.CaretColumn})", ex);
        }
    }

    /// <summary>Reads a module's text and hands it to the surface, with its squiggles.</summary>
    private void ShowModuleInSurface(string component, string? projectId = null)
    {
        using var found = FindComponent(component, projectId, out var owner);
        if (found is null)
        {
            return;
        }

        var source = ProjectReader.ReadSource(found);
        if (source is null)
        {
            return;
        }

        // What "shown" means from here on: bare names arriving from the page resolve to this
        // project first, and markers for a same-named module elsewhere stay off this surface.
        // The analysis service hears it too, so completions, hover, and live text for a shared
        // name address the workbook actually on screen.
        var movedFile = !string.Equals(_shownProject, owner, StringComparison.OrdinalIgnoreCase);
        _shownProject = owner;
        if (_analysis is not null)
        {
            _analysis.PreferredProject = owner;
        }

        // THE TESTS PANE SPEAKS FOR THE FILE THE DEVELOPER IS IN, so it has to hear when that
        // changes. Which files the pane lists is decided by `Shown()`, and one of its three rules
        // is "the file being worked in, even with nothing to say" - so moving between workbooks
        // changes the list, the install chip's answer, and whether the file select appears at all.
        // Nothing published for it. Every read of the `tests` route was correct and the PANE was
        // whatever the last install or refresh left behind: with DebugFixture and TwinFixture open
        // and TwinFixture on screen, the route answered both files and "missing" while the pane
        // still drew one file and a green "XlideAssert Installed" from 36 minutes earlier
        // (2026-08-21, the owner's report).
        //
        // Only when the FILE moves. Switching modules inside one workbook changes nothing the
        // pane shows, and this is on the path every tab change takes.
        if (movedFile)
        {
            PublishTests();
        }

        // A project the engine has never been seeded with - a workbook just opened or created.
        // Nothing else would ask for the pass: only this session's own writes call Reanalyse,
        // which is how an externally added workbook stayed unanalysed forever. Gated, because
        // a new workbook shows two panes in quick succession and the first pass has not had
        // time to make the project known before the second show asks.
        if (owner is not null && _analysis is { } analysisService && !analysisService.KnowsProject(owner)
            && Environment.TickCount64 - _lastUnknownProjectPass > 2000)
        {
            _lastUnknownProjectPass = Environment.TickCount64;
            Log.Info($"analysis: {component}'s project is new, analysing everything");
            analysisService.Reanalyse();
        }

        // Opening a module also selects it, the way the editor's own tree behaves.
        _propertiesTarget = component;
        _propertiesControl = null;

        var display = DisplayFromProjectId(owner);
        var written = WrittenKey(component, display);
        var toldBefore = _writtenModules.TryGetValue(written, out var already) ? already : null;
        _writtenModules[written] = source;
        _editorSurface?.Show(component, display, source);

        // The engine's live copy starts from what is being shown; the keystrokes stream from
        // here as edits.
        //
        // ONLY WHEN IT HAS MOVED. This sent the module's whole text on EVERY activation, so
        // clicking a tab handed the engine a megabyte and a half of text it already had and made
        // it re-parse the module from scratch - with the pipe serving one call at a time, every
        // other request queued behind that. Switching to a 64,802-line tab froze the editor for
        // two to three seconds, reported twice by the owner (2026-08-21); the same switch is
        // under a tenth of that when the text has not changed. The surface's own Show above
        // already worked this way, and posts nothing when the page's copy matches.
        if (!string.Equals(toldBefore, source, StringComparison.Ordinal))
        {
            _analysis?.NotifyLiveText(component, source, null);
        }
        Log.Info($"editor surface: showing {component}, {source.Length} character(s)");

        // A hold belongs to the module it began in; the switch is the caret leaving the line.
        _ = _activeLineHold.Release();

        // The findings for this module were computed before it was opened, so they are applied here
        // rather than waiting for the next analysis pass.
        PublishMarkersToSurface();
        PublishFindingsToSurface();
        PublishBreakpoints();
        PublishProperties();
    }

    /// <summary>
    /// Tells the surface which modules the editor has open, for its tab strip.
    ///
    /// The list comes from the editor's own collection of open panes rather than from the project's
    /// components, so the tabs are the modules the developer actually has open, not every module
    /// that exists. Reading a component's pane would create one, which would put a tab up for a
    /// module nobody opened.
    /// </summary>
    /// <summary>The mode the title last reported, so a refresh costs no COM call.</summary>
    private string? _titleMode = "design";

    /// <summary>
    /// Workbook display names by project id, filled where the tree already walks every project.
    ///
    /// The title bar needs the workbook on every tab switch, and a tab switch can cross workbooks.
    /// Reading it from the debugger's tick cached the WRONG one: the module updated on each switch
    /// and the workbook did not, so a tab in the other workbook was labelled with the first one's
    /// name (2026-08-07). Deriving it from the id instead spells it "debugfixture.xlsm", because
    /// the id is a lowercased path.
    /// </summary>
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _projectNames =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Puts the current workbook, module and mode on the title bar.
    ///
    /// CALLED WHEREVER THE SHOWN MODULE CHANGES, not only from the tick that reads the mode. It
    /// lived on that tick alone at first and the title lagged a tab behind: switching modules left
    /// the previous one's name on the window until something else made the session poll, and the
    /// poll stops when the editor is idle, which is exactly when somebody is reading the title.
    ///
    /// NOT A POLL of its own, and it must not become one. The mode and the workbook are cached
    /// from the tick that already had to read them for the debugger; the module comes off the
    /// surface; and Apply writes only when the composed caption differs from what is up. So
    /// hanging this on the tab strip's publish, which is an event, costs a few string comparisons
    /// and no round trip to the editor.
    ///
    /// APPLY IS THE ONLY COMPARISON, and this must not grow another. It had one: a guard here that
    /// returned early when the mode, workbook and module all matched what was last written. That
    /// compares our record of the caption against our record of the caption, and the thing it has
    /// to notice is the HOST rewriting the window underneath us - which changes neither.
    ///
    /// Pressing Reset twice was enough. The first press leaves break, the mode changes, the guard
    /// lets it through. The second press does nothing to the execution state and everything to the
    /// caption: the editor rewrites the window as "Microsoft Visual Basic for Applications", the
    /// mode is still design, the module is still the same, and the guard sent the refresh home. The
    /// product's name was off its own window until the next tab switch (measured 2026-08-09: three
    /// polls ran in that state, each one calling this, each one returning early).
    ///
    /// Apply compares against what the window ACTUALLY READS, which is the only comparison that can
    /// see that happen, and it writes only on a difference, so this cannot chase its own tail.
    /// </summary>
    private void RefreshWindowTitle()
    {
        if (_hostChrome is not { } chrome)
        {
            return;
        }

        var shown = _editorSurface?.Module;

        // The workbook comes from the SHOWN project, read now, for the same reason the module
        // does: a tab switch can cross workbooks, and a cached one labels the new tab with the old
        // workbook's name. Cased from the map when it is known, since the id is a lowercased path.
        var workbook = _shownProject is { Length: > 0 } id
            && _projectNames.TryGetValue(id, out var cased)
                ? cased
                : DisplayFromProjectId(_shownProject);

        chrome.Mode = _titleMode;
        chrome.Workbook = workbook;
        chrome.Module = shown;
        chrome.Apply();
    }

    private void PublishModules()
    {
        // Timed because this runs on every poll tick and its cost was asserted rather than
        // measured (the audit's B23). Microseconds, because the unchanged pass sits under a
        // millisecond and the sample ring drops zeros; perf().publishUs serves the figures.
        var publishTimer = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            PublishModulesCore();
        }
        finally
        {
            PerfCounters.Publish(publishTimer.Elapsed.Ticks / 10);
        }
    }

    private void PublishModulesCore()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        // The tab strip changing is the moment the shown module can have changed, which is the
        // moment the title bar is wrong until somebody says so.
        RefreshWindowTitle();

        // A refusing collection changes nothing: the strip keeps its last picture and the
        // tracker's recovery republishes it.
        if (ReadOpenModules() is { } modules)
        {
            // Whether anything was live is read BEFORE the prune empties the table: the
            // empty-workspace branch below is what tells the page, and pruning first would
            // leave nothing for its condition to see.
            var hadDocuments = surface.OpenDocuments.Count > 0;

            // Documents follow the pane list both ways: one closed natively leaves the table
            // (its unwritten edits flushed on the way out), and the page prunes its models
            // from the same published list.
            surface.PruneDocuments([.. modules.Select(m => (m.Name, m.Project))]);

            // Designer tabs join the same strip. Pruned here only of closed WORKBOOKS (their
            // project id stops resolving to a display); a form removed while its tab stands is
            // collected lazily, when the tab's next markup request answers "no component".
            _designerTabs.RemoveAll(tab => DisplayFromProjectId(tab.ProjectId) is not { Length: > 0 });
            if (_activeDesignerTab is { } held && !_designerTabs.Contains(held))
            {
                _activeDesignerTab = null;
            }

            // A NATIVE move takes the active slot back: the developer clicked a code pane, or
            // an activation landed one. Only a move TO A PANE counts - a transition to no
            // module at all is a pane closing or an empty workspace, and treating it as focus
            // stripped a designer tab of the slot in the very publish that granted it, which
            // is how a designer tab opened over an empty workspace stood inactive and unshown
            // (2026-08-13, the second face of "clicking back does not activate").
            var nativeActive = (surface.Module, _shownProject);
            if (nativeActive != _lastNativeActive)
            {
                _lastNativeActive = nativeActive;
                if (nativeActive.Item1 is not null)
                {
                    _activeDesignerTab = null;
                }
            }

            // Closing the LAST pane leaves the surface holding a document nobody can see a
            // tab for, when no window event reaches the tracker to say so - the mirror of
            // the empty view that outlived its panes (2026-08-06). The object model is the
            // authority both ways: an empty open list with a module still shown IS the
            // empty workspace, and every close route passes through this publish. A standing
            // designer tab keeps the workspace: it is a tab, just not the host's.
            if (modules.Count == 0 && _designerTabs.Count == 0 && hadDocuments)
            {
                Log.Info("editor surface: the last module closed, showing the empty workspace");
                _watchingEmpty = true;
                surface.Clear();
                UpdatePolling();
            }

            // Dirty is a WORKBOOK fact - the editor persists all of a workbook's modules
            // together (probed 2026-08-04: a module edit flips Workbook.Saved false, Save
            // flips it true) - so it is read once per workbook and worn by every tab the
            // workbook owns.
            //
            // But the host's flag alone over-reports: it never flips back when an edit is
            // undone to the saved text, because the host does not diff. The dot should
            // (developer, 2026-08-04), so module text is snapshotted whenever the workbook is
            // known clean, and while the host says dirty the dot shows only if some module's
            // known text actually differs from its snapshot. A module with no snapshot to
            // compare keeps the flag's word.
            // WorkbookSaved crosses into the object model, so it is asked once per workbook even
            // though the answer is used once per module.
            var savedByProject = new Dictionary<string, bool?>(StringComparer.OrdinalIgnoreCase);

            // Any live document's text counts, not just the active one: a background tab's
            // typing is exactly as unsaved as the shown one's.
            string? CurrentTextOf(string module, string? display) =>
                surface.TextOf(module, display)
                    ?? (_writtenModules.TryGetValue(WrittenKey(module, display), out var written) ? written : null);

            // One module, one answer. The dot used to be a workbook fact worn by every tab the
            // workbook owned, on the grounds that saving persists all of a workbook's modules
            // together. That is true of SAVING and false of the question a tab strip is asked,
            // which is which of these have I changed (developer, 2026-08-06). A workbook-wide dot
            // marks four untouched tabs for one edit and tells nobody which one to look at.
            bool DirtyOf(string module, string? project)
            {
                var display = DisplayFromProjectId(project);
                if (display is not { Length: > 0 })
                {
                    return false;
                }

                if (!savedByProject.TryGetValue(display, out var saved))
                {
                    saved = WorkbookSaved(display);
                    savedByProject[display] = saved;
                }

                // Unreadable must not invent a dot, and a saved workbook has nothing unsaved in it.
                // A clean moment is also when a baseline is the saved truth, so it is taken here.
                if (saved != false)
                {
                    if (saved == true && CurrentTextOf(module, display) is { } clean)
                    {
                        _savedBaselines[BaselineKey(display, module)] = clean;
                    }

                    return false;
                }

                var text = CurrentTextOf(module, display);
                if (text is null)
                {
                    // Something in this workbook is unsaved, but nothing here says it was this
                    // module, and a dot on a tab whose text was never read is a guess.
                    return false;
                }

                if (!_savedBaselines.TryGetValue(BaselineKey(display, module), out var baseline))
                {
                    // First sight of this module, with the workbook already dirty. Taking that as
                    // proof THIS module made it dirty put a dot on a tab nobody had typed in, the
                    // moment it was opened -- and on a workbook that has never been saved,
                    // permanently, because the clean moment that writes baselines never comes
                    // (developer, 2026-08-06: a fresh tab showing the unsaved dot immediately).
                    //
                    // The text as first read becomes the baseline instead, which is also what
                    // Don't Save restores, so the dot and the revert agree about what unchanged
                    // means.
                    _savedBaselines[BaselineKey(display, module)] = text;
                    return false;
                }

                return !string.Equals(text, baseline, StringComparison.Ordinal);
            }

            // Designer tabs append after the mirrored panes, wearing their face. Their dirty
            // is always false HERE: unapplied markup is the page's own state, and the page
            // wears that dot itself rather than asking the host to relay it.
            var designerRows = _designerTabs
                .Select(tab => (tab.Module, Project: DisplayFromProjectId(tab.ProjectId)))
                .ToList();

            string[] names = [.. modules.Select(m => m.Name), .. designerRows.Select(d => d.Module)];
            string?[] projects = [.. modules.Select(m => m.Project), .. designerRows.Select(d => d.Project)];
            bool[] dirty = [.. modules.Select(m => DirtyOf(m.Name, m.Project)), .. designerRows.Select(_ => false)];
            string?[]? faces = designerRows.Count == 0
                ? null
                : [.. modules.Select(_ => (string?)null), .. designerRows.Select(_ => (string?)"design")];

            string? active;
            string? activeProject;
            string? activeFace;
            if (_activeDesignerTab is { } designer)
            {
                active = designer.Module;
                activeProject = DisplayFromProjectId(designer.ProjectId);
                activeFace = "design";
            }
            else
            {
                active = surface.Module;
                activeProject = DisplayFromProjectId(_shownProject);
                activeFace = null;
            }

            // Sent on change only: the polls re-derive this several times a second during an
            // episode, and an identical strip is not news to the page or the log.
            var key = string.Join("|", names) + "\n" + string.Join("|", projects)
                + "\n" + string.Join("|", dirty) + "\n" + active + "\n" + activeProject
                + "\n" + string.Join("|", faces ?? []) + "\n" + activeFace;
            if (key == _lastModulesKey)
            {
                return;
            }

            _lastModulesKey = key;
            Log.Verbose($"modules: publish [{string.Join(",", names)}] dirty [{string.Join(",", dirty)}]"
                + (designerRows.Count == 0 ? string.Empty : $" designer [{string.Join(",", designerRows.Select(d => d.Module))}]"));
            surface.ShowModules(names, projects, active, activeProject, dirty, faces, activeFace);
        }
    }

    /// <summary>What the strip was last sent, so unchanged republishes send nothing.</summary>
    private string? _lastModulesKey;

    /// <summary>The host application, kept between reads; a refusal drops it for a re-find.</summary>
    private DispatchObject? _hostApp;

    /// <summary>
    /// Each module's text as of the workbook's last known-clean moment, keyed by workbook and
    /// module. This is what lets an edit undone back to the saved text drop its dot: the host
    /// flag never un-dirties, but the text can be compared.
    /// </summary>
    private readonly Dictionary<string, string> _savedBaselines = new(StringComparer.Ordinal);

    private static string BaselineKey(string display, string module) =>
        (display + "\0" + module).ToLowerInvariant();

    /// <summary>
    /// The application's workbook wearing this display name, through the trust-free route, or
    /// null when the application or the name cannot be found. One walk, shared by the Saved
    /// read and the save itself: the two were the same Workbooks/Count/GetItem/match-on-Name
    /// loop written twice, differing only in what they did on the match (the audit's B23).
    /// Ownership transfers - the caller disposes; exceptions propagate to the callers, whose
    /// recoveries differ on purpose.
    /// </summary>
    private DispatchObject? FindWorkbookByDisplay(string display)
    {
        _hostApp ??= HostApplication.Find();
        if (_hostApp is null)
        {
            return null;
        }

        using var books = _hostApp.GetObject("Workbooks");
        var count = books?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            var book = books!.GetItem(i);
            if (book is not null
                && string.Equals(book.GetString("Name"), display, StringComparison.OrdinalIgnoreCase))
            {
                return book;
            }

            book?.Dispose();
        }

        return null;
    }

    /// <summary>
    /// A workbook's Saved flag, by display name, through the same trust-free application route
    /// the evaluator uses. Null when it cannot be read - a missing dot is a small wrong, a
    /// lying dot is a large one, so unknown must never invent one.
    /// </summary>
    private bool? WorkbookSaved(string display)
    {
        try
        {
            using var book = FindWorkbookByDisplay(display);
            return book?.GetBool("Saved");
        }
        catch (Exception)
        {
            // The application answer went stale - a workbook closed mid-read, or the host is
            // busy. Re-found on the next read.
            _hostApp?.Dispose();
            _hostApp = null;
        }

        return null;
    }

    /// <summary>
    /// What the developer calls a project, from its identity alone: the file's name for a saved
    /// workbook, the identity itself otherwise. Lowercase for saved ones - comparisons on the
    /// page side are case-insensitive.
    /// </summary>
    /// <summary>
    /// A module's text reduced to something two sides can be compared on.
    ///
    /// Line endings are normalised and trailing blank lines dropped, because the host and the
    /// page genuinely disagree about both and neither disagreement is a defect: VBA stores CRLF
    /// and counts a trailing line the page does not draw. Everything else must match exactly -
    /// a single changed character is a real difference and has to register as one.
    ///
    /// Null for no text at all, which is a different answer from empty text and is reported as
    /// such: a module the surface does not hold is not a module it holds wrongly.
    /// </summary>
    private static string? ContentKey(string? text)
    {
        if (text is null)
        {
            return null;
        }

        var normalised = text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
        return $"{normalised.Length}:{normalised.GetHashCode(StringComparison.Ordinal)}";
    }

    private string? DisplayFromProjectId(string? projectId)
    {
        if (string.IsNullOrEmpty(projectId))
        {
            return null;
        }

        // The name the tree was built with, when there is one. An unsaved workbook's id carries
        // its COM identity ("vbaproject#24ce5821b58") because its name is not unique, and that
        // is an id to key by, never a thing to show a developer.
        if (_projectNames.TryGetValue(projectId, out var shown) && shown.Length > 0)
        {
            return shown;
        }

        return projectId.Contains('\\') || projectId.Contains('/')
            ? Path.GetFileName(projectId)
            : projectId.Contains('#')
                ? projectId[..projectId.IndexOf('#')]
                : projectId;
    }

    /// <summary>
    /// The modules with open panes, or null when the pane collection refuses entirely.
    ///
    /// From the collection, per item and tolerantly, and never from the tracker's window map. The
    /// windows cannot carry this list: with maximised panes the editor keeps a window only for
    /// the ACTIVE pane and destroys the others, so the map holds one entry however many panes are
    /// open. What the collection cannot be trusted with is a member that has just died, so a dead
    /// member is skipped rather than taking the list with it.
    /// </summary>
    private List<(string Name, string? Project)>? ReadOpenModules()
    {
        try
        {
            using var panes = _editor.GetObject("CodePanes");
            var count = panes?.GetInt32("Count") ?? 0;

            var modules = new List<(string Name, string? Project)>(count);
            for (var i = 1; i <= count; i++)
            {
                try
                {
                    using var pane = panes!.GetItem(i);
                    using var module = pane?.GetObject("CodeModule");
                    using var component = module?.GetObject("Parent");

                    if (component?.GetString("Name") is not { Length: > 0 } name
                        || IsScratchComponent(name))
                    {
                        continue;
                    }

                    // The workbook the tab belongs to, by the name the tree uses, so the strip
                    // can say WHICH Module1 when two workbooks hold one.
                    string? owner = null;
                    try
                    {
                        using var collection = component.GetObject("Collection");
                        using var project = collection?.GetObject("Parent");
                        if (project is not null)
                        {
                            // THE SAME NAME THE TREE USES, which is not always the project's own.
                            //
                            // Two unsaved workbooks are both "VBAProject", so the tree numbers
                            // them; this labelled a tab's project with the raw name instead, and
                            // the two halves of the same window disagreed. The strip published
                            // `projects: ["VBAProject", ...]` beside `activeProject:
                            // "VBAProject 01"`, so activating a module in a numbered workbook
                            // matched no tab and quietly did nothing: clicking Sheet1 under
                            // VBAProject 01 left whatever was already open on screen, which
                            // reads as the click landing on the wrong workbook (2026-08-08).
                            owner = DisplayFromProjectId(ProjectReader.Identity(project).Id);
                        }
                    }
                    catch (Exception)
                    {
                        // A tab without a workbook is still a tab.
                    }

                    if (!modules.Any(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(m.Project, owner, StringComparison.OrdinalIgnoreCase)))
                    {
                        modules.Add((name, owner));
                    }
                }
                catch (Exception)
                {
                    // The pane that would not answer is the one that is going away.
                }
            }

            return modules;
        }
        catch (Exception ex)
        {
            Log.Info($"modules: the pane list would not answer ({ex.GetType().Name})");
            return null;
        }
    }

    /// <summary>
    /// Sends the surface the whole project tree, for its explorer.
    ///
    /// Every component, not only the ones with a pane open: this is what the developer navigates
    /// by, so it has to show modules that have never been opened. Reading a component's pane would
    /// create one, so nothing here touches CodeModule.
    /// </summary>
    private void PublishProjects()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        try
        {
            using var projects = _editor.GetObject("VBProjects");
            var projectCount = projects?.GetInt32("Count") ?? 0;

            /*
             * TWO UNSAVED WORKBOOKS ARE BOTH CALLED "VBAProject", AND THE TREE SHOWED BOTH.
             *
             * A workbook that has never been saved has no file name, and for a while the only
             * name the tree could call it was the project's own - "VBAProject" for every new
             * workbook, two of them two identical rows (reported 2026-08-08). The display walks
             * the same axis as saved workbooks now: the host's own name, "Book1", unique by the
             * host's construction (#14). The numbering below stays as the backstop for any name
             * collision that still arises - a host that answers no document name falls back to
             * the project name, and two of THOSE side by side still need telling apart.
             *
             * A name shared by more than one project is numbered: "VBAProject 01",
             * "VBAProject 02", in the order the editor lists them. A name that is unique is left
             * exactly as it is, because "Book1.xlsm 01" would be noise.
             *
             * Counted first, over the whole set, because numbering needs to know whether there is
             * anything to disambiguate FROM. The names are then reused below rather than read
             * again, so the tree, the title bar and every id-to-name lookup agree on one spelling.
             */
            var displays = new List<string>(projectCount);
            for (var i = 1; i <= projectCount; i++)
            {
                using var project = projects!.GetItem(i);
                displays.Add(project is null ? "VBAProject" : WorkbookDisplayName(project));
            }

            var shared = displays
                .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var numbered = new List<string>(displays.Count);
            foreach (var name in displays)
            {
                if (!shared.Contains(name))
                {
                    numbered.Add(name);
                    continue;
                }

                seen[name] = seen.TryGetValue(name, out var count) ? count + 1 : 1;
                numbered.Add($"{name} {seen[name]:00}");
            }

            var tree = new List<SurfaceProject>(projectCount);
            var live = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 1; i <= projectCount; i++)
            {
                using var project = projects!.GetItem(i);
                using var components = project?.GetObject("VBComponents");
                if (project is null || components is null)
                {
                    continue;
                }

                // The folder rides each component (#23): from the cache the analysis pass fills,
                // or a one-time read of the declarations for a module no pass has described.
                var identity = ProjectReader.Identity(project).Id;
                var members = new List<SurfaceComponent>();
                ForEachRealComponent(project, (component, name) =>
                    members.Add(new SurfaceComponent(name, component.GetInt32("Type"), FolderOf(identity, component, name))));
                PruneFolders(identity, members.Select(member => member.Name));

                // The cased name against the id, so the title bar can name the workbook on a tab
                // switch without a COM call. The id is a lowercased path and everything derived
                // from it reads "debugfixture.xlsm"; this is the spelling the shell uses.
                //
                // The NUMBERED name, from the pass above, so the tree and the title bar call a
                // workbook the same thing. A name only the tree knew would leave the title bar
                // saying "VBAProject" for both of two workbooks the tree had told apart.
                var display = numbered[i - 1];
                _projectNames[ProjectReader.Identity(project).Id] = display;

                tree.Add(new SurfaceProject(display, [.. members]));
                live.Add(ProjectReader.Identity(project).Id);
            }

            surface.ShowProjects([.. tree]);

            // A NAME IS FOR A PROJECT THAT EXISTS. This map never forgot, so a project the
            // editor holds for a beat while a closed workbook is torn down kept its entry for
            // the rest of the session - and anything that trusted "the tree knows this one"
            // went on trusting it, which is how a ghost file called VBAProject kept appearing
            // in the Tests pane after a close (measured 2026-08-20).
            foreach (var stale in _projectNames.Keys.Where(id => !live.Contains(id)).ToList())
            {
                _projectNames.TryRemove(stale, out _);
            }

            // The tree is where a file opening or closing is noticed, so it is where the Tests
            // pane hears about either: its own picture is merged per analysis snapshot, and
            // neither event produces one.
            ReconcileTestFiles(live);
        }
        catch (Exception ex)
        {
            Log.Error("explorer: the project tree could not be read", ex);
        }
    }

    /// <summary>
    /// What the developer calls a project: its workbook's file name. The project's own name is
    /// almost always the default 'VBAProject', which distinguishes nothing; the file name is how
    /// the companion editor's tree names workbooks too. A workbook never saved has no file name
    /// and raises when asked for one, and keeps the project name instead.
    /// </summary>
    private static string WorkbookDisplayName(DispatchObject project)
    {
        try
        {
            if (project.GetString("FileName") is { Length: > 0 } fileName)
            {
                return Path.GetFileName(fileName);
            }
        }
        catch (Exception)
        {
            // Unsaved: the property raises rather than answering empty.
        }

        // A workbook never saved still has a name - the one the host gave it. "Book1" is what
        // the titlebar, the native explorer ("VBAProject (Book1)") and Workbooks("Book1") all
        // say, and the host numbers those unique by construction. Naming these by the PROJECT
        // instead made two new workbooks two identical "VBAProject" rows - the reason the
        // numbering below exists at all - and left Book1 unaddressable through every route,
        // while the refusal promised that "a display name resolves" (#14). The saved case above
        // is already Workbook.Name; this keeps the unsaved case on the same axis.
        return UnsavedWorkbookName(project) ?? project.GetString("Name") ?? "VBAProject";
    }

    /// <summary>
    /// The host's own name for a never-saved workbook ("Book1", "Document1"), read off the
    /// document module's Properties - the road the Properties panel already reads workbook
    /// rows by, so the trust switch stays out of it. The module is found by SHAPE, not name:
    /// "ThisWorkbook" is localized, a string-valued "FullName" property among a document
    /// component's rows is not, and only the workbook's own module carries one. Null when no
    /// component answers the shape, and the caller's project-name fallback is the name these
    /// workbooks had before this existed.
    /// </summary>
    private static string? UnsavedWorkbookName(DispatchObject project)
    {
        try
        {
            using var components = project.GetObject("VBComponents");
            var count = components?.GetInt32("Count") ?? 0;
            for (var i = 1; i <= count; i++)
            {
                using var component = components!.GetItem(i);
                if (component?.GetInt32("Type") != 100)
                {
                    continue;
                }

                if (DocumentPropertyValue(component, "FullName") is not { Length: > 0 })
                {
                    continue;   // document-shaped, but a sheet: no FullName row
                }

                if (DocumentPropertyValue(component, "Name") is { Length: > 0 } answer)
                {
                    return answer;
                }
            }
        }
        catch (Exception)
        {
            // A project mid-teardown answers nothing; so does a host with no document modules.
        }

        return null;
    }

    /// <summary>One string-valued row of a document component's Properties, or null - a
    /// component that does not carry the row answers by throwing, and that is an answer.</summary>
    private static string? DocumentPropertyValue(DispatchObject component, string property)
    {
        try
        {
            using var properties = component.GetObject("Properties");
            using var row = properties?.CallObject("Item", property);
            return row?.GetString("Value");
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// The project shown under a workbook name, or null when none matches.
    ///
    /// THE NAME THE TREE SHOWED, WHICH IS NOT ALWAYS THE PROJECT'S OWN. Two unsaved workbooks are
    /// both called "VBAProject", so the tree numbers them "VBAProject 01" and "VBAProject 02",
    /// and that is what comes back on a click. Comparing against the project's own name matched
    /// neither, so every caller fell back to the shown project: clicking Sheet1 under either new
    /// workbook opened DebugFixture's Sheet1 instead (reported twice, 2026-08-08).
    ///
    /// This is the one place every route resolves a name through, so the numbering is understood
    /// here rather than at each caller.
    /// </summary>
    private DispatchObject? FindProjectByDisplayName(string? displayName)
    {
        if (string.IsNullOrEmpty(displayName))
        {
            return null;
        }

        // What the tree meant, when the tree is what asked. `_projectNames` is id to the name it
        // was shown under, so reversing it turns a numbered name back into an identity.
        string? wantedId = null;
        foreach (var (id, shown) in _projectNames)
        {
            if (string.Equals(shown, displayName, StringComparison.OrdinalIgnoreCase))
            {
                wantedId = id;
                break;
            }
        }

        /*
         * AN IDENTITY NAMES A PROJECT TOO, and this product hands identities out.
         *
         * `projects` answers a projectId that is the workbook's full path and `projectHolding`
         * returns it to a caller, so the answer to "which workbook holds this module" was an
         * argument this could not match: not a shown name, not a file name. Every caller then took
         * its own fallback - the shown project, or the ACTIVE one - and answered confidently about
         * a workbook nobody asked about. With two workbooks each holding a Helpers that is simply
         * the wrong module, silently (measured 2026-08-11).
         *
         * Nine call sites resolve names through here, several of them `?? ActiveVBProject`, so
         * this is the place to understand every spelling rather than each caller.
         */
        if (wantedId is null)
        {
            foreach (var id in _projectNames.Keys)
            {
                if (string.Equals(id, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    wantedId = id;
                    break;
                }
            }
        }

        // A path nothing has been shown under yet still names a workbook by its file name, which
        // is what an unnumbered display is.
        var byFileName = wantedId is null
            && (displayName.Contains('\\', StringComparison.Ordinal)
                || displayName.Contains('/', StringComparison.Ordinal))
                ? Path.GetFileName(displayName)
                : null;

        try
        {
            using var projects = _editor.GetObject("VBProjects");
            var count = projects?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                var project = projects!.GetItem(i);
                if (project is null)
                {
                    continue;
                }

                // The item is held in a plain local because a match hands it to the caller, so it
                // cannot be a `using`. That makes the read below the dangerous part: both routes
                // to a name read a property, an unsaved workbook raises rather than answering, and
                // the catch outside this loop would swallow the throw with the wrapper still in
                // hand. Dispose on the way out and let it carry on to that catch.
                bool matches;
                try
                {
                    matches = wantedId is not null
                        ? string.Equals(ProjectReader.Identity(project).Id, wantedId, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(WorkbookDisplayName(project), displayName, StringComparison.OrdinalIgnoreCase)
                            || (byFileName is not null
                                && string.Equals(WorkbookDisplayName(project), byFileName, StringComparison.OrdinalIgnoreCase));
                }
                catch
                {
                    project.Dispose();
                    throw;
                }

                if (matches)
                {
                    return project;
                }

                project.Dispose();
            }
        }
        catch (Exception ex)
        {
            Log.Info($"project: '{displayName}' could not be looked up ({ex.GetType().Name})");
        }

        return null;
    }

    /// <summary>
    /// Closes the editor's own windows for the panels this product replaces.
    ///
    /// Closed rather than covered. The editor hides a tool window on request and a hidden window
    /// cannot be uncovered by anything the editor does afterwards, which is the failure mode that
    /// covering them would have: the editor raises its own windows on all sorts of occasions and
    /// wins every one of those races. Closing them also gives the document area their space, which
    /// is what the surface is measured against.
    ///
    /// The objects stay alive and the project is untouched, so anything reading them keeps working.
    /// Only the windows for panels that exist in the surface are closed; a native window with no
    /// replacement is left alone, because taking it away would remove the feature rather than
    /// restyle it.
    /// </summary>
    private void HideReplacedWindows()
    {
        if (_windowsHidden)
        {
            return;
        }

        _windowsHidden = true;

        // The project explorer and the Immediate window have surface replacements. The properties
        // window does not need its native form; it is closed for the dock space it occupies, and
        // the menu route opens ours. The Object Browser (2) is put away too: the editor
        // remembers it open across sessions, and a remembered one reappeared blank at start-up
        // (2026-08-05). The Locals and Watch windows are NOT hidden here: both become ghost
        // palettes a moment later - the editor only feeds a window with a paintable surface
        // (lessons 25 and 29), and hiding one removes the feature rather than restyling it.
        const int immediateWindow = 5;
        ReadOnlySpan<int> replaced = [immediateWindow, 2, 6, 7];

        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is null || !replaced.Contains(window.GetInt32("Type")))
                {
                    continue;
                }

                if (window.GetInt32("Type") != immediateWindow)
                {
                    if (window.GetBool("Visible"))
                    {
                        window.SetBool("Visible", false);
                        Log.Info($"window: closed the editor's own '{window.GetString("Caption")}'");
                    }

                    continue;
                }

                HideImmediateWindow(window);
            }
        }
        catch (Exception ex)
        {
            Log.Error("window: the replaced windows could not be closed", ex);
        }
    }

    /// <summary>
    /// The ghost palettes' reading thread, started once both palettes are prepared. It owns the
    /// Locals and Watches readers; the host thread asks it to read and looks at what landed.
    /// </summary>
    private GhostReaderThread? _ghostReaders;

    /// <summary>The floated Locals palette, its original extended styles kept for restoration.</summary>
    private nint _localsPalette;
    private long _localsPaletteExStyle;

    /// <summary>The floated Watches palette, its original extended styles kept for restoration.</summary>
    private nint _watchPalette;
    private long _watchPaletteExStyle;

    /// <summary>The native window kinds this product ghosts, as the object model numbers them.</summary>
    private const int WatchWindowType = 3;
    private const int LocalsWindowType = 4;

    /// <summary>
    /// Turns a native panel window into an invisible data engine: floated through the object
    /// model, ghosted, parked off screen, and read by handle.
    ///
    /// The editor only feeds an ON-SCREEN window (lesson 25), but "on screen" turned out to mean
    /// "has a paintable surface". A LAYERED window renders into its own surface regardless of
    /// occlusion or position, so a floated palette with WS_EX_LAYERED at alpha zero, parked far
    /// off the virtual screen, is fed faithfully through every break and step while being
    /// impossible to see, click, or discover. Probed 2026-08-04: the counter tracked 1 through 4
    /// across steps at alpha 1, at alpha 0, and at -20000,-20000. The themed panel renders what
    /// the reader reads; nothing native is ever visible.
    ///
    /// Floating uses LinkedWindows.Remove on the window's own linked frame, pure object model.
    /// If any step refuses, the ghost is skipped, and the police pass hides the docked native
    /// window: the canvas stays pure and the panel sits idle.
    ///
    /// ONE ROUTINE FOR BOTH PANELS. Locals and Watches were prepared by two copies of this that
    /// differed in a window number and a word in the log, and the copy had already lost every
    /// comment above (2026-08-09). The reader is not created here either way: it lives on the
    /// ghost reading thread, which starts once both palettes are prepared.
    /// </summary>
    private bool PrepareGhostPalette(int windowType, string name, out nint palette, out long exStyle)
    {
        palette = 0;
        exStyle = 0;

        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is null || window.GetInt32("Type") != windowType)
                {
                    continue;
                }

                window.SetBool("Visible", true);

                // Undocked; a window already floating answers Remove with an error worth nothing.
                try
                {
                    using var frame = window.GetObject("LinkedWindowFrame");
                    using var linked = frame?.GetObject("LinkedWindows");
                    linked?.InvokeWithObject("Remove", window);
                }
                catch (Exception ex)
                {
                    Log.Verbose($"{name}: the palette would not undock ({ex.GetType().Name}); it may already be floating");
                }

                // Small: the reader reads the store, not the viewport, and the store does not
                // shrink with the window.
                try
                {
                    window.SetInt32("Left", 300);
                    window.SetInt32("Top", 300);
                    window.SetInt32("Width", 240);
                    window.SetInt32("Height", 150);
                }
                catch (Exception ex)
                {
                    Log.Verbose($"{name}: the palette would not be placed ({ex.GetType().Name})");
                }

                var caption = window.GetString("Caption");
                var found = caption is { Length: > 0 }
                    ? CodePaneTracker.FindTopLevelByCaption(caption)
                    : 0;

                if (found == 0)
                {
                    Log.Info($"{name}: the floated palette was not found; the native window stays");
                    return false;
                }

                palette = found;
                exStyle = Win32.GetWindowLongPtr(found, Win32.GwlExStyle);
                Win32.SetWindowLongPtr(found, Win32.GwlExStyle,
                    (nint)(exStyle | Win32.WsExLayered | Win32.WsExTransparent | Win32.WsExNoActivate));
                Win32.SetLayeredWindowAttributes(found, 0, 0, Win32.LwaAlpha);
                Win32.SetWindowPos(found, 0, -20000, -20000, 0, 0,
                    Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate);

                Log.Verbose($"{name}: palette {found:X} floated and ghosted");
                return true;
            }
        }
        catch (Exception ex)
        {
            Log.Info($"{name}: the ghost could not be prepared ({ex.GetType().Name}: {ex.Message})");
        }

        return false;
    }

    /// <summary>
    /// Gives a palette back to the native editor: opaque, its styles restored, on screen, and
    /// hidden until someone asks for it. A stopped session must leave a usable window.
    /// </summary>
    private void RestoreGhostPalette(int windowType, string name, ref nint palette, long exStyle)
    {
        if (palette == 0)
        {
            return;
        }

        try
        {
            Win32.SetWindowPos(palette, 0, 300, 300, 0, 0,
                Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate);
            Win32.SetLayeredWindowAttributes(palette, 0, 255, Win32.LwaAlpha);
            Win32.SetWindowLongPtr(palette, Win32.GwlExStyle, (nint)exStyle);

            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;
            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is not null && window.GetInt32("Type") == windowType)
                {
                    window.SetBool("Visible", false);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"{name}: the palette could not be restored ({ex.GetType().Name})");
        }

        palette = 0;
    }

    private void PrepareLocalsGhost() =>
        PrepareGhostPalette(LocalsWindowType, "locals", out _localsPalette, out _localsPaletteExStyle);

    private void PrepareWatchGhost() =>
        PrepareGhostPalette(WatchWindowType, "watch", out _watchPalette, out _watchPaletteExStyle);

    private void RestoreLocalsPalette() =>
        RestoreGhostPalette(LocalsWindowType, "locals", ref _localsPalette, _localsPaletteExStyle);

    private void RestoreWatchPalette() =>
        RestoreGhostPalette(WatchWindowType, "watch", ref _watchPalette, _watchPaletteExStyle);

    /// <summary>
    /// Hides every docked native toolbar, which the surface's toolbar and menus replace.
    ///
    /// Hidden only because everything on them is somewhere else now: each button is a menu
    /// command, the menus are all on the surface's bar, and every run and step command keeps the
    /// key it always had. The toggles that would bring the bars back are suppressed from the
    /// surface's View menu, so a hidden bar stays hidden. A docked bar left visible would claim
    /// rows the surface covers - since it stopped retreating for chrome, that is a toolbar on
    /// screen that cannot be pressed.
    ///
    /// Floating bars are left alone: they float above the surface and contest nothing. So is the
    /// menu bar object, whose commands back the surface's own menus.
    /// </summary>
    private void HideNativeToolbars()
    {
        const int menuBarType = 1;
        const int floating = 4;
        const int popup = 5;

        try
        {
            using var bars = _editor.GetObject("CommandBars");
            var count = bars?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var bar = bars!.GetItem(i);
                if (bar is null || !bar.GetBool("Visible") || bar.GetInt32("Type") == menuBarType)
                {
                    continue;
                }

                var position = bar.GetInt32("Position");
                if (position == floating || position == popup)
                {
                    continue;
                }

                bar.SetBool("Visible", false);
                Log.Info($"window: hid the editor's docked '{bar.GetString("Name")}' toolbar");
            }
        }
        catch (Exception ex)
        {
            Log.Info($"window: the editor's toolbars could not be hidden ({ex.GetType().Name})");
        }
    }

    /*
     * Frame colours, as the compositor wants them: one byte each of blue, green and red, in that
     * order, which is the reverse of how they are written everywhere else. They match the surface's
     * dark theme so the window is one thing rather than a dark document in a pale frame.
     */
    private static readonly int BorderColour = 0x002D2D2D;
    private static readonly int CaptionColour = 0x001E1E1E;
    private static readonly int CaptionTextColour = 0x00D4D4D4;

    /// <summary>
    /// Asks the system to draw the editor's title bar dark.
    ///
    /// The title bar is drawn by the desktop compositor, not by the editor, so nothing the editor
    /// or this add-in paints can reach it. The compositor will draw it dark on request, and that
    /// request is the only way to change it.
    ///
    /// The attribute was renumbered once, before it was documented. Both numbers are tried because
    /// which one works depends on the build of Windows rather than on anything observable here.
    /// </summary>
    private static void DarkenTitleBar(nint frame)
    {
        var dark = 1;

        if (Win32.DwmSetWindowAttribute(frame, Win32.UseDarkTitleBar, in dark, sizeof(int)) < 0)
        {
            Win32.DwmSetWindowAttribute(frame, Win32.UseDarkTitleBarLegacy, in dark, sizeof(int));
        }

        // The border and the caption are drawn by the compositor too, and dark mode alone leaves
        // the border light: a pale rectangle around an otherwise dark window. Colours are given
        // explicitly so the frame matches the surface rather than approximately matching it.
        // These are refused on Windows versions that predate them, which is why nothing checks.
        Win32.DwmSetWindowAttribute(frame, Win32.BorderColor, in BorderColour, sizeof(int));
        Win32.DwmSetWindowAttribute(frame, Win32.CaptionColor, in CaptionColour, sizeof(int));
        Win32.DwmSetWindowAttribute(frame, Win32.CaptionTextColor, in CaptionTextColour, sizeof(int));

        // The frame is redrawn only when it is told something about it changed.
        Win32.SetWindowPos(
            frame,
            0,
            0,
            0,
            0,
            0,
            Win32.SwpNoMove | Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate | Win32.SwpFrameChanged);

        Log.Info("window: asked for a dark title bar");
    }

    /// <summary>
    /// Closes the Immediate window and remembers which window it was.
    ///
    /// It is shown first, deliberately. Its class is shared with the code panes and with the Locals
    /// and Watch windows and its caption is localised, so the only thing that identifies it is
    /// which window stops being visible when it is closed. That comparison needs it to have been
    /// visible, and the editor reports it as visible before it has created it.
    ///
    /// The window survives being hidden, keeping its handle and its contents, which is what makes
    /// Debug.Print readable afterwards.
    /// </summary>
    private void HideImmediateWindow(DispatchObject window)
    {
        window.SetBool("Visible", true);

        // The identification that survives timing: the object model names the window's
        // localised caption, and the handle carries the same caption, visible or not. The old
        // way - diffing which pane stopped being visible across the hide - lost whenever the
        // hide had not yet reached the window list, and answered "0 windows changed"; that is
        // a hidden Immediate whose Debug.Print output nothing can mirror. The caption is kept
        // so a failed attachment can be retried when the first evaluation actually needs it.
        _immediateCaption = window.GetString("Caption");
        var handle = _immediateCaption is { Length: > 0 } caption
            ? CodePaneTracker.FindPaneByCaption(caption)
            : 0;

        var before = handle == 0 ? CodePaneTracker.VisiblePanes() : null;

        window.SetBool("Visible", false);
        Log.Info($"window: closed the editor's own '{_immediateCaption}'");

        if (handle == 0 && before is not null)
        {
            // The old diff as the fallback for a host where the caption read nothing.
            before.ExceptWith(CodePaneTracker.VisiblePanes());
            if (before.Count == 1)
            {
                handle = before.First();
            }
            else
            {
                Log.Info($"immediate: {before.Count} windows changed when it closed, cannot tell which it is");
            }
        }

        AttachImmediateReader(handle);
    }

    /// <summary>The Immediate window's localised caption, kept so attachment can be retried.</summary>
    private string? _immediateCaption;

    /// <summary>
    /// Publishes every finding to the surface's panel, across all modules - except the ones the
    /// active-line hold is keeping back, so the panel and the badges never announce a verdict
    /// about a line still being typed.
    /// </summary>
    /// <summary>
    /// Forgets one module's findings and redraws. For a module whose text has been replaced behind
    /// the analyzer's back: stale findings outlive their code otherwise, because a live answer is
    /// only accepted while its module is the one on screen.
    /// </summary>
    private void DropFindingsFor(string component, string? projectDisplay)
    {
        var remaining = _findings
            .Where(finding => !(
                string.Equals(finding.Module, component, StringComparison.OrdinalIgnoreCase)
                && (projectDisplay is null
                    || string.Equals(DisplayFromProjectId(finding.Project), projectDisplay, StringComparison.OrdinalIgnoreCase))))
            .ToList();

        if (remaining.Count == _findings.Count)
        {
            return;
        }

        Log.Info($"findings: dropped {_findings.Count - remaining.Count} for {component}, its text was replaced");
        _findings = remaining;
        PublishMarkersToSurface();
        PublishFindingsToSurface();
    }

    private void PublishFindingsToSurface()
    {
        // The engine's findings and this session's own - the attribute annotations' drift -
        // in one list, held to the same active-line rule.
        _editorSurface?.ShowFindings([.. _findings.Concat(AttributeFindings())
            .Where(f => !_activeLineHold.Hides(f.Module, f.StartLine, f.EndLine))
            .Select(f => new SurfaceFinding(
                f.Module,
                f.Code,
                f.Message,
                f.Severity,
                f.StartLine,
                f.StartColumn,
                DisplayFromProjectId(f.Project)))]);
    }

    /// <summary>
    /// A close asked for by the developer - the tab's X, its middle-click, Ctrl+W, or the tab
    /// menu. A module whose text still differs from the workbook's last saved text does not just
    /// close: the developer is asked first, and the answer comes back through the same message
    /// with a choice on it. "save" saves the workbook - the editor persists all of a workbook's
    /// modules together, so saving the workbook is what saving a module means - and "discard"
    /// writes the saved text back over the module, which is the closest thing to closing without
    /// saving that a document living inside a workbook can have.
    ///
    /// The question gates on the module's OWN text, not the workbook dot. The dot is a workbook
    /// fact and can stand on a sibling's changes; a question here offers a revert of THIS module,
    /// so it is only asked when this module's changes can be named - and therefore reverted.
    /// </summary>
    /// <summary>
    /// What a close request came to. The tab's X does not read this - it either sees its tab go or
    /// sees the question appear - but a script driving the same path has neither, and answering it
    /// a bare true was the same lie the `open` branch beside it was fixed for. `Awaiting` is the
    /// third answer that matters: a confirm on screen is not a close and not a failure, and a
    /// caller told which one it got can answer the question instead of guessing.
    /// </summary>
    internal readonly record struct CloseOutcome(bool Closed, string Detail, string? Awaiting)
    {
        public static CloseOutcome Went(string detail = "closed") => new(true, detail, null);

        public static CloseOutcome Stayed(string detail) => new(false, detail, null);

        public static CloseOutcome Asking(string question) =>
            new(false, "the developer was asked before closing", question);
    }

    private CloseOutcome OnModuleCloseRequested(string component, string? projectDisplay, string? action)
    {
        var display = projectDisplay is { Length: > 0 }
            ? projectDisplay
            : DisplayFromProjectId(_shownProject);

        switch (action)
        {
            case "save":
            {
                // Whatever is mid-debounce belongs to the save.
                _editorSurface?.FlushEdits();

                if (display is null || !SaveWorkbookOf(display))
                {
                    // Closing on a failed save would lose the very thing the developer asked
                    // to keep. The tab stays; the log says why.
                    Log.Warn($"close: {display ?? "the workbook"} would not save; {component}'s tab stays");
                    return CloseOutcome.Stayed($"{display ?? "the workbook"} would not save");
                }

                CloseModule(component, projectDisplay);
                return CloseOutcome.Went("saved and closed");
            }

            case "discard":
            {
                // A refused revert still closes the tab, so this is not a failure to report as
                // one - but it is the difference between "the edits are gone" and "the edits are
                // still in the module", which a caller cannot see from a closed tab.
                string? revertRefused = null;

                if (display is not null
                    && _savedBaselines.TryGetValue(BaselineKey(display, component), out var baseline))
                {
                    // The debounced write of the abandoned text must not chase the revert -
                    // this document's write only; a sibling tab's typing keeps its debounce.
                    _editorSurface?.DiscardEdits(component, display);

                    // Discarding means putting the SAVED text back, so a refusal here leaves the
                    // module holding the edits the developer just asked to throw away. Said out
                    // loud, because the tab is about to close over it.
                    if (WriteModule(component, baseline, ProjectIdFromDisplay(display), hostRewrite: true)
                        is { } refused)
                    {
                        _editorSurface?.Notify($"The saved text could not be put back: {refused}");
                        revertRefused = refused;
                    }

                    // The surface may still show the abandoned text. The close below replaces
                    // the document anyway; this covers the close that cannot find a pane, so
                    // whatever stays on screen is the module's truth.
                    ResyncFromModule();

                    // The findings still describe the text that was just thrown away, and nothing
                    // else will replace them: a live answer is dropped unless its module is still
                    // the one on screen, and this one is about to close. So the panel kept
                    // reporting errors from code that no longer existed (developer, 2026-08-06).
                    //
                    // Dropped rather than recomputed here. Whatever the reverted text really
                    // contains comes back from the next analysis pass, which is the same path that
                    // would have reported it had the module never been opened.
                    DropFindingsFor(component, display);

                    Log.Info($"close: {component} reverted to {display}'s saved text");
                }
                else
                {
                    // No snapshot to go back to. Closing with the text kept is the only honest
                    // remainder - inventing a revert target would destroy real work.
                    Log.Info($"close: {component} has no saved text to revert to; closing as it is");
                }

                CloseModule(component, projectDisplay);
                return revertRefused is null
                    ? CloseOutcome.Went("discarded and closed")
                    : CloseOutcome.Went($"closed, but the saved text could not be put back: {revertRefused}");
            }

            default:
            {
                if (display is not null && ModuleDiffersFromSaved(component, display))
                {
                    Log.Verbose($"close: {component} differs from {display}'s saved text; asking");
                    _editorSurface?.ConfirmClose(component, projectDisplay);
                    return CloseOutcome.Asking("confirm");
                }

                CloseModule(component, projectDisplay);
                return CloseOutcome.Went();
            }
        }
    }

    /// <summary>
    /// Whether a module's current text is known to differ from the workbook's last saved text.
    /// The same comparison the tab dot makes, with the opposite default: unknown is not
    /// different here, because this gates a question whose "Don't Save" performs a revert, and
    /// a revert needs both sides of the difference in hand.
    /// </summary>
    private bool ModuleDiffersFromSaved(string component, string display)
    {
        var current = _editorSurface?.TextOf(component, display)
            ?? (_writtenModules.TryGetValue(WrittenKey(component, display), out var written) ? written : null);

        return current is not null
            && _savedBaselines.TryGetValue(BaselineKey(display, component), out var baseline)
            && !string.Equals(current, baseline, StringComparison.Ordinal);
    }

    /// <summary>
    /// Saves a workbook by display name, through the same trust-free application route the
    /// Saved flag is read by. False when the workbook cannot be found or the save is refused -
    /// the caller keeps the tab open on false, so it needs the truth, not best effort.
    /// </summary>
    private bool SaveWorkbookOf(string display)
    {
        try
        {
            using var book = FindWorkbookByDisplay(display);
            if (book is null)
            {
                Log.Warn($"close: {display} is not among the application's workbooks");
                return false;
            }

            book.Invoke("Save");
            Log.Info($"close: saved {display}");

            // A SAVE IS A ROUND BOUNDARY. "When I last saved" is already how a developer thinks
            // about a last-good state, so the change log draws its line in the same place. It
            // costs nothing - a round is a divider, not a copy - and it means a log read tomorrow
            // is grouped the way the work actually went.
            CloseChangeRounds($"saved {display}");
            return true;
        }
        catch (Exception ex)
        {
            Log.Error($"close: {display} could not be saved", ex);
            _hostApp?.Dispose();
            _hostApp = null;
        }

        return false;
    }

    /// <summary>
    /// Closes a module's pane, which is what closing its tab means. The editor activates another
    /// pane afterwards and the surface follows it; closing the last one hides the surface, which
    /// is the empty state there is.
    ///
    /// The pane is looked for in the editor's own open list, never through the component. Reading
    /// a component's pane CREATES one when none is open, so closing that way closes forever: a
    /// second click on a dead tab was conjuring a fresh pane just to destroy it, and the editor
    /// spent a minute churning through the wreckage.
    /// </summary>
    private void CloseModule(string component, string? projectDisplay = null)
    {
        var projectId = ProjectIdFromDisplay(projectDisplay);

        try
        {
            if (CloseThroughObjectModel(component, projectId))
            {
                return;
            }

            // The pane collection would not say, but the tracker already knows which window shows
            // the module, and asking the window to close is exactly what its own close box does.
            // This is what keeps closing alive while the collection is refusing everyone.
            if (_codePanes is not null)
            {
                foreach (var pane in _codePanes.Panes)
                {
                    if (string.Equals(pane.Component, component, StringComparison.OrdinalIgnoreCase)
                        && (projectId is null || pane.Project is null
                            || string.Equals(pane.Project, projectId, StringComparison.OrdinalIgnoreCase))
                        && Win32.IsWindow(pane.Window))
                    {
                        Win32.PostMessage(pane.Window, Win32.WmClose, 0, 0);
                        Log.Info($"module: asked {component}'s window to close");
                        return;
                    }
                }
            }

            Log.Info($"module: {component} has no open pane to close");
        }
        finally
        {
            // The teardown a close starts finishes AFTER this returns, and its window events
            // can land while a refresh is already mid-flight. The tracker queues those now,
            // but a close is the one moment staleness is guaranteed visible - the tab must
            // leave the strip - so the next few polls re-derive the picture unconditionally
            // rather than trusting the event stream alone.
            _resyncPanePolls = 3;
            _pollsRemaining = Math.Max(_pollsRemaining, (int)(2_000 / DebugPollMilliseconds));
            UpdatePolling();

            // Closing a pane makes the editor activate the next one, and activation takes the
            // keyboard with it. Without this, the second Ctrl+W in a row went to a window the
            // developer cannot see.
            _editorSurface?.Focus();
        }
    }

    /// <summary>Polls that must re-derive the pane picture regardless of events. See CloseModule.</summary>
    private int _resyncPanePolls;

    /// <summary>Closes a module's pane through the editor's own pane list. False when it cannot.</summary>
    private bool CloseThroughObjectModel(string component, string? projectId = null)
    {
        try
        {
            using var panes = _editor.GetObject("CodePanes");
            var count = panes?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                try
                {
                    using var pane = panes!.GetItem(i);
                    using var module = pane?.GetObject("CodeModule");
                    using var owner = module?.GetObject("Parent");
                    if (!string.Equals(owner?.GetString("Name"), component, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // The tab said which workbook it belongs to; a same-named pane in another
                    // workbook is not the one being closed.
                    if (projectId is not null && owner is not null)
                    {
                        using var collection = owner.GetObject("Collection");
                        using var project = collection?.GetObject("Parent");
                        if (project is not null
                            && !string.Equals(ProjectReader.Identity(project).Id, projectId, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                    }

                    using var window = pane!.GetObject("Window");
                    window?.Invoke("Close");
                    Log.Info($"module: closed {component}");
                    return true;
                }
                catch (Exception)
                {
                    // A pane that will not answer is already going away.
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"module: the pane list would not answer for {component} ({ex.GetType().Name})");
        }

        return false;
    }

    /// <summary>
    /// Adds a component to the active project and opens it: 1 module, 2 class module, 3 form.
    ///
    /// Through the project's own collection rather than a menu, because the operation is the
    /// documented one and it names what it made, which is what lets the new module open at once.
    /// </summary>
    private void InsertComponent(int kind, string? projectName)
    {
        if (kind is not (1 or 2 or 3))
        {
            return;
        }

        // The page's menu already hides "New UserForm" where MSForms does not exist, but the
        // message is a wire anyone can speak: the refusal has to live where the insert does.
        if (kind == 3 && !HostApp.CarriesMsForms)
        {
            _editorSurface?.Notify($"{HostApp.Name} VBA has no UserForms, so none can be added");
            return;
        }

        try
        {
            // The workbook the menu was opened on, when it was; the active project otherwise.
            using var project = FindProjectByDisplayName(projectName) ?? _editor.GetObject("ActiveVBProject");
            using var components = project?.GetObject("VBComponents");
            using var added = components?.CallObject("Add", kind);
            var name = added?.GetString("Name");
            if (added is not null)
            {
                // Before the change log records the add, so the module is recorded as it was born.
                SeedOptionExplicit(added);
            }

            Log.Info($"project: inserted {name ?? "?"} (kind {kind})");

            // A MODULE ARRIVING IS A CHANGE. Only the write path recorded anything until now, so
            // a module could be added, filled and removed and the log would show the filling with
            // no sign of either end (the owner asked, 2026-08-22).
            if (name is not null && added is not null)
            {
                RecordChange(
                    name, ProjectIdFromDisplay(projectName) ?? _shownProject,
                    Core.Changes.ChangeKind.Added, null, ProjectReader.ReadSource(added) ?? string.Empty,
                    componentKind: ComponentKind(kind));
            }

            if (name is not null)
            {
                ShowModule(name);
            }

            // Not just the analyzer: the strip and the tree both have to hear about it too, and
            // this used to tell only the analyzer - so a module inserted from the page appeared
            // in neither list until something else happened to republish them.
            ComponentsChanged();
        }
        catch (Exception ex)
        {
            Log.Error($"project: a component of kind {kind} could not be inserted", ex);
            _editorSurface?.Notify($"The component could not be inserted: {ex.Message}");
        }
    }

    /// <summary>
    /// Removes a component and everything this session was still holding on its behalf.
    ///
    /// The COM call is one line; the rest is why this is a method. A component that goes leaves
    /// four kinds of record behind, and each one that survives is a lie about a module that no
    /// longer exists: the page's unwritten edits, this session's baseline of what it last wrote,
    /// the breakpoints recorded against its lines, and the properties panel's target.
    ///
    /// The unwritten edits are DISCARDED rather than flushed, and the order matters. Pruning a
    /// closed document flushes its pending text, which is right for a tab being closed and wrong
    /// here: the write would land on a component that is already gone, and it reaches WriteModule,
    /// which complains to the developer about a module they just deleted. So the edits are dropped
    /// before the removal, not left for the prune that follows it.
    ///
    /// A document module is refused rather than attempted. ThisWorkbook and a sheet's code belong
    /// to the workbook and the host answers Remove on them with a bare HRESULT.
    ///
    /// Returns null when the component is gone, and why it is not otherwise. The same shape the
    /// write and show routes were given this release, and for the same reason: a refusal that only
    /// reaches the log is a refusal the caller reports as success.
    /// </summary>
    private string? RemoveComponent(string component, string? projectName)
    {
        if (string.IsNullOrWhiteSpace(component))
        {
            return "no module was named";
        }

        try
        {
            var projectId = ProjectIdFromDisplay(projectName);
            using var doomed = FindComponent(component, projectId, out var foundIn);
            if (doomed is null)
            {
                Log.Warn($"project: nothing named {component} to remove"
                    + (projectName is null ? string.Empty : $" in {projectName}"));
                return $"There is no module named {component} to remove"
                    + (projectName is null ? "." : $" in {projectName}.");
            }

            // Read back the name the host holds rather than trusting the one the tree sent: the
            // messages below name what was actually removed, and the editor unifies identifier
            // case, so the two can differ in spelling.
            var removed = doomed.GetString("Name") ?? component;

            if (doomed.GetInt32("Type") == DocumentComponent)
            {
                Log.Warn($"project: {removed} is a document module and cannot be removed");
                return $"{removed} belongs to the workbook and cannot be removed.";
            }

            var owner = DisplayFromProjectId(foundIn ?? projectId);

            // Before the component goes, so the prune that follows finds nothing to write back.
            _editorSurface?.DiscardEdits(removed, owner);

            using var project = FindProjectByDisplayName(owner) ?? _editor.GetObject("ActiveVBProject");
            using var components = project?.GetObject("VBComponents");
            if (components is null)
            {
                Log.Warn($"project: {removed} could not be removed, the project would not open its components");
                return $"{removed} could not be removed: the project would not open its component list.";
            }

            // WHAT IT HELD, read while it still exists. A removal is the one change whose "before"
            // cannot be recovered afterwards, which makes it the one most worth having.
            var lastWords = ProjectReader.ReadSource(doomed);
            var lastKind = ComponentKind(doomed.GetInt32("Type"));

            // Remove takes the COMPONENT, not an index.
            components.InvokeWithObject("Remove", doomed);

            RecordChange(
                removed, foundIn ?? projectId, Core.Changes.ChangeKind.Removed, lastWords, null,
                componentKind: lastKind);

            // The bare-name key as well as the workbook-qualified one: a baseline recorded before
            // the workbook could be told is keyed without it, and the same fallback AdoptRename
            // walks for a rename applies to a removal.
            foreach (var key in new[] { WrittenKey(removed, owner), WrittenKey(removed, null) })
            {
                _writtenModules.Remove(key);
                _breakpoints.Remove(key);
            }

            Log.Info($"project: removed {removed}" + (owner is null ? string.Empty : $" from {owner}"));

            // A FORM'S DESIGNER TAB GOES WITH THE FORM. Nothing collected it here, so a removed
            // form left its tab standing over a canvas describing something that no longer
            // exists - and the next request for it published a reason onto the canvas instead
            // ("EntryForm2 has no designer", reported by the owner 2026-08-16, which is what a
            // form mid-teardown answers). The collector already exists for the case where a
            // request FINDS OUT; a removal knows, so it says so.
            CloseDesignerTab(removed, owner);

            // The strip, the tree, the analyzer and the Properties panel, the same four that an
            // insert has to tell. The panel used to be cleared by hand right here, which covered
            // this route and no other: PublishProperties drops a missing aim itself now, so a
            // removal through any path leaves it honest.
            ComponentsChanged();
            return null;
        }
        catch (Exception ex)
        {
            Log.Error($"project: {component} could not be removed", ex);
            return $"{component} could not be removed: {ex.Message}";
        }
    }

    /// <summary>
    /// Shows a module: the editor's active pane is set to it, and the surface is told to show it.
    ///
    /// Both are told directly. Show creates and displays a pane, but it does not reliably
    /// activate one that is already open behind another, and an activation that does not happen
    /// produces no window event and therefore, before this, no switch: clicking a tab did
    /// nothing exactly when both modules were already open. The active pane matters because the
    /// run and debug commands act on it; the surface matters because it is what the developer
    /// sees; neither is left to depend on the editor choosing to move a window.
    /// </summary>
    /// <summary>
    /// Null when the module is on screen. Anything else is why it is not.
    ///
    /// A name nothing answers to used to return here in silence, and the route above it replied ok
    /// to the caller, so asking for a module that is not there and asking for one that is looked
    /// identical from outside. The same shape as the write route, found the same afternoon.
    /// </summary>
    private string? ShowModule(string component, string? projectDisplay = null)
    {
        try
        {
            // Before the document changes underneath them.
            _editorSurface?.FlushEdits();

            // The tree says which workbook it means; a bare name resolves shown-project-first.
            var projectId = ProjectIdFromDisplay(projectDisplay);

            using var pane = FindCodePane(component, projectId);
            if (pane is null)
            {
                Log.Warn($"modules: nothing named {component} to show"
                    + (projectDisplay is null ? string.Empty : $" in {projectDisplay}"));
                return $"no module named {component}"
                    + (projectDisplay is null ? string.Empty : $" in {projectDisplay}");
            }

            pane.Invoke("Show");

            try
            {
                _editor.SetObject("ActiveCodePane", pane);
            }
            catch (Exception ex)
            {
                Log.Info($"modules: {component} could not be made the active pane ({ex.GetType().Name})");
            }

            if (_editorSurface?.Module != component || projectId is not null && _shownProject != projectId)
            {
                ShowModuleInSurface(component, projectId);
            }

            PublishModules();

            // Activating the native pane moved keyboard focus onto it, and the developer is not
            // looking at it. Without this, the shortcut that switches module works exactly once.
            _editorSurface?.Focus();
            return null;
        }
        catch (Exception ex)
        {
            Log.Error($"modules: {component} could not be shown", ex);
            return $"{component} could not be shown: {ex.Message}";
        }
    }

    /// <summary>Finds a component by name across every open project, or null when there is none.</summary>
    /// <summary>
    /// The project of the editor's own active pane, when its component carries this name. The
    /// caption-matched pane picture cannot tell two same-named modules apart; the object
    /// model's ActiveCodePane names both its component and its project without a caption.
    /// </summary>
    private string? ActivePaneOwner(string component)
    {
        try
        {
            using var active = _editor.GetObject("ActiveCodePane");
            using var module = active?.GetObject("CodeModule");
            using var found = module?.GetObject("Parent");
            if (found?.GetString("Name") is not { } name
                || !string.Equals(name, component, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            using var collection = found.GetObject("Collection");
            using var project = collection?.GetObject("Parent");
            return project is null ? null : ProjectReader.Identity(project).Id;
        }
        catch (Exception)
        {
            // No active pane, or an editor mid-teardown: the ambiguity simply stands.
            return null;
        }
    }

    private DispatchObject? FindComponent(string component) => FindComponent(component, null, out _);

    /// <summary>
    /// The component carrying this name - within one project when its identity is given, in
    /// whichever project answers first otherwise. The identity of the owning project comes back
    /// either way, which is how the session learns what "shown" means in a world where two
    /// workbooks may both hold a Module1.
    /// </summary>
    private DispatchObject? FindComponent(string component, string? projectId, out string? foundProject)
    {
        foundProject = null;

        using var projects = _editor.GetObject("VBProjects");
        var count = projects?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            using var project = projects!.GetItem(i);
            if (project is null)
            {
                continue;
            }

            var identity = ProjectReader.Identity(project).Id;
            if (projectId is not null && !string.Equals(identity, projectId, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var components = project.GetObject("VBComponents");
            if (components is null)
            {
                continue;
            }

            var componentCount = components.GetInt32("Count");
            for (var j = 1; j <= componentCount; j++)
            {
                var candidate = components.GetItem(j);
                if (candidate is null)
                {
                    continue;
                }

                // A plain local, because a match hands the component to the caller. The name read
                // is the part that can raise, and nothing here catches: the throw would leave the
                // finalizer thread to release the wrapper, which is not a release at all.
                string? name;
                try
                {
                    name = candidate.GetString("Name");
                }
                catch
                {
                    candidate.Dispose();
                    throw;
                }

                if (name == component)
                {
                    foundProject = identity;
                    return candidate;
                }

                candidate.Dispose();
            }
        }

        return null;
    }

    /// <summary>Finds the code pane a component's module is displayed in, opening one if needed.</summary>
    private DispatchObject? FindCodePane(string component, string? projectId = null)
    {
        // An explicit project wins; the shown project stands in when the name is the module on
        // the surface; anything else searches every project the way it always has.
        var owner = projectId ?? (IsShownComponent(component) ? _shownProject : null);

        using var found = FindComponent(component, owner, out _);
        using var module = found?.GetObject("CodeModule");

        // Reading CodePane on a module that has never been opened creates the pane, which is what
        // makes navigating to a module the user has not opened work at all.
        return module?.GetObject("CodePane");
    }

    /// <summary>True when this name is the module currently on the surface.</summary>
    private bool IsShownComponent(string component) =>
        string.Equals(component, _editorSurface?.Module, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The identity behind a workbook name the tree uses, or null when nothing matches. The
    /// tree names workbooks by file name, which is what FindProjectByDisplayName matches.
    /// </summary>
    private string? ProjectIdFromDisplay(string? display)
    {
        if (string.IsNullOrEmpty(display))
        {
            return null;
        }

        /*
         * THE NUMBERED NAME FIRST, because that is the name the surface was given.
         *
         * Two unsaved workbooks are both called "VBAProject", so the tree numbers them
         * "VBAProject 01" and "VBAProject 02" and that is what comes back on a click. No project
         * answers to those names, so the search below found nothing, the caller fell back to the
         * shown project, and clicking a module in either new workbook opened a module of
         * whichever workbook was already on screen (reported 2026-08-08, minutes after the
         * numbering went in and caused it).
         *
         * `_projectNames` is the map the tree was built from, id to the name it was shown under,
         * so reversing it answers exactly what the surface meant. Only then fall through to
         * matching a project's own name, which is what an unnumbered display is.
         */
        /*
         * AN IDENTITY IS ACCEPTED TOO, because this product hands them out and then would not
         * take them back.
         *
         * `projects` answers a `projectId` that is the workbook's full path, and `projectHolding`
         * on the client passes that straight to a caller - so the one route built to answer "which
         * workbook holds this module" produced an argument every route taking a `project` refused.
         * Refused SILENTLY: an unmatched name resolves to null, the caller falls back to whichever
         * workbook answers first, and with two workbooks each holding a Helpers the answer is the
         * wrong one. Measured 2026-08-11 with both fixtures open: `readModule("Helpers", twin.projectId)`
         * returned RenameFixture's copy, and `pane open` with the same argument opened the wrong
         * workbook's module and reported ran:true.
         *
         * That is the "a name is not an identity across workbooks" defect the two-workbook case
         * exists to catch, living in the identity plumbing itself. `project` (singular) confuses it
         * further by answering `projectId` as the DISPLAY name, which is why every suite works: they
         * all pass that one.
         */
        foreach (var id in _projectNames.Keys)
        {
            if (string.Equals(id, display, StringComparison.OrdinalIgnoreCase))
            {
                return id;
            }
        }

        foreach (var (id, shown) in _projectNames)
        {
            if (string.Equals(shown, display, StringComparison.OrdinalIgnoreCase))
            {
                return id;
            }
        }

        // Which understands a path and an identity as well as a shown name, so there is nothing
        // left to try here.
        using var project = FindProjectByDisplayName(display);
        return project is null ? null : ProjectReader.Identity(project).Id;
    }

    /// <summary>
    /// Sends each open document the squiggles that belong to it.
    ///
    /// Findings arrive for whole projects and the surface holds one model per open module
    /// (decision 12), so they are filtered per document. A document with none is sent an empty
    /// set rather than skipped: that is what clears squiggles the user has just fixed.
    ///
    /// The active-line hold call sites name the modules their hold change touched. A hold hides
    /// findings by module name, so beginning or releasing one can alter the markers of the held
    /// and released modules only - every other open document's set is provably identical to what
    /// it already shows, and re-sending those sets on every line the caret enters while typing
    /// was the audit's C12. No names means what it always meant: publish to every document.
    /// </summary>
    private void PublishMarkersToSurface(params string?[] touchedModules)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        foreach (var (module, project) in surface.OpenDocuments)
        {
            if (touchedModules.Length > 0
                && !touchedModules.Any(touched => touched is not null
                    && string.Equals(touched, module, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            // The attribute annotations' drift squiggles too: a finding the Problems pane lists
            // but the editor does not mark has no lightbulb, so its fixes cannot be reached from
            // the line they are about (the attributes suite, 2026-09-05).
            var markers = _findings.Concat(AttributeFindings())
                .Where(f => string.Equals(f.Module, module, StringComparison.OrdinalIgnoreCase)
                    && (f.Project is null || project is null
                        || string.Equals(DisplayFromProjectId(f.Project), project, StringComparison.OrdinalIgnoreCase))
                    && !_activeLineHold.Hides(f.Module, f.StartLine, f.EndLine))
                .Select(f => new EditorMarker(
                    f.StartLine,
                    f.StartColumn,
                    f.EndLine,
                    f.EndColumn,
                    f.Severity,
                    f.Message,
                    f.Code))
                .ToArray();

            surface.ShowDiagnostics(module, project, markers);
        }
    }

    /// <summary>The editor's frame window, kept for placements recomputed outside window events.</summary>
    private nint _frame;

    /// <summary>The editor's document area, kept for the same reason.</summary>
    private nint _documentArea;

    /// <summary>
    /// Whether the surface is meant to be on screen right now. A recomputed placement must not
    /// show a surface the session decided to hide because no pane is visible.
    /// </summary>
    private bool _surfaceShown;

    /// <summary>
    /// Where the surface goes inside the frame.
    ///
    /// Covering, it takes the frame's entire client area: the menu bar row included, because the
    /// surface draws its own menu bar backed by the same menus, and the document inset too, because
    /// the frame draws a pale line a pixel or two inside itself that no compositor attribute
    /// reaches. Anything less leaves native chrome showing through a themed product.
    ///
    /// Native tool windows do not shrink it: the canvas is purely xlide, and every native
    /// window is replaced, ghosted, floated, or policed away - because the
    /// old retreat to the document area handed the menu bar and toolbar rows back to the native
    /// editor every time one opened, and the product visibly reverted.
    ///
    /// While the page is still coming up, the loader takes the whole client area too, menu row
    /// included: a strip of native chrome over a branded splash reads as a defect, not as a menu.
    /// Only a loader that has stalled hands the menu row back, so a page that never arrives
    /// cannot keep every menu covered.
    /// </summary>
    private unsafe PixelRect SurfaceBounds(nint frame, nint documentArea, bool covering)
    {
        var document = ClientAreaIn(documentArea, frame);

        if (covering)
        {
            return FullClientBounds(frame, document);
        }

        if (_editorSurface is { IsReady: false } surface)
        {
            return surface.IsLoaderStalled
                ? BelowMenuBounds(frame, document)
                : FullClientBounds(frame, document);
        }

        return document;
    }

    private static unsafe PixelRect FullClientBounds(nint frame, PixelRect fallback)
    {
        Rect client;
        if (!Win32.GetClientRect(frame, &client))
        {
            return fallback;
        }

        return new PixelRect(0, 0, client.Right - client.Left, client.Bottom - client.Top);
    }

    /// <summary>
    /// The stalled loader's retreat: the client area below the native menu bar, which goes back
    /// to being usable while the loader explains itself. Falls back to the document area when
    /// the bar's height cannot be read.
    /// </summary>
    private unsafe PixelRect BelowMenuBounds(nint frame, PixelRect fallback)
    {
        Rect client;
        if (!Win32.GetClientRect(frame, &client))
        {
            return fallback;
        }

        var menuBottom = MenuBarHeight();
        var height = client.Bottom - client.Top;
        if (menuBottom <= 0 || menuBottom >= height)
        {
            return fallback;
        }

        // A PixelRect is edges, not origin-and-size: the fourth value is the bottom edge, which
        // is the client's own. Passing a height here left a menu-bar-sized band of native frame
        // below the loader, invisible in every other rectangle because their tops are zero.
        return new PixelRect(0, menuBottom, client.Right - client.Left, height);
    }

    /// <summary>The native menu bar's height in pixels, or zero when it cannot be read.</summary>
    private int MenuBarHeight()
    {
        try
        {
            using var menuBar = VbeMenus.FindMenuBar(_editor);
            return menuBar?.GetInt32("Height") ?? 0;
        }
        catch (Exception ex)
        {
            Log.Info($"surface: the menu bar's height could not be read ({ex.GetType().Name})");
            return 0;
        }
    }

    /// <summary>
    /// Whether the surface may cover everything native inside the frame.
    ///
    /// Only a page that has not loaded yet says no: covering the menu bar with a surface that
    /// cannot draw its own menus takes every menu away, so the native bar stays until the
    /// replacement is genuinely standing. Native tool windows stopped saying no long ago -
    /// today every one is replaced, ghosted, floated, or policed away, and the canvas is
    /// purely xlide.
    /// </summary>
    private bool CanCoverChrome() => _editorSurface is { IsReady: true };

    /// <summary>
    /// Recomputes where the surface belongs, now.
    ///
    /// For the moments that are not window events: a menu item has just opened or closed a native
    /// window, or the page has just come up. Waiting for the next window event leaves the wrong
    /// thing covered in the meantime.
    /// </summary>
    /// <summary>
    /// The synchronous resize path: bounds and nothing else. The full pass reads the object
    /// model and sweeps children for band regions, and per WM_SIZE tick of a drag that was
    /// the visible lag ("the resize draw lags behind the mouse", 2026-08-04). Inside the
    /// resize only the rectangle matters; the full pass follows at the settle.
    /// </summary>
    private void PlaceSurfaceFast()
    {
        if (_editorSurface is null || !_surfaceShown || _frame == 0 || _documentArea == 0
            || !Win32.IsWindowVisible(_frame))
        {
            return;
        }

        var startedAt = Environment.TickCount64;

        _editorSurface.Follow(SurfaceBounds(_frame, _documentArea, CanCoverChrome()), visible: true);

        // The full pass - window policing, band silencing, chrome - is object-model work,
        // and it was running once per frame event of a drag: measurable latency and repaint
        // churn (2026-08-05). The bounds followed above; everything else holds still until
        // the events pause, and then ONE full pass re-derives it all.
        _editorSurface.ArmPlacementSettle(PlacementSettleMilliseconds);

        PerfCounters.PlacementFast(Environment.TickCount64 - startedAt);
    }

    /// <summary>One quiet moment after the last frame event; then the full pass, once.</summary>
    private const uint PlacementSettleMilliseconds = 150;

    /// <summary>What the last followed pane event amounted to, geometry aside. An event that
    /// amounts to the same takes the fast path instead of the full cascade.</summary>
    private string? _lastFollowSubstance;

    /// <summary>
    /// A frame or document-area event: the frame moved, resized, was shown, or is going away.
    /// Bounds follow synchronously - that is what keeps the surface glued to the window - and
    /// the full pass waits for the settle, EXCEPT where eventfulness is the point: a hiding
    /// frame is the editor closing and the surface must go with it now, with no object-model
    /// work at all (lesson 27).
    /// </summary>
    private void OnFrameChanged()
    {
        if (_frame != 0 && !Win32.IsWindowVisible(_frame))
        {
            // The full pass takes its hide branch before any object-model call.
            RefreshSurfacePlacement();
            return;
        }

        PlaceSurfaceFast();
    }

    private void RefreshSurfacePlacement()
    {
        var placementStarted = Environment.TickCount64;
        if (_editorSurface is null || !_surfaceShown || _frame == 0 || _documentArea == 0)
        {
            Log.Verbose($"placement: skipped (surface {(_editorSurface is null ? "none" : "up")}, " +
                        $"shown {_surfaceShown}, frame {_frame:X}, documents {_documentArea:X})");
            return;
        }

        // The frame hiding is the editor's window closing: Alt+F4, its X, a shutdown's first
        // act. Placement work there is worse than pointless - reacting to the hide event with
        // the cutout pass put object-model calls INSIDE the editor's own close handling, and
        // the editor faulted under them, taking the host down (three crash records,
        // 2026-08-04: VBE7, ntdll, and this shim faulting by turn, each at a close; the
        // object-model pass of that day is the police pass of this one). A hidden
        // frame needs nothing covered; the surface hides with it, and the show event that
        // brings the frame back re-derives everything.
        // THE ONE PLACE THAT NOTICES THE EDITOR'S WINDOW COMING AND GOING, so it is where the
        // polling decision is re-taken. Only on the change: UpdatePolling restarts the timer, and
        // placement runs on every window event, every loader pulse and every pane follow - calling
        // it each time would reset a 1,000ms timer far more often than once a second and the tick
        // would never arrive at all.
        var frameVisible = Win32.IsWindowVisible(_frame);
        if (frameVisible != _frameVisible)
        {
            _frameVisible = frameVisible;
            UpdatePolling();
        }

        if (!frameVisible)
        {
            Log.Verbose("placement: the frame is not visible, hiding the surface with it");
            _editorSurface.Follow(default, visible: false);

            // The palette is owned, and owned windows do not follow a hidden owner. Left
            // alone it floats over the workbook after the editor closes and greets the next
            // Alt+F11 uninvited; it returns only when summoned.
            _browserPalette?.Hide();
            return;
        }

        var covering = CanCoverChrome();

        // Before anything paints this cycle: on the synchronous WM_SIZE route this runs ahead
        // of the native layout, so a hidden band never gets to draw at all.
        SetNativeChromeBands(visible: !covering);

        var bounds = SurfaceBounds(_frame, _documentArea, covering);
        Log.Verbose($"placement: {bounds.Width}x{bounds.Height} at {bounds.Left},{bounds.Top}, covering {covering}");

        _editorSurface.Follow(bounds, visible: true);
        _editorSurface.SetChrome(menuBar: covering);

        // The canvas is purely xlide (developer, 2026-08-05), no exceptions: every native
        // window is a panel, a ghost, a palette of our own, or policed away. The cutout-hole
        // machinery that once let the Object Browser live here is dormant in the overlay,
        // retired with the native Browser itself.
        if (covering)
        {
            PoliceNativeToolWindows();
        }

        // Only a pass that did the full object-model work counts; the early exits are free.
        PerfCounters.PlacementFull(Environment.TickCount64 - placementStarted);
    }

    /// <summary>
    /// Silences or restores the windows the native menu bar and toolbars live in, by region
    /// rather than by visibility.
    ///
    /// Covering them was not enough: the bands paint without clipping their siblings, so
    /// every resize stamped the native menu bar straight over the surface - "the native menu
    /// bar bleeds thru", 2026-08-04. Hiding their windows was not enough either: the editor's
    /// own layout shows them again on every resize, and the beat between its show and our
    /// next hide is the same flash. An EMPTY window region ends the argument: the window
    /// stays exactly as visible as the editor believes, its layout never changes, and it owns
    /// no pixels to paint with - Office never sets or resets regions on these, so nothing
    /// fights back. The commands lose nothing: menu reading and execution go through the
    /// object model, which has driven invisible bars since the Standard toolbar was retired.
    ///
    /// Restored whenever the surface is not covering the chrome, which is the loading phase:
    /// a stalled loader retreats below the native menu bar, and the bar it retreats below
    /// must be a real one.
    /// </summary>
    private void SetNativeChromeBands(bool visible)
    {
        if (_frame == 0)
        {
            return;
        }

        var docks = CodePaneTracker.FindChildrenByClass(_frame, "MsoCommandBarDock");
        var changed = 0;

        foreach (var dock in docks)
        {
            // Region complexity: 0 is no region set, 1 is an empty one. Checked first because
            // setting a region forces a repaint, and this runs on every placement pass.
            var probe = Win32.CreateRectRgn(0, 0, 0, 0);
            var kind = Win32.GetWindowRgn(dock, probe);
            Win32.DeleteObject(probe);

            if (visible)
            {
                if (kind != 0)
                {
                    _ = Win32.SetWindowRgn(dock, 0, true);
                    changed++;
                }
            }
            else if (kind != 1)
            {
                _ = Win32.SetWindowRgn(dock, Win32.CreateRectRgn(0, 0, 0, 0), false);
                changed++;
            }
        }

        if (changed > 0)
        {
            Log.Info($"chrome bands: {(visible ? "restored" : "silenced")} {changed} of {docks.Count} native band window(s)");
        }
    }

    /// <summary>
    /// Keeps the canvas purely xlide: no native tool window may stand docked and visible in
    /// the frame. This replaced the cutout-hole machinery (developer, 2026-08-05: "I'd like
    /// our canvas to be purely xlide") - holes meant tracking native windows pixel-for-pixel
    /// through a poll, and every window turned out to have a better home. The replaced
    /// windows (Project Explorer 6, Properties 7) are re-hidden if any native route re-shows
    /// them, and so is the Object Browser (2) now that its floating palette replaces it
    /// outright; the Locals and Watches ghosts (4, 3) are already exactly where they belong,
    /// visible to the object model and to nothing else. The Immediate window (5) stays
    /// hidden behind its mirror.
    /// </summary>
    private void PoliceNativeToolWindows()
    {
        var story = Log.VerboseEnabled ? new System.Text.StringBuilder() : null;

        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                var type = window?.GetInt32("Type") ?? -1;
                if (window is null || type is not (2 or 3 or 4 or 6 or 7))
                {
                    continue;
                }

                var caption = window.GetString("Caption");
                if (!window.GetBool("Visible"))
                {
                    continue;
                }

                if (type is 3 or 4)
                {
                    // The ghosts are visible to the object model on purpose, and their
                    // palettes are not frame children. A DOCKED one is a ghost that failed
                    // to prepare or was re-docked by hand; hiding it keeps the canvas pure,
                    // at the price of that panel sitting idle until the next session.
                    if (caption is { Length: > 0 } && CodePaneTracker.FindChildByCaption(caption) != 0)
                    {
                        window.SetBool("Visible", false);
                        story?.Append($"; hid docked '{caption}'");
                    }

                    continue;
                }

                window.SetBool("Visible", false);
                story?.Append($"; hid '{caption}'");
            }
        }
        catch (Exception ex)
        {
            Log.Verbose($"police: the editor's windows could not be read ({ex.GetType().Name})");
        }

        if (story is { Length: > 0 })
        {
            Log.Verbose($"police{story}");
        }
    }

    /// <summary>
    /// One window's client area expressed in another's client coordinates, which is the space a
    /// child of the second is positioned in.
    ///
    /// The window manager does the mapping. Working the origin out from window and client
    /// rectangles means assuming symmetric borders and that nothing but a caption and a menu sits
    /// above the client area, and each of those is wrong somewhere: maximised windows,
    /// right-to-left layouts, and per-monitor scaling break a different one. The arithmetic version
    /// of this put the surface a toolbar's height too high, which is how it came to cover the
    /// toolbar.
    /// </summary>
    private static unsafe PixelRect ClientAreaIn(nint window, nint target)
    {
        Rect client;
        if (!Win32.GetClientRect(window, &client))
        {
            return default;
        }

        var corners = stackalloc Point[2];
        corners[0] = new Point { X = client.Left, Y = client.Top };
        corners[1] = new Point { X = client.Right, Y = client.Bottom };

        // The call reports a failure and a legitimate zero shift identically, so the last error is
        // cleared first and consulted only when it returns zero.
        Marshal.SetLastSystemError(0);
        if (Win32.MapWindowPoints(window, target, corners, 2) == 0 && Marshal.GetLastSystemError() != 0)
        {
            return default;
        }

        // Normalised, because a right-to-left parent mirrors the mapping and swaps the corners.
        return new PixelRect(
            Math.Min(corners[0].X, corners[1].X),
            Math.Min(corners[0].Y, corners[1].Y),
            Math.Max(corners[0].X, corners[1].X),
            Math.Max(corners[0].Y, corners[1].Y));
    }

    /// <summary>
    /// Starts watching where the editor puts its code panes. Nothing is drawn over them yet; this
    /// establishes the map an editor surface will be positioned by, and proves it stays correct
    /// while the user rearranges the editor.
    /// </summary>
    private void TrackCodePanes()
    {
        try
        {
            _codePanes = new CodePaneTracker(_editor);
            _codePanes.Changed += panes =>
            {
                // A drag reports a new rectangle per mouse move, and following it is routine, not
                // an event: logging each wrote hundreds of identical lines in a second. Which
                // panes exist and whether they show is what changes rarely and reads as a story.
                var composition = string.Join("|",
                    panes.Select(pane => pane.Component + (pane.IsVisible ? string.Empty : "~")));

                if (composition != _lastPaneComposition)
                {
                    _lastPaneComposition = composition;
                    Log.Info($"code panes: {panes.Count} open");
                    foreach (var pane in panes)
                    {
                        Log.Info($"  {pane.Component} at {pane.Bounds.Left},{pane.Bounds.Top} " +
                                 $"{pane.Bounds.Width}x{pane.Bounds.Height}" + (pane.IsVisible ? string.Empty : " (hidden)"));
                    }
                }

                FollowActivePane(panes);
            };

            // While the editor is refusing refreshes, the window events that normally drive them
            // prove nothing. The poll timer becomes the way back: it keeps asking until the
            // editor answers, and the answer is what removes a closed tab from the strip.
            _codePanes.RefreshFailed += () =>
            {
                _pollsRemaining = Math.Max(_pollsRemaining, (int)(20_000 / DebugPollMilliseconds));
                UpdatePolling();
            };

            // The frame resizing is not a pane event. With no visible pane - the empty
            // workspace - nothing else hears it, and the surface sat at its old size while the
            // window grew around it. The frame's own events re-derive placement in every state.
            _codePanes.FrameChanged = OnFrameChanged;

            // The title bar is the last thing in the window still announcing the product this one
            // replaced, and the editor rewrites it as the active module changes, so it is retaken
            // on every rename of that window rather than set once.
            /*
             * ANY rename in the editor retakes the title, not only a rename OF the title's window.
             *
             * This used to fire only when the renamed window WAS the frame the chrome owns, which
             * reads as the careful thing to do and meant the caption was taken over once at
             * start-up and never again. Two measurements, on 2026-08-07, between them explain it:
             *
             *   - a real rename arrives as `rename event for F10AF8` while the chrome owns
             *     13209A6, so the equality never holds and Apply is never called;
             *   - and overwriting 13209A6's caption by hand produces NO event at all, so that
             *     window's own renames do not reach this hook in the first place.
             *
             * The frame is retitled at the same moments its neighbours are, so their events are a
             * perfectly good cue for asking whether ours still says what we put there. Apply is
             * built for exactly this: it reads the caption, compares it against the one it last
             * wrote, and writes nothing when they match, so being asked often costs a string
             * compare and answers almost always no.
             *
             * An event rather than a tick, deliberately. A poll was tried here and removed: it
             * stops when the editor is idle, so it did not fix the case it was added for, and a
             * caption that changes at a known moment belongs on that moment.
             */
            _codePanes.CaptionChanged = _ => _hostChrome?.Apply();

            // Any destroy might have been a hidden pane, which the tracker's own picture cannot
            // show (it only holds panes it can match - the active one, in practice). A moment of
            // polls re-reads the object model's open list and republishes; the page skips the
            // rebuild when nothing changed, so a dying tooltip costs a diff and no work.
            _codePanes.WindowDestroyed = () =>
            {
                _resyncPanePolls = Math.Max(_resyncPanePolls, 2);
                _pollsRemaining = Math.Max(_pollsRemaining, (int)(1_000 / DebugPollMilliseconds));
                UpdatePolling();
            };

            _codePanes.SurfaceStirred = CheckDesignerTabsForOutsideEdits;

            _codePanes.Start();
        }
        catch (Exception ex)
        {
            Log.Error("code panes: tracking could not be started", ex);
        }
    }

    public void Stop()
    {
        if (_stopped)
        {
            return;
        }

        _stopped = true;
        Log.Info("session stopping");

        // First out: no debug request may land on a session mid-teardown. The inside door goes
        // with it, and BEFORE the automation references below - it hangs on the add-in object,
        // and clearing that property is itself a COM call.
        RetireInsideDoor();
        _apiServer?.Dispose();
        _apiServer = null;

        // Order matters. Hooks and subclasses come out first, then windows, then automation
        // references, so nothing can call back into a half-released session.
        //
        // The engine goes before any of it. It is a separate process answering on another thread,
        // and letting it run on would mean a reply arriving after the objects meant to handle it
        // are gone. The wait is bounded because the host is shutting down and a hung engine must
        // not hold it there; the job object guarantees the process dies regardless.
        if (_analysis is not null)
        {
            var analysis = _analysis;
            _analysis = null;
            analysis.DisposeAsync().AsTask().Wait(TimeSpan.FromSeconds(3));
        }

        // Before anything is torn down, and before the engine is stopped: whatever the developer
        // typed last must reach the module, or closing the host loses it.
        _editorSurface?.FlushEdits();

        _frameSubclass?.Dispose();
        _frameSubclass = null;

        // The native menu and toolbar windows come back before the surface goes: a session
        // that stops - shutdown, disconnection, revival teardown - must leave the editor
        // whole, not silently menu-less.
        SetNativeChromeBands(visible: true);

        _immediateReader?.Dispose();
        _immediateReader = null;

        // The reading thread stops before the palettes change back: a reader must not touch a
        // window mid-restoration. Its join is bounded - see GhostReaderThread.Dispose.
        _ghostReaders?.Dispose();
        _ghostReaders = null;

        RestoreLocalsPalette();
        RestoreWatchPalette();

        _typeLibraries?.Dispose();
        _typeLibraries = null;

        // The palette owns a browser and a top-level window of its own; both go the same way
        // the surface's do, before the host can unload the library their code lives in.
        _browserPalette?.Dispose();
        _browserPalette = null;

        _codePanes?.Dispose();
        _codePanes = null;

        // The application reference is held between reads, so by the time a session stops there is
        // almost always one in hand: WorkbookSaved takes it on every poll tick. Until now it was
        // released only by the two catch blocks that drop a stale one, which is enough when the
        // process is dying and nothing when it is not - a watchdog revival and a non-shutdown
        // disconnection both stop a session inside a living Excel, and the wrapper left behind is
        // released by the finalizer thread instead. That is the FailFast this codebase has already
        // paid for: see DispatchObject.Dispose.
        _hostApp?.Dispose();
        _hostApp = null;

        // The window is the host's, and an add-in that unloads should leave its title bar saying
        // what it said before. After the tracker, so nothing is left to put our name back on.
        _hostChrome?.Dispose();
        _hostChrome = null;

        // Before the editor tears its own windows down. The surface owns a browser and a window
        // parented to the editor frame; leaving them for the host to destroy leaves browser
        // processes with no parent and a window procedure in a library about to be unloaded.
        _editorSurface?.Dispose();
        _editorSurface = null;

        Log.Info("session stopped");
    }

    /// <summary>
    /// Records what the add-in can see. This is the first proof that the object model is reachable,
    /// and it is the first thing to read in a support log when a load goes wrong.
    /// </summary>
    private void ReportEnvironment()
    {
        try
        {
            var version = _editor.GetString("Version");
            Log.Info($"editor version {version ?? "unknown"}");
        }
        catch (Exception ex)
        {
            Log.Error("could not read the editor version", ex);
        }

        try
        {
            using var host = _editor.GetObject("MainWindow");
            var caption = host?.GetString("Caption");
            Log.Info($"main window caption '{caption ?? "unknown"}'");
        }
        catch (Exception ex)
        {
            Log.Error("could not read the main window", ex);
        }
    }

    private void ReportOpenProjects()
    {
        try
        {
            using var projects = _editor.GetObject("VBProjects");
            if (projects is null)
            {
                Log.Warn("the editor exposed no project collection");
                return;
            }

            var count = projects.GetInt32("Count");
            Log.Info($"{count} project(s) loaded");

            for (var i = 1; i <= count; i++)
            {
                using var project = projects.GetItem(i);
                if (project is null)
                {
                    continue;
                }

                var name = project.GetString("Name");
                using var components = project.GetObject("VBComponents");
                var componentCount = components?.GetInt32("Count") ?? 0;
                Log.Info($"  project '{name}' with {componentCount} component(s)");
            }
        }
        catch (Exception ex)
        {
            Log.Error("could not enumerate projects", ex);
        }
    }

    public void Dispose()
    {
        Stop();

        // Reverse acquisition order: the tool window was obtained from the editor, so it goes first.
        _addIn?.Dispose();
        _editor.Dispose();
    }
}
