using System.Diagnostics;
using System.Text;
using Xlide.Vbe.Core.Changes;
using Xlide.Vbe.Core.Scm;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.Scm;
using Xlide.Vbe.Shim.Sync;

namespace Xlide.Vbe.Shim.AddIn;

/*
 * SOURCE CONTROL: git behind the folder a project's modules are exported to.
 *
 * ONE BRAIN, TWO DOORS. The pane's request and the api's `scm` route reach the same three steps:
 * GatherScm reads everything the object model and the change log must supply, ON the host
 * thread, into strings; RunScm runs git and computes the rows OFF it, on a pool thread, and
 * touches no COM at all; ScmApply writes modules, opens tabs and answers, back ON the host
 * thread. An action that changes the project - an import, a checkout - then re-gathers and runs
 * the status again, so what it answers is what the pane will draw next. The api door crosses to
 * the host only for the gather and the apply, because git can take seconds and the host lane's
 * budget is three; the pane's door is already on the host thread when its request arrives, so
 * it gathers inline, and answers through RunOnHostThread like the sync dialog does.
 *
 * Saving exports: a save made through this editor writes the folder in true-up mode, so the
 * folder always equals the last saved workbook. Commit applies the annotations, saves if dirty,
 * exports, and commits the ticked rows' files with `git commit --only`, so files somebody staged
 * outside are left alone. See docs/source-control.md for the design and the wire contract.
 */
internal sealed partial class AddInSession
{
    private static readonly TimeSpan GitDeadline = TimeSpan.FromSeconds(20);

    /// <summary>Fetch, pull and push talk to a network; a credential manager may open a browser.</summary>
    private static readonly TimeSpan GitRemoteDeadline = TimeSpan.FromSeconds(120);

    private const string HistoryFacePrefix = "history:";

    /// <summary>git.exe as last found, or null while it is absent.</summary>
    private GitClient? _gitClient;

    /// <summary>When git was last looked for and not found, so a session without it does not
    /// start a process per status.</summary>
    private long _gitMissedAtTicks;

    /// <summary>
    /// git.exe, found on the first use and LOOKED FOR AGAIN when it was not there: the noGit
    /// state tells the developer to install Git for Windows and press Refresh, and a lookup made
    /// once per session would make that sentence a lie until the next restart. A miss is
    /// remembered for five seconds, because each look starts a process. On the pool thread.
    /// </summary>
    private GitClient? GitNow()
    {
        if (_gitClient is { } found)
        {
            return found;
        }

        var now = Environment.TickCount64;
        if (now - Interlocked.Read(ref _gitMissedAtTicks) < 5000)
        {
            return null;
        }

        Interlocked.Exchange(ref _gitMissedAtTicks, now);
        _gitClient = GitClient.Locate();
        return _gitClient;
    }

    /// <summary>
    /// What this session knows per project under source control: the repository watcher and the
    /// time of the last commit made here, which is the boundary the suggested commit message
    /// reads the change log from. Keyed by project identity, pruned with the workbook.
    /// </summary>
    private readonly Dictionary<string, ScmProjectState> _scm = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>How often the folder or the repository moved on. Only ever compared by the pane.</summary>
    private int _scmStamp;

    /// <summary>
    /// The past-version tabs open on the strip, in opening order: a module's text at one commit,
    /// read-only. Product state like the designer tabs, host-listed so the strip, the `ui` route
    /// and the host agree, and keyed by PROJECT ID for the same reason those are.
    /// </summary>
    private readonly List<HistoryTab> _historyTabs = [];

    /// <summary>The past-version tab holding the active slot, when one is.</summary>
    private HistoryTab? _activeHistoryTab;

    private sealed class ScmProjectState : IDisposable
    {
        public ScmWatch? Watch { get; set; }

        public DateTimeOffset? LastCommitAt { get; set; }

        public void Dispose()
        {
            Watch?.Dispose();
            Watch = null;
        }
    }

    /// <summary>One past-version tab: which module, which project, which commit.</summary>
    private readonly record struct HistoryTab(string Module, string ProjectId, string Short, string Ref, string Text);

    /// <summary>Another open project's part in a repository question: it may share the root.</summary>
    private sealed record ScmOpenProject(string ProjectId, string Display, string Repository, bool Dirty);

    /// <summary>Carried across a re-gather: what the round before did, for the final answer.</summary>
    private sealed record ScmCarry(string Kind, string Detail, string[] Imported, ScmSkippedReply[] Skipped);

    /// <summary>Everything a request needs that had to be read on the host thread.</summary>
    private sealed record ScmGather(
        string Action,
        string ProjectId,
        string Display,
        bool Unsaved,
        string Folder,
        string SuggestedFolder,
        bool Dirty,
        bool InDesignMode,
        List<ScmLiveModule> Live,
        List<(string From, string To)> Renames,
        List<RoundSummary> Rounds,
        List<ScmOpenProject> Others,
        HashSet<string> UnwrittenBefore,
        string? Module,
        string? Ref,
        string? Message,
        string? Name,
        string? Email,
        int Limit,
        List<string> Named,
        string? LiveKind,
        string? LiveCode,
        bool LiveUnwritten,
        ScmCarry? Carry);

    private enum ScmFollowUp
    {
        None,
        Committed,
        Export,
        Import,
        Checkout,
        Restore,
        Open,
    }

    /// <summary>What the pool thread worked out, and what the host must still do with it.</summary>
    private sealed record ScmWork(
        ScmGather Gather,
        string Json,
        ScmFollowUp FollowUp,
        string? Root,
        List<string> Files,
        List<ScmOpenProject> Projects,
        ScmSkippedReply[] Skipped,
        string? Text,
        string? Kind,
        string? Hash,
        string Detail);

    /// <summary>The apply's outcome: the answer, or another round through the same three steps.</summary>
    private sealed record ScmStep(string? Json, ScmGather? Next);

    private static bool IsHistoryFace([System.Diagnostics.CodeAnalysis.NotNullWhen(true)] string? face) =>
        face is not null && face.StartsWith(HistoryFacePrefix, StringComparison.Ordinal);

    private static string HistoryFaceOf(string shortHash) => HistoryFacePrefix + shortHash;

    private static string ShortOf(string hash) => hash.Length > 7 ? hash[..7] : hash;

    private static string ScmError(string error) =>
        System.Text.Json.JsonSerializer.Serialize(new ScmErrorReply(error), ScmJsonContext.Default.ScmErrorReply);

    private static string ScmStatusJson(ScmStatusReply status) =>
        System.Text.Json.JsonSerializer.Serialize(status, ScmJsonContext.Default.ScmStatusReply);

    // ---------------------------------------------------------------- the pane's door

    /// <summary>
    /// The Source Control pane asking. On the host thread inside the browser callback, so the
    /// gather runs inline and git leaves for a pool thread at once; the answer comes back through
    /// RunOnHostThread, because posting to the page is the host thread's business.
    /// </summary>
    private void OnScmRequested(int requestId, IReadOnlyDictionary<string, string> arguments, string body)
    {
        ScmGather? gathered = null;
        string? refusal;
        try
        {
            gathered = GatherScm(arguments, body, null, out refusal);
        }
        catch (Exception ex)
        {
            Log.Error("scm: the pane's request could not be gathered", ex);
            refusal = ScmError(ex.Message.Trim());
        }

        if (gathered is null)
        {
            _editorSurface?.ShowScmResult(requestId, refusal ?? ScmError("nothing was gathered"));
            return;
        }

        DriveScmFromHost(gathered, json => _editorSurface?.ShowScmResult(requestId, json));
    }

    /// <summary>
    /// Runs the work on a pool thread and the apply back here, and again when the apply asks for
    /// another round - a checkout re-reads the project it just imported into before answering.
    /// </summary>
    private void DriveScmFromHost(ScmGather gathered, Action<string> answer)
    {
        var surface = _editorSurface;
        _ = Task.Run(() =>
        {
            var work = RunScm(gathered);
            var queued = surface?.RunOnHostThread(() =>
            {
                ScmStep step;
                try
                {
                    step = ScmApply(work);
                }
                catch (Exception ex)
                {
                    Log.Error("scm: the apply failed", ex);
                    step = new ScmStep(ScmError(ex.Message.Trim()), null);
                }

                if (step.Next is not null)
                {
                    DriveScmFromHost(step.Next, answer);
                    return;
                }

                answer(step.Json ?? ScmError("the request produced no answer"));
            });

            if (queued != true)
            {
                Log.Info("scm: the answer was dropped, the surface is gone");
            }
        });
    }

    // ---------------------------------------------------------------- the api's door

    /// <summary>
    /// The `scm` route, on the pool thread. The same three steps as the pane's door, with the
    /// two host halves crossed under the door's own budget and busy reply, and the apply
    /// attributed to `by=` so a restore records the caller the way every other write does.
    /// </summary>
    private string AnswerScmFromPool(IReadOnlyDictionary<string, string> query, string body)
    {
        var host = _editorSurface;
        if (host is null)
        {
            return ScmError("the surface is not up yet");
        }

        query.TryGetValue("by", out var by);

        ScmGather? gathered = null;
        string? refusal = null;
        var busy = CrossForScm(host, () => AttributedTo(by, () =>
        {
            gathered = GatherScm(query, body, null, out refusal);
            return 0;
        }));
        if (busy is not null)
        {
            return busy;
        }

        if (gathered is null)
        {
            return refusal ?? ScmError("nothing was gathered");
        }

        while (true)
        {
            var work = RunScm(gathered);
            ScmStep? step = null;
            busy = CrossForScm(host, () => step = AttributedTo(by, () => ScmApply(work)));
            if (busy is not null)
            {
                return busy;
            }

            if (step is null)
            {
                return ScmError("the apply produced no answer");
            }

            if (step.Next is null)
            {
                return step.Json ?? ScmError("the request produced no answer");
            }

            gathered = step.Next;
        }
    }

    /// <summary>
    /// One crossing to the host thread for the route, with the door's bookkeeping: the lane
    /// holder named while the work is on the thread, and the same busy reply every other route
    /// gives when the crossing does not answer in time. Null when the work ran.
    /// </summary>
    private string? CrossForScm(EditorSurface host, Action work)
    {
        var standingBefore = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);
        string? failure = null;

        using var crossing = CrossToHost(host, () =>
        {
            _laneHolder = "scm";
            System.Threading.Interlocked.Exchange(ref _laneHeldSince, Environment.TickCount64);
            try
            {
                work();
            }
            catch (Exception ex)
            {
                Log.Error("scm: the host half of the request failed", ex);
                failure = ScmError(ex.Message.Trim());
            }
            finally
            {
                _laneHolder = null;
            }
        });

        if (crossing.Answered)
        {
            return failure;
        }

        if (!crossing.Queued)
        {
            return ScmError("the editor surface has no window to run work on, so nothing was run - "
                + "the session is shutting down or the surface is being rebuilt");
        }

        var holder = _laneHolder;
        var blocked = AnswerBlockedRequest(
            standingBefore, crossing.Done, () => failure, keep: false, holder,
            holder is null ? 0 : Environment.TickCount64 - System.Threading.Interlocked.Read(ref _laneHeldSince));
        return Encoding.UTF8.GetString(blocked.Bytes);
    }

    // ---------------------------------------------------------------- step one: gather (host)

    /// <summary>
    /// Everything the request needs from the object model and the change log, read on the host
    /// thread. The actions that only change what is REMEMBERED - settings, forget, browse - act
    /// here, because remembering is a file write and needs no git. Commit's save and export
    /// happen here too: they are host work that must precede the git work, and nothing sits
    /// between them and the request. Null with a refusal when the request cannot proceed.
    /// </summary>
    private ScmGather? GatherScm(
        IReadOnlyDictionary<string, string> arguments, string body, ScmCarry? carry, out string? refusal)
    {
        refusal = null;
        arguments.TryGetValue("action", out var action);
        action = action is { Length: > 0 } ? action.Trim() : "status";
        var answersStatus = ActionAnswersStatus(action);

        arguments.TryGetValue("project", out var asked);
        var projectId = ResolveNamedProject(asked, out var complaint) ?? _shownProject;
        if (complaint is not null)
        {
            refusal = ScmError(complaint);
            return null;
        }

        if (string.IsNullOrEmpty(projectId))
        {
            if (!answersStatus)
            {
                refusal = ScmError("no project is shown, and none was named");
                return null;
            }

            return new ScmGather(
                action, string.Empty, string.Empty, true, string.Empty, string.Empty, false, false,
                [], [], [], [], new HashSet<string>(StringComparer.Ordinal), null, null, null, null, null,
                0, [], null, null, false, carry);
        }

        var display = DisplayFromProjectId(projectId) ?? projectId;
        var unsaved = !projectId.Contains('\\') && !projectId.Contains('/');

        var settings = LoadSyncSettings();
        var choice = settings.For(projectId);
        var folder = choice.Repository;
        var suggested = choice.Folder.Length > 0
            ? choice.Folder
            : unsaved
                ? string.Empty
                : Path.Combine(Path.GetDirectoryName(projectId) ?? string.Empty, Path.GetFileNameWithoutExtension(projectId));

        arguments.TryGetValue("folder", out var askedFolder);
        if (action == "browse")
        {
            // Modal to the editor's own window, like the sync dialog's chooser, and BLOCKING:
            // this is the pane's button, never a harness's.
            askedFolder = FolderPicker.Choose(
                CodePaneTracker.MainWindow(),
                "Keep this project's modules under source control in this folder",
                folder.Length > 0 ? folder : suggested);
            if (askedFolder is null)
            {
                action = "status";
            }
        }

        if (action is "settings" or "browse")
        {
            if (string.IsNullOrWhiteSpace(askedFolder))
            {
                refusal = ScmError("settings needs folder=<path>, the folder to keep this project's modules in");
                return null;
            }

            var full = Path.GetFullPath(askedFolder.Trim());

            // TWO OPEN PROJECTS NEVER SHARE A FOLDER: a true-up export of one would delete the
            // other's modules. Two projects can share a REPOSITORY, under two subfolders.
            foreach (var otherId in _projectNames.Keys)
            {
                if (string.Equals(otherId, projectId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var theirs = settings.For(otherId).Repository;
                if (theirs.Length > 0 && SamePath(theirs, full))
                {
                    refusal = ScmError($"{DisplayFromProjectId(otherId) ?? otherId} already keeps its modules in "
                        + $"{full}; two open projects cannot share a folder, because a true-up export of one "
                        + "would delete the other's modules. Choose a folder of its own, in the same repository if you like.");
                    return null;
                }
            }

            if (!SamePath(folder, full))
            {
                ScmStateFor(projectId).Dispose();
            }

            SaveSyncSettings(settings.With(projectId, choice with { Repository = full }));
            Log.Info($"scm: {display} keeps its modules in {full}");
            folder = full;
        }
        else if (action == "forget")
        {
            SaveSyncSettings(settings.With(projectId, choice with { Repository = string.Empty }));
            ScmStateFor(projectId).Dispose();
            Log.Info($"scm: {display} is no longer under source control here");
            folder = string.Empty;
        }

        arguments.TryGetValue("module", out var module);
        arguments.TryGetValue("ref", out var reference);
        arguments.TryGetValue("message", out var message);
        arguments.TryGetValue("name", out var name);
        arguments.TryGetValue("email", out var email);
        arguments.TryGetValue("limit", out var limitText);
        var limit = int.TryParse(limitText, out var parsedLimit) ? parsedLimit : 0;
        var named = body.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

        if (action == "commit" && string.IsNullOrWhiteSpace(message))
        {
            refusal = ScmError("commit needs message=<text>; a commit with no message is not one anybody can read later");
            return null;
        }

        // THE DIRTY SET IS TAKEN BEFORE THE FLUSH, as HandleSync takes it: the flush that lets
        // the rows read the typing clears the very flag an import refuses by.
        var unwrittenBefore = new HashSet<string>(
            (_editorSurface?.DocumentTable ?? [])
                .Where(doc => doc.Unwritten)
                .Select(doc => $"{(doc.Project ?? string.Empty).ToLowerInvariant()}\0{doc.Module.ToLowerInvariant()}"),
            StringComparer.Ordinal);

        if (action == "commit" && folder.Length > 0 && !unsaved)
        {
            // COMMIT IS SAVE, THEN EXPORT, THEN COMMIT. The annotations go first, as a save's
            // do; a dirty workbook is saved, and the save exports; a clean one is exported, so
            // the folder equals the workbook either way before git looks at it.
            if (_settings.ApplyAttributesOnSave)
            {
                ApplyAnnotationsBeforeSave(projectId);
            }

            if (WorkbookSaved(display) == false)
            {
                if (!SaveWorkbookOf(display))
                {
                    refusal = ScmError($"{display} could not be saved, so nothing was committed");
                    return null;
                }
            }
            else
            {
                ExportForSourceControl(projectId);
            }
        }

        var dirty = WorkbookSaved(display) == false;
        var inDesignMode = ProjectModeNow() == DesignMode;

        // The live modules, read AFTER anything above rewrote them, and after the typing has
        // reached them. The object model is apartment bound, so this walk is the reason the
        // gather runs here at all; everything after it is strings.
        var live = new List<ScmLiveModule>();
        if (folder.Length > 0 && !unsaved)
        {
            using var project = FindProjectByDisplayName(projectId);
            if (project is null)
            {
                refusal = ScmError($"the project could not be reached: nothing open is {display}");
                return null;
            }

            _editorSurface?.FlushEdits();
            using var components = project.GetObject("VBComponents");
            var count = components?.GetInt32("Count") ?? 0;
            for (var i = 1; i <= count; i++)
            {
                using var component = components!.GetItem(i);
                var componentName = component?.GetString("Name");
                if (component is null || string.IsNullOrEmpty(componentName))
                {
                    continue;
                }

                live.Add(new ScmLiveModule(
                    componentName,
                    ProjectReader.TypeName(component.GetInt32("Type")),
                    ProjectReader.ReadSource(component) ?? string.Empty));
            }
        }

        var log = ChangeLogFor(projectId);
        var state = ScmStateFor(projectId);
        var renames = new List<(string From, string To)>();
        var rounds = new List<RoundSummary>();
        if (log is not null)
        {
            foreach (var target in log.RestoreTargets(0, includeOpen: true))
            {
                if (!string.Equals(target.Module, target.NameNow, StringComparison.OrdinalIgnoreCase))
                {
                    renames.Add((target.Module, target.NameNow));
                }
            }

            foreach (var round in log.Rounds(int.MaxValue))
            {
                if (state.LastCommitAt is { } since && round.Ended <= since)
                {
                    continue;
                }

                rounds.Add(new RoundSummary(
                    round.Author,
                    round.Label,
                    [.. round.Entries.Select(entry => (entry.Kind.ToString(), entry.Module, entry.From))]));
            }
        }

        // The other open projects with a repository, for the two questions that cross workbooks:
        // a checkout imports into every project in the same root and refuses if any is dirty.
        var others = new List<ScmOpenProject>();
        foreach (var otherId in _projectNames.Keys)
        {
            if (string.Equals(otherId, projectId, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var theirs = settings.For(otherId).Repository;
            if (theirs.Length == 0)
            {
                continue;
            }

            var otherDisplay = DisplayFromProjectId(otherId) ?? otherId;
            others.Add(new ScmOpenProject(otherId, otherDisplay, theirs, WorkbookSaved(otherDisplay) == false));
        }

        var liveModule = module is { Length: > 0 }
            ? live.Find(one => string.Equals(one.Name, module, StringComparison.OrdinalIgnoreCase))
            : null;
        var liveUnwritten = module is { Length: > 0 }
            && (unwrittenBefore.Contains($"{display.ToLowerInvariant()}\0{module.ToLowerInvariant()}")
                || _editorSurface?.HasUnwritten(module, display) == true);

        return new ScmGather(
            action, projectId, display, unsaved, folder, suggested, dirty, inDesignMode,
            live, renames, rounds, others, unwrittenBefore,
            module, reference, message, name, email, limit, named,
            liveModule?.Kind, liveModule?.Code, liveUnwritten, carry);
    }

    private static bool ActionAnswersStatus(string action) => action is "status" or "settings" or "forget"
        or "browse" or "init" or "identity" or "export" or "open" or "checkout" or "fetch" or "pull"
        or "push" or "abort";

    private ScmProjectState ScmStateFor(string projectId)
    {
        if (!_scm.TryGetValue(projectId, out var state))
        {
            state = new ScmProjectState();
            _scm[projectId] = state;
        }

        return state;
    }

    private static bool SamePath(string left, string right)
    {
        if (left.Length == 0 || right.Length == 0)
        {
            return left.Length == 0 && right.Length == 0;
        }

        try
        {
            return string.Equals(
                Path.GetFullPath(left).TrimEnd('\\', '/'),
                Path.GetFullPath(right).TrimEnd('\\', '/'),
                StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex) when (ex is ArgumentException or IOException or NotSupportedException)
        {
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }

    // ---------------------------------------------------------------- step two: work (pool)

    /// <summary>
    /// A git repository as one request sees it: where it is, how the folder sits inside it, and
    /// who commits are signed by.
    /// </summary>
    private sealed record ScmRepo(GitClient Git, string Root, string Rel, ScmIdentityReply? Identity);

    private static GitResult Git(GitClient git, string directory, TimeSpan deadline, params string[] args) =>
        git.RunAsync(directory, args, deadline).GetAwaiter().GetResult();

    /// <summary>
    /// NO COM IN HERE, deliberately: everything it touches is the strings the gather read, the
    /// folder, and git.exe. That is what lets both doors run it off the host thread, and every
    /// git call it makes is bounded by a deadline.
    /// </summary>
    private ScmWork RunScm(ScmGather gathered)
    {
        try
        {
            return RunScmCore(gathered);
        }
        catch (Exception ex)
        {
            Log.Error($"scm: {gathered.Action} failed", ex);
            return Answer(gathered, ScmError($"{gathered.Action} failed: {ex.Message.Trim()}"));
        }
    }

    private static ScmWork Answer(ScmGather gathered, string json, string? root = null) =>
        new(gathered, json, ScmFollowUp.None, root, [], [], [], null, null, null, string.Empty);

    private ScmWork RunScmCore(ScmGather g)
    {
        var answersStatus = ActionAnswersStatus(g.Action);
        var git = GitNow();
        var version = git?.Version ?? string.Empty;

        // The states before a repository, in the order the pane draws them.
        if (git is null)
        {
            return answersStatus
                ? Answer(g, ScmStatusJson(BareStatus(g, "noGit", "git.exe was not found", version)))
                : Answer(g, ScmError("git.exe was not found on PATH or in Git for Windows' folders. Install Git "
                    + "for Windows (git-scm.com/download/win) and try again."));
        }

        if (g.ProjectId.Length == 0)
        {
            return Answer(g, ScmStatusJson(BareStatus(g, "noProject", "no project is shown, and none was named", version)));
        }

        if (g.Unsaved)
        {
            return answersStatus
                ? Answer(g, ScmStatusJson(BareStatus(g, "unsaved", "save the workbook first: an unsaved workbook has no path and no identity that survives a restart", version)))
                : Answer(g, ScmError($"{g.Display} has never been saved; save it first"));
        }

        if (g.Folder.Length == 0)
        {
            return answersStatus
                ? Answer(g, ScmStatusJson(BareStatus(g, "noFolder", "choose the folder to keep this project's modules in", version)))
                : Answer(g, ScmError($"{g.Display} is not under source control here; choose a folder first (scm?action=settings&folder=)"));
        }

        // Export needs a folder and nothing else: the folder is written whether or not a
        // repository stands above it, exactly as a save writes it.
        if (g.Action == "export")
        {
            return new ScmWork(g, string.Empty, ScmFollowUp.Export, RootOf(git, g.Folder), [], [], [], null, null, null, string.Empty);
        }

        var root = RootOf(git, g.Folder);
        var detail = g.Carry?.Detail ?? "read";

        if (g.Action == "init")
        {
            if (root is null)
            {
                detail = Initialise(git, g.Folder);
                root = RootOf(git, g.Folder);
            }
            else
            {
                detail = $"{g.Folder} is already inside the repository at {root}";
            }
        }

        if (root is null)
        {
            return answersStatus
                ? Answer(g, ScmStatusJson(BareStatus(g, "noRepository", "no repository holds this folder; Initialize runs git init in it", version)))
                : Answer(g, ScmError($"no repository holds {g.Folder}; initialise one first (scm?action=init)"));
        }

        var rel = Path.GetRelativePath(root, g.Folder).Replace('\\', '/');
        if (g.Action == "identity")
        {
            if (string.IsNullOrWhiteSpace(g.Name) || string.IsNullOrWhiteSpace(g.Email))
            {
                return Answer(g, ScmError("identity needs name= and email=, the name and address commits are signed with"));
            }

            var wroteName = Git(git, root, GitDeadline, "config", "user.name", g.Name.Trim());
            var wroteEmail = Git(git, root, GitDeadline, "config", "user.email", g.Email.Trim());
            if (!wroteName.Ok || !wroteEmail.Ok)
            {
                return Answer(g, ScmError($"the identity could not be written: {(wroteName.Ok ? wroteEmail : wroteName).Words}"));
            }

            detail = $"commits here are signed {g.Name.Trim()} <{g.Email.Trim()}>";
        }

        var repo = new ScmRepo(git, root, rel, ReadIdentity(git, root));

        switch (g.Action)
        {
            case "log":
                return Answer(g, LogReply(g, repo), root);

            case "diff":
                return Answer(g, DiffReply(g, repo, withText: false), root);

            case "show":
                return Answer(g, DiffReply(g, repo, withText: true), root);

            case "blame":
                return Answer(g, BlameReply(g, repo), root);

            case "restore":
                return RestoreWork(g, repo);

            case "fetch" or "pull" or "push":
            {
                var remotes = Git(git, root, GitDeadline, "remote");
                var remoteNames = remotes.StdOut.Split(
                    '\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (remoteNames.Length == 0)
                {
                    return Answer(g, ScmError($"{g.Action} needs a remote, and {root} has none configured; add one with "
                        + "`git remote add origin <url>`"), root);
                }

                // A FIRST PUSH SETS THE UPSTREAM ITSELF. Bare `git push` on a branch with no
                // upstream stops to ask for `--set-upstream`, which is the one thing a button
                // cannot answer: origin when there is one, else the only remote there is.
                string[] args;
                var hasUpstream = Git(git, root, GitDeadline, "rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}").Ok;
                if (g.Action == "push" && !hasUpstream)
                {
                    var remote = remoteNames.Contains("origin", StringComparer.Ordinal) ? "origin" : remoteNames[0];
                    args = ["push", "-u", remote, "HEAD"];
                }
                else
                {
                    args = [g.Action];
                }

                var ran = git.RunAsync(root, args, GitRemoteDeadline).GetAwaiter().GetResult();
                if (!ran.Ok)
                {
                    return Answer(g, ScmError($"{g.Action} failed: {ran.Words}"), root);
                }

                detail = ran.Words.Length > 0 ? ran.Words : $"{g.Action}: nothing to report";
                break;
            }

            case "abort":
            {
                var aborted = Git(git, root, GitDeadline, "merge", "--abort");
                if (!aborted.Ok)
                {
                    return Answer(g, ScmError($"abort failed: {aborted.Words}"), root);
                }

                detail = "the merge was aborted";
                break;
            }

            case "checkout":
                return CheckoutWork(g, repo);
        }

        var status = ReadStatus(g, repo, detail);

        switch (g.Action)
        {
            case "commit":
                return CommitWork(g, repo, status);

            case "import":
                return ImportWork(g, repo, status);

            case "open":
                return OpenWork(g, repo, status);
        }

        // A status after an import carries what the import did; after anything else it is the
        // status, with the detail of what happened.
        if (g.Carry is { Kind: "import" } imported)
        {
            return Answer(g, System.Text.Json.JsonSerializer.Serialize(
                new ScmImportReply(imported.Detail, imported.Imported, imported.Skipped, status),
                ScmJsonContext.Default.ScmImportReply), root);
        }

        return Answer(g, ScmStatusJson(status), root);
    }

    private static ScmStatusReply BareStatus(ScmGather g, string state, string detail, string version) =>
        new(detail, g.Display, g.ProjectId, state, g.Folder, string.Empty, version, string.Empty, string.Empty,
            0, 0, g.Dirty, null, [], [], [], [], null, string.Empty, g.SuggestedFolder, ScmWords.Covers);

    /// <summary>The repository root above a folder, or null when no repository holds it.</summary>
    private static string? RootOf(GitClient git, string folder)
    {
        if (!Directory.Exists(folder))
        {
            return null;
        }

        var answered = Git(git, folder, GitDeadline, "rev-parse", "--show-toplevel");
        if (!answered.Ok || answered.StdOut.Trim().Length == 0)
        {
            return null;
        }

        return Path.GetFullPath(answered.StdOut.Trim().Replace('/', '\\'));
    }

    private static string Initialise(GitClient git, string folder)
    {
        Directory.CreateDirectory(folder);
        var made = Git(git, folder, GitDeadline, "init");
        if (!made.Ok)
        {
            throw new InvalidOperationException($"git init failed: {made.Words}");
        }

        // A module round-trips byte for byte only with autocrlf off: the export writes CRLF and
        // the compare normalises, but a repository that rewrote endings on the way in would show
        // every file as changed to everybody else. An EXISTING repository's configuration is
        // never touched; this runs only in the one just made.
        Git(git, folder, GitDeadline, "config", "core.autocrlf", "false");

        var ignore = Path.Combine(folder, ".gitignore");
        if (!File.Exists(ignore))
        {
            File.WriteAllText(ignore, GitIgnore.Content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        return $"initialised a repository in {folder}";
    }

    private static ScmIdentityReply? ReadIdentity(GitClient git, string root)
    {
        var name = Git(git, root, GitDeadline, "config", "--get", "user.name");
        var email = Git(git, root, GitDeadline, "config", "--get", "user.email");
        return name.Ok && email.Ok && name.StdOut.Trim().Length > 0 && email.StdOut.Trim().Length > 0
            ? new ScmIdentityReply(name.StdOut.Trim(), email.StdOut.Trim())
            : null;
    }

    private static string RepoPath(string rel, string file) => rel == "." ? file : $"{rel}/{file}";

    /// <summary>Whether a root-relative path is a file directly inside the module folder.</summary>
    private static bool InFolder(string rel, string path)
    {
        var cut = path.LastIndexOf('/');
        var directory = cut < 0 ? "." : path[..cut];
        return string.Equals(directory, rel, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Every module file at HEAD directly under the folder: file name to text, header and all.</summary>
    private static Dictionary<string, string> HeadFiles(ScmRepo repo)
    {
        var files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var listed = Git(repo.Git, repo.Root, GitDeadline, "ls-tree", "-r", "--name-only", "HEAD", "--", repo.Rel);
        if (!listed.Ok)
        {
            // An unborn branch has no HEAD, and every module is then an added row.
            return files;
        }

        var wanted = new List<string>();
        foreach (var line in listed.StdOut.Split('\n'))
        {
            var path = line.TrimEnd('\r');
            if (path.Length == 0 || !InFolder(repo.Rel, path) || !ModuleSync.IsModuleFileName(Path.GetFileName(path)))
            {
                continue;
            }

            wanted.Add($"HEAD:{path}");
        }

        if (wanted.Count == 0)
        {
            return files;
        }

        var input = Encoding.UTF8.GetBytes(string.Join('\n', wanted) + "\n");
        var batch = repo.Git.RunAsync(repo.Root, ["cat-file", "--batch"], GitDeadline, input).GetAwaiter().GetResult();
        foreach (var (request, text) in GitCatFile.ParseBatch(batch.StdOutBytes, wanted))
        {
            // The request is "HEAD:<path>", and a colon is not a separator to Path.GetFileName,
            // so the prefix is cut first: keyed by the request whole, every module read as both
            // deleted and added after the first commit (the suite's first live run, 2026-09-08).
            var path = request[(request.IndexOf(':', StringComparison.Ordinal) + 1)..];
            files[Path.GetFileName(path)] = text;
        }

        return files;
    }

    private static Dictionary<string, string> FolderFiles(string folder)
    {
        var files = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var file in ModuleSyncService.ReadFolder(folder))
        {
            if (file.ReadError is null && ModuleSync.IsModuleFileName(file.FileName))
            {
                files[file.FileName] = file.Source;
            }
        }

        return files;
    }

    private static ScmLastCommitReply? LastCommit(ScmRepo repo)
    {
        var logged = Git(repo.Git, repo.Root, GitDeadline,
            "log", "-1", $"--format={GitArguments.LogFormat}", "--name-status", "-M", "--", repo.Rel);
        if (!logged.Ok)
        {
            return null;
        }

        var commits = GitLog.Parse(logged.StdOut);
        return commits.Count == 0
            ? null
            : new ScmLastCommitReply(commits[0].Hash, ShortOf(commits[0].Hash), commits[0].Author, commits[0].When, commits[0].Subject);
    }

    private static ScmRowReply RowOf(ScmRow row) => new(row.Module, row.Kind, row.File, row.Status, row.From);

    /// <summary>The whole picture: git's state, the rows, the folder, the history's head.</summary>
    private static ScmStatusReply ReadStatus(ScmGather g, ScmRepo repo, string detail)
    {
        var state = GitStatus.Parse(
            Git(repo.Git, repo.Root, GitDeadline, "status", "--porcelain=v2", "--branch", "-z", "--", repo.Rel).StdOut);
        var branches = GitBranches.Parse(
            Git(repo.Git, repo.Root, GitDeadline, "for-each-ref",
                "--format=%(refname:short)%09%(HEAD)%09%(upstream:short)", "refs/heads").StdOut);

        var rows = ScmRows.Compute(g.Live, HeadFiles(repo), g.Renames);
        var outside = ScmRows.Outside(g.Live, FolderFiles(g.Folder));
        var conflicts = state.Conflicts.Select(path => Path.GetFileName(path)).ToArray();

        var word = repo.Identity is null ? "noIdentity"
            : conflicts.Length > 0 ? "conflicted"
            : "ready";

        return new ScmStatusReply(
            detail, g.Display, g.ProjectId, word, g.Folder, repo.Root, repo.Git.Version,
            state.Head, state.Upstream ?? string.Empty, state.Ahead, state.Behind, g.Dirty, repo.Identity,
            [.. rows.Select(RowOf)],
            [.. outside.Select(RowOf)],
            [.. branches.Select(branch => new ScmBranchRow(branch.Name, branch.Current, branch.Upstream))],
            conflicts,
            LastCommit(repo),
            CommitMessage.Suggest(g.Rounds),
            g.SuggestedFolder,
            ScmWords.Covers);
    }

    // ---------------------------------------------------------------- the reads: log, diff, show, blame

    private static string StatusWord(string letter) => letter switch
    {
        "A" => "added",
        "M" => "modified",
        "D" => "deleted",
        "R" => "renamed",
        _ => letter,
    };

    private static bool IsModuleSideFile(string fileName) =>
        ModuleSync.IsModuleFileName(fileName) || ModuleSync.IsDesignFileName(fileName) || ModuleSync.IsSidecarFileName(fileName);

    private static string LogReply(ScmGather g, ScmRepo repo)
    {
        var limit = g.Limit > 0 ? g.Limit : 50;
        var args = new List<string>
        {
            "log", $"--format={GitArguments.LogFormat}", "--name-status", "-M", "-n", limit.ToString(System.Globalization.CultureInfo.InvariantCulture), "--",
        };

        if (g.Module is { Length: > 0 })
        {
            // The module's file, or every file it could be when it is not live any more.
            if (g.LiveKind is not null)
            {
                args.Add(RepoPath(repo.Rel, ModuleSync.FileNameFor(g.Module, g.LiveKind)));
            }
            else
            {
                args.Add(RepoPath(repo.Rel, $"{g.Module}.bas"));
                args.Add(RepoPath(repo.Rel, $"{g.Module}.cls"));
                args.Add(RepoPath(repo.Rel, $"{g.Module}.frm"));
            }
        }
        else
        {
            args.Add(repo.Rel);
        }

        var logged = repo.Git.RunAsync(repo.Root, args, GitDeadline).GetAwaiter().GetResult();
        if (!logged.Ok)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new ScmLogReply("no commits yet", []), ScmJsonContext.Default.ScmLogReply);
        }

        var commits = GitLog.Parse(logged.StdOut)
            .Select(commit => new ScmCommitRow(
                commit.Hash, ShortOf(commit.Hash), commit.Author, commit.Email, commit.When, commit.Subject, commit.Body,
                [.. commit.Files
                    .Where(file => InFolder(repo.Rel, file.Path) && IsModuleSideFile(Path.GetFileName(file.Path)))
                    .Select(file => new ScmFileRow(
                        ModuleSync.ModuleNameFromFileName(Path.GetFileName(file.Path)),
                        Path.GetFileName(file.Path),
                        StatusWord(file.Status)))]))
            .ToArray();

        return System.Text.Json.JsonSerializer.Serialize(
            new ScmLogReply($"{commits.Length} commit(s)", commits), ScmJsonContext.Default.ScmLogReply);
    }

    /// <summary>The text as the comparison sees it: header off, endings normalised, no trailing newline.</summary>
    private static string Code(string text) => ModuleSync.CodeWithoutHeader(text).TrimEnd('\n');

    /// <summary>
    /// A module's file at a ref: its name and text, or null when the ref holds none. The live
    /// kind says which extension to ask for; a module no longer live is looked for under all
    /// three.
    /// </summary>
    private static (string File, string Text)? FileAtRef(ScmGather g, ScmRepo repo, string module, string reference)
    {
        var candidates = g.LiveKind is not null
            ? [ModuleSync.FileNameFor(module, g.LiveKind)]
            : new[] { $"{module}.bas", $"{module}.cls", $"{module}.frm" };

        foreach (var file in candidates)
        {
            var shown = Git(repo.Git, repo.Root, GitDeadline, "show", $"{reference}:{RepoPath(repo.Rel, file)}");
            if (shown.Ok)
            {
                return (file, shown.StdOut);
            }
        }

        return null;
    }

    /// <summary>The full hash a ref names, or null when it names no commit.</summary>
    private static string? ResolveCommit(ScmRepo repo, string reference)
    {
        var resolved = Git(repo.Git, repo.Root, GitDeadline, "rev-parse", "--verify", "--quiet", $"{reference}^{{commit}}");
        return resolved.Ok && resolved.StdOut.Trim().Length > 0 ? resolved.StdOut.Trim() : null;
    }

    private static SyncDiffRow[] DiffRows(string left, string right) =>
        [.. ModuleSync.Condense(ModuleSync.Diff(left, right)).Select(SyncDiffRowFor)];

    private static string DiffReply(ScmGather g, ScmRepo repo, bool withText)
    {
        if (g.Module is not { Length: > 0 })
        {
            return ScmError($"{g.Action} needs module=<name>");
        }

        var reference = g.Ref is { Length: > 0 } ? g.Ref : "HEAD";
        var at = FileAtRef(g, repo, g.Module, reference);
        if (at is null && g.LiveCode is null)
        {
            return ScmError($"no module named {g.Module}, live or at {reference}");
        }

        if (at is null && withText)
        {
            return ScmError($"{g.Module} is not in {reference}");
        }

        var left = at is null ? string.Empty : Code(at.Value.Text);
        var right = g.LiveCode is null ? string.Empty : Code(g.LiveCode);
        var rows = DiffRows(left, right);
        var detail = at is null ? $"{g.Module} is not in {reference}; every line is live"
            : g.LiveCode is null ? $"{g.Module} is not in the project; every line is at {reference}"
            : $"{g.Module} at {reference} against the editor";

        return withText
            ? System.Text.Json.JsonSerializer.Serialize(
                new ScmTextReply(detail, g.Module, reference, left, rows), ScmJsonContext.Default.ScmTextReply)
            : System.Text.Json.JsonSerializer.Serialize(
                new ScmDiffReply(detail, g.Module, reference, rows), ScmJsonContext.Default.ScmDiffReply);
    }

    private static string BlameReply(ScmGather g, ScmRepo repo)
    {
        if (g.Module is not { Length: > 0 })
        {
            return ScmError("blame needs module=<name>");
        }

        if (g.LiveCode is null)
        {
            return ScmError($"no module named {g.Module} in {g.Display}");
        }

        var liveLines = ModuleSync.NormaliseEol(g.LiveCode).Split('\n').Length;
        var head = ResolveCommit(repo, "HEAD");
        var at = head is null ? null : FileAtRef(g, repo, g.Module, "HEAD");
        if (head is null || at is null)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new ScmBlameReply(
                    head is null ? "no commits yet" : $"{g.Module} is not committed yet",
                    g.Module, head ?? string.Empty, [], [.. Enumerable.Range(1, liveLines)]),
                ScmJsonContext.Default.ScmBlameReply);
        }

        var blamed = Git(repo.Git, repo.Root, GitDeadline, "blame", "--porcelain", "HEAD", "--", RepoPath(repo.Rel, at.Value.File));
        if (!blamed.Ok)
        {
            return ScmError($"blame failed: {blamed.Words}");
        }

        var map = GitBlame.MapToLive(GitBlame.ParsePorcelain(blamed.StdOut), at.Value.Text, g.LiveCode);
        return System.Text.Json.JsonSerializer.Serialize(
            new ScmBlameReply(
                $"{map.Lines.Count} committed line(s), {map.Uncommitted.Count} not",
                g.Module, head,
                [.. map.Lines.Select(line => new ScmBlameRow(line.Line, line.Hash, line.ShortHash, line.Author, line.When, line.Summary))],
                [.. map.Uncommitted]),
            ScmJsonContext.Default.ScmBlameReply);
    }

    // ---------------------------------------------------------------- the writes' pool halves

    private static ScmWork CommitWork(ScmGather g, ScmRepo repo, ScmStatusReply status)
    {
        if (repo.Identity is null)
        {
            return Answer(g, ScmStatusJson(status with { Detail = "set the name and email commits are signed with first" }), repo.Root);
        }

        if (status.Conflicts.Length > 0)
        {
            return Answer(g, ScmError("the folder has unresolved conflicts; resolve them, or Abort the merge, then commit"), repo.Root);
        }

        var rows = status.Rows;
        var chosen = new List<ScmRowReply>();
        var skipped = new List<ScmSkippedReply>();
        if (g.Named.Count == 0)
        {
            chosen.AddRange(rows);
        }
        else
        {
            foreach (var name in g.Named)
            {
                var row = Array.Find(rows, one => string.Equals(one.Module, name, StringComparison.OrdinalIgnoreCase));
                if (row is null)
                {
                    skipped.Add(new ScmSkippedReply(name, "no change against the branch head"));
                }
                else
                {
                    chosen.Add(row);
                }
            }
        }

        if (chosen.Count == 0)
        {
            return Answer(g, ScmError(rows.Length == 0
                ? "nothing to commit: every module matches the branch head"
                : "nothing to commit: none of the named modules has changed"), repo.Root);
        }

        // The files a row stands for: its code file, the earlier name's file for a rename, and a
        // form's design and sidecar beside its .frm. Only what exists on disk or in HEAD is
        // named to git; a path it knows nothing about is an error, not a no-op.
        var paths = new List<string>();
        void Name(string file)
        {
            var path = RepoPath(repo.Rel, file);
            if (!paths.Contains(path, StringComparer.OrdinalIgnoreCase)
                && (File.Exists(Path.Combine(g.Folder, file)) || HeadHas(repo, path)))
            {
                paths.Add(path);
            }
        }

        var headPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var listed = Git(repo.Git, repo.Root, GitDeadline, "ls-tree", "-r", "--name-only", "HEAD", "--", repo.Rel);
        if (listed.Ok)
        {
            foreach (var line in listed.StdOut.Split('\n'))
            {
                headPaths.Add(line.TrimEnd('\r'));
            }
        }

        bool HeadHas(ScmRepo _, string path) => headPaths.Contains(path);

        foreach (var row in chosen)
        {
            Name(row.File);
            if (row.From is { Length: > 0 })
            {
                Name(ModuleSync.FileNameFor(row.From, row.Kind));
            }

            if (string.Equals(row.Kind, "userform", StringComparison.OrdinalIgnoreCase))
            {
                Name(Path.ChangeExtension(row.File, ".frx"));
                Name(ModuleSync.DesignFileNameFor(row.Module));
            }
        }

        if (paths.Count == 0)
        {
            return Answer(g, ScmError("nothing to commit: none of the chosen rows has a file in the folder"), repo.Root);
        }

        var added = repo.Git.RunAsync(repo.Root, ["add", "-A", "--", .. paths], GitDeadline).GetAwaiter().GetResult();
        if (!added.Ok)
        {
            return Answer(g, ScmError($"git add failed: {added.Words}"), repo.Root);
        }

        var committed = repo.Git.RunAsync(repo.Root, ["commit", "--only", "-m", g.Message!.Trim(), "--", .. paths], GitDeadline)
            .GetAwaiter().GetResult();
        if (!committed.Ok)
        {
            return Answer(g, ScmError($"git commit failed: {committed.Words}"), repo.Root);
        }

        var hash = ResolveCommit(repo, "HEAD") ?? string.Empty;
        var after = ReadStatus(g, repo, $"committed {chosen.Count} module(s) as {ShortOf(hash)}");
        var json = System.Text.Json.JsonSerializer.Serialize(
            new ScmCommitReply(after.Detail, hash, ShortOf(hash), [.. chosen.Select(row => row.Module)], [.. skipped], after),
            ScmJsonContext.Default.ScmCommitReply);

        return new ScmWork(g, json, ScmFollowUp.Committed, repo.Root, [], [], [], null, null, hash, after.Detail);
    }

    private static ScmWork ImportWork(ScmGather g, ScmRepo repo, ScmStatusReply status)
    {
        if (!g.InDesignMode)
        {
            return Answer(g, ScmError("the project is stopped in the debugger; importing now would reset it. Press Reset, then import again"), repo.Root);
        }

        if (status.Conflicts.Length > 0)
        {
            return Answer(g, ScmError("the folder has unresolved conflicts; resolve them, or Abort the merge, then import"), repo.Root);
        }

        var candidates = status.Outside.Where(row => row.Status is "folderNewer" or "missingInProject").ToList();
        var files = new List<string>();
        var skipped = new List<ScmSkippedReply>();
        if (g.Named.Count == 0)
        {
            files.AddRange(candidates.Select(row => row.File));
        }
        else
        {
            foreach (var name in g.Named)
            {
                var row = candidates.Find(one => string.Equals(one.Module, name, StringComparison.OrdinalIgnoreCase));
                if (row is null)
                {
                    skipped.Add(new ScmSkippedReply(name,
                        status.Outside.Any(one => string.Equals(one.Module, name, StringComparison.OrdinalIgnoreCase))
                            ? "the folder has no file for it"
                            : "the folder holds the same text"));
                }
                else
                {
                    files.Add(row.File);
                }
            }
        }

        if (files.Count == 0 && skipped.Count == 0)
        {
            return Answer(g, ScmError("nothing to import: the folder matches the project"), repo.Root);
        }

        return new ScmWork(g, string.Empty, ScmFollowUp.Import, repo.Root, files, [], [.. skipped], null, null, null, string.Empty);
    }

    private static ScmWork OpenWork(ScmGather g, ScmRepo repo, ScmStatusReply status)
    {
        if (g.Module is not { Length: > 0 })
        {
            return Answer(g, ScmError("open needs module=<name> and ref=<commit>"), repo.Root);
        }

        var reference = g.Ref is { Length: > 0 } ? g.Ref : "HEAD";
        var hash = ResolveCommit(repo, reference);
        if (hash is null)
        {
            return Answer(g, ScmError($"{reference} names no commit in {repo.Root}"), repo.Root);
        }

        var at = FileAtRef(g, repo, g.Module, hash);
        if (at is null)
        {
            return Answer(g, ScmError($"{g.Module} is not in {reference}"), repo.Root);
        }

        var json = ScmStatusJson(status with { Detail = $"opened {g.Module} at {ShortOf(hash)}" });
        return new ScmWork(
            g, json, ScmFollowUp.Open, repo.Root, [], [], [], ModuleSync.BodyForImport(at.Value.Text), null, hash, reference);
    }

    private static ScmWork RestoreWork(ScmGather g, ScmRepo repo)
    {
        if (g.Module is not { Length: > 0 })
        {
            return Answer(g, ScmError("restore needs module=<name> and ref=<commit>"), repo.Root);
        }

        if (!g.InDesignMode)
        {
            return Answer(g, ScmError("the project is stopped in the debugger. Restoring now would reset it and lose "
                + "the run, so nothing was touched. Press Reset in the editor, or POST command?name=reset, and restore again"), repo.Root);
        }

        var reference = g.Ref is { Length: > 0 } ? g.Ref : "HEAD";
        var hash = ResolveCommit(repo, reference);
        if (hash is null)
        {
            return Answer(g, ScmError($"{reference} names no commit in {repo.Root}"), repo.Root);
        }

        var at = FileAtRef(g, repo, g.Module, hash);
        if (at is null)
        {
            return Answer(g, ScmError($"{g.Module} is not in {reference}"), repo.Root);
        }

        // The import's own body, one terminator short of the file, so a restored module has the
        // lines the commit's module had and not an empty one more.
        var kind = g.LiveKind ?? ModuleSync.ClassifyFile(at.Value.File, at.Value.Text);
        return new ScmWork(
            g, string.Empty, ScmFollowUp.Restore, repo.Root, [], [], [], ModuleSync.BodyForImport(at.Value.Text),
            kind, hash, reference);
    }

    private static ScmWork CheckoutWork(ScmGather g, ScmRepo repo)
    {
        if (g.Ref is not { Length: > 0 })
        {
            return Answer(g, ScmError("checkout needs ref=<branch or commit>"), repo.Root);
        }

        // Refused over unsaved work, the way git refuses a checkout over uncommitted work: the
        // checkout imports into every open project in this repository, and an import over a
        // dirty workbook would carry the folder over edits nobody has saved.
        if (g.Dirty)
        {
            return Answer(g, ScmError($"{g.Display} has unsaved changes; save it first, because the checkout imports the "
                + "branch's modules into it"), repo.Root);
        }

        var sharing = new List<ScmOpenProject>();
        foreach (var other in g.Others)
        {
            var theirRoot = RootOf(repo.Git, other.Repository);
            if (theirRoot is null || !SamePath(theirRoot, repo.Root))
            {
                continue;
            }

            if (other.Dirty)
            {
                return Answer(g, ScmError($"{other.Display} shares this repository and has unsaved changes; save it first, "
                    + "because the checkout imports the branch's modules into it too"), repo.Root);
            }

            sharing.Add(other);
        }

        var switched = Git(repo.Git, repo.Root, GitDeadline, "checkout", g.Ref.Trim());
        if (!switched.Ok)
        {
            return Answer(g, ScmError($"checkout failed: {switched.Words}"), repo.Root);
        }

        var projects = new List<ScmOpenProject> { new(g.ProjectId, g.Display, g.Folder, g.Dirty) };
        projects.AddRange(sharing);
        return new ScmWork(g, string.Empty, ScmFollowUp.Checkout, repo.Root, [], projects, [], null, null, null,
            switched.Words.Length > 0 ? switched.Words : $"checked out {g.Ref.Trim()}");
    }

    // ---------------------------------------------------------------- step three: apply (host)

    /// <summary>
    /// The host half after git: writes into the project, tabs, the watcher, and the answer. An
    /// action that changed the project asks for another round, so the status it answers is the
    /// one the pane will draw.
    /// </summary>
    private ScmStep ScmApply(ScmWork work)
    {
        var g = work.Gather;
        if (work.Root is not null && g.Folder.Length > 0 && g.ProjectId.Length > 0
            && LoadSyncSettings().For(g.ProjectId).Repository is { Length: > 0 } named
            && SamePath(named, g.Folder))
        {
            // Read again, not taken from the gather: a forget can land while the round is out with
            // git, and a watcher armed for the folder it gathered would hold that folder, and stamp
            // the page for it, until the workbook closed.
            EnsureScmWatch(g.ProjectId, work.Root, g.Folder);
        }

        switch (work.FollowUp)
        {
            case ScmFollowUp.Committed:
            {
                var state = ScmStateFor(g.ProjectId);
                state.LastCommitAt = DateTimeOffset.UtcNow;
                Log.Info($"scm: {g.Display} {work.Detail}");
                _editorSurface?.ShowScmStamp(++_scmStamp);
                return new ScmStep(work.Json, null);
            }

            case ScmFollowUp.Export:
            {
                var detail = ExportForSourceControl(g.ProjectId) ?? "nothing to export";
                return new ScmStep(null, ReGather(g, new ScmCarry("status", detail, [], [])));
            }

            case ScmFollowUp.Import:
            {
                var (imported, skipped) = ImportFromFolder(g.ProjectId, g.Display, g.Folder, work.Files, g.UnwrittenBefore);
                skipped.AddRange(work.Skipped);
                var detail = $"imported {imported.Count}, skipped {skipped.Count}";
                Log.Info($"scm: {g.Display} {detail}");
                _editorSurface?.ShowScmStamp(++_scmStamp);
                return new ScmStep(null, ReGather(g, new ScmCarry("import", detail, [.. imported], [.. skipped])));
            }

            case ScmFollowUp.Checkout:
            {
                var said = new StringBuilder(work.Detail);
                foreach (var project in work.Projects)
                {
                    var (imported, skipped) = ImportFromFolder(project.ProjectId, project.Display, project.Repository, null,
                        string.Equals(project.ProjectId, g.ProjectId, StringComparison.OrdinalIgnoreCase)
                            ? g.UnwrittenBefore
                            : new HashSet<string>(StringComparer.Ordinal));
                    said.Append("; imported ").Append(imported.Count).Append(" into ").Append(project.Display);
                    if (skipped.Count > 0)
                    {
                        said.Append(" (").Append(skipped.Count).Append(" skipped)");
                    }
                }

                Log.Info($"scm: {said}");
                _editorSurface?.ShowScmStamp(++_scmStamp);
                return new ScmStep(null, ReGather(g, new ScmCarry("status", said.ToString(), [], [])));
            }

            case ScmFollowUp.Restore:
                return new ScmStep(RestoreFromRef(g, work), null);

            case ScmFollowUp.Open:
                OpenHistoryTab(g, work);
                return new ScmStep(work.Json, null);

            default:
                return new ScmStep(work.Json, null);
        }
    }

    private ScmGather? ReGather(ScmGather g, ScmCarry carry)
    {
        var again = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["action"] = "status",
            ["project"] = g.ProjectId,
        };
        return GatherScm(again, string.Empty, carry, out _);
    }

    private void EnsureScmWatch(string projectId, string root, string folder)
    {
        var state = ScmStateFor(projectId);
        if (state.Watch is { } standing && SamePath(standing.Root, root) && SamePath(standing.Folder, folder))
        {
            return;
        }

        state.Watch?.Dispose();
        state.Watch = new ScmWatch(root, folder, () =>
        {
            // On a pool thread here; the stamp is the page's business, so it goes to the host.
            var surface = _editorSurface;
            surface?.RunOnHostThread(() => surface.ShowScmStamp(++_scmStamp));
        });
        Log.Info($"scm: watching {folder} and {root}\\.git");
    }

    /// <summary>Lets go of a closed workbook's watcher and commit time.</summary>
    private void ForgetScmState(IReadOnlyCollection<string> liveProjectIds)
    {
        foreach (var id in _scm.Keys.Where(id => !liveProjectIds.Contains(id)).ToList())
        {
            _scm[id].Dispose();
            _scm.Remove(id);
        }
    }

    private void DisposeScmWatches()
    {
        foreach (var state in _scm.Values)
        {
            state.Dispose();
        }
    }

    /// <summary>
    /// Writes the folder a project keeps under source control, in true-up mode, through the
    /// built-in planner and the same apply the sync dialog's Export runs. Called from both save
    /// paths and from the pane's Export; nothing to do, and nothing touched, when the project
    /// has no repository. Idempotent: an unchanged module is not written. Answers what it did,
    /// or null when there was no folder to write.
    /// </summary>
    private string? ExportForSourceControl(string? projectId)
    {
        if (string.IsNullOrEmpty(projectId))
        {
            return null;
        }

        var folder = LoadSyncSettings().For(projectId).Repository;
        if (folder.Length == 0)
        {
            return null;
        }

        var began = Stopwatch.GetTimestamp();
        try
        {
            using var project = FindProjectByDisplayName(projectId);
            if (project is null)
            {
                Log.Warn($"scm: {projectId} could not be reached for the export");
                return "the project could not be reached, so the folder was not written";
            }

            _editorSurface?.FlushEdits();
            var display = DisplayFromProjectId(projectId) ?? projectId;
            Directory.CreateDirectory(folder);
            var live = ModuleSyncService.ReadLiveModules(project, _controlDefaults);
            var plan = ModuleSync.PlanExport(projectId, display, folder, live, ModuleSyncService.ReadFolder(folder), ExportMode.TrueUp);
            var chosen = new HashSet<string>(plan.Items.Where(item => item.Checked).Select(item => item.Id), StringComparer.Ordinal);
            var applied = ModuleSyncService.Apply(
                project, plan, chosen,
                (component, text, owner) => WriteModule(component, text, owner, hostRewrite: true, keepEveryCharacter: true));

            var detail = $"exported {display}: {applied.Changed.Count} written, {applied.Removed.Count} removed, "
                + $"{applied.Failed.Count} failed";
            Log.Info($"scm: {detail} in {(long)Stopwatch.GetElapsedTime(began).TotalMilliseconds}ms"
                + (applied.Failed.Count > 0 ? $" ({string.Join("; ", applied.Failed)})" : string.Empty));
            _editorSurface?.ShowScmStamp(++_scmStamp);
            return detail;
        }
        catch (Exception ex)
        {
            Log.Error($"scm: the export of {projectId} failed", ex);
            return $"the export failed: {ex.Message.Trim()}";
        }
    }

    /// <summary>
    /// Reads files from the folder into a project through the same planner and apply the sync
    /// dialog's Import runs, with its unwritten-edits guard, then tells the tree, the engine
    /// and the attributes, exactly as HandleSync does. Null files means every row the plan ticks.
    /// </summary>
    private (List<string> Imported, List<ScmSkippedReply> Skipped) ImportFromFolder(
        string projectId, string display, string folder, List<string>? files, HashSet<string> unwrittenBefore)
    {
        var imported = new List<string>();
        var skipped = new List<ScmSkippedReply>();

        using var project = FindProjectByDisplayName(projectId);
        if (project is null)
        {
            skipped.Add(new ScmSkippedReply(display, "the project could not be reached"));
            return (imported, skipped);
        }

        _editorSurface?.FlushEdits();
        var live = ModuleSyncService.ReadLiveModules(project, _controlDefaults);
        var plan = ModuleSync.PlanImport(projectId, display, folder, live, ModuleSyncService.ReadFolder(folder), ImportMode.UpdateOnly);

        // Chosen by FILE NAME, never by composing an id: the built-in planner's ids are its own
        // business. A form's design row rides with its code file, by module name.
        HashSet<string> chosen;
        if (files is null)
        {
            chosen = new HashSet<string>(plan.Items.Where(item => item.Checked).Select(item => item.Id), StringComparer.Ordinal);
        }
        else
        {
            var wantedFiles = new HashSet<string>(files, StringComparer.OrdinalIgnoreCase);
            var wantedModules = new HashSet<string>(files.Select(ModuleSync.ModuleNameFromFileName), StringComparer.OrdinalIgnoreCase);
            chosen = new HashSet<string>(
                plan.Items
                    .Where(item => item.Status != SyncStatus.Unchanged
                        && (wantedFiles.Contains(item.FileName) || (item.IsDesign && wantedModules.Contains(item.ModuleName))))
                    .Select(item => item.Id),
                StringComparer.Ordinal);
        }

        var applied = ModuleSyncService.Apply(
            project, plan, chosen,
            (component, text, owner) =>
            {
                var guardKey = $"{(DisplayFromProjectId(owner) ?? string.Empty).ToLowerInvariant()}\0{component.ToLowerInvariant()}";
                if (unwrittenBefore.Contains(guardKey)
                    || _editorSurface?.HasUnwritten(component, DisplayFromProjectId(owner)) == true)
                {
                    return $"{component} was not imported: it held edits you had not written yet, and importing "
                        + "would have replaced them. They are written to the module now - review them, then import again.";
                }

                return WriteModule(component, text, owner, hostRewrite: true, keepEveryCharacter: true);
            });

        imported.AddRange(applied.Changed);
        foreach (var why in applied.Failed.Concat(applied.Skipped))
        {
            var module = files?.Select(ModuleSync.ModuleNameFromFileName)
                .FirstOrDefault(name => why.StartsWith(name, StringComparison.OrdinalIgnoreCase)) ?? string.Empty;
            skipped.Add(new ScmSkippedReply(module, why));
        }

        if (applied.Changed.Count > 0 || applied.Removed.Count > 0)
        {
            PublishProjects();
            _analysis?.Reanalyse();
            ApplyAnnotationsAfterImport(projectId, applied.Changed);
        }

        Log.Info($"scm: import into {display}: {applied.Changed.Count} changed, {applied.Skipped.Count} skipped, "
            + $"{applied.Removed.Count} removed, {applied.Failed.Count} failed");
        return (imported, skipped);
    }

    /// <summary>
    /// Writes a module's text at a commit into the project through the ordinary module write,
    /// bracketed the way the change log's own restore is, so it lands as a round labelled with
    /// the commit and can itself be restored away. A module the project lacks is re-added when
    /// it is a standard or class module; a form's design is not in the file's code, and a
    /// document cannot be added, so those are refused in words.
    /// </summary>
    private string RestoreFromRef(ScmGather g, ScmWork work)
    {
        var module = g.Module!;
        var text = work.Text ?? string.Empty;
        var shortHash = ShortOf(work.Hash ?? string.Empty);
        var display = DisplayFromProjectId(g.ProjectId);
        var log = ChangeLogFor(g.ProjectId);
        var now = DateTimeOffset.UtcNow;

        // The running round closes first, so the restore lands as a round of its own rather
        // than folding into somebody's edits.
        log?.Close(null, now);

        string did;
        string? why = null;
        if (g.LiveUnwritten || _editorSurface?.HasUnwritten(module, display) == true)
        {
            did = "skipped";
            why = "It holds edits you have not written yet. Save or discard them, then restore again.";
        }
        else
        {
            using var standing = FindComponent(module, g.ProjectId, out _);
            if (standing is null)
            {
                var kind = work.Kind ?? "standard";
                var addKind = kind.ToLowerInvariant() switch
                {
                    "class" => 2,
                    "standard" => 1,
                    _ => 0,
                };

                if (addKind == 0)
                {
                    did = "skipped";
                    why = $"{module} is a {kind} and is not in the project; a {kind} cannot be re-added from its code alone.";
                }
                else if (AddComponentCore(addKind, module, display, g.ProjectId, out _) is { } refused)
                {
                    did = "failed";
                    why = refused;
                }
                else
                {
                    var filled = WriteModule(module, text, g.ProjectId, hostRewrite: true);
                    did = filled is null ? "added" : "failed";
                    why = filled;
                }
            }
            else if (ModuleSync.SameText(Code(ProjectReader.ReadSource(standing) ?? string.Empty), Code(text)))
            {
                did = "unchanged";
                why = $"Already holds its text from {shortHash}.";
            }
            else
            {
                var wrote = WriteModule(module, text, g.ProjectId, hostRewrite: true);
                did = wrote is null ? "written" : "failed";
                why = wrote;
            }
        }

        log?.Close($"restore {module} from {shortHash}", DateTimeOffset.UtcNow);
        var detail = $"{module} from {shortHash}: {did}{(why is null ? string.Empty : $" - {why}")}";
        Log.Info($"scm: restore {detail}");
        _editorSurface?.ShowChangesStamp(++_changeStamp);
        _editorSurface?.ShowScmStamp(++_scmStamp);

        return System.Text.Json.JsonSerializer.Serialize(
            new ScmRestoreReply(detail, module, work.Detail, did, why), ScmJsonContext.Default.ScmRestoreReply);
    }

    // ---------------------------------------------------------------- the past-version tabs

    private void OpenHistoryTab(ScmGather g, ScmWork work)
    {
        var shortHash = ShortOf(work.Hash ?? string.Empty);
        var module = g.Live.Find(one => string.Equals(one.Name, g.Module, StringComparison.OrdinalIgnoreCase))?.Name ?? g.Module!;
        var index = _historyTabs.FindIndex(tab => SameHistoryTab(tab, module, g.ProjectId, shortHash));
        if (index < 0)
        {
            _historyTabs.Add(new HistoryTab(module, g.ProjectId, shortHash, work.Detail, work.Text ?? string.Empty));
            index = _historyTabs.Count - 1;
            Log.Info($"history tab: opened {module} @ {shortHash}");
        }

        _activeHistoryTab = _historyTabs[index];
        _activeDesignerTab = null;
        PublishModules();
    }

    private static bool SameHistoryTab(HistoryTab tab, string module, string projectId, string shortHash) =>
        string.Equals(tab.Module, module, StringComparison.OrdinalIgnoreCase)
        && string.Equals(tab.ProjectId, projectId, StringComparison.OrdinalIgnoreCase)
        && string.Equals(tab.Short, shortHash, StringComparison.OrdinalIgnoreCase);

    /// <summary>The page clicked a past-version tab: it takes the active slot, and nothing moves underneath.</summary>
    private void ActivateHistoryTab(string module, string? projectDisplay, string face)
    {
        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject ?? string.Empty;
        var shortHash = face[HistoryFacePrefix.Length..];
        var index = _historyTabs.FindIndex(tab => SameHistoryTab(tab, module, projectId, shortHash));
        if (index < 0)
        {
            Log.Info($"history tab: nothing open for {module} @ {shortHash}");
            return;
        }

        _activeHistoryTab = _historyTabs[index];
        _activeDesignerTab = null;
        PublishModules();
    }

    /// <summary>Closes a past-version tab; closing what is not open changes nothing.</summary>
    private void CloseHistoryTab(string module, string? projectDisplay, string face)
    {
        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject ?? string.Empty;
        var shortHash = face[HistoryFacePrefix.Length..];
        var removed = _historyTabs.RemoveAll(tab => SameHistoryTab(tab, module, projectId, shortHash));

        if (_activeHistoryTab is { } shown && SameHistoryTab(shown, module, projectId, shortHash))
        {
            _activeHistoryTab = null;
        }

        if (removed > 0)
        {
            Log.Info($"history tab: closed {module} @ {shortHash}");
            PublishModules();
        }
    }
}
