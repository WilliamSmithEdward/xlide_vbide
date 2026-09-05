using System.Text;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The test runner's session half: one brain for the Tests pane and the xlide api, so the two
/// doors cannot disagree about what a run is. Discovery, the support-module gate, the run loop
/// and its per-test streaming all live in <see cref="TestRunService"/>; this partial owns the
/// standing state - the latest outcome per test, whether a run is in flight - and the message
/// that repaints the pane after every change.
///
/// EVERY OPEN FILE, not just the active one. A session holds as many VBA projects as the host
/// has files open, each with its own tests and its own copy of the support module, and a runner
/// that could only see the front one made the other file's tests invisible rather than absent.
/// A test is therefore addressed by the pair (file, id): a module name is not an identity across
/// files, so two open workbooks may each hold an InvoiceTests.Adds and they are not the same test.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>One open file's whole test picture: its identity, its support state, its tests.</summary>
    private sealed record TestFile(
        string ProjectId,
        string Name,
        string Support,
        List<TestRunService.TestCase> Tests);

    /// <summary>The latest outcome per (file, test), kept across runs so the pane remembers.</summary>
    private readonly Dictionary<string, TestRunService.TestResult> _testOutcomes = new(StringComparer.OrdinalIgnoreCase);

    private bool _testsRunning;
    private string? _testsCurrent;
    private string? _testsCurrentProject;

    /// <summary>
    /// When the last run finished. The pane says it out loud, because a tally with no clock
    /// beside it cannot tell a result that just landed from one left over from an hour ago -
    /// and a rerun that changes nothing looks identical to a rerun that never happened.
    /// </summary>
    private DateTimeOffset? _testsRanAt;

    /// <summary>
    /// While a run is in flight: the files the run started with. Every landing result repaints
    /// the pane, and a repaint that re-walked every module of every project over COM turned an
    /// N-test run into N whole-session reads - measured as the dominant cost of a run the day
    /// after the runner shipped. Patched statuses over this instead; the run's last repaint
    /// clears it and walks fresh, because a test is allowed to have edited modules.
    /// </summary>
    private List<TestFile>? _testsLive;

    /// <summary>
    /// What each project last discovered, so an analysis snapshot for one file can repaint the
    /// pane without re-reading the others. Refreshed wholesale by <see cref="ReadTestFiles"/>,
    /// which is also what forgets a file that has been closed.
    ///
    /// EVERY OPEN FILE IS IN HERE, including one with no tests and no support module. What the
    /// pane paints is decided by <see cref="Shown"/> instead - see the three symptoms written
    /// there for why this must not be the place a file is left out.
    /// </summary>
    private readonly Dictionary<string, TestFile> _testsSeen = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The key an outcome is filed under: the file's identity, then the test's.</summary>
    private static string TestKey(string projectId, string id) => $"{projectId}\0{id}";

    /// <summary>
    /// The pane's whole picture right now: every open file's tests merged with their latest
    /// outcomes. Discovery walks the live projects on every call - tests are code, and code
    /// moves under the pane - except mid-run, where the run's own snapshot stands in.
    /// </summary>
    /// <summary>True once a walk has covered every open file, so the cache can be trusted.</summary>
    private bool _testsWalked;

    internal SetTestsMessage TestsSnapshot() => ComposeTests(_testsLive ?? Shown(ReadTestFiles()));

    /// <summary>
    /// The same picture WITHOUT re-reading the project, when this session has already walked it
    /// once: the analysis pass hands every file's text to <see cref="OnAnalysisSnapshot"/> as it
    /// moves, so the cache is as fresh as the last pass - which is the same freshness the pane
    /// has always run on. A walk of a large project is a third of a second of COM on the host
    /// thread (81,795 lines, measured 2026-08-20), and showing a pane should not cost that.
    /// </summary>
    private SetTestsMessage TestsSnapshotCached()
    {
        if (_testsLive is { } live)
        {
            return ComposeTests(live);
        }

        if (!_testsWalked)
        {
            return ComposeTests(ReadTestFiles());
        }

        List<TestFile> known;
        lock (_testsSeen)
        {
            known = Shown([.. _testsSeen.Values]);
        }

        return ComposeTests(known);
    }

    /// <summary>
    /// Walks every open project: its standard modules once, its tests out of that text, and its
    /// support state. EVERY open file comes back, including one with nothing to say - the pane
    /// leaves those out through <see cref="Shown"/>, and the actions below want them all: a file
    /// named to an action has to be findable, and a file the walk has covered has to be known
    /// afterwards or the reconcile asks for the walk again on every tree publish.
    /// </summary>
    private List<TestFile> ReadTestFiles()
    {
        var files = new List<TestFile>();

        // The editor's own idea of active, read here because this is the host thread. It stands
        // in for the shown project in the beat before a module has been adopted, which is the
        // one moment a pane can be painted with nothing shown.
        var activeId = ActiveProjectId();
        _walkedActiveProjectId = activeId;

        using var projects = _editor.GetObject("VBProjects");
        var count = projects?.GetInt32("Count") ?? 0;
        for (var i = 1; i <= count; i++)
        {
            using var project = projects!.GetItem(i);
            if (project is null)
            {
                continue;
            }

            try
            {
                var (id, own) = ProjectReader.Identity(project);

                // THE NAME THE TREE USES, not the project's own. Two workbooks that have never
                // been saved are both called "VBAProject", so the tree numbers them - and this
                // pane calling both of them "VBAProject" made two files one indistinguishable
                // option, collapsed their scope keys into one, and left Current Module unable
                // to match the tab it was following (measured 2026-08-20 with two unsaved
                // books open beside the fixtures). One name per file across the whole surface.
                // The tree's name, and ONLY a name the tree knows. Closing a workbook leaves
                // the editor holding its project for a beat while it is torn down, and that
                // ghost is briefly the ACTIVE one - which let it into the pane, where both
                // panes flashed a file called "VBAProject" that nobody had opened (measured
                // 2026-08-20 while closing the twin). A project the tree has never published
                // is one the developer cannot see anyway; the gate on it now lives in
                // <see cref="Shown"/>, with the active-file rule it belongs to.
                var display = DisplayFromProjectId(id) ?? own;
                var tests = TestRunService.DiscoverFrom(TestRunService.ReadStandardModules(project));
                var support = TestRunService.SupportState(project);
                files.Add(new TestFile(id, display, support, tests));
            }
            catch (Exception ex)
            {
                // A project that will not answer is one file's problem, not the pane's.
                Log.Warn($"tests: a project could not be read for discovery, {ex.Message}");
            }
        }

        var ordered = Order(files, PaneActiveProjectId());
        _testsWalked = true;
        lock (_testsSeen)
        {
            _testsSeen.Clear();
            foreach (var file in ordered)
            {
                _testsSeen[file.ProjectId] = file;
            }
        }

        return ordered;
    }

    /// <summary>The developer's own file first, then the rest by name: the pane reads top-down.</summary>
    private static List<TestFile> Order(IEnumerable<TestFile> files, string? activeId) =>
    [
        .. files
            .OrderByDescending(file => string.Equals(file.ProjectId, activeId, StringComparison.OrdinalIgnoreCase))
            .ThenBy(file => file.Name, StringComparer.OrdinalIgnoreCase),
    ];

    /// <summary>
    /// The files the pane lists, in the order it reads them: every file the tree has published,
    /// the developer's own first - because the file select has to be able to offer all of them,
    /// and the install chip has to be able to land on any one.
    ///
    /// STORAGE HOLDS EVERY OPEN FILE AND ONLY THIS LEAVES ONE OUT. The rule used to live in the
    /// walk, which kept a file with nothing to say only while it was active, while the analysis
    /// pass - the other writer of the same cache - dropped it outright. Two writers, two rules,
    /// three symptoms (all measured 2026-08-21 with a 7-module, 81,795-line file open):
    ///
    ///   - opening a file with no tests offered NO install: the pane had never heard of it, so
    ///     the chip spoke for some other file and reported "installed" for a file that had
    ///     nothing installed;
    ///   - a file listed by a walk VANISHED the moment its text moved, because the pass that
    ///     read the change dropped it on the way past;
    ///   - the reconcile asks whether every live file is known, and for such a file the answer
    ///     could never become yes, so EVERY tree publish walked the whole session again: 95ms
    ///     per module add with every file known against 228-412ms with one that could not be.
    ///
    /// The tree's name is still the gate on the active file, for the reason the walk gives
    /// above: a workbook being torn down is briefly the active project, and a file the tree has
    /// never published is one the developer cannot see.
    /// </summary>
    private List<TestFile> Shown(IEnumerable<TestFile> files)
    {
        var active = PaneActiveProjectId();

        // EVERY FILE THE TREE KNOWS, now that the pane has a file select on it.
        //
        // This used to keep a file only while it held tests, carried the support module, or was
        // the one being worked in. That was right for a list of ROWS and wrong for a list of
        // FILES: a workbook with no tests and no XlideAssert vanished from the select the moment
        // the developer looked at something else, so the one file a developer most needs to point
        // the install chip at was the one file they could not choose (the owner, 2026-08-21).
        //
        // The tree's name is still the gate, for the reason the walk gives: a workbook being torn
        // down is briefly the active project, and a file the tree has never published is one the
        // developer cannot see. A file with nothing to say now says so, with a count of zero.
        return Order(files.Where(file => DisplayFromProjectId(file.ProjectId) is not null), active);
    }

    private string? ActiveProjectId()
    {
        try
        {
            using var active = _editor.GetObject("ActiveVBProject");
            return active is null ? null : ProjectReader.Identity(active).Id;
        }
        catch (Exception)
        {
            return null;
        }
    }

    private SetTestsMessage ComposeTests(IReadOnlyList<TestFile> files)
    {
        var rows = new List<TestRowMessage>();
        var standings = new TestFileMessage[files.Count];
        for (var at = 0; at < files.Count; at++)
        {
            var file = files[at];
            standings[at] = new TestFileMessage(file.Name, file.Support, file.Tests.Count);
            foreach (var test in file.Tests)
            {
                var outcome = _testOutcomes.GetValueOrDefault(TestKey(file.ProjectId, test.Id));
                var running = _testsRunning
                    && string.Equals(_testsCurrentProject, file.ProjectId, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(_testsCurrent, test.Id, StringComparison.OrdinalIgnoreCase);
                rows.Add(new TestRowMessage(
                    test.Id, file.Name, test.Module, test.Procedure, test.Line,
                    running ? "running" : outcome?.Status ?? (test.SkipReason is not null ? "skip-marked" : "none"),
                    outcome?.Message, outcome?.DurationMs ?? 0, outcome?.Output ?? [],
                    test.Tags, test.Owner, test.Requirement, test.TimeoutMs, test.ExpectedError));
            }
        }

        return new SetTestsMessage(
            "setTests", SummarySupport(files), _testsRunning, _testsCurrent,
            _testsRanAt?.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
            standings, [.. rows]);
    }

    /// <summary>
    /// The session's one-word support standing, worst first, across EVERY file the pane lists.
    ///
    /// It used to count only files that hold tests, on the reasoning that a file with no tests
    /// needs no XlideAssert and should not drag the answer red. That reads well and it lies: with
    /// nothing open holding tests the condition is vacuously satisfied, so the chip went green and
    /// its tooltip said "every open file that holds tests carries an XlideAssert matching this
    /// product" - true, and about no files at all. The owner reported it as a chip that stayed
    /// green with a workbook selected that plainly did not have it (2026-08-21).
    ///
    /// So the whole-session answer is now the worst standing among the files the pane lists, which
    /// is what "All Files" has to mean if choosing one file means that file. A file with nothing
    /// to say still says it: the chip offers the install, and the tooltip names which files.
    /// </summary>
    private static string SummarySupport(IReadOnlyList<TestFile> files)
    {
        if (files.Count == 0)
        {
            return "missing";
        }

        if (files.Any(file => file.Support == "missing"))
        {
            return "missing";
        }

        return files.Any(file => file.Support == "outdated") ? "outdated" : "installed";
    }

    /// <summary>The last auto-published discovery shape, so unchanged passes stay silent.</summary>
    private string _testsFingerprint = string.Empty;

    /// <summary>
    /// AUTO-REDISCOVERY: the analysis pass just read a project because its text moved, and hands
    /// the snapshot here. Discovery and the support check are pure text scans over sources
    /// already in memory - no COM read of any kind - and the pane repaints only when the
    /// discovered shape actually changed, through a hop to the host thread. The snapshot is one
    /// FILE's, so it is merged into what the other files last said rather than replacing it.
    /// Runs on the pass's worker thread until that hop.
    /// </summary>
    internal void OnAnalysisSnapshot(string projectId, IReadOnlyList<Xlide.Vbe.Core.Engine.EngineModule> modules)
    {
        // The folder view rides the same snapshot, for the same reason: every module's text is
        // in hand, and a folder is a comment at the top of it. So does the drift between the
        // attribute annotations and the saved module's attributes.
        RememberFolders(projectId, modules);
        RememberAttributeDrift(projectId, modules);

        var pairs = new List<(string Name, string Source)>();
        string? assertSource = null;
        foreach (var module in modules)
        {
            if (!string.Equals(module.Type, "standard", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(module.ModuleName, TestRunService.AssertModuleName, StringComparison.OrdinalIgnoreCase))
            {
                assertSource = module.Source;
                continue;
            }

            if (!TestRunService.IsGeneratedModule(module.ModuleName))
            {
                pairs.Add((module.ModuleName, module.Source));
            }
        }

        var discovered = TestRunService.DiscoverFrom(pairs);
        var support = TestRunService.SupportStateOf(assertSource);

        // WHICH MODULES THIS FILE HOLDS. A module added, removed or renamed in the native
        // editor is not a gesture inside this product either, so nothing republishes the tree
        // for it - and the tree is where the Problems pane's module list comes from. The pass
        // only reaches here when the project's text moved, and the names are already in hand,
        // so the check costs a string compare. Republished only when the SET changed: an edit
        // inside a module must not walk the components every keystroke.
        var names = string.Join(";", modules.Select(one => one.ModuleName).OrderBy(one => one, StringComparer.OrdinalIgnoreCase));
        lock (_moduleShapes)
        {
            if (!_moduleShapes.TryGetValue(projectId, out var was) || was != names)
            {
                _moduleShapes[projectId] = names;
                if (was is not null)
                {
                    _editorSurface?.RunOnHostThread(PublishProjects);
                }
            }
        }

        // The pass names a project by its identity; the pane names a file by what it is called.
        // A file this pass is the first to mention takes its display name from the tree's map,
        // and falls back to the identity's own last segment - which is the file name - when the
        // map has not heard of it yet.
        var display = DisplayFromProjectId(projectId)
            ?? _testsSeen.GetValueOrDefault(projectId)?.Name
            ?? System.IO.Path.GetFileName(projectId);
        List<TestFile> files;
        lock (_testsSeen)
        {
            // Stored whatever it holds, and READ UNDER THE SAME LOCK. The copy used to be taken
            // outside it, so a pass composing on its own thread enumerated the dictionary the
            // host thread's walk was clearing and refilling - the pass would have thrown
            // "collection was modified", been caught as a failed pass, and left the pane holding
            // whatever it had (found by reading, 2026-08-21).
            _testsSeen[projectId] = new TestFile(projectId, display, support, discovered);
            files = Shown([.. _testsSeen.Values]);
        }
        var fingerprint = new StringBuilder();
        foreach (var file in files)
        {
            fingerprint.Append(file.Name).Append('/').Append(file.Support);
            foreach (var test in file.Tests)
            {
                fingerprint.Append(';').Append(test.Id).Append('@').Append(test.Line)
                    .Append('|').Append(test.SkipReason).Append('|').Append(test.XfailReason)
                    .Append('|').Append(string.Join(',', test.Tags)).Append('|').Append(test.ExpectedError)
                    .Append('|').Append(test.TimeoutMs).Append('|').Append(test.Owner).Append('|').Append(test.Requirement);
            }
        }

        var shape = fingerprint.ToString();
        if (shape == _testsFingerprint)
        {
            return;
        }

        _testsFingerprint = shape;
        _editorSurface?.RunOnHostThread(() =>
        {
            // A run in flight owns the pane; its final repaint walks fresh anyway.
            if (!_testsRunning)
            {
                _editorSurface?.ShowTests(ComposeTests(files));
            }
        });
    }

    /// <summary>
    /// The active project's identity WITHOUT touching COM - the analysis pass runs off the host
    /// thread, and reaching for ActiveVBProject from there is a call across apartments. The
    /// shown project is what the developer is looking at, which is what the ordering wants.
    /// </summary>
    private string? ActiveProjectIdOrNull() => _shownProject is { Length: > 0 } shown ? shown : null;

    /// <summary>
    /// The file the pane treats as the developer's own: the module they are looking at, and - in
    /// the beat before the editor has adopted one, which is the whole of a fresh workbook's
    /// first moments - whatever the last walk found active. Never COM from here: the analysis
    /// pass composes off the host thread.
    /// </summary>
    private string? PaneActiveProjectId() => ActiveProjectIdOrNull() ?? _walkedActiveProjectId;

    /// <summary>What the last walk found active, so a pane painted before any module is shown
    /// still knows which file the install chip belongs to.</summary>
    private volatile string? _walkedActiveProjectId;

    /// <summary>The project set the last pass reported, so an unchanged one publishes nothing.</summary>
    private string _lastProjectSet = string.Empty;

    /// <summary>Each project's module names as last seen, so only a changed SET republishes.</summary>
    private readonly Dictionary<string, string> _moduleShapes = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The analysis pass has counted the open projects. A workbook opening or closing in the
    /// host is not a gesture inside this product, so nothing here republishes the tree for it -
    /// it happened to be corrected only when an unrelated pane change came along. The tree
    /// carries both panes' file lists, so it is republished from one place, on the host thread
    /// the COM walk needs, and only when the set has actually changed.
    /// </summary>
    internal void OnProjectsObserved(IReadOnlyCollection<string> live)
    {
        var shape = string.Join(";", live.OrderBy(id => id, StringComparer.OrdinalIgnoreCase));
        if (shape == _lastProjectSet)
        {
            return;
        }

        _lastProjectSet = shape;
        _editorSurface?.RunOnHostThread(PublishProjects);
    }

    /// <summary>
    /// A FILE HAS OPENED OR CLOSED, and the Tests pane hears about it here - from the tree
    /// publish, which is where the editor notices either. Its own picture is built by merging
    /// per-file analysis snapshots, and neither event produces one: nothing is analysed for a
    /// project that has gone, and a project REOPENED inside the same session is answered from
    /// the engine's cache as unchanged, so no snapshot arrives for it either. Left to itself
    /// the pane kept the closed file's tests and never learned the reopened one's.
    ///
    /// A closing costs no COM read - the rows come from what each file last said. A file this
    /// session has never walked does cost one, once, which is what a new file is worth.
    /// </summary>
    private void ReconcileTestFiles(IReadOnlyCollection<string> liveProjectIds)
    {
        if (_testsRunning)
        {
            // A run owns the pane, and its own finally repaints from a fresh walk anyway.
            return;
        }

        // FORGETTING COMES FIRST, WHATEVER ELSE THIS TICK DOES. A file closing and another
        // opening in the same beat used to take the "something new is open" road and return
        // before any of this ran, so the closed file's results survived the very event that
        // should have dropped them.
        var forgotten = ForgetClosedFiles(liveProjectIds);

        bool unknown;
        lock (_testsSeen)
        {
            unknown = liveProjectIds.Any(id => !_testsSeen.ContainsKey(id));
        }

        if (unknown)
        {
            // Something new is open. Walk it properly - outside the lock, because the walk is
            // COM and the analysis thread merges its snapshots under the same one.
            PublishTests();
            return;
        }

        if (!forgotten)
        {
            return;
        }

        List<TestFile> left;
        lock (_testsSeen)
        {
            left = Shown([.. _testsSeen.Values]);
        }

        _testsFingerprint = string.Empty;
        _editorSurface?.ShowTests(ComposeTests(left));
    }

    /// <summary>
    /// Drops everything this session remembers about files it no longer holds, and answers
    /// whether anything went. A file that has left takes its discovery, its module-name shape
    /// and ITS RESULTS with it: kept, the results came back on reopen as green ticks and a red
    /// against code that anything could have edited while the file was away, under a clock from
    /// a run that happened before it returned (measured 2026-08-20 - close the twin, reopen it,
    /// and every row still claimed its last result). A reopened file reads "not run", which is
    /// the only thing this product actually knows about it.
    /// </summary>
    private bool ForgetClosedFiles(IReadOnlyCollection<string> liveProjectIds)
    {
        // EACH STORE IS PRUNED AGAINST THE LIVE SET, not against another store. Driving this
        // from _testsSeen looked right and was not: every walk REWRITES that map, so by the
        // time a close was noticed an unrelated refresh had already dropped the file from it,
        // "what went" came out empty, and the outcomes stayed - which is exactly how a
        // reopened file got its old green ticks back while its rows had visibly gone
        // (measured 2026-08-20, twice, the second time because the first fix looked like it
        // worked: the rows vanish either way).
        var dropped = false;
        lock (_testsSeen)
        {
            foreach (var id in _testsSeen.Keys.Where(id => !liveProjectIds.Contains(id)).ToList())
            {
                _testsSeen.Remove(id);
                dropped = true;
            }
        }

        lock (_moduleShapes)
        {
            foreach (var id in _moduleShapes.Keys.Where(id => !liveProjectIds.Contains(id)).ToList())
            {
                _moduleShapes.Remove(id);
            }
        }

        // An outcome is filed under (file, test), so the file it belongs to is the key's head.
        var forgotten = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in _testOutcomes.Keys.ToList())
        {
            var cut = key.IndexOf('\0', StringComparison.Ordinal);
            var owner = cut < 0 ? key : key[..cut];
            if (!liveProjectIds.Contains(owner))
            {
                _testOutcomes.Remove(key);
                forgotten.Add(owner);
                dropped = true;
            }
        }

        if (forgotten.Count > 0)
        {
            Log.Info($"tests: {string.Join(", ", forgotten)} closed; their results are forgotten");
        }

        // AND THE CHANGE LOG LETS GO OF A FILE THAT IS GONE. The log itself stays on disk - it is
        // a record, and a record that deletes itself when a workbook closes is not one - but this
        // session stops holding it, so a workbook opened again is read fresh rather than answered
        // from what it looked like an hour ago. The pane drops it at the same moment, because a
        // file that is not open is not in its file list.
        foreach (var id in _changeLogs.Keys.Where(id => !liveProjectIds.Contains(id)).ToList())
        {
            _changeLogs.Remove(id);
        }

        return dropped;
    }

    private void PublishTests()
    {
        try
        {
            _lastTestsPublishTicks = Environment.TickCount64;
            _editorSurface?.ShowTests(TestsSnapshot());
        }
        catch (Exception ex)
        {
            Log.Warn($"tests: the pane could not be repainted, {ex.Message}");
        }
    }

    /// <summary>When the pane was last repainted, so a fast suite does not repaint per test.</summary>
    private long _lastTestsPublishTicks;

    /// <summary>
    /// At most one repaint per this many milliseconds while a run streams. The pane is a
    /// projection of the WHOLE picture - that is what keeps it from drifting - so every
    /// repaint carries every row: measured at 400 tests, one repaint is 100KB and a run
    /// repainting per result pushed 38MB through the bridge for a suite that takes seconds.
    /// Twelve frames a second still reads as live, and a test slower than a frame is never
    /// throttled at all.
    /// </summary>
    private const long TestsRepaintMilliseconds = 80;

    /// <summary>
    /// A mid-run repaint, which may be skipped. Safe ONLY because the run's finally publishes
    /// unconditionally: whatever a skipped frame would have shown, the next one or the last
    /// one does, and the picture at rest is always exact.
    /// </summary>
    private void PublishTestsStreaming()
    {
        if (Environment.TickCount64 - _lastTestsPublishTicks < TestsRepaintMilliseconds)
        {
            return;
        }

        PublishTests();
    }

    /// <summary>
    /// Every gesture the Tests pane makes, and every verb the api's tests route takes: one
    /// entry, so the two mirror by construction. `file` narrows to one open file the way the
    /// pane's scope selector does; without it a verb means every file that has tests.
    /// `tags` and `outcomes` are the pane's filter facets as comma lists, narrowing any run
    /// verb the way the filters narrow the list - which is what makes the pane's Run Displayed
    /// and the door's run&amp;tags=&amp;outcomes= the same selection by construction.
    /// Answers in words; the pane's real answer is the stream of setTests repaints.
    /// </summary>
    private string HandleTestsAction(
        string action, string? target, string? file, string? tags = null, string? outcomes = null)
    {
        // BEFORE THE WALK, because the walk is the thing this action exists to avoid: showing
        // the pane repaints from what the analysis pass has already told us.
        if (action == "show")
        {
            try
            {
                _editorSurface?.ShowTests(TestsSnapshotCached());
            }
            catch (Exception ex)
            {
                Log.Warn($"tests: the pane could not be repainted, {ex.Message}");
            }

            return "shown";
        }

        var files = ReadTestFiles();
        if (files.Count == 0)
        {
            return "no project is open";
        }

        // A named file narrows everything below it. Named and not found is an answer, not a
        // silent fall back to the active file - which would run the wrong file's tests.
        List<TestFile> scope = files;
        if (file is { Length: > 0 })
        {
            scope = [.. files.Where(one => Names(one, file))];
            if (scope.Count == 0)
            {
                return $"no open file called {file}";
            }
        }

        switch (action)
        {
            // The walk above IS the refresh; publishing from it rather than through
            // PublishTests, which would walk every open file a second time for the same answer.
            case "refresh":
                _editorSurface?.ShowTests(ComposeTests(Shown(files)));
                _lastTestsPublishTicks = Environment.TickCount64;
                return "refreshed";

            case "install":
            {
                // Unscoped, this installs into every file that needs it - which is what the chip
                // says out loud when it is speaking for All Files, down to the count in its own
                // label. It used to skip files holding no tests, on the rule that one is "never
                // written to on a guess"; that rule and the chip's promise cannot both be kept,
                // and now that the developer picks the file themselves the press is not a guess
                // (the owner, 2026-08-21).
                var wanted = file is { Length: > 0 }
                    ? scope
                    : [.. scope.Where(one => one.Support != "installed")];

                // Nothing needs it, so the press is a deliberate re-install of the file in hand.
                if (wanted.Count == 0)
                {
                    wanted = [scope[0]];
                }

                var did = new List<string>();
                foreach (var one in wanted)
                {
                    using var project = ProjectFor(one);
                    if (project is null)
                    {
                        continue;
                    }

                    did.Add($"{one.Name}: {TestRunService.InstallSupport(project)}");
                }

                Log.Info($"tests: support module {string.Join("; ", did)}");

                // The install changes what the projects contain, so the tree and the engine are
                // told, the way a component added by hand tells them - without this the
                // analyzer's copy has no XlideAssert and squiggles every assertion as an
                // undeclared variable (the owner's screenshot, 2026-08-20).
                PublishProjects();
                _analysis?.Reanalyse();
                PublishTests();
                return did.Count == 0 ? "nothing to install into" : string.Join("; ", did);
            }

            case "run" or "runOne" or "runFile" or "runFailed" or "runModule":
            {
                if (_testsRunning)
                {
                    return "a test run is already in flight";
                }

                // The facets, validated before anything runs on their say-so. An unknown tag is
                // a legitimate empty selection - tags are the author's own vocabulary - but the
                // outcome groups are this product's five words, and a stray one is a typo to
                // name rather than a run of nothing to shrug at.
                var wantedTags = SplitFacet(tags);
                var wantedOutcomes = SplitFacet(outcomes);
                if (wantedOutcomes.FirstOrDefault(group =>
                    !OutcomeGroups.Contains(group, StringComparer.OrdinalIgnoreCase)) is { } stray)
                {
                    return $"'{stray}' is not an outcome group; use passed, failed, xfail, skipped or notRun";
                }

                // What the verb selects, per file, before the support gate: a gate that refused
                // on a file the run would never have touched would be a gate on the wrong file.
                var chosen = new List<(TestFile File, List<TestRunService.TestCase> Tests)>();
                foreach (var one in scope)
                {
                    var mine = action switch
                    {
                        "runOne" => one.Tests.Where(test =>
                            string.Equals(test.Id, target, StringComparison.OrdinalIgnoreCase)).ToList(),
                        "runModule" => one.Tests.Where(test =>
                            string.Equals(test.Module, target, StringComparison.OrdinalIgnoreCase)).ToList(),
                        // A module narrows a rerun the same way it narrows a run: the pane's
                        // Failed button sends its scope, and a rerun that ran failures the pane
                        // has scoped out of sight would disagree with the list above it.
                        "runFailed" => one.Tests.Where(test =>
                            (target is not { Length: > 0 }
                                || string.Equals(test.Module, target, StringComparison.OrdinalIgnoreCase))
                            && _testOutcomes.GetValueOrDefault(TestKey(one.ProjectId, test.Id))?.Status
                                is "failed" or "error" or "xpass").ToList(),
                        _ => [.. one.Tests],
                    };

                    // The facets narrow whatever the verb picked, per file - which is what
                    // keeps them workbook-aware: each test is judged against its own file's
                    // outcomes, and a two-file run stays a two-file run.
                    if (wantedTags.Length > 0)
                    {
                        mine = [.. mine.Where(test => TagsAdmit(test, wantedTags))];
                    }

                    if (wantedOutcomes.Length > 0)
                    {
                        mine = [.. mine.Where(test => wantedOutcomes.Contains(
                            OutcomeGroup(StandingOf(one.ProjectId, test)), StringComparer.OrdinalIgnoreCase))];
                    }

                    if (mine.Count > 0)
                    {
                        chosen.Add((one, mine));
                    }
                }

                if (chosen.Count == 0)
                {
                    // Named by what was asked for, in the order the caller said it: the module
                    // or test first, then the file, then the facets, so an empty answer says
                    // which ask was empty.
                    var facets = FacetWords(wantedTags, wantedOutcomes);
                    var narrowed = (target is { Length: > 0 } ? $" in {target}" : string.Empty) + Where(file);
                    return action switch
                    {
                        "runOne" => $"no test named {target}{Where(file)}{facets}",
                        "runModule" => $"no tests in {target}{Where(file)}{facets}",
                        "runFailed" => $"nothing has failed{narrowed}{facets}",
                        _ => $"no tests to run{Where(file)}{facets}",
                    };
                }

                // Every file the run will touch has to carry a current XlideAssert, because the
                // generated runner calls it inside that file.
                var unready = chosen.Where(one => one.File.Support != "installed").ToList();
                if (unready.Count > 0)
                {
                    var named = string.Join(", ", unready.Select(one => one.File.Name));
                    return unready.Any(one => one.File.Support == "missing")
                        ? $"the {TestRunService.AssertModuleName} module is not installed in {named} - install it first (tests?action=install)"
                        : $"the {TestRunService.AssertModuleName} module is out of date in {named} - reinstall it first (tests?action=install)";
                }

                RunTests(files, chosen);
                return chosen.Count == 1
                    ? $"ran {chosen[0].Tests.Count} in {chosen[0].File.Name}"
                    : $"ran {chosen.Sum(one => one.Tests.Count)} across {chosen.Count} files";
            }

            case "debug" when target is { Length: > 0 }:
            {
                var holders = scope
                    .Select(one => (File: one, Test: one.Tests.FirstOrDefault(test =>
                        string.Equals(test.Id, target, StringComparison.OrdinalIgnoreCase))))
                    .Where(pair => pair.Test is not null)
                    .ToList();
                if (holders.Count == 0)
                {
                    return $"no test named {target}{Where(file)}";
                }

                if (holders.Count > 1)
                {
                    return $"{target} is in {string.Join(" and ", holders.Select(one => one.File.Name))}"
                        + " - name the file (file=) to say which";
                }

                // Gated like a run: a test that calls a missing XlideAssert is a compile-error
                // modal, not a debug session.
                var (home, wanted) = holders[0];
                if (home.Support != "installed")
                {
                    return home.Support == "missing"
                        ? $"the {TestRunService.AssertModuleName} module is not installed in {home.Name} - install it first"
                        : $"the {TestRunService.AssertModuleName} module is out of date in {home.Name} - reinstall it first";
                }

                using var project = ProjectFor(home);
                if (project is null)
                {
                    return $"{home.Name} would not answer";
                }

                Log.Info($"tests: debugging {wanted!.Id} in {home.Name}");
                return TestRunService.DebugTest(project, wanted);
            }

            default:
                return $"'{action}' is not a tests action; use refresh, install, run, runFile, runModule, runOne, runFailed or debug";
        }
    }

    /// <summary>Whether a file answers to a name: what it is shown as, or its identity.</summary>
    private static bool Names(TestFile file, string wanted) =>
        string.Equals(file.Name, wanted, StringComparison.OrdinalIgnoreCase)
        || string.Equals(file.ProjectId, wanted, StringComparison.OrdinalIgnoreCase);

    private static string Where(string? file) => file is { Length: > 0 } ? $" in {file}" : string.Empty;

    /// <summary>A comma- or space-separated facet list, split the way the tags directive splits.</summary>
    private static string[] SplitFacet(string? list) =>
        list is { Length: > 0 }
            ? list.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            : [];

    /// <summary>
    /// Whether a test wears ANY of the wanted tags - the union the pane's checkboxes mean, the
    /// same way its outcome chips mean a union. The word `untagged` is reserved: it admits a
    /// test carrying no tags at all, so the pane's "(untagged)" choice has a wire spelling.
    /// A project that really tags tests "untagged" gets both meanings at once, documented
    /// rather than disambiguated - there is no spelling that could not also be a tag.
    /// </summary>
    private static bool TagsAdmit(TestRunService.TestCase test, string[] wanted) =>
        wanted.Any(tag =>
            (string.Equals(tag, "untagged", StringComparison.OrdinalIgnoreCase) && test.Tags.Length == 0)
            || test.Tags.Contains(tag, StringComparer.OrdinalIgnoreCase));

    /// <summary>The status a test would wear in the pane right now: its outcome, or its mark.</summary>
    private string StandingOf(string projectId, TestRunService.TestCase test) =>
        _testOutcomes.GetValueOrDefault(TestKey(projectId, test.Id))?.Status
            ?? (test.SkipReason is not null ? "skip-marked" : "none");

    private static readonly string[] OutcomeGroups = ["passed", "failed", "xfail", "skipped", "notRun"];

    /// <summary>
    /// The five filter groups the pane's outcome chips press, folded from a row status the
    /// same way testspane.ts folds them - so outcomes=failed selects exactly the rows the
    /// Failed chip governs: failed, error and xpass together.
    /// </summary>
    private static string OutcomeGroup(string status) => status switch
    {
        "passed" => "passed",
        "failed" or "error" or "xpass" => "failed",
        "xfail" => "xfail",
        "skipped" or "skip-marked" => "skipped",
        _ => "notRun",
    };

    /// <summary>How an empty answer names the facets that emptied it.</summary>
    private static string FacetWords(string[] tags, string[] outcomes)
    {
        var parts = new List<string>();
        if (tags.Length > 0)
        {
            parts.Add($"tag{(tags.Length == 1 ? string.Empty : "s")} {string.Join(", ", tags)}");
        }

        if (outcomes.Length > 0)
        {
            parts.Add($"outcome{(outcomes.Length == 1 ? string.Empty : "s")} {string.Join(", ", outcomes)}");
        }

        return parts.Count == 0 ? string.Empty : $" with {string.Join(" and ", parts)}";
    }

    private DispatchObject? ProjectFor(TestFile file) =>
        FindProjectByDisplayName(file.Name) ?? FindProjectByDisplayName(file.ProjectId);

    /// <summary>
    /// The run itself, on the host thread the caller already holds: file by file, sequential,
    /// one Application.Run per test, the pane repainted the moment each result lands. The
    /// door's caller waits for the whole answer; the pane's caller watches it stream.
    ///
    /// The pane keeps painting EVERY discovered test throughout, not just the chosen ones, or a
    /// rerun of one file would blank every other file's results for its duration.
    /// </summary>
    private void RunTests(
        List<TestFile> everything,
        List<(TestFile File, List<TestRunService.TestCase> Tests)> chosen)
    {
        _testsRunning = true;
        _testsBreakCleared = false;
        _testsCurrent = null;
        _testsCurrentProject = null;

        /*
         * WHAT IS ABOUT TO BE RUN HAS NO RESULT UNTIL IT HAS RUN.
         *
         * Outcomes used to survive until something overwrote them, which is fine when every
         * chosen test lands and a lie when one does not. Measured against a project with one
         * unrelated module that would not parse: the run could not execute a line, and the pane
         * went on showing five rows reading `passed` from the run before the code was broken,
         * with a sixth stuck on `running`. A developer looking at that sees green.
         *
         * Only the CHOSEN ones. A rerun of one file must not blank another file's results, which
         * is the same reason the pane keeps painting every discovered test throughout.
         */
        foreach (var (file, tests) in chosen)
        {
            foreach (var test in tests)
            {
                _testOutcomes.Remove(TestKey(file.ProjectId, test.Id));
            }
        }

        _testsLive = Shown(everything);
        PublishTests();
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var landed = 0;
        try
        {
            foreach (var (file, tests) in chosen)
            {
                // ONE FILE'S FAILURE IS ONE FILE'S. A project that refuses the generated module -
                // a compile error standing in it, a workbook the host has half let go of - used
                // to take every file after it down with it: the exception left the loop and the
                // remaining files' tests simply never ran, showing as "not run" with nothing
                // said. Each file is now attempted, and a file that could not be run says so on
                // its own rows rather than through the log alone.
                var landedHere = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    using var project = ProjectFor(file);
                    if (project is null)
                    {
                        throw new InvalidOperationException($"{file.Name} would not answer.");
                    }

                    _testsCurrentProject = file.ProjectId;
                    var reads = TestRunService.ReadStandardModules(project);
                    TestRunService.Run(
                        project, reads, tests, new TestRunService.Selection(null, null, []), failFast: false,
                        starting: id =>
                        {
                            _testsCurrent = id;
                            PublishTestsStreaming();
                        },
                        landed: result =>
                        {
                            landed++;
                            landedHere.Add(result.Id);
                            _testOutcomes[TestKey(file.ProjectId, result.Id)] = result;
                            _testsCurrent = null;
                            PublishTestsStreaming();
                        });

                    if (TestRunService.LeftDesignMode(project))
                    {
                        Log.Warn($"tests: the run left {file.Name} out of design mode - a test left something standing");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"tests: {file.Name} could not be run", ex);
                    var why = $"{file.Name} could not be run: {ex.Message.Trim()}";
                    foreach (var test in tests.Where(one => !landedHere.Contains(one.Id)))
                    {
                        _testOutcomes[TestKey(file.ProjectId, test.Id)] = new TestRunService.TestResult(
                            test.Id, test.Module, test.Procedure, test.Line, "error", why, 0, [], test.Tags);
                    }

                    _testsCurrent = null;
                    PublishTests();
                }
            }
        }
        finally
        {
            _testsRunning = false;
            _testsCurrent = null;
            _testsCurrentProject = null;
            _testsLive = null;
            _testsRanAt = DateTimeOffset.Now;
            watch.Stop();
            PublishTests();

            // The run added and removed its generated modules; the tree and the engine hear
            // about it so neither carries a ghost of a module that no longer exists - and a
            // test is allowed to have edited real modules, which the engine should re-read.
            PublishProjects();
            _analysis?.Reanalyse();
            Log.Info($"tests: {landed} result(s) in {watch.ElapsedMilliseconds}ms");
        }
    }

    private void OnTestsAction(string action, string? target, string? file, string? tags, string? outcomes)
    {
        try
        {
            var answer = HandleTestsAction(action, target, file, tags, outcomes);
            Log.Info($"tests: {action}{(target is { Length: > 0 } ? $" {target}" : string.Empty)}"
                + $"{(file is { Length: > 0 } ? $" in {file}" : string.Empty)}"
                + $"{(tags is { Length: > 0 } ? $" tags={tags}" : string.Empty)}"
                + $"{(outcomes is { Length: > 0 } ? $" outcomes={outcomes}" : string.Empty)} -> {answer}");
        }
        catch (Exception ex)
        {
            Log.Error($"tests: {action} failed", ex);
            _testsRunning = false;
            _testsLive = null;
            PublishTests();
        }
    }
}
