using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The test runner's session half: one brain for the Tests pane and the debug api, so the two
/// doors cannot disagree about what a run is. Discovery, the support-module gate, the run loop
/// and its per-test streaming all live in <see cref="TestRunService"/>; this partial owns the
/// standing state - the latest outcome per test, whether a run is in flight - and the message
/// that repaints the pane after every change.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>The latest outcome per test id, kept across runs so the pane remembers.</summary>
    private readonly Dictionary<string, TestRunService.TestResult> _testOutcomes = new(StringComparer.OrdinalIgnoreCase);

    private bool _testsRunning;
    private string? _testsCurrent;

    /// <summary>
    /// While a run is in flight: the support state and the discovery the run started with.
    /// Every landing result repaints the pane, and a repaint that re-walked every module's
    /// text over COM turned an N-test run into N whole-project reads - measured as the
    /// dominant cost of a run the day after the runner shipped. Patched statuses over this
    /// cache instead; the run's last repaint clears it and walks fresh, because a test is
    /// allowed to have edited modules.
    /// </summary>
    private (string Support, List<TestRunService.TestCase> Discovered)? _testsLive;

    /// <summary>
    /// The pane's whole picture right now: support state, run state, and every discovered test
    /// merged with its latest outcome. Discovery walks the live modules on every call - tests
    /// are code, and code moves under the pane - except mid-run, where the run's own snapshot
    /// stands in.
    /// </summary>
    internal SetTestsMessage TestsSnapshot()
    {
        string support;
        List<TestRunService.TestCase> discovered;
        if (_testsLive is { } live)
        {
            (support, discovered) = live;
        }
        else
        {
            using var project = _editor.GetObject("ActiveVBProject");
            if (project is null)
            {
                return new SetTestsMessage("setTests", "missing", false, null, []);
            }

            support = TestRunService.SupportState(project);
            discovered = TestRunService.Discover(project);
        }

        var rows = new TestRowMessage[discovered.Count];
        for (var i = 0; i < discovered.Count; i++)
        {
            var test = discovered[i];
            var outcome = _testOutcomes.GetValueOrDefault(test.Id);
            var status = _testsRunning && string.Equals(_testsCurrent, test.Id, StringComparison.OrdinalIgnoreCase)
                ? "running"
                : outcome?.Status ?? (test.SkipReason is not null ? "skip-marked" : "none");
            rows[i] = new TestRowMessage(
                test.Id, test.Module, test.Procedure, test.Line, status,
                outcome?.Message, outcome?.DurationMs ?? 0, outcome?.Output ?? [],
                test.Tags, test.Owner, test.Requirement, test.TimeoutMs, test.ExpectedError);
        }

        return new SetTestsMessage("setTests", support, _testsRunning, _testsCurrent, rows);
    }

    private void PublishTests()
    {
        try
        {
            _editorSurface?.ShowTests(TestsSnapshot());
        }
        catch (Exception ex)
        {
            Log.Warn($"tests: the pane could not be repainted, {ex.Message}");
        }
    }

    /// <summary>
    /// Every gesture the Tests pane makes, and every verb the api's tests route takes: one
    /// entry, so the two mirror by construction. Answers in words; the pane's real answer is
    /// the stream of setTests repaints.
    /// </summary>
    private string HandleTestsAction(string action, string? target)
    {
        using var project = _editor.GetObject("ActiveVBProject");
        if (project is null)
        {
            return "no project is open";
        }

        switch (action)
        {
            case "refresh":
                PublishTests();
                return "refreshed";

            case "install":
            {
                var installed = TestRunService.InstallSupport(project);
                Log.Info($"tests: support module {installed}");

                // The install changes what the project contains, so the tree and the engine
                // are told, the way a component added by hand tells them - without this the
                // analyzer's copy has no XlideAssert and squiggles every assertion as an
                // undeclared variable (the owner's screenshot, 2026-08-20).
                PublishProjects();
                _analysis?.Reanalyse();
                PublishTests();
                return installed;
            }

            case "run" or "runOne" or "runFailed" or "runModule":
            {
                if (_testsRunning)
                {
                    return "a test run is already in flight";
                }

                var support = TestRunService.SupportState(project);
                if (support != "installed")
                {
                    return support == "missing"
                        ? $"the {TestRunService.AssertModuleName} module is not installed - install it first (tests?action=install)"
                        : $"the {TestRunService.AssertModuleName} module is out of date - reinstall it first (tests?action=install)";
                }

                // One project walk serves discovery, the pane's mid-run snapshot, AND the
                // dispatcher's target scan - the module read is the expensive COM call here.
                var reads = TestRunService.ReadStandardModules(project);
                var discovered = TestRunService.DiscoverFrom(reads);
                var selection = action switch
                {
                    "runOne" => new TestRunService.Selection(null, target, []),
                    "runModule" => new TestRunService.Selection(target, null, []),
                    _ => new TestRunService.Selection(null, null, []),
                };

                // A misspelled target answers by name, not with a run that ran nothing.
                if (action == "runOne" && !discovered.Any(test =>
                    string.Equals(test.Id, target, StringComparison.OrdinalIgnoreCase)))
                {
                    return $"no test named {target}";
                }

                if (action == "runModule" && !discovered.Any(test =>
                    string.Equals(test.Module, target, StringComparison.OrdinalIgnoreCase)))
                {
                    return $"no tests in {target}";
                }

                // runFailed narrows what RUNS; the pane keeps painting the whole discovery,
                // or every green row would vanish for the duration of the rerun.
                var runSet = discovered;
                if (action == "runFailed")
                {
                    var failed = _testOutcomes.Values
                        .Where(one => one.Status is "failed" or "error" or "xpass")
                        .Select(one => one.Id)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    runSet = [.. discovered.Where(test => failed.Contains(test.Id))];
                    if (runSet.Count == 0)
                    {
                        return "nothing has failed";
                    }
                }

                RunTests(project, reads, discovered, runSet, selection, support, failFast: false);
                return "ran";
            }

            case "debug" when target is { Length: > 0 }:
            {
                var wanted = TestRunService.Discover(project)
                    .FirstOrDefault(test => string.Equals(test.Id, target, StringComparison.OrdinalIgnoreCase));
                if (wanted is null)
                {
                    return $"no test named {target}";
                }

                Log.Info($"tests: debugging {wanted.Id}");
                return TestRunService.DebugTest(project, wanted);
            }

            default:
                return $"'{action}' is not a tests action; use refresh, install, run, runModule, runOne, runFailed or debug";
        }
    }

    /// <summary>
    /// The run itself, on the host thread the caller already holds: sequential, one
    /// Application.Run per test, the pane repainted the moment each result lands. The door's
    /// caller waits for the whole answer; the pane's caller watches it stream.
    /// </summary>
    private void RunTests(
        DispatchObject project,
        IReadOnlyList<(string Name, string Source)> reads,
        List<TestRunService.TestCase> discovered,
        IReadOnlyList<TestRunService.TestCase> runSet,
        TestRunService.Selection selection,
        string support,
        bool failFast)
    {
        _testsRunning = true;
        _testsCurrent = null;
        _testsLive = (support, discovered);
        PublishTests();
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var landed = 0;
        try
        {
            TestRunService.Run(
                project, reads, runSet, selection, failFast,
                starting: id =>
                {
                    _testsCurrent = id;
                    PublishTests();
                },
                landed: result =>
                {
                    landed++;
                    _testOutcomes[result.Id] = result;
                    _testsCurrent = null;
                    PublishTests();
                });
        }
        finally
        {
            _testsRunning = false;
            _testsCurrent = null;
            _testsLive = null;
            watch.Stop();
            PublishTests();

            // The run added and removed its generated modules; the tree and the engine hear
            // about it so neither carries a ghost of a module that no longer exists - and a
            // test is allowed to have edited real modules, which the engine should re-read.
            PublishProjects();
            _analysis?.Reanalyse();
            Log.Info($"tests: {landed} result(s) in {watch.ElapsedMilliseconds}ms");
            if (TestRunService.LeftDesignMode(project))
            {
                Log.Warn("tests: the run left the project out of design mode - a test left something standing");
            }
        }
    }

    private void OnTestsAction(string action, string? target)
    {
        try
        {
            var answer = HandleTestsAction(action, target);
            Log.Info($"tests: {action}{(target is { Length: > 0 } ? $" {target}" : string.Empty)} -> {answer}");
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
