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
    /// The pane's whole picture right now: support state, run state, and every discovered test
    /// merged with its latest outcome. Discovery walks the live modules on every call - tests
    /// are code, and code moves under the pane.
    /// </summary>
    internal SetTestsMessage TestsSnapshot()
    {
        using var project = _editor.GetObject("ActiveVBProject");
        if (project is null)
        {
            return new SetTestsMessage("setTests", "missing", false, null, []);
        }

        var support = TestRunService.SupportState(project);
        var discovered = TestRunService.Discover(project);
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

                var discovered = TestRunService.Discover(project);
                var selection = action switch
                {
                    "runOne" => new TestRunService.Selection(null, target, []),
                    "runModule" => new TestRunService.Selection(target, null, []),
                    _ => new TestRunService.Selection(null, null, []),
                };

                if (action == "runFailed")
                {
                    var failed = _testOutcomes.Values
                        .Where(one => one.Status is "failed" or "error" or "xpass")
                        .Select(one => one.Id)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    discovered = [.. discovered.Where(test => failed.Contains(test.Id))];
                    if (discovered.Count == 0)
                    {
                        return "nothing has failed";
                    }
                }

                RunTests(project, discovered, selection, failFast: false);
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
        IReadOnlyList<TestRunService.TestCase> discovered,
        TestRunService.Selection selection,
        bool failFast)
    {
        _testsRunning = true;
        _testsCurrent = null;
        PublishTests();
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var landed = 0;
        try
        {
            TestRunService.Run(
                project, discovered, selection, failFast,
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
            watch.Stop();
            PublishTests();
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
            PublishTests();
        }
    }
}
