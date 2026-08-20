using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/*
 * The VBA test runner, living where the tests live.
 *
 * The authoring surface is the one xlide_vscode ships - `' @xlide-test` directives over
 * zero-argument Subs in standard modules, the XlideAssert module's latched first-failure
 * protocol, a generated RunTest(testId) returning one JSON object per test - so a project's
 * tests move between the two products without an edit. The EXECUTION is this product's own:
 * no staged copy, no owned hidden host, no PowerShell bridge. The tests run in the live
 * project through the same Application.Run road the Immediate window takes, one test per
 * call, and each result streams to the page as it lands.
 *
 * What the out-of-process runner had that this one deliberately does not: a watchdog that
 * can kill the host. In-process VBA cannot be preempted, so a test's `timeout=` metadata is
 * carried and shown but not enforced - the developer's Ctrl+Break remains the escape, the
 * same one the native F5 has.
 */
internal static partial class TestRunService
{
    internal const string AssertModuleName = "XlideAssert";
    internal const string DispatchModuleName = "XlideTestDispatch";
    internal const string RunnerModulePrefix = "XlideRun";
    private const int StandardModule = 1;
    private const int DesignMode = 2;

    /// <summary>One discovered test: identity, place, and every directive fact about it.</summary>
    internal sealed record TestCase(
        string Module,
        string Procedure,
        string Id,
        int Line,
        string[] Tags,
        string? Owner,
        string? Requirement,
        int? TimeoutMs,
        string? ExpectedError,
        string? SkipReason,
        string? XfailReason);

    /// <summary>One test's outcome, in the vocabulary the results table speaks.</summary>
    internal sealed record TestResult(
        string Id,
        string Module,
        string Procedure,
        int Line,
        string Status,
        string? Message,
        double DurationMs,
        string[] Output,
        string[] Tags);

    /// <summary>What a run selection means: nothing set runs everything discovered.</summary>
    internal sealed record Selection(string? Module, string? Test, string[] Tags);

    // ---------------------------------------------------------------- discovery

    /// <summary>
    /// Every standard module's name and whole text, read once. The module read is the
    /// expensive COM call here, and both consumers - test discovery and the dispatcher's
    /// target scan - used to make it separately, which doubled every run's project walk.
    /// The runner's own generated modules are left out; XlideAssert rides along because the
    /// dispatcher offers its public zero-argument Subs as targets, the way the sibling does.
    /// </summary>
    internal static List<(string Name, string Source)> ReadStandardModules(DispatchObject project)
    {
        var reads = new List<(string, string)>();
        using var components = project.GetObject("VBComponents");
        if (components is null)
        {
            return reads;
        }

        var count = components.GetInt32("Count");
        for (var i = 1; i <= count; i++)
        {
            using var component = components.GetItem(i);
            if (component is null || component.GetInt32("Type") != StandardModule)
            {
                continue;
            }

            var name = component.GetString("Name") ?? string.Empty;
            if (IsGeneratedModule(name) && !string.Equals(name, AssertModuleName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var code = component.GetObject("CodeModule");
            var lineCount = code?.GetInt32("CountOfLines") ?? 0;
            if (code is null || lineCount == 0)
            {
                continue;
            }

            reads.Add((name, code.CallToString("Lines", 1, lineCount)));
        }

        return reads;
    }

    /// <summary>
    /// Every test the project declares, walked straight off the live modules: standard modules
    /// only, a contiguous comment block ending immediately above a zero-argument PUBLIC Sub,
    /// carrying at least one `@xlide-test` directive. Private Subs are excluded here on
    /// purpose - the generated runner calls `Module.Proc`, which cannot compile against a
    /// Private target. (xlide_vscode discovers them and then fails the whole generated module;
    /// filed there rather than copied here.)
    /// </summary>
    internal static List<TestCase> Discover(DispatchObject project) => DiscoverFrom(ReadStandardModules(project));

    internal static List<TestCase> DiscoverFrom(IReadOnlyList<(string Name, string Source)> modules)
    {
        var tests = new List<TestCase>();
        foreach (var (name, source) in modules)
        {
            if (string.Equals(name, AssertModuleName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            tests.AddRange(DiscoverInSource(name, source));
        }

        tests.Sort((left, right) =>
        {
            var byModule = string.Compare(left.Module, right.Module, StringComparison.OrdinalIgnoreCase);
            return byModule != 0 ? byModule : left.Line.CompareTo(right.Line);
        });
        return tests;
    }

    internal static bool IsGeneratedModule(string name) =>
        string.Equals(name, AssertModuleName, StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, DispatchModuleName, StringComparison.OrdinalIgnoreCase)
        || name.StartsWith(RunnerModulePrefix, StringComparison.OrdinalIgnoreCase)
        || string.Equals(name, "XlideImmediateScratch", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The directive grammar, ported line for line: `' @xlide-test`, `-skip`, `-xfail`, then
    /// key=value metadata with quoted or bare values, a bare `expected-error` meaning any, and
    /// a trailing ` -- comment` stripped before parsing.
    /// </summary>
    internal static List<TestCase> DiscoverInSource(string moduleName, string source)
    {
        var found = new List<TestCase>();
        var lines = source.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');

        for (var index = 0; index < lines.Length; index++)
        {
            var declaration = ProcedureDeclaration().Match(lines[index]);
            if (!declaration.Success || declaration.Groups["access"].Value.Contains("Private", StringComparison.OrdinalIgnoreCase)
                || declaration.Groups["access"].Value.Contains("Friend", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var metadata = PrecedingDirectives(lines, index);
            if (metadata is null)
            {
                continue;
            }

            var procedure = declaration.Groups["name"].Value;
            found.Add(metadata with
            {
                Module = moduleName,
                Procedure = procedure,
                Id = $"{moduleName}.{procedure}",
                Line = index + 1,
            });
        }

        return found;
    }

    /// <summary>
    /// A zero-argument Sub's declaration line. The empty parentheses are required ON the
    /// declaration line - a parameter list continued with `_` is by definition not empty, and
    /// a test taking parameters is not runnable anyway.
    /// </summary>
    [GeneratedRegex(@"^\s*(?<access>(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?)Sub\s+(?<name>[^\W\d]\w*)\s*\(\s*\)", RegexOptions.IgnoreCase)]
    private static partial Regex ProcedureDeclaration();

    [GeneratedRegex(@"^\s*'\s*@(xlide-test(?:-(?<suffix>skip|xfail))?)\b(?<rest>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex DirectiveLine();

    [GeneratedRegex("([A-Za-z][A-Za-z0-9_-]*)=(?:\"([^\"]*)\"|'([^']*)'|(\\S+))")]
    private static partial Regex MetadataPair();

    private static TestCase? PrecedingDirectives(string[] lines, int declarationIndex)
    {
        string[] tags = [];
        string? owner = null, requirement = null, expectedError = null, skipReason = null, xfailReason = null;
        int? timeoutMs = null;
        var discovered = false;

        for (var index = declarationIndex - 1; index >= 0; index--)
        {
            var line = lines[index];
            if (line.Trim().Length == 0 || !line.TrimStart().StartsWith('\''))
            {
                break;
            }

            var directive = DirectiveLine().Match(line);
            if (!directive.Success)
            {
                continue;
            }

            discovered = true;
            var suffix = directive.Groups["suffix"].Value.ToLowerInvariant();
            var rest = directive.Groups["rest"].Value;

            // ` -- anything after` is the author's own aside, never metadata.
            var aside = Regex.Match(rest, @"\s--");
            if (aside.Success)
            {
                rest = rest[..aside.Index];
            }

            var consumed = new List<(int Start, int End)>();
            foreach (Match pair in MetadataPair().Matches(rest))
            {
                consumed.Add((pair.Index, pair.Index + pair.Length));
                var key = pair.Groups[1].Value.ToLowerInvariant();
                var value = pair.Groups[2].Success ? pair.Groups[2].Value
                    : pair.Groups[3].Success ? pair.Groups[3].Value
                    : pair.Groups[4].Value;
                switch (key)
                {
                    case "tags":
                        tags = [.. value.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)];
                        break;
                    case "owner":
                        owner = value;
                        break;
                    case "requirement" or "req":
                        requirement = value;
                        break;
                    case "timeout" or "timeoutms":
                        timeoutMs = ParseTimeout(value);
                        break;
                    case "expected-error" or "expectederror":
                        expectedError = value;
                        break;
                    case "reason":
                        if (suffix == "skip") { skipReason = value; }
                        if (suffix == "xfail") { xfailReason = value; }
                        break;
                }
            }

            // A bare `expected-error` token between the pairs means "any error will do".
            var leftovers = new StringBuilder(rest);
            for (var backward = consumed.Count - 1; backward >= 0; backward--)
            {
                leftovers.Remove(consumed[backward].Start, consumed[backward].End - consumed[backward].Start);
            }

            foreach (var token in leftovers.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (token.Equals("expected-error", StringComparison.OrdinalIgnoreCase)
                    || token.Equals("expectederror", StringComparison.OrdinalIgnoreCase))
                {
                    expectedError = "any";
                }
            }

            if (suffix == "skip" && skipReason is null) { skipReason = string.Empty; }
            if (suffix == "xfail" && xfailReason is null) { xfailReason = string.Empty; }
        }

        return discovered
            ? new TestCase(string.Empty, string.Empty, string.Empty, 0, tags, owner, requirement,
                timeoutMs, expectedError, skipReason, xfailReason)
            : null;
    }

    private static int? ParseTimeout(string value)
    {
        var text = value.Trim().ToLowerInvariant();
        var multiplier = 1;
        if (text.EndsWith("ms", StringComparison.Ordinal))
        {
            text = text[..^2];
        }
        else if (text.EndsWith('s'))
        {
            text = text[..^1];
            multiplier = 1000;
        }

        return int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) && parsed > 0
            ? parsed * multiplier
            : null;
    }

    // ---------------------------------------------------------------- support module

    /// <summary>"missing", "outdated" or "installed": whether XlideAssert is in the project as shipped.</summary>
    internal static string SupportState(DispatchObject project)
    {
        using var components = project.GetObject("VBComponents");
        using var installed = Find(components, AssertModuleName);
        if (installed is null)
        {
            return "missing";
        }

        using var code = installed.GetObject("CodeModule");
        var lineCount = code?.GetInt32("CountOfLines") ?? 0;
        var held = lineCount > 0 ? code!.CallToString("Lines", 1, lineCount) : string.Empty;

        // Case-insensitively, because VBA re-cases identifiers PROJECT-WIDE to the latest
        // declaration: a module elsewhere declaring `number` turns this module's `Err.Number`
        // into `Err.number` without an edit here, and byte equality then cried outdated over
        // text the developer never touched. The sibling product's installed copy also spells
        // some parameters in lower case; folding keeps the two products agreeing.
        return string.Equals(
            Normalized(held), Normalized(AssertModuleSource), StringComparison.OrdinalIgnoreCase)
            ? "installed"
            : "outdated";
    }

    /// <summary>
    /// Writes XlideAssert into the project, replacing an out-of-date copy. This is a REAL
    /// module the developer's tests compile against, so it is installed once and saved with
    /// the workbook, not injected per run the way the runner is.
    /// </summary>
    internal static string InstallSupport(DispatchObject project)
    {
        using var components = project.GetObject("VBComponents");
        if (components is null)
        {
            return "the project would not answer its components";
        }

        using var existing = Find(components, AssertModuleName);
        if (existing is not null)
        {
            using var code = existing.GetObject("CodeModule");
            var lineCount = code?.GetInt32("CountOfLines") ?? 0;
            if (lineCount > 0)
            {
                code!.Invoke("DeleteLines", 1, lineCount);
            }

            code!.Invoke("AddFromString", AssertModuleSource);
            return "installed";
        }

        using var component = components.CallObject("Add", StandardModule);
        if (component is null)
        {
            return "the project would not accept the support module";
        }

        component.SetString("Name", AssertModuleName);
        using var fresh = component.GetObject("CodeModule");
        fresh?.Invoke("AddFromString", AssertModuleSource);
        return "installed";
    }

    /// <summary>EOL-normalized, attribute lines dropped, trimmed: the same equality the sibling product uses.</summary>
    private static string Normalized(string source) =>
        string.Join('\n', source
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n')
            .Where(line => !line.TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase)))
        .Trim();

    private static DispatchObject? Find(DispatchObject? components, string name)
    {
        if (components is null)
        {
            return null;
        }

        var count = components.GetInt32("Count");
        for (var i = 1; i <= count; i++)
        {
            var candidate = components.GetItem(i);
            if (candidate is not null
                && string.Equals(candidate.GetString("Name"), name, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }

            candidate?.Dispose();
        }

        return null;
    }

    // ---------------------------------------------------------------- codegen

    /// <summary>Full-unicode VBA identifier, the same gate the sibling product applies before codegen.</summary>
    [GeneratedRegex(@"^[\p{L}_][\p{L}\p{M}\p{N}_]*$")]
    private static partial Regex VbaIdentifier();

    internal static string? IdentifierComplaint(IEnumerable<TestCase> tests)
    {
        foreach (var test in tests)
        {
            if (!VbaIdentifier().IsMatch(test.Module) || !VbaIdentifier().IsMatch(test.Procedure))
            {
                return $"'{test.Id}' is not a name the generated runner can spell; rename the module or procedure.";
            }
        }

        return null;
    }

    /// <summary>The per-run runner: RunTest(testId) answering one JSON object per call.</summary>
    internal static string BuildRunnerModule(IReadOnlyList<TestCase> tests)
    {
        var body = new StringBuilder();
        body.Append("Option Explicit\r\n\r\n");
        body.Append("Public Function RunTest(ByVal testId As String) As String\r\n");
        body.Append($"    {AssertModuleName}.ResetTestState\r\n");
        body.Append("    On Error GoTo Caught\r\n");
        body.Append("    Select Case testId\r\n");
        foreach (var test in tests)
        {
            body.Append($"        Case \"{test.Id.Replace("\"", "\"\"", StringComparison.Ordinal)}\"\r\n");
            body.Append($"            Call {test.Module}.{test.Procedure}\r\n");
        }

        body.Append("        Case Else\r\n");
        body.Append("            RunTest = FailureJson(5, \"XLIDE.TestRunner\", \"Unknown XLIDE test: \" & testId)\r\n");
        body.Append("            Exit Function\r\n");
        body.Append("    End Select\r\n");
        body.Append("    On Error GoTo 0\r\n");
        body.Append($"    If Len({AssertModuleName}.LastFailureMessage()) > 0 Then\r\n");
        body.Append($"        RunTest = FailureJson({AssertModuleName}.AssertionErrorNumber(), \"XLIDE.Assert\", {AssertModuleName}.LastFailureMessage())\r\n");
        body.Append("    Else\r\n");
        body.Append($"        RunTest = \"{{\"\"outcome\"\":\"\"passed\"\",\"\"output\"\":\" & {AssertModuleName}.OutputJson() & \"}}\"\r\n");
        body.Append("    End If\r\n");
        body.Append("    Exit Function\r\n");
        body.Append("Caught:\r\n");
        body.Append("    Dim actualNumber As Long\r\n");
        body.Append("    Dim actualSource As String\r\n");
        body.Append("    Dim actualDescription As String\r\n");
        body.Append("    actualNumber = Err.Number\r\n");
        body.Append("    actualSource = Err.Source\r\n");
        body.Append("    actualDescription = Err.Description\r\n");
        body.Append("    On Error GoTo 0\r\n");
        body.Append("    RunTest = FailureJson(actualNumber, actualSource, actualDescription)\r\n");
        body.Append("End Function\r\n\r\n");
        body.Append("Private Function FailureJson(ByVal Number As Long, ByVal Source As String, ByVal Message As String) As String\r\n");
        body.Append($"    FailureJson = \"{{\"\"outcome\"\":\"\"failed\"\",\"\"number\"\":\" & CStr(Number) & \",\"\"source\"\":\"\"\" & JsonEscape(Source) & \"\"\",\"\"message\"\":\"\"\" & JsonEscape(Message) & \"\"\",\"\"output\"\":\" & {AssertModuleName}.OutputJson() & \"}}\"\r\n");
        body.Append("End Function\r\n\r\n");
        body.Append(JsonEscapeFunction);
        return body.ToString();
    }

    /// <summary>
    /// The dispatcher Assert.Throws runs its target through: every public zero-argument Sub in
    /// every standard module, direct-called inside one execution context so an error crosses
    /// back as recorded state rather than as a host modal. `Option Compare Text` with
    /// original-cased keys keeps the matching on VBA's own case rule. Chunked at 100 targets a
    /// procedure, under VBA's compiled-procedure cap.
    /// </summary>
    internal static string BuildDispatchModule(IReadOnlyList<(string Name, string Source)> modules)
    {
        var targets = new List<(string Module, string Procedure)>();
        foreach (var (name, source) in modules)
        {
            foreach (var line in source.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
            {
                var declaration = ProcedureDeclaration().Match(line);
                if (declaration.Success
                    && !declaration.Groups["access"].Value.Contains("Private", StringComparison.OrdinalIgnoreCase)
                    && !declaration.Groups["access"].Value.Contains("Friend", StringComparison.OrdinalIgnoreCase)
                    && VbaIdentifier().IsMatch(name)
                    && VbaIdentifier().IsMatch(declaration.Groups["name"].Value))
                {
                    targets.Add((name, declaration.Groups["name"].Value));
                }
            }
        }

        var bareCounts = targets
            .GroupBy(one => one.Procedure, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        var body = new StringBuilder();
        body.Append("Option Explicit\r\n");
        body.Append("Option Compare Text\r\n\r\n");
        body.Append("' Generated for one XLIDE test run; removed when the run ends.\r\n");
        body.Append("Public Sub XlideInvokeTarget(ByVal macroName As String)\r\n");
        body.Append("    On Error GoTo Caught\r\n");
        body.Append($"    {AssertModuleName}.RecordTargetOutcome 0, \"\", \"\"\r\n");
        body.Append("    Dim targetKey As String\r\n");
        body.Append("    targetKey = Trim$(macroName)\r\n");

        const int chunkSize = 100;
        var chunks = (targets.Count + chunkSize - 1) / chunkSize;
        for (var chunk = 0; chunk < Math.Max(1, chunks); chunk++)
        {
            body.Append($"    If XlideDispatch{chunk}(targetKey) Then Exit Sub\r\n");
        }

        body.Append("    Err.Raise 5, \"XLIDE.TestDispatch\", \"Unknown test target: \" & macroName\r\n");
        body.Append("    Exit Sub\r\n");
        body.Append("Caught:\r\n");
        body.Append($"    {AssertModuleName}.RecordTargetOutcome Err.Number, Err.Source, Err.Description\r\n");
        body.Append("End Sub\r\n");

        for (var chunk = 0; chunk < Math.Max(1, chunks); chunk++)
        {
            body.Append($"\r\nPrivate Function XlideDispatch{chunk}(ByVal targetKey As String) As Boolean\r\n");
            body.Append($"    XlideDispatch{chunk} = True\r\n");
            body.Append("    Select Case targetKey\r\n");
            foreach (var (module, procedure) in targets.Skip(chunk * chunkSize).Take(chunkSize))
            {
                var keys = $"\"{module}.{procedure}\"";
                if (bareCounts[procedure] == 1)
                {
                    keys += $", \"{procedure}\"";
                }

                body.Append($"        Case {keys}\r\n");
                body.Append($"            {module}.{procedure}\r\n");
            }

            body.Append("        Case Else\r\n");
            body.Append($"            XlideDispatch{chunk} = False\r\n");
            body.Append("    End Select\r\n");
            body.Append("End Function\r\n");
        }

        return body.ToString();
    }

    private const string JsonEscapeFunction =
        "Private Function JsonEscape(ByVal Value As String) As String\r\n"
        + "    Dim escaped As String\r\n"
        + "    escaped = Replace(Value, Chr$(92), Chr$(92) & Chr$(92))\r\n"
        + "    escaped = Replace(escaped, Chr$(34), Chr$(92) & Chr$(34))\r\n"
        + "    escaped = Replace(escaped, vbCrLf, Chr$(92) & \"n\")\r\n"
        + "    escaped = Replace(escaped, vbCr, Chr$(92) & \"n\")\r\n"
        + "    escaped = Replace(escaped, vbLf, Chr$(92) & \"n\")\r\n"
        + "    JsonEscape = escaped\r\n"
        + "End Function\r\n";

    // ---------------------------------------------------------------- execution

    /// <summary>
    /// Runs the selected tests in the live project, one Application.Run per test, reporting
    /// each result through <paramref name="landed"/> the moment it exists. The runner and the
    /// dispatcher are injected for the run and removed in the finally; XlideAssert stays,
    /// because the developer's tests compile against it.
    /// </summary>
    internal static List<TestResult> Run(
        DispatchObject project,
        IReadOnlyList<(string Name, string Source)> modules,
        IReadOnlyList<TestCase> discovered,
        Selection selection,
        bool failFast,
        Action<string> starting,
        Action<TestResult> landed)
    {
        var results = new List<TestResult>();
        var chosen = discovered.Where(test => Selected(test, selection)).ToList();

        // Skips never reach the host, and land first so the panel says why at once.
        foreach (var skip in chosen.Where(one => one.SkipReason is not null))
        {
            var row = Resolve(skip, "skipped",
                skip.SkipReason is { Length: > 0 } why ? why : "Skipped by its directive.", 0, []);
            results.Add(row);
            landed(row);
        }

        var runnable = chosen.Where(one => one.SkipReason is null).ToList();
        if (runnable.Count == 0)
        {
            return results;
        }

        if (IdentifierComplaint(runnable) is { } complaint)
        {
            foreach (var test in runnable)
            {
                var row = Resolve(test, "error", complaint, 0, []);
                results.Add(row);
                landed(row);
            }

            return results;
        }

        using var components = project.GetObject("VBComponents");
        if (components is null)
        {
            throw new InvalidOperationException("The project would not answer its components.");
        }

        var runnerName = $"{RunnerModulePrefix}{Environment.TickCount64.ToString("x", CultureInfo.InvariantCulture).PadLeft(8, '0')[^8..]}";
        RemoveGeneratedRunModules(components);

        try
        {
            AddModule(components, runnerName, BuildRunnerModule(runnable));
            AddModule(components, DispatchModuleName, BuildDispatchModule(modules));

            using var application = HostApplication.Find()
                ?? throw new InvalidOperationException("The host application could not be reached.");

            var file = SafeFileName(project);
            var target = file is null || Engine.HostApp.Name == "word"
                ? $"{runnerName}.RunTest"
                : $"'{file}'!{runnerName}.RunTest";

            var stopped = false;
            foreach (var test in runnable)
            {
                if (stopped)
                {
                    var held = Resolve(test, "skipped", "Not run because fail-fast stopped the run.", 0, []);
                    results.Add(held);
                    landed(held);
                    continue;
                }

                starting(test.Id);
                var watch = System.Diagnostics.Stopwatch.StartNew();
                TestResult row;
                try
                {
                    var answer = application.CallToString("Run", target, test.Id);
                    watch.Stop();
                    row = FromRunnerJson(test, answer, watch.Elapsed.TotalMilliseconds);
                }
                catch (Exception ex)
                {
                    watch.Stop();
                    // Run itself refused: a compile error in the project, macros disabled, or
                    // the developer pressed Reset mid-test. The editor's message is the answer.
                    row = Resolve(test, "error", ex.Message.Trim(), watch.Elapsed.TotalMilliseconds, []);
                }

                results.Add(row);
                landed(row);
                if (failFast && row.Status is "failed" or "error" or "xpass")
                {
                    stopped = true;
                }
            }
        }
        finally
        {
            try
            {
                RemoveGeneratedRunModules(components);
            }
            catch (Exception ex)
            {
                Log.Warn($"tests: the generated run modules could not be removed, {ex.Message}");
            }
        }

        return results;
    }

    /// <summary>
    /// Runs ONE test with no trap around it: an error or a breakpoint drops the editor into
    /// its own debugger, exactly as F5 on the Sub would. The call does not return until the
    /// debug session ends, which is the honest shape of debugging.
    /// </summary>
    internal static string DebugTest(DispatchObject project, TestCase test)
    {
        using var application = HostApplication.Find();
        if (application is null)
        {
            return "the host application could not be reached";
        }

        var file = SafeFileName(project);
        var target = file is null || Engine.HostApp.Name == "word"
            ? $"{test.Module}.{test.Procedure}"
            : $"'{file}'!{test.Module}.{test.Procedure}";
        application.Invoke("Run", target);
        return $"ran {test.Id} under the debugger";
    }

    private static bool Selected(TestCase test, Selection selection)
    {
        if (selection.Test is { Length: > 0 }
            && !string.Equals(test.Id, selection.Test, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (selection.Module is { Length: > 0 }
            && !string.Equals(test.Module, selection.Module, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return selection.Tags.Length == 0
            || selection.Tags.All(wanted => test.Tags.Contains(wanted, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>The two-shape runner JSON becomes a result row, with the expected-error and xfail flips applied.</summary>
    internal static TestResult FromRunnerJson(TestCase test, string answer, double durationMs)
    {
        string outcome;
        long number = 0;
        string source = string.Empty, message = string.Empty;
        string[] output = [];
        try
        {
            using var parsed = JsonDocument.Parse(answer);
            var root = parsed.RootElement;
            outcome = root.TryGetProperty("outcome", out var outcomeValue) ? outcomeValue.GetString() ?? "" : "";
            if (root.TryGetProperty("number", out var numberValue))
            {
                number = numberValue.GetInt64();
            }

            if (root.TryGetProperty("source", out var sourceValue))
            {
                source = sourceValue.GetString() ?? string.Empty;
            }

            if (root.TryGetProperty("message", out var messageValue))
            {
                message = messageValue.GetString() ?? string.Empty;
            }

            if (root.TryGetProperty("output", out var outputValue) && outputValue.ValueKind == JsonValueKind.Array)
            {
                output = [.. outputValue.EnumerateArray().Select(one => one.GetString() ?? string.Empty)];
            }
        }
        catch (JsonException)
        {
            return Resolve(test, "error",
                answer.Length == 0
                    ? "The test answered nothing - it may have been reset mid-run."
                    : $"The runner answered something that is not a result: {Truncate(answer)}",
                durationMs, []);
        }

        var passed = outcome == "passed";
        var failureText = source == "XLIDE.Assert" || number == 0
            ? message
            : $"VBA error {number} from {source}: {message}";

        // expected-error inverts the verdict before xfail gets its say, in that order.
        if (test.ExpectedError is { Length: > 0 } expected)
        {
            if (passed)
            {
                passed = false;
                failureText = expected == "any"
                    ? "Expected a VBA error, but no error was raised."
                    : $"Expected VBA error {expected}, but no error was raised.";
            }
            else if (expected == "any"
                || (long.TryParse(expected, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var wanted) && number == wanted))
            {
                passed = true;
                failureText = string.Empty;
            }
            else
            {
                failureText = number == 0
                    ? $"Expected VBA error {expected}, but got a failure without a VBA error number: {message}"
                    : $"Expected VBA error {expected}, but got VBA error {number}: {message}";
            }
        }

        if (test.XfailReason is not null)
        {
            return passed
                ? Resolve(test, "xpass",
                    $"Expected failure did not occur{(test.XfailReason.Length > 0 ? $": {test.XfailReason}" : ".")}",
                    durationMs, output)
                : Resolve(test, "xfail", failureText, durationMs, output);
        }

        return passed
            ? Resolve(test, "passed", null, durationMs, output)
            : Resolve(test, "failed", failureText, durationMs, output);
    }

    private static string Truncate(string text) =>
        text.Length <= 200 ? text : $"{text[..200]}...";

    private static TestResult Resolve(TestCase test, string status, string? message, double durationMs, string[] output) =>
        new(test.Id, test.Module, test.Procedure, test.Line, status, message, durationMs, output, test.Tags);

    private static void AddModule(DispatchObject components, string name, string body)
    {
        using var component = components.CallObject("Add", StandardModule)
            ?? throw new InvalidOperationException($"The project would not accept the {name} module.");
        component.SetString("Name", name);
        using var code = component.GetObject("CodeModule");
        code?.Invoke("AddFromString", body);
    }

    private static void RemoveGeneratedRunModules(DispatchObject components)
    {
        var count = components.GetInt32("Count");
        for (var i = count; i >= 1; i--)
        {
            using var candidate = components.GetItem(i);
            var name = candidate?.GetString("Name") ?? string.Empty;
            if (candidate is not null
                && (name.StartsWith(RunnerModulePrefix, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, DispatchModuleName, StringComparison.OrdinalIgnoreCase)))
            {
                components.InvokeWithObject("Remove", candidate);
            }
        }
    }

    private static string? SafeFileName(DispatchObject? project)
    {
        try
        {
            var full = project?.GetString("FileName");
            return string.IsNullOrEmpty(full) ? null : Path.GetFileName(full);
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>True when the project has left design mode - a test left something running.</summary>
    internal static bool LeftDesignMode(DispatchObject? project)
    {
        try
        {
            return project?.GetInt32("Mode") != DesignMode;
        }
        catch (Exception)
        {
            return false;
        }
    }
}
