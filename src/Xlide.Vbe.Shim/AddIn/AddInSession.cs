using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// One connected lifetime of the add-in inside one editor instance.
///
/// The session owns every resource that must be released before the host tears down: automation
/// references, window hooks, tool windows, and the engine connection. It is stopped from
/// OnBeginShutdown, which is the last moment at which touching the object model is safe.
/// </summary>
internal sealed class AddInSession : IDisposable
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
    /// releases the hold and republishes from the unfiltered findings above — no re-analysis.
    /// </summary>
    private readonly ActiveLineHold _activeLineHold = new();

    /// <summary>
    /// What each module read back as the last time this add-in wrote it, keyed by
    /// <see cref="WrittenKey"/> — one baseline per (workbook, module), never per bare name.
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
    /// A change from the page's dialog: adopted, written through, and echoed back — the echo is
    /// the page's confirmation that the choice will survive the session.
    /// </summary>
    private void OnSettingsChanged(ProductSettings updated)
    {
        _settings = updated.Normalized();

        try
        {
            var path = SettingsPath;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, _settings.ToJson());
            Log.Info($"settings: saved (blockLayout {_settings.BlockLayout}"
                + $", continueComment {_settings.ContinueCommentOnNewline}"
                + $", mirrorSpacing {_settings.MirrorCommentSpacing})");
        }
        catch (Exception ex)
        {
            Log.Error("settings: could not be written; the choice holds for this session only", ex);
        }

        _editorSurface?.ShowSettings(_settings);
    }

    private bool _stopped;

    public AddInSession(DispatchObject editor, DispatchObject? addIn)
    {
        _editor = editor;
        _addIn = addIn;
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
        ReportOpenProjects();
        TrackCodePanes();
        StartAnalysis();

        // An editor with no panes at all never fires the pane events the surface normally
        // arrives on, and a developer opening a fresh workbook's editor would meet the native
        // gray shell. The surface goes up now, showing the empty workspace and the explorer.
        TryShowEmptyWorkspace();

#if DEBUG
        // The dev build's local door for the harness: state, windows, commands by name.
        // Compiled out of Release entirely. Re-landed from post-v010-experiments after the
        // crash storm that retired the first landing was root-caused (the 16-byte variant,
        // lesson 33); the locals and watches routes now read the ghost reader thread's
        // published snapshots, which did not exist the first time.
        _debugServer = DebugServer.Start(AnswerDebugRequest);
#endif
    }

#if DEBUG
    private DebugServer? _debugServer;

    private static DebugServer.DebugReply DebugError(string error) =>
        DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
            new DebugErrorReply(error), DebugJsonContext.Default.DebugErrorReply));

    /// <summary>A request's wait budget, clamped to something a stuck page cannot outlast.</summary>
    private static int WaitMilliseconds(DebugServer.DebugRequest request, int fallback)
    {
        ArgumentNullException.ThrowIfNull(request);
        return request.Query.TryGetValue("waitMs", out var text) && int.TryParse(text, out var asked)
            ? Math.Clamp(asked, 100, 120_000)
            : fallback;
    }

    /// <summary>A string as a JavaScript (JSON) literal, quotes included.</summary>
    private static string JsonString(string value) =>
        "\"" + System.Text.Json.JsonEncodedText.Encode(value) + "\"";

    /// <summary>Identifies each pending script, so two callers cannot collect each other's.</summary>
    private static int _pageScriptId;

    /// <summary>
    /// Runs a script in a page and answers with its result as JSON.
    ///
    /// A PROMISE is awaited rather than reported as an empty object, which is what the
    /// browser's own ExecuteScript does with one: every async probe written against this
    /// door returned `{}` and looked like a page fault until the shape was recognised
    /// (2026-08-06). The script is evaluated inside a wrapper that stashes a pending promise
    /// on the page and hands back a ticket; the ticket is collected on a poll until the
    /// promise settles or the budget runs out.
    ///
    /// Only the START of each call crosses to the host thread. The browser delivers its
    /// answer by calling back on that same thread, so THIS thread — a pool thread — is the
    /// one that waits; waiting on the host thread would be waiting for a callback that
    /// cannot arrive until the waiting stops.
    /// </summary>
    private (bool Answered, int ErrorCode, string Result, string? Error) RunPageScript(
        string script, string? surface, int budgetMs)
    {
        var host = _editorSurface;
        if (host is null)
        {
            return (false, 0, string.Empty, "the surface is not up yet");
        }

        var ticket = Interlocked.Increment(ref _pageScriptId);
        var wrapper = $$"""
            (function () {
              var script = {{JsonString(script)}};
              var id = {{ticket}};
              window.__xlideEval = window.__xlideEval || {};
              try {
                var value = (0, eval)(script);
                if (value && typeof value.then === "function") {
                  window.__xlideEval[id] = { pending: true };
                  Promise.resolve(value).then(
                    function (settled) { window.__xlideEval[id] = { value: settled === undefined ? null : settled }; },
                    function (failed) { window.__xlideEval[id] = { error: String((failed && failed.message) || failed) }; });
                  return JSON.stringify({ pending: id });
                }
                return JSON.stringify({ value: value === undefined ? null : value });
              } catch (error) {
                return JSON.stringify({ error: String((error && error.message) || error) });
              }
            })()
            """;

        var first = RunPageScriptOnce(wrapper, surface, Math.Min(budgetMs, 10_000));
        if (first.Error is not null)
        {
            return first;
        }

        var opened = ReadWrapped(first.Result);
        if (!opened.Pending)
        {
            return (first.Answered, first.ErrorCode, opened.Payload, null);
        }

        // A promise: collect the ticket until it settles. The page keeps running between
        // polls, which is the whole point — this thread is not holding anything it needs.
        var deadline = Environment.TickCount64 + budgetMs;
        var collector = $$"""
            (function () {
              var id = {{ticket}};
              var held = (window.__xlideEval || {})[id];
              if (!held) { return JSON.stringify({ error: "the pending result was lost" }); }
              if (held.pending) { return JSON.stringify({ pending: id }); }
              delete window.__xlideEval[id];
              return JSON.stringify(held);
            })()
            """;

        /*
         * The wait BACKS OFF rather than sitting at 40ms.
         *
         * A flat 40ms before the first poll put a floor of about 70ms under every promise this
         * door returns — the initial call, the sleep, and the poll, each of the last two a page
         * round trip. Most of what comes back is a hover or a completion that settled in single
         * digits and then waited four times as long to be collected.
         *
         * It showed up as a scaling curve that would not move: hover measured 77ms on a
         * 109-line module and 77ms on an 11,000-line one, while the analyzer's own share went
         * from 1ms to 9ms (2026-08-08). The door was the whole measurement.
         *
         * Starting at 2ms and doubling to 40ms keeps a fast answer fast without spinning on a
         * slow one: a 200ms analysis is polled about a dozen times rather than five, which costs
         * nothing anybody can measure.
         */
        var pause = 2;

        while (Environment.TickCount64 < deadline)
        {
            Thread.Sleep(pause);
            pause = Math.Min(40, pause * 2);

            var poll = RunPageScriptOnce(collector, surface, 3000);
            if (poll.Error is not null)
            {
                return poll;
            }

            var settled = ReadWrapped(poll.Result);
            if (!settled.Pending)
            {
                return (true, poll.ErrorCode, settled.Payload, null);
            }
        }

        return (false, 0, string.Empty, $"the script did not settle within {budgetMs}ms");
    }

    /// <summary>Unwraps one wrapper answer: its payload, and whether it is still a ticket.</summary>
    private static (bool Pending, string Payload) ReadWrapped(string answer)
    {
        try
        {
            // The wrapper returns a STRING, so the browser's answer is that string as JSON.
            using var outer = System.Text.Json.JsonDocument.Parse(answer);
            var inner = outer.RootElement.ValueKind == System.Text.Json.JsonValueKind.String
                ? outer.RootElement.GetString() ?? "{}"
                : answer;

            using var parsed = System.Text.Json.JsonDocument.Parse(inner);
            if (parsed.RootElement.TryGetProperty("pending", out _))
            {
                return (true, string.Empty);
            }

            if (parsed.RootElement.TryGetProperty("error", out var error))
            {
                return (false, System.Text.Json.JsonSerializer.Serialize(
                    new DebugErrorReply(error.GetString() ?? "the script failed"),
                    DebugJsonContext.Default.DebugErrorReply));
            }

            return (false, parsed.RootElement.TryGetProperty("value", out var value)
                ? value.GetRawText()
                : "null");
        }
        catch (Exception)
        {
            // An answer that is not the wrapper's shape is passed through as it came: a
            // caller reading a raw result is better served than one told nothing.
            return (false, answer);
        }
    }

    /// <summary>One ExecuteScript round trip, with the host-thread hop and its deadlines.</summary>
    private (bool Answered, int ErrorCode, string Result, string? Error) RunPageScriptOnce(
        string script, string? surface, int budgetMs)
    {
        var host = _editorSurface;
        if (host is null)
        {
            return (false, 0, string.Empty, "the surface is not up yet");
        }

        string? result = null;
        var errorCode = 0;
        var started = false;
        using var scriptDone = new ManualResetEventSlim(false);
        using var scheduled = new ManualResetEventSlim(false);

        host.RunOnHostThread(() =>
        {
            try
            {
                var browser = surface == "palette" ? _browserPalette?.Browser : _editorSurface?.Browser;
                started = browser is not null && browser.ExecuteScript(script, (code, json) =>
                {
                    errorCode = code;
                    result = json;
                    scriptDone.Set();
                });
            }
            finally
            {
                scheduled.Set();
            }
        });

        if (!scheduled.Wait(TimeSpan.FromSeconds(3)))
        {
            return (false, 0, string.Empty, "the host thread did not start the script in time");
        }

        if (!started)
        {
            return (false, 0, string.Empty, "that surface has no page to run script in");
        }

        var answered = scriptDone.Wait(budgetMs);
        return (answered, errorCode, result ?? string.Empty, null);
    }

    /// <summary>
    /// Puts a small ring buffer in front of the page's console, so the `console` route can
    /// answer what the page said. Installed at every ready, including a reload's — the page
    /// that comes back is a new one and carries none of this.
    ///
    /// The console is WRAPPED rather than replaced: everything still reaches DevTools when a
    /// client is attached. Only uncaught errors go to the shim log (bounded, deliberately);
    /// this is for the rest, which is otherwise invisible during a live test.
    /// </summary>
    private void InstallConsoleRing()
    {
        const string install = """
            (function () {
              if (window.__xlideConsoleInstalled) { return "already"; }
              window.__xlideConsoleInstalled = true;
              window.__xlideConsole = [];

              ["log", "info", "warn", "error", "debug"].forEach(function (level) {
                var original = console[level];
                console[level] = function () {
                  try {
                    var parts = [];
                    for (var i = 0; i < arguments.length; i++) {
                      var one = arguments[i];
                      parts.push(typeof one === "string" ? one
                        : (one && one.message) ? String(one.message)
                        : (function () { try { return JSON.stringify(one); } catch (e) { return String(one); } })());
                    }
                    window.__xlideConsole.push(level + ": " + parts.join(" "));
                    if (window.__xlideConsole.length > 500) { window.__xlideConsole.shift(); }
                  } catch (ignored) { }
                  return original.apply(console, arguments);
                };
              });

              return "installed";
            })()
            """;

        // Fire and forget on a pool thread: this runs from the ready handler, which is ON
        // the host thread, and the script's answer arrives on that same thread — waiting
        // here would wait for a callback that cannot arrive until the waiting stops.
        _ = Task.Run(() => RunPageScriptOnce(install, null, 4000));
    }

    /// <summary>
    /// Where a surface's page sits inside the captured frame, in the frame's own pixels.
    ///
    /// A page reports coordinates relative to its own client area and cannot see where that
    /// area is; a frame capture is the whole window. The difference between the two window
    /// rectangles is what turns one into the other.
    /// </summary>
    private (int X, int Y) SurfaceOriginInFrame(string? which)
    {
        var surfaceWindow = which == "palette" ? _browserPalette?.Handle ?? 0 : _editorSurface?.SurfaceWindow ?? 0;
        if (surfaceWindow == 0 || _frame == 0)
        {
            return (0, 0);
        }

        // The palette is captured as its own window, so its page starts at its own origin.
        if (which == "palette")
        {
            return (0, 0);
        }

        unsafe
        {
            Interop.Rect surfaceRect;
            Interop.Rect frameRect;
            if (!Interop.Win32.GetWindowRect(surfaceWindow, &surfaceRect)
                || !Interop.Win32.GetWindowRect(_frame, &frameRect))
            {
                return (0, 0);
            }

            return (surfaceRect.Left - frameRect.Left, surfaceRect.Top - frameRect.Top);
        }
    }

    /// <summary>When the page bundle beside the running shim was built, or "(unknown)".</summary>
    private static string BundleBuiltUtc()
    {
        var directory = Interop.ShimModule.Directory;
        var bundle = directory is null ? null : Path.Combine(directory, "ui", "editor", "dist", "editor.js");
        return bundle is not null && File.Exists(bundle)
            ? File.GetLastWriteTimeUtc(bundle).ToString("s", System.Globalization.CultureInfo.InvariantCulture)
            : "(unknown)";
    }

    /// <summary>
    /// Whether the page is running a bundle older than the one on disk — the question behind
    /// "why is my fix not in the page", which cost three rounds of confusion in one day. The
    /// page stamps itself to the second at build time; a stamp before the file's own write
    /// time by more than a minute means the browser is serving something cached.
    /// </summary>
    private static bool StampIsBehind(string pageStamp, string bundleStamp) =>
        DateTime.TryParse(pageStamp, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var page)
        && DateTime.TryParse(bundleStamp, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var bundle)
        && bundle - page > TimeSpan.FromMinutes(1);

    /// <summary>
    /// The whole visible arrangement, read from the page: the four dock sections with their
    /// group trees, the editor groups with what each shows, and the sizes. One answer in
    /// place of the dozen ad-hoc measurements this layout's development needed.
    /// </summary>
    private const string LayoutScript = """
        (function () {
          var round = function (n) { return Math.round(n); };
          var boxOf = function (el) {
            var b = el.getBoundingClientRect();
            return { x: round(b.x), y: round(b.y), w: round(b.width), h: round(b.height) };
          };

          var sections = ["left", "right", "top", "bottom"].map(function (side) {
            var dock = document.getElementById("dock-" + side);
            if (!dock || dock.hidden) { return { side: side, standing: false, groups: [] }; }
            var groups = [].slice.call(dock.querySelectorAll(".dock-group")).map(function (group) {
              return {
                tabs: [].slice.call(group.querySelectorAll(".panel-tab")).map(function (tab) {
                  return { pane: tab.dataset.panel, active: tab.classList.contains("active") };
                }),
                box: boxOf(group)
              };
            });
            return { side: side, standing: true, box: boxOf(dock), groups: groups };
          });

          var workspace = window.xlideBridge && window.xlideBridge.workspace;
          var editors = [];
          if (workspace) {
            editors = [].slice.call(document.querySelectorAll(".editor-group")).map(function (group) {
              return {
                tabs: [].slice.call(group.querySelectorAll(".tab")).map(function (tab) {
                  return {
                    module: tab.dataset.module,
                    project: tab.dataset.project || null,
                    active: tab.classList.contains("active"),
                    dirty: tab.classList.contains("dirty")
                  };
                }),
                activeGroup: group.classList.contains("active-group"),
                box: boxOf(group)
              };
            });
          }

          var area = document.getElementById("editor-area");
          var empty = document.getElementById("empty-view");

          // An OBJECT, not a string: the runner already carries the value across as JSON,
          // and stringifying here would deliver the whole answer as one quoted string that
          // every caller then has to unescape.
          return {
            sections: sections,
            editorGroups: editors,
            editorArea: area ? boxOf(area) : null,
            emptyWorkspace: !!(empty && !empty.hidden),
            documents: workspace ? window.xlideBridge.documents.all() : [],
            dragging: !!document.querySelector(".drag-dim")
          };
        })()
        """;

    /// <summary>
    /// Answers one debug-door request. Arrives on a pool thread. Routes that read files,
    /// ring buffers, or the reader thread's published snapshots answer right here; routes
    /// that read the session or drive the editor cross to the host thread and wait with a
    /// deadline. The immediate route only schedules - a statement that hits a breakpoint
    /// does not return until the developer continues, and an api that waited on it would
    /// jam.
    /// </summary>
    private DebugServer.DebugReply AnswerDebugRequest(DebugServer.DebugRequest request)
    {
        // Sweep FIRST, before the routes that answer without the host thread.
        //
        // The sweep used to sit below them, which read as "every request heals first" and was
        // not: dialogs, dismiss and guard all return before reaching it, and those are exactly
        // the routes a caller uses while something is standing. Armed and watching, the guard
        // therefore never ran once — fourteen seconds of polling with a modal on screen and an
        // empty cleared list (2026-08-07).
        //
        // The heartbeat is no help here and this is why it cannot be the trigger: a VBA modal
        // PUMPS messages, so the host thread kept answering in under 140ms the whole time it was
        // blocked. What is standing is the only evidence that something is standing.
        ClearDialogsWeRaised();

        switch (request.Route)
        {
            case "log":
            {
                var path = Log.Path;
                if (path is null)
                {
                    return DebugError("no log file");
                }

                var since = request.Query.TryGetValue("since", out var sinceText)
                    && long.TryParse(sinceText, out var parsed) ? parsed : 0;
                request.Query.TryGetValue("match", out var match);
                var max = request.Query.TryGetValue("max", out var maxText)
                    && int.TryParse(maxText, out var cap) ? Math.Clamp(cap, 1, 5000) : 500;

                // waitMs turns the log into something a probe can AWAIT. Without it every
                // test that cares about an event sleeps a guessed interval and greps, which
                // is slow when the guess is generous and flaky when it is not - the whole
                // sleep-and-hope class of harness bug. With it, "wait until the log says the
                // module was written, or three seconds pass" is one request that returns the
                // moment it is true.
                var waitMs = request.Query.TryGetValue("waitMs", out var waitText)
                    && int.TryParse(waitText, out var waited) ? Math.Clamp(waited, 0, 30000) : 0;

                var deadline = Environment.TickCount64 + waitMs;
                while (true)
                {
                    var (lines, next) = ReadLogSlice(path, since, match, max);
                    if (lines.Count > 0 || Environment.TickCount64 >= deadline)
                    {
                        return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                            new DebugLogReply([.. lines], next), DebugJsonContext.Default.DebugLogReply));
                    }

                    Thread.Sleep(100);
                }
            }

            case "messages":
            {
                var last = request.Query.TryGetValue("last", out var lastText)
                    && int.TryParse(lastText, out var parsed) ? Math.Clamp(parsed, 1, 200) : 50;
                var rows = WebView.WebView2Surface.MessageTap.Snapshot(last)
                    .Select(entry => new DebugMessageRow(entry.Seq, entry.At, entry.Surface, entry.Direction, entry.Text))
                    .ToArray();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugMessagesReply(rows), DebugJsonContext.Default.DebugMessagesReply));
            }

            case "capture":
            {
                request.Query.TryGetValue("window", out var which);
                var target = which switch
                {
                    "palette" => _browserPalette?.Handle ?? 0,
                    _ => _frame,
                };
                var bytes = DebugCapture.CaptureBmp(target);
                if (bytes is null)
                {
                    return DebugError($"window {which ?? "frame"} would not render");
                }

                // With a selector, the picture is cut down to that element. A whole frame is
                // a big image in which a 54-pixel drop zone cannot be seen, and a surface
                // built by reading numbers rather than looking at it is built with one eye
                // shut (2026-08-06). The page says where the element is; the crop is here,
                // because the pixels are here.
                if (request.Query.TryGetValue("selector", out var cropSelector) && cropSelector.Length > 0)
                {
                    var pad = request.Query.TryGetValue("pad", out var padText) && int.TryParse(padText, out var asked)
                        ? Math.Clamp(asked, 0, 200)
                        : 8;

                    var where = RunPageScript(
                        $$"""
                        (function () {
                          var element = document.querySelector({{JsonString(cropSelector)}});
                          if (!element) { return null; }
                          var box = element.getBoundingClientRect();
                          // Page coordinates plus the browser's own origin on screen: the
                          // page cannot see where its window is, so the host adds that.
                          return {
                            x: Math.round(box.x), y: Math.round(box.y),
                            w: Math.round(box.width), h: Math.round(box.height)
                          };
                        })()
                        """,
                        which,
                        4000);

                    if (where.Error is not null || where.Result.Trim() is "null" or "")
                    {
                        return DebugError($"nothing matches {cropSelector} on that surface");
                    }

                    try
                    {
                        using var box = System.Text.Json.JsonDocument.Parse(where.Result);
                        var x = box.RootElement.GetProperty("x").GetInt32();
                        var y = box.RootElement.GetProperty("y").GetInt32();
                        var w = box.RootElement.GetProperty("w").GetInt32();
                        var h = box.RootElement.GetProperty("h").GetInt32();

                        // The page's coordinates are relative to the BROWSER's client area,
                        // which sits at the surface's own origin inside the frame.
                        var origin = SurfaceOriginInFrame(which);
                        var cropped = DebugCapture.CropBmp(
                            bytes, 0, 0,
                            origin.X + x - pad, origin.Y + y - pad,
                            w + pad * 2, h + pad * 2);

                        return cropped is null
                            ? DebugError($"{cropSelector} is not on screen")
                            : new DebugServer.DebugReply("image/bmp", cropped);
                    }
                    catch (Exception ex)
                    {
                        return DebugError($"the element's box could not be read ({ex.GetType().Name})");
                    }
                }

                return new DebugServer.DebugReply("image/bmp", bytes);
            }

            /*
             * THE OUTCOME, not the request.
             *
             * This posted the line to the host thread and answered `{ran: true}` without waiting,
             * so a caller learned that the evaluation had been ASKED FOR and nothing else. What
             * the expression came to, and whether it failed, went only to the page. That is why
             * the Immediate window had a route and no suite: nothing could read what it said.
             *
             * The same rule the rest of this door already follows. `closeActive` reports whether
             * the tab actually closed rather than that a close was requested; `compile` answers
             * the errors as data rather than leaving them on screen. A route that reports its own
             * invocation is a route that cannot be asserted on.
             *
             * Waits, therefore, and answers the evaluator's own verdict. Without `text` it READS
             * instead: the whole window as it stands, which is the other half nobody had.
             */
            case "immediate":
            {
                var surface = _editorSurface;
                if (surface is null)
                {
                    return DebugError("the surface is not up yet");
                }

                request.Query.TryGetValue("text", out var text);

                if (string.IsNullOrEmpty(text))
                {
                    return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                        new DebugImmediateReply(false, _immediateReader?.Text() ?? string.Empty, false),
                        DebugJsonContext.Default.DebugImmediateReply));
                }

                /*
                 * Started on the host thread and NOT waited on there, then answered from here,
                 * which is the shape `compile` already uses and for the same reason.
                 *
                 * A line that will not compile raises the editor's own "Compile error" box. That
                 * box owns the host thread, so anything waiting on that thread waits for the box,
                 * and the box is waiting for somebody to press OK. The door's thread is the only
                 * one still moving, so it is the one that has to notice and clear it.
                 *
                 * The first version of this waited on an event and reported a timeout. It made
                 * things worse rather than better: the request returned after ten seconds, the
                 * dialog stopped being one this request had raised, and nothing cleared it at all
                 * -- so a mistyped line left a modal standing in front of the editor for the rest
                 * of the session instead of for thirteen seconds.
                 */
                /*
                 * A SESSION THIS PRODUCT LEFT STOPPED IS CLEARED HERE, ON THE DOOR'S THREAD.
                 *
                 * THE ROOT CAUSE, which three attempts on the host thread could not reach. When a
                 * line will not compile, the editor's "Compile error" box goes up and dismissing
                 * it leaves VBA stopped INSIDE the scratch procedure, with `Application.Run`
                 * suspended mid-call. A suspended frame unwinds only when the host thread returns
                 * to its message loop -- so a recovery running ON that thread, inside a
                 * RunOnHostThread callback, is holding the one thing that has to happen for the
                 * recovery to work. Reset was issued, repeatedly, for eight seconds, and could not
                 * take: not because the budget was short but because no budget can be long enough
                 * when waiting is itself what prevents the wait from ending (2026-08-07).
                 *
                 * Issued from here it is an ordinary request, the host thread goes back to its
                 * pump between calls, the frame unwinds, and the mode is design again. The polling
                 * below is not a timing guess for the same reason `compile` polls: the door's
                 * thread is the only one still moving, and what it waits for can actually happen
                 * while it waits.
                 */
                if (ScratchBreakStanding())
                {
                    Log.Info("immediate: the editor is stopped in the scratch module, clearing it");
                    surface.RunOnHostThread(() => ExecuteEditorCommand(VbeCommands.Command.Reset));

                    var clearBy = Environment.TickCount64 + 5000;
                    while (Environment.TickCount64 < clearBy && ScratchBreakStanding())
                    {
                        Thread.Sleep(100);
                    }

                    if (ScratchBreakStanding())
                    {
                        var stuck = "The last line left the editor stopped inside this product's "
                            + "own scratch module and it could not be cleared. Press Reset in the "
                            + "editor.";

                        Log.Warn($"immediate: {stuck}");
                        return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                            new DebugImmediateReply(false, stuck, true),
                            DebugJsonContext.Default.DebugImmediateReply));
                    }

                    surface.RunOnHostThread(RemoveScratchModule);
                }

                var raisedBefore = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);

                var evaluated = new ManualResetEventSlim(false);
                var outcome = string.Empty;
                var failed = false;

                surface.RunOnHostThread(() =>
                {
                    try
                    {
                        var result = EvaluateImmediate(text);
                        outcome = result.Text;
                        failed = result.Failed;
                    }
                    finally
                    {
                        evaluated.Set();
                    }
                });

                var complained = new List<string>();
                var deadline = Environment.TickCount64 + WaitMilliseconds(request, 15000);

                while (Environment.TickCount64 < deadline && !evaluated.IsSet)
                {
                    evaluated.Wait(120);

                    foreach (var raised in DialogWatch.Dialogs())
                    {
                        if (!raisedBefore.Add(raised.Window))
                        {
                            continue;
                        }

                        // The box's own words ARE the answer. A compile error says what is wrong
                        // with the line, which is exactly what the developer typed it to find out,
                        // and it is more use in the panel than on top of the editor.
                        complained.Add(raised.Text.Length > 0 ? raised.Text : raised.Caption);

                        var pressed = DialogWatch.SafeAnswerFor(raised) ?? "OK";
                        DialogWatch.Dismiss(raised.Caption, pressed);
                        Log.Info($"immediate: \"{raised.Text}\" answered with {pressed}");
                    }
                }

                var ran = evaluated.Wait(2000);
                evaluated.Dispose();

                // What the editor complained about outranks what the evaluator managed to return.
                // A cleared compile box leaves the run answering an empty string, which reads as a
                // successful evaluation of nothing.
                if (complained.Count > 0)
                {
                    outcome = string.Join(" ", complained).Replace("\r", " ").Replace("\n", " ").Trim();
                    failed = true;
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugImmediateReply(ran, outcome, failed),
                    DebugJsonContext.Default.DebugImmediateReply));
            }

            case "locals":
            {
                // Straight from the reader thread's published snapshot: an immutable record
                // behind a volatile read, safe from any thread, and exactly what the panel
                // renders. The first landing kept mirror fields; the thread made them
                // unnecessary.
                var snapshot = _ghostReaders?.Locals;
                var rows = snapshot is null ? [] : new SurfaceLocalRow[snapshot.Rows.Count];
                for (var i = 0; i < rows.Length; i++)
                {
                    var row = snapshot!.Rows[i];
                    rows[i] = new SurfaceLocalRow(row.Expression, row.Value, row.Type);
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugLocalsReply(snapshot?.Context, rows),
                    DebugJsonContext.Default.DebugLocalsReply));
            }

            case "watches":
            {
                var reading = _ghostReaders?.Watches;
                var rows = reading is null ? [] : new SurfaceWatchRow[reading.Count];
                for (var i = 0; i < rows.Length; i++)
                {
                    var row = reading![i];
                    rows[i] = new SurfaceWatchRow(row.Expression, row.Value, row.Type, row.Context);
                }

                // _inBreak is written on the host thread; a bool read is atomic and a poll
                // tick of staleness is the nature of this api.
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugWatchesReply(_inBreak, rows),
                    DebugJsonContext.Default.DebugWatchesReply));
            }

            case "problems":
            {
                // The findings list is replaced whole, never mutated, so a reference read from
                // this thread sees a complete edition.
                var held = _findings;
                request.Query.TryGetValue("module", out var onlyModule);
                var rows = held
                    .Where(finding => onlyModule is null
                        || string.Equals(finding.Module, onlyModule, StringComparison.OrdinalIgnoreCase))
                    .Select(finding => new DebugFindingRow(
                        finding.Module, finding.StartLine, finding.StartColumn,
                        finding.Severity, finding.Code ?? string.Empty, finding.Message))
                    .ToArray();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugProblemsReply(rows), DebugJsonContext.Default.DebugProblemsReply));
            }

            case "drainfinalizers":
            {
                /*
                 * MAKES A LEAKED WRAPPER FAIL NOW, WHERE IT CAN BE ATTRIBUTED.
                 *
                 * A COM wrapper nothing disposed is released by the FINALIZER thread, and for
                 * the editor's objects that is an access violation the runtime cannot throw: it
                 * FailFasts, and Excel goes with it. The damage arrives whenever a collection
                 * happens to run, which is minutes after whatever created the wrapper and in a
                 * stack that names nothing about it. Three crashes on 2026-08-07 and two more on
                 * 2026-08-08 were all read as unrelated because of that delay.
                 *
                 * This collapses the delay. Run an operation, call this, and if the host dies
                 * then THAT operation created the wrapper. It is the bisecting tool the previous
                 * hunts did not have.
                 *
                 * NOT A LEAK COUNTER, and the distinction is the whole reason the last attempt at
                 * a `gc` route was deleted: that one reported a clean bill of health while 8,734
                 * wrappers were pending, because it measured the heap rather than the outcome.
                 * This measures the OUTCOME - the host is alive afterwards, or it is not - and
                 * says nothing else. `stats.comWrappersLive` is still where a count comes from.
                 */
                var before = Com.ComRuntime.WrappersLive;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugDrainReply(
                        WrappersLiveBefore: before,
                        WrappersLiveAfter: Com.ComRuntime.WrappersLive,
                        Survived: true),
                    DebugJsonContext.Default.DebugDrainReply));
            }

            case "history":
            {
                // The session as a script. After a live investigation the useful sequence is
                // normally reconstructed from a scrollback and gets a step wrong; this hands
                // it back ready to run, so a bug found by hand becomes a probe by copying.
                var requests = DebugServer.Requests();
                var script = new System.Text.StringBuilder();
                script.AppendLine("# Replay of a debug api session. Point it at a live instance:");
                script.AppendLine("#   $api = \"http://127.0.0.1:PORT/TOKEN\"");
                foreach (var line in requests)
                {
                    var space = line.IndexOf(' ', StringComparison.Ordinal);
                    var verb = line[..space];
                    var rest = line[(space + 1)..];
                    script.AppendLine(verb == "POST"
                        ? $"Invoke-RestMethod \"$api/{rest}\" -Method Post -TimeoutSec 20"
                        : $"Invoke-RestMethod \"$api/{rest}\" -TimeoutSec 20");
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugHistoryReply(requests, script.ToString(), DebugServer.RouteCosts()),
                    DebugJsonContext.Default.DebugHistoryReply));
            }

            case "assert" when request.Query.TryGetValue("that", out var claim) && claim.Length > 0:
            {
                // A probe's expectation, stated once and waited on. Every check written here
                // so far has been the same four lines - poll a route, read a field, compare,
                // give up eventually - and each rewrite is a chance to sleep instead of wait
                // or to swallow the answer. The named claims are the ones the harness keeps
                // needing; anything more specific belongs in eval or a real test.
                var timeout = request.Query.TryGetValue("timeoutMs", out var timeoutText)
                    && int.TryParse(timeoutText, out var wanted) ? Math.Clamp(wanted, 0, 60000) : 10000;
                request.Query.TryGetValue("value", out var expected);

                var deadline = Environment.TickCount64 + timeout;
                string? saw = null;
                var held = false;

                while (true)
                {
                    (held, saw) = EvaluateClaim(claim, expected);
                    if (held || Environment.TickCount64 >= deadline)
                    {
                        break;
                    }

                    Thread.Sleep(150);
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugAssertReply(held, claim, expected ?? "(none)", saw ?? "(nothing)"),
                    DebugJsonContext.Default.DebugAssertReply));
            }

            case "journal":
            {
                // Everything a bug report needs, in one request, because evidence gathered
                // six requests apart is evidence of six different moments - and because the
                // moment worth capturing is usually the one already passing. State, the
                // dialogs standing, the counters, the recent log, and the last page traffic,
                // all read as close together as this door can manage.
                var lines = request.Query.TryGetValue("lines", out var linesText)
                    && int.TryParse(linesText, out var wanted) ? Math.Clamp(wanted, 1, 2000) : 200;

                var logLines = Log.Path is { } journalPath
                    ? ReadLogSlice(journalPath, TailOffset(journalPath, lines), null, lines).Lines
                    : [];

                var (placementSamples, marshalSamples) = PerfCounters.Samples();
                var messages = WebView.WebView2Surface.MessageTap.Snapshot(40)
                    .Select(entry => new DebugMessageRow(entry.Seq, entry.At, entry.Surface, entry.Direction, entry.Text))
                    .ToArray();

                // The session facts need the host thread, and the whole point of a journal is
                // that it still says something when that thread is busy - so a failure there
                // is reported inside the journal rather than instead of it.
                string? sessionState = null;
                if (_editorSurface is { } journalHost)
                {
                    using var ready = new ManualResetEventSlim(false);
                    journalHost.RunOnHostThread(() =>
                    {
                        try
                        {
                            sessionState = AnswerDebugRequestOnHost(
                                new DebugServer.DebugRequest("state", request.Query, string.Empty));
                        }
                        catch (Exception ex)
                        {
                            sessionState = $"{{\"error\":\"{ex.GetType().Name}\"}}";
                        }
                        finally
                        {
                            ready.Set();
                        }
                    });

                    if (!ready.Wait(TimeSpan.FromSeconds(3)))
                    {
                        sessionState = null;
                    }
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugJournalReply(
                        CapturedAt: DateTime.Now.ToString("O"),
                        Pid: Environment.ProcessId,
                        State: sessionState ?? "(the host thread did not answer)",
                        HeartbeatAgeMs: PerfCounters.HeartbeatAgeMs,
                        Dialogs: DialogWatch.Dialogs()
                            .Select(row => new DebugDialogRow(row.Window, row.Caption, row.Text, row.Buttons, row.Enabled))
                            .ToArray(),
                        PlacementMs: placementSamples,
                        MarshalMs: marshalSamples,
                        Messages: messages,
                        Log: [.. logLines]),
                    DebugJsonContext.Default.DebugJournalReply));
            }

            case "perf":
            {
                // Raw recent durations, so a probe can compute a median and a p95 rather
                // than reason from a running maximum that one outlier owns forever.
                var (placementSamples, marshalSamples) = PerfCounters.Samples();

                // reset=1 forgets the analyzer figures, so an experiment measures what it
                // provokes rather than everything since the editor opened. Session start is
                // the wrong window for "is THIS change slow".
                if (request.Query.TryGetValue("reset", out var perfReset) && perfReset != "0")
                {
                    EngineCounters.Reset();
                }

                var (engineMethods, engineSlowest, engineWindow) = EngineCounters.Snapshot();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugPerfReply(
                        placementSamples,
                        marshalSamples,
                        PerfCounters.HeartbeatAgeMs,
                        engineMethods,
                        engineSlowest,
                        engineWindow),
                    DebugJsonContext.Default.DebugPerfReply));
            }

            case "ui":
            {
                // The surface as the PAGE describes it: tabs with the labels the strip drew,
                // the tree's expansion, which panes and dialogs are up, what has not arrived
                // yet, where the caret is.
                //
                // The page answers because the page knows. Every earlier version of this
                // question was a querySelectorAll written fresh in whichever probe was asking,
                // and a scraped row cannot tell "collapsed" from "rendered wrong" — the render
                // being stale is the defect worth catching, and scraping it measures the wrong
                // half. See ui/editor/src/devsurface.ts.
                // line/column, or word, adds the `at` field: what is painted at that position
                // and what squiggles cover it. Asked for by argument rather than always, because
                // reading the rendered span means touching the DOM for a line that may not be on
                // screen, and most callers want the surface rather than one word of it.
                request.Query.TryGetValue("line", out var atLine);
                request.Query.TryGetValue("column", out var atColumn);
                request.Query.TryGetValue("word", out var atWord);

                var arguments = string.IsNullOrEmpty(atWord)
                    ? $"{(int.TryParse(atLine, out var l) ? l : 0)}, {(int.TryParse(atColumn, out var c) ? c : 1)}"
                    : $"null, null, {System.Text.Json.Nodes.JsonValue.Create(atWord)!.ToJsonString()}";

                var ui = RunPageScript($"window.xlideUi.state({arguments})", null, WaitMilliseconds(request, 5000));
                return ui.Error is { } uiError
                    ? DebugError(uiError)
                    : DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                        new DebugEvalReply(ui.Answered, ui.ErrorCode, ui.Result, Unwrap(ui.Result)),
                        DebugJsonContext.Default.DebugEvalReply));
            }

            case "act" when request.Query.TryGetValue("do", out var actionName) && actionName.Length > 0:
            {
                // The surface DRIVEN, through the methods a click reaches rather than through
                // synthesised events. The tab close box is why: it arms at pointerdown and
                // fires at pointerup, so `element.click()` on it does nothing, silently, and a
                // probe written that way reports a working feature broken (2026-08-07).
                //
                // Arguments ride as query values and arrive as strings; the page coerces them.
                var arguments = new System.Text.Json.Nodes.JsonObject();
                foreach (var (key, value) in request.Query)
                {
                    if (key is not ("do" or "token" or "waitMs"))
                    {
                        arguments[key] = value;
                    }
                }

                var quotedName = System.Text.Json.Nodes.JsonValue.Create(actionName)!.ToJsonString();
                var act = RunPageScript(
                    $"window.xlideUi.act({quotedName}, {arguments.ToJsonString()})",
                    null,
                    WaitMilliseconds(request, 8000));

                return act.Error is { } actError
                    ? DebugError(actError)
                    : DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                        new DebugEvalReply(act.Answered, act.ErrorCode, act.Result, Unwrap(act.Result)),
                        DebugJsonContext.Default.DebugEvalReply));
            }

            case "eval" when request.Body.Length > 0 || request.Query.ContainsKey("script"):
            {
                // The page's own DOM, asked directly: the questions that are one line ("how
                // many tabs does the strip show?", "is the empty view up?") answered without
                // a DevTools client. Pixels cannot answer those, and this needs no protocol.
                //
                // `ui` above answers most of them now, and better. Reach for this when the
                // question is genuinely new; if it gets asked twice, it belongs in devsurface.
                var script = request.Body.Length > 0 ? request.Body : request.Query["script"];
                request.Query.TryGetValue("surface", out var which);

                var run = RunPageScript(script, which, WaitMilliseconds(request, 5000));
                return run.Error is { } evalError
                    ? DebugError(evalError)
                    : DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                        new DebugEvalReply(run.Answered, run.ErrorCode, run.Result, Unwrap(run.Result)),
                        DebugJsonContext.Default.DebugEvalReply));
            }

            case "await" when request.Body.Length > 0 || request.Query.ContainsKey("script"):
            {
                // A condition, waited for IN the page rather than by a caller looping over
                // eval. Every such loop was a round trip per tick and a sleep chosen by
                // guess; the ones written during the workspace work raced the thing they
                // were watching more than once (2026-08-06). One request, one answer, and
                // the elapsed time says whether the condition was already true or arrived.
                var predicate = request.Body.Length > 0 ? request.Body : request.Query["script"];
                request.Query.TryGetValue("surface", out var awaitSurface);
                var budget = WaitMilliseconds(request, 10000);

                // The predicate is compiled to a function ONCE, here, while this script is
                // still the browser's own synchronous evaluation — which the page's content
                // policy exempts. Evaluating a string from inside a later timer callback is
                // not exempt and is refused outright ("unsafe-eval is not an allowed
                // source"), so a waiter that eval'd per tick never ran its predicate at all
                // and reported every condition as unmet (2026-08-06).
                var waiter = $$"""
                    (function () {
                      var test;
                      try {
                        test = (0, eval)("(function () { return (" + {{JsonString(predicate)}} + "); })");
                      } catch (error) {
                        return Promise.resolve({ met: false, elapsedMs: 0,
                          detail: "the predicate would not compile: " + String((error && error.message) || error) });
                      }

                      var deadline = Date.now() + {{budget}};
                      var started = Date.now();
                      return new Promise(function (resolve) {
                        (function tick() {
                          var met = false;
                          var detail = "";
                          try {
                            var value = test();
                            met = !!value;
                            detail = met ? "" : String(value);
                          } catch (error) {
                            detail = String((error && error.message) || error);
                          }
                          if (met || Date.now() > deadline) {
                            resolve({ met: met, elapsedMs: Date.now() - started, detail: detail });
                            return;
                          }
                          setTimeout(tick, 60);
                        })();
                      });
                    })()
                    """;

                // The page's own deadline expires first; the transport's is the backstop for
                // a page that stopped running timers at all.
                var awaited = RunPageScript(waiter, awaitSurface, budget + 4000);
                if (awaited.Error is { } awaitError)
                {
                    return DebugError(awaitError);
                }

                var met = false;
                var elapsed = 0;
                var detail = string.Empty;
                try
                {
                    using var parsed = System.Text.Json.JsonDocument.Parse(awaited.Result);
                    met = parsed.RootElement.TryGetProperty("met", out var metValue)
                        && metValue.ValueKind == System.Text.Json.JsonValueKind.True;
                    elapsed = parsed.RootElement.TryGetProperty("elapsedMs", out var elapsedValue)
                        && elapsedValue.TryGetInt32(out var ms) ? ms : 0;
                    detail = parsed.RootElement.TryGetProperty("detail", out var detailValue)
                        ? detailValue.GetString() ?? string.Empty : string.Empty;
                }
                catch (Exception ex)
                {
                    detail = $"the page's answer could not be read ({ex.GetType().Name})";
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugAwaitReply(met, elapsed, detail), DebugJsonContext.Default.DebugAwaitReply));
            }

            case "console":
            {
                // What the page said to itself. Only UNCAUGHT errors reach the shim log —
                // deliberately, because forwarding every line would drown it — so a warning
                // or a console.error the page handled is invisible without a DevTools client
                // attached, which is exactly the situation during a live test. The ring is
                // installed at page ready and read here.
                var last = request.Query.TryGetValue("last", out var lastText) && int.TryParse(lastText, out var asked)
                    ? Math.Clamp(asked, 1, 500)
                    : 100;

                var read = RunPageScript(
                    $$"""
                    (function () {
                      var ring = window.__xlideConsole;
                      if (!ring) { return { installed: false, lines: [] }; }
                      return { installed: true, lines: ring.slice(-{{last}}) };
                    })()
                    """,
                    null,
                    4000);

                return read.Error is { } consoleError
                    ? DebugError(consoleError)
                    : DebugServer.DebugReply.Json(read.Result);
            }

            case "inspect" when request.Query.TryGetValue("selector", out var selector) && selector.Length > 0:
            {
                // What the page actually has, where it is, and — with `styles` — what those
                // properties computed to, plus WHICH RULES claimed them.
                //
                // The rule list is the point. This page shares a document with a large
                // bundled stylesheet, and a structural class of ours (`.row` on a split
                // container) silently inherited `align-items: baseline` from an unrelated
                // rule, collapsing every cell to its tab strip's height. It read as a flex
                // bug in our own code and took an hour; the loop that finally found it —
                // walk every stylesheet, keep the rules this element matches — is this
                // route (2026-08-06).
                request.Query.TryGetValue("styles", out var wanted);
                var withRules = request.Query.TryGetValue("rules", out var rulesFlag) && rulesFlag != "0";
                var cap = request.Query.TryGetValue("max", out var maxText) && int.TryParse(maxText, out var asked)
                    ? Math.Clamp(asked, 1, 50)
                    : 10;

                var inspect = $$"""
                    (function () {
                      var found = [].slice.call(document.querySelectorAll({{JsonString(selector)}}));
                      var wanted = {{JsonString(wanted ?? string.Empty)}}
                        .split(",").map(function (one) { return one.trim(); }).filter(Boolean);

                      return {
                        selector: {{JsonString(selector)}},
                        matched: found.length,
                        elements: found.slice(0, {{cap}}).map(function (element) {
                          var box = element.getBoundingClientRect();
                          var computed = getComputedStyle(element);
                          var styles = {};
                          wanted.forEach(function (name) { styles[name] = computed.getPropertyValue(name); });

                          var rules = [];
                          if ({{(withRules ? "true" : "false")}}) {
                            for (var s = 0; s < document.styleSheets.length; s++) {
                              var sheet = document.styleSheets[s];
                              var list;
                              try { list = sheet.cssRules; } catch (blocked) { continue; }
                              for (var r = 0; r < list.length; r++) {
                                var rule = list[r];
                                if (!rule.selectorText) { continue; }
                                var matches = false;
                                try { matches = element.matches(rule.selectorText); } catch (bad) { matches = false; }
                                if (!matches) { continue; }
                                // Only rules that speak to a property being asked about, or
                                // every matching rule when nothing was named.
                                if (wanted.length === 0) { rules.push(rule.selectorText); continue; }
                                for (var w = 0; w < wanted.length; w++) {
                                  if (rule.style.getPropertyValue(wanted[w])) {
                                    rules.push(rule.selectorText + " { " + wanted[w] + ": "
                                      + rule.style.getPropertyValue(wanted[w]) + " }");
                                  }
                                }
                              }
                            }
                          }

                          return {
                            tag: element.tagName.toLowerCase(),
                            id: element.id || "",
                            classes: element.className && element.className.toString ? element.className.toString() : "",
                            hidden: !!element.hidden || computed.display === "none",
                            x: Math.round(box.x), y: Math.round(box.y),
                            w: Math.round(box.width), h: Math.round(box.height),
                            styles: styles,
                            rules: rules
                          };
                        })
                      };
                    })()
                    """;

                var inspected = RunPageScript(inspect, null, 5000);
                return inspected.Error is { } inspectError
                    ? DebugError(inspectError)
                    : DebugServer.DebugReply.Json(inspected.Result);
            }

            case "bench" when request.Query.TryGetValue("what", out var what) && what.Length > 0:
            {
                // Numbers for the things a developer feels, run enough times to have a
                // shape. The counters elsewhere say what the host spent; this says what the
                // SURFACE costs, which is where the risk moved when the workspace learned to
                // split and dock.
                var runs = request.Query.TryGetValue("n", out var runsText) && int.TryParse(runsText, out var count)
                    ? Math.Clamp(count, 1, 200)
                    : 20;

                var body = what switch
                {
                    // The live-model claim, measured: switching tabs should be page-local
                    // and free, where the old one-model surface paid a host round trip and
                    // a full document load for every switch.
                    "tabswitch" => """
                        var groups = window.xlideBridge.workspace;
                        var docs = window.xlideBridge.documents.all();
                        if (docs.length < 2) { return { detail: "needs two open documents", samples: [] }; }
                        var editor = groups.activeEditor();
                        var samples = [];
                        for (var i = 0; i < RUNS; i++) {
                          var target = docs[i % docs.length];
                          var model = window.xlideBridge.documents.get(target.module, target.project);
                          var began = performance.now();
                          editor.setModel(model);
                          editor.render(true);
                          samples.push(performance.now() - began);
                        }
                        return { detail: docs.length + " documents", samples: samples };
                        """,

                    // A full re-measure of every editor, which is what a splitter drag and
                    // every dock change costs.
                    "layout" => """
                        var editors = window.xlideBridge.workspace.editors();
                        var samples = [];
                        for (var i = 0; i < RUNS; i++) {
                          var began = performance.now();
                          editors.forEach(function (one) { one.layout(); });
                          samples.push(performance.now() - began);
                        }
                        return { detail: editors.length + " editor(s)", samples: samples };
                        """,

                    // Typing, as the editor sees it: an edit applied to the model and the
                    // page's own work to show it. The host's half is the write timer, which
                    // the log and the marshal counters already carry.
                    "type" => """
                        var editor = window.xlideBridge.workspace.activeEditor();
                        var model = editor.getModel();
                        if (!model) { return { detail: "no model", samples: [] }; }
                        var samples = [];
                        for (var i = 0; i < RUNS; i++) {
                          var began = performance.now();
                          model.pushEditOperations(null, [{
                            range: { startLineNumber: 1, startColumn: 1, endLineNumber: 1, endColumn: 1 },
                            text: "'x\n"
                          }], function () { return null; });
                          editor.render(true);
                          samples.push(performance.now() - began);
                        }
                        for (var u = 0; u < RUNS; u++) { model.undo(); }
                        return { detail: model.uri.toString(), samples: samples };
                        """,

                    _ => null,
                };

                if (body is null)
                {
                    return DebugError($"unknown benchmark {what}; try tabswitch, layout, or type");
                }

                var bench = RunPageScript(
                    $$"""
                    (function () {
                      var RUNS = {{runs}};
                      {{body}}
                    })()
                    """,
                    null,
                    Math.Max(15000, runs * 200));

                if (bench.Error is { } benchError)
                {
                    return DebugError(benchError);
                }

                var samples = new List<double>();
                var detail = string.Empty;
                try
                {
                    using var parsed = System.Text.Json.JsonDocument.Parse(bench.Result);
                    detail = parsed.RootElement.TryGetProperty("detail", out var detailValue)
                        ? detailValue.GetString() ?? string.Empty : string.Empty;
                    if (parsed.RootElement.TryGetProperty("samples", out var sampleValues))
                    {
                        foreach (var sample in sampleValues.EnumerateArray())
                        {
                            samples.Add(Math.Round(sample.GetDouble(), 3));
                        }
                    }
                }
                catch (Exception ex)
                {
                    return DebugError($"the benchmark's answer could not be read ({ex.GetType().Name})");
                }

                if (samples.Count == 0)
                {
                    return DebugError($"the benchmark ran nothing: {detail}");
                }

                var ordered = samples.OrderBy(one => one).ToArray();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugBenchReply(
                        what,
                        ordered.Length,
                        ordered[0],
                        ordered[ordered.Length / 2],
                        ordered[Math.Min(ordered.Length - 1, (int)(ordered.Length * 0.95))],
                        ordered[^1],
                        [.. samples],
                        detail),
                    DebugJsonContext.Default.DebugBenchReply));
            }

            case "trip" when request.Query.TryGetValue("what", out var tripWhat) && tripWhat.Length > 0:
            {
                // What a person WAITS for, which is never one layer.
                //
                // `bench` times the page's own work and `perf` reports the host's, and both
                // have looked healthy while the surface felt slow, because the cost was in
                // neither: it was the crossing between them, which nothing measured. These
                // are wall clock from asking to observable, taken here rather than on either
                // side, so the door's own cost is inside the number.
                //
                // `pagecall` is the floor. Every other figure here contains it, and without
                // it a 40ms feature and a 40ms door are the same reading.
                //
                // WHAT CANNOT BE MEASURED HERE, and it is a constraint on every route, not
                // just this one: a route body runs ON THE HOST THREAD, and a web message
                // posted to the page is delivered by that same thread's pump. So a body that
                // posts and then waits to see the effect waits forever - the post cannot be
                // delivered until the body returns. A caret trip written that way sat through
                // four seconds per sample and reported that the caret never moved, while the
                // identical sequence from OUTSIDE landed in 50ms on the first poll
                // (2026-08-07). Thread.Sleep does not help: it yields the CPU and pumps
                // nothing.
                //
                // RunPageScript survives this because ExecuteScript's answer comes back by a
                // path the blocked thread still completes; PostWebMessageAsString does not.
                // Anything of that second kind is measured ACROSS requests, in the client.
                //
                // A `hostcall` scenario lived here briefly and was worse than nothing. It called
                // RunOnHostThread and timed the return, and RunOnHostThread ENQUEUES and sets a
                // timer rather than waiting - so it reported 0.001ms and read as "reaching the
                // host thread is free". The queued action could not have run anyway, for the
                // reason above. The honest figure for that crossing is perf().marshalMs, which
                // every api request already samples from the far side of it.
                var tripRuns = request.Query.TryGetValue("n", out var tripRunsText)
                    && int.TryParse(tripRunsText, out var tripCount)
                    ? Math.Clamp(tripCount, 1, 50)
                    : 10;

                var tripSamples = new List<double>();
                var tripDetail = string.Empty;
                var tripSurface = _editorSurface;

                if (tripSurface is null)
                {
                    return DebugError("no surface is up");
                }

                switch (tripWhat)
                {
                    case "pagecall":
                    {
                        for (var run = 0; run < tripRuns; run++)
                        {
                            var began = System.Diagnostics.Stopwatch.StartNew();
                            var pinged = RunPageScript("1", null, 5000);
                            began.Stop();
                            if (pinged.Error is { } pageError)
                            {
                                return DebugError(pageError);
                            }
                            tripSamples.Add(Math.Round(began.Elapsed.TotalMilliseconds, 3));
                        }

                        tripDetail = "a script into the page and its answer back";
                        break;
                    }

                    default:
                        return DebugError(
                            $"unknown trip {tripWhat}; pagecall is the only one. Anything that has to "
                            + "observe an effect delivered BY the host thread cannot be measured from in "
                            + "here at all - see the note on this route - and belongs in the client, the "
                            + "way tripCaret() does. For the cost of reaching the host thread, read "
                            + "perf().marshalMs, which every api request already samples from the far side");
                }

                var tripOrdered = tripSamples.OrderBy(one => one).ToArray();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugBenchReply(
                        tripWhat,
                        tripOrdered.Length,
                        tripOrdered[0],
                        tripOrdered[tripOrdered.Length / 2],
                        tripOrdered[Math.Min(tripOrdered.Length - 1, (int)(tripOrdered.Length * 0.95))],
                        tripOrdered[^1],
                        [.. tripSamples],
                        tripDetail),
                    DebugJsonContext.Default.DebugBenchReply));
            }

            case "layout" when request.Query.TryGetValue("reset", out var resetFlag) && resetFlag != "0":
            {
                // Putting the arrangement back. A probe that drags panes about is testing
                // the right thing and leaving the wrong thing behind: the layout is
                // persistent state, and clearing its storage key does not undo what the
                // page already holds in memory. This resets and reloads in one request, so
                // a probe's cleanup is one line that cannot half-work (2026-08-06).
                var reset = $$"""
                    (function () {
                      try { localStorage.removeItem("xlide.docks.v1"); } catch (blocked) { }
                      location.reload();
                      return "reset";
                    })()
                    """;

                _ = RunPageScript(reset, null, 3000);

                var back = Environment.TickCount64;
                var restored = false;
                while (Environment.TickCount64 - back < WaitMilliseconds(request, 20000))
                {
                    Thread.Sleep(150);
                    var probe = RunPageScript("!!(window.xlideBridge && window.xlideBridge.workspace)", null, 1500);
                    if (probe.Error is null && probe.Result.Trim() == "true")
                    {
                        restored = true;
                        break;
                    }
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(restored, 0), DebugJsonContext.Default.DebugCommandReply));
            }

            case "layout":
            {
                // The whole visible arrangement in one answer: which panes are docked where,
                // which editor groups exist and what each shows, and the sizes. Built by
                // hand out of a dozen ad-hoc evals while the layout was being written, which
                // is the argument for it existing (2026-08-06).
                var layout = RunPageScript(LayoutScript, null, 5000);
                return layout.Error is { } layoutError
                    ? DebugError(layoutError)
                    : DebugServer.DebugReply.Json(layout.Result);
            }

            case "reload":
            {
                // Reload the page and WAIT for it to come back, answering with the bundle it
                // is now running. The manual version — reload, sleep a guess, hope — was run
                // a dozen times in one afternoon, and a guess that is too short reports on
                // the page that is going away (2026-08-06).
                var reloadHost = _editorSurface;
                if (reloadHost is null)
                {
                    return DebugError("the surface is not up yet");
                }

                var startedAt = Environment.TickCount64;
                _ = RunPageScript("location.reload()", null, 2000);

                // Ready is the PAGE's own word for it, not a script answering: a page part
                // way through booting can run script and still have no bridge.
                var reloadBudget = WaitMilliseconds(request, 20000);
                var ready = false;
                while (Environment.TickCount64 - startedAt < reloadBudget)
                {
                    Thread.Sleep(150);
                    var probe = RunPageScript("!!(window.xlideBridge && window.xlideBridge.workspace)", null, 1500);
                    if (probe.Error is null && probe.Result.Trim() == "true")
                    {
                        ready = true;
                        break;
                    }
                }

                var stamp = reloadHost.PageBuildStamp ?? "(none reported)";
                var bundle = BundleBuiltUtc();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugReloadReply(
                        ready,
                        (int)(Environment.TickCount64 - startedAt),
                        stamp,
                        bundle,
                        Stale: StampIsBehind(stamp, bundle)),
                    DebugJsonContext.Default.DebugReloadReply));
            }

            case "dialogs":
            {
                // No host thread anywhere in this route, deliberately: it answers while the
                // editor is blocked, which is the only time it matters.
                var rows = DialogWatch.Dialogs()
                    .Select(row => new DebugDialogRow(row.Window, row.Caption, row.Text, row.Buttons, row.Enabled))
                    .ToArray();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugDialogsReply(rows, PerfCounters.HeartbeatAgeMs),
                    DebugJsonContext.Default.DebugDialogsReply));
            }

            case "compile":
            {
                // Does this project compile, and if not, what does it say?
                //
                // Not just the menu command. A compile error is a MODAL, so running it and
                // waiting on the host thread hangs the thread that raised it — which is how a
                // probe left one standing for six minutes, and why the answer nobody could read
                // was on screen the whole time (2026-08-07). The command is started and not
                // waited for; the answering happens here, on the door's own thread, which is the
                // only one still moving while a modal owns the editor.
                if (_editorSurface is not { } compileSurface)
                {
                    return DebugError("the surface is not up yet");
                }

                var standing = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);
                var command = VbeCommands.ForName("compile");
                compileSurface.RunOnHostThread(() => ExecuteEditorCommand(command));

                var said = new List<string>();
                var settle = Environment.TickCount64 + WaitMilliseconds(request, 6000);

                while (Environment.TickCount64 < settle)
                {
                    Thread.Sleep(150);

                    foreach (var raised in DialogWatch.Dialogs())
                    {
                        if (standing.Contains(raised.Window))
                        {
                            continue;
                        }

                        said.Add(raised.Text.Length > 0 ? raised.Text : raised.Caption);
                        standing.Add(raised.Window);

                        // Read, then cleared. A compile error left on screen is the hang this
                        // route exists to stop happening.
                        var compileAnswer = DialogWatch.SafeAnswerFor(raised) ?? "OK";
                        DialogWatch.Dismiss(raised.Caption, compileAnswer);
                        Log.Info($"compile: \"{raised.Text}\" answered with {compileAnswer}");
                    }
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugCompileReply(said.Count == 0, [.. said], DisplayFromProjectId(_shownProject) ?? string.Empty),
                    DebugJsonContext.Default.DebugCompileReply));
            }

            case "type" when request.Body.Length > 0 || request.Query.ContainsKey("text"):
            {
                // Types into the editor the way a person does, so the behaviour that only happens
                // WHILE typing can be tested: smart Enter, comment continuation, auto-indent.
                //
                // Through the editor's own keyboard pipeline — `trigger("keyboard", "type")` —
                // not by setting the text. Setting text goes around every handler that makes
                // typing feel like anything, which means a probe that sets text is testing
                // nothing this product does. \n is sent as a real Enter for the same reason: it
                // is the keystroke the block layout hangs off.
                var typing = request.Body.Length > 0 ? request.Body : request.Query["text"];

                // Typed with GAPS, because typing has gaps.
                //
                // Smart Enter runs from a content-change listener and defers its own work to a
                // microtask, the way the editor's auto-indent lands first. A script that types a
                // newline and then the next line synchronously never lets that run, so the
                // continuation is computed against a line that already has the next line on it —
                // and the first version of this route reported that comment continuation was
                // broken when it was not (2026-08-07). One turn of the loop between segments is
                // the difference between typing and setting text.
                var script = $$"""
                    (async function () {
                      var editor = window.xlideBridge.workspace.activeEditor();
                      editor.focus();
                      var settle = function () {
                        return new Promise(function (done) { setTimeout(done, 24); });
                      };

                      var text = {{JsonString(typing)}};
                      var parts = text.split("\n");

                      for (var i = 0; i < parts.length; i++) {
                        if (i > 0) {
                          editor.trigger("keyboard", "type", { text: "\n" });
                          await settle();
                        }
                        if (parts[i].length > 0) {
                          editor.trigger("keyboard", "type", { text: parts[i] });
                          await settle();
                        }
                      }

                      var at = editor.getPosition();
                      return JSON.stringify({
                        line: at ? at.lineNumber : null,
                        column: at ? at.column : null,
                        text: editor.getModel() ? editor.getModel().getValue() : null
                      });
                    })()
                    """;

                var typed = RunPageScript(script, null, WaitMilliseconds(request, 8000));
                return typed.Error is { } typeError
                    ? DebugError(typeError)
                    : DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                        new DebugEvalReply(typed.Answered, typed.ErrorCode, typed.Result, Unwrap(typed.Result)),
                        DebugJsonContext.Default.DebugEvalReply));
            }

            case "mark" when request.Query.TryGetValue("text", out var marker) && marker.Length > 0:
            {
                // A labelled line in the log, and the offset it landed at.
                //
                // Reading a log for what one step did means finding where that step began, and
                // "scroll up until it looks like the right place" is how a session ends up
                // reasoning about the wrong three seconds. A probe that marks its steps can ask
                // for exactly the slice between two marks — `log({ since })` with the offset this
                // hands back.
                // The offset is taken BEFORE the marker is written, so reading from it returns
                // the marker itself — a slice that starts with the words the caller chose is a
                // slice they can be sure is theirs.
                var at = Log.Path is { } logPath && File.Exists(logPath)
                    ? new FileInfo(logPath).Length
                    : 0;

                Log.Info($"---- {marker} ----");

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugMarkReply(marker, at),
                    DebugJsonContext.Default.DebugMarkReply));
            }

            case "guard":
            {
                // No host thread here either: turning the guard on is exactly what a caller does
                // when the host thread has already stopped answering.
                if (request.Query.TryGetValue("on", out var wanted))
                {
                    _guardEverything = wanted is "1" or "true" or "yes" or "on";
                    Log.Info($"debug api: the dialog guard is {(_guardEverything ? "on" : "off")}");
                }

                string[] cleared;
                lock (_dialogGate)
                {
                    cleared = [.. _guardCleared];
                    if (request.Query.ContainsKey("forget"))
                    {
                        _guardCleared.Clear();
                    }
                }

                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugGuardReply(_guardEverything, cleared, DialogWatch.Dialogs().Length),
                    DebugJsonContext.Default.DebugGuardReply));
            }

            case "dismiss" when request.Query.TryGetValue("button", out var button) && button.Length > 0:
            {
                // Explicit, unlike the automatic guard: the caller names the button, so this
                // one will press OK if asked. The guard's safe-button rule protects requests
                // that never meant to open a dialog at all; a person asking by name knows.
                request.Query.TryGetValue("caption", out var caption);
                var dismissed = DialogWatch.Dismiss(caption, button);
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(dismissed, 0), DebugJsonContext.Default.DebugCommandReply));
            }

            case "stats":
            {
                var placement = PerfCounters.PlacementSnapshot();
                var marshal = PerfCounters.MarshalSnapshot();
                var refresh = PerfCounters.RefreshSnapshot();
                var follow = PerfCounters.FollowSnapshot();
                var messages = WebView.WebView2Surface.MessageTap.Totals;
                using var self = System.Diagnostics.Process.GetCurrentProcess();
                return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugStatsReply(
                        UptimeSeconds: (Environment.TickCount64 - PerfCounters.StartedAt) / 1000,
                        ManagedMemoryBytes: GC.GetTotalMemory(forceFullCollection: false),
                        WorkingSetBytes: Environment.WorkingSet,
                        HandleCount: self.HandleCount,
                        GcCounts: [GC.CollectionCount(0), GC.CollectionCount(1), GC.CollectionCount(2)],
                        PlacementFullPasses: placement.FullPasses,
                        PlacementFastPasses: placement.FastPasses,
                        PlacementFastTotalMs: placement.FastTotalMs,
                        PlacementFastMaxMs: placement.FastMaxMs,
                        WindowEvents: refresh.Events,
                        RefreshPasses: refresh.Passes,
                        RefreshTotalMs: refresh.TotalMs,
                        RefreshMaxMs: refresh.MaxMs,
                        OverlayMs: follow.OverlayMs,
                        BrowserMs: follow.BrowserMs,
                        BrowserCalls: follow.BrowserCalls,
                        PlacementLastMs: placement.LastMs,
                        PlacementMaxMs: placement.MaxMs,
                        MarshalCount: marshal.Count,
                        MarshalLastMs: marshal.LastMs,
                        MarshalMaxMs: marshal.MaxMs,
                        LogLines: PerfCounters.LogLineCount,
                        PollIntervalMs: PerfCounters.PollIntervalMs,
                        MessagesToPage: messages.ToPage,
                        MessagesToHost: messages.ToHost,
                        HeartbeatAgeMs: PerfCounters.HeartbeatAgeMs,
                        DialogsStanding: DialogWatch.Dialogs().Length,
                        ComWrappersTaken: Com.ComRuntime.WrappersTaken,
                        ComWrappersGivenBack: Com.ComRuntime.WrappersGivenBack,
                        ComWrappersLive: Com.ComRuntime.WrappersLive),
                    DebugJsonContext.Default.DebugStatsReply));
            }
        }

        var host = _editorSurface;
        if (host is null)
        {
            return DebugError("the surface is not up yet");
        }

        // The sweep already ran, at the top of this method, so every route heals — including the
        // ones that answer without the host thread and used to return before reaching it. A modal
        // this door raised earlier may still be standing, and waiting for a timeout to notice is
        // the wrong instrument: a VBA modal PUMPS messages, so marshaled work still runs and no
        // timeout ever comes (measured 2026-08-06 — state answered normally while the Macros
        // dialog owned the editor), while the developer is looking at a stuck editor throughout.

        // What was already standing before this request. Anything that appears while it is
        // in flight was raised BY it, and only those may be answered automatically: a dialog
        // the developer opened is theirs, and closing it under them would be worse than any
        // hang. See the timeout path below.
        var standingBefore = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);

        string? answer = null;
        using var done = new ManualResetEventSlim(false);
        var marshalStarted = Environment.TickCount64;
        host.RunOnHostThread(() =>
        {
            try
            {
                answer = AnswerDebugRequestOnHost(request);
            }
            catch (Exception ex)
            {
                answer = System.Text.Json.JsonSerializer.Serialize(
                    new DebugErrorReply($"{ex.GetType().Name}: {ex.Message}"), DebugJsonContext.Default.DebugErrorReply);
            }
            finally
            {
                done.Set();
            }
        });

        var answered = done.Wait(TimeSpan.FromSeconds(3));

        // Every marshaled request doubles as a probe of the host thread's responsiveness;
        // the stats route serves what this line measures.
        PerfCounters.Marshal(Environment.TickCount64 - marshalStarted);

        // Whatever appeared while this request ran, the door raised - and the dangerous case
        // is the one that ANSWERS successfully and leaves a modal standing behind it: Run
        // with the caret outside a procedure returns "ran" and then opens the Macros dialog,
        // which owns the host thread from that moment on. Recording it here is what lets the
        // NEXT request heal instead of timing out forever.
        // Keep and sweep are the same watch with different destinations. Asking synchronously
        // whether this request raised a dialog does not work - it lands after the request
        // returns, which is why the watch is delayed at all - and a synchronous keep check
        // therefore protected nothing: References opened with keep=1 and was swept anyway
        // (2026-08-06).
        RememberRaisedDialogs(standingBefore, keep: request.Query.ContainsKey("keep"));

        if (answered && answer is not null)
        {
            return DebugServer.DebugReply.Json(answer);
        }

        // A request that asked to keep what it opens is not rescued from it: opening a modal
        // was the point, and the caller dismisses it when finished.
        return AnswerBlockedRequest(standingBefore, done, () => answer, request.Query.ContainsKey("keep"));
    }

    /// <summary>
    /// Dialogs this door is answerable for: they were absent when a request began and present
    /// in the moments after it, so that request raised them.
    ///
    /// Attribution took three tries and the failures are the design. A snapshot taken as the
    /// request ends catches nothing, because a dialog arrives microseconds after the command
    /// returns - Run answers "ran" and the Macros dialog comes next. Comparing against
    /// "whatever was standing when the door last looked" then swept a dialog the DEVELOPER
    /// had opened between requests, which is the one outcome worth avoiding entirely
    /// (measured 2026-08-06: an Add Watch opened by hand was cancelled underneath). What
    /// works is watching for a short while AFTER each request, on a pool thread, and owning
    /// only what appears in that window.
    /// </summary>
    private readonly HashSet<string> _dialogsWeRaised = new(StringComparer.Ordinal);
    private readonly Lock _dialogGate = new();

    /// <summary>How long after a request a dialog may appear and still be counted as its doing.</summary>
    private static readonly int[] DialogWatchDelaysMs = [250, 750, 1750];

    /// <summary>
    /// Answers any dialog that appeared while this door was working, with a SAFE button:
    /// Cancel, then Close, then No. Never OK, Yes, Save, Delete, or Run - a dialog nobody
    /// read must not be agreed with, and every safe button means "as you were".
    ///
    /// Two conditions, both required. The dialog must have appeared since the door last
    /// looked, and the host thread must have stopped ticking for three seconds - a poll that
    /// is still running means nothing is wedged and nothing needs rescuing. A dialog the
    /// developer opened while the door was idle is in the snapshot already and is left alone,
    /// however long it stands.
    /// </summary>
    private void ClearDialogsWeRaised()
    {
        string[] ours;
        lock (_dialogGate)
        {
            if (_dialogsWeRaised.Count == 0 && !_guardEverything)
            {
                return;
            }

            ours = [.. _dialogsWeRaised];
        }

        foreach (var dialog in DialogWatch.Dialogs())
        {
            if (_dialogsToKeep.Contains(dialog.Window))
            {
                continue;
            }

            var mine = ours.Contains(dialog.Window);

            // A dialog this door did not raise is cleared only while a caller has asked for the
            // guard, and only when it is a NOTICE. Declining a question nobody asked this door to
            // raise would be answering for the developer; clearing a notice only takes an already
            // finished announcement off the screen — and off the host thread it is holding.
            if (!mine && !(_guardEverything && DialogWatch.IsNotice(dialog)))
            {
                continue;
            }

            var answer = DialogWatch.SafeAnswerFor(dialog);
            var pressed = answer is not null && DialogWatch.Dismiss(dialog.Caption, answer) ? answer : null;

            Log.Info(pressed is null
                ? $"debug api: \"{dialog.Caption}\" has the editor and offers no safe button; leaving it"
                : $"debug api: cleared {(mine ? "our" : "a standing")} dialog \"{dialog.Caption}\""
                    + $"{(dialog.Text.Length > 0 ? $" ({dialog.Text})" : string.Empty)} with {pressed}, "
                    + $"host thread quiet for {PerfCounters.HeartbeatAgeMs}ms");

            if (pressed is not null && !mine)
            {
                _guardCleared.Add($"{dialog.Caption}: {dialog.Text}".Trim().TrimEnd(':'));
            }

            lock (_dialogGate)
            {
                _dialogsWeRaised.Remove(dialog.Window);
            }
        }

        // Anything that closed on its own stops being this door's business.
        var alive = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);
        lock (_dialogGate)
        {
            _dialogsWeRaised.RemoveWhere(window => !alive.Contains(window));
        }
    }

    /// <summary>
    /// Dialogs a caller asked to keep. A request that means to open one - Call Stack is the
    /// standing example - passes keep=1, and what it raises is exempted from the sweep for
    /// as long as it stands. Without this the guard would helpfully cancel the very dialog
    /// the request existed to open.
    /// </summary>
    private readonly HashSet<string> _dialogsToKeep = new(StringComparer.Ordinal);

    /// <summary>
    /// Whether to clear a NOTICE this door did not raise. Off unless a caller asks.
    ///
    /// The rule that a dialog the developer opened is theirs is right for a person at the
    /// keyboard and wrong for a harness: a compile error raised by an experiment stood for six
    /// minutes with the host thread behind it, and nothing in the session could say so because
    /// every other route answers normally while a modal pumps messages (2026-08-07). A harness
    /// turns this on for its run; nothing turns it on by itself.
    /// </summary>
    private volatile bool _guardEverything;

    /// <summary>What the guard has taken off the screen, so a run can report what it swallowed.</summary>
    private readonly List<string> _guardCleared = [];


    /// <summary>
    /// Watches, briefly and on a pool thread, for a dialog this request raised. The delays
    /// are what makes attribution honest: a dialog appears after the command that opened it
    /// returns, and a dialog that appears when no request has just run is the developer's.
    /// </summary>
    private void RememberRaisedDialogs(HashSet<string> standingBefore, bool keep)
    {
        _ = Task.Run(async () =>
        {
            try
            {
                foreach (var delay in DialogWatchDelaysMs)
                {
                    await Task.Delay(delay).ConfigureAwait(false);

                    foreach (var dialog in DialogWatch.Dialogs())
                    {
                        if (standingBefore.Contains(dialog.Window))
                        {
                            continue;
                        }

                        bool noted;
                        lock (_dialogGate)
                        {
                            noted = keep
                                ? _dialogsToKeep.Add(dialog.Window)
                                : _dialogsWeRaised.Add(dialog.Window);
                        }

                        if (noted)
                        {
                            Log.Info(keep
                                ? $"debug api: keeping \"{dialog.Caption}\", as the request asked"
                                : $"debug api: a request raised \"{dialog.Caption}\"; "
                                    + "it will be cleared unless the request asked to keep it");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("debug api: the dialog watch failed", ex);
            }
        });
    }

    /// <summary>
    /// What to say - and do - when the host thread did not answer.
    ///
    /// A bare timeout is the least useful true statement an api can make, and the editor's
    /// commonest reason for one is a MODAL DIALOG: it owns the host thread until somebody
    /// answers it, and every route that needs that thread goes dark for as long as it stands
    /// (twice in one day, a probe left one up and the editor simply stopped). Window
    /// enumeration needs no host thread, so the door can still see what is in the way.
    ///
    /// A dialog that was NOT standing when this request began was raised by this request, and
    /// answering it is undoing our own mess, so it is dismissed and the request retried once.
    /// Only a SAFE button is ever pressed: Cancel, then Close, then No. Never OK, Yes, Save,
    /// Delete, or Run - a dialog nobody read must not be agreed with. A dialog that was
    /// already standing belongs to the developer and is only reported.
    /// </summary>
    /// <summary>
    /// Whether a named claim holds right now, and what was actually seen. Read from the
    /// snapshots the reader thread publishes and from fields the host thread writes, so a
    /// claim can be tested while that thread is busy - which is exactly when a harness is
    /// waiting on one.
    /// </summary>
    private (bool Held, string Saw) EvaluateClaim(string claim, string? expected)
    {
        switch (claim)
        {
            case "stopped":
                return (_inBreak, _inBreak ? "stopped" : "running");

            case "running":
                return (!_inBreak, _inBreak ? "stopped" : "running");

            case "surfaceReady":
                return (_surfaceShown, _surfaceShown ? "ready" : "not ready");

            case "shownModule":
            {
                var shown = _editorSurface?.Module;
                return (shown is not null
                    && (expected is null || shown.Equals(expected, StringComparison.OrdinalIgnoreCase)),
                    shown ?? "(none)");
            }

            case "noDialogs":
            {
                var standing = DialogWatch.Dialogs();
                return (standing.Length == 0, standing.Length == 0 ? "none" : standing[0].Caption);
            }

            case "localsHas":
            {
                var rows = _ghostReaders?.Locals?.Rows;
                var names = rows is null ? [] : rows.Select(row => row.Expression).ToArray();
                return (expected is not null
                    && names.Any(name => name.Equals(expected, StringComparison.OrdinalIgnoreCase)),
                    names.Length == 0 ? "(no locals)" : string.Join(", ", names));
            }

            case "watchHas":
            {
                var rows = _ghostReaders?.Watches;
                var names = rows is null ? [] : rows.Select(row => row.Expression).ToArray();
                return (expected is not null
                    && names.Any(name => name.Equals(expected, StringComparison.OrdinalIgnoreCase)),
                    names.Length == 0 ? "(no watches)" : string.Join(", ", names));
            }

            case "problemFree":
            {
                var held = _findings;
                return (held.Count == 0, held.Count == 0 ? "none" : $"{held.Count} finding(s)");
            }

            case "responsive":
            {
                var age = PerfCounters.HeartbeatAgeMs;
                return (age < 3000, $"{age}ms since the last poll");
            }

            default:
                return (false, $"unknown claim {claim}");
        }
    }

    /// <summary>
    /// An offset far enough back to hold roughly the requested number of lines. A journal
    /// wants the END of the log, and reading a megabyte to reach it would make the capture
    /// itself part of the problem it is capturing.
    /// </summary>
    private static long TailOffset(string path, int lines)
    {
        try
        {
            var length = new FileInfo(path).Length;

            // Lines here run long: timestamps, a level, a thread, and often a serialized
            // message. Two hundred characters apiece is a generous guess that errs towards
            // reading more than asked, which the line cap then trims.
            var guess = (long)lines * 200;
            return length > guess ? length - guess : 0;
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>One read of the log from an offset, filtered, with the offset to ask from next.</summary>
    private static (List<string> Lines, long Next) ReadLogSlice(string path, long since, string? match, int max)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (since > 0 && since <= file.Length)
        {
            file.Position = since;
        }

        using var reader = new StreamReader(file, System.Text.Encoding.UTF8);
        var lines = new List<string>();
        while (reader.ReadLine() is { } line && lines.Count < max)
        {
            if (match is null || line.Contains(match, StringComparison.OrdinalIgnoreCase))
            {
                lines.Add(line);
            }
        }

        return (lines, file.Length);
    }

    /// <summary>
    /// A page script's answer, unwrapped as far as it is wrapped.
    ///
    /// The browser returns a result as JSON, so a script returning a string returns a QUOTED
    /// string; a script that builds its answer with JSON.stringify — which every useful one does,
    /// because that is how a structure crosses — returns it quoted twice. Unwrapping stops at the
    /// first thing that is not itself a JSON document, so a plain string stays a plain string.
    /// </summary>
    private static System.Text.Json.Nodes.JsonNode? Unwrap(string result)
    {
        System.Text.Json.Nodes.JsonNode? node;
        try
        {
            node = System.Text.Json.Nodes.JsonNode.Parse(result);
        }
        catch (System.Text.Json.JsonException)
        {
            return null;
        }

        for (var depth = 0; depth < 4; depth++)
        {
            if (node is not System.Text.Json.Nodes.JsonValue value
                || !value.TryGetValue<string>(out var inner))
            {
                break;
            }

            try
            {
                node = System.Text.Json.Nodes.JsonNode.Parse(inner);
            }
            catch (System.Text.Json.JsonException)
            {
                break;
            }
        }

        return node;
    }

    private static DebugServer.DebugReply AnswerBlockedRequest(
        HashSet<string> standingBefore,
        ManualResetEventSlim done,
        Func<string?> answerSoFar,
        bool keep)
    {
        var blocking = keep
            ? null
            : DialogWatch.Dialogs().FirstOrDefault(row => !standingBefore.Contains(row.Window));

        if (blocking is null)
        {
            var standing = DialogWatch.Dialogs();
            return DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                new DebugBlockedReply(
                    Error: "the host thread did not answer in time",
                    HeartbeatAgeMs: PerfCounters.HeartbeatAgeMs,
                    BlockedBy: standing.Length > 0 ? standing[0].Caption : null,
                    Buttons: standing.Length > 0 ? standing[0].Buttons : [],
                    Dismissed: null,
                    Retried: false),
                DebugJsonContext.Default.DebugBlockedReply));
        }

        var safe = DialogWatch.SafeAnswerFor(blocking);
        var pressed = safe is not null && DialogWatch.Dismiss(blocking.Caption, safe) ? safe : null;

        Log.Info(pressed is null
            ? $"debug api: \"{blocking.Caption}\" is blocking the host thread and has no safe button"
            : $"debug api: \"{blocking.Caption}\" was raised by this request; answered with {pressed}");

        // The dismissal releases the host thread, and the work this request asked for was
        // queued before the dialog appeared, so it may complete on its own.
        var completed = pressed is not null && done.Wait(TimeSpan.FromSeconds(3));

        return completed && answerSoFar() is { } answer
            ? DebugServer.DebugReply.Json(answer)
            : DebugServer.DebugReply.Json(System.Text.Json.JsonSerializer.Serialize(
                new DebugBlockedReply(
                    Error: pressed is null
                        ? "a dialog this request raised is blocking the host thread, and it has no safe button to press"
                        : "a dialog this request raised was dismissed, but the request did not finish",
                    HeartbeatAgeMs: PerfCounters.HeartbeatAgeMs,
                    BlockedBy: blocking.Caption,
                    Buttons: blocking.Buttons,
                    Dismissed: pressed,
                    Retried: false),
                DebugJsonContext.Default.DebugBlockedReply));
    }

    private unsafe string AnswerDebugRequestOnHost(DebugServer.DebugRequest request)
    {
        switch (request.Route)
        {
            case "state":
            {
                Rect frameRect = default;
                Rect documentsRect = default;
                if (_frame != 0)
                {
                    Win32.GetWindowRect(_frame, &frameRect);
                }

                if (_documentArea != 0)
                {
                    Win32.GetWindowRect(_documentArea, &documentsRect);
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugStateReply(
                        Configuration: "debug",
                        ShownModule: _editorSurface?.Module,
                        ShownProject: DisplayFromProjectId(_shownProject),
                        DebugMode: _lastPublishedMode,
                        HasUnwrittenEdits: _editorSurface?.HasUnwrittenEdits ?? false,
                        // Whether the engine is ANSWERING, not whether this session got as far as
                        // constructing the service that talks to it. The old reading was true from
                        // start-up to shutdown whatever the engine did, so killing the engine
                        // process left the door reporting engineUp: true while the editor drew
                        // squiggles from the last pass that ran (found 2026-08-08 by killing it).
                        EngineUp: _analysis?.IsReady == true,
                        Frame: $"0x{_frame:X}",
                        FrameRect: $"{frameRect.Left},{frameRect.Top},{frameRect.Right},{frameRect.Bottom}",
                        DocumentArea: $"0x{_documentArea:X}",
                        DocumentAreaRect: $"{documentsRect.Left},{documentsRect.Top},{documentsRect.Right},{documentsRect.Bottom}",
                        PaletteOpen: _browserPalette is not null,
                        PaletteVisible: _browserPalette is { } palette && Win32.IsWindowVisible(palette.Handle),
                        SurfaceReady: _surfaceShown,
                        DevToolsPort: WebView.WebView2Surface.DevToolsPort),
                    DebugJsonContext.Default.DebugStateReply);
            }

            case "doctor":
            {
                // The questions that are asked at the START of every confusing session, and
                // that cost the most when nobody thinks to ask them. Chief among them: is
                // the code running the code I just built? A shim and a page built minutes
                // apart, or a session serving a bundle from somewhere else entirely, produce
                // symptoms that look like anything except what they are - three rounds of
                // "why is my fix not in the log" on 2026-08-06 were exactly this.
                var shimPath = Interop.ShimModule.Directory;
                var shimFile = shimPath is null ? null : Path.Combine(shimPath, "Xlide.Vbe.Shim.dll");
                var bundle = shimPath is null
                    ? null
                    : Path.Combine(shimPath, "ui", "editor", "dist", "editor.js");

                var findings = new List<string>();

                if (shimFile is not null && File.Exists(shimFile) && bundle is not null && File.Exists(bundle))
                {
                    var shimBuilt = File.GetLastWriteTimeUtc(shimFile);
                    var bundleBuilt = File.GetLastWriteTimeUtc(bundle);
                    var apart = (shimBuilt - bundleBuilt).Duration();
                    if (apart > TimeSpan.FromMinutes(30))
                    {
                        findings.Add($"the shim and the page bundle were built {apart.TotalMinutes:N0} "
                            + "minutes apart; one of them is probably stale");
                    }
                }
                else
                {
                    findings.Add("the shim directory does not hold both a shim and a page bundle");
                }

                if (_editorSurface?.PageBuildStamp is null)
                {
                    findings.Add("the page never reported a build stamp; it may not have finished booting");
                }

                // A standing dialog owns the host thread, and every OTHER route answers normally
                // while it does — so a session can look healthy for minutes while nothing it is
                // asked to do can run. It was found by a person looking at the screen, which is
                // the one instrument a harness does not have (2026-08-07).
                foreach (var standing in DialogWatch.Dialogs())
                {
                    var says = standing.Text.Length > 0 ? $": {standing.Text}" : string.Empty;
                    findings.Add($"a dialog is standing and owns the host thread{says} "
                        + $"(buttons: {string.Join(", ", standing.Buttons)})");
                }

                if (_analysis is null)
                {
                    findings.Add("the analysis engine is not up, so diagnostics will stay empty");
                }
                else if (!_analysis.IsReady)
                {
                    // Distinct from the case above, and the more dangerous one: the service is
                    // there, the last pass's findings are still drawn, and nothing new will ever
                    // be analysed because the engine is started once and never restarted. The
                    // doctor called that healthy until 2026-08-08.
                    findings.Add("the analysis engine is not answering, so the findings on screen "
                        + "are from the last pass that ran and will not change (the engine is not restarted)");
                }

                if (_ghostReaders is null)
                {
                    findings.Add("the ghost readers are not attached, so Locals and Watch cannot fill");
                }

                // Only diagnostic while something should be ticking. An idle editor stops
                // polling by design, and a doctor that called that a fault would cry wolf on
                // every quiet session - which it did, the first time it ran (2026-08-06).
                if (PerfCounters.PollingExpected && PerfCounters.HeartbeatAgeMs > 5000)
                {
                    findings.Add($"the host thread has not ticked for {PerfCounters.HeartbeatAgeMs}ms "
                        + "while it should be polling; something is holding it (check the dialogs route)");
                }

                if (DialogWatch.Dialogs() is { Length: > 0 } standingDialogs)
                {
                    findings.Add($"a dialog is standing: \"{standingDialogs[0].Caption}\"");
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugDoctorReply(
                        Healthy: findings.Count == 0,
                        ShimPath: shimFile ?? "(unknown)",
                        ShimBuiltUtc: shimFile is not null && File.Exists(shimFile)
                            ? File.GetLastWriteTimeUtc(shimFile).ToString("O")
                            : "(missing)",
                        BundleBuiltUtc: bundle is not null && File.Exists(bundle)
                            ? File.GetLastWriteTimeUtc(bundle).ToString("O")
                            : "(missing)",
                        PageBuildStamp: _editorSurface?.PageBuildStamp ?? "(none reported)",
                        // Whether the engine is ANSWERING, not whether this session got as far as
                        // constructing the service that talks to it. The old reading was true from
                        // start-up to shutdown whatever the engine did, so killing the engine
                        // process left the door reporting engineUp: true while the editor drew
                        // squiggles from the last pass that ran (found 2026-08-08 by killing it).
                        EngineUp: _analysis?.IsReady == true,
                        GhostReadersUp: _ghostReaders is not null,
                        SurfaceReady: _surfaceShown,
                        Findings: [.. findings]),
                    DebugJsonContext.Default.DebugDoctorReply);
            }

            case "engine" when request.Query.TryGetValue("module", out var engineModule) && engineModule.Length > 0:
            {
                // WHAT THE ENGINE IS HOLDING, which nothing could see until 2026-08-08.
                //
                // Every finding is computed against this copy and it is maintained incrementally,
                // so a squiggle drawn in the wrong place is always the same question: does this
                // match the surface? A finding was seen six columns out after a format and there
                // was no way to ask which side had drifted.
                var wantEngineText = request.Query.TryGetValue("text", out var engineText) && engineText != "0";
                var surface = _editorSurface?.TextOf(engineModule, DisplayFromProjectId(_shownProject));

                try
                {
                    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                    var held = _analysis?.LiveSourceAsync(engineModule, wantEngineText, deadline.Token)
                        .GetAwaiter().GetResult();

                    var engineHeld = held?.TryGetProperty("held", out var heldValue) == true && heldValue.GetBoolean();
                    var engineLines = held?.TryGetProperty("lines", out var linesValue) == true ? linesValue.GetInt32() : 0;
                    var engineSource = held?.TryGetProperty("source", out var sourceValue) == true
                        && sourceValue.ValueKind == System.Text.Json.JsonValueKind.String
                        ? sourceValue.GetString() : null;

                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugEngineSourceReply(
                            engineModule,
                            engineHeld,
                            engineLines,
                            surface?.Split('\n').Length ?? 0,
                            ContentKey(engineSource),
                            ContentKey(surface),
                            wantEngineText ? engineSource : null,
                            wantEngineText ? surface : null),
                        DebugJsonContext.Default.DebugEngineSourceReply);
                }
                catch (Exception ex)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply($"the engine's copy could not be read ({ex.GetType().Name})"),
                        DebugJsonContext.Default.DebugErrorReply);
                }
            }

            case "native":
            {
                // THE HOST'S OWN EDITOR, underneath the surface that covers it.
                //
                // Everything else this door reports is the page, or the workbook. Neither is
                // what Run, Step, Compile and ToggleBreakpoint act on: those act on the native
                // ACTIVE CODE PANE and the caret inside it. A page showing one module while the
                // editor's active pane is another is a Run that executes elsewhere and a
                // breakpoint set on the wrong line, with nothing on screen to say so.
                //
                // Asked for by the developer 2026-08-08 -- "are you validating the vbe native
                // editor surface is staying in sync" -- and the honest answer was no: every
                // check until then read the page and the workbook and never the panes below.
                using var activePane = _editor.GetObject("ActiveCodePane");

                string? activeModule = null;
                string? activeProject = null;
                var caretLine = 0;
                var caretColumn = 0;
                var nativeLines = 0;
                string? nativeText = null;

                if (activePane is not null)
                {
                    Span<int> selection = stackalloc int[4];
                    try
                    {
                        activePane.InvokeInt32s("GetSelection", selection);
                        caretLine = selection[0];
                        caretColumn = selection[1];
                    }
                    catch (Exception)
                    {
                        // A pane mid-teardown answers nothing; the rest of the picture stands.
                    }

                    using var codeModule = activePane.GetObject("CodeModule");
                    using var component = codeModule?.GetObject("Parent");
                    activeModule = component?.GetString("Name");

                    // THE PANE'S TEXT, not a proxy for it. Names agreeing is not parity and
                    // neither is a line count: what a developer means by "the editor is in sync"
                    // is that the code is the same code. A surface holding an empty document for
                    // a module the host has 42 lines of passed every name comparison there was,
                    // and showed a blank editor (2026-08-08).
                    //
                    // Hashed rather than shipped, because this is asked after every step of a
                    // randomised walk; `text=1` carries the actual text for the run that fails.
                    nativeLines = codeModule?.GetInt32("CountOfLines") ?? 0;
                    nativeText = component is null ? null : ProjectReader.ReadSource(component);

                    using var collection = component?.GetObject("Collection");
                    using var owner = collection?.GetObject("Parent");
                    activeProject = owner is null ? null : DisplayFromProjectId(ProjectReader.Identity(owner).Id);
                }

                var surfaceText = _editorSurface?.Text;
                var wantText = request.Query.TryGetValue("text", out var wantsText) && wantsText != "0";

                // EVERY open pane, each with the host's content and the surface's side by side.
                //
                // The active one is not the only one that can drift. A background tab holds a
                // copy the developer is not looking at, so a module written from outside while
                // its tab sits behind another goes stale with nothing to notice until it is
                // clicked — and then it is the developer who notices.
                var paneRows = (ReadOpenModules() ?? [])
                    .Select(pane =>
                    {
                        string? hostText = null;
                        try
                        {
                            using var found = FindComponent(pane.Name, ProjectIdFromDisplay(pane.Project), out _);
                            hostText = found is null ? null : ProjectReader.ReadSource(found);
                        }
                        catch (Exception)
                        {
                            // A component mid-teardown answers nothing; it is reported as unknown
                            // rather than as a disagreement.
                        }

                        return new DebugNativePaneRow(
                            pane.Name,
                            pane.Project,
                            ContentKey(hostText),
                            ContentKey(_editorSurface?.TextOf(pane.Name, pane.Project)));
                    })
                    .ToArray();

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugNativeReply(
                        activeModule,
                        activeProject,
                        caretLine,
                        caretColumn,
                        paneRows,
                        _editorSurface?.Module,
                        DisplayFromProjectId(_shownProject),
                        nativeLines,
                        surfaceText?.Split('\n').Length ?? 0,
                        ContentKey(nativeText),
                        ContentKey(surfaceText),
                        wantText ? nativeText : null,
                        wantText ? surfaceText : null),
                    DebugJsonContext.Default.DebugNativeReply);
            }

            case "windows":
            {
                var rows = new List<DebugWindowRow>();
                using var windows = _editor.GetObject("Windows");
                var count = windows?.GetInt32("Count") ?? 0;
                for (var i = 1; i <= count; i++)
                {
                    using var window = windows!.GetItem(i);
                    if (window is not null)
                    {
                        rows.Add(new DebugWindowRow(
                            window.GetInt32("Type"),
                            window.GetString("Caption") ?? string.Empty,
                            window.GetBool("Visible")));
                    }
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugWindowsReply([.. rows]), DebugJsonContext.Default.DebugWindowsReply);
            }

            case "outline" when request.Query.TryGetValue("module", out var outlineModule) && outlineModule.Length > 0:
            {
                // A module's shape, from the analyzer, so a caller can assert on structure rather
                // than read the text back and parse it a second time — in a second language, with
                // a second set of bugs.
                if (_analysis is not { } outlineAnalysis)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply("the analysis engine is not up"),
                        DebugJsonContext.Default.DebugErrorReply);
                }

                request.Query.TryGetValue("project", out var outlineProject);

                try
                {
                    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                    var answered = outlineAnalysis
                        .OutlineAsync(outlineModule, ProjectIdFromDisplay(outlineProject), source: null, deadline.Token)
                        .GetAwaiter().GetResult();

                    if (answered is null)
                    {
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugErrorReply($"'{outlineModule}' could not be outlined"),
                            DebugJsonContext.Default.DebugErrorReply);
                    }

                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugOutlineReply(
                            outlineModule,
                            [.. answered.Select(p => new DebugProcedureRow(p.Name, p.Kind, p.Line))]),
                        DebugJsonContext.Default.DebugOutlineReply);
                }
                catch (Exception ex)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply($"outline failed: {ex.Message.Trim()}"),
                        DebugJsonContext.Default.DebugErrorReply);
                }
            }

            case "component" when request.Query.TryGetValue("action", out var componentAction):
            {
                // Adding, renaming and removing components, from INSIDE.
                //
                // This is what a fixture is made of, and until now it was the one thing a harness
                // had to reach in through `Workbook.VBProject` for — which needs "Trust access to
                // the VBA project object model" turned on. The add-in is already past that gate:
                // the host hands it the VBE at OnConnection. So the fixture can be built through
                // the door, and the setting can stay off (2026-08-07).
                request.Query.TryGetValue("name", out var componentName);
                request.Query.TryGetValue("project", out var componentProject);
                var componentOwner = ProjectIdFromDisplay(componentProject) ?? _shownProject;

                try
                {
                    switch (componentAction)
                    {
                        case "add":
                        {
                            // 1 standard, 2 class, 3 form — the VBE's own numbering.
                            var kind = request.Query.TryGetValue("kind", out var kindText)
                                && int.TryParse(kindText, out var parsedKind) ? parsedKind : 1;

                            if (kind is not (1 or 2 or 3))
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply($"kind {kind} is not 1, 2 or 3"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            using var project = FindProjectByDisplayName(componentProject)
                                ?? _editor.GetObject("ActiveVBProject");
                            using var components = project?.GetObject("VBComponents");
                            using var added = components?.CallObject("Add", kind);
                            if (added is null)
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply("the project would not add a component"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            // Named here rather than left as Module1, because a fixture is its
                            // names. The editor refuses some outright — Circle is owned by the
                            // Excel object library — and says so with a bare HRESULT, so the
                            // refusal is reported with the name that caused it.
                            if (componentName is { Length: > 0 })
                            {
                                try
                                {
                                    added.SetString("Name", componentName);
                                }
                                catch (Exception ex)
                                {
                                    // Taken back out. A refused name otherwise leaves a Module1
                                    // nobody asked for, in a project a fixture is about to make
                                    // claims about — and the next run finds it and is confused by
                                    // it. Add either produces the component that was asked for or
                                    // produces nothing.
                                    try { components?.InvokeWithObject("Remove", added); }
                                    catch (Exception undo) { Log.Warn($"component: could not undo the add ({undo.GetType().Name})"); }

                                    return System.Text.Json.JsonSerializer.Serialize(
                                        new DebugErrorReply(
                                            $"'{componentName}' was refused as a name, so nothing was added ({ex.Message.Trim()})"),
                                        DebugJsonContext.Default.DebugErrorReply);
                                }
                            }

                            var finalName = added.GetString("Name") ?? string.Empty;
                            Log.Info($"component: added {finalName} (kind {kind})");

                            // The strip AND the tree. Neither republishes on its own, and they are
                            // separate publishes: the first version of this route refreshed the
                            // tabs only, so the explorer went on listing three components while
                            // the strip showed eight — a surface describing two different
                            // projects at once (the developer, 2026-08-07).
                            ComponentsChanged();

                            return System.Text.Json.JsonSerializer.Serialize(
                                new DebugComponentReply(true, finalName, "add"),
                                DebugJsonContext.Default.DebugComponentReply);
                        }

                        case "remove":
                        {
                            if (componentName is not { Length: > 0 })
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply("remove needs a name"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            using var doomed = FindComponent(componentName, componentOwner, out var removedFrom);
                            if (doomed is null)
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply($"'{componentName}' is not a component of this project"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            using var owningProject = FindProjectByDisplayName(
                                DisplayFromProjectId(removedFrom ?? componentOwner))
                                ?? _editor.GetObject("ActiveVBProject");
                            using var holding = owningProject?.GetObject("VBComponents");
                            if (holding is null)
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply("the project would not open its component list"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            // Remove takes the COMPONENT, not an index, so it goes through the
                            // object-argument path rather than the integer one.
                            holding.InvokeWithObject("Remove", doomed);

                            Log.Info($"component: removed {componentName}");
                            ComponentsChanged();

                            return System.Text.Json.JsonSerializer.Serialize(
                                new DebugComponentReply(true, componentName, "remove"),
                                DebugJsonContext.Default.DebugComponentReply);
                        }

                        case "rename":
                        {
                            // The COMPONENT only. Renaming a module AND everything that names it
                            // is `renameModule` through the page, which is a different operation
                            // with an engine behind it; this is the fixture-building primitive.
                            if (componentName is not { Length: > 0 }
                                || !request.Query.TryGetValue("newName", out var newName)
                                || newName.Length == 0)
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply("rename needs name and newName"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            using var target = FindComponent(componentName, componentOwner, out _);
                            if (target is null)
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply($"'{componentName}' is not a component of this project"),
                                    DebugJsonContext.Default.DebugErrorReply);
                            }

                            target.SetString("Name", newName);
                            var readBack = target.GetString("Name") ?? newName;
                            Log.Info($"component: renamed {componentName} to {readBack}");
                            ComponentsChanged();

                            return System.Text.Json.JsonSerializer.Serialize(
                                new DebugComponentReply(true, readBack, "rename"),
                                DebugJsonContext.Default.DebugComponentReply);
                        }

                        default:
                            return System.Text.Json.JsonSerializer.Serialize(
                                new DebugErrorReply($"unknown action {componentAction}; use add, remove or rename"),
                                DebugJsonContext.Default.DebugErrorReply);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"component: {componentAction} failed", ex);
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply($"{componentAction} failed: {ex.Message.Trim()}"),
                        DebugJsonContext.Default.DebugErrorReply);
                }
            }

            case "pane" when request.Query.TryGetValue("action", out var paneAction)
                && request.Query.TryGetValue("module", out var paneModule) && paneModule.Length > 0:
            {
                // Opening and CLOSING a module's pane.
                //
                // `caret` opens one on the way to a line, and until now nothing closed one — so
                // every test of what the tab strip does when a tab goes had to reach into the
                // page's private workspace through eval, which is a test of the probe as much as
                // of the thing. Four defects in the strip this week were found that way and each
                // one needed the reach rewritten (2026-08-07).
                request.Query.TryGetValue("project", out var paneProject);
                var paneOwner = ProjectIdFromDisplay(paneProject) ?? _shownProject;

                switch (paneAction)
                {
                    case "open":
                        // The workbook is PASSED ON. It was computed and then dropped, so
                        // `project=` did nothing at all on open and a bare name resolved
                        // shown-project-first — meaning the second workbook's copy of a shared
                        // module name could not be opened from a script by any argument.
                        //
                        // That is why the two-workbook state was unreachable from the harness, and
                        // why every defect in this class had to be found by hand. A stress walk
                        // seeded with both workbooks' Helpers silently held only one of them and
                        // passed its label checks vacuously (2026-08-07).
                        ShowModule(paneModule, DisplayFromProjectId(paneOwner));
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply);

                    case "close":
                    {
                        // Through the same gate the tab's own X uses, so a module with unwritten
                        // edits gets the question rather than the guillotine — and `action` is how
                        // a caller answers it in advance.
                        request.Query.TryGetValue("answer", out var closeAnswer);
                        OnModuleCloseRequested(paneModule, DisplayFromProjectId(paneOwner), closeAnswer);
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply);
                    }

                    default:
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugErrorReply($"unknown action {paneAction}; use open or close"),
                            DebugJsonContext.Default.DebugErrorReply);
                }
            }

            case "undoRename":
            {
                // The same path the editor's own Undo Rename takes. Here so a probe can prove a
                // rename is reversible without driving the page, which is the half a rename test
                // could never assert before.
                UndoRename(0);
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply);
            }

            case "breakpoints":
            {
                // Reading what is set. There has been a way to SET a breakpoint since the door
                // landed and no way to ask what is set, which makes every debugger assertion a
                // matter of remembering what the test did rather than looking (2026-08-07).
                // Reported from the record's own spellings, not from the key. The key is
                // lowercased so that two workbooks holding Helpers and helpers are holding the
                // same module, and a first version handed that key back: the route answered
                // `helpers @ renamefixture.xlsm`, so a caller comparing against the name on
                // screen matched nothing. A door that mangles its own answers is worse than one
                // that refuses (2026-08-08).
                var rows = _breakpoints.Values
                    .Where(record => record.Lines.Count > 0)
                    .OrderBy(record => record.Project, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(record => record.Module, StringComparer.OrdinalIgnoreCase)
                    .Select(record => new DebugBreakpointRow(record.Module, record.Project, [.. record.Lines]))
                    .ToArray();

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugBreakpointsReply(rows, _lastPublishedMode),
                    DebugJsonContext.Default.DebugBreakpointsReply);
            }

            case "settings":
            {
                // Read them, or change one without restating the rest.
                //
                // The page's own update takes the WHOLE settings object, so changing one thing
                // from a harness meant spelling out all seven and getting a default wrong in the
                // process. Here a caller names what it wants changed and everything else stands.
                var settings = _settings;

                if (request.Query.Count > 0)
                {
                    bool Flag(string name, bool current) =>
                        request.Query.TryGetValue(name, out var asked)
                            ? asked is "1" or "true" or "yes" or "on"
                            : current;

                    settings = new ProductSettings
                    {
                        BlockLayout = request.Query.TryGetValue("blockLayout", out var layout)
                            ? layout
                            : settings.BlockLayout,
                        ContinueCommentOnNewline = Flag("continueCommentOnNewline", settings.ContinueCommentOnNewline),
                        MirrorCommentSpacing = Flag("mirrorCommentSpacing", settings.MirrorCommentSpacing),
                        TreeFollowsEditor = Flag("treeFollowsEditor", settings.TreeFollowsEditor),
                        FormatIndentSize = request.Query.TryGetValue("formatIndentSize", out var indent)
                            && int.TryParse(indent, out var asked) ? asked : settings.FormatIndentSize,
                        FormatCanonicalKeywords = Flag("formatCanonicalKeywords", settings.FormatCanonicalKeywords),
                    }.Normalized();

                    OnSettingsChanged(settings);
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugSettingsReply(
                        settings.BlockLayout,
                        settings.ContinueCommentOnNewline,
                        settings.MirrorCommentSpacing,
                        settings.TreeFollowsEditor,
                        settings.FormatIndentSize,
                        settings.FormatCanonicalKeywords),
                    DebugJsonContext.Default.DebugSettingsReply);
            }

            /*
             * EVERY open workbook, which nothing could ask for.
             *
             * `project` answers about ONE: the one named, or the active one. With two workbooks
             * open there was no way to discover the other's name from the host at all, so a probe
             * either knew it in advance or asked the page's tree, which is the surface's view
             * rather than the object model's. The language suite failed exactly there: it asked
             * `project()`, got whichever workbook happened to be active, and looked for its own
             * fixture's module inside the other one.
             *
             * That matters more here than in most products. Two workbooks holding a module of the
             * same name is a designed case, and three separate defects have lived in it. A suite
             * that cannot name the workbook it means cannot test any of them.
             *
             * The plural of a noun, beside its singular, the way `breakpoints` sits beside
             * `breakpoint`. Cheap on purpose: names and counts, not contents. Ask `project` for
             * what is inside one.
             */
            case "projects":
            {
                var found = new List<DebugProjectRow>();

                using (var projects = _editor.GetObject("VBProjects"))
                {
                    var count = projects?.GetInt32("Count") ?? 0;
                    for (var i = 1; i <= count; i++)
                    {
                        try
                        {
                            using var project = projects!.GetItem(i);
                            if (project is null)
                            {
                                continue;
                            }

                            var identity = ProjectReader.Identity(project);
                            var display = WorkbookDisplayName(project);

                            // Components counted rather than listed, and the scratch module left
                            // out of the count for the same reason it is left out everywhere else:
                            // it is ours, and a fixture that counts it counts wrong.
                            var components = 0;
                            using (var list = project.GetObject("VBComponents"))
                            {
                                var total = list?.GetInt32("Count") ?? 0;
                                for (var c = 1; c <= total; c++)
                                {
                                    using var component = list!.GetItem(c);
                                    if (component?.GetString("Name") is { Length: > 0 } name
                                        && !IsScratchComponent(name))
                                    {
                                        components++;
                                    }
                                }
                            }

                            found.Add(new DebugProjectRow(
                                display ?? identity.Id,
                                identity.Id,
                                components,
                                string.Equals(DisplayFromProjectId(_shownProject), display, StringComparison.OrdinalIgnoreCase)));
                        }
                        catch (Exception ex)
                        {
                            Log.Info($"projects: entry {i} could not be read ({ex.GetType().Name})");
                        }
                    }
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugProjectsReply([.. found]), DebugJsonContext.Default.DebugProjectsReply);
            }

            case "project":
            {
                // What is actually THERE, as opposed to what the surface is showing.
                //
                // This is the question a fixture asks twice — once to build and once to check —
                // and it was the last one that could only be answered by reaching in through
                // `Workbook.VBProject`, which needs the trust setting. Answered from inside, where
                // the add-in already is.
                request.Query.TryGetValue("project", out var wantedProject);

                using var project = FindProjectByDisplayName(wantedProject)
                    ?? _editor.GetObject("ActiveVBProject");

                if (project is null)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply("no VBA project is active"),
                        DebugJsonContext.Default.DebugErrorReply);
                }

                var rows = new List<DebugComponentRow>();
                using (var components = project.GetObject("VBComponents"))
                {
                    var count = components?.GetInt32("Count") ?? 0;
                    for (var i = 1; i <= count; i++)
                    {
                        try
                        {
                            using var component = components!.GetItem(i);
                            if (component?.GetString("Name") is not { Length: > 0 } name
                                || IsScratchComponent(name))
                            {
                                continue;
                            }

                            var type = component.GetInt32("Type");
                            using var code = component.GetObject("CodeModule");

                            // A pane exists once the module has been LOOKED at. Reading CodePane
                            // would create one, which would make asking the question change the
                            // answer, so this asks the open list instead.
                            var open = ReadOpenModules()?.Any(pane =>
                                string.Equals(pane.Name, name, StringComparison.OrdinalIgnoreCase)) ?? false;

                            rows.Add(new DebugComponentRow(
                                name,
                                ComponentKind(type),
                                type,
                                code?.GetInt32("CountOfLines") ?? 0,
                                open));
                        }
                        catch (Exception ex)
                        {
                            Log.Verbose($"project: component {i} could not be read ({ex.GetType().Name})");
                        }
                    }
                }

                // The identity of the project THIS REPLY DESCRIBES, read off that project.
                //
                // It used to be DisplayFromProjectId(_shownProject) — the workbook the surface
                // happened to be showing — so asking about the second workbook answered with the
                // second workbook's components under the FIRST workbook's id. The reply
                // contradicted itself, and a caller doing the obvious thing (read `projectId`,
                // pass it to `pane` or `module`) was then addressing the wrong workbook while
                // holding a reply that looked right.
                //
                // That is why the two-workbook state could never be set up from a script, which
                // is why every defect in this class — navigation, tab labels, breakpoints — had
                // to be found by hand (2026-08-07).
                // Through DisplayFromProjectId, so the shape is unchanged: the field carries the
                // workbook FILE NAME, which is the form every route's `project=` argument takes.
                // The raw identity is a full path, and handing that back would fix the wrong
                // project only to make the value unusable.
                var identity = ProjectReader.Identity(project);

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugProjectReply(
                        project.GetString("Name") ?? string.Empty,
                        DisplayFromProjectId(identity.Id),
                        project.GetInt32("Mode"),
                        [.. rows]),
                    DebugJsonContext.Default.DebugProjectReply);
            }

            case "documents":
            {
                // What the surface actually HOLDS, as opposed to what the strip draws. A module
                // with a tab and no text is the state most of a workspace is in, and it is what
                // an empty peek window and a blank pane both turned out to be.
                var rows = _editorSurface?.DocumentTable
                    .Select(row => new DebugDocumentRow(row.Module, row.Project, row.Lines, row.Unwritten, row.Active))
                    .ToArray() ?? [];

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugDocumentsReply(rows, _editorSurface?.Module),
                    DebugJsonContext.Default.DebugDocumentsReply);
            }

            case "command" when request.Query.TryGetValue("name", out var name) && name.Length > 0:
            {
                var command = VbeCommands.ForName(name);
                if (command == 0)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply($"unknown command name {name}"), DebugJsonContext.Default.DebugErrorReply);
                }

                ExecuteEditorCommand(command);
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(true, command), DebugJsonContext.Default.DebugCommandReply);
            }

            case "breakpoint"
                when request.Query.TryGetValue("module", out var module) && module.Length > 0
                    && request.Query.TryGetValue("line", out var lineText)
                    && int.TryParse(lineText, out var breakLine) && breakLine >= 1:
            {
                // The same manner a person uses: go to the line, toggle there. This is what
                // makes break mode a harness-reachable state, the regression net the
                // debugger milestone needs before it starts.
                //
                // state=on|off makes it IDEMPOTENT, which is what a script wants: the bare
                // toggle cost a live run its breakpoint when a retry cleared what the first
                // call had set (2026-08-06). Without the argument it still toggles, the way
                // the key does.
                request.Query.TryGetValue("project", out var project);
                request.Query.TryGetValue("state", out var wanted);
                GoTo(module, breakLine, 1, project);

                // Read against the module the GoTo above has just made the shown one, in the
                // workbook it belongs to. Keyed by name alone this read the TWIN's record when
                // two workbooks shared the module name, so `state=on` saw a breakpoint that was
                // not there and did nothing.
                var alreadySet = _editorSurface?.Module is { } shownModule
                    && BreakpointsFor(shownModule).Contains(breakLine);
                var shouldSet = wanted switch
                {
                    "on" => true,
                    "off" => false,
                    _ => !alreadySet,
                };

                if (shouldSet != alreadySet)
                {
                    ToggleBreakpoint(breakLine);
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(shouldSet, VbeCommands.Command.ToggleBreakpoint), DebugJsonContext.Default.DebugCommandReply);
            }

            case "module" when request.Query.TryGetValue("name", out var moduleName) && moduleName.Length > 0:
            {
                request.Query.TryGetValue("project", out var projectDisplay);
                var projectId = ProjectIdFromDisplay(projectDisplay);

                if (request.Body.Length > 0)
                {
                    // A write goes through the session's own writer, so it carries everything
                    // a host rewrite carries: the baseline bookkeeping and the engine's
                    // live-copy correction (the stale-problems lesson). This is also the
                    // bridge's first limb: another editor pushing code into a running VBE.
                    WriteModule(moduleName, request.Body, projectId, hostRewrite: true);
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply);
                }

                // live=1 reads the SURFACE's copy rather than the workbook's.
                //
                // They differ for as long as the developer has typed and the write-back timer has
                // not fired, which is exactly the window every typing behaviour lives in: smart
                // Enter, comment continuation and auto-indent all produce text that only exists in
                // the editor until it is written. Without this there was no way to read what
                // typing produced, so those features could only be checked by eye (2026-08-08).
                if (request.Query.TryGetValue("live", out var liveFlag) && liveFlag != "0")
                {
                    var live = _editorSurface?.TextOf(moduleName, DisplayFromProjectId(projectId));
                    return live is null
                        ? System.Text.Json.JsonSerializer.Serialize(
                            new DebugErrorReply($"the surface holds no text for {moduleName}"),
                            DebugJsonContext.Default.DebugErrorReply)
                        : System.Text.Json.JsonSerializer.Serialize(
                            new DebugModuleReply(moduleName, DisplayFromProjectId(projectId), live),
                            DebugJsonContext.Default.DebugModuleReply);
                }

                using var found = FindComponent(moduleName, projectId, out var foundProject);
                var source = found is null ? null : ProjectReader.ReadSource(found);
                return source is null
                    ? System.Text.Json.JsonSerializer.Serialize(
                        new DebugErrorReply($"no module named {moduleName}"), DebugJsonContext.Default.DebugErrorReply)
                    : System.Text.Json.JsonSerializer.Serialize(
                        new DebugModuleReply(moduleName, DisplayFromProjectId(foundProject), source),
                        DebugJsonContext.Default.DebugModuleReply);
            }

            case "caret"
                when request.Query.TryGetValue("line", out var caretLineText)
                    && int.TryParse(caretLineText, out var caretLine) && caretLine >= 1:
            {
                // Aiming, not scrolling. Every editor command acts on the caret - the host
                // syncs it into the native pane first - so a Run or a Step meant for one
                // procedure has to put the caret inside it, and revealLine cannot: pressing
                // Run with the caret on line 1 opens the editor's Macros dialog and waits
                // (2026-08-06). An optional module navigates there first.
                var caretColumn = request.Query.TryGetValue("column", out var columnText)
                    && int.TryParse(columnText, out var parsedColumn) ? parsedColumn : 1;

                if (request.Query.TryGetValue("module", out var caretModule) && caretModule.Length > 0)
                {
                    request.Query.TryGetValue("project", out var caretProject);
                    GoTo(caretModule, caretLine, caretColumn, caretProject);
                }

                _editorSurface?.SetCaret(caretLine, caretColumn);
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply);
            }

            case "placement":
                RefreshSurfacePlacement();
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply);

            default:
                /*
                 * "UNKNOWN ROUTE" IS OFTEN A LIE, and it sends the reader to the wrong place.
                 *
                 * Seventeen of the cases above are guarded on their arguments -- `case "caret"
                 * when the line parses`, and so on -- and a guard that does not hold falls through
                 * to here. So calling a route that exists, with an argument it will not accept,
                 * is answered by being told the route does not exist. Measured 2026-08-07:
                 * `caret?line=-1` answered "unknown route caret", and the next minute was spent
                 * looking for a spelling mistake in a name that was spelled correctly.
                 *
                 * Naming both possibilities costs nothing and points at the argument, which is
                 * what it is nearly always going to be: a route name is copied from the docs and
                 * an argument is computed.
                 */
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugErrorReply(
                        $"no route '{request.Route}' accepted this request. Either there is no such "
                        + "route, or there is and its required arguments were missing or rejected: "
                        + "many routes are guarded on theirs. "
                        + $"Given: {(request.Query.Count == 0 ? "(no arguments)" : string.Join(", ", request.Query.Select(pair => $"{pair.Key}={pair.Value}")))}"),
                    DebugJsonContext.Default.DebugErrorReply);
        }
    }
#endif

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

        // In the frame's message chain, so a resize re-places the surface synchronously —
        // before the native layout paints — instead of a posted event later. The event route
        // stays as the correcting pass.
        _frameSubclass ??= FrameSubclass.Install(host, PlaceSurfaceFast);

        _editorSurface.KeyPressed = OnSurfaceKey;
        _editorSurface.ModuleRequested = ShowModule;
        _editorSurface.NavigateRequested = GoTo;
        _editorSurface.CommandRequested = RunCommand;
        // The document names its workbook by display name; the write path resolves that to the
        // project identity so a same-named module in another workbook is never the one written.
        _editorSurface.TextChanged = (component, project, text) =>
            WriteModule(component, text, ProjectIdFromDisplay(project));
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
        _editorSurface.PanelChanged = OnPanelChanged;
        _editorSurface.MenuRequested = OnMenuRequested;
        _editorSurface.MenuExecuteRequested = OnMenuExecuteRequested;
        _editorSurface.PropertyEditRequested = OnPropertyEdit;
        _editorSurface.ComponentSelected = OnComponentSelected;
        _editorSurface.ModuleCloseRequested = OnModuleCloseRequested;
        _editorSurface.ComponentInsertRequested = InsertComponent;
        _editorSurface.CompletionRequested = OnCompletionRequested;
        _editorSurface.HoverRequested = OnHoverRequested;
        _editorSurface.SignatureHelpRequested = OnSignatureHelpRequested;
        _editorSurface.SmartEnterRequested = OnSmartEnterRequested;
        _editorSurface.CanonicalCaseRequested = OnCanonicalCaseRequested;
        _editorSurface.LoopSyncRequested = OnLoopSyncRequested;
        _editorSurface.CodeActionsRequested = OnCodeActionsRequested;
        _editorSurface.NavigationRequested = OnNavigationRequested;
        _editorSurface.RenameRequested = OnRenameRequested;
        _editorSurface.ModuleRenameRequested = OnModuleRenameRequested;
        _editorSurface.OutlineRequested = OnOutlineRequested;
        _editorSurface.SemanticTokensRequested = OnSemanticTokensRequested;
        _editorSurface.LiveAnalysisDue = OnLiveAnalysisDue;
        _editorSurface.LiveTextPushed = (module, full, edits) => _analysis?.NotifyLiveText(module, full, edits);

        // The active-line hold: typing on a line hides the verdicts about it, and the caret
        // settling anywhere else brings them back. Both handlers run on the host thread, and
        // both republish only when the hold actually changed — a keystroke on an already-held
        // line and a caret resting where it was cost nothing.
        _editorSurface.LineTyped = line =>
        {
            if (_editorSurface?.Module is { } typedModule && _activeLineHold.Begin(typedModule, line))
            {
                PublishMarkersToSurface();
                PublishFindingsToSurface();
            }
        };
        _editorSurface.CaretLineSettled = line =>
        {
            if (_activeLineHold.Release(_editorSurface?.Module, line))
            {
                PublishMarkersToSurface();
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
            // nothing — the tabs stayed gone (the tap showed every republish EXCEPT
            // setModules, 2026-08-06).
            _lastModulesKey = null;
            _lastLanguageFactsKey = null;
            _editorSurface?.ShowInstallPath(Interop.ShimModule.Directory);
            _hostChrome ??= HostChrome.Install(CodePaneTracker.MainWindow(), Interop.ShimModule.Directory);
            PublishModules();
            PublishProjects();
            PublishFindingsToSurface();
            PublishMarkersToSurface();
            PublishBreakpoints();
            PublishProperties();
            UpdateDebugState();

#if DEBUG
            InstallConsoleRing();
#endif
        };

        // While the loader shows, placement is re-asserted on its heartbeat: the editor is still
        // arranging itself — restoring its size, raising its own bands — and with no pane open
        // there is no window event to notice any of it. Without this, the loader keeps covering
        // the window as it was at the first placement, and a band of native chrome outlives it.
        _editorSurface.LoadingPulse = RefreshSurfacePlacement;

        // Now rather than at start-up. The editor answers that these windows are visible before
        // it has created them, so hiding one then closes something with no window behind it and
        // there is nothing to identify afterwards.
        HideReplacedWindows();
        HideNativeToolbars();
        PrepareLocalsGhost();
        PrepareWatchGhost();

        // Both ghosts are read from one dedicated thread; the host thread only asks and looks.
        // Reading from here re-enters the editor's own accessibility provider and dies in
        // native faults — GhostReaderThread carries the story.
        _ghostReaders = GhostReaderThread.Start(_localsPalette, _watchPalette);

        DarkenTitleBar(host);

        return true;
    }

    /// <summary>True while the surface is up over an editor that has no panes anywhere.</summary>
    private bool _watchingEmpty;

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
    private void GoTo(string component, int line, int column, string? projectDisplay = null)
    {
        try
        {
            var projectId = ProjectIdFromDisplay(projectDisplay);
            using var pane = FindCodePane(component, projectId);
            if (pane is null)
            {
                Log.Info($"navigate: no pane for {component}");
                return;
            }

            pane.Invoke("Show");
            pane.Invoke("SetSelection", line, column, line, column);

            // The surface shows one (module, WORKBOOK) pair, and a module name is not unique
            // across workbooks: two open workbooks can each hold a Helpers. Comparing the name
            // alone made a navigation to the other workbook's Helpers a no-op on the surface —
            // the native pane moved and the surface stayed where it was, showing a different
            // workbook's module of the same name (2026-08-07). A navigation that names no
            // project keeps the old meaning: whichever one is shown.
            var alreadyShowing =
                string.Equals(_editorSurface?.Module, component, StringComparison.OrdinalIgnoreCase)
                && (projectId is null
                    || string.Equals(_shownProject, projectId, StringComparison.OrdinalIgnoreCase));

            if (alreadyShowing)
            {
                _editorSurface!.Reveal(line);
            }
            else if (_editorSurface is not null)
            {
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
        }
        catch (Exception ex)
        {
            Log.Error($"navigate: could not go to {component}({line},{column})", ex);
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

            // The read of the projects belongs to the thread that owns them. The door is the
            // overlay's action timer, and it answers false while there is no surface to carry
            // it — the service retries rather than reading from the wrong thread.
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
            // frame routes already follow rectangles synchronously. Everything below — the
            // full placement with policing and bands, the module and project publishes, debug
            // state, resync — is for events that changed WHAT is on screen, not where its
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
                // may read the module — F5 is one keystroke away — so the text is made true.
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
    private void ExecuteEditorCommand(int command)
    {
        if (command == 0)
        {
            return;
        }

        // The Object Browser is ours (developer, 2026-08-05): a floating themed window of our
        // own over the typelib catalog, outside the canvas. The native window — which cannot
        // float, cannot be adopted, and paints only docked (lesson 32) — is never opened at all.
        if (command == VbeCommands.Command.ObjectBrowser)
        {
            OpenBrowserPalette();
            return;
        }

        // The editor runs what the module holds, and acts on its own caret. Both are brought up to
        // date here, at the one moment it matters: running code the developer has not finished
        // typing is worse than a short pause before it starts. A toolbar button also takes focus
        // off the surface, which is when the two carets are furthest apart.
        _editorSurface?.FlushEdits();
        SyncCaretToPane();

        // Toggling a breakpoint is bookkeeping as well as a command. The editor cannot report which
        // lines carry one, so the record kept here is the only thing the surface can draw from, and
        // a route that skips it sets a breakpoint that is real and invisible.
        if (command == VbeCommands.Command.ToggleBreakpoint)
        {
            ToggleBreakpoint(_editorSurface?.CaretLine ?? 0);
            return;
        }



        var ran = VbeCommands.Execute(_editor, command);

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

        // A save just cleaned the workbook — but the flag flips a beat AFTER the command
        // returns, so one immediate republish read the old value and the dot lingered. The
        // next few polls re-derive it; the change-key keeps the repeats free.
        if (command == VbeCommands.Command.Save)
        {
            PublishModules();
            _resyncPanePolls = Math.Max(_resyncPanePolls, 3);
        }

        WatchDebugState();

        // A command can change the native window landscape — the Object Browser above all —
        // and no event the tracker recognises announces it. The menu route always re-derived
        // placement after executing; this route learned the same manners (2026-08-05, the
        // Browser opening invisible under the surface).
        RefreshSurfacePlacement();
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
    /// applied, so replacing once is both faster and the only version whose failure mode is a
    /// module unchanged rather than a module half written.
    ///
    /// Writing resets the project, which discards any running state. That is what the host's own
    /// editor does when a module is edited, so it is parity rather than a regression, and it is
    /// why this is debounced rather than done per keystroke.
    /// </summary>
    private void WriteModule(string component, string text, string? ownerProject = null, bool hostRewrite = false)
    {
        try
        {
            // A write is normally about the module on the surface, so it goes to the shown
            // project's component — never to a same-named module in another workbook that
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
                return;
            }

            // The baseline belongs to the workbook actually found: a line diff computed against
            // another workbook's same-named module would write a merge of the two.
            var writtenKey = WrittenKey(component, DisplayFromProjectId(foundOwner ?? owner));

            // The changed lines alone, when the last read-back says where they are. Replacing a
            // whole large module makes the host reparse every line of it — seconds, on the
            // thread the keystrokes live on — where typing only ever touches a few. The whole
            // replace below remains for a module with no baseline or a rewrite too large to
            // call an edit.
            var baseline = _writtenModules.TryGetValue(writtenKey, out var known) ? known : null;
            var wroteDiff = baseline is not null && TryWriteLineDiff(module, baseline, text);

            if (!wroteDiff)
            {
                var existing = module.GetInt32("CountOfLines");
                if (existing > 0)
                {
                    module.Invoke("DeleteLines", 1, existing);
                }

                // A module with nothing in it is a legitimate state, and asking the host to add
                // an empty string to one is not.
                if (text.Length > 0)
                {
                    module.Invoke("AddFromString", text);
                }
            }

            // Read straight back and remembered, but not pushed into the surface.
            //
            // The editor rewrites what it is given, and its rewrites are the kind a developer is in
            // the middle of doing for themselves: it completes the parentheses on a procedure and
            // inserts a blank body for one. Sending that back mid-keystroke duplicated what had
            // just been typed and inserted lines nobody asked for. What it holds is remembered as
            // the baseline instead, so a later comparison sees changes made by something else and
            // not the editor's own tidying of our own write.
            var stored = ProjectReader.ReadSource(found);
            _writtenModules[writtenKey] = stored ?? text;

            if (hostRewrite)
            {
                // The engine's live copy of a module OUTRANKS its seeded copy everywhere —
                // search, completion, diagnosis — which is what keeps answers exact
                // mid-keystroke, and the typing path maintains it by streaming edits ahead of
                // the write. A host rewrite bypasses the page, so the live copy is corrected
                // here, or the engine keeps diagnosing text that no longer exists — the
                // problems of a discarded edit survived the close and the reopen (2026-08-05).
                _analysis?.NotifyLiveText(component, stored ?? text, null, owner);

                // AND THE SURFACE, which this used to leave behind entirely.
                //
                // "A host rewrite bypasses the page" was the reason given for correcting the
                // engine here, and the page was then never corrected at all: a module written
                // from outside while its pane is open kept whatever text the surface last had.
                // Found 2026-08-08 with a freshly built fixture — the workbook held 42 lines,
                // the native pane held them too, and the editor showed an EMPTY document, so
                // every breakpoint on it was refused as "not an executable statement" because
                // the line being asked about did not exist on the surface.
                //
                // Sync only touches a document the surface actually holds, and leaves an
                // unwritten one alone, so this cannot overwrite an edit in flight.
                _editorSurface?.Sync(component, DisplayFromProjectId(foundOwner ?? owner), stored ?? text);
            }

            Log.Info($"write: {component}, {text.Length} character(s){(wroteDiff ? " as a line diff" : string.Empty)}"
                     + (stored is not null && stored != text ? " (the editor reformatted it)" : string.Empty)
                     + (hostRewrite ? " (host rewrite; the engine's live copy follows)" : string.Empty));

            // The write just made the workbook dirty; the tab dots follow. The page skips the
            // rebuild when nothing it draws has changed, so a write that changed no flag is free.
            PublishModules();

            // The full pass re-reads every module, reseeds the engine, and diagnoses the whole
            // project — work worth doing, but not per pause: the live pass keeps the shown
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
        }
        catch (Exception ex)
        {
            Log.Error($"write: {component} could not be updated", ex);
        }
    }

    /// <summary>
    /// Writes only the lines that changed between the baseline and the new text: the common
    /// prefix and suffix are found, and the window between them is deleted and re-inserted at
    /// one anchor, so nothing shifts under anything. False when the change is too large to be
    /// typing — a paste of a module's worth of text is a whole replace, honestly.
    /// </summary>
    private static bool TryWriteLineDiff(DispatchObject module, string baseline, string text)
    {
        const int LargestDiffLines = 400;

        if (baseline == text)
        {
            return true;
        }

        var oldLines = SplitPhysicalLines(baseline);
        var newLines = SplitPhysicalLines(text);

        var prefix = 0;
        var maxPrefix = Math.Min(oldLines.Length, newLines.Length);
        while (prefix < maxPrefix && oldLines[prefix] == newLines[prefix])
        {
            prefix++;
        }

        var suffix = 0;
        var maxSuffix = Math.Min(oldLines.Length, newLines.Length) - prefix;
        while (suffix < maxSuffix
               && oldLines[oldLines.Length - 1 - suffix] == newLines[newLines.Length - 1 - suffix])
        {
            suffix++;
        }

        var oldWindow = oldLines.Length - prefix - suffix;
        var newWindow = newLines.Length - prefix - suffix;

        if (oldWindow > LargestDiffLines || newWindow > LargestDiffLines)
        {
            return false;
        }

        // Never ask to delete lines that are not there.
        //
        // An EMPTY module has CountOfLines 0, but the empty baseline splits into one empty line —
        // so the window says "delete 1 from line 1" and the editor refuses the whole write with
        // "Invalid procedure call or argument". Nothing is written, and the only place it is said
        // is the log: the write route's reply looks like every other success. It took the code out
        // of the first module of every fixture built through the door, leaving a workbook that
        // looked right and no longer exercised what it existed for (2026-08-07).
        var present = module.GetInt32("CountOfLines");
        if (prefix + oldWindow > present)
        {
            oldWindow = Math.Max(0, present - prefix);
        }

        if (oldWindow > 0)
        {
            module.Invoke("DeleteLines", prefix + 1, oldWindow);
        }

        if (newWindow > 0)
        {
            module.Invoke("InsertLines", prefix + 1,
                string.Join("\r\n", newLines.Skip(prefix).Take(newWindow)));
        }

        return true;
    }

    private static string[] SplitPhysicalLines(string text) =>
        text.Split(["\r\n", "\n"], StringSplitOptions.None);

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
    /// did, and the `breakpoints` route then answered `helpers @ renamefixture.xlsm` — so a
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

    /// <summary>Project modes, as the editor numbers them.</summary>
    private const int BreakMode = 1;
    private const int DesignMode = 2;

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
    private static bool CanBreakOn(string? line)
    {
        var code = line?.Trim();
        if (string.IsNullOrEmpty(code))
        {
            return false;
        }

        if (code.StartsWith('\'') || code.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (StartsWithWord(code, "Option", "Attribute", "Declare", "Dim", "Const", "Type", "Enum", "End Type", "End Enum"))
        {
            return false;
        }

        // A modifier followed by anything that is not a procedure is a declaration.
        foreach (var modifier in (string[])["Public", "Private", "Friend", "Static", "Global"])
        {
            if (StartsWithWord(code, modifier))
            {
                var rest = code[modifier.Length..].TrimStart();
                return StartsWithWord(rest, "Sub", "Function", "Property");
            }
        }

        return true;
    }

    private static bool StartsWithWord(string text, params ReadOnlySpan<string> words)
    {
        foreach (var word in words)
        {
            if (!text.StartsWith(word, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // A whole word, so "Constant" is not "Const" and "Dimension" is not "Dim".
            if (text.Length == word.Length || !char.IsLetterOrDigit(text[word.Length]) && text[word.Length] != '_')
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Answers the Search panel: the engine searches the modules it holds — live text where a
    /// module is being edited — and the hits come back with workbook display names. The edits
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
    /// What the Browser lists at its top level: the open projects first — the developer's
    /// own code is what they browse most — then every referenced type library.
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
            using var code = component?.GetObject("CodeModule");
            var count = code?.GetInt32("CountOfLines") ?? 0;
            var source = count > 0 ? code!.GetStringIndexed("Lines", 1, count) : null;
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
    /// scan, not a parse — it reads the declaration lines the developer wrote and skips
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
    /// module the hit lives in — ReplaceLine, surgical, no module reset. The shown module's
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
                    // resynced NOW — waiting for a pane event left the editor showing the old
                    // text while the panel claimed the replacement had happened ("replace is
                    // not working", 2026-08-04) — and the full pass re-reads the rest.
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

                // A replace is a host rewrite: the engine's live copy of this module — which
                // outranks its seeded copy in every answer — still holds the pre-replace text
                // for any module the developer has typed in, and the dirty-dot comparison
                // still holds the pre-replace text as current. Both adopt the truth here.
                if (inModule > 0 && component is not null
                    && ProjectReader.ReadSource(component) is { } adopted)
                {
                    _writtenModules[WrittenKey(group.Key.Module, DisplayFromProjectId(group.Key.ProjectId))] = adopted;
                    _analysis?.NotifyLiveText(group.Key.Module, adopted, null, group.Key.ProjectId);
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

        if (!clearing && !CanBreakOn(_editorSurface?.LineAt(line)))
        {
            // Refused silently on screen, by design (the developer, 2026-08-04): the hover
            // preview already showed an orange cross where no breakpoint can go, and a click
            // there draws nothing — nothing may ever appear that looks like a breakpoint the
            // developer did not get. The page mirrors CanBreakOn for that preview.
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

            if (!VbeCommands.Execute(_editor, VbeCommands.Command.ToggleBreakpoint))
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
    private void PublishProperties()
    {
        var surface = _editorSurface;
        var target = _propertiesTarget ?? surface?.Module;
        if (surface is null || target is null)
        {
            return;
        }

        try
        {
            using var found = FindComponent(target);
            using var properties = found?.GetObject("Properties");
            if (found is null || properties is null)
            {
                return;
            }

            var componentType = found.GetInt32("Type");
            var count = properties.GetInt32("Count");

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

                using var property = properties.GetItem(name);
                if (property is null)
                {
                    continue;
                }

                var shownName = componentType != DocumentComponent
                    && string.Equals(name, "Name", StringComparison.OrdinalIgnoreCase)
                    ? "(Name)"
                    : name;

                entries.Add(DescribeProperty(shownName, property));
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
    /// One property, rendered for the panel. Whether it is offered for editing comes from the type
    /// it currently holds: values of simple types are editable, objects and the unreadable are not.
    /// The editor can still refuse an edit, and that refusal is reported when it happens.
    /// </summary>
    private static SurfacePropertyEntry DescribeProperty(string shownName, DispatchObject property)
    {
        try
        {
            var (kind, display) = property.ReadProperty("Value");
            var writable = kind is VarEnum.VT_BSTR or VarEnum.VT_BOOL
                or VarEnum.VT_I2 or VarEnum.VT_I4 or VarEnum.VT_INT
                or VarEnum.VT_R4 or VarEnum.VT_R8 or VarEnum.VT_EMPTY;

            return new SurfacePropertyEntry(shownName, display, writable, kind == VarEnum.VT_BOOL);
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
        PublishProperties();
    }

    /// <summary>
    /// Writes a property the developer edited in the panel, in the type the property currently
    /// holds. A refusal is reported in the editor's own words; a rename is adopted everywhere the
    /// old name was a key.
    /// </summary>
    private void OnPropertyEdit(string component, string name, string value)
    {
        try
        {
            using var found = FindComponent(component);
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

            using var properties = found.GetObject("Properties");
            using var property = properties?.GetItem(name);
            if (property is null)
            {
                Log.Info($"properties: {component}.{name} no longer exists");
                PublishProperties();
                return;
            }

            if (!WriteProperty(property, value, out var complaint))
            {
                _editorSurface?.Notify($"{name}: {complaint}");
                PublishProperties();
                return;
            }

            Log.Info($"properties: {component}.{name} = '{value}'");

            // Renaming changes the key everything else holds: the write baseline, the breakpoint
            // record, the tabs, the explorer, and the name the surface files the document under.
            var renamed = string.Equals(name, "Name", StringComparison.OrdinalIgnoreCase)
                ? found.GetString("Name")
                : null;

            if (renamed is not null && !string.Equals(renamed, component, StringComparison.OrdinalIgnoreCase))
            {
                AdoptRename(component, renamed);
            }
            else
            {
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
    /// component has to say so — and each one remembering a different subset is exactly how this
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
        // answers within a tick or two. Until it does — or while reads are failing — an empty
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
    /// Tells the page which debug mode the editor is in — "design", "run", or "break" — so
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

            if (mode != BreakMode)
            {
                if (_inBreak)
                {
                    _inBreak = false;
                    _editorSurface?.ShowCurrentLine(null);

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

            if (name is not null && name != _editorSurface?.Module)
            {
                ShowModuleInSurface(name);
            }

            _editorSurface?.ShowCurrentLine(line);
            _editorSurface?.Reveal(line);

            if (!_inBreak)
            {
                Log.Info($"debug: stopped at {name}({line})");
            }

            _inBreak = true;
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

    /// <summary>Whether the developer is looking at the Immediate panel.</summary>
    private bool _watchingImmediate;

    /// <summary>
    /// Tracks each panel's own visibility transitions. With two docks the page can show more
    /// than one panel at once, so a message about one panel says nothing about the others —
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
        var interval = _pollsRemaining > 0 ? DebugPollMilliseconds
            : _watchingImmediate ? ImmediatePollMilliseconds
            : _watchingEmpty ? EmptyWorkspacePollMilliseconds
            : 0;

        _editorSurface?.Poll(interval);
#if DEBUG
        PerfCounters.Poll(interval);
#endif
    }

    /// <summary>
    /// How often the project tree is refreshed while the editor has no panes. With no panes
    /// there are no window events, and the explorer is the only way the first module gets
    /// opened, so it cannot be allowed to sit stale.
    /// </summary>
    private const uint EmptyWorkspacePollMilliseconds = 1000;

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
        // identities resolve once per workbook rather than once per document — the resolve
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

    /// <summary>One tick of the execution watch.</summary>
    private void PollDebugState()
    {
#if DEBUG
        // Stamped per TICK, not per configuration change: this is the pulse the debug api
        // reports, and it should mean "the host thread ran our periodic work just now".
        // Note the honest limit - an idle editor stops polling altogether, so an old
        // heartbeat means blocked only while something should be watching (see doctor).
        PerfCounters.Beat();
#endif

        // While the editor has no panes, the tree is the only living thing on screen and the
        // only route to a first module, so it follows the project as it grows.
        if (_watchingEmpty)
        {
            PublishProjects();
            AdoptOpenModuleIfEmpty();
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
        // this change: it only ever holds the pane windows it can match — the active one, in
        // practice — so closing a HIDDEN pane leaves its picture identical, Changed never
        // fires, and the strip kept showing the closed module's tab ("the tab X doesn't close
        // it if it's not focused", 2026-08-04, the real mechanism at last). The strip's truth
        // is the object model's open list; after a close it is re-read and re-sent, and the
        // page skips the rebuild when nothing actually changed.
        if (_resyncPanePolls > 0)
        {
            _resyncPanePolls--;
            _codePanes?.Refresh();
        }
        else if (_codePanes is { Stale: true })
        {
            _codePanes.Refresh();
        }

        // Every tick, not only the resync ones: a save made on the host's side of the fence —
        // Excel's own Ctrl+S, an autosave — flips the dirty flags with no event we hear, and
        // the dots must follow. The change-key inside makes an unchanged strip cost a read
        // and no message.
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
    /// Starts reading the Immediate window, its handle already worked out — by caption match
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

    /// <summary>
    /// Gives the page a module's text without activating it.
    ///
    /// The page holds a module's text once it has been ACTIVATED, so a workspace opened onto eight
    /// modules holds one — and anything that draws a module it is not showing (peeking a
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
        // one did not — so the component was never found and the answer never came (2026-08-07).
        var projectId = ProjectIdFromDisplay(projectDisplay) ?? _shownProject;

        surface.RunOnHostThread(() =>
        {
            try
            {
                using var component = FindComponent(moduleName, projectId, out var owner);
                using var code = component?.GetObject("CodeModule");
                if (code is null)
                {
                    Log.Info($"document: {moduleName} could not be found to publish");
                    return;
                }

                var count = code.GetInt32("CountOfLines");
                var text = count > 0 ? code.GetStringIndexed("Lines", 1, count) ?? string.Empty : string.Empty;
                Log.Info($"document: publishing {moduleName}, {count} line(s), without activating it");
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
        // matter. The window certainly exists by now — start-up toggled it visible — and the
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
                Log.Info("immediate: the line left the project stopped, resetting");
                try
                {
                    ExecuteEditorCommand(VbeCommands.Command.Reset);
                    Log.Info("immediate: reset executed");
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
    /// Answers a completion request from the surface.
    ///
    /// The module name and its live text are captured here, on the host thread; the engine round
    /// trip happens off it, because completions ride on every keystroke and the developer must
    /// never wait for one; and the answer is marshalled back, because the browser may only be
    /// spoken to from the thread that owns it. A request that fails answers empty rather than
    /// never: the editor is left waiting on nothing.
    /// </summary>
    private void OnCompletionRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowCompletions(requestId, []);
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceCompletionItem[] items = [];

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.CompleteAsync(module, source, offset, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    items = [.. answered.Select(item => new SurfaceCompletionItem(
                        item.Label,
                        item.Kind,
                        item.Detail,
                        item.Documentation,
                        item.InsertText,
                        item.FilterText,
                        item.SortText))];
                }

                Log.Info($"completion: {module}@{offset} -> {items.Length} item(s)");
            }
            catch (Exception ex)
            {
                Log.Info($"completion: {module}@{offset} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowCompletions(requestId, items));
        });
    }

    /// <summary>
    /// Answers a hover request from the surface, the same way a completion is answered: capture
    /// on the host thread, resolve off it, marshal the answer back. A request that fails answers
    /// empty rather than never.
    /// </summary>
    private void OnHoverRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowHover(requestId, null);
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceHoverPayload? payload = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.HoverAsync(module, source, offset, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    payload = new SurfaceHoverPayload(
                        answered.Signature,
                        answered.Details,
                        answered.Documentation,
                        answered.Span.Start,
                        answered.Span.End);
                }

                Log.Info($"hover: {module}@{offset} -> {(payload is null ? "nothing" : payload.Signature)}");
            }
            catch (Exception ex)
            {
                Log.Info($"hover: {module}@{offset} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowHover(requestId, payload));
        });
    }

    /// <summary>Answers a call-tip request from the surface, the same way a hover is answered.</summary>
    private void OnSignatureHelpRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowSignatureHelp(requestId, null);
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceSignatureInfo? payload = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.SignatureHelpAsync(module, source, offset, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    payload = new SurfaceSignatureInfo(
                        answered.Label,
                        [.. answered.Parameters.Select(parameter =>
                            new SurfaceSignatureParameter(parameter.Label, parameter.Documentation))],
                        answered.ActiveParameter,
                        answered.Documentation,
                        answered.Details);
                }

                Log.Info($"signature: {module}@{offset} -> {(payload is null ? "nothing" : payload.Label)}");
            }
            catch (Exception ex)
            {
                Log.Info($"signature: {module}@{offset} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowSignatureHelp(requestId, payload));
        });
    }

    /// <summary>
    /// Answers a Smart Enter request from the surface: what the Enter that just went in should
    /// leave behind. Answered the way a completion is: capture on the host thread, resolve off
    /// it, marshal the answer back. A request that fails answers empty rather than never.
    /// </summary>
    private void OnSmartEnterRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowSmartEnter(requestId, [], null);
            return;
        }

        // Captured on the host thread, before the hop: the developer's typing choices decide
        // what Enter leaves behind, and reading them off the session from a background task
        // would race a settings change.
        var typing = _settings;

        _ = Task.Run(async () =>
        {
            SurfaceTextEdit[] edits = [];
            int? caret = null;

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.SmartEnterAsync(module, source, offset, typing, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    edits = [.. answered.Edits.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))];
                    caret = answered.Caret;
                }

                if (edits.Length > 0)
                {
                    Log.Info($"smartEnter: {module}@{offset} -> {edits.Length} edit(s)");
                }
            }
            catch (Exception ex)
            {
                Log.Info($"smartEnter: {module}@{offset} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowSmartEnter(requestId, edits, caret));
        });
    }

    /// <summary>
    /// Answers a canonical-case request from the surface: the case corrections for a span,
    /// resolved from the same project facts completion uses.
    /// </summary>
    private void OnCanonicalCaseRequested(int requestId, int start, int end, bool single, bool completeHeader)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowCanonicalCase(requestId, []);
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceTextEdit[] edits = [];

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.CanonicalCaseAsync(module, source, start, end, single, completeHeader, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    edits = [.. answered.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))];
                }

                if (edits.Length > 0)
                {
                    Log.Info($"canonicalCase: {module}@{start}..{end} -> {edits.Length} edit(s)");
                }
            }
            catch (Exception ex)
            {
                Log.Info($"canonicalCase: {module}@{start}..{end} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowCanonicalCase(requestId, edits));
        });
    }

    /// <summary>
    /// Answers a loop-sync request from the surface: the paired iterator rename, when the edit
    /// at the offset touched one side of a For/Next pair.
    /// </summary>
    private void OnLoopSyncRequested(int requestId, int offset)
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        var source = surface?.Text;

        if (surface is null || module is null || source is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowLoopSync(requestId, []);
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceTextEdit[] edits = [];

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.LoopSyncAsync(module, source, offset, deadline.Token)
                    .ConfigureAwait(false);

                if (answered is not null)
                {
                    edits = [.. answered.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))];
                }

                if (edits.Length > 0)
                {
                    Log.Info($"loopSync: {module}@{offset} -> {edits.Length} edit(s)");
                }
            }
            catch (Exception ex)
            {
                Log.Info($"loopSync: {module}@{offset} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowLoopSync(requestId, edits));
        });
    }

    /// <summary>
    /// Answers a quick-fix request from the surface: what can be fixed over a span, and the edits
    /// that would fix it. Answered the same way a hover is, and empty on failure for the same
    /// reason — a lightbulb that does not appear is what the developer already sees when there is
    /// nothing to fix.
    /// </summary>
    private void OnCodeActionsRequested(int requestId, int start, int end)
    {
        var surface = _editorSurface;
        var module = surface?.Module;

        if (surface is null || module is null || _analysis is not { } analysis)
        {
            _editorSurface?.ShowCodeActions(requestId, []);
            return;
        }

        _ = Task.Run(async () =>
        {
            SurfaceCodeAction[] actions = [];

            try
            {
                using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(3));
                var answered = await analysis.CodeActionsAsync(module, start, end, deadline.Token)
                    .ConfigureAwait(false);

                actions = [.. answered.Select(action => new SurfaceCodeAction(
                    action.Title,
                    action.IsPreferred ?? false,
                    action.Code,
                    action.Span.Start,
                    action.Span.End,
                    [.. action.Edits.Select(edit => new SurfaceTextEdit(edit.Start, edit.End, edit.Text))]))];

                if (actions.Length > 0)
                {
                    Log.Info($"codeAction: {module}@{start}..{end} -> {actions.Length} fix(es)");
                }
            }
            catch (Exception ex)
            {
                Log.Info($"codeAction: {module}@{start}..{end} failed ({ex.GetType().Name})");
            }

            surface.RunOnHostThread(() => surface.ShowCodeActions(requestId, actions));
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
            // Not "this module is empty" — "nothing here can answer yet". The page keeps what
            // it already shows.
            _editorSurface?.ShowOutline(requestId, [], failed: true);
            return;
        }

        // The tree names the workbook its row belongs to, which is what makes a shared module
        // name unfold the right workbook's procedures.
        var projectId = ProjectIdFromDisplay(projectDisplay);

        // No source travels with the request: the engine's live copy is exact — didChange rides
        // the same FIFO pipe ahead of this — and serialising a 918KB module once a second to
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
                    // of the workbook — and the page only knows the ones with a tab open. The
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

            // The writes are the host's own object model, so they belong on the host thread —
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
                        WriteModule(entry.Module, entry.Source, _shownProject, hostRewrite: true);
                        written.Add(entry.Module);

                        // The page holds a model per OPEN module. Syncing one that is not open is
                        // harmless — the page has nothing by that name to sync — and syncing one
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
    /// identifier, or is a keyword. Only then is the component renamed — the one step that can
    /// still be refused by the host — and only if THAT succeeds is a single line of code written.
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
    /// owns both, and both entry points — the explorer's Rename and renaming a module's name in
    /// code — land here so there is one order rather than two.
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
    /// One slot. A second rename replaces it — undo goes back one step, the way the operation is
    /// one step — and the texts are the ones read out of the modules immediately before writing,
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
                using var code = component?.GetObject("CodeModule");
                if (code is null)
                {
                    continue;
                }

                var count = code.GetInt32("CountOfLines");
                captured.Add((module, count > 0 ? code.GetStringIndexed("Lines", 1, count) ?? string.Empty : string.Empty));
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
    private void UndoRename(int requestId)
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        if (_undoableRename is not { } undo)
        {
            surface.ShowRenamed(requestId, null, null, [], 0, "There is no rename to undo.");
            return;
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
                WriteModule(target, text, undo.ProjectId, hostRewrite: true);
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
        // it has now — including the one about to change.
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
                WriteModule(target, entry.Source, projectId, hostRewrite: true);
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
                    location.Preview))];

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
    /// expression being typed. A result computed for text that has moved on is dropped — the
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
    private void RunCommand(string name)
    {
        var command = VbeCommands.ForName(name);
        if (command == 0)
        {
            Log.Info($"command: '{name}' is not one of ours");
            return;
        }

        ExecuteEditorCommand(command);
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
        _shownProject = owner;
        if (_analysis is not null)
        {
            _analysis.PreferredProject = owner;
        }

        // A project the engine has never been seeded with — a workbook just opened or created.
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

        var display = DisplayFromProjectId(owner);
        _writtenModules[WrittenKey(component, display)] = source;
        _editorSurface?.Show(component, display, source);

        // The engine's live copy starts from what is being shown; the keystrokes stream from
        // here as edits.
        _analysis?.NotifyLiveText(component, source, null);
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

        if (chrome.Mode == _titleMode && chrome.Workbook == workbook && chrome.Module == shown)
        {
            return;
        }

        chrome.Mode = _titleMode;
        chrome.Workbook = workbook;
        chrome.Module = shown;
        chrome.Apply();
    }

    private void PublishModules()
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

            // Closing the LAST pane leaves the surface holding a document nobody can see a
            // tab for, when no window event reaches the tracker to say so - the mirror of
            // the empty view that outlived its panes (2026-08-06). The object model is the
            // authority both ways: an empty open list with a module still shown IS the
            // empty workspace, and every close route passes through this publish.
            if (modules.Count == 0 && hadDocuments)
            {
                Log.Info("editor surface: the last module closed, showing the empty workspace");
                _watchingEmpty = true;
                surface.Clear();
                UpdatePolling();
            }

            // Dirty is a WORKBOOK fact — the editor persists all of a workbook's modules
            // together (probed 2026-08-04: a module edit flips Workbook.Saved false, Save
            // flips it true) — so it is read once per workbook and worn by every tab the
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

            string[] names = [.. modules.Select(m => m.Name)];
            string?[] projects = [.. modules.Select(m => m.Project)];
            bool[] dirty = [.. modules.Select(m => DirtyOf(m.Name, m.Project))];
            var active = surface.Module;
            var activeProject = DisplayFromProjectId(_shownProject);

            // Sent on change only: the polls re-derive this several times a second during an
            // episode, and an identical strip is not news to the page or the log.
            var key = string.Join("|", names) + "\n" + string.Join("|", projects)
                + "\n" + string.Join("|", dirty) + "\n" + active + "\n" + activeProject;
            if (key == _lastModulesKey)
            {
                return;
            }

            _lastModulesKey = key;
            Log.Verbose($"modules: publish [{string.Join(",", names)}] dirty [{string.Join(",", dirty)}]");
            surface.ShowModules(names, projects, active, activeProject, dirty);
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
    /// A workbook's Saved flag, by display name, through the same trust-free application route
    /// the evaluator uses. Null when it cannot be read — a missing dot is a small wrong, a
    /// lying dot is a large one, so unknown must never invent one.
    /// </summary>
    private bool? WorkbookSaved(string display)
    {
        try
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
                using var book = books!.GetItem(i);
                if (book is not null
                    && string.Equals(book.GetString("Name"), display, StringComparison.OrdinalIgnoreCase))
                {
                    return book.GetBool("Saved");
                }
            }
        }
        catch (Exception)
        {
            // The application answer went stale — a workbook closed mid-read, or the host is
            // busy. Re-found on the next read.
            _hostApp?.Dispose();
            _hostApp = null;
        }

        return null;
    }

    /// <summary>
    /// What the developer calls a project, from its identity alone: the file's name for a saved
    /// workbook, the identity itself otherwise. Lowercase for saved ones — comparisons on the
    /// page side are case-insensitive.
    /// </summary>
    /// <summary>
    /// A module's text reduced to something two sides can be compared on.
    ///
    /// Line endings are normalised and trailing blank lines dropped, because the host and the
    /// page genuinely disagree about both and neither disagreement is a defect: VBA stores CRLF
    /// and counts a trailing line the page does not draw. Everything else must match exactly —
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

    private static string? DisplayFromProjectId(string? projectId)
    {
        if (string.IsNullOrEmpty(projectId))
        {
            return null;
        }

        return projectId.Contains('\\') || projectId.Contains('/')
            ? Path.GetFileName(projectId)
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
                            owner = ProjectReader.Identity(project).DisplayName;
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

            var tree = new List<SurfaceProject>(projectCount);
            for (var i = 1; i <= projectCount; i++)
            {
                using var project = projects!.GetItem(i);
                using var components = project?.GetObject("VBComponents");
                if (project is null || components is null)
                {
                    continue;
                }

                var componentCount = components.GetInt32("Count");
                var members = new List<SurfaceComponent>(componentCount);

                for (var j = 1; j <= componentCount; j++)
                {
                    using var component = components.GetItem(j);
                    if (component?.GetString("Name") is { Length: > 0 } name && !IsScratchComponent(name))
                    {
                        members.Add(new SurfaceComponent(name, component.GetInt32("Type")));
                    }
                }

                // The cased name against the id, so the title bar can name the workbook on a tab
                // switch without a COM call. The id is a lowercased path and everything derived
                // from it reads "debugfixture.xlsm"; this is the spelling the shell uses.
                var display = WorkbookDisplayName(project);
                _projectNames[ProjectReader.Identity(project).Id] = display;

                tree.Add(new SurfaceProject(display, [.. members]));
            }

            surface.ShowProjects([.. tree]);
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

        return project.GetString("Name") ?? "VBAProject";
    }

    /// <summary>The project shown under a workbook name, or null when none matches.</summary>
    private DispatchObject? FindProjectByDisplayName(string? displayName)
    {
        if (string.IsNullOrEmpty(displayName))
        {
            return null;
        }

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

                if (string.Equals(WorkbookDisplayName(project), displayName, StringComparison.OrdinalIgnoreCase))
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
        // palettes a moment later — the editor only feeds a window with a paintable surface
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

    /// <summary>
    /// Turns the native Locals window into the panel's invisible data engine: floated through
    /// the object model, ghosted, parked off screen, and read by handle.
    ///
    /// The editor only feeds an on-screen Locals window (lesson 25) — but "on screen" turned
    /// out to mean "has a paintable surface". A LAYERED window renders into its own surface
    /// regardless of occlusion or position, so a floated palette with WS_EX_LAYERED at alpha
    /// zero, parked far off the virtual screen, is fed faithfully through every break and
    /// step while being impossible to see, click, or discover. Probed 2026-08-04: counter
    /// tracked 1 through 4 across steps at alpha 1, alpha 0, and at -20000,-20000
    /// (Probe-GhostLocals.ps1). The themed panel renders what the reader reads; nothing
    /// native is ever visible.
    ///
    /// Floating uses LinkedWindows.Remove on the window's own linked frame — pure object
    /// model. If any step refuses, the ghost is skipped, and the police pass hides the
    /// docked native window: the canvas stays pure and the panel sits idle.
    /// </summary>
    private void PrepareLocalsGhost()
    {
        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is null || window.GetInt32("Type") != 4)
                {
                    continue;
                }

                window.SetBool("Visible", true);

                // Undocked; a window already floating answers Remove with an error worth
                // nothing.
                try
                {
                    using var frame = window.GetObject("LinkedWindowFrame");
                    using var linked = frame?.GetObject("LinkedWindows");
                    linked?.InvokeWithObject("Remove", window);
                }
                catch (Exception)
                {
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
                catch (Exception)
                {
                }

                var caption = window.GetString("Caption");
                var palette = caption is { Length: > 0 }
                    ? CodePaneTracker.FindTopLevelByCaption(caption)
                    : 0;

                if (palette == 0)
                {
                    Log.Info("locals: the floated palette was not found; the native window stays");
                    return;
                }

                _localsPalette = palette;
                _localsPaletteExStyle = Win32.GetWindowLongPtr(palette, Win32.GwlExStyle);
                Win32.SetWindowLongPtr(palette, Win32.GwlExStyle,
                    (nint)(_localsPaletteExStyle | Win32.WsExLayered | Win32.WsExTransparent | Win32.WsExNoActivate));
                Win32.SetLayeredWindowAttributes(palette, 0, 0, Win32.LwaAlpha);
                Win32.SetWindowPos(palette, 0, -20000, -20000, 0, 0,
                    Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate);

                // The reader is not created here: it lives on the ghost reading thread, which
                // starts once both palettes are prepared and reports its own connect result.
                Log.Verbose($"locals: palette {palette:X} floated and ghosted");
                return;
            }
        }
        catch (Exception ex)
        {
            Log.Info($"locals: the ghost could not be prepared ({ex.GetType().Name}: {ex.Message})");
        }
    }

    /// <summary>
    /// Gives the palette back to the native editor: opaque, its styles restored, on screen,
    /// and hidden until someone asks for it. A stopped session must leave a usable window.
    /// </summary>
    private void RestoreLocalsPalette()
    {
        if (_localsPalette == 0)
        {
            return;
        }

        try
        {
            Win32.SetWindowPos(_localsPalette, 0, 300, 300, 0, 0,
                Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate);
            Win32.SetLayeredWindowAttributes(_localsPalette, 0, 255, Win32.LwaAlpha);
            Win32.SetWindowLongPtr(_localsPalette, Win32.GwlExStyle, (nint)_localsPaletteExStyle);

            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;
            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is not null && window.GetInt32("Type") == 4)
                {
                    window.SetBool("Visible", false);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"locals: the palette could not be restored ({ex.GetType().Name})");
        }

        _localsPalette = 0;
    }

    /// <summary>The floated Watches palette, its original extended styles kept for restoration.</summary>
    private nint _watchPalette;
    private long _watchPaletteExStyle;

    /// <summary>
    /// Turns the native Watches window into the Watch panel's invisible data engine, by the
    /// same route as the Locals ghost above (lesson 29): floated through the object model,
    /// layered at alpha zero, parked off screen, read by handle. If any step refuses, the
    /// ghost is skipped, and the police pass hides the docked native window.
    /// </summary>
    private void PrepareWatchGhost()
    {
        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is null || window.GetInt32("Type") != 3)
                {
                    continue;
                }

                window.SetBool("Visible", true);

                try
                {
                    using var frame = window.GetObject("LinkedWindowFrame");
                    using var linked = frame?.GetObject("LinkedWindows");
                    linked?.InvokeWithObject("Remove", window);
                }
                catch (Exception)
                {
                }

                try
                {
                    window.SetInt32("Left", 300);
                    window.SetInt32("Top", 300);
                    window.SetInt32("Width", 240);
                    window.SetInt32("Height", 150);
                }
                catch (Exception)
                {
                }

                var caption = window.GetString("Caption");
                var palette = caption is { Length: > 0 }
                    ? CodePaneTracker.FindTopLevelByCaption(caption)
                    : 0;

                if (palette == 0)
                {
                    Log.Info("watch: the floated palette was not found; the native window stays");
                    return;
                }

                _watchPalette = palette;
                _watchPaletteExStyle = Win32.GetWindowLongPtr(palette, Win32.GwlExStyle);
                Win32.SetWindowLongPtr(palette, Win32.GwlExStyle,
                    (nint)(_watchPaletteExStyle | Win32.WsExLayered | Win32.WsExTransparent | Win32.WsExNoActivate));
                Win32.SetLayeredWindowAttributes(palette, 0, 0, Win32.LwaAlpha);
                Win32.SetWindowPos(palette, 0, -20000, -20000, 0, 0,
                    Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate);

                // The reader is not created here: it lives on the ghost reading thread, the
                // same manner as the Locals ghost above.
                Log.Verbose($"watch: palette {palette:X} floated and ghosted");
                return;
            }
        }
        catch (Exception ex)
        {
            Log.Info($"watch: the ghost could not be prepared ({ex.GetType().Name}: {ex.Message})");
        }
    }

    /// <summary>
    /// Gives the Watches palette back to the native editor: opaque, its styles restored, on
    /// screen, and hidden until someone asks for it.
    /// </summary>
    private void RestoreWatchPalette()
    {
        if (_watchPalette == 0)
        {
            return;
        }

        try
        {
            Win32.SetWindowPos(_watchPalette, 0, 300, 300, 0, 0,
                Win32.SwpNoSize | Win32.SwpNoZOrder | Win32.SwpNoActivate);
            Win32.SetLayeredWindowAttributes(_watchPalette, 0, 255, Win32.LwaAlpha);
            Win32.SetWindowLongPtr(_watchPalette, Win32.GwlExStyle, (nint)_watchPaletteExStyle);

            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;
            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is not null && window.GetInt32("Type") == 3)
                {
                    window.SetBool("Visible", false);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"watch: the palette could not be restored ({ex.GetType().Name})");
        }

        _watchPalette = 0;
    }

    /// <summary>
    /// Hides every docked native toolbar, which the surface's toolbar and menus replace.
    ///
    /// Hidden only because everything on them is somewhere else now: each button is a menu
    /// command, the menus are all on the surface's bar, and every run and step command keeps the
    /// key it always had. The toggles that would bring the bars back are suppressed from the
    /// surface's View menu, so a hidden bar stays hidden. A docked bar left visible would claim
    /// rows the surface covers — since it stopped retreating for chrome, that is a toolbar on
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
        // way — diffing which pane stopped being visible across the hide — lost whenever the
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
    /// Publishes every finding to the surface's panel, across all modules — except the ones the
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
        _editorSurface?.ShowFindings([.. _findings
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
    /// A close asked for by the developer — the tab's X, its middle-click, Ctrl+W, or the tab
    /// menu. A module whose text still differs from the workbook's last saved text does not just
    /// close: the developer is asked first, and the answer comes back through the same message
    /// with a choice on it. "save" saves the workbook — the editor persists all of a workbook's
    /// modules together, so saving the workbook is what saving a module means — and "discard"
    /// writes the saved text back over the module, which is the closest thing to closing without
    /// saving that a document living inside a workbook can have.
    ///
    /// The question gates on the module's OWN text, not the workbook dot. The dot is a workbook
    /// fact and can stand on a sibling's changes; a question here offers a revert of THIS module,
    /// so it is only asked when this module's changes can be named — and therefore reverted.
    /// </summary>
    private void OnModuleCloseRequested(string component, string? projectDisplay, string? action)
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
                    return;
                }

                CloseModule(component, projectDisplay);
                return;
            }

            case "discard":
            {
                if (display is not null
                    && _savedBaselines.TryGetValue(BaselineKey(display, component), out var baseline))
                {
                    // The debounced write of the abandoned text must not chase the revert —
                    // this document's write only; a sibling tab's typing keeps its debounce.
                    _editorSurface?.DiscardEdits(component, display);

                    WriteModule(component, baseline, ProjectIdFromDisplay(display), hostRewrite: true);

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
                    // remainder — inventing a revert target would destroy real work.
                    Log.Info($"close: {component} has no saved text to revert to; closing as it is");
                }

                CloseModule(component, projectDisplay);
                return;
            }

            default:
            {
                if (display is not null && ModuleDiffersFromSaved(component, display))
                {
                    Log.Verbose($"close: {component} differs from {display}'s saved text; asking");
                    _editorSurface?.ConfirmClose(component, projectDisplay);
                    return;
                }

                CloseModule(component, projectDisplay);
                return;
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
    /// Saved flag is read by. False when the workbook cannot be found or the save is refused —
    /// the caller keeps the tab open on false, so it needs the truth, not best effort.
    /// </summary>
    private bool SaveWorkbookOf(string display)
    {
        try
        {
            _hostApp ??= HostApplication.Find();
            if (_hostApp is null)
            {
                return false;
            }

            using var books = _hostApp.GetObject("Workbooks");
            var count = books?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var book = books!.GetItem(i);
                if (book is not null
                    && string.Equals(book.GetString("Name"), display, StringComparison.OrdinalIgnoreCase))
                {
                    book.Invoke("Save");
                    Log.Info($"close: saved {display}");
                    return true;
                }
            }

            Log.Warn($"close: {display} is not among the application's workbooks");
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
            // but a close is the one moment staleness is guaranteed visible — the tab must
            // leave the strip — so the next few polls re-derive the picture unconditionally
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

        try
        {
            // The workbook the menu was opened on, when it was; the active project otherwise.
            using var project = FindProjectByDisplayName(projectName) ?? _editor.GetObject("ActiveVBProject");
            using var components = project?.GetObject("VBComponents");
            using var added = components?.CallObject("Add", kind);
            var name = added?.GetString("Name");

            Log.Info($"project: inserted {name ?? "?"} (kind {kind})");

            if (name is not null)
            {
                ShowModule(name);
            }

            // Not just the analyzer: the strip and the tree both have to hear about it too, and
            // this used to tell only the analyzer — so a module inserted from the page appeared
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
    /// Shows a module: the editor's active pane is set to it, and the surface is told to show it.
    ///
    /// Both are told directly. Show creates and displays a pane, but it does not reliably
    /// activate one that is already open behind another, and an activation that does not happen
    /// produces no window event and therefore, before this, no switch: clicking a tab did
    /// nothing exactly when both modules were already open. The active pane matters because the
    /// run and debug commands act on it; the surface matters because it is what the developer
    /// sees; neither is left to depend on the editor choosing to move a window.
    /// </summary>
    private void ShowModule(string component, string? projectDisplay = null)
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
                return;
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
        }
        catch (Exception ex)
        {
            Log.Error($"modules: {component} could not be shown", ex);
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
    /// The component carrying this name — within one project when its identity is given, in
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
                if (candidate?.GetString("Name") == component)
                {
                    foundProject = identity;
                    return candidate;
                }

                candidate?.Dispose();
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

        using var project = FindProjectByDisplayName(display);
        return project is null ? null : ProjectReader.Identity(project).Id;
    }

    /// <summary>
    /// Sends each open document the squiggles that belong to it.
    ///
    /// Findings arrive for whole projects and the surface holds one model per open module
    /// (decision 12), so they are filtered per document. A document with none is sent an empty
    /// set rather than skipped: that is what clears squiggles the user has just fixed.
    /// </summary>
    private void PublishMarkersToSurface()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        foreach (var (module, project) in surface.OpenDocuments)
        {
            var markers = _findings
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
    /// window is replaced, ghosted, floated, or policed away — because the
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
    /// replacement is genuinely standing. Native tool windows stopped saying no long ago —
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

#if DEBUG
        var startedAt = Environment.TickCount64;
#endif

        _editorSurface.Follow(SurfaceBounds(_frame, _documentArea, CanCoverChrome()), visible: true);

        // The full pass — window policing, band silencing, chrome — is object-model work,
        // and it was running once per frame event of a drag: measurable latency and repaint
        // churn (2026-08-05). The bounds followed above; everything else holds still until
        // the events pause, and then ONE full pass re-derives it all.
        _editorSurface.ArmPlacementSettle(PlacementSettleMilliseconds);

#if DEBUG
        PerfCounters.PlacementFast(Environment.TickCount64 - startedAt);
#endif
    }

    /// <summary>One quiet moment after the last frame event; then the full pass, once.</summary>
    private const uint PlacementSettleMilliseconds = 150;

    /// <summary>What the last followed pane event amounted to, geometry aside. An event that
    /// amounts to the same takes the fast path instead of the full cascade.</summary>
    private string? _lastFollowSubstance;

    /// <summary>
    /// A frame or document-area event: the frame moved, resized, was shown, or is going away.
    /// Bounds follow synchronously — that is what keeps the surface glued to the window — and
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
#if DEBUG
        var placementStarted = Environment.TickCount64;
#endif
        if (_editorSurface is null || !_surfaceShown || _frame == 0 || _documentArea == 0)
        {
            Log.Verbose($"placement: skipped (surface {(_editorSurface is null ? "none" : "up")}, " +
                        $"shown {_surfaceShown}, frame {_frame:X}, documents {_documentArea:X})");
            return;
        }

        // The frame hiding is the editor's window closing: Alt+F4, its X, a shutdown's first
        // act. Placement work there is worse than pointless — reacting to the hide event with
        // the cutout pass put object-model calls INSIDE the editor's own close handling, and
        // the editor faulted under them, taking the host down (three crash records,
        // 2026-08-04: VBE7, ntdll, and this shim faulting by turn, each at a close; the
        // object-model pass of that day is the police pass of this one). A hidden
        // frame needs nothing covered; the surface hides with it, and the show event that
        // brings the frame back re-derives everything.
        if (!Win32.IsWindowVisible(_frame))
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

#if DEBUG
        // Only a pass that did the full object-model work counts; the early exits are free.
        PerfCounters.PlacementFull(Environment.TickCount64 - placementStarted);
#endif
    }

    /// <summary>
    /// Silences or restores the windows the native menu bar and toolbars live in, by region
    /// rather than by visibility.
    ///
    /// Covering them was not enough: the bands paint without clipping their siblings, so
    /// every resize stamped the native menu bar straight over the surface — "the native menu
    /// bar bleeds thru", 2026-08-04. Hiding their windows was not enough either: the editor's
    /// own layout shows them again on every resize, and the beat between its show and our
    /// next hide is the same flash. An EMPTY window region ends the argument: the window
    /// stays exactly as visible as the editor believes, its layout never changes, and it owns
    /// no pixels to paint with — Office never sets or resets regions on these, so nothing
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
    /// our canvas to be purely xlide") — holes meant tracking native windows pixel-for-pixel
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

            // The frame resizing is not a pane event. With no visible pane — the empty
            // workspace — nothing else hears it, and the surface sat at its old size while the
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
            // show (it only holds panes it can match — the active one, in practice). A moment of
            // polls re-reads the object model's open list and republishes; the page skips the
            // rebuild when nothing changed, so a dying tooltip costs a diff and no work.
            _codePanes.WindowDestroyed = () =>
            {
                _resyncPanePolls = Math.Max(_resyncPanePolls, 2);
                _pollsRemaining = Math.Max(_pollsRemaining, (int)(1_000 / DebugPollMilliseconds));
                UpdatePolling();
            };

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

#if DEBUG
        // First out: no debug request may land on a session mid-teardown.
        _debugServer?.Dispose();
        _debugServer = null;
#endif

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
        // that stops — shutdown, disconnection, revival teardown — must leave the editor
        // whole, not silently menu-less.
        SetNativeChromeBands(visible: true);

        _immediateReader?.Dispose();
        _immediateReader = null;

        // The reading thread stops before the palettes change back: a reader must not touch a
        // window mid-restoration. Its join is bounded — see GhostReaderThread.Dispose.
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
