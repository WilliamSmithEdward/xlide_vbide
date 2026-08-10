using System.Text.Json;
using System.Runtime.InteropServices;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.Sync;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The debug api: the session's side of the local HTTP door.
///
/// SPLIT OFF AS A PARTIAL RATHER THAN REWRITTEN. Nothing here changed in the move. It lived in
/// AddInSession.cs, where two of its methods - the route switch and the on-host route switch - ran
/// to 1,407 and 1,019 lines and were, between them, 22% of a 10,784-line file that was itself 21%
/// of the product. The third longest member in that file is 223 lines, so the bulk was not spread
/// out: it was here, and it is one concern.
///
/// It is all inside `#if DEBUG`, which is what makes the split so clean: the region was already
/// contiguous, and a Release build has none of it. The gate proves that separately by looking for
/// the door's own strings in the published Release binary.
/// </summary>
internal sealed partial class AddInSession
{
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
                        ComWrappersDisposed: Com.ComRuntime.WrappersDisposed,
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
                        FrameCaption: HostChrome.CaptionOf(_frame),
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

            case "sync":
                return HandleSync(request.Query, request.Body);

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
                            // 1 standard, 2 class, 3 form, the VBE's own numbering, and the
                            // words for them as well.
                            //
                            // The words matter because this only parsed an int, and anything
                            // else fell through to the default: "kind=class" handed back a
                            // STANDARD module and still answered ok, so a caller asking for a
                            // class got one that could not hold a Friend member, and the
                            // analyzer was right to complain about it. Nothing said no. An
                            // unparseable kind is now refused rather than guessed at.
                            request.Query.TryGetValue("kind", out var kindText);
                            var kind = (kindText ?? string.Empty).Trim().ToLowerInvariant() switch
                            {
                                "" => 1,
                                "1" or "module" or "standard" => 1,
                                "2" or "class" => 2,
                                "3" or "form" or "userform" => 3,
                                _ => 0,
                            };

                            if (kind == 0)
                            {
                                return System.Text.Json.JsonSerializer.Serialize(
                                    new DebugErrorReply(
                                        $"kind '{kindText}' is not one of 1/module/standard, "
                                        + "2/class, 3/form"),
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
                        // And its answer is the show's answer: opening a module that is not there
                        // replied ok, which is the same lie the write route told about a module
                        // that is not there (2026-08-09).
                        var showed = ShowModule(paneModule, DisplayFromProjectId(paneOwner));
                        return showed is null
                            ? System.Text.Json.JsonSerializer.Serialize(
                                new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply)
                            : System.Text.Json.JsonSerializer.Serialize(
                                new DebugErrorReply(showed), DebugJsonContext.Default.DebugErrorReply);

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
                        SyncEngine = request.Query.TryGetValue("syncEngine", out var planner)
                            ? planner
                            : settings.SyncEngine,
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
                        settings.SyncEngine),
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
                    //
                    // AND ITS ANSWER IS THE WRITER'S ANSWER. This replied ok unconditionally, so a
                    // write the editor refused - which can cost the module its previous text - was
                    // indistinguishable here from one that landed. The fixture builder works around
                    // it by reading the line count back, which is the right instinct and should not
                    // have been necessary (2026-08-09).
                    var complaint = WriteModule(moduleName, request.Body, projectId, hostRewrite: true);
                    return complaint is null
                        ? System.Text.Json.JsonSerializer.Serialize(
                            new DebugCommandReply(true, 0), DebugJsonContext.Default.DebugCommandReply)
                        : System.Text.Json.JsonSerializer.Serialize(
                            new DebugErrorReply(complaint), DebugJsonContext.Default.DebugErrorReply);
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
}
