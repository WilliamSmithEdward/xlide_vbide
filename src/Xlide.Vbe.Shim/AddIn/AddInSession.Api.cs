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
/// The xlide api: the session's side of the local HTTP door.
///
/// SPLIT OFF AS A PARTIAL RATHER THAN REWRITTEN. Nothing here changed in the move. It lived in
/// AddInSession.cs, where two of its methods - the route switch and the on-host route switch - ran
/// to 1,407 and 1,019 lines and were, between them, 22% of a 10,784-line file that was itself 21%
/// of the product. The third longest member in that file is 223 lines, so the bulk was not spread
/// out: it was here, and it is one concern.
///
/// It WAS all inside `#if DEBUG`, which is what made the split so clean: the region was already
/// contiguous. The gate is a RUNTIME one now (the owner, 2026-08-22: "let's just have one api...
/// but we just include what we call xlide api in the production build... but it's off by default
/// in production, so we're not duplicating code"). A shipped build carries every route here and
/// listens on none of them until someone turns the door on from the agent card.
/// </summary>
internal sealed partial class AddInSession
{
    private ApiServer? _apiServer;

    /// <summary>
    /// The host UI thread, captured at construction (OnConnection runs there). Two doors need
    /// it: CrossToHost goes inline instead of deadlocking when a request is ALREADY on the host
    /// thread - which is every request through the inside door - and RunPageScriptOnce refuses
    /// there, because a page script's answer arrives by the very pump that thread would be
    /// blocking to wait.
    /// </summary>
    private readonly int _hostThreadId = Environment.CurrentManagedThreadId;

    /// <summary>The in-process door, held so its COM wrapper outlives every caller.</summary>
    private InsideDoor? _insideDoor;
    private Com.DispatchObject? _insideDoorRef;

    private static ApiServer.ApiReply ApiError(string error) =>
        ApiServer.ApiReply.Json(HostError(error));

    /// <summary>
    /// The error reply as the ON-HOST switch answers it: the serialized string, because
    /// AnswerDebugRequestOnHost returns strings and the dispatch wraps them. This existing
    /// as a name is what stops the serializer call being spelled out at every refusal -
    /// it was spelled out twenty-eight times (the audit's B13). ApiError above is the
    /// pool-side wrapping of the same convention.
    /// </summary>
    private static string HostError(string error) =>
        System.Text.Json.JsonSerializer.Serialize(
            new DebugErrorReply(error), DebugJsonContext.Default.DebugErrorReply);

    /// <summary>The bare "it ran" the on-host action routes answer when the act itself is the
    /// whole story. Anything with an outcome to report builds its reply by hand.</summary>
    private static string HostOk(string detail = "") =>
        System.Text.Json.JsonSerializer.Serialize(
            new DebugCommandReply(true, 0, detail), DebugJsonContext.Default.DebugCommandReply);

    /// <summary>
    /// The root the agent routes build their advertised URLs on. Null only if the HTTP door
    /// failed to start while the in-process door still serves - a caller there gets URLs that
    /// say so rather than URLs that dangle.
    /// </summary>
    private string AgentBaseUrl() => _apiServer?.BaseUrl ?? "http://(the-http-door-did-not-start)";

    /// <summary>
    /// The designerSaveDirty act's refusal, or null when every dirty document applied. The act
    /// answers did:false with its refusals in detail; anything unparseable is read as clean,
    /// because the flush must never turn a working export into a refusal about JSON.
    /// </summary>
    private static string? FlushRefusal(string actAnswer)
    {
        try
        {
            using var parsed = System.Text.Json.JsonDocument.Parse(actAnswer);
            var root = parsed.RootElement;
            if (root.ValueKind == System.Text.Json.JsonValueKind.String)
            {
                using var inner = System.Text.Json.JsonDocument.Parse(root.GetString() ?? "{}");
                return RefusalOf(inner.RootElement);
            }

            return RefusalOf(root);

            static string? RefusalOf(System.Text.Json.JsonElement act) =>
                act.TryGetProperty("did", out var did) && did.ValueKind == System.Text.Json.JsonValueKind.False
                    && act.TryGetProperty("detail", out var detail)
                    ? detail.GetString() ?? "a designer document refused to apply"
                    : null;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>When the running shim was built, the same reading doctor reports.</summary>
    private static string ShimBuiltUtc()
    {
        var directory = Interop.ShimModule.Directory;
        var shim = directory is null ? null : Path.Combine(directory, "Xlide.Vbe.Shim.dll");
        return shim is not null && File.Exists(shim)
            ? File.GetLastWriteTimeUtc(shim).ToString("O")
            : "unknown";
    }

    /// <summary>
    /// Puts the in-process door where running code can pick it up: the Running Object Table,
    /// under its own ProgID, so `GetObject(, "Xlide.Api")` answers from VBA and
    /// GetActiveObject("Xlide.Api") from any automation client.
    ///
    /// THE ROT IS THE DOOR'S ADDRESS BECAUSE THE OBVIOUS ADDRESS DOES NOT EXIST. This add-in is
    /// a VBE add-in, and `Application.COMAddIns` holds only OFFICE add-ins - measured 2026-08-18:
    /// PowerMap, DataStreamer, PowerPivot, and no VBE add-in at all. The VBE's own collection
    /// (`Application.VBE.AddIns(...).Object`) is reachable only through Application.VBE, which
    /// the trust switch gates - so a door hung there could never be found by the very callers it
    /// is for. The ROT is how Excel itself is found (`GetObject(, "Excel.Application")`), needs
    /// no traversal, and inherits that mechanism's one ambiguity: with several instances
    /// running, GetObject binds ONE of them - the agent reply carries `pid`, which is how a
    /// caller checks it got the session it meant.
    ///
    /// The add-in's own `Object` property is still tried, best effort, for callers that do
    /// hold the VBE - and this VBE refuses it BOTH ways (reference put and ordinary put,
    /// measured 2026-08-18), so today the attempt only writes a log line. It stays because it
    /// costs one call, the log records the truth, and a host that ever accepts it gains a
    /// second address for free. The ROT is the door's real address either way.
    /// </summary>
    private void OfferInsideDoor()
    {
        try
        {
            var door = new InsideDoor(AnswerInsideRequest);
            var unknown = Com.ComRuntime.Wrappers.GetOrCreateComInterfaceForObject(
                door, System.Runtime.InteropServices.CreateComInterfaceFlags.None);

            var wrapper = Com.DispatchObject.Attach(unknown);
            if (wrapper is null)
            {
                Marshal.Release(unknown);
                Log.Info("inside door: our own object did not answer as dispatch; not offered");
                return;
            }

            _insideDoor = door;
            _insideDoorRef = wrapper;

            // The ProgID -> CLSID mapping GetObject resolves through. HKCU, rewritten by every
            // session and left in place at retirement - see RetireInsideDoor for why deleting
            // it broke the other Excel.
            using (var classes = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(
                $@"Software\Classes\{Core.ProductIdentity.ApiProgId}\CLSID"))
            {
                classes.SetValue(null, $"{{{Core.ProductIdentity.ApiClsid}}}");
            }

            var classId = new Guid(Core.ProductIdentity.ApiClsid);
            var registered = RegisterActiveObject(wrapper.Pointer, in classId, 0, out _insideDoorRotTicket);
            if (registered == HResult.Ok)
            {
                Log.Info($"inside door: GetObject(, \"{Core.ProductIdentity.ApiProgId}\") answers");
            }
            else
            {
                _insideDoorRotTicket = 0;
                Log.Info($"inside door: the running object table refused (0x{registered:X8})");
            }

            // Best effort, and genuinely secondary: only trust-holding callers can traverse to it.
            if (_addIn is not null)
            {
                try
                {
                    _addIn.SetObjectByValue("Object", wrapper);
                    Log.Info("inside door: the add-in's Object property answers too");
                }
                catch (Exception ex)
                {
                    Log.Info($"inside door: the add-in refused its Object property ({ex.GetType().Name})");
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"inside door: could not be offered ({ex.GetType().Name}: {ex.Message})");
        }
    }

    /// <summary>
    /// Takes the door back down: out of the ROT and off the add-in. The ProgID -> CLSID key in
    /// HKCU\Software\Classes deliberately STAYS: with two Excels running, the first session to
    /// stop would otherwise delete the mapping out from under the survivor, and every
    /// `GetObject(, "Xlide.Api")` on the machine would fail until some later session start
    /// rewrote it (found in the 2026-08-19 hunt). A mapping with no running instance behind it
    /// refuses GetObject exactly the way no mapping does, so leaving it costs nothing.
    /// </summary>
    private void RetireInsideDoor()
    {
        if (_insideDoorRotTicket != 0)
        {
            _ = RevokeActiveObject(_insideDoorRotTicket, 0);
            _insideDoorRotTicket = 0;
        }

        if (_insideDoorRef is not null && _addIn is not null)
        {
            try
            {
                _addIn.ClearObject("Object");
            }
            catch (Exception)
            {
                // The put may never have landed; a refused clear says nothing new.
            }
        }

        _insideDoorRef?.Dispose();
        _insideDoorRef = null;
        _insideDoor = null;
    }

    private uint _insideDoorRotTicket;

    [System.Runtime.InteropServices.LibraryImport("oleaut32.dll")]
    private static partial int RegisterActiveObject(
        nint unknown, in Guid classId, uint flags, out uint ticket);

    [System.Runtime.InteropServices.LibraryImport("oleaut32.dll")]
    private static partial int RevokeActiveObject(uint ticket, nint reserved);

    /// <summary>
    /// One inside-door request: policy first, then the same switch the HTTP door serves.
    ///
    /// The one policy is THREADING, not trust: page routes refuse fast because they could only
    /// ever time out here (see RunPageScriptOnce). The door deliberately does not gate on the
    /// host's "Trust access to the VBA project object model" switch - the owner's call,
    /// 2026-08-18: the end-user experience must not require a Trust Center trip, and the gate
    /// protected nothing while the HTTP door beside it answers the same routes for any local
    /// caller without ever reading that switch.
    /// </summary>
    private string AnswerInsideRequest(string target, string body)
    {
        /*
         * ONE NAME REACHES EVERY SESSION (the owner, 2026-08-19: "i don't like the rot door
         * limitation"). GetObject binds one registration however many hosts are live - the
         * first, then the survivor - so whoever holds the name FEDERATES: a target opening
         * `@who/` names another session, resolved against the discovery files and proxied to
         * that session's own HTTP door with its own key. `@word/state` from Excel's door
         * answers Word's state; `@12345/agent` addresses a pid exactly. Request("sessions")
         * lists the fleet. A self-address just serves locally, and a proxied route escapes
         * this door's host-thread limits for free - the peer answers on its own pool.
         */
        var trimmed = target.Trim().TrimStart('/');
        if (trimmed.StartsWith('@'))
        {
            var slash = trimmed.IndexOf('/', StringComparison.Ordinal);
            if (slash <= 1)
            {
                return HostError(
                    "@ addresses another session and takes its route after the slash: "
                    + "Request(\"@word/state\"), Request(\"@12345/agent\"). Request(\"sessions\") lists them.");
            }

            return AnswerAddressedRequest(trimmed[1..slash], trimmed[(slash + 1)..], body);
        }

        // The route and its query, the same split the HTTP door reads from a URL - kept tiny
        // and local because the HTTP parse is entangled with the token prefix this door has no
        // need of.
        var route = target;
        var query = new Dictionary<string, string>(StringComparer.Ordinal);
        var mark = target.IndexOf('?', StringComparison.Ordinal);
        if (mark >= 0)
        {
            route = target[..mark];
            foreach (var pair in target[(mark + 1)..].Split('&'))
            {
                var eq = pair.IndexOf('=', StringComparison.Ordinal);
                if (eq > 0)
                {
                    query[Uri.UnescapeDataString(pair[..eq])] = Uri.UnescapeDataString(pair[(eq + 1)..]);
                }
            }
        }

        route = route.Trim().TrimStart('/');

        if (route.Length == 0)
        {
            return HostError(
                "Request takes a route, e.g. Request(\"agent\") - Request(\"agent/routes\") "
                + "lists them all, and Request(\"@word/state\") reaches another live session "
                + "(Request(\"sessions\") lists those)");
        }

        if (AgentGuide.PolicyOf(route) == AgentGuide.DoorPolicy.HttpOnly)
        {
            return HostError(
                $"'{route}' answers only at the HTTP door ({AgentBaseUrl()}): it waits on "
                + "work only a pumping host thread can finish, and this call is standing on "
                + "that thread. Request(\"agent/routes\") says which door each route answers.");
        }

        // A DESIGNATED DEVIATION from the HTTP door: assert defaults to timeoutMs=0 here. Its
        // wait polls HOST state, and an inside caller is standing on the host thread - so from
        // this door a false claim can never become true by waiting; it can only hold Excel
        // frozen for the whole default ten seconds and then answer what one look would have
        // answered at once. A caller that names a timeout on purpose keeps it, and burning it
        // is then a choice. timeoutMs=0 evaluates the claim exactly once.
        if (route == "assert" && !query.ContainsKey("timeoutMs"))
        {
            query["timeoutMs"] = "0";
        }

        var reply = AnswerApiRequest(new ApiServer.ApiRequest(route, query, body));

        return reply.ContentType.StartsWith("application/json", StringComparison.Ordinal)
            ? System.Text.Encoding.UTF8.GetString(reply.Bytes)
            : HostError($"'{route}' answers {reply.ContentType}, which only the HTTP door can carry");
    }

    /// <summary>
    /// Every live session's discovery facts, self included: pid and host for naming, port and
    /// token for the `@` proxy's own use. Dead pids are skipped the way the discovery sweep
    /// buries them; an unreadable file is a session mid-write and is skipped the same.
    /// </summary>
    private static List<(int Pid, string Host, string StartedAt, int Port, string Token)> LiveSessions()
    {
        var fleet = new List<(int, string, string, int, string)>();
        var directory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            Core.ProductIdentity.DataFolderName);

        foreach (var file in Directory.EnumerateFiles(directory, "xlide-api-*.json"))
        {
            try
            {
                using var parsed = JsonDocument.Parse(File.ReadAllText(file));
                var root = parsed.RootElement;
                var pid = root.GetProperty("pid").GetInt32();
                if (pid != Environment.ProcessId)
                {
                    using var alive = System.Diagnostics.Process.GetProcessById(pid);
                }

                fleet.Add((
                    pid,
                    root.TryGetProperty("host", out var host) ? host.GetString() ?? "excel" : "excel",
                    root.TryGetProperty("startedAt", out var started) ? started.GetString() ?? "" : "",
                    root.GetProperty("port").GetInt32(),
                    root.GetProperty("token").GetString() ?? ""));
            }
            catch
            {
                // A corpse or a file mid-write; the fleet is what answers, not what lingers.
            }
        }

        fleet.Sort((left, right) => left.Item1.CompareTo(right.Item1));
        return fleet;
    }

    private static readonly System.Net.Http.HttpClient PeerDoor = new()
    {
        Timeout = TimeSpan.FromSeconds(4),
    };

    /// <summary>
    /// The `@` proxy: resolves `who` (a pid, or a host name like word) against the live fleet
    /// and forwards the request to that session's HTTP door. Self-addresses serve locally.
    /// The peer's token travels in the URL it already lives in and is never echoed to the
    /// caller; a body makes the forward a POST, its absence a GET, which is the split every
    /// route table row documents.
    /// </summary>
    private string AnswerAddressedRequest(string who, string target, string body)
    {
        var fleet = LiveSessions();
        List<(int Pid, string Host, string StartedAt, int Port, string Token)> matches;
        if (int.TryParse(who, out var wantedPid))
        {
            matches = fleet.FindAll(one => one.Pid == wantedPid);
        }
        else
        {
            matches = fleet.FindAll(one => string.Equals(one.Host, who, StringComparison.OrdinalIgnoreCase));
        }

        if (matches.Count == 0)
        {
            var known = fleet.Count == 0
                ? "none are live"
                : string.Join(", ", fleet.ConvertAll(one => $"{one.Host} (pid {one.Pid})"));
            return HostError($"@{who}: no live session answers to that; {known}");
        }

        if (matches.Count > 1)
        {
            var pids = string.Join(", ", matches.ConvertAll(one => one.Pid.ToString()));
            return HostError($"@{who} is {matches.Count} sessions (pids {pids}); address one by pid");
        }

        // ONE HOP, designated. A nested address would recurse here once per prefix on a
        // self-address - a caller stacking thousands of `@pid/` prefixes would overflow the
        // stack ON THE HOST THREAD and take the host with it - and on a peer it would arrive
        // at an HTTP door that has never heard of `@` and answer a riddle. A hop that lands
        // is a session; sessions do not forward.
        if (target.TrimStart('/').StartsWith('@'))
        {
            return HostError(
                $"@{who}: the address takes ONE hop - name the final session directly "
                + "(Request(\"sessions\") lists them all)");
        }

        var found = matches[0];
        if (found.Pid == Environment.ProcessId)
        {
            return AnswerInsideRequest(target, body);
        }

        try
        {
            var url = $"http://127.0.0.1:{found.Port}/{found.Token}/{target}";
            using var request = new System.Net.Http.HttpRequestMessage(
                body.Length > 0 ? System.Net.Http.HttpMethod.Post : System.Net.Http.HttpMethod.Get, url);
            if (body.Length > 0)
            {
                request.Content = new System.Net.Http.StringContent(
                    body, System.Text.Encoding.UTF8, "application/json");
            }

            using var response = PeerDoor.Send(request);
            var contentType = response.Content.Headers.ContentType?.MediaType ?? "";
            using var reader = new StreamReader(response.Content.ReadAsStream(), System.Text.Encoding.UTF8);
            var answer = reader.ReadToEnd();
            return contentType.StartsWith("application/json", StringComparison.Ordinal)
                ? answer
                : HostError($"@{who}: '{target}' answers {contentType}, which only the HTTP door can carry");
        }
        catch (Exception ex)
        {
            return HostError(
                $"@{who} (pid {found.Pid}) did not answer: {ex.GetType().Name}. "
                + "The session may be mid-teardown; Request(\"sessions\") re-lists the fleet.");
        }
    }

    /// <summary>Shapes a page script's answer the one way every eval-style route answers it:
    /// the script's error verbatim, or the answer with its result unwrapped for the caller.</summary>
    private static ApiServer.ApiReply PageReply(
        (bool Answered, int ErrorCode, string Result, string? Error) ran) =>
        ran.Error is { } error
            ? ApiError(error)
            : ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                new DebugEvalReply(ran.Answered, ran.ErrorCode, ran.Result, Unwrap(ran.Result)),
                DebugJsonContext.Default.DebugEvalReply));

    /// <summary>
    /// Polls until the page's bridge answers, for the routes that just tore the page down and
    /// must not report on the page that is going away. Ready is the page's own word for it: a
    /// page part way through booting can run script and still have no bridge.
    /// </summary>
    private bool WaitForWorkspace(long budgetMs)
    {
        var began = Environment.TickCount64;
        while (Environment.TickCount64 - began < budgetMs)
        {
            Thread.Sleep(150);
            var probe = RunPageScript("!!(window.xlideBridge && window.xlideBridge.workspace)", null, 1500);
            if (probe.Error is null && probe.Result.Trim() == "true")
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Min, median, p95 and max over the samples - the one quantile convention, so the
    /// bench and trip routes cannot come to mean two different things by p95.</summary>
    private static ApiServer.ApiReply BenchReply(string what, List<double> samples, string detail)
    {
        var ordered = samples.OrderBy(one => one).ToArray();
        return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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

    /// <summary>
    /// One crossing to the host thread: queues the work, waits the standard three seconds, and
    /// samples the crossing for the stats route. Every marshaled request doubles as a probe of
    /// host responsiveness - INCLUDING the journal's nested state read, whose hand-rolled copy
    /// of this scaffold was the one crossing the perf route could not see (the audit's B13).
    /// The caller owns disposal, because the dispatch's blocked path keeps waiting on Done
    /// after dismissing the dialog that owned the thread.
    /// </summary>
    private readonly record struct HostCrossing(bool Answered, ManualResetEventSlim Done) : IDisposable
    {
        public void Dispose() => Done.Dispose();
    }

    /// <summary>The route whose marshaled work is ON the host thread right now - the lane's
    /// current holder - and the tick it took the lane. Written from the host thread, read by
    /// `stats` on a pool thread, hence the volatile and the interlocked ticks.</summary>
    private volatile string? _laneHolder;
    private long _laneHeldSince;

    private HostCrossing CrossToHost(EditorSurface host, Action work)
    {
        // Already on the host thread - a request through the inside door. Queueing here would
        // be a deadlock wearing a timeout's clothes: the queued work waits for the thread that
        // is standing right here waiting for the queue. Run it now. Not sampled as a marshal,
        // because no marshal happened and a flood of zero-cost samples would flatter the stats.
        if (Environment.CurrentManagedThreadId == _hostThreadId)
        {
            var inline = new ManualResetEventSlim(true);
            try
            {
                work();
            }
            catch
            {
                inline.Dispose();
                throw;
            }

            return new HostCrossing(true, inline);
        }

        var done = new ManualResetEventSlim(false);
        var began = Environment.TickCount64;
        try
        {
            host.RunOnHostThread(() =>
            {
                try
                {
                    work();
                }
                finally
                {
                    done.Set();
                }
            });
        }
        catch
        {
            done.Dispose();
            throw;
        }

        var answered = done.Wait(TimeSpan.FromSeconds(3));
        PerfCounters.Marshal(Environment.TickCount64 - began);
        return new HostCrossing(answered, done);
    }

    /// <summary>
    /// A boolean query argument, read one way.
    ///
    /// THERE WERE TWO RULES IN THIS FILE and they disagreed about the same text. Six sites took
    /// anything that was not the literal "0" as true, so `reset=false` cleared the counters and
    /// `text=no` asked for the text. Two sites wanted an affirmative word, so `live=1` worked and
    /// `live=yes` did not, depending which route you were calling. Both readings are defensible
    /// and a caller cannot hold both, which is what makes it a defect rather than a preference.
    ///
    /// Absent, or a word neither list knows, answers the fallback: a route's default is the
    /// route's business, and a typo must not silently mean the opposite of what it looks like.
    /// </summary>
    private static bool Flag(ApiServer.ApiRequest request, string name, bool fallback = false)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!request.Query.TryGetValue(name, out var value))
        {
            return fallback;
        }

        return value.ToLowerInvariant() switch
        {
            "1" or "true" or "yes" or "on" => true,
            "0" or "false" or "no" or "off" => false,

            // Present and unreadable. "reset=maybe" is a caller who believes they asked for
            // something, and the honest answer is the route's default rather than a guess.
            _ => fallback,
        };
    }

    /// <summary>
    /// The project a request named, or null when it named none.
    ///
    /// A NAMED PROJECT THAT DOES NOT RESOLVE IS AN ERROR, and this is the whole point of the
    /// method. Every route used to write `ProjectIdFromDisplay(asked) ?? _shownProject` or
    /// `FindProjectByDisplayName(asked) ?? ActiveVBProject`, so a value that matched nothing
    /// silently became a different workbook - and with two workbooks each holding a Helpers, the
    /// caller got the wrong module and no indication of it. Naming a workbook and being answered
    /// about another is worse than being refused, because a refusal can be read.
    ///
    /// An ABSENT argument is not an error: "the one on screen" is a reasonable default and most
    /// callers have exactly one workbook open. Only a name that was given and did not land.
    /// </summary>
    private string? ResolveNamedProject(string? asked, out string? complaint)
    {
        complaint = null;

        if (string.IsNullOrEmpty(asked))
        {
            return null;
        }

        if (ProjectIdFromDisplay(asked) is { } resolved)
        {
            return resolved;
        }

        var open = _projectNames.Values.Where(name => name.Length > 0).Distinct().ToArray();
        complaint = $"no open workbook answers to '{asked}'"
            + (open.Length > 0 ? $". Open: {string.Join(", ", open)}" : "; none are open")
            + ". A display name, a full path or a project identity all resolve.";

        return null;
    }

    /// <summary>A request's wait budget, clamped to something a stuck page cannot outlast.</summary>
    private static int WaitMilliseconds(ApiServer.ApiRequest request, int fallback)
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
    /// answer by calling back on that same thread, so THIS thread - a pool thread - is the
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
        // polls, which is the whole point - this thread is not holding anything it needs.
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
         * door returns - the initial call, the sleep, and the poll, each of the last two a page
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
                return (false, HostError(error.GetString() ?? "the script failed"));
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

        // From the host thread itself this can only ever time out: ExecuteScript's answer is
        // delivered by the pump, and the pump is what this wait would be blocking. That was
        // already true for doctor's page probe when it runs on-host - it burned its three-second
        // scheduling wait and reported nothing - and it is true for every inside-door call.
        // Saying so at once beats proving it slowly.
        if (Environment.CurrentManagedThreadId == _hostThreadId)
        {
            return (false, 0, string.Empty,
                "a page script cannot be awaited from the host thread itself: its answer arrives "
                + "by the pump this wait would block. Page routes answer at the HTTP door.");
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
    /// ASKS the page what it recorded going wrong. The fallback behind the pushed report.
    ///
    /// The FIRST rather than the last: a bundle that dies on load throws once and then produces a
    /// cascade of consequences, and the consequences are the part that is easy to find.
    ///
    /// This can be asked of a page that never booted, which is the whole reason it exists - a
    /// module throwing during initialisation leaves the JavaScript context perfectly alive, so the
    /// ring boot.js installed before it is still there to be read.
    /// </summary>
    private (string? Error, bool Asked) AskPageForItsError()
    {
        /*
         * A SECOND, well inside the deadline of the route that calls it.
         *
         * The first version asked for three, which is the whole budget a host-thread route gets,
         * so doctor stopped answering at all in the one state this was added for: the page dead,
         * the finding ready, and the route timing out before it could say so. Measured immediately
         * after adding it, against a bundle broken on purpose (2026-08-09). A diagnostic that
         * spends the caller's entire budget is not a diagnostic.
         */
        var read = RunPageScript(
            """
            (function () {
              var ring = window.__xlideConsole;
              if (!ring) { return null; }
              for (var i = 0; i < ring.length; i++) {
                if (ring[i].indexOf("UNCAUGHT") === 0 || ring[i].indexOf("UNHANDLED REJECTION") === 0) {
                  return ring[i];
                }
              }
              return null;
            })()
            """,
            null,
            1000);

        // "The page recorded nothing" and "the page could not be asked" are opposite answers, and
        // reporting the second as the first would say a broken page is probably still starting.
        if (!read.Answered || read.Error is not null)
        {
            return (null, false);
        }

        if (string.IsNullOrWhiteSpace(read.Result) || read.Result == "null")
        {
            return (null, true);
        }

        // The page answers as JSON, so a string comes back quoted. Unwrapped with JsonDocument
        // rather than a generic Deserialize: this library is published ahead-of-time and the
        // reflecting overloads are refused outright there.
        var line = read.Result.Trim();
        if (line.StartsWith('"'))
        {
            try
            {
                using var parsed = System.Text.Json.JsonDocument.Parse(line);
                line = parsed.RootElement.GetString() ?? line;
            }
            catch (System.Text.Json.JsonException)
            {
                // Left as it came. A finding that says something odd beats no finding.
            }
        }

        // One line, trimmed: a stack can be a dozen frames and a finding is meant to be read.

        var firstLine = line.Split('\n')[0].Trim();
        return (firstLine.Length > 300 ? firstLine[..300] + "..." : firstLine, true);
    }

    /// <summary>
    /// A FALLBACK. The page installs this ring itself, in boot.js, ahead of its own bundle.
    ///
    /// This ran at page-ready and was the only installer, which made the ring useless for the
    /// failure it most needed to cover: a bundle that throws while its modules initialise never
    /// reaches ready, so the ring was never created and the `console` route answered
    /// `{"installed": false, "lines": []}` at exactly the moment somebody was asking why the page
    /// was blank (2026-08-09). Ownership moved into the page, where nothing can run before it.
    ///
    /// Kept because a page served from somewhere without boot.js - an older bundle, a hand-built
    /// dist - should still say something rather than nothing. It no-ops when boot.js has run,
    /// which it detects through the same flag boot.js sets.
    ///
    /// The console is WRAPPED rather than replaced: everything still reaches DevTools when a
    /// client is attached.
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
        // the host thread, and the script's answer arrives on that same thread - waiting
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
    /// Whether the page is running a bundle older than the one on disk - the question behind
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
    /// deadline. The immediate route runs the line on the host thread and waits for its
    /// result HERE, on the door's thread, so a Compile-error box it raises does not jam the
    /// wait; the one thing no deadline outlasts is a line that stops at a real breakpoint,
    /// which is the debugger working.
    /// </summary>
    private ApiServer.ApiReply AnswerApiRequest(ApiServer.ApiRequest request)
    {
        // Sweep FIRST, before the routes that answer without the host thread.
        //
        // The sweep used to sit below them, which read as "every request heals first" and was
        // not: dialogs, dismiss and guard all return before reaching it, and those are exactly
        // the routes a caller uses while something is standing. Armed and watching, the guard
        // therefore never ran once - fourteen seconds of polling with a modal on screen and an
        // empty cleared list (2026-08-07).
        //
        // The heartbeat is no help here and this is why it cannot be the trigger: a VBA modal
        // PUMPS messages, so the host thread kept answering in under 140ms the whole time it was
        // blocked. What is standing is the only evidence that something is standing.
        ClearDialogsWeRaised();

        // THE HOST-FREE ROUTES ANSWER THEIR OWN REFUSALS. assert and dismiss exist precisely so
        // a caller can get an answer while a modal owns the host thread - and their argument
        // guards sat on their switch cases, so a call with a missing argument fell out of the
        // switch and was marshalled to the very thread these routes exist to avoid: dismiss
        // with no button, made while a dialog stood, timed out instead of saying "button?".
        // The refusal happens here, before the switch, on the pool thread (the audit's B19).
        // guard needs no twin: its case carries no argument guard, so it always answers.
        if (request.Route == "assert"
            && !(request.Query.TryGetValue("that", out var assertClaim) && assertClaim.Length > 0))
        {
            return ApiError($"assert needs that=<claim>; known claims are {string.Join(", ", KnownClaims)}");
        }

        if (request.Route == "dismiss"
            && !(request.Query.TryGetValue("button", out var dismissButton) && dismissButton.Length > 0))
        {
            return ApiError("dismiss needs button=<label>, and takes caption=<title> to pick between dialogs");
        }

        // EXPORT FLUSHES THE DESIGNERS FIRST (the owner, 2026-08-19): the designer's document
        // is the transaction log and the form only catches up on a save, so an export that
        // skipped the save shipped the LAST save - the same bug Run had, fixed by the same
        // rule. Pool-side because applying needs the page to pump; the sync dialog's own path
        // makes the same call page-side before it sends. A page that is not up has no dirty
        // designers, so a failed script is a clean continue - and a document that REFUSES to
        // apply stops the export and says why, exactly as it stops a run.
        if (request.Route == "sync"
            && request.Query.TryGetValue("direction", out var syncFlushDirection)
            && string.Equals(syncFlushDirection, "export", StringComparison.OrdinalIgnoreCase))
        {
            var flushed = RunPageScript("window.xlideUi.act('designerSaveDirty', {})", null, 20000);
            if (flushed.Error is null && FlushRefusal(flushed.Result) is { } refusal)
            {
                return ApiError($"the export did not run: a designer document refuses to apply - {refusal}");
            }
        }

        switch (request.Route)
        {
            // The agent front door. Pool-side and host-free on purpose: orientation must answer
            // even while the host thread is busy or wedged, which is exactly when a caller that
            // has never seen this api is most likely to be asking what it is looking at.
            case "agent":
            {
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    AgentGuide.FrontDoor(
                        AgentBaseUrl(), Engine.HostApp.Name, Environment.ProcessId,
                        _analysis?.IsReady == true, ShimBuiltUtc()),
                    DebugJsonContext.Default.DebugAgentReply));
            }

            // The fleet, host-free like the front door: the discovery files name every live
            // session, and this is the list the inside door's `@` prefix addresses. Ports,
            // tokens and agent urls are deliberately absent from the reply - one door's caller
            // is not handed the keys to every other door; the `@` proxy uses them unseen.
            case "sessions":
            {
                var fleet = LiveSessions();
                var rows = new DebugSessionRow[fleet.Count];
                for (var at = 0; at < fleet.Count; at++)
                {
                    rows[at] = new DebugSessionRow(
                        fleet[at].Pid, fleet[at].Host, fleet[at].StartedAt,
                        fleet[at].Pid == Environment.ProcessId);
                }

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugSessionsReply(rows), DebugJsonContext.Default.DebugSessionsReply));
            }

            case "agent/routes":
            {
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    AgentGuide.RouteTable(AgentBaseUrl()),
                    DebugJsonContext.Default.DebugAgentRoutesReply));
            }

            case "agent/route" when request.Query.TryGetValue("name", out var helpName) && helpName.Length > 0:
            {
                if (AgentGuide.OneRoute(helpName) is not { } row)
                {
                    return ApiError(
                        $"no route named '{helpName}'; the names are "
                        + string.Join(", ", AgentGuide.Routes.Select(one => one.Name)));
                }

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    row, DebugJsonContext.Default.DebugAgentRouteRow));
            }

            case "agent/route":
                return ApiError("agent/route needs name=<route>; agent/routes lists them all");

            case "agent/examples":
            {
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    AgentGuide.Examples(AgentBaseUrl()),
                    DebugJsonContext.Default.DebugAgentExamplesReply));
            }

            // What the language service KNOWS, passed through from the engine verbatim: the
            // engine's reply is the contract, and reshaping it here would be a second copy that
            // drifts. Pool-side because the pipe answers on its own thread - no host needed.
            case "model":
            {
                if (_analysis?.Engine is not { } modelEngine)
                {
                    return ApiError("the analysis engine is not up");
                }

                request.Query.TryGetValue("type", out var modelType);

                try
                {
                    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                    var known = modelEngine.KnowledgeModelAsync(modelType, deadline.Token)
                        .GetAwaiter().GetResult();

                    return known is { } model
                        ? ApiServer.ApiReply.Json(model.GetRawText())
                        : ApiError("the engine did not answer the model request");
                }
                catch (Exception ex)
                {
                    return ApiError($"model failed: {ex.Message.Trim()}");
                }
            }

            case "analyzer":
            {
                if (_analysis?.Engine is not { } rulesEngine)
                {
                    return ApiError("the analysis engine is not up");
                }

                try
                {
                    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                    var known = rulesEngine.KnowledgeAnalyzerAsync(deadline.Token)
                        .GetAwaiter().GetResult();

                    return known is { } rules
                        ? ApiServer.ApiReply.Json(rules.GetRawText())
                        : ApiError("the engine did not answer the rules request");
                }
                catch (Exception ex)
                {
                    return ApiError($"analyzer failed: {ex.Message.Trim()}");
                }
            }

            case "log":
            {
                var path = Log.Path;
                if (path is null)
                {
                    return ApiError("no log file");
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
                        return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugMessagesReply(rows), DebugJsonContext.Default.DebugMessagesReply));
            }

            case "capture":
            {
                request.Query.TryGetValue("window", out var which);

                // `form` is a RUNNING form, matched on caption like the userform close - the one
                // picture of a designer's work nothing in the object model can answer for, since
                // MSForms draws windowless and the designer's collection describes the stored
                // form rather than the one on screen.
                request.Query.TryGetValue("caption", out var wantedCaption);
                var target = which switch
                {
                    "palette" => _browserPalette?.Handle ?? 0,
                    "form" => DialogWatch.RunningFormHandle(wantedCaption),
                    _ => _frame,
                };

                if (which == "form" && target == 0)
                {
                    return ApiError(wantedCaption is { Length: > 0 }
                        ? $"no running form whose caption holds '{wantedCaption}'"
                        : "no form is running");
                }

                var bytes = DebugCapture.CaptureBmp(target);
                if (bytes is null)
                {
                    return ApiError($"window {which ?? "frame"} would not render");
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
                        return ApiError($"nothing matches {cropSelector} on that surface");
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
                            ? ApiError($"{cropSelector} is not on screen")
                            : new ApiServer.ApiReply("image/bmp", cropped);
                    }
                    catch (Exception ex)
                    {
                        return ApiError($"the element's box could not be read ({ex.GetType().Name})");
                    }
                }

                return new ApiServer.ApiReply("image/bmp", bytes);
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
                    return ApiError("the surface is not up yet");
                }

                request.Query.TryGetValue("text", out var text);

                if (string.IsNullOrEmpty(text))
                {
                    return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                /*
                 * TWO WAYS THIS PRODUCT LEAVES THE EDITOR STOPPED, and both are cleared here.
                 *
                 * The first is a line that stopped inside the scratch module, which is what the
                 * comment above is about. The second is a COMPILE ERROR: evaluating anything in a
                 * project that will not compile makes the editor raise its box and then drop out
                 * of design mode a moment AFTER the evaluation has returned - measured at 40ms -
                 * so nothing on the way out can see it. What was left was an editor answering
                 * "Not available while execution is stopped" to every later evaluation and
                 * refusing every write, for ever, over one syntax error somebody was still
                 * typing. Reproduced in three calls: write a module ending `Public Sub Broken(`,
                 * evaluate anything, and the project never comes back.
                 *
                 * Cleared on the way IN rather than on the way out, and that is the whole point.
                 * A first attempt recovered right after the evaluation and reset the project
                 * while the evaluation was still in flight: it turned the useful "Compile error:
                 * Expected: identifier" into a bare COM error and took seventeen seconds to do
                 * it. Here there is nothing in flight to race.
                 */
                bool StoppedByUs() =>
                    ScratchBreakStanding() || (_immediateLeftItStopped && _inBreak);

                if (StoppedByUs())
                {
                    Log.Info("immediate: the editor is stopped by something this product ran, clearing it");

                    // Reset asks "proceed anyway?", and the rescue that answers a dialog blocking
                    // the host thread declines every real question - which is right for a question
                    // nobody here asked, and leaves THIS one cancelled and the wait below timing
                    // out for a reason nothing reported (issue #6). Held across the wait, because
                    // the confirmation appears while the command is still running.
                    using (Diagnostics.DialogWatch.ExpectingConfirmation(6000))
                    {
                        surface.RunOnHostThread(() => ExecuteEditorCommand(VbeCommands.Command.Reset));

                        var clearBy = Environment.TickCount64 + 5000;
                        while (Environment.TickCount64 < clearBy && StoppedByUs())
                        {
                            Thread.Sleep(100);
                        }
                    }

                    _immediateLeftItStopped = false;

                    if (StoppedByUs())
                    {
                        var stuck = "The last line left the editor stopped and it could not be "
                            + "cleared. Press Reset in the editor, or POST command?name=reset.";

                        Log.Warn($"immediate: {stuck}");
                        return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                            new DebugImmediateReply(false, stuck, true),
                            DebugJsonContext.Default.DebugImmediateReply));
                    }

                    if (!_inBreak)
                    {
                        surface.RunOnHostThread(RemoveScratchModule);
                    }
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

                // Set when a dialog has been answered that the evaluation cannot come back from,
                // so the wait below stops rather than running out its budget - see the note at
                // the break.
                var nothingLeftToWaitFor = false;

                while (Environment.TickCount64 < deadline && !evaluated.IsSet && !nothingLeftToWaitFor)
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

                        // MARKED HERE, not after the reply is composed, because the editor drops
                        // out of design mode about forty milliseconds from now and the poll that
                        // notices runs every hundred and fifty. Setting this at the end of the
                        // route lost that race every time: the poll saw the break, found the flag
                        // still false, and left it standing (measured 2026-08-24).
                        //
                        // AND ONLY IF THE EVALUATION HAS NOT COMPLETED, which is the whole safety
                        // of it and is knowable right here. A compile error is raised INSTEAD of
                        // running, so the line never started and there is no session to lose. A
                        // dialog raised after the evaluation completed belongs to code that DID
                        // run - `Debug.Print TheirFunction()` failing inside their own procedure -
                        // and the stop that follows is theirs, with their call stack on it. That
                        // is never cleared for them.
                        //
                        // Read off the completion event rather than off the `ran` computed later:
                        // later is exactly what loses the race to the poll.
                        if (!evaluated.IsSet)
                        {
                            _immediateLeftItStopped = true;

                            /*
                             * AND STOP WAITING FOR A COMPLETION THAT CANNOT COME.
                             *
                             * The loop's condition is `!evaluated.IsSet`, and a compile error is
                             * raised INSTEAD of running the line - so the evaluation never
                             * completes, the event is never set, and this waited out the whole
                             * fifteen-second budget before giving up. Two more seconds followed
                             * it. Measured three times at 17.0, 17.1 and 17.1 seconds, on the
                             * host thread, for a developer who typed in the Immediate window
                             * while their project had a syntax error somewhere - which is the
                             * ordinary state of a project being edited.
                             *
                             * The answer was already in hand at that point: the box's own words
                             * are the outcome, recorded a few lines above. Nothing was gained by
                             * the remaining fourteen and a half seconds except a frozen editor.
                             *
                             * A short grace first, because dismissing the box is occasionally
                             * what lets the evaluation finish, and a result is better than a
                             * complaint when both are available.
                             */
                            evaluated.Wait(500);
                            nothingLeftToWaitFor = true;
                            break;
                        }
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

                // TAKEN BACK IF THE EVALUATION RAN. The flag was set optimistically the moment
                // the editor complained, because the project drops out of design mode about
                // forty milliseconds later and the poll that clears it runs every hundred and
                // fifty - so noting it at the end of this route lost the race every time.
                //
                // Running is what distinguishes the two stops. A compile error means nothing of
                // the developer's ever started, so there is no session to lose and the break is
                // ours to clear. An evaluation that RAN and then stopped may be sitting at their
                // own breakpoint, in their own code, and that stop is theirs.
                if (ran)
                {
                    _immediateLeftItStopped = false;
                }

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugDrainReply(
                        WrappersLiveBefore: before,
                        WrappersLiveAfter: Com.ComRuntime.WrappersLive,
                        Survived: true),
                    DebugJsonContext.Default.DebugDrainReply));
            }

            case "session" when request.Query.TryGetValue("action", out var sessionAction):
            {
                // The one lifecycle a developer meets and no test could reach: a shutdown begun
                // and then cancelled. OnBeginShutdown stops the session before the host's save
                // prompt appears, and pressing Cancel there leaves the host running with the
                // add-in a corpse - the field failure of 2026-08-02. The watchdog is the guard,
                // and until now it could only be exercised by closing Excel by hand.
                if (sessionAction != "cancelledShutdown")
                {
                    return ApiError($"unknown action {sessionAction}; use cancelledShutdown");
                }

                var surface = _editorSurface;
                if (surface is null)
                {
                    return ApiError("the surface is not up; there is nothing to tear down");
                }

                // RESPOND FIRST, TEAR DOWN AFTER, and the order is the whole trick. Triggering
                // the shutdown inline would run Stop(), which disposes the very ApiServer
                // writing this reply, so the client would see a dropped connection instead of an
                // answer. The reply goes out here; a short pool-thread delay lets it flush, and
                // only then is BeginSimulatedShutdown posted to the HOST thread, where it must
                // run. The editor frame stays standing, so the watchdog reads a cancelled
                // shutdown and revives the session - a fresh ApiServer on a fresh port with a
                // fresh startedAt, rewritten into the discovery file, which is how the client
                // reconnects. See lessons on the shutdown watchdog.
                _ = System.Threading.Tasks.Task.Run(async () =>
                {
                    await System.Threading.Tasks.Task.Delay(400);
                    surface.RunOnHostThread(() => XlideAddIn.Current?.BeginSimulatedShutdown());
                });

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(true, 0,
                        "the session will stop and revive; reconnect by re-reading "
                        + "xlide-api-<pid>.json, whose port and startedAt the revived session rewrites"),
                    DebugJsonContext.Default.DebugCommandReply));
            }

            case "history":
            {
                // The session as a script. After a live investigation the useful sequence is
                // normally reconstructed from a scrollback and gets a step wrong; this hands
                // it back ready to run, so a bug found by hand becomes a probe by copying.
                var requests = ApiServer.Requests();
                var script = new System.Text.StringBuilder();
                script.AppendLine("# Replay of a xlide api session. Point it at a live instance:");
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

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugHistoryReply(requests, script.ToString(), ApiServer.RouteCosts()),
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

                // A NAME THAT IS NOT A CLAIM IS AN ERROR, not a claim that has yet to come true.
                // Polled like one it took the full timeout and then answered in the same shape a
                // real failure answers in, so a typo read as the product being broken and cost ten
                // seconds per occurrence while it did. `bench` and `trip` both refuse an unknown
                // argument outright; this does the same.
                if (Array.IndexOf(KnownClaims, claim) < 0)
                {
                    return ApiError($"unknown claim {claim}; known claims are {string.Join(", ", KnownClaims)}");
                }

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

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                    using var stateCrossing = CrossToHost(journalHost, () =>
                    {
                        try
                        {
                            sessionState = AnswerDebugRequestOnHost(
                                new ApiServer.ApiRequest("state", request.Query, string.Empty));
                        }
                        catch (Exception ex)
                        {
                            sessionState = $"{{\"error\":\"{ex.GetType().Name}\"}}";
                        }
                    });

                    if (!stateCrossing.Answered)
                    {
                        sessionState = null;
                    }
                }

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                if (Flag(request, "reset"))
                {
                    // BOTH, because this reply carries both. Resetting only the engine's figures
                    // left every other number in it describing a different stretch of time, with
                    // nothing saying which - see PerfCounters.Reset.
                    EngineCounters.Reset();
                    PerfCounters.Reset();
                }

                var (engineMethods, engineSlowest, engineWindow) = EngineCounters.Snapshot();
                var (hostReadCount, hostReadChars, hostReadFull, hostReadSkipped, hostReadSamples) = PerfCounters.HostReadSnapshot();
                _ = hostReadCount;
                var (publishCount, publishSamples) = PerfCounters.PublishSnapshot();
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugPerfReply(
                        placementSamples,
                        marshalSamples,
                        PerfCounters.HeartbeatAgeMs,
                        engineMethods,
                        engineSlowest,
                        engineWindow,
                        hostReadSamples,
                        hostReadChars,
                        hostReadFull,
                        hostReadSkipped,
                        publishSamples,
                        publishCount),
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
                // and a scraped row cannot tell "collapsed" from "rendered wrong" - the render
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
                return PageReply(ui);
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

                return PageReply(act);
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
                return PageReply(run);
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
                // still the browser's own synchronous evaluation - which the page's content
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
                    return ApiError(awaitError);
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

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugAwaitReply(met, elapsed, detail), DebugJsonContext.Default.DebugAwaitReply));
            }

            case "console":
            {
                // What the page said to itself. Only UNCAUGHT errors reach the shim log -
                // deliberately, because forwarding every line would drown it - so a warning
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
                    ? ApiError(consoleError)
                    : ApiServer.ApiReply.Json(read.Result);
            }

            case "inspect" when request.Query.TryGetValue("selector", out var selector) && selector.Length > 0:
            {
                // What the page actually has, where it is, and - with `styles` - what those
                // properties computed to, plus WHICH RULES claimed them.
                //
                // The rule list is the point. This page shares a document with a large
                // bundled stylesheet, and a structural class of ours (`.row` on a split
                // container) silently inherited `align-items: baseline` from an unrelated
                // rule, collapsing every cell to its tab strip's height. It read as a flex
                // bug in our own code and took an hour; the loop that finally found it -
                // walk every stylesheet, keep the rules this element matches - is this
                // route (2026-08-06).
                request.Query.TryGetValue("styles", out var wanted);
                var withRules = Flag(request, "rules");
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
                    ? ApiError(inspectError)
                    : ApiServer.ApiReply.Json(inspected.Result);
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
                    return ApiError($"unknown benchmark {what}; try tabswitch, layout, or type");
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
                    return ApiError(benchError);
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
                    return ApiError($"the benchmark's answer could not be read ({ex.GetType().Name})");
                }

                if (samples.Count == 0)
                {
                    return ApiError($"the benchmark ran nothing: {detail}");
                }

                return BenchReply(what, samples, detail);
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
                    return ApiError("no surface is up");
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
                                return ApiError(pageError);
                            }
                            tripSamples.Add(Math.Round(began.Elapsed.TotalMilliseconds, 3));
                        }

                        tripDetail = "a script into the page and its answer back";
                        break;
                    }

                    default:
                        return ApiError(
                            $"unknown trip {tripWhat}; pagecall is the only one. Anything that has to "
                            + "observe an effect delivered BY the host thread cannot be measured from in "
                            + "here at all - see the note on this route - and belongs in the client, the "
                            + "way tripCaret() does. For the cost of reaching the host thread, read "
                            + "perf().marshalMs, which every api request already samples from the far side");
                }

                return BenchReply(tripWhat, tripSamples, tripDetail);
            }

            case "layout" when Flag(request, "reset"):
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

                var restored = WaitForWorkspace(WaitMilliseconds(request, 20000));

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                    ? ApiError(layoutError)
                    : ApiServer.ApiReply.Json(layout.Result);
            }

            case "reload":
            {
                // Reload the page and WAIT for it to come back, answering with the bundle it
                // is now running. The manual version - reload, sleep a guess, hope - was run
                // a dozen times in one afternoon, and a guess that is too short reports on
                // the page that is going away (2026-08-06).
                var reloadHost = _editorSurface;
                if (reloadHost is null)
                {
                    return ApiError("the surface is not up yet");
                }

                var startedAt = Environment.TickCount64;
                _ = RunPageScript("location.reload()", null, 2000);

                // The budget runs from before the reload was posted, so the wait gets what is
                // left of it rather than a fresh allowance.
                var reloadBudget = WaitMilliseconds(request, 20000);
                var ready = WaitForWorkspace(reloadBudget - (Environment.TickCount64 - startedAt));

                var stamp = reloadHost.PageBuildStamp ?? "(none reported)";
                var bundle = BundleBuiltUtc();
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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

                // NOT drained here. `guard` owns the draining, and a read that emptied the list
                // would make this route the reason the next caller sees nothing - the same class
                // of problem it is here to expose. The last few are enough to tell "nothing
                // opened" from "something opened and was taken away".
                string[] cleared;
                lock (_dialogGate)
                {
                    cleared = [.. _guardCleared.TakeLast(5)];
                }

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugDialogsReply(rows, PerfCounters.HeartbeatAgeMs, cleared),
                    DebugJsonContext.Default.DebugDialogsReply));
            }

            case "compile":
            {
                // Does this project compile, and if not, what does it say?
                //
                // Not just the menu command. A compile error is a MODAL, so running it and
                // waiting on the host thread hangs the thread that raised it - which is how a
                // probe left one standing for six minutes, and why the answer nobody could read
                // was on screen the whole time (2026-08-07). The command is started and not
                // waited for; the answering happens here, on the door's own thread, which is the
                // only one still moving while a modal owns the editor.
                if (_editorSurface is not { } compileSurface)
                {
                    return ApiError("the surface is not up yet");
                }

                var standing = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);
                var command = VbeCommands.ForName("compile");

                // Whether the command STARTED is a separate question from what it said, and the
                // two used to collapse into one another: a greyed Compile item ran nothing, no
                // dialog appeared because nothing had compiled, and the reply said the project
                // compiled cleanly. The debugger suite reads this as its precondition, so a false
                // pass here turns every check behind it into a test of the dialog guard.
                //
                // The outcome is written on the host thread and read on this one, and the gate is
                // what orders the two. It is waited on AFTER the watch loop, never before: a
                // compile error is a modal raised inside the command itself, so the host thread
                // stays in there until the loop below answers it, and waiting first would be
                // waiting for the thing the waiting prevents.
                var started = VbeCommands.CommandRun.No("the host thread did not answer");
                using var ran = new ManualResetEventSlim(false);
                compileSurface.RunOnHostThread(() =>
                {
                    try
                    {
                        started = ExecuteEditorCommand(command);
                    }
                    finally
                    {
                        ran.Set();
                    }
                });

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

                // Set on the host thread the moment the command returns, so a wait that succeeds
                // is also the barrier that makes `started` safe to read here. By now the loop has
                // answered any modal the command raised, so this is short or already past.
                var reachedTheHost = ran.Wait(TimeSpan.FromSeconds(2));
                var startedOk = reachedTheHost && started.Ran;

                // WHAT COMPILED MEANS. No dialog appeared within waitMs, and the command that
                // would have raised one actually ran. It is not a positive report from the
                // compiler, because the host has no such report to give: the editor answers a
                // compile with a modal or with nothing at all. A caller that needs more than
                // "nothing objected" should read `errors`, and one that needs to know the
                // question was even asked should read `started`.
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugCompileReply(
                        startedOk && said.Count == 0,
                        [.. said],
                        DisplayFromProjectId(_shownProject) ?? string.Empty,
                        startedOk,
                        reachedTheHost ? started.Detail : "the command never reached the host thread"),
                    DebugJsonContext.Default.DebugCompileReply));
            }

            case "type" when request.Body.Length > 0 || request.Query.ContainsKey("text"):
            {
                // Types into the editor the way a person does, so the behaviour that only happens
                // WHILE typing can be tested: smart Enter, comment continuation, auto-indent.
                //
                // Through the editor's own keyboard pipeline - `trigger("keyboard", "type")` -
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
                // continuation is computed against a line that already has the next line on it -
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
                return PageReply(typed);
            }

            case "mark" when request.Query.TryGetValue("text", out var marker) && marker.Length > 0:
            {
                // A labelled line in the log, and the offset it landed at.
                //
                // Reading a log for what one step did means finding where that step began, and
                // "scroll up until it looks like the right place" is how a session ends up
                // reasoning about the wrong three seconds. A probe that marks its steps can ask
                // for exactly the slice between two marks - `log({ since })` with the offset this
                // hands back.
                // The offset is taken BEFORE the marker is written, so reading from it returns
                // the marker itself - a slice that starts with the words the caller chose is a
                // slice they can be sure is theirs.
                var at = Log.Path is { } logPath && File.Exists(logPath)
                    ? new FileInfo(logPath).Length
                    : 0;

                Log.Info($"---- {marker} ----");

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugMarkReply(marker, at),
                    DebugJsonContext.Default.DebugMarkReply));
            }

            case "guard":
            {
                // No host thread here either: turning the guard on is exactly what a caller does
                // when the host thread has already stopped answering.
                if (request.Query.ContainsKey("on"))
                {
                    // The fallback is what it already is, so an unreadable value leaves the guard
                    // alone rather than turning it off - this one governs whether dialogs are
                    // dismissed, and guessing "off" from a typo is the expensive direction.
                    _guardEverything = Flag(request, "on", _guardEverything);
                    Log.Info($"xlide api: the dialog guard is {(_guardEverything ? "on" : "off")}");
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

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(dismissed, 0), DebugJsonContext.Default.DebugCommandReply));
            }

            case "userform":
            {
                // No host thread, twice over: enumeration is pure Win32, and the one moment
                // this route matters - watching or closing a modally RUNNING form - is
                // exactly when the host thread is parked inside the Run command until that
                // form goes away. A host-bound route here would deadlock against its caller.
                request.Query.TryGetValue("action", out var formAction);
                request.Query.TryGetValue("caption", out var formCaption);

                if (string.Equals(formAction, "close", StringComparison.OrdinalIgnoreCase))
                {
                    var closed = DialogWatch.CloseRunningForm(formCaption);
                    return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                        new DebugCommandReply(closed, 0, closed ? "closed" : "no running form matched"),
                        DebugJsonContext.Default.DebugCommandReply));
                }

                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                    new DebugUserFormsReply(DialogWatch.RunningForms()),
                    DebugJsonContext.Default.DebugUserFormsReply));
            }

            case "stats":
            {
                var placement = PerfCounters.PlacementSnapshot();
                var marshal = PerfCounters.MarshalSnapshot();
                var refresh = PerfCounters.RefreshSnapshot();
                var follow = PerfCounters.FollowSnapshot();
                var messages = WebView.WebView2Surface.MessageTap.Totals;
                using var self = System.Diagnostics.Process.GetCurrentProcess();
                return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
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
                        ComWrappersLive: Com.ComRuntime.WrappersLive,
                        LaneHolder: _laneHolder,
                        LaneHeldMs: _laneHolder is null
                            ? 0
                            : Environment.TickCount64 - System.Threading.Interlocked.Read(ref _laneHeldSince)),
                    DebugJsonContext.Default.DebugStatsReply));
            }
        }

        var host = _editorSurface;
        if (host is null)
        {
            return ApiError("the surface is not up yet");
        }

        // The sweep already ran, at the top of this method, so every route heals - including the
        // ones that answer without the host thread and used to return before reaching it. A modal
        // this door raised earlier may still be standing, and waiting for a timeout to notice is
        // the wrong instrument: a VBA modal PUMPS messages, so marshaled work still runs and no
        // timeout ever comes (measured 2026-08-06 - state answered normally while the Macros
        // dialog owned the editor), while the developer is looking at a stuck editor throughout.

        // What was already standing before this request. Anything that appears while it is
        // in flight was raised BY it, and only those may be answered automatically: a dialog
        // the developer opened is theirs, and closing it under them would be worse than any
        // hang. See the timeout path below.
        var standingBefore = DialogWatch.Dialogs().Select(row => row.Window).ToHashSet(StringComparer.Ordinal);

        // The crossing samples PerfCounters.Marshal itself: every marshaled request doubles as
        // a probe of the host thread's responsiveness, and the stats route serves the sample.
        string? answer = null;
        using var crossing = CrossToHost(host, () =>
        {
            // THE LANE'S OWN EYES. One work item that will not finish starves every request
            // behind it while the heartbeat ticks on - a door dark for four minutes was read
            // as "VBA is running your code" when the holder was a reload's teardown, and
            // nothing anywhere could name it (#12, chaos seed 2009959200). The route is
            // recorded while its work is ON the thread, `stats` serves it without needing
            // that thread, and a hold past five seconds writes itself into the log so a
            // post-mortem needs no live observer.
            _laneHolder = request.Route;
            System.Threading.Interlocked.Exchange(ref _laneHeldSince, Environment.TickCount64);
            try
            {
                answer = AnswerDebugRequestOnHost(request);
            }
            catch (Exception ex)
            {
                answer = HostError($"{ex.GetType().Name}: {ex.Message}");
            }
            finally
            {
                var heldMs = Environment.TickCount64
                    - System.Threading.Interlocked.Read(ref _laneHeldSince);
                if (heldMs > 5000)
                {
                    Log.Info($"marshal: '{request.Route}' held the host lane for {heldMs}ms");
                }

                _laneHolder = null;
            }
        });

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

        if (crossing.Answered && answer is not null)
        {
            return ApiServer.ApiReply.Json(answer);
        }

        // A request that asked to keep what it opens is not rescued from it: opening a modal
        // was the point, and the caller dismisses it when finished.
        return AnswerBlockedRequest(standingBefore, crossing.Done, () => answer, request.Query.ContainsKey("keep"));
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
            // finished announcement off the screen - and off the host thread it is holding.
            if (!mine && !(_guardEverything && DialogWatch.IsNotice(dialog)))
            {
                continue;
            }

            var answer = DialogWatch.SafeAnswerFor(dialog);
            var pressed = answer is not null && DialogWatch.Dismiss(dialog.Caption, answer) ? answer : null;

            Log.Info(pressed is null
                ? $"xlide api: \"{dialog.Caption}\" has the editor and offers no safe button; leaving it"
                : $"xlide api: cleared {(mine ? "our" : "a standing")} dialog \"{dialog.Caption}\""
                    + $"{(dialog.Text.Length > 0 ? $" ({dialog.Text})" : string.Empty)} with {pressed}, "
                    + $"host thread quiet for {PerfCounters.HeartbeatAgeMs}ms");

            lock (_dialogGate)
            {
                // UNDER THE GATE, like every read of it. This add was outside any lock while the
                // `guard` route enumerated the same list from a request thread, which is a list
                // being written during someone else's enumeration: rare, and an exception in a
                // watcher rather than a wrong answer. It matters more now that `dialogs` reads it
                // too, so there are two readers and one writer instead of one of each.
                //
                // OUR OWN SWEEPS LAND HERE TOO. They used to be left out (`!mine`), so a dialog
                // this door raised and healed left no trace anywhere a caller could read:
                // `recentlyCleared` answered empty, and "raised and swept" was indistinguishable
                // from "never raised" - the exact ambiguity the field exists to resolve, and the
                // same confusion that once had a suite read an empty list as a failure to raise
                // (2026-08-08). Measured 2026-08-28: a bare References raise, swept on the next
                // request, with recentlyCleared answering [] throughout. The compile-error
                // detection below had already paid for the gap once and routes around it.
                if (pressed is not null)
                {
                    _guardCleared.Add($"{dialog.Caption}: {dialog.Text}".Trim().TrimEnd(':'));
                }

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
    /// Whether the last thing the guard took off the screen was a compile error.
    ///
    /// The one piece of evidence that separates a test run stopped by a project that will not
    /// compile - which nothing can continue from, and which the debug poll clears - from a
    /// developer's own breakpoint inside a test, which is theirs and is never touched.
    ///
    /// Under the gate, like every read of this list: it is written from a watcher on a pool
    /// thread while request threads enumerate it.
    /// </summary>
    private bool GuardClearedACompileError()
    {
        // THE WATCH ANSWERS MOST OF THEM, and it is the one that answers this. `_guardCleared`
        // used to leave out a dialog raised by the request that is running - which a test run's
        // compile box always is - so asking only that list found nothing, every time, while the
        // log showed the watch answering the box on another thread two lines earlier. The list
        // carries the door's own sweeps too since 2026-08-28; the watch stays first because it
        // answers the boxes that never cross the sweep at all.
        if (Diagnostics.DialogWatch.AnsweredRecently("Compile error"))
        {
            return true;
        }

        lock (_dialogGate)
        {
            for (var at = _guardCleared.Count - 1; at >= 0 && at >= _guardCleared.Count - 3; at--)
            {
                if (_guardCleared[at].Contains("Compile error", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }


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
                                ? $"xlide api: keeping \"{dialog.Caption}\", as the request asked"
                                : $"xlide api: a request raised \"{dialog.Caption}\"; "
                                    + "it will be cleared unless the request asked to keep it");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("xlide api: the dialog watch failed", ex);
            }
        });
    }

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

            // THE HOST'S ANSWER TO THE SAME QUESTION. `shownModule` above asks the surface what it
            // believes it is showing, which is this session's own record: useful, and not what Run,
            // Step and ToggleBreakpoint act on. Those act on the native active pane, so a claim
            // about where a command will land has to read the pane. Named separately rather than
            // changing what `shownModule` means, because a route that quietly starts answering a
            // different question breaks every caller that was right about the old one.
            case "nativeModule":
            {
                try
                {
                    using var activePane = _editor.GetObject("ActiveCodePane");
                    using var codeModule = activePane?.GetObject("CodeModule");
                    using var component = codeModule?.GetObject("Parent");
                    var active = component?.GetString("Name");

                    return (active is not null
                        && (expected is null || active.Equals(expected, StringComparison.OrdinalIgnoreCase)),
                        active ?? "(no active pane)");
                }
                catch (Exception ex)
                {
                    return (false, $"the active pane could not be read ({ex.GetType().Name})");
                }
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
    /// Every claim <see cref="EvaluateClaim"/> answers, so the route can refuse a name it does not
    /// know instead of waiting for it to come true. One list rather than two: a vocabulary kept in
    /// the switch and repeated in a guard is a vocabulary that drifts.
    /// </summary>
    private static readonly string[] KnownClaims =
    [
        "stopped",
        "running",
        "surfaceReady",
        "shownModule",
        "nativeModule",
        "noDialogs",
        "localsHas",
        "watchHas",
        "problemFree",
        "responsive",
    ];

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
    /// string; a script that builds its answer with JSON.stringify - which every useful one does,
    /// because that is how a structure crosses - returns it quoted twice. Unwrapping stops at the
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
    /// answering it is undoing our own mess, so it is dismissed. Only a SAFE button is ever
    /// pressed: Cancel, then Close, then No. Never OK, Yes, Save, Delete, or Run - a dialog
    /// nobody read must not be agreed with. A dialog that was already standing belongs to the
    /// developer and is only reported.
    ///
    /// THE REQUEST IS NOT RE-SENT. Dismissing the dialog releases the host thread, and the work
    /// this request asked for was queued before the dialog appeared, so what happens next is
    /// that the queued work is given three seconds to finish on its own. Retrying would run it
    /// twice. This comment said "retried once" for as long as it sat above the wrong method, and
    /// the reply carried a `retried` field that was the constant false to match.
    /// </summary>
    private static ApiServer.ApiReply AnswerBlockedRequest(
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

            /*
             * WHY IT DID NOT ANSWER, as far as this side can tell without it.
             *
             * "The host thread did not answer in time" is true and tells a caller nothing it can
             * act on. The commonest cause by far is the developer's own code RUNNING - a loop
             * with no exit condition owns that thread until it finishes, and this door's own
             * Break and Reset are marshalled onto the same thread, so they queue behind the very
             * thing they exist to stop. Measured 2026-08-24: `Do ... Loop` with no exit left
             * Excel not responding, every route timing out, and break and reset among them.
             *
             * The heartbeat is the evidence. It is stamped by the host thread's own tick, so an
             * old one means that thread has not been round its loop - which is exactly the
             * difference between a door that is broken and a door waiting behind running code.
             * A caller told that can act; a caller told "did not answer" cannot.
             *
             * The remedy is named because this door does not have one. Ctrl+Break is a keyboard
             * interrupt VBA's own pump handles, and nothing marshalled can substitute for it.
             */
            var asleep = PerfCounters.HeartbeatAgeMs;
            var why = asleep > 2000
                ? $"the host thread has not answered for {asleep}ms. That usually means VBA is "
                    + "running your code - a loop with no exit owns the editor's thread until it "
                    + "ends, and this door's own Break and Reset need that same thread, so they "
                    + "cannot interrupt it. Press Ctrl+Break in the editor."
                : "the host thread did not answer in time";

            return ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                new DebugBlockedReply(
                    Error: why,
                    HeartbeatAgeMs: asleep,
                    BlockedBy: standing.Length > 0 ? standing[0].Caption : null,
                    Buttons: standing.Length > 0 ? standing[0].Buttons : [],
                    Dismissed: null),
                DebugJsonContext.Default.DebugBlockedReply));
        }

        var safe = DialogWatch.SafeAnswerFor(blocking);
        var pressed = safe is not null && DialogWatch.Dismiss(blocking.Caption, safe) ? safe : null;

        Log.Info(pressed is null
            ? $"xlide api: \"{blocking.Caption}\" is blocking the host thread and has no safe button"
            : $"xlide api: \"{blocking.Caption}\" was raised by this request; answered with {pressed}");

        // The dismissal releases the host thread, and the work this request asked for was
        // queued before the dialog appeared, so it may complete on its own.
        var completed = pressed is not null && done.Wait(TimeSpan.FromSeconds(3));

        return completed && answerSoFar() is { } answer
            ? ApiServer.ApiReply.Json(answer)
            : ApiServer.ApiReply.Json(System.Text.Json.JsonSerializer.Serialize(
                new DebugBlockedReply(
                    Error: pressed is null
                        ? "a dialog this request raised is blocking the host thread, and it has no safe button to press"
                        : "a dialog this request raised was dismissed, but the request did not finish",
                    HeartbeatAgeMs: PerfCounters.HeartbeatAgeMs,
                    BlockedBy: blocking.Caption,
                    Buttons: blocking.Buttons,
                    Dismissed: pressed),
                DebugJsonContext.Default.DebugBlockedReply));
    }

    /// <summary>
    /// Answers one request, with everything it writes attributed to whoever asked.
    ///
    /// A write through this door is never the developer at the keyboard - it is a program, whether
    /// or not it says which. `by=` names the caller and is what makes a change log worth reading
    /// once two agents are working; without it the honest record is that something wrote and did
    /// not say who, which is exactly what `unattributed` means.
    /// </summary>
    private string AnswerDebugRequestOnHost(ApiServer.ApiRequest request)
    {
        request.Query.TryGetValue("by", out var by);
        return AttributedTo(by, () => AnswerDebugRouteOnHost(request));
    }

    private unsafe string AnswerDebugRouteOnHost(ApiServer.ApiRequest request)
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
                        FrameVisible: _frame != 0 && Win32.IsWindowVisible(_frame),
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

                    /*
                     * ONE DIRECTION ONLY, and the direction is the whole point.
                     *
                     * A SHIM newer than the bundle beside it is the real hazard: something
                     * published the shim and did not redeploy the page, so the session is serving
                     * an old surface over new host code.
                     *
                     * A BUNDLE newer than the shim is the ordinary page loop - `tools\page.ps1` a
                     * dozen times in an afternoon over a shim nobody has touched - and the shim is
                     * not stale, it is finished. Measuring the gap with Duration() flagged that as
                     * suspicious too, and being refused for it costs more than the check saves:
                     * every generator opens by making a blank workbook and only then asks the
                     * doctor, so a refusal here leaves FormFixture.xlsm a blank book with no
                     * project in it. That happened twice on 2026-08-17 and cost a rebuild each
                     * time. Whether the shim matches its own SOURCES is a different question, and
                     * tools\dev.ps1 already answers it by comparing the newest .cs against the
                     * published DLL and refusing to go on.
                     */
                    var pageBehind = shimBuilt - bundleBuilt;
                    if (pageBehind > TimeSpan.FromMinutes(30))
                    {
                        findings.Add($"the shim is {pageBehind.TotalMinutes:N0} minutes newer than "
                            + "the page bundle beside it; the page was probably not redeployed");
                    }
                }
                else
                {
                    findings.Add("the shim directory does not hold both a shim and a page bundle");
                }

                if (_editorSurface?.PageBuildStamp is null)
                {
                    /*
                     * A CAUSE, not the symptom restated.
                     *
                     * This used to say only that the page had not reported a stamp, which is the
                     * observation written out longhand: it names what did not happen and nothing
                     * about why. The page's own black box knows why - boot.js installs before the
                     * bundle and catches a module that throws on load - and the JavaScript context
                     * survives that throw, so it can be read out of a page that never booted.
                     *
                     * Read here rather than left to a second call, because the whole point of a
                     * doctor is that one question is enough (2026-08-09: the page died, doctor
                     * described the silence, and the reason was found by reading source).
                     */
                    // PUSHED FIRST, asked second. boot.js posts the error to the host the moment
                    // it happens, so the usual case costs nothing and works even when the surface
                    // is too far gone to answer a script. The script read stays as the fallback
                    // for a page serving an older bundle that does not push.
                    var pushed = _editorSurface?.FirstPageError;
                    var (died, asked) = pushed is not null ? (pushed, true) : AskPageForItsError();
                    findings.Add(died is not null
                        ? $"the page never reported a build stamp because it THREW while loading: {died}"
                        : asked
                            ? "the page never reported a build stamp, and it recorded no error, so it "
                                + "is more likely still starting than broken"
                            : "the page never reported a build stamp, and it did not answer when asked "
                                + "what went wrong, so the surface itself is not responding");
                }

                // A standing dialog owns the host thread, and every OTHER route answers normally
                // while it does - so a session can look healthy for minutes while nothing it is
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
                var wantEngineText = Flag(request, "text");
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
                    return HostError($"the engine's copy could not be read ({ex.GetType().Name})");
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
                var wantText = Flag(request, "text");

                // EVERY open pane, each with the host's content and the surface's side by side.
                //
                // The active one is not the only one that can drift. A background tab holds a
                // copy the developer is not looking at, so a module written from outside while
                // its tab sits behind another goes stale with nothing to notice until it is
                // clicked - and then it is the developer who notices.
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
                            window.GetBool("Visible"),
                            window.GetInt32("HWnd")));
                    }
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugWindowsReply([.. rows]), DebugJsonContext.Default.DebugWindowsReply);
            }

            case "menus":
            {
                /*
                 * THE EDITOR'S MENUS, WITH THEIR IDS, including the ones the surface suppresses.
                 *
                 * The suppression table in VbeMenus is a list of numbers, and until this existed
                 * the only way to learn a number was to enumerate the bar by hand once and write
                 * it into a comment. That is how a table of magic numbers goes quietly out of
                 * date, and it is why removing a menu was a measurement session rather than an
                 * edit. `suppressed` is this session's own answer about each one, so what the
                 * table does can be read back rather than inferred.
                 *
                 * `path` is the position chain, comma separated; absent means the bar itself.
                 */
                request.Query.TryGetValue("path", out var menuPath);
                var positions = (menuPath ?? string.Empty)
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Select(part => int.TryParse(part, out var one) ? one : 0)
                    .Where(one => one > 0)
                    .ToArray();

                var menuRows = VbeMenus.Describe(_editor, positions)
                    .Select(row => new DebugMenuRow(
                        row.Index, row.Id, row.Caption, row.Popup, row.Enabled, row.Suppressed))
                    .ToArray();

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugMenusReply(string.Join(",", positions), menuRows),
                    DebugJsonContext.Default.DebugMenusReply);
            }

            case "analysis":
            {
                /*
                 * MACHINE-WIDE ANALYZER RULE SETTINGS, the same mechanism the rules modal, the
                 * problems pane's menu and the lightbulb use - one ApplyRuleSeverityAsync for
                 * all of them, so the api cannot drift from the UI.
                 *
                 * GET lists every rule the bundled analyzer ships, each with the override values
                 * it PERMITS and the override standing on this machine. `rule=` with `severity=`
                 * changes one: off, warning, error, information, or default to clear. The guard
                 * is the analyzer's own - an illegal move is refused here in words, because the
                 * engine would silently ignore it and nothing would say why.
                 *
                 * This is policy for the MACHINE, persisted in settings.json in user space. One
                 * finding at one line wants the inline directives instead, which travel with the
                 * code; the suppression quick fix writes those.
                 */
                request.Query.TryGetValue("rule", out var analysisRule);
                request.Query.TryGetValue("severity", out var analysisSeverity);

                string? analysisDetail = null;

                // The INLINE half, drivable from the same route: module + line + code writes the
                // directive the problems pane's menu writes, through the identical mechanism.
                if (request.Query.TryGetValue("module", out var suppressModule)
                    && suppressModule.Length > 0)
                {
                    if (!request.Query.TryGetValue("line", out var suppressLineText)
                        || !int.TryParse(suppressLineText, out var suppressLine)
                        || suppressLine < 1
                        || !request.Query.TryGetValue("code", out var suppressCode)
                        || suppressCode.Length == 0)
                    {
                        return HostError("suppressing inline takes module=, line= (1-based) "
                            + "and code=");
                    }

                    request.Query.TryGetValue("project", out var suppressProject);
                    analysisDetail = SuppressFinding(
                        suppressModule, suppressProject, suppressLine, suppressCode)
                        ?? $"suppressed {suppressCode} at {suppressModule}({suppressLine}) "
                            + "with an inline directive";
                }

                if (analysisRule is { Length: > 0 } || analysisSeverity is { Length: > 0 })
                {
                    if (analysisRule is not { Length: > 0 } || analysisSeverity is not { Length: > 0 })
                    {
                        return HostError("changing a rule takes both rule= and severity=; "
                            + "a bare GET lists the catalog");
                    }

                    analysisDetail = ApplyRuleSeverityAsync(analysisRule, analysisSeverity)
                        .GetAwaiter().GetResult();
                }

                var analysisCatalog = AnalysisRuleCatalogAsync().GetAwaiter().GetResult();
                if (analysisCatalog is null)
                {
                    return HostError("the analysis engine is not up, so the rule catalog "
                        + "cannot be read");
                }

                var standing = _settings.AnalysisRuleSeverityOverrides;
                var analysisRows = analysisCatalog.Rules
                    .Select(rule => new DebugAnalysisRuleRow(
                        rule.Code,
                        rule.Title,
                        rule.Category,
                        rule.DefaultSeverity,
                        rule.Allowed,
                        rule.SuppressionScopes,
                        standing is not null && standing.TryGetValue(rule.Code, out var held)
                            ? held
                            : null))
                    .ToArray();

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugAnalysisReply(analysisDetail, SettingsPath, analysisRows),
                    DebugJsonContext.Default.DebugAnalysisReply);
            }

            case "bars" when request.Query.TryGetValue("name", out var barsName) && barsName.Length > 0:
            {
                /*
                 * WHERE A COMMAND LIVES, AND WHICH COPY ANSWERS.
                 *
                 * `command` reports one word - "currently disabled" - and until this existed there
                 * was no way to ask which of the several controls carrying that command said it.
                 * Reset alone is on the menu bar, on two toolbars and on three context menus, and
                 * this surface hides every toolbar the editor came with, so a greyed answer could
                 * be a copy nobody can see. Issue #9 is a break with every debug command refused
                 * and no visible cause; a reading that cannot name the copy cannot tell a real
                 * refusal from a stale one.
                 *
                 * `mode` beside it is the ACTIVE PROJECT's, read live rather than from the polled
                 * value the page is told, because a stale published mode looks exactly the same
                 * from the outside.
                 */
                var barsId = VbeCommands.ForName(barsName);
                if (barsId == 0 && !int.TryParse(barsName, out barsId))
                {
                    return HostError($"unknown command name {barsName}; pass a name "
                        + "VbeCommands.ForName knows, or a bare identifier");
                }

                var places = VbeCommands.Places(_editor, barsId)
                    .Select(place => new DebugBarRow(
                        place.Bar, place.BarVisible, place.Enabled, place.State))
                    .ToArray();

                string? activeProject = null;
                var liveMode = -1;
                string? modeError = null;
                try
                {
                    using var active = _editor.GetObject("ActiveVBProject");
                    activeProject = active?.GetString("Name");
                    liveMode = active?.GetInt32("Mode") ?? -1;
                }
                catch (Exception ex)
                {
                    // NAMED, NEVER SWALLOWED. The debug-state poll swallows this same failure and
                    // leaves the page holding whatever mode it last heard, which is one of the
                    // ways a session can report break long after the break is over.
                    modeError = $"{ex.GetType().Name}: {ex.Message}";
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugBarsReply(
                        barsName,
                        barsId,
                        places,
                        places.Count(row => row.Enabled),
                        activeProject,
                        liveMode,
                        _lastPublishedMode,
                        modeError),
                    DebugJsonContext.Default.DebugBarsReply);
            }

            case "outline" when request.Query.TryGetValue("module", out var outlineModule) && outlineModule.Length > 0:
            {
                // A module's shape, from the analyzer, so a caller can assert on structure rather
                // than read the text back and parse it a second time - in a second language, with
                // a second set of bugs.
                if (_analysis is not { } outlineAnalysis)
                {
                    return HostError("the analysis engine is not up");
                }

                request.Query.TryGetValue("project", out var outlineProject);

                var outlineOwner = ResolveNamedProject(outlineProject, out var outlineUnknown);
                if (outlineUnknown is not null)
                {
                    return HostError(outlineUnknown);
                }

                try
                {
                    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
                    var answered = outlineAnalysis
                        .OutlineAsync(outlineModule, outlineOwner, source: null, deadline.Token)
                        .GetAwaiter().GetResult();

                    if (answered is null)
                    {
                        return HostError($"'{outlineModule}' could not be outlined");
                    }

                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugOutlineReply(
                            outlineModule,
                            [.. answered.Select(p => new DebugProcedureRow(p.Name, p.Kind, p.Line))]),
                        DebugJsonContext.Default.DebugOutlineReply);
                }
                catch (Exception ex)
                {
                    return HostError($"outline failed: {ex.Message.Trim()}");
                }
            }

            case "sync":
                return HandleSync(request.Query, request.Body);

            // The VBA test runner, mirrored: the same brain the Tests pane presses. No action
            // answers the snapshot alone; `run` waits for the whole run (the pane streams the
            // same results live), `debug` waits for the debug session to end, which is the
            // honest shape of debugging - the fleet's other doors still answer meanwhile.
            case "tests":
            {
                request.Query.TryGetValue("action", out var testsAction);
                request.Query.TryGetValue("module", out var testsModule);
                request.Query.TryGetValue("test", out var testsTarget);
                request.Query.TryGetValue("file", out var testsFile);

                var testsDetail = "listed";
                if (testsAction is { Length: > 0 } and not "list")
                {
                    // `run` says what it means by what it carries: a test, a module, a file, or
                    // nothing at all - which is every test in every open file.
                    var testsVerb = testsAction == "run" && testsTarget is { Length: > 0 } ? "runOne"
                        : testsAction == "run" && testsModule is { Length: > 0 } ? "runModule"
                        : testsAction == "run" && testsFile is { Length: > 0 } ? "runFile"
                        : testsAction;
                    testsDetail = HandleTestsAction(
                        testsVerb,
                        testsTarget is { Length: > 0 } ? testsTarget : testsModule,
                        testsFile);
                }

                // From the CACHE, not a fresh walk: every action above has just walked the open
                // files itself, so a second read here answers the same thing at the same cost
                // again - a third of a second on a large project, on every route call
                // (measured 2026-08-20). A bare list with nothing walked yet still reads.
                var testsNow = TestsSnapshotCached();
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugTestsReply(
                        testsDetail, testsNow.Support, testsNow.Running, testsNow.CurrentTest,
                        testsNow.RanAt, testsNow.Files, testsNow.Rows),
                    DebugJsonContext.Default.DebugTestsReply);
            }

            /*
             * THE CHANGE LOG: what happened to this project's module code, by whom, in rounds.
             *
             * READ-ONLY, and deliberately so. There is no revert verb here and there is not going
             * to be one: putting text back is a WRITE, and this product already has a hardened one
             * at `module`. An agent that wants to undo something reads the text it wants from
             * `action=text` and writes it, which lands in this log like any other write - so the
             * clever thing an agent does stays visible, attributable, and reversible by the same
             * means. Cleverness in the caller, dumbness in the store.
             *
             * `action=snapshot&label=` ends the round that is running and names it. It costs
             * nothing - a round is a divider, not a copy - so call it as often as the work has
             * shape. Label it after the fact: an agent describes what it DID far better than what
             * it was about to do, and one that forgets to open a round still leaves a usable one.
             */
            case "changes":
            {
                request.Query.TryGetValue("action", out var changesAction);
                request.Query.TryGetValue("project", out var changesProject);
                request.Query.TryGetValue("module", out var changesModule);
                request.Query.TryGetValue("label", out var changesLabel);
                request.Query.TryGetValue("which", out var changesWhich);
                request.Query.TryGetValue("round", out var changesRound);
                request.Query.TryGetValue("limit", out var changesLimit);

                var changesAt = int.TryParse(changesRound, out var parsedRound) ? parsedRound : 0;
                var changesMost = int.TryParse(changesLimit, out var parsedLimit) ? parsedLimit : 200;

                switch (changesAction)
                {
                    case "text":
                        return ChangeTextReply(changesProject, changesAt, changesModule, changesWhich);

                    case "diff":
                        return ChangeDiffReply(changesProject, changesAt, changesModule);

                    case "snapshot":
                        CloseChangeRounds(changesLabel);
                        return ChangesReply(
                            changesLabel is { Length: > 0 } named
                                ? $"the round is closed: {named}"
                                : "the round is closed",
                            changesProject, changesMost);

                    case "accept":
                    {
                        var accepting = ChangeLogFor(ChangeLogProject(changesProject));
                        accepting?.Accept(DateTimeOffset.UtcNow);
                        return ChangesReply("accepted", changesProject, changesMost);
                    }

                    default:
                        return ChangesReply("listed", changesProject, changesMost);
                }
            }

            // What a control of a KIND holds untouched, measured from a bare instance of the
            // coclass MSForms registers - no form, no workbook, nothing on screen. This is the
            // inventory the markup projection compares against to decide what a developer has
            // actually changed, and reading it is how anyone checks that claim.
            case "defaults" when request.Query.TryGetValue("type", out var defaultsType) && defaultsType.Length > 0:
            {
                var known = _controlDefaults.For(defaultsType);
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugDefaultsReply(defaultsType, known.Count,
                        [.. known.OrderBy(one => one.Key, StringComparer.OrdinalIgnoreCase)
                            .Select(one => new DebugDefaultRow(one.Key, one.Value))]),
                    DebugJsonContext.Default.DebugDefaultsReply);
            }

            // The markup language's whole vocabulary, exactly as the designer tab's completions
            // and hovers hold it: one answer, two doors. `module` names a live form so the Form's
            // own entry can be described; without it every other kind still answers.
            case "vocabulary":
            {
                request.Query.TryGetValue("module", out var vocabularyModule);
                request.Query.TryGetValue("project", out var vocabularyProject);

                using var vocabularyForm = vocabularyModule is { Length: > 0 }
                    ? FindComponent(vocabularyModule, ResolveNamedProject(vocabularyProject, out _), out _)
                    : null;
                return System.Text.Json.JsonSerializer.Serialize(
                    new FormMarkupVocabularyMessage(
                        "formMarkupVocabulary",
                        FormMarkupVocabulary.Of(_controlDefaults, _propertyTypes, vocabularyForm)),
                    EditorMessageContext.Default.FormMarkupVocabularyMessage);
            }

            case "designer" when request.Query.TryGetValue("module", out var designerModule) && designerModule.Length > 0:
            {
                // A UserForm's design, read and mutated through the MSForms designer object
                // model - docs/userform-designer.md's M1 instrument. No action reads; add,
                // remove and set are the three mutations a form fixture is made of, and they
                // go through the same model the native toolbox calls, which is what keeps a
                // form built here byte-compatible with one built by hand.
                request.Query.TryGetValue("project", out var designerProject);

                if (!request.Query.TryGetValue("action", out var designerAction) || designerAction.Length == 0)
                {
                    // format=markup answers the same walk projected through Core's printer:
                    // the form as text, in the markup layer's dialect.
                    return request.Query.TryGetValue("format", out var designerFormat)
                        && string.Equals(designerFormat, "markup", StringComparison.OrdinalIgnoreCase)
                        ? DesignerMarkup(designerModule, designerProject)
                        : DesignerRead(designerModule, designerProject);
                }

                switch (designerAction)
                {
                    case "applyMarkup":
                        return request.Body.Length > 0
                            ? DesignerApplyMarkup(designerModule, designerProject, request.Body)
                            : HostError("applyMarkup takes the markup document as the request body");

                    case "add":
                    {
                        request.Query.TryGetValue("type", out var addType);
                        if (addType is not { Length: > 0 })
                        {
                            return HostError("add needs type=<control kind or ProgID>");
                        }

                        request.Query.TryGetValue("name", out var addName);
                        request.Query.TryGetValue("parent", out var addParent);
                        request.Query.TryGetValue("left", out var addLeft);
                        request.Query.TryGetValue("top", out var addTop);
                        request.Query.TryGetValue("width", out var addWidth);
                        request.Query.TryGetValue("height", out var addHeight);
                        return DesignerAdd(designerModule, designerProject, addType, addName, addParent,
                            DesignerNumber(addLeft), DesignerNumber(addTop),
                            DesignerNumber(addWidth), DesignerNumber(addHeight));
                    }

                    case "remove":
                        return request.Query.TryGetValue("name", out var removeName) && removeName.Length > 0
                            ? DesignerRemove(designerModule, designerProject, removeName)
                            : HostError("remove needs name=<control>");

                    case "set":
                    {
                        request.Query.TryGetValue("name", out var setName);
                        request.Query.TryGetValue("property", out var setProperty);
                        request.Query.TryGetValue("value", out var setValue);
                        request.Query.TryGetValue("as", out var setAs);
                        return setProperty is { Length: > 0 } && setValue is not null
                            ? DesignerSet(designerModule, designerProject, setName, setProperty, setValue, setAs)
                            : HostError("set needs property= and value=; name= targets a control, omitted targets the form");
                    }

                    case "zorder":
                    {
                        request.Query.TryGetValue("name", out var zName);
                        request.Query.TryGetValue("to", out var zTo);
                        return zName is { Length: > 0 }
                            ? DesignerZOrder(designerModule, designerProject, zName,
                                !string.Equals(zTo, "back", StringComparison.OrdinalIgnoreCase))
                            : HostError("zorder needs name=<control> and to=front|back");
                    }

                    case "autosize":
                    {
                        request.Query.TryGetValue("name", out var fitName);
                        return designerModule is { Length: > 0 } && fitName is { Length: > 0 }
                            ? DesignerAutoSize(designerModule, designerProject, fitName)
                            : HostError("autosize needs module=<form> and name=<control>");
                    }

                    case "liveness":
                        return DesignerLiveness();

                    case "baseline":
                        return designerModule is { Length: > 0 }
                            ? DesignerBaseline(designerModule, designerProject)
                            : HostError("baseline needs module=<form>");

                    default:
                        return HostError(
                            $"designer action '{designerAction}' is not add, remove, set, zorder, "
                            + "liveness or baseline");
                }
            }

            case "component" when request.Query.TryGetValue("action", out var componentAction):
            {
                // Adding, renaming and removing components, from INSIDE.
                //
                // This is what a fixture is made of, and until now it was the one thing a harness
                // had to reach in through `Workbook.VBProject` for - which needs "Trust access to
                // the VBA project object model" turned on. The add-in is already past that gate:
                // the host hands it the VBE at OnConnection. So the fixture can be built through
                // the door, and the setting can stay off (2026-08-07).
                request.Query.TryGetValue("name", out var componentName);
                request.Query.TryGetValue("project", out var componentProject);
                var componentOwner = ResolveNamedProject(componentProject, out var componentUnknown)
                    ?? _shownProject;
                if (componentUnknown is not null)
                {
                    return HostError(componentUnknown);
                }

                // A STOPPED PROJECT WILL NOT TAKE COMPONENTS EITHER, and says so only as a bare
                // COM error naming nothing - the same shape as the write path's, and refused for
                // the same reason (issue #6). Adding or removing a module in break mode means
                // resetting the developer's run, which is theirs to decide.
                if (ProjectModeNow() != DesignMode)
                {
                    return HostError($"the project is stopped in the debugger, so '{componentName}' "
                        + "was not touched. Changing components now would reset it and lose the run. "
                        + "Press Reset in the editor, or POST command?name=reset, and ask again.");
                }

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
                                return HostError(
                                    $"kind '{kindText}' is not one of 1/module/standard, 2/class, 3/form");
                            }

                            // A NAME ALREADY TAKEN IS SAID IN WORDS. The editor answers a
                            // duplicate with a bare `Unexpected HRESULT`, which names neither the
                            // problem nor the module - a caller reads it as the product having
                            // broken rather than as a name it can simply change. Measured
                            // eighteen times across two chaos walks before anyone worked out the
                            // two runs had chosen the same name.
                            if (componentName is { Length: > 0 }
                                && FindComponent(componentName, componentOwner, out _) is { } taken)
                            {
                                taken.Dispose();
                                return HostError($"'{componentName}' is already a component of "
                                    + "this project; choose another name or remove that one first");
                            }

                            // The api mirrors the UI: Access's VBE has no MSForms, offers no
                            // Insert > UserForm, and would fail the Add(3) with a COM mumble.
                            // The refusal is designated instead (the owner, 2026-08-19).
                            if (kind == 3 && !HostApp.CarriesMsForms)
                            {
                                return HostError(
                                    $"{HostApp.Name} VBA has no UserForms, so kind=form cannot be added here");
                            }

                            using var project = FindProjectByDisplayName(componentProject)
                                ?? _editor.GetObject("ActiveVBProject");
                            using var components = project?.GetObject("VBComponents");
                            using var added = components?.CallObject("Add", kind);
                            if (added is null)
                            {
                                return HostError("the project would not add a component");
                            }

                            // Named here rather than left as Module1, because a fixture is its
                            // names. The editor refuses some outright - Circle is owned by the
                            // Excel object library - and says so with a bare HRESULT, so the
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
                                    // claims about - and the next run finds it and is confused by
                                    // it. Add either produces the component that was asked for or
                                    // produces nothing.
                                    try { components?.InvokeWithObject("Remove", added); }
                                    catch (Exception undo) { Log.Warn($"component: could not undo the add ({undo.GetType().Name})"); }

                                    return HostError(
                                        $"'{componentName}' was refused as a name, so nothing was added ({ex.Message.Trim()})");
                                }
                            }

                            var finalName = added.GetString("Name") ?? string.Empty;
                            Log.Info($"component: added {finalName} (kind {kind})");

                            // THE SECOND PLACE A MODULE CAN ARRIVE. The menu's insert has its own
                            // implementation of this, and hooking one of the two is how the change
                            // log came to record a module being filled with no sign of it arriving
                            // (2026-08-22). Both are told now; that there are two is its own
                            // smell, and one worth taking out separately rather than in a change
                            // about the log.
                            RecordChange(
                                finalName, componentOwner, Core.Changes.ChangeKind.Added,
                                null, ProjectReader.ReadSource(added) ?? string.Empty);

                            // The strip AND the tree. Neither republishes on its own, and they are
                            // separate publishes: the first version of this route refreshed the
                            // tabs only, so the explorer went on listing three components while
                            // the strip showed eight - a surface describing two different
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
                                return HostError("remove needs a name");
                            }

                            // THE PRODUCT'S OWN REMOVAL, not a second one written here.
                            //
                            // This used to be its own COM call: find the component, call Remove,
                            // republish. That was true of the collection and false of everything
                            // else - the page's unwritten edits were left to be flushed into a
                            // module that no longer existed, and this session went on holding a
                            // baseline and breakpoints for it. A harness removing a component saw
                            // a different machine state than a developer removing the same one
                            // from the tree, which is exactly what the api is not allowed to do.
                            var refused = RemoveComponent(componentName, componentProject);
                            if (refused is not null)
                            {
                                return HostError(refused);
                            }

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
                                return HostError("rename needs name and newName");
                            }

                            using var target = FindComponent(componentName, componentOwner, out _);
                            if (target is null)
                            {
                                return HostError($"'{componentName}' is not a component of this project");
                            }

                            target.SetString("Name", newName);
                            var readBack = target.GetString("Name") ?? newName;
                            Log.Info($"component: renamed {componentName} to {readBack}");

                            // THE SESSION'S OWN ADOPTION, for the reason `remove` above gives.
                            // Setting the name on the component renames it in the collection and
                            // nowhere else: the write baseline and the breakpoint record are both
                            // keyed by module name, so they stayed under the old one and stopped
                            // describing anything. A breakpoint the developer can see and the
                            // session can no longer find is the same defect the removal path was
                            // fixed for. AdoptRename re-keys both, moves the properties target,
                            // reloads the surface when the shown module is the one that moved,
                            // and calls ComponentsChanged itself.
                            AdoptRename(componentName, readBack);

                            return System.Text.Json.JsonSerializer.Serialize(
                                new DebugComponentReply(true, readBack, "rename"),
                                DebugJsonContext.Default.DebugComponentReply);
                        }

                        default:
                            return HostError($"unknown action {componentAction}; use add, remove or rename");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error($"component: {componentAction} failed", ex);
                    return HostError($"{componentAction} failed: {ex.Message.Trim()}");
                }
            }

            case "pane" when request.Query.TryGetValue("action", out var paneAction)
                && request.Query.TryGetValue("module", out var paneModule) && paneModule.Length > 0:
            {
                // Opening and CLOSING a module's pane.
                //
                // `caret` opens one on the way to a line, and until now nothing closed one - so
                // every test of what the tab strip does when a tab goes had to reach into the
                // page's private workspace through eval, which is a test of the probe as much as
                // of the thing. Four defects in the strip this week were found that way and each
                // one needed the reach rewritten (2026-08-07).
                request.Query.TryGetValue("project", out var paneProject);
                var paneOwner = ResolveNamedProject(paneProject, out var paneUnknown) ?? _shownProject;
                if (paneUnknown is not null)
                {
                    return HostError(paneUnknown);
                }

                switch (paneAction)
                {
                    case "open":
                        // The workbook is PASSED ON. It was computed and then dropped, so
                        // `project=` did nothing at all on open and a bare name resolved
                        // shown-project-first - meaning the second workbook's copy of a shared
                        // module name could not be opened from a script by any argument.
                        //
                        // That is why the two-workbook state was unreachable from the harness, and
                        // why every defect in this class had to be found by hand. A stress walk
                        // seeded with both workbooks' Helpers silently held only one of them and
                        // passed its label checks vacuously (2026-08-07).
                        // And its answer is the show's answer: opening a module that is not there
                        // replied ok, which is the same lie the write route told about a module
                        // that is not there (2026-08-09).
                        //
                        // `face=design` opens a form's DESIGNER TAB instead - the same method
                        // the page's own activate uses, so the api leaves the state the click
                        // would (the mirror rule). Its refusals go to the surface as notices,
                        // the way the click's do; the route reads the tab's presence back from
                        // the published strip rather than being answered twice.
                        if (request.Query.TryGetValue("face", out var paneFace) && paneFace == "design")
                        {
                            OpenDesignerTab(paneModule, DisplayFromProjectId(paneOwner));
                            var openedTab = _designerTabs.Any(tab =>
                                string.Equals(tab.Module, paneModule, StringComparison.OrdinalIgnoreCase));
                            return openedTab ? HostOk() : HostError($"{paneModule} has no designer tab to open");
                        }

                        var showed = ShowModule(paneModule, DisplayFromProjectId(paneOwner));
                        return showed is null ? HostOk() : HostError(showed);

                    case "close":
                    {
                        // Through the same gate the tab's own X uses, so a module with unwritten
                        // edits gets the question rather than the guillotine - and `action` is how
                        // a caller answers it in advance.
                        //
                        // And its answer is the close's answer, for the reason the `open` branch
                        // above gives: a save that would not save, a revert the module refused,
                        // and a confirm now standing on screen all left the tab where it was and
                        // all replied ok.
                        if (request.Query.TryGetValue("face", out var closeFace) && closeFace == "design")
                        {
                            // A designer tab has no unsaved-text question on this side; the
                            // page owns unapplied markup and asks its own.
                            CloseDesignerTab(paneModule, DisplayFromProjectId(paneOwner));
                            return HostOk();
                        }

                        request.Query.TryGetValue("answer", out var closeAnswer);
                        var closed = OnModuleCloseRequested(paneModule, DisplayFromProjectId(paneOwner), closeAnswer);
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugCloseReply(closed.Closed, closed.Detail, closed.Awaiting),
                            DebugJsonContext.Default.DebugCloseReply);
                    }

                    case "closeNative":
                    {
                        // The OTHER direction: the host's own pane window closing, which is what
                        // a developer's click on the native child's close box produces - and the
                        // direction the 2026-08-04 dead-tab defect lived in, because a HIDDEN
                        // pane's close fires no Changed and the strip kept its tab. Until this
                        // existed that path could only be produced through Application.VBE from
                        // outside, which project trust gates; in here the pane list is ours.
                        //
                        // A DESIGNATED DEVIATION from `close` above, and the deviation is the
                        // subject: the native close box asks no unwritten-edits question - that
                        // question belongs to the page's tab X - so none is asked here. The tab
                        // disappears because the tracker notices the window go, not because this
                        // route told the page anything.
                        var went = CloseThroughObjectModel(paneModule, paneOwner);
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugCloseReply(
                                went,
                                went
                                    ? "the pane's window was closed through the editor's own pane list"
                                    : $"no pane of {paneModule} exists to close",
                                null),
                            DebugJsonContext.Default.DebugCloseReply);
                    }

                    default:
                        return HostError($"unknown action {paneAction}; use open, close or closeNative");
                }
            }

            case "palette" when request.Query.TryGetValue("action", out var paletteAction):
            {
                // The palette's lifecycle, drivable. Summoning it always had a door (the
                // objectBrowser command, which means SUMMON and not toggle); putting it away had
                // none, so every probe that opened it either left it standing in front of the
                // next probe or reached for a window message from outside (2026-08-12). Hiding
                // is what the palette's own close box does - its WM_CLOSE handler hides, state
                // intact, and the next summons presents the same page.
                if (paletteAction != "hide")
                {
                    return HostError($"unknown action {paletteAction}; use hide (objectBrowser is the summons)");
                }

                if (_browserPalette is not { } paletteToHide)
                {
                    return System.Text.Json.JsonSerializer.Serialize(
                        new DebugVisibilityReply(false, "no palette exists; nothing to hide", false),
                        DebugJsonContext.Default.DebugVisibilityReply);
                }

                paletteToHide.Hide();
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugVisibilityReply(
                        true,
                        "hidden, state intact; the objectBrowser command presents it again",
                        Win32.IsWindowVisible(paletteToHide.Handle)),
                    DebugJsonContext.Default.DebugVisibilityReply);
            }

            case "frame" when request.Query.TryGetValue("action", out var frameAction):
            {
                // The editor window itself, drivable: the one pair of gestures every developer
                // makes daily and no route could produce - closing the editor and bringing it
                // back. Closing shipped a crash once (lesson 27), and until now only a probe
                // sending window messages from OUTSIDE the process could exercise it.
                if (_frame == 0)
                {
                    return HostError("the session has no editor frame to act on");
                }

                switch (frameAction)
                {
                    case "close":
                        // POSTED, not sent, and that is the design: SC_CLOSE delivered by the
                        // pump after this body returns is byte-for-byte the developer's X click,
                        // and the editor runs its whole close path - the hide, the palette
                        // follow, the placement retreat - outside this request. So `visible`
                        // still reads true in this reply; the outcome is observed on
                        // state.frameVisible, the rule every posted effect on this door lives by.
                        Win32.PostMessage(_frame, Win32.WmSysCommand, Win32.ScClose, 0);
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugVisibilityReply(
                                true,
                                "SC_CLOSE posted; observe state.frameVisible for the outcome",
                                Win32.IsWindowVisible(_frame)),
                            DebugJsonContext.Default.DebugVisibilityReply);

                    case "show":
                    {
                        // The same end state Excel's own Developer > Visual Basic button reaches:
                        // the editor's main window made visible. Synchronous, so the answer here
                        // IS the outcome.
                        using var mainWindow = _editor.GetObject("MainWindow");
                        mainWindow?.SetBool("Visible", true);
                        var showing = Win32.IsWindowVisible(_frame);
                        return System.Text.Json.JsonSerializer.Serialize(
                            new DebugVisibilityReply(
                                showing,
                                showing ? "the editor window is on screen" : "MainWindow was told to show and the frame still reads hidden",
                                showing),
                            DebugJsonContext.Default.DebugVisibilityReply);
                    }

                    default:
                        return HostError($"unknown action {frameAction}; use close or show");
                }
            }

            case "undoRename":
            {
                // The same path the editor's own Undo Rename takes. Here so a probe can prove a
                // rename is reversible without driving the page, which is the half a rename test
                // could never assert before.
                //
                // A DESIGNATED DEVIATION: this reaches the session's undo directly and does NOT go
                // through the page's "Undo Rename" context-menu action, so it proves the operation
                // is reversible and not that the menu item works. Drive the menu item itself with
                // act("editorAction", {id: "xlide.undoRename"}) when that is the question.
                //
                // What it restored comes back now. It used to answer ran:true whether it had put
                // six modules back, stopped halfway with the project in neither state, or found
                // nothing to undo - and the sentence saying which went to the page under a request
                // id this route invented, where no caller could read it.
                var undone = UndoRename(0);
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugUndoRenameReply(
                        undone.Undone, undone.From, undone.To, undone.Modules, undone.Stopped),
                    DebugJsonContext.Default.DebugUndoRenameReply);
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
                    // The current value is the fallback, which is what "name what you want
                    // changed and everything else stands" means when a value cannot be read.
                    bool Flag(string name, bool current) =>
                        AddInSession.Flag(request, name, current);

                    // FROM THE CURRENT RECORD, not a fresh one: `new ProductSettings { ... }`
                    // names the queryable fields and zeroes the rest, so changing the indent
                    // size over the door wiped the developer's api.enabled answer and the
                    // analyzer rule overrides. Same absent-means-unchanged rule as the page's
                    // updateSettings, same fix.
                    settings = (settings with
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
                        DesignerSnap = request.Query.TryGetValue("designerSnap", out var snapMode)
                            ? snapMode
                            : settings.DesignerSnap,
                        DesignerGridSize = request.Query.TryGetValue("designerGridSize", out var grid)
                            && int.TryParse(grid, out var gridAsked) ? gridAsked : settings.DesignerGridSize,
                    }).Normalized();

                    OnSettingsChanged(settings);
                }

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugSettingsReply(
                        settings.BlockLayout,
                        settings.ContinueCommentOnNewline,
                        settings.MirrorCommentSpacing,
                        settings.TreeFollowsEditor,
                        settings.FormatIndentSize,
                        settings.SyncEngine,
                        settings.DesignerSnap,
                        settings.DesignerGridSize),
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

                            // Components counted rather than listed; ForEachRealComponent keeps
                            // the scratch module out of the count and contains an unreadable
                            // entry to itself - this copy's hand-rolled walk sat inside the try
                            // wrapping the whole project, so one bad component dropped the
                            // entire project from the reply.
                            var components = 0;
                            ForEachRealComponent(project, (_, _) => components++);

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

            case "workbook" when request.Query.TryGetValue("action", out var workbookAction)
                && string.Equals(workbookAction, "close", StringComparison.OrdinalIgnoreCase):
            {
                /*
                 * THE ONE SAFE WAY TO CLOSE A WORKBOOK FROM OUTSIDE. The immediate window used
                 * to be the only route with the reach, and it is the one route that cannot
                 * survive the gesture: its line runs as a procedure IN the active project, so
                 * closing that project's workbook tears the code out from under itself and the
                 * editor dies with it - measured 2026-08-28, one try, one dead Excel (#13). The
                 * evaluator refuses those lines now and points here. This close is a plain COM
                 * call on the host thread with no VBA running, which is exactly the state the
                 * teardown is safe in.
                 *
                 * `saveChanges` is REQUIRED and explicit - the native close asks the developer
                 * this exact question, and a route that picked silently would either lose work
                 * or write a file nobody asked for. NOT named `keep`: this door reserves that
                 * word for dialog protection on every request. A DESIGNATED DEVIATION from the
                 * prompt itself: the caller answers the question in the request instead of
                 * meeting the dialog, which is what makes the route drivable at all. saveChanges
                 * on a never-saved workbook raises the host's own Save As, which the door's
                 * dialog watch then handles like any dialog this door raised.
                 */
                request.Query.TryGetValue("project", out var closeTarget);
                if (closeTarget is not { Length: > 0 })
                {
                    return HostError("workbook?action=close needs project=<name>; "
                        + "a close aimed at nothing closes nothing");
                }

                if (!(request.Query.TryGetValue("saveChanges", out var saveWord)
                    && saveWord is "0" or "1" or "true" or "false"))
                {
                    return HostError("workbook?action=close needs saveChanges=0|1 - the same "
                        + "question the native close asks, answered in the request");
                }

                var closeId = ProjectIdFromDisplay(closeTarget);
                var closeDisplay = closeId is null ? null : DisplayFromProjectId(closeId);
                using var book = closeDisplay is null ? null : FindWorkbookByDisplay(closeDisplay);
                if (book is null)
                {
                    return HostError($"no open workbook answers to '{closeTarget}'. "
                        + $"Open: {string.Join(", ", _projectNames.Values)}");
                }

                // VT_I4 on the wire: the host coerces 0/1 to SaveChanges' boolean, the same
                // road every argumented Invoke here takes.
                var saveIt = saveWord is "1" or "true";
                book.Invoke("Close", saveIt ? 1 : 0);
                Log.Info($"workbook: closed {closeDisplay}, changes {(saveIt ? "saved" : "discarded")}");

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugWorkbookReply(true, closeDisplay!, saveIt),
                    DebugJsonContext.Default.DebugWorkbookReply);
            }

            case "project":
            {
                // What is actually THERE, as opposed to what the surface is showing.
                //
                // This is the question a fixture asks twice - once to build and once to check -
                // and it was the last one that could only be answered by reaching in through
                // `Workbook.VBProject`, which needs the trust setting. Answered from inside, where
                // the add-in already is.
                request.Query.TryGetValue("project", out var wantedProject);

                // A NAMED workbook that does not resolve is refused rather than answered about the
                // active one. This route is how a fixture asks "what is in this workbook", twice,
                // and answering about a different one is how a fixture check passes against the
                // wrong project.
                if (ResolveNamedProject(wantedProject, out var projectUnknown) is null
                    && projectUnknown is not null)
                {
                    return HostError(projectUnknown);
                }

                using var project = FindProjectByDisplayName(wantedProject)
                    ?? _editor.GetObject("ActiveVBProject");

                if (project is null)
                {
                    return HostError("no VBA project is active");
                }

                var rows = new List<DebugComponentRow>();

                // The saved document's attribute headers, which is the only place the
                // default-instance flag can be read from - and therefore the only way a caller
                // can see why the analyzer did or did not report a bare class name. Read once
                // for the whole walk; null throughout for a project never saved.
                var savedHeaders = Xlide.Vbe.Core.Vba.SavedModules.For(Engine.ProjectReader.SavedPathOf(project));

                ForEachRealComponent(project, (component, name) =>
                {
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
                        open,
                        // Asked for a CLASS and nothing else, exactly as the seed asks it: a
                        // document and a form have a default instance by their kind and a
                        // standard module cannot have one, so the flag is only ever consulted
                        // here, and answering it for the others would invite reading it as one.
                        type == 2 ? savedHeaders?.PredeclaredIdOf(name) : null));
                });

                // The identity of the project THIS REPLY DESCRIBES, read off that project.
                //
                // It used to be DisplayFromProjectId(_shownProject) - the workbook the surface
                // happened to be showing - so asking about the second workbook answered with the
                // second workbook's components under the FIRST workbook's id. The reply
                // contradicted itself, and a caller doing the obvious thing (read `projectId`,
                // pass it to `pane` or `module`) was then addressing the wrong workbook while
                // holding a reply that looked right.
                //
                // That is why the two-workbook state could never be set up from a script, which
                // is why every defect in this class - navigation, tab labels, breakpoints - had
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
                    return HostError($"unknown command name {name}");
                }

                var outcome = ExecuteEditorCommand(command);
                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(outcome.Ran, command, outcome.Detail),
                    DebugJsonContext.Default.DebugCommandReply);
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

                // A NAVIGATION THAT DID NOT ARRIVE MUST NOT BE TOGGLED OVER. Everything below
                // reads and writes "the shown module", which the GoTo is what makes correct; when
                // it finds no pane the shown module is whatever was there before, so the
                // breakpoint is read from one module and set on another, and the reply names the
                // line the caller asked for either way.
                if (GoTo(module, breakLine, 1, project) is { } lost)
                {
                    return HostError(lost);
                }

                // Read against the module the GoTo above has just made the shown one, in the
                // workbook it belongs to. Keyed by name alone this read the TWIN's record when
                // two workbooks shared the module name, so `state=on` saw a breakpoint that was
                // not there and did nothing.
                var alreadySet = _editorSurface?.Module is { } shownModule
                    && BreakpointsFor(shownModule).Contains(breakLine);
                // A STATE THIS DOES NOT KNOW IS REFUSED, not quietly taken as a toggle.
                //
                // Absent means toggle, which is the documented default and what the glyph margin
                // sends. But `state=set` - a perfectly reasonable guess at the vocabulary, and one
                // this author made - used to fall into that same default, so asking twice for a
                // breakpoint to BE set left it clear, and nothing said the word had not been
                // understood. That is the trap the analyzer's own host contract names for
                // `moduleKind`: an unrecognised string that degrades silently is worse than one
                // that fails, because the caller is told it worked.
                if (wanted is { Length: > 0 } && wanted is not ("on" or "off"))
                {
                    return HostError($"'{wanted}' is not a breakpoint state; pass on, off, "
                        + "or leave it out to toggle");
                }

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

                // The record as it stands NOW, not the state that was asked for. ToggleBreakpoint
                // declines a line the editor will not carry one on - a blank line, a declaration -
                // and the decline reaches the surface as a notice, so answering `wanted` here
                // reported a breakpoint on lines that have never held one.
                var nowSet = _editorSurface?.Module is { } afterModule
                    && BreakpointsFor(afterModule).Contains(breakLine);

                return System.Text.Json.JsonSerializer.Serialize(
                    new DebugCommandReply(
                        nowSet,
                        VbeCommands.Command.ToggleBreakpoint,
                        nowSet == shouldSet
                            ? (nowSet ? "the breakpoint is set" : "the breakpoint is clear")
                            : $"the editor would not put a breakpoint on line {breakLine}"),
                    DebugJsonContext.Default.DebugCommandReply);
            }

            case "module" when request.Query.TryGetValue("name", out var moduleName) && moduleName.Length > 0:
            {
                request.Query.TryGetValue("project", out var projectDisplay);
                var projectId = ResolveNamedProject(projectDisplay, out var moduleUnknown);
                if (moduleUnknown is not null)
                {
                    return HostError(moduleUnknown);
                }

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
                    return complaint is null ? HostOk() : HostError(complaint);
                }

                // live=1 reads the SURFACE's copy rather than the workbook's.
                //
                // They differ for as long as the developer has typed and the write-back timer has
                // not fired, which is exactly the window every typing behaviour lives in: smart
                // Enter, comment continuation and auto-indent all produce text that only exists in
                // the editor until it is written. Without this there was no way to read what
                // typing produced, so those features could only be checked by eye (2026-08-08).
                if (Flag(request, "live"))
                {
                    var live = _editorSurface?.TextOf(moduleName, DisplayFromProjectId(projectId));
                    return live is null
                        ? HostError($"the surface holds no text for {moduleName}")
                        : System.Text.Json.JsonSerializer.Serialize(
                            new DebugModuleReply(moduleName, DisplayFromProjectId(projectId), live),
                            DebugJsonContext.Default.DebugModuleReply);
                }

                using var found = FindComponent(moduleName, projectId, out var foundProject);
                var source = found is null ? null : ProjectReader.ReadSource(found);
                return source is null
                    ? HostError($"no module named {moduleName}")
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
                    // A navigation that finds no pane leaves the caret in the module already
                    // shown, and this route answered ok over it. The caller's next act is
                    // usually Run, Step or a breakpoint, all of which then land somewhere the
                    // caller did not ask for and believes they did not go.
                    request.Query.TryGetValue("project", out var caretProject);
                    if (GoTo(caretModule, caretLine, caretColumn, caretProject) is { } lost)
                    {
                        return HostError(lost);
                    }
                }

                _editorSurface?.SetCaret(caretLine, caretColumn);
                return HostOk("the caret was set");
            }

            case "placement":
                RefreshSurfacePlacement();
                return HostOk();

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
                return HostError(
                    $"no route '{request.Route}' accepted this request. Either there is no such "
                    + "route, or there is and its required arguments were missing or rejected: "
                    + "many routes are guarded on theirs. "
                    + $"Given: {(request.Query.Count == 0 ? "(no arguments)" : string.Join(", ", request.Query.Select(pair => $"{pair.Key}={pair.Value}")))}");
        }
    }
}
