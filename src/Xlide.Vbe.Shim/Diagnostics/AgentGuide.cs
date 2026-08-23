namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// The api explaining itself, to a caller that has never seen it.
///
/// The audience is an AGENT: a program - usually a language model with an HTTP client - that a
/// developer has pointed at a running session and asked to work with. A person learns this door
/// from docs/xlide-api.md; an agent gets handed a discovery file and one URL, so everything the
/// docs would have told it has to be reachable BY REQUEST, as a breadcrumb trail: the discovery
/// file names the `agent` route, `agent` names the rest, and every reply carries `next` links so
/// there is no step at which the caller has to already know the answer.
///
/// ONE TABLE, THREE READERS. Each route's row is the single source for (1) the `agent/routes`
/// reply an agent reads, (2) the per-route detail `agent/route?name=` serves, and (3) the
/// in-process door's policy for that route - which is why the policy lives HERE and not in the
/// door: a route added with a row gets documentation and a door ruling in the same line, and a
/// route added without one is refused by the stricter default rather than silently opened.
///
/// The table is hand-maintained against the two route switches in AddInSession.DebugApi.cs; the
/// suite walks every row whose bare GET is harmless and fails on a row the switch no longer
/// answers.
/// </summary>
internal static class AgentGuide
{
    /// <summary>
    /// How a route answers through the IN-PROCESS door (`GetObject(, "Xlide.Api")`), which runs
    /// on the host thread with the caller's code on the stack.
    ///
    /// - <see cref="Open"/>: answers. The door deliberately does NOT gate on the host's "Trust
    ///   access to the VBA project object model" switch - the owner's call (2026-08-18): the
    ///   end-user experience must not require flipping a Trust Center setting, and the gate
    ///   protected nothing anyway, because the HTTP door beside this one has the same
    ///   capabilities, never honored that switch, and is reachable by any local code that could
    ///   have reached this door.
    /// - <see cref="HttpOnly"/>: refused inside with a pointer to the HTTP door. These routes
    ///   wait on work only a pumping host thread can finish - a page script's answer, a modal
    ///   watched from outside - and the inside caller IS the host thread, so the wait could
    ///   only ever time out.
    /// </summary>
    internal enum DoorPolicy
    {
        Open,
        HttpOnly,
    }

    /// <summary>One route, as the front door teaches it.</summary>
    /// <param name="Name">The route segment after the token.</param>
    /// <param name="Method">The CONVENTIONAL method. The server accepts GET and POST on every
    /// route; reads are written GET and acts POST so a transcript reads honestly.</param>
    /// <param name="Args">Query arguments, `body:` for a raw-text body, `?` marking optional.</param>
    /// <param name="What">What the route answers or does, one sentence.</param>
    /// <param name="Example">A literal request line against the base URL.</param>
    /// <param name="BareGetIsSafe">Whether a bare argument-less GET reads without acting, so a
    /// caller (or the suite) can try it freely.</param>
    /// <param name="Inside">How the in-process door treats it.</param>
    /// <param name="Note">The caveat a caller finds out the hard way otherwise.</param>
    internal sealed record RouteHelp(
        string Name,
        string Method,
        string Args,
        string What,
        string Example,
        bool BareGetIsSafe,
        DoorPolicy Inside,
        string? Note = null);

    internal static readonly RouteHelp[] Routes =
    [
        // The trail itself.
        new("agent", "GET", "-",
            "The front door: what this api is, which Office host it is running in, and where to go next.",
            "GET agent", true, DoorPolicy.Open),
        new("agent/routes", "GET", "-",
            "Every route: arguments, what it does, an example, and how the in-process door treats it.",
            "GET agent/routes", true, DoorPolicy.Open),
        new("agent/route", "GET", "name=<route>",
            "One route in detail.",
            "GET agent/route?name=module", true, DoorPolicy.Open),
        new("agent/examples", "GET", "-",
            "Runnable recipes: ordered request lines for the common jobs, from reading a module to a breakpoint round trip.",
            "GET agent/examples", true, DoorPolicy.Open),

        // What the product knows. Product knowledge, not project state - open to any caller.
        new("model", "GET", "type=<name>?",
            "What the language service knows about this host's object model: the type inventory bare, or one type's members - name it as code writes it (Worksheet) or qualified (Excel.Worksheet).",
            "GET model?type=Worksheet", true, DoorPolicy.Open,
            "In a host with no model wired yet the answer is known:false with a note - the same "
            + "honesty the language features practice, silence over another host's members."),
        new("analyzer", "GET", "-",
            "The analyzer's rule catalogue: code, title, severity, category, evidence kind, confidence, and the MS-VBAL authority where one applies.",
            "GET analyzer", true, DoorPolicy.Open),

        // Orientation and health.
        new("state", "GET", "-",
            "The session at a glance: shown module and project, debug mode, unwritten edits, whether the engine answers, the editor frame.",
            "GET state", true, DoorPolicy.Open),
        new("sessions", "GET", "-",
            "Every live session on this machine: pid, host, and which one is answering. The fleet the inside door's @ prefix addresses - GetObject(, \"Xlide.Api\").Request(\"@word/state\") answers for Word from whichever host holds the name.",
            "GET sessions", true, DoorPolicy.Open,
            "Ports, tokens and agent urls are deliberately absent: one door's caller is not "
            + "handed the keys to every other door. Address a peer with @pid or @host through "
            + "the inside door, or read its own discovery file."),
        new("tests", "GET", "action, module, test, file",
            "The VBA test runner over EVERY open file: '@xlide-test' directives over zero-argument Subs in standard modules, the XlideAssert latched-failure protocol, results per test. No action lists, answering a files array - each file's own support state and test count - and a row per test carrying the file it came from. file= narrows any verb to one open file, which a module name cannot do on its own because it is only unique inside its own file. action=install writes/updates XlideAssert in one file with file= or in every file that holds tests without it; action=run runs everything everywhere, run&file= one file, run&module= that module wherever it is, run&test= one test; action=runFailed reruns what failed, narrowed by module= or file=; action=debug&test= runs one test untrapped so breakpoints and errors drop into the debugger. ranAt says when the last run finished.",
            "GET tests?action=run", true, DoorPolicy.Open,
            "run answers when the whole run has finished - the Tests pane streams the same "
            + "results live. XlideAssert is a module INSIDE a file, so support is per file: a "
            + "run refuses for the file that lacks it while the file that has it still runs. "
            + "debug answers when the debug session ends. A test's timeout "
            + "metadata is carried but not enforced: in-process VBA cannot be preempted, and "
            + "Ctrl+Break remains the escape, exactly as with the native F5."),
        new("doctor", "GET", "-",
            "Staleness and health findings, chief among them whether the running shim, page and engine are the ones last built.",
            "GET doctor", true, DoorPolicy.Open),
        new("stats", "GET", "-",
            "Counters: host-thread heartbeat, marshal costs, per-route costs, COM wrapper balance.",
            "GET stats", true, DoorPolicy.Open),
        new("perf", "GET", "reset=1?",
            "The performance counters the session keeps about itself. `reset=1` forgets the "
            + "analyzer figures first.",
            "GET perf", true, DoorPolicy.Open,
            "Reset before provoking something, so the figures measure what you did rather than "
            + "everything since the editor opened - session start is the wrong window for asking "
            + "whether THIS change is slow."),
        new("log", "GET", "since=<tick>? match=<text>? max=<n>? waitMs=<ms>?",
            "The shim log, sliced; with match+waitMs it returns the moment a matching line is written.",
            "GET log?match=modules:%20publish&waitMs=10000", true, DoorPolicy.Open,
            "The log grows from every thread, so waiting on it works from inside too."),
        new("journal", "GET", "lines=<n>?",
            "One-call evidence bundle: state, standing dialogs, counters, recent log, recent page traffic.",
            "GET journal", true, DoorPolicy.Open),
        new("history", "GET", "-",
            "The door's last 300 state-changing requests, plus a runnable script of them.",
            "GET history", true, DoorPolicy.Open),
        new("mark", "POST", "text=<marker>",
            "Writes a marker line into the shim log, so a caller can bracket its own experiment.",
            "POST mark?text=before-rename", true, DoorPolicy.Open),

        // Projects, modules, code.
        new("projects", "GET", "-",
            "Every open VBA project.",
            "GET projects", true, DoorPolicy.Open),
        new("project", "GET", "project=<display>?",
            "One project in detail; the shown one when unnamed.",
            "GET project", true, DoorPolicy.Open),
        new("documents", "GET", "-",
            "The modules the EDITOR holds text for, one row each: module, project, lines, whether "
            + "it has unwritten edits, and whether it is active. Not the open workbooks - `projects` "
            + "answers those. An empty answer means no module has been activated yet, which is the "
            + "ordinary state of a session nobody has clicked into.",
            "GET documents", true, DoorPolicy.Open,
            "A module with a TAB and no text is the state most of a workspace is in, so this list "
            + "is shorter than the tab strip and shorter still than the project. Reading it as a "
            + "list of workbooks makes an empty answer look like the api cannot see an open file."),
        new("module", "GET|POST", "name=<module> project=<display>? live=1? body:<source on POST>",
            "Reads a module's code, or replaces it with the POSTed body. `live=1` reads the "
            + "EDITOR's copy instead of the workbook's.",
            "GET module?name=Module1", true, DoorPolicy.Open,
            "The two copies differ for as long as the developer has typed and the write-back has "
            + "not fired, which is the window every typing behaviour lives in: smart Enter, comment "
            + "continuation and auto-indent all produce text that exists only in the editor until "
            + "it is written. Without `live=1` there is no way to read what typing produced."),
        new("outline", "GET", "module=<name> project=<display>?",
            "A module's procedures from the analyzer: name, kind, line.",
            "GET outline?module=Module1", true, DoorPolicy.Open),
        new("engine", "GET", "module=<name> text=1?",
            "The analysis engine's copy of a module's source, for comparing against the host's. "
            + "`text=1` returns both texts rather than only their shapes.",
            "GET engine?module=Module1", true, DoorPolicy.Open,
            "The question this answers is which side drifted: a finding was once seen six columns "
            + "out after a format, with no way to ask whether the engine or the surface had moved."),
        new("problems", "GET", "module=<name>?",
            "The analyzer's findings, optionally for one module.",
            "GET problems", true, DoorPolicy.Open),
        new("component", "POST", "action=add|remove|rename name=<component> kind=<kind on add> newName=<on rename> project=<display>?",
            "Adds, removes or renames a project component the way the editor would. kind=form is refused in Access, whose VBA has no UserForms.",
            "POST component?action=add&name=Helper&kind=module", true, DoorPolicy.Open),
        new("compile", "POST", "-",
            "Runs Debug > Compile and answers with the error dialog's own words, or that it compiled cleanly.",
            "POST compile", false, DoorPolicy.HttpOnly,
            "Starts the command and reads its modal from the pool thread; inside callers hold the very thread that must pump."),
        new("sync", "GET|POST", "action=plan|apply direction=import|export folder=<path> project=<display>? mode=? select=?",
            "Import and export through the same service the sync dialog uses: plan first, then apply.",
            "GET sync?action=plan&direction=export&folder=C:%5Ctemp%5Csrc", false, DoorPolicy.Open),

        // The editor surface.
        new("windows", "GET", "-",
            "Every VBE window per workbook: panes, designers, captions, visibility.",
            "GET windows", true, DoorPolicy.Open),
        new("menus", "GET", "path=<Menu%20Path>?",
            "The editor's menu tree, or one menu's items with their enabled state.",
            "GET menus", true, DoorPolicy.Open),
        new("native", "GET", "text=1?",
            "The native code panes as the HOST holds them: which modules, which is active, and "
            + "where the caret is. `text=1` returns the pane's text and the surface's beside it.",
            "GET native", true, DoorPolicy.Open,
            "Run, Step, Compile and ToggleBreakpoint act on the native ACTIVE PANE and the caret "
            + "inside it - not on what the page is showing. A page showing one module while the "
            + "active pane is another is a Run that executes elsewhere, with nothing on screen to "
            + "say so."),
        new("pane", "POST", "action=open|close|closeNative module=<name> project=<display>? face=design? answer=save|discard?",
            "Opens or closes an editor tab for a module; face=design opens a form's designer tab.",
            "POST pane?action=open&module=Module1", true, DoorPolicy.Open),
        new("frame", "POST", "action=close|show",
            "The editor window itself: close posts the developer's X click, show is the Developer-tab summons.",
            "POST frame?action=show", true, DoorPolicy.Open),
        new("palette", "POST", "action=hide",
            "Hides the palette window, state intact; the objectBrowser command is the summons.",
            "POST palette?action=hide", true, DoorPolicy.Open),
        new("placement", "POST", "-",
            "Re-applies the remembered window placement to the editor frame.",
            "POST placement", true, DoorPolicy.Open),
        new("settings", "GET|POST", "blockLayout=? formatIndentSize=? syncEngine=? designerSnap=? designerGridSize=?",
            "Reads the product settings, or writes the ones named.",
            "GET settings", true, DoorPolicy.Open),
        new("command", "POST", "name=<command>",
            "Runs a named editor command at the caret - run, save, stepInto and the rest of the command table.",
            "POST command?name=save", true, DoorPolicy.Open,
            "name=run stops answering until the breakpoint hits - the timeout is normal; poll "
            + "state for debugMode=break. From inside, a macro is already on the stack, so the "
            + "host refuses run and step; save and friends work."),
        new("undoRename", "POST", "-",
            "Reverts the last rename through the session's own undo, and says what it put back.",
            "POST undoRename", false, DoorPolicy.Open),

        // The debugger.
        new("breakpoint", "POST", "module=<name> line=<n> state=on|off project=<display>?",
            "Sets or clears a breakpoint. state is explicit so a retry cannot toggle away what the first call set.",
            "POST breakpoint?module=Module1&line=8&state=on", true, DoorPolicy.Open),
        new("caret", "POST", "line=<n> column=<n>? module=<name>? project=<display>?",
            "Places the caret, which is what every editor command acts on; aim it before command?name=run.",
            "POST caret?module=Module1&line=10", true, DoorPolicy.Open),
        new("breakpoints", "GET", "-",
            "Every breakpoint the session knows.",
            "GET breakpoints", true, DoorPolicy.Open),
        new("locals", "GET", "-",
            "The Locals rows at the current break, from the ghost reader's snapshot - never touches the break itself.",
            "GET locals", true, DoorPolicy.Open),
        new("watches", "GET", "-",
            "The Watch rows, same snapshot discipline as locals.",
            "GET watches", true, DoorPolicy.Open),
        new("immediate", "POST", "text=<line, URL-encoded> waitMs=<ms>?",
            "Runs a line in the Immediate window and answers what it came to, error box included. Without text it READS the Immediate window instead.",
            "POST immediate?text=?2%2B2", false, DoorPolicy.HttpOnly,
            "The line rides the QUERY, not the body. Inside callers are already executing VBA; run the expression yourself instead."),

        // Dialogs and safety.
        new("dialogs", "GET", "-",
            "Standing dialogs with their captions, their TEXT, and their buttons; answers even while a modal owns the host thread.",
            "GET dialogs", true, DoorPolicy.Open),
        new("dismiss", "POST", "button=<label> caption=<title>?",
            "Presses a named button on a standing dialog.",
            "POST dismiss?button=Cancel", true, DoorPolicy.Open),
        new("guard", "GET|POST", "on=1|0? forget=1?",
            "The dialog guard: while on, notices the door did not raise are cleared too, and listed here.",
            "GET guard", true, DoorPolicy.Open),
        new("assert", "GET", "that=<claim> timeoutMs=<ms>? value=<expected>?",
            "Waits for a named claim - stopped, running, surfaceReady, shownModule, nativeModule, noDialogs, localsHas, watchHas, problemFree, responsive - and answers what it SAW.",
            "GET assert?that=problemFree&timeoutMs=5000", true, DoorPolicy.Open,
            "From inside, host state cannot change while you hold the thread, so the inside "
            + "door defaults timeoutMs to 0 - one look, answered now. Name a timeout there "
            + "only if you mean to hold the host while it burns."),
        new("userform", "GET|POST", "action=close? caption=<title>?",
            "The running (shown) forms' windows; action=close closes one.",
            "GET userform", true, DoorPolicy.Open),
        new("session", "POST", "action=cancelledShutdown",
            "Lifecycle experiments: tears the session down and lets the watchdog revive it, as a cancelled host shutdown would.",
            "POST session?action=cancelledShutdown", false, DoorPolicy.HttpOnly,
            "It stops the very session the inside caller is standing in."),
        new("drainfinalizers", "POST", "-",
            "Runs a garbage collection and drains the COM finalizer queue, for leak accounting.",
            "POST drainfinalizers", false, DoorPolicy.Open),

        // The form designer.
        new("designer", "GET|POST", "module=<form> project=<display>? format=markup? action=applyMarkup|add|remove|set|zorder|autosize|liveness|baseline + per-action args",
            "The form designer: GET reads the design (format=markup for the dialect), actions edit it - add(type,name?,parent?,left,top,width,height), remove(name), set(name,property,value,as?), zorder(name,to), autosize(name).",
            "GET designer?module=UserForm1&format=markup", true, DoorPolicy.Open),
        new("defaults", "GET", "type=<control kind>",
            "What a control of a kind holds untouched, measured from a bare instance - the designer's baseline for what a developer changed.",
            "GET defaults?type=CommandButton", true, DoorPolicy.Open),
        new("vocabulary", "GET", "module=<form>? project=<display>?",
            "The form markup language's whole vocabulary, exactly as the designer tab's completions hold it.",
            "GET vocabulary", true, DoorPolicy.Open),

        // The page (the editor's web surface). All of these ride a page script round trip.
        new("ui", "GET", "line=? column=? word=?",
            "The page's own model state, by the page's own report.",
            "GET ui", true, DoorPolicy.HttpOnly),
        new("act", "POST", "do=<action> + action's args as query",
            "Drives the surface through the methods a click reaches - tabs, menus, gestures - by name.",
            "POST act?do=editorAction&id=xlide.undoRename", false, DoorPolicy.HttpOnly),
        new("eval", "POST", "body:<script> surface=palette?",
            "Runs script in the page and answers its value; promises are awaited.",
            "POST eval  (body: document.title)", false, DoorPolicy.HttpOnly),
        new("await", "POST", "body:<predicate> waitMs=<ms>? surface=?",
            "Polls a predicate IN the page until true; answers met and elapsedMs.",
            "POST await?waitMs=5000  (body: !!window.xlideBridge)", false, DoorPolicy.HttpOnly),
        new("console", "GET", "last=<n>?",
            "The page's console ring - the warnings and handled errors that never reach the shim log.",
            "GET console", true, DoorPolicy.HttpOnly),
        new("inspect", "GET", "selector=<css> styles=<a,b>? rules=1? max=<n>?",
            "The elements a selector matches: boxes, classes, computed styles, and with rules=1 the CSS rules that claim them.",
            "GET inspect?selector=.dock-split&styles=display", false, DoorPolicy.HttpOnly),
        new("capture", "GET", "selector=<css>? pad=<px>? window=? caption=?",
            "A screenshot, cropped to an element when a selector is given. Answers PNG bytes, not JSON.",
            "GET capture?selector=.designer-canvas", false, DoorPolicy.HttpOnly),
        new("layout", "GET|POST", "reset=1 to restore",
            "The whole visible arrangement - docks, groups, tabs, sizes; POST layout?reset=1 puts it back.",
            "GET layout", true, DoorPolicy.HttpOnly),
        new("type", "POST", "text=<keys>|body",
            "Types into the surface as a person would, through the page's input path.",
            "POST type?text=Sub%20Demo", false, DoorPolicy.HttpOnly),
        new("bench", "GET", "what=tabswitch|layout|type n=<runs>?",
            "Times a named page scenario: min, median, p95, max, and the samples.",
            "GET bench?what=tabswitch&n=40", false, DoorPolicy.HttpOnly),
        new("trip", "GET", "what=pagecall n=<runs>?",
            "Times the door's own round trips, so a slow route cannot masquerade as a slow feature.",
            "GET trip?what=pagecall", false, DoorPolicy.HttpOnly),
        new("reload", "POST", "waitMs=<ms>?",
            "Reloads the page and WAITS for it to come back, answering the build stamp it is now running.",
            "POST reload", false, DoorPolicy.HttpOnly),
        new("messages", "GET", "last=<n>?",
            "The recent page-to-shim bridge traffic.",
            "GET messages", true, DoorPolicy.Open),
    ];

    /// <summary>The row for a route, matched on the name column's route segment.</summary>
    internal static RouteHelp? Find(string name)
    {
        foreach (var row in Routes)
        {
            if (string.Equals(row.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return row;
            }
        }

        return null;
    }

    /// <summary>
    /// The in-process door's ruling for a route. A route without a row is HTTP-ONLY, not open:
    /// fail toward the door that works for everything, so forgetting a row inconveniences an
    /// inside caller but can never hand one a wait that cannot finish.
    /// </summary>
    internal static DoorPolicy PolicyOf(string route)
    {
        if (route.StartsWith("agent", StringComparison.OrdinalIgnoreCase))
        {
            return DoorPolicy.Open;
        }

        return Find(route)?.Inside ?? DoorPolicy.HttpOnly;
    }

    /// <summary>The `agent` reply: identity, orientation, and the trail onward.</summary>
    internal static DebugAgentReply FrontDoor(
        string baseUrl, string host, int pid, bool engineUp, string shimBuiltUtc)
    {
        return new DebugAgentReply(
            Product: "xlide_vbide",
            What: "A VBA IDE that lives inside the Office VBA editor: a modern editor surface, "
                + "a language service (completions, hover, diagnostics, navigation), module sync, "
                + "a debugger bridge, and a UserForm designer. This api is a local HTTP door into "
                + "the one RUNNING session that answered this request.",
            Host: host,
            Pid: pid,
            ShimBuiltUtc: shimBuiltUtc,
            EngineUp: engineUp,
            BaseUrl: baseUrl,
            HowToCall: "GET or POST {baseUrl}/{route}?args - the server accepts either method on "
                + "every route; the documented method is the convention. POST bodies are raw text "
                + "(module source, page scripts). Every reply is JSON except capture, which is a "
                + "PNG. There is no pagination and no auth beyond the token already in the base URL.",
            InsideDoor: "Code already running on this machine - a workbook's VBA, or any "
                + "automation client - reaches the same routes without HTTP: "
                + "GetObject(, \"Xlide.Api\").Request(\"route?args\"[, body]) answers the same "
                + "JSON, and .Guide answers this very reply. No Trust Center setting is "
                + "required. With several Office instances running, GetObject binds one of "
                + "them; check pid in Request(\"agent\"). Routes marked httpOnly in the table "
                + "answer only at this HTTP door.",
            StartHere:
            [
                "GET agent/routes for the route table.",
                "GET state to see what the session is showing.",
                "GET projects for the open workbooks, and GET documents for the modules the editor is holding text for - two different questions.",
                "GET module?name=<module> to read code; POST the body back to write it.",
                "GET model and GET analyzer for what the language service knows: the host's object model and the diagnostic rules.",
                "GET agent/examples for full recipes, including the breakpoint round trip.",
            ],
            Next: Trail(baseUrl));
    }

    /// <summary>The `agent/routes` reply.</summary>
    internal static DebugAgentRoutesReply RouteTable(string baseUrl)
    {
        return new DebugAgentRoutesReply(
            Count: Routes.Length,
            MethodNote: "The server accepts GET and POST on every route; the method column is "
                + "the convention that keeps a transcript honest. Routes whose bareGetIsSafe is "
                + "true read without acting and can be tried freely.",
            InsideDoorNote: "insideDoor says how each route answers through the in-process door "
                + "(GetObject(, \"Xlide.Api\") from VBA or GetActiveObject(\"Xlide.Api\") from "
                + "automation): open answers there; httpOnly is refused there because it waits "
                + "on a pumping host thread and answers only at this HTTP door. No Trust Center "
                + "setting is required for either door.",
            Routes: [.. Routes.Select(row => new DebugAgentRouteRow(
                row.Name, row.Method, row.Args, row.What, row.Example,
                row.BareGetIsSafe, PolicyWord(row.Inside), row.Note))],
            Next: Trail(baseUrl));
    }

    /// <summary>The `agent/route?name=` reply, or null when the name has no row.</summary>
    internal static DebugAgentRouteRow? OneRoute(string name)
    {
        return Find(name) is { } row
            ? new DebugAgentRouteRow(
                row.Name, row.Method, row.Args, row.What, row.Example,
                row.BareGetIsSafe, PolicyWord(row.Inside), row.Note)
            : null;
    }

    /// <summary>The `agent/examples` reply: the common jobs as ordered request lines.</summary>
    internal static DebugAgentExamplesReply Examples(string baseUrl)
    {
        return new DebugAgentExamplesReply(
            Examples:
            [
                new DebugAgentExample(
                    "See what you are working with",
                    [
                        "GET agent            - who am I talking to, which Office host",
                        "GET state            - what the editor is showing",
                        "GET documents        - the open workbooks",
                        "GET projects         - the open VBA projects",
                    ],
                    null),
                new DebugAgentExample(
                    "Read and write a module",
                    [
                        "GET module?name=Module1              - the code, as the host holds it",
                        "POST module?name=Module1  (body: the new source)",
                        "GET problems?module=Module1          - what the analyzer thinks of it",
                    ],
                    "POST module writes an EXISTING module; component?action=add creates one. "
                    + "Add project=<display name> to either call when two workbooks are open."),
                new DebugAgentExample(
                    "Create a module, prove the edit landed everywhere, clean up",
                    [
                        "POST component?action=add&name=Scratch&kind=module",
                        "POST module?name=Scratch  (body: the source)",
                        "GET module?name=Scratch              - the host agrees",
                        "GET engine?module=Scratch            - the analyzer agrees",
                        "POST component?action=remove&name=Scratch",
                    ],
                    null),
                new DebugAgentExample(
                    "Read and edit a form's design",
                    [
                        "GET designer?module=UserForm1&format=markup   - the design as markup",
                        "POST designer?module=UserForm1&action=set&name=OkButton&property=Caption&value=Go",
                        "GET designer?module=UserForm1                 - the control tree agrees",
                    ],
                    "POST designer?action=applyMarkup with a markup body replaces the whole design."),
                new DebugAgentExample(
                    "Run to a breakpoint and look around",
                    [
                        "POST module?name=Demo  (body: a Sub with a few locals)",
                        "POST breakpoint?module=Demo&line=4&state=on",
                        "POST caret?module=Demo&line=2        - commands act on the caret",
                        "POST command?name=run                - EXPECT a timeout: the host stops answering until the break",
                        "GET state                            - poll until debugMode=break",
                        "GET locals",
                        "POST command?name=reset",
                    ],
                    "The command?name=run timeout is normal, not a failure; the break is the answer."),
                new DebugAgentExample(
                    "When something looks stuck",
                    [
                        "GET dialogs          - a modal owns the editor until somebody answers it; this route answers regardless",
                        "GET doctor           - staleness and health findings",
                        "GET journal          - one-call evidence bundle for a bug report",
                    ],
                    null),
            ],
            Next: Trail(baseUrl));
    }

    private static string PolicyWord(DoorPolicy policy) =>
        policy == DoorPolicy.Open ? "open" : "httpOnly";

    /// <summary>The breadcrumbs every agent reply carries.</summary>
    private static Dictionary<string, string> Trail(string baseUrl) => new(StringComparer.Ordinal)
    {
        ["routes"] = $"{baseUrl}/agent/routes",
        ["examples"] = $"{baseUrl}/agent/examples",
        ["objectModel"] = $"{baseUrl}/model",
        ["analyzerRules"] = $"{baseUrl}/analyzer",
        ["state"] = $"{baseUrl}/state",
        ["doctor"] = $"{baseUrl}/doctor",
    };
}
