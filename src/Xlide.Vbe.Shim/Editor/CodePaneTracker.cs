using System.Runtime.InteropServices;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>One code pane window the editor is currently showing.</summary>
/// <param name="Window">Handle of the pane window.</param>
/// <param name="Component">Component the pane is showing, or null when it could not be matched.</param>
/// <param name="Bounds">Screen rectangle of the pane.</param>
/// <param name="IsVisible">Whether the pane is currently shown.</param>
/// <param name="Project">
/// Identity of the project the component belongs to, or null when it could not be told apart -
/// a caption names only the component, so when two projects share the component's name the
/// pane's owner is ambiguous from here.
/// </param>
internal readonly record struct CodePane(
    nint Window,
    string? Component,
    PixelRect Bounds,
    bool IsVisible,
    string? Project = null);

/// <summary>A component that has a pane open, with the identity of the project that owns it.</summary>
internal readonly record struct OpenComponent(string Name, string? Project);

/// <summary>
/// Keeps an up to date picture of which code panes exist, where they are, and what they show.
///
/// This is the map an editor surface is positioned by, so it has to be correct rather than
/// approximately correct. Two sources are combined because neither is sufficient alone. Window
/// enumeration finds the pane windows and their rectangles, which the object model does not expose.
/// The object model lists which components actually have panes open, which the windows cannot say
/// reliably, because a pane window shares its class with the Immediate window and its caption is
/// localised.
///
/// Everything here runs on the host user interface thread.
/// </summary>
internal sealed class CodePaneTracker : IDisposable
{
    /// <summary>Window class the editor uses for code panes, and for the Immediate window.</summary>
    internal const string PaneClass = "VbaWindow";

    /// <summary>Window class of the editor's own frame.</summary>
    internal const string FrameClass = "wndclass_desked_gsk";

    private static CodePaneTracker? _enumerating;

    private readonly DispatchObject _editor;
    private readonly List<CodePane> _panes = [];
    private readonly List<nint> _candidates = [];

    private WindowEventHook? _hook;
    private bool _refreshing;

    /// <summary>
    /// Component names the editor last reported as having a pane open.
    ///
    /// Reading this costs several calls into the object model, and a single splitter drag raises a
    /// location change for every affected window on every mouse move. Re-reading it per event would
    /// turn one gesture into hundreds of cross-process calls on the thread the user is typing on.
    /// It changes only when a pane opens or closes, so it is cached and invalidated by exactly
    /// those events.
    /// </summary>
    private List<OpenComponent>? _openComponents;

    public CodePaneTracker(DispatchObject editor) => _editor = editor;

    /// <summary>Panes as of the last refresh.</summary>
    public IReadOnlyList<CodePane> Panes => _panes;

    /// <summary>Raised after the set of panes, or any pane's position, has changed.</summary>
    public event Action<IReadOnlyList<CodePane>>? Changed;

    /// <summary>
    /// Raised when a refresh could not be completed because the editor is not answering, which
    /// happens while it runs the developer's code. Whoever owns a timer should keep one running
    /// and retry, because the window events that normally drive refreshes stop mattering the
    /// moment they all fail, and the picture stays wrong until something asks again.
    /// </summary>
    public event Action? RefreshFailed;

    /// <summary>True while the picture is known to be out of date because refreshes are failing,
    /// or because a burst of events outran the trailing passes and its tail was dropped.</summary>
    public bool Stale => _failedRefreshes > 0 || _refreshDropped;

    /// <summary>Set when a burst outran the trailing passes; the poll's retry clears it.</summary>
    private bool _refreshDropped;

    /// <summary>Consecutive refreshes the editor has refused, for recovery and log discipline.</summary>
    private int _failedRefreshes;

    public void Start()
    {
        _hook = WindowEventHook.Install(OnWindowEvent);

        // A Toolbox shown BEFORE the hook armed fires no event this session will ever hear:
        // a workbook saved with designer state restores it during load, and a host that
        // loads the add-in late leaves that show in the past. One Win32 sweep at arm time
        // closes the gap - the object model's Visible cannot, for this species (measured
        // 2026-08-14: route-end OM sweeps ran dozens of times while a restored Toolbox
        // stood).
        foreach (var standing in VisibleMiniFrames())
        {
            Log.Info($"code panes: putting the boot toolbox down {standing:X}");
            Win32.ShowWindow(standing, 0);
        }

        Refresh();
    }

    /// <summary>
    /// Raised on any window destruction in the process. A dying pane cannot be told from a
    /// dying tooltip here - destruction strips a window of its class name before the event
    /// arrives - so the listener treats every destroy as "a pane may have closed" and
    /// re-derives cheaply for a moment. The need: the tracker only ever holds the pane
    /// windows it can match, the active one in practice, so a HIDDEN pane's close changes
    /// nothing in its picture and Changed stays silent while the strip shows a dead tab.
    /// </summary>
    public Action? WindowDestroyed { get; set; }

    /// <summary>
    /// Raised when the editor's own frame or its document area moves or resizes.
    ///
    /// This is a different fact from <see cref="Changed"/>, and the difference bit: Changed
    /// fires only when the PANE list differs, so an editor with no visible pane - the empty
    /// workspace - resized in silence, and the surface sat at its old size while the window
    /// grew around it (the developer's report, 2026-08-04). The frame's own window events are
    /// the one signal that exists in every state.
    /// </summary>
    public Action? FrameChanged { get; set; }

    /// <summary>
    /// A window was renamed, with its handle. The editor renames its own main window as the active
    /// project and module change, which is the only notice anything holding that title bar gets.
    /// </summary>
    public Action<nint>? CaptionChanged { get; set; }

    /// <summary>
    /// The editor's main window in this process, or zero before it exists. Found by class and
    /// filtered by process, because the class is not unique to us: a second host running the same
    /// editor has one too, and dressing someone else's window would be a bug nobody could explain.
    /// </summary>
    public static nint MainWindow()
    {
        var ours = Win32.GetCurrentProcessId();

        nint frame = 0;
        while ((frame = Win32.FindWindowEx(0, frame, FrameClass, null)) != 0)
        {
            Win32.GetWindowThreadProcessId(frame, out var owner);
            if (owner == ours)
            {
                return frame;
            }
        }

        return 0;
    }

    /// <summary>One class the editor's floating palettes wear.</summary>
    private const string FloatingPaletteClass = "VBFloatingPalette";

    /// <summary>
    /// The OTHER class a palette wears: the Office tool-window frame, "F3 MinFrame" plus a
    /// per-process hex suffix, so it is matched by prefix. The Toolbox restored by a
    /// workbook saved with designer state shows as THIS species at session boot (measured
    /// 2026-08-14, the owner's fourth report: the standing window's class, read from
    /// outside, was "F3 MinFrame 8a860000" - and every defense before this one, plus the
    /// verification enums that declared them working, filtered on VBFloatingPalette).
    /// </summary>
    private const string MiniFramePrefix = "F3 MinFrame";

    /// <summary>
    /// The Toolbox window's handle, resolved from the object model the first time a palette
    /// shows and re-resolved whenever a shown palette does not match it. Type 10 in the
    /// editor's window-type enumeration, matched by HANDLE because captions are localised.
    /// </summary>
    private nint _toolboxWindow;

    private bool IsTheToolbox(nint window) =>
        window != 0 && (window == _toolboxWindow || window == ResolveToolboxWindow());

    /*
     * The boot sweep's enumeration callback must be static and unmanaged, so its hits go
     * into a static list behind a gate - the dialog watch's own pattern. Collection is the
     * only thing done under the gate.
     */
    private static readonly Lock MiniFrameGate = new();
    private static readonly List<nint> MiniFrameHits = [];

    /// <summary>Every visible top-level MinFrame this process owns - native tool-window
    /// frames, none of which is ever this product's.</summary>
    private static unsafe nint[] VisibleMiniFrames()
    {
        lock (MiniFrameGate)
        {
            MiniFrameHits.Clear();
            Win32.EnumWindows((nint)(delegate* unmanaged<nint, nint, int>)&OnTopLevelMiniFrame, 0);
            return [.. MiniFrameHits];
        }
    }

    [UnmanagedCallersOnly]
    private static int OnTopLevelMiniFrame(nint window, nint parameter)
    {
        Win32.GetWindowThreadProcessId(window, out var owner);
        if (owner == Win32.GetCurrentProcessId()
            && Win32.IsWindowVisible(window)
            && Win32.ReadClassName(window).StartsWith(MiniFramePrefix, StringComparison.Ordinal))
        {
            MiniFrameHits.Add(window);
        }

        return MiniFrameHits.Count < 32 ? 1 : 0;
    }

    /// <summary>Reads the Toolbox's handle off the object model and remembers it. Zero when
    /// the editor will not answer - the next asker resolves again.</summary>
    private nint ResolveToolboxWindow()
    {
        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;
            for (var i = 1; i <= count; i++)
            {
                using var candidate = windows!.GetItem(i);
                if (candidate?.GetInt32("Type") == 10)
                {
                    _toolboxWindow = candidate.GetInt32("HWnd");
                    return _toolboxWindow;
                }
            }
        }
        catch
        {
            // An editor that will not answer keeps its palette this once.
        }

        return 0;
    }

    private void OnWindowEvent(WindowEvent windowEvent)
    {
        // Ahead of the layout gate, which does not admit renames: the editor rewrites its own
        // caption whenever the active project or module changes, and that is exactly when anything
        // that has taken the title bar over has to take it over again.
        if (windowEvent.IsNameChange && !windowEvent.IsCaret)
        {
            CaptionChanged?.Invoke(windowEvent.Window);

            /*
             * A RENAMED COMPONENT KEEPS ITS PANE, so the cache below never hears about it.
             *
             * That cache holds the components with panes open and is invalidated when a pane
             * appears or disappears, on the reasoning that nothing else can change the set. A
             * rename changes what a component is CALLED without touching a window, and the names
             * are what a pane's caption is matched against: a stale one matches nothing, and an
             * unmatched pane is dropped rather than guessed at, which loses the tab.
             *
             * Measured before changing it, and it did not reproduce: renaming a module with its
             * pane shown, in the background, with references and without, all followed correctly,
             * because something else in the rename path happens to invalidate this. That is luck
             * to depend on, and the same shape of assumption cost a day elsewhere this week, so
             * the event that actually corresponds to the change now clears it.
             *
             * Only for the editor's own windows. The host rewrites its frame caption constantly,
             * and invalidating on every one of those would re-read the component list several
             * times a second to notice nothing.
             */
            if (IsEditorClass(Win32.ReadClassName(windowEvent.Window)))
            {
                _openComponents = null;
            }
        }

        if (!windowEvent.AffectsLayout || windowEvent.IsCaret)
        {
            return;
        }

        var className = Win32.ReadClassName(windowEvent.Window);

        // THE TOOLBOX NEVER SHOWS - the owner's rule, restated four times before every hole
        // was found (2026-08-13 twice, 2026-08-14 twice). It arrives on POSTED messages
        // after every synchronous put-down this product makes, and at SESSION BOOT when a
        // workbook saved with designer state restores it - so the put-down lives here, in
        // the event that showed it, for BOTH of its window species. An Office MinFrame (the
        // species the boot restore uses, which every earlier defense filtered out by class)
        // goes down UNCONDITIONALLY: every native tool window this product replaces wears
        // that frame and none of ours ever does, and the object model's Visible cannot even
        // touch it (measured 2026-08-14: route-end sweeps ran dozens of times while a
        // restored Toolbox stood). A VBFloatingPalette goes down only when it IS the
        // Toolbox, matched by HANDLE off the object model - Type 10, never the localised
        // caption - because the Watches and Locals ghosts share that class.
        if (windowEvent.IsShow
            && (className.StartsWith(MiniFramePrefix, StringComparison.Ordinal)
                || (className == FloatingPaletteClass && IsTheToolbox(windowEvent.Window))))
        {
            Log.Info($"window event: putting the toolbox down {windowEvent.Window:X} ({className})");
            Win32.ShowWindow(windowEvent.Window, 0);
        }

        // Moves are logged only for the editor's own windows. The hook hears the whole
        // process, and resizing the HOST streams thousands of move events for its own
        // controls per second - a line each was most of the resize lag all by itself.
        // Appearing and disappearing stays logged for everything: those are rare and every
        // one of them is a story.
        if (!windowEvent.IsLocationChange || IsEditorClass(className))
        {
            Log.Verbose($"window event: {windowEvent.Describe()} {className} {windowEvent.Window:X}");
        }

        // A pane appearing or disappearing changes which components have panes open. So does a
        // rename, which is handled above where the name change arrives. Moves and resizes need
        // nothing but window rectangles.
        if (windowEvent.IsCreate || windowEvent.IsDestroy)
        {
            _openComponents = null;
        }

#if DEBUG
        Diagnostics.PerfCounters.WindowEvent();
#endif

        // A window MOVING somewhere else in the process is not news here.
        //
        // The hook hears everything the host owns, and resizing a frame moves every toolbar,
        // scrollbar and child control inside it. One drag of the editor's frame measured 1,298
        // window events for 60 size steps, and a full refresh ran for every one of them: 188ms
        // of that gesture, the largest single cost in it and three times what moving the surface
        // itself took (2026-08-06). What a refresh computes is pane rectangles and which panes
        // are open, and a window OUTSIDE the editor moving can change neither.
        //
        // Appearing and disappearing still refreshes for everything, whatever the class: those
        // are rare, and one of them is how a pane closing gets noticed at all.
        if (!windowEvent.IsLocationChange || MovesPanes(className))
        {
            Refresh();
        }

        // After the refresh, so a pane-driven placement (via Changed) has already happened and
        // the frame-driven one sees final rectangles.
        if (className == FrameClass || className == "MDIClient")
        {
            FrameChanged?.Invoke();
        }

        if (windowEvent.IsDestroy)
        {
            // NOT SHORTCUT FROM HERE, and it was tried. A destroy event names the pane's CHILD
            // windows rather than the pane, so matching one against a tracked pane never fired
            // once; what makes a close feel immediate is the fast resync poll the session runs
            // afterwards (see AddInSession.UpdatePolling, 2026-08-08).
            WindowDestroyed?.Invoke();
        }
    }

    /// <summary>The window classes whose movements concern the editor surface.</summary>
    private static bool IsEditorClass(string? className) => className
        is FrameClass or "MDIClient" or PaneClass or "XlideEditorOverlay"
        or "MsoCommandBarDock" or "MsoCommandBar";

    /// <summary>
    /// Whether a window MOVING can have moved a code pane. Narrower than the editor's own
    /// windows, which is what deserves a log line rather than what deserves a refresh.
    ///
    /// A pane lives inside the MDI client, inside the frame. Those three moving can move a pane;
    /// nothing else in the editor can. Measured over one resize drag, 60 size steps produced 531
    /// of these events and only 236 came from the three (2026-08-06):
    ///
    /// - The toolbars, 236 of them, are half the traffic and move no pane. What they DO change is
    ///   how much room is left for the MDI client - and the MDI client says so itself, with its
    ///   own event, which is still heard.
    /// - Our own overlay, 59 of them, moved because we moved it a moment earlier. A refresh
    ///   caused by our own placement is a loop that tells us nothing.
    ///
    /// Appearing and disappearing is unaffected and still refreshes for every class, because a
    /// pane closing is exactly the thing that must never be missed.
    /// </summary>
    private static bool MovesPanes(string? className) => className
        is FrameClass or "MDIClient" or PaneClass;

    /// <summary>Set when events arrive while a refresh is running, so none of them are lost.</summary>
    private bool _refreshQueued;

    /// <summary>Rebuilds the picture from both sources.</summary>
    public void Refresh()
    {
        // Reading window rectangles can itself raise events on some systems. Re-entering here
        // would recurse without bound - but DROPPING the re-entrant call loses the burst's
        // tail, and the tail is where the truth lives: closing a hidden pane fires its destroy
        // events while the refresh its hide events started is still running, that refresh
        // still sees the dying window, and with the trailing events swallowed the strip showed
        // a closed module's tab until something else stirred ("the tab X doesn't close it",
        // 2026-08-04, second life). Queued instead of dropped, with a bounded trailing loop:
        // one more pass sees the settled truth, and the cap keeps a self-raising read from
        // becoming the recursion the guard was built against.
        if (_refreshing)
        {
            _refreshQueued = true;
            return;
        }

        _refreshing = true;

#if DEBUG
        var refreshStartedAt = Environment.TickCount64;
#endif

        try
        {
            var passes = 0;

            do
            {
                _refreshQueued = false;
                RefreshOnce();
            }
            while (_refreshQueued && ++passes < 4);

#if DEBUG
            Diagnostics.PerfCounters.Refresh(Environment.TickCount64 - refreshStartedAt);
#endif

            if (_refreshQueued)
            {
                // The burst outran the trailing passes. Not silently: the picture is declared
                // stale, and the poll the listener arms is what finishes the story once the
                // storm has passed - a teardown's tail once outran every pass here and left a
                // closed module's tab on the strip.
                _refreshDropped = true;
                Log.Verbose("code panes: a burst outran the trailing passes, deferring to the poll");
                RefreshFailed?.Invoke();
            }
            else
            {
                _refreshDropped = false;
            }
        }
        finally
        {
            _refreshing = false;
        }
    }

    private void RefreshOnce()
    {
        try
        {
            var open = _openComponents ??= ReadOpenComponents();
            var found = FindPaneWindows();

            // The reads succeeded, so the editor is answering again. Said once, with the count,
            // rather than staying silent about a minute of blindness.
            if (_failedRefreshes > 0)
            {
                Log.Info($"code panes: the editor is answering again, after {_failedRefreshes} abandoned refresh(es)");
                _failedRefreshes = 0;
            }

            var openNames = open.Select(candidate => candidate.Name).ToList();
            var updated = new List<CodePane>(found.Count);

            foreach (var window in found)
            {
                var caption = Win32.ReadWindowText(window);
                if (CodePaneCaption.IsKnownNonCodePane(caption))
                {
                    continue;
                }

                var component = CodePaneCaption.MatchComponent(caption, openNames);
                if (component is null)
                {
                    // Unmatched windows are dropped rather than guessed at. Placing a surface over
                    // the wrong window is worse than placing none.
                    continue;
                }

                updated.Add(new CodePane(
                    window,
                    component,
                    ReadBounds(window),
                    Win32.IsWindowVisible(window),
                    ProjectOf(open, component)));
            }

            // Each pass tells the development log what it saw, so a wrong picture names the
            // exact pass that adopted it. Collapses while nothing changes.
            if (Log.VerboseEnabled)
            {
                var composition = string.Join("|",
                    updated.Select(pane => pane.Component + (pane.IsVisible ? string.Empty : "~")));
                Log.Verbose($"code panes: pass saw {updated.Count} [{composition}]{(SameAs(updated) ? " unchanged" : "")}");
            }

            // Events arrive in bursts and most of them change nothing. Telling listeners that
            // nothing changed would make every one of them re-do its own work, which for a surface
            // positioned over a pane means moving a window that is already in the right place.
            if (SameAs(updated))
            {
                return;
            }

            _panes.Clear();
            _panes.AddRange(updated);

            Changed?.Invoke(_panes);
        }
        catch (Exception ex)
        {
            // The editor stops answering while it runs the developer's code and while it tears
            // windows down, and its refusals arrive as more than one exception type. Expected,
            // but not once per window event: a run that lasts a minute used to write two
            // thousand identical lines. The first failure is reported, the repetition is
            // summarised, and the recovery above closes the story.
            _failedRefreshes++;

            if (_failedRefreshes == 1 || _failedRefreshes % 200 == 0)
            {
                var detail = ex is COMException com ? $"0x{com.HResult:X8}" : $"{ex.GetType().Name}: {ex.Message}";
                Log.Info($"code panes: refresh abandoned, {detail} ({_failedRefreshes} in a row)");
            }
            else
            {
                // The failures between the loud ones, for the development log: a silent
                // stretch of retries once hid a strip that stayed stale for seconds.
                Log.Verbose($"code panes: refresh abandoned ({_failedRefreshes} in a row)");
            }

            RefreshFailed?.Invoke();
        }
    }

    /// <summary>True when a freshly built list is identical to the one already published.</summary>
    private bool SameAs(List<CodePane> updated)
    {
        if (updated.Count != _panes.Count)
        {
            return false;
        }

        for (var i = 0; i < updated.Count; i++)
        {
            if (updated[i] != _panes[i])
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Whether the current episode of an unavailable pane list has been reported.</summary>
    private bool _paneListUnavailableLogged;

    /// <summary>
    /// Names a pane window may legitimately be showing.
    ///
    /// The editor's own pane list is asked first, because it is exact. It is also fragile: after
    /// a pane closes it can spend a long time refusing every read while the rest of the object
    /// model answers normally. When that happens the component names of every project stand in.
    /// A pane window's caption can only name a component, so matching against the superset
    /// changes nothing for windows that exist, and it keeps the tracker alive, which is what
    /// keeps tabs switchable, while the pane list sulks.
    /// </summary>
    /// <summary>
    /// The project a matched component belongs to, when exactly one open component carries the
    /// name. Two projects sharing the name means the caption alone cannot say, and null is the
    /// honest answer rather than a guess.
    /// </summary>
    private static string? ProjectOf(List<OpenComponent> open, string component)
    {
        string? project = null;
        var matches = 0;

        foreach (var candidate in open)
        {
            if (string.Equals(candidate.Name, component, StringComparison.OrdinalIgnoreCase))
            {
                matches++;
                project = candidate.Project;
            }
        }

        return matches == 1 ? project : null;
    }

    private List<OpenComponent> ReadOpenComponents()
    {
        try
        {
            var names = ReadPaneComponents();
            _paneListUnavailableLogged = false;
            return names;
        }
        catch (Exception ex)
        {
            if (!_paneListUnavailableLogged)
            {
                _paneListUnavailableLogged = true;
                var stage = ex.Data["stage"] as string ?? "?";
                Log.Info($"code panes: the pane list is unavailable at {stage} "
                         + $"({ex.GetType().Name}: {ex.Message}); matching against every project component");
            }

            return ReadAllComponentNames();
        }
    }

    /// <summary>The exact list, from the editor's pane collection, with the failing step named.</summary>
    private List<OpenComponent> ReadPaneComponents()
    {
        var stage = "CodePanes";

        try
        {
            var names = new List<OpenComponent>();

            using var panes = _editor.GetObject("CodePanes");
            if (panes is null)
            {
                return names;
            }

            stage = "Count";
            var count = panes.GetInt32("Count");

            for (var i = 1; i <= count; i++)
            {
                stage = $"Item({i})";
                using var pane = panes.GetItem(i);
                stage = $"Item({i}).CodeModule";
                using var module = pane?.GetObject("CodeModule");
                stage = $"Item({i}).Parent";
                using var component = module?.GetObject("Parent");

                stage = $"Item({i}).Name";
                var name = component?.GetString("Name");
                if (!string.IsNullOrEmpty(name))
                {
                    names.Add(new OpenComponent(name, ProjectIdentityOf(component!)));
                }
            }

            return names;
        }
        catch (Exception ex)
        {
            ex.Data["stage"] = stage;
            throw;
        }
    }

    /// <summary>
    /// Identity of the project that owns a component, or null when it will not say. One more
    /// hop than the component read itself - the component's collection's parent is the project
    /// - and deliberately forgiving: a pane whose project cannot be named is still a pane.
    /// </summary>
    private static string? ProjectIdentityOf(DispatchObject component)
    {
        try
        {
            using var collection = component.GetObject("Collection");
            using var project = collection?.GetObject("Parent");
            return project is null ? null : ProjectReader.Identity(project).Id;
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>Every component in every project: the superset the fallback matches against.</summary>
    private List<OpenComponent> ReadAllComponentNames()
    {
        var names = new List<OpenComponent>();

        try
        {
            using var projects = _editor.GetObject("VBProjects");
            var projectCount = projects?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= projectCount; i++)
            {
                try
                {
                    using var project = projects!.GetItem(i);
                    if (project is null)
                    {
                        continue;
                    }

                    var identity = ProjectReader.Identity(project).Id;
                    using var components = project.GetObject("VBComponents");
                    var componentCount = components?.GetInt32("Count") ?? 0;

                    for (var j = 1; j <= componentCount; j++)
                    {
                        using var component = components!.GetItem(j);
                        var name = component?.GetString("Name");
                        if (!string.IsNullOrEmpty(name))
                        {
                            names.Add(new OpenComponent(name, identity));
                        }
                    }
                }
                catch (Exception)
                {
                    // A project that will not answer contributes nothing; the others still do.
                }
            }
        }
        catch (Exception)
        {
            // Both sources refusing at once is the abandoned-refresh case; the caller's catch
            // owns reporting it.
        }

        return names;
    }

    /// <summary>
    /// Every visible window of the pane class, anywhere under the editor frame.
    ///
    /// Enumerated rather than walked as direct children: the panes are not all at the same depth,
    /// and asking only for direct children finds none of the docked ones.
    /// </summary>
    internal static unsafe HashSet<nint> VisiblePanes()
    {
        _visible = [];

        try
        {
            var ours = Win32.GetCurrentProcessId();

            nint frame = 0;
            while ((frame = Win32.FindWindowEx(0, frame, FrameClass, null)) != 0)
            {
                Win32.GetWindowThreadProcessId(frame, out var owner);
                if (owner != ours)
                {
                    continue;
                }

                Win32.EnumChildWindows(
                    frame,
                    (nint)(delegate* unmanaged<nint, nint, int>)&OnVisiblePane,
                    0);
            }

            return _visible;
        }
        finally
        {
            _visible = null;
        }
    }

    private static HashSet<nint>? _visible;

    /// <summary>
    /// The pane-class window carrying this exact caption, visible or hidden, or zero.
    ///
    /// This is how the Immediate window is identified: the object model names its localised
    /// caption, the handle carries the same caption, and neither changes with visibility - so
    /// the identification cannot lose the race that diffing visible panes across an
    /// asynchronous hide loses. The caption is unique among the editor's tool windows; code
    /// panes carry their module titles.
    /// </summary>
    internal static unsafe nint FindPaneByCaption(string caption) => FindByCaption(caption, paneClassOnly: true);

    /// <summary>
    /// The frame descendant carrying this exact caption regardless of window class, or zero.
    ///
    /// The pane-class filter above is right for the Immediate window, which shares its class with
    /// the code panes. The Object Browser does not - its class is its own - so callers locating an
    /// arbitrary tool window by its object-model caption match on the caption alone. Tool captions
    /// are localised words like "Locals"; module panes carry their file-qualified titles, so the
    /// two vocabularies do not collide.
    /// </summary>
    internal static unsafe nint FindChildByCaption(string caption) => FindByCaption(caption, paneClassOnly: false);

    private static unsafe nint FindByCaption(string caption, bool paneClassOnly)
    {
        _captionWanted = caption;
        _captionFound = 0;

        try
        {
            var ours = Win32.GetCurrentProcessId();

            nint frame = 0;
            while (_captionFound == 0 && (frame = Win32.FindWindowEx(0, frame, FrameClass, null)) != 0)
            {
                Win32.GetWindowThreadProcessId(frame, out var owner);
                if (owner != ours)
                {
                    continue;
                }

                Win32.EnumChildWindows(
                    frame,
                    (nint)(delegate* unmanaged<nint, nint, int>)&OnCaptionCandidate,
                    paneClassOnly ? 1 : 0);
            }

            return _captionFound;
        }
        finally
        {
            _captionWanted = null;
        }
    }

    /// <summary>
    /// The top-level window in this process carrying this exact caption, or zero. This is how a
    /// FLOATING tool palette is found: it is nobody's child, so the frame walks above cannot
    /// reach it.
    /// </summary>
    internal static unsafe nint FindTopLevelByCaption(string caption)
    {
        _captionWanted = caption;
        _captionFound = 0;

        try
        {
            Win32.EnumWindows(
                (nint)(delegate* unmanaged<nint, nint, int>)&OnTopLevelCaptionCandidate,
                0);
            return _captionFound;
        }
        finally
        {
            _captionWanted = null;
        }
    }

    [UnmanagedCallersOnly]
    private static int OnTopLevelCaptionCandidate(nint window, nint parameter)
    {
        Win32.GetWindowThreadProcessId(window, out var owner);
        if (owner == Win32.GetCurrentProcessId()
            && _captionWanted is { } wanted
            && Win32.ReadWindowText(window) == wanted)
        {
            _captionFound = window;
            return 0;
        }

        return 1;
    }

    /// <summary>Every descendant of a window carrying this exact class.</summary>
    internal static unsafe List<nint> FindChildrenByClass(nint root, string className)
    {
        _classWanted = className;
        _classFound = [];

        try
        {
            Win32.EnumChildWindows(
                root,
                (nint)(delegate* unmanaged<nint, nint, int>)&OnClassCandidate,
                0);
            return _classFound;
        }
        finally
        {
            _classWanted = null;
        }
    }

    private static string? _classWanted;
    private static List<nint> _classFound = [];

    [UnmanagedCallersOnly]
    private static int OnClassCandidate(nint window, nint parameter)
    {
        if (_classWanted is { } wanted && Win32.ReadClassName(window) == wanted)
        {
            _classFound.Add(window);
        }

        return 1;
    }

    private static string? _captionWanted;
    private static nint _captionFound;

    [UnmanagedCallersOnly]
    private static int OnCaptionCandidate(nint window, nint parameter)
    {
        if (_captionWanted is { } wanted
            && (parameter == 0 || Win32.ReadClassName(window) == PaneClass)
            && Win32.ReadWindowText(window) == wanted)
        {
            _captionFound = window;
            return 0;
        }

        return 1;
    }

    [UnmanagedCallersOnly]
    private static int OnVisiblePane(nint window, nint parameter)
    {
        if (_visible is not null && Win32.IsWindowVisible(window) && Win32.ReadClassName(window) == PaneClass)
        {
            _visible.Add(window);
        }

        return 1;
    }

    private unsafe List<nint> FindPaneWindows()
    {
        _candidates.Clear();

        var frame = FindEditorFrame();
        if (frame == 0)
        {
            return _candidates;
        }

        // The callback is a plain function pointer under ahead-of-time compilation, so the instance
        // is reached through a static field for the duration of the call. Enumeration is synchronous
        // and single threaded, so the field cannot be observed by anyone else.
        _enumerating = this;

        try
        {
            Win32.EnumChildWindows(
                frame,
                (nint)(delegate* unmanaged<nint, nint, int>)&OnChildWindow,
                0);
        }
        finally
        {
            _enumerating = null;
        }

        return _candidates;
    }

    [UnmanagedCallersOnly]
    private static int OnChildWindow(nint window, nint parameter)
    {
        try
        {
            var tracker = _enumerating;
            if (tracker is not null && Win32.ReadClassName(window) == PaneClass)
            {
                tracker._candidates.Add(window);
            }
        }
        catch
        {
            // Nothing may escape into the enumeration.
        }

        return 1;
    }

    /// <summary>Finds the editor's own frame window, which every pane is a descendant of.</summary>
    private unsafe nint FindEditorFrame()
    {
        using var main = _editor.GetObject("MainWindow");
        if (main is null)
        {
            return 0;
        }

        // The object model exposes no window handle, so the frame is located by walking up from any
        // window the editor owns. Starting from a pane we already know is cheapest when we have one.
        foreach (var pane in _panes)
        {
            if (Win32.IsWindow(pane.Window))
            {
                var root = Win32.GetAncestor(pane.Window, Win32.GaRoot);
                if (root != 0 && Win32.ReadClassName(root) == FrameClass)
                {
                    return root;
                }
            }
        }

        return FindFrameByEnumeration();
    }

    /// <summary>The editor's frame window, found without needing any pane to exist.</summary>
    internal static nint FindFrame() => FindFrameByEnumeration();

    private static unsafe nint FindFrameByEnumeration()
    {
        // GetTopWindow-style enumeration over the desktop would reach other processes. Walking our
        // own windows is enough: the editor frame is a top-level window in this process.
        var found = 0L;
        var self = (uint)Environment.ProcessId;

        for (var window = Win32.GetTopWindow(0); window != 0; window = Win32.GetWindow(window, Win32.GwHwndNext))
        {
            uint owner;
            Win32.GetWindowThreadProcessId(window, &owner);
            if (owner != self)
            {
                continue;
            }

            if (Win32.ReadClassName(window) == FrameClass)
            {
                found = window;
                break;
            }
        }

        return (nint)found;
    }


    /// <summary>
    /// The pane's client area in screen coordinates: the part that shows text, without the pane's
    /// own borders.
    ///
    /// The client area rather than the window rectangle, because a pane maximised inside the
    /// editor's document area has a window rectangle that reaches outside it. Its own borders and
    /// the space its caption would occupy sit above and to the left of the area anything is drawn
    /// in. A surface placed on the window rectangle is pushed up out of view by exactly that much,
    /// and shows the module scrolled down by the first couple of lines.
    /// </summary>
    private static unsafe PixelRect ReadBounds(nint window)
    {
        Rect client;
        if (!Win32.GetClientRect(window, &client))
        {
            return default;
        }

        var corners = stackalloc Point[2];
        corners[0] = new Point { X = client.Left, Y = client.Top };
        corners[1] = new Point { X = client.Right, Y = client.Bottom };

        // A destination of zero means the screen.
        Marshal.SetLastSystemError(0);
        if (Win32.MapWindowPoints(window, 0, corners, 2) == 0 && Marshal.GetLastSystemError() != 0)
        {
            return default;
        }

        return new PixelRect(
            Math.Min(corners[0].X, corners[1].X),
            Math.Min(corners[0].Y, corners[1].Y),
            Math.Max(corners[0].X, corners[1].X),
            Math.Max(corners[0].Y, corners[1].Y));
    }

    public void Dispose()
    {
        _hook?.Dispose();
        _hook = null;
        _panes.Clear();
    }
}
