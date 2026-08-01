using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.UI;

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
    private readonly DispatchObject? _addIn;
    private DispatchObject? _toolWindow;
    private CodePaneTracker? _codePanes;
    private AnalysisService? _analysis;
    private EditorSurface? _editorSurface;

    /// <summary>
    /// The most recent findings for every module, kept so a module can be decorated the moment it
    /// is shown. Analysis runs per project and the surface shows one module at a time, so without
    /// this a module opened between two passes carries no squiggles until the next one.
    /// </summary>
    private IReadOnlyList<Finding> _findings = [];

    private bool _stopped;

    public AddInSession(DispatchObject editor, DispatchObject? addIn)
    {
        _editor = editor;
        _addIn = addIn;
    }

    /// <summary>Automation object for the editor itself.</summary>
    public DispatchObject Editor => _editor;

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

        // Subscribed before the panel can exist, so a double-click that lands during start-up is
        // handled rather than dropped.
        PanelBus.NavigateRequested = GoTo;

        CreateToolWindow();
        TrackCodePanes();
        StartAnalysis();
    }

    /// <summary>
    /// Takes the user to a finding: the native pane is selected and the caret placed on it, and the
    /// surface over that pane scrolls to match.
    ///
    /// The native pane is moved as well as the surface, because it stays the text of record and
    /// what the debugger drives. Leaving it where it was would put the two out of step the first
    /// time the user pressed F8.
    /// </summary>
    private void GoTo(string component, int line, int column)
    {
        try
        {
            using var pane = FindCodePane(component);
            if (pane is null)
            {
                Log.Info($"navigate: no pane for {component}");
                return;
            }

            pane.Invoke("Show");
            pane.Invoke("SetSelection", line, column, line, column);

            if (_editorSurface?.Module == component)
            {
                _editorSurface.Reveal(line);
            }

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

                PanelBus.PublishFindings(findings);
                PublishMarkersForShownModule();
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
            var pane = panes.FirstOrDefault(p => p.IsVisible);

            if (pane.Window == 0)
            {
                _editorSurface?.Follow(default, visible: false);
                return;
            }

            // The surface goes in beside the pane, as a child of whatever the pane's own parent is.
            //
            // That parent is the editor's document area, and being inside it is what keeps the
            // surface off everything that is not a code pane: the window manager clips a child to
            // its parent, so the toolbars, the docked panels and the splitters between them cannot
            // be covered however wrong the rectangle is. Parenting to the frame instead made the
            // surface a sibling of the toolbars, which is exactly what it then drew over.
            var host = Win32.GetParent(pane.Window);
            if (host == 0)
            {
                return;
            }

            // A pane can be reparented, by being undocked or by the editor rebuilding its layout.
            // The surface belongs to one parent, so a change means a new one rather than a move.
            if (_editorSurface is not null && _editorSurface.Host != host)
            {
                Log.Info("editor surface: the document area changed, rebuilding");
                _editorSurface.Dispose();
                _editorSurface = null;
            }

            _editorSurface ??= EditorSurface.Create(host, default);
            if (_editorSurface is null)
            {
                return;
            }

            // The surface covers the whole document area, not the rectangle of one pane.
            //
            // Sizing it to a pane meant it had to be moved and re-raised every time the editor
            // activated a different one, and the editor paints the pane it is activating before any
            // of that can happen. The result was the old pane flashing into view on every module
            // switch, and a surface that could be left underneath whatever had just been raised.
            // Covering the area once removes both: switching modules becomes a message to a surface
            // that never moved and was never uncovered.
            //
            // The native panes keep running underneath, unchanged and never seen. They remain the
            // text of record, the compile target, and what the debugger drives.
            _editorSurface.Follow(ClientBounds(host), visible: true);

            if (pane.Component is not null && pane.Component != _editorSurface.Module)
            {
                ShowModuleInSurface(pane.Component);
            }
        }
        catch (Exception ex)
        {
            Log.Error("editor surface: could not follow the active pane", ex);
        }
    }

    /// <summary>Reads a module's text and hands it to the surface, with its squiggles.</summary>
    private void ShowModuleInSurface(string component)
    {
        using var found = FindComponent(component);
        if (found is null)
        {
            return;
        }

        var source = ProjectReader.ReadSource(found);
        if (source is null)
        {
            return;
        }

        _editorSurface?.Show(component, source);
        Log.Info($"editor surface: showing {component}, {source.Length} character(s)");

        // The findings for this module were computed before it was opened, so they are applied here
        // rather than waiting for the next analysis pass.
        PublishMarkersForShownModule();
    }

    /// <summary>Finds a component by name across every open project, or null when there is none.</summary>
    private DispatchObject? FindComponent(string component)
    {
        using var projects = _editor.GetObject("VBProjects");
        var count = projects?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            using var project = projects!.GetItem(i);
            using var components = project?.GetObject("VBComponents");
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
                    return candidate;
                }

                candidate?.Dispose();
            }
        }

        return null;
    }

    /// <summary>Finds the code pane a component's module is displayed in, opening one if needed.</summary>
    private DispatchObject? FindCodePane(string component)
    {
        using var found = FindComponent(component);
        using var module = found?.GetObject("CodeModule");

        // Reading CodePane on a module that has never been opened creates the pane, which is what
        // makes navigating to a module the user has not opened work at all.
        return module?.GetObject("CodePane");
    }

    /// <summary>
    /// Sends the surface the squiggles belonging to whichever module it is showing.
    ///
    /// Findings arrive for a whole project and the surface shows one module, so they are filtered
    /// here. A module with none is sent an empty set rather than skipped: that is what clears
    /// squiggles the user has just fixed.
    /// </summary>
    private void PublishMarkersForShownModule()
    {
        var surface = _editorSurface;
        var module = surface?.Module;
        if (surface is null || module is null)
        {
            return;
        }

        var markers = _findings
            .Where(f => string.Equals(f.Module, module, StringComparison.OrdinalIgnoreCase))
            .Select(f => new EditorMarker(
                f.StartLine,
                f.StartColumn,
                f.EndLine,
                f.EndColumn,
                f.Severity,
                f.Message,
                f.Code))
            .ToArray();

        surface.ShowDiagnostics(markers);
    }

    /// <summary>
    /// A window's client area, in the coordinates a child of it is positioned in.
    ///
    /// A client rectangle always starts at the origin, so this needs no conversion and carries no
    /// assumption about borders, captions, or scaling. An earlier version worked the origin out
    /// from the window and client rectangles and placed the surface a toolbar's height too high,
    /// which is what put it over the toolbar.
    /// </summary>
    private static unsafe PixelRect ClientBounds(nint window)
    {
        Rect client;
        return Win32.GetClientRect(window, &client)
            ? new PixelRect(0, 0, client.Right - client.Left, client.Bottom - client.Top)
            : default;
    }

    /// <summary>Height the panel asks for, in the units the editor's layout uses.</summary>
    private const int PanelHeight = 260;

    /// <summary>
    /// Gives the tool window a usable size, if the editor lets us.
    ///
    /// A docked window ignores this, which is correct: its size belongs to the layout the user
    /// arranged. Only a floating one takes it, which is the case that starts out too small to read.
    /// </summary>
    private static void TrySize(DispatchObject window, int width, int height)
    {
        // Set independently. A docked window belongs to a band that owns one of its dimensions and
        // refuses that one while accepting the other, so setting them together loses the second to
        // the first one's refusal.
        try
        {
            window.SetInt32("Height", height);
        }
        catch (Exception ex)
        {
            Log.Info($"tool window: kept its own height ({ex.GetType().Name})");
        }

        try
        {
            window.SetInt32("Width", width);
        }
        catch (Exception ex)
        {
            Log.Info($"tool window: kept its own width ({ex.GetType().Name})");
        }
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
                Log.Info($"code panes: {panes.Count} open");
                foreach (var pane in panes)
                {
                    Log.Info($"  {pane.Component} at {pane.Bounds.Left},{pane.Bounds.Top} " +
                             $"{pane.Bounds.Width}x{pane.Bounds.Height}" + (pane.IsVisible ? string.Empty : " (hidden)"));
                }

                FollowActivePane(panes);
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

        // The bus outlives the session, so a handler left on it would keep this one reachable and
        // let a panel still on screen call into it after everything below has been released.
        PanelBus.NavigateRequested = null;

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

        _codePanes?.Dispose();
        _codePanes = null;

        // Before the editor tears its own windows down. The surface owns a browser and a window
        // parented to the editor frame; leaving them for the host to destroy leaves browser
        // processes with no parent and a window procedure in a library about to be unloaded.
        _editorSurface?.Dispose();
        _editorSurface = null;

        _toolWindow?.Dispose();
        _toolWindow = null;

        Log.Info("session stopped");
    }

    /// <summary>
    /// Asks the editor for a docked tool window sited on our control.
    ///
    /// The editor creates the control itself from its registered program identifier, so this call
    /// is what makes the whole hosting arrangement start. Everything that can go wrong is on the
    /// far side of it and reports nothing: an unregistered class, a missing interface, or a control
    /// that refuses activation all produce the same failed call or an empty pane. The step-by-step
    /// logging is the only view into which of those happened.
    /// </summary>
    private void CreateToolWindow()
    {
        if (_addIn is null)
        {
            Log.Warn("tool window: the editor supplied no add-in instance, so it cannot be created");
            return;
        }

        try
        {
            Log.Info($"tool window: creating '{ProductIdentity.ToolWindowHostProgId}'");

            using var windows = _editor.GetObject("Windows");
            if (windows is null)
            {
                Log.Error("tool window: the editor exposed no window collection");
                return;
            }

            // CreateToolWindow(AddInInst, ProgId, Caption, GuidPosition, ByRef DocObj) As Window.
            // The editor's own type library declares the parameters as an add-in interface pointer,
            // three strings, and an in-out automation pointer that receives the sited control.
            //
            // The add-in argument is a plain descriptor over a reference this session already owns,
            // so it is deliberately not disposed: doing so would release a reference we did not add.
            var addInInstance = ComVariant.CreateRaw(VarEnum.VT_DISPATCH, _addIn.Pointer);
            using var progId = ComVariant.Create(ProductIdentity.ToolWindowHostProgId);
            using var caption = ComVariant.Create(ProductIdentity.ToolWindowCaption);
            using var position = ComVariant.Create(ProductIdentity.ToolWindowPositionGuid);

            var window = windows.CallWithByRefObject(
                "CreateToolWindow",
                [addInInstance, progId, caption, position],
                out var control);

            // The control handed back is our own object seen through the editor. The editor keeps
            // its own reference to it for as long as the window exists.
            control?.Dispose();

            if (window is null)
            {
                Log.Error("tool window: the editor returned no window");
                return;
            }

            Log.Info("tool window: got window");
            _toolWindow = window;

            window.SetBool("Visible", true);
            Log.Info("tool window: visible");

            // Left where the editor puts it, and deliberately not docked.
            //
            // The editor refuses to size a tool window in either state: setting Width or Height
            // throws whether it is floating or docked, and a docked one is given a band six pixels
            // high with a negative client area, so docking makes it invisible rather than usable.
            // Measured, not assumed: see docs/lessons.md.
            //
            // This window is therefore not where the product's panels belong. They live in the
            // surface, which owns its own layout completely. This one stays as the foothold the
            // editor gives an add-in, and as somewhere to report that the add-in is loaded.
            TrySize(window, 460, PanelHeight);
        }
        catch (COMException ex)
        {
            Log.Error($"tool window: the editor refused to create it, 0x{ex.HResult:X8}", ex);
        }
        catch (Exception ex)
        {
            Log.Error("tool window: creation failed", ex);
        }
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
        _toolWindow?.Dispose();
        _toolWindow = null;
        _addIn?.Dispose();
        _editor.Dispose();
    }
}
