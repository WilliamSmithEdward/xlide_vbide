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
        CreateToolWindow();
        TrackCodePanes();
        StartAnalysis();
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

            // The surface is a sibling of the pane inside the editor frame, so its rectangle has to
            // be expressed in the frame's coordinates rather than the screen's.
            var frame = Win32.GetAncestor(pane.Window, Win32.GaRoot);
            if (frame == 0)
            {
                return;
            }

            _editorSurface ??= EditorSurface.Create(frame, default);
            if (_editorSurface is null)
            {
                return;
            }

            _editorSurface.Follow(ToFrameCoordinates(frame, pane.Bounds), visible: true);

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

    /// <summary>Reads a module's text and hands it to the surface.</summary>
    private void ShowModuleInSurface(string component)
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
                using var candidate = components.GetItem(j);
                if (candidate?.GetString("Name") != component)
                {
                    continue;
                }

                var source = ProjectReader.ReadSource(candidate);
                if (source is not null)
                {
                    _editorSurface?.Show(component, source);
                    Log.Info($"editor surface: showing {component}, {source.Length} character(s)");
                }

                return;
            }
        }
    }

    /// <summary>
    /// Converts a screen rectangle into the frame's client coordinates, which is what positioning a
    /// child of the frame requires.
    /// </summary>
    private static unsafe PixelRect ToFrameCoordinates(nint frame, PixelRect screen)
    {
        Rect frameRect;
        if (!Win32.GetWindowRect(frame, &frameRect))
        {
            return screen;
        }

        // The frame's client area starts inside its border, and the difference is what a child is
        // positioned relative to. Measuring both rectangles gives it without assuming a border size.
        Rect client;
        if (!Win32.GetClientRect(frame, &client))
        {
            return screen;
        }

        var borderX = ((frameRect.Right - frameRect.Left) - (client.Right - client.Left)) / 2;
        var originX = frameRect.Left + borderX;
        var originY = frameRect.Bottom - borderX - (client.Bottom - client.Top);

        return new PixelRect(
            screen.Left - originX,
            screen.Top - originY,
            screen.Right - originX,
            screen.Bottom - originY);
    }

    /// <summary>
    /// Gives the tool window a usable size, if the editor lets us.
    ///
    /// A docked window ignores this, which is correct: its size belongs to the layout the user
    /// arranged. Only a floating one takes it, which is the case that starts out too small to read.
    /// </summary>
    private static void TrySize(DispatchObject window, int width, int height)
    {
        try
        {
            window.SetInt32("Width", width);
            window.SetInt32("Height", height);
        }
        catch (Exception ex)
        {
            Log.Info($"tool window: kept its own size ({ex.GetType().Name})");
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

            // Sized after it is shown, because the editor refuses the dimensions of a window that
            // is not yet on screen. It creates one barely larger than an icon, which is unusable
            // for a list; after this the editor remembers whatever size the user settles on.
            TrySize(window, 460, 620);
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
