using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;

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

                // Until the surface renders them, the log is where they surface. Capped, because a
                // project with thousands of findings would otherwise write a novel on every pass.
                foreach (var finding in findings.Take(20))
                {
                    Log.Info($"  {finding.Module}({finding.StartLine},{finding.StartColumn}) " +
                             $"{finding.Severity} {finding.Code}: {finding.Message}");
                }

                if (findings.Count > 20)
                {
                    Log.Info($"  and {findings.Count - 20} more");
                }
            };

            _analysis.Start();
        }
        catch (Exception ex)
        {
            Log.Error("analysis: could not be started", ex);
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
