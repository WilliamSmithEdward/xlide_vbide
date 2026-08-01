using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

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
        _addIn?.Dispose();
        _editor.Dispose();
    }
}
