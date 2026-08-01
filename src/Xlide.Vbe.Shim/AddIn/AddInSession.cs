using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core;
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
        HideReplacedWindows();
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

                PublishMarkersForShownModule();
                PublishFindingsToSurface();
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

            // A pane can be reparented, by being undocked or by the editor rebuilding its layout.
            // The surface belongs to one parent, so a change means a new one rather than a move.
            if (_editorSurface is not null && _editorSurface.Host != host)
            {
                Log.Info("editor surface: the document area changed, rebuilding");
                _editorSurface.Dispose();
                _editorSurface = null;
            }

            if (_editorSurface is null)
            {
                _editorSurface = EditorSurface.Create(host, default);
                if (_editorSurface is null)
                {
                    return;
                }

                _editorSurface.KeyPressed = OnSurfaceKey;
                _editorSurface.ModuleRequested = ShowModule;
                _editorSurface.NavigateRequested = GoTo;
                _editorSurface.CommandRequested = RunCommand;
                _editorSurface.TextChanged = WriteModule;
                _editorSurface.BreakpointToggleRequested = ToggleBreakpoint;
                _editorSurface.Polled = PollDebugState;
            }

            // The surface covers the whole document area, not the rectangle of one pane. Switching
            // module is then a message to a surface that never moved and was never uncovered.
            //
            // The native panes keep running underneath, unchanged and never seen. They remain the
            // text of record, the compile target, and what the debugger drives.
            _editorSurface.Follow(ClientAreaIn(documentArea, host), visible: true);

            if (pane.Component is not null && pane.Component != _editorSurface.Module)
            {
                // Before the document is replaced. Loading a module resets the surface, so an edit
                // that has not been written yet would go with the document it belonged to.
                _editorSurface.FlushEdits();
                ShowModuleInSurface(pane.Component);
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

        var command = VbeCommands.ForKey(virtualKey, shift, control);
        Log.Info($"key: 0x{virtualKey:X2}{(shift ? " shift" : string.Empty)}{(control ? " ctrl" : string.Empty)}"
                 + $" -> {(command == 0 ? "not ours" : command.ToString(System.Globalization.CultureInfo.InvariantCulture))}");

        if (command == 0)
        {
            return false;
        }

        // The editor runs what the module holds, and the caret it uses is its own. Both are brought
        // up to date here, at the one moment it matters: running code the developer has not
        // finished typing is worse than a short pause before it starts.
        _editorSurface?.FlushEdits();
        SyncCaretToPane();

        // Toggling goes through the bookkeeping rather than straight at the command, so the
        // breakpoint drawn on the surface and the one the editor holds stay the same set however
        // the developer asked for it.
        if (command == VbeCommands.Command.ToggleBreakpoint)
        {
            ToggleBreakpoint(_editorSurface?.CaretLine ?? 0);
            return true;
        }

        VbeCommands.Execute(_editor, command);
        WatchDebugState();
        return true;
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
    private void WriteModule(string component, string text)
    {
        try
        {
            using var found = FindComponent(component);
            using var module = found?.GetObject("CodeModule");
            if (found is null || module is null)
            {
                Log.Warn($"write: {component} has no code module");
                return;
            }

            var existing = module.GetInt32("CountOfLines");
            if (existing > 0)
            {
                module.Invoke("DeleteLines", 1, existing);
            }

            // A module with nothing in it is a legitimate state, and asking the host to add an
            // empty string to one is not.
            if (text.Length > 0)
            {
                module.Invoke("AddFromString", text);
            }

            // Read straight back, because the editor rewrites what it is given. It respells
            // keywords and normalises spacing as it takes a module in, so the text it now holds is
            // not the text that was sent. Adopting its version here is what keeps the two the same
            // document; without it every later comparison sees a difference that is the editor's
            // doing and not the developer's, and the surface slowly drifts away from the truth.
            var stored = ProjectReader.ReadSource(found);
            if (stored is not null && stored != text)
            {
                Log.Info($"write: {component} was normalised by the editor, adopting its version");
                _editorSurface?.Sync(component, stored);
            }

            Log.Info($"write: {component}, {text.Length} character(s)");

            // The analyzer reads the module, so it has nothing new to say until the module has
            // been written. Without this the squiggles describe the text as it was before the
            // developer started typing.
            _analysis?.Reanalyse();
        }
        catch (Exception ex)
        {
            Log.Error($"write: {component} could not be updated", ex);
        }
    }

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
    private readonly Dictionary<string, SortedSet<int>> _breakpoints = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Whether execution was stopped last time it was looked at.</summary>
    private bool _inBreak;

    /// <summary>Project modes, as the editor numbers them.</summary>
    private const int BreakMode = 1;
    private const int DesignMode = 2;

    /// <summary>How often the execution state is looked at while anything might be running.</summary>
    private const uint DebugPollMilliseconds = 150;

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

    /// <summary>Toggles a breakpoint on a line of the module currently shown.</summary>
    private void ToggleBreakpoint(int line)
    {
        var module = _editorSurface?.Module;
        if (module is null || line < 1)
        {
            return;
        }

        if (!CanBreakOn(_editorSurface?.LineAt(line)))
        {
            Log.Info($"breakpoint: {module}({line}) is not an executable statement");
            return;
        }

        try
        {
            // The command acts on the caret, so the caret is put on the line first. Everything the
            // developer typed goes with it: a breakpoint is set by line number, and writing the
            // module afterwards would move it.
            _editorSurface?.FlushEdits();

            using var pane = FindCodePane(module);
            if (pane is null)
            {
                return;
            }

            pane.Invoke("SetSelection", line, 1, line, 1);

            if (!VbeCommands.Execute(_editor, VbeCommands.Command.ToggleBreakpoint))
            {
                return;
            }

            if (!_breakpoints.TryGetValue(module, out var lines))
            {
                lines = [];
                _breakpoints[module] = lines;
            }

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

        _editorSurface?.ShowBreakpoints(
            _breakpoints.TryGetValue(module, out var lines) ? [.. lines] : []);
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
    private void UpdateDebugState()
    {
        try
        {
            using var project = _editor.GetObject("ActiveVBProject");
            var mode = project?.GetInt32("Mode") ?? DesignMode;

            if (mode != BreakMode)
            {
                if (_inBreak)
                {
                    _inBreak = false;
                    _editorSurface?.ShowCurrentLine(null);
                    Log.Info($"debug: mode {mode}, not stopped");
                }

                return;
            }

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
            Log.Error("debug: the execution state could not be read", ex);
        }
    }

    /// <summary>Starts watching the execution state, for a while.</summary>
    private void WatchDebugState()
    {
        // Twenty seconds of watching. Long enough for a procedure that does some work before it
        // reaches a breakpoint, short enough that a run which never stops does not poll all day.
        _pollsRemaining = (int)(20_000 / DebugPollMilliseconds);
        _editorSurface?.Poll(DebugPollMilliseconds);
        UpdateDebugState();
    }

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
        var module = surface?.Module;

        if (surface is null || module is null || surface.HasUnwrittenEdits)
        {
            return;
        }

        try
        {
            using var found = FindComponent(module);
            var stored = found is null ? null : ProjectReader.ReadSource(found);

            if (stored is not null && stored != surface.Text)
            {
                Log.Info($"resync: {module} changed underneath the surface, adopting the module");
                surface.Sync(module, stored);
                _analysis?.Reanalyse();
            }
        }
        catch (Exception ex)
        {
            Log.Error($"resync: {module} could not be compared with the module", ex);
        }
    }

    /// <summary>One tick of the execution watch.</summary>
    private void PollDebugState()
    {
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
            _editorSurface?.Poll(0);
        }
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

        // Same as for a keystroke. A toolbar button also takes focus away from the surface, which
        // is exactly when the two are furthest apart.
        _editorSurface?.FlushEdits();
        SyncCaretToPane();
        VbeCommands.Execute(_editor, command);
        WatchDebugState();
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
        PublishFindingsToSurface();
        PublishBreakpoints();
    }

    /// <summary>
    /// Tells the surface which modules the editor has open, for its tab strip.
    ///
    /// The list comes from the editor's own collection of open panes rather than from the project's
    /// components, so the tabs are the modules the developer actually has open, not every module
    /// that exists. Reading a component's pane would create one, which would put a tab up for a
    /// module nobody opened.
    /// </summary>
    private void PublishModules()
    {
        var surface = _editorSurface;
        if (surface is null)
        {
            return;
        }

        try
        {
            using var panes = _editor.GetObject("CodePanes");
            var count = panes?.GetInt32("Count") ?? 0;

            var modules = new List<string>(count);
            for (var i = 1; i <= count; i++)
            {
                using var pane = panes!.GetItem(i);
                using var module = pane?.GetObject("CodeModule");
                using var component = module?.GetObject("Parent");

                if (component?.GetString("Name") is { Length: > 0 } name && !modules.Contains(name))
                {
                    modules.Add(name);
                }
            }

            surface.ShowModules([.. modules], surface.Module);
        }
        catch (Exception ex)
        {
            Log.Error("modules: the open panes could not be listed", ex);
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
                    if (component?.GetString("Name") is { Length: > 0 } name)
                    {
                        members.Add(new SurfaceComponent(name, component.GetInt32("Type")));
                    }
                }

                tree.Add(new SurfaceProject(project.GetString("Name") ?? "VBAProject", [.. members]));
            }

            surface.ShowProjects([.. tree]);
        }
        catch (Exception ex)
        {
            Log.Error("explorer: the project tree could not be read", ex);
        }
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
        // Project explorer and properties. The Immediate, Locals and Watch windows stay: they are
        // not replaced yet, and hiding them would take the feature away rather than restyle it.
        ReadOnlySpan<int> replaced = [6, 7];

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

                if (window.GetBool("Visible"))
                {
                    window.SetBool("Visible", false);
                    Log.Info($"window: closed the editor's own '{window.GetString("Caption")}'");
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error("window: the replaced windows could not be closed", ex);
        }
    }

    /// <summary>Publishes every finding to the surface's panel, across all modules.</summary>
    private void PublishFindingsToSurface()
    {
        _editorSurface?.ShowFindings([.. _findings.Select(f => new SurfaceFinding(
            f.Module,
            f.Code,
            f.Message,
            f.Severity,
            f.StartLine,
            f.StartColumn))]);
    }

    /// <summary>Brings a module's pane to the front, which the surface then follows.</summary>
    private void ShowModule(string component)
    {
        try
        {
            using var pane = FindCodePane(component);
            pane?.Invoke("Show");
        }
        catch (Exception ex)
        {
            Log.Error($"modules: {component} could not be shown", ex);
        }
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

        // Before anything is torn down, and before the engine is stopped: whatever the developer
        // typed last must reach the module, or closing the host loses it.
        _editorSurface?.FlushEdits();

        _codePanes?.Dispose();
        _codePanes = null;

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
