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
    private ImmediateEvaluator? _immediate;
    private ImmediateReader? _immediateReader;
    private bool _windowsHidden;
    private EditorSurface? _editorSurface;

    /// <summary>
    /// The most recent findings for every module, kept so a module can be decorated the moment it
    /// is shown. Analysis runs per project and the surface shows one module at a time, so without
    /// this a module opened between two passes carries no squiggles until the next one.
    /// </summary>
    private IReadOnlyList<Finding> _findings = [];

    /// <summary>
    /// What each module read back as the last time this add-in wrote it.
    ///
    /// This is the baseline a later comparison is made against, and it is deliberately not the
    /// surface's text. The editor rewrites what it is given as it takes a module in: it completes
    /// a procedure's parentheses, inserts the blank body of one, and respells keywords. Comparing
    /// the module against the surface would see all of that as a change and pull it back into the
    /// document, on top of somebody who is still typing. Comparing it against this sees only what
    /// changed the module after we wrote it, which is what "something else changed it" means.
    /// </summary>
    private readonly Dictionary<string, string> _writtenModules = new(StringComparer.OrdinalIgnoreCase);

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
                _surfaceShown = false;
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

            // Remembered so that placement can be recomputed at moments that are not window
            // events: a menu item opening a native window, or the page announcing it is ready.
            _frame = host;
            _documentArea = documentArea;

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
                _editorSurface.EvaluateRequested = EvaluateImmediate;
                _editorSurface.PanelChanged = OnPanelChanged;
                _editorSurface.MenuRequested = OnMenuRequested;
                _editorSurface.MenuExecuteRequested = OnMenuExecuteRequested;
                _editorSurface.PropertyEditRequested = OnPropertyEdit;
                _editorSurface.ComponentSelected = OnComponentSelected;

                // The moment the page is up is the moment the menu bar can be covered, and it is
                // not a window event, so nothing else would recompute the bounds.
                _editorSurface.Ready = RefreshSurfacePlacement;

                // Now rather than at start-up. The editor answers that these windows are visible
                // before it has created them, so hiding one then closes something with no window
                // behind it and there is nothing to identify afterwards.
                HideReplacedWindows();
                HideRedundantToolbar();
                DarkenTitleBar(host);
            }

            // The surface covers the whole document area, not the rectangle of one pane. Switching
            // module is then a message to a surface that never moved and was never uncovered.
            //
            // The native panes keep running underneath, unchanged and never seen. They remain the
            // text of record, the compile target, and what the debugger drives.
            var covering = CanCoverChrome();
            _surfaceShown = true;
            _editorSurface.Follow(SurfaceBounds(host, documentArea, covering), visible: true);
            _editorSurface.SetChrome(menuBar: covering);

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

        VbeCommands.Execute(_editor, command);
        WatchDebugState();
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

            // Read straight back and remembered, but not pushed into the surface.
            //
            // The editor rewrites what it is given, and its rewrites are the kind a developer is in
            // the middle of doing for themselves: it completes the parentheses on a procedure and
            // inserts a blank body for one. Sending that back mid-keystroke duplicated what had
            // just been typed and inserted lines nobody asked for. What it holds is remembered as
            // the baseline instead, so a later comparison sees changes made by something else and
            // not the editor's own tidying of our own write.
            var stored = ProjectReader.ReadSource(found);
            _writtenModules[component] = stored ?? text;

            Log.Info($"write: {component}, {text.Length} character(s)"
                     + (stored is not null && stored != text ? " (the editor reformatted it)" : string.Empty));

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
            // Said out loud, not only logged. The developer pressed something and nothing happened,
            // and "nothing happened" is indistinguishable from a fault unless the reason is given.
            Log.Info($"breakpoint: {module}({line}) is not an executable statement");
            _editorSurface?.Notify($"No breakpoint on line {line}: only executable statements can carry one.");
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
        if (_writtenModules.Remove(oldName, out var baseline))
        {
            _writtenModules[newName] = baseline;
        }

        if (_breakpoints.Remove(oldName, out var lines))
        {
            _breakpoints[newName] = lines;
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

    /// <summary>Whether the developer is looking at the Immediate panel.</summary>
    private bool _watchingImmediate;

    /// <summary>Remembers which panel is showing, and watches the output only when it is.</summary>
    private void OnPanelChanged(string name, bool open)
    {
        _watchingImmediate = open && name == "immediate";
        UpdatePolling();
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
            : 0;

        _editorSurface?.Poll(interval);
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
            if (stored is null)
            {
                return;
            }

            // Against what the module said last time, not against the surface. The two differ by
            // the editor's own reformatting from the moment anything is written, and that
            // difference is not a change anybody made.
            if (_writtenModules.TryGetValue(module, out var baseline) && baseline == stored)
            {
                return;
            }

            Log.Info($"resync: {module} changed outside the surface, adopting the module");
            _writtenModules[module] = stored;
            surface.Sync(module, stored);
            _analysis?.Reanalyse();
        }
        catch (Exception ex)
        {
            Log.Error($"resync: {module} could not be compared with the module", ex);
        }
    }

    /// <summary>One tick of the execution watch.</summary>
    private void PollDebugState()
    {
        _immediateReader?.Poll();
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
    /// Starts reading the Immediate window, having worked out which window it is.
    ///
    /// The one that stopped being visible when it was closed is the one that was closed. Its class
    /// is shared with the code panes and with the Locals and Watch windows, and its caption is
    /// localised, so there is nothing else to tell it apart by. It keeps its handle once hidden,
    /// which is what makes it readable afterwards.
    /// </summary>
    private void AttachImmediateReader(HashSet<nint> before)
    {
        before.ExceptWith(CodePaneTracker.VisiblePanes());

        if (before.Count != 1)
        {
            Log.Info($"immediate: {before.Count} windows changed when it closed, cannot tell which it is");
            return;
        }

        var window = before.First();
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
    private void EvaluateImmediate(string line)
    {
        _editorSurface?.FlushEdits();

        var evaluator = _immediate ??= new ImmediateEvaluator(_editor);
        var result = evaluator.Evaluate(line, _inBreak);

        _editorSurface?.ShowImmediateResult(result.Text, result.Failed);

        // Evaluating adds and removes a module, which the analyzer would otherwise report on.
        _analysis?.Reanalyse();
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

            case VbeCommands.Command.PropertiesWindow:
                _editorSurface?.RunEditorCommand("xlide.panel.properties");
                return true;

            case VbeCommands.Command.ProjectExplorer:
                _editorSurface?.Notify("The project explorer is part of the editor and always shown.");
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

        // Opening a module also selects it, the way the editor's own tree behaves.
        _propertiesTarget = component;

        _writtenModules[component] = source;
        _editorSurface?.Show(component, source);
        Log.Info($"editor surface: showing {component}, {source.Length} character(s)");

        // The findings for this module were computed before it was opened, so they are applied here
        // rather than waiting for the next analysis pass.
        PublishMarkersForShownModule();
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
        if (_windowsHidden)
        {
            return;
        }

        _windowsHidden = true;

        // The project explorer and the Immediate window have surface replacements. The properties
        // window does not yet; it is closed for the dock space it occupies, and the menu can bring
        // it back, at which point the surface retreats so it can be seen. The Locals and Watch
        // windows stay untouched: nothing replaces them yet, and hiding a window with no
        // replacement removes the feature rather than restyling it.
        const int immediateWindow = 5;
        ReadOnlySpan<int> replaced = [immediateWindow, 6, 7];

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
    /// Hides the editor's own toolbar, which the surface's own toolbar has replaced.
    ///
    /// Hidden only because everything on it is somewhere else now: saving, undo and redo, the object
    /// browser, and every run and step command are on the surface's toolbar or on the keys they
    /// always were. Hiding a bar whose commands had nowhere else to go would take those commands
    /// away, which is not a theme change.
    ///
    /// The menu bar is left alone. Its menus reach a great deal that has no replacement yet, and a
    /// consistent colour is not worth losing References, Options, or the object browser's menu.
    /// </summary>
    private void HideRedundantToolbar()
    {
        try
        {
            using var bars = _editor.GetObject("CommandBars");
            var count = bars?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var bar = bars!.GetItem(i);

                // The programmatic name, which is not localised the way a caption is.
                if (bar?.GetString("Name") != "Standard" || !bar.GetBool("Visible"))
                {
                    continue;
                }

                bar.SetBool("Visible", false);
                Log.Info("window: hid the editor's own toolbar");
                return;
            }
        }
        catch (Exception ex)
        {
            Log.Info($"window: the editor's toolbar could not be hidden ({ex.GetType().Name})");
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

        var before = CodePaneTracker.VisiblePanes();

        window.SetBool("Visible", false);
        Log.Info($"window: closed the editor's own '{window.GetString("Caption")}'");

        AttachImmediateReader(before);
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
    /// Retreating, it takes only the document area, and the native chrome shows again. That is the
    /// right way round: a visible seam is a smaller problem than covering a window or a toolbar the
    /// developer just asked for.
    /// </summary>
    private static unsafe PixelRect SurfaceBounds(nint frame, nint documentArea, bool covering)
    {
        var document = ClientAreaIn(documentArea, frame);

        if (!covering)
        {
            return document;
        }

        Rect client;
        if (!Win32.GetClientRect(frame, &client))
        {
            return document;
        }

        return new PixelRect(0, 0, client.Right - client.Left, client.Bottom - client.Top);
    }

    /// <summary>
    /// Whether the surface may cover everything native inside the frame.
    ///
    /// Three things say no. A page that has not loaded yet: covering the menu bar with a surface
    /// that cannot draw its own menus takes every menu away, so the native bar stays until the
    /// replacement is genuinely standing. A native tool window the developer opened: covering it
    /// would hide what they just asked for. And a docked toolbar they have shown, for the same
    /// reason.
    /// </summary>
    private bool CanCoverChrome() =>
        _editorSurface is { IsReady: true } && !AnyToolWindowOpen() && !AnyDockedToolbarVisible();

    /// <summary>
    /// Recomputes where the surface belongs, now.
    ///
    /// For the moments that are not window events: a menu item has just opened or closed a native
    /// window, or the page has just come up. Waiting for the next window event leaves the wrong
    /// thing covered in the meantime.
    /// </summary>
    private void RefreshSurfacePlacement()
    {
        if (_editorSurface is null || !_surfaceShown || _frame == 0 || _documentArea == 0)
        {
            return;
        }

        var covering = CanCoverChrome();
        _editorSurface.Follow(SurfaceBounds(_frame, _documentArea, covering), visible: true);
        _editorSurface.SetChrome(menuBar: covering);
    }

    /// <summary>
    /// Whether any of the editor's own docked windows is showing.
    ///
    /// Asked of the object model rather than worked out from windows on screen. Every one of these
    /// shares its window class with the code panes, so a window that is visible says nothing about
    /// what it is: an earlier version compared visible windows against the panes it was tracking
    /// and a second code pane, open but behind the first, read as a tool window every time.
    /// </summary>
    private bool AnyToolWindowOpen()
    {
        // Object browser, watches, locals, properties: the ones the surface has no replacement
        // for. Properties is closed at start-up for its dock space, but the menu can reopen it,
        // and reopened it must be seen.
        ReadOnlySpan<int> tools = [2, 3, 4, 7];

        try
        {
            using var windows = _editor.GetObject("Windows");
            var count = windows?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= count; i++)
            {
                using var window = windows!.GetItem(i);
                if (window is not null && tools.Contains(window.GetInt32("Type")) && window.GetBool("Visible"))
                {
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            // Erring towards the smaller rectangle, which covers less and hides nothing.
            Log.Info($"surface: the editor's windows could not be read ({ex.GetType().Name})");
            return true;
        }

        return false;
    }

    /// <summary>
    /// Whether the developer has a native toolbar docked and showing.
    ///
    /// The surface's own toolbar replaces the Standard bar, which is hidden, but the others can be
    /// shown from the View menu, and a docked bar occupies real rows at the frame's edge. Covering
    /// those rows would put a toolbar on screen that cannot be pressed.
    /// </summary>
    private bool AnyDockedToolbarVisible()
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

                // A floating bar floats above everything and contests nothing.
                var position = bar.GetInt32("Position");
                if (position != floating && position != popup)
                {
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            Log.Info($"surface: the editor's toolbars could not be read ({ex.GetType().Name})");
            return true;
        }

        return false;
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

        _immediateReader?.Dispose();
        _immediateReader = null;

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
