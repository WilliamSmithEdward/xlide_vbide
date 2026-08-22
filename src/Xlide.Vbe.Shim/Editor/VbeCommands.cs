using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Runs the editor's own commands: running a procedure, stepping, breaking, toggling a breakpoint.
///
/// The editor exposes none of these on its object model. There is no Run method and no debugger
/// object. What it does expose is the command bars its menus are built from, and every menu item on
/// them can be executed. So this drives the editor exactly as the menu does, which means the editor
/// does the work in its own way and everything that depends on it keeps working.
///
/// Commands are found by identifier, never by caption. Captions are localised, so matching them
/// would work on one machine and silently do nothing on the next.
/// </summary>
internal static class VbeCommands
{
    /// <summary>
    /// Bars searched first, by their programmatic names, which are not localised the way captions
    /// are. Every command this drives sits on one of these, so the search almost never goes further.
    /// </summary>
    private static readonly string[] PreferredBars = ["Debug", "Edit"];

    /// <summary>
    /// Identifiers of the editor commands this drives.
    ///
    /// These are established by enumerating the editor's own menus, not taken from documentation.
    /// A `Describe` here used to be the instrument that did it, writing every identifier and
    /// caption to the log; <see cref="VbeMenus.Describe"/> superseded it and reports suppression
    /// as well as identity, and the `menus` debug route serves the same reading over the door.
    /// Re-read them there if a host build ever disagrees with the values below.
    /// </summary>
    public static class Command
    {
        public const int Run = 186;
        public const int Break = 189;
        public const int Reset = 228;

        /// <summary>
        /// Design Mode, the Run menu's one item that was not already a button. Measured off a
        /// running editor on 2026-08-09, the day the Run menu was retired; a toggle, so pressing
        /// it twice is where it started.
        /// </summary>
        public const int DesignMode = 212;
        public const int ToggleBreakpoint = 51;
        public const int StepInto = 188;
        public const int StepOver = 194;
        public const int StepOut = 2559;
        public const int RunToCursor = 1811;
        public const int Compile = 578;
        public const int SetNextStatement = 1812;
        public const int ShowNextStatement = 1813;
        public const int QuickWatch = 229;
        public const int AddWatch = 1820;
        public const int EditWatch = 940;
        public const int CallStack = 620;
        public const int CommentBlock = 192;
        public const int Save = 3;
        public const int ObjectBrowser = 473;
        public const int UncommentBlock = 2552;
        public const int Undo = 128;
        public const int Redo = 129;
        public const int Find = 141;
        public const int Replace = 313;
        public const int ClearAllBreakpoints = 579;
        public const int ImmediateWindow = 2554;
        public const int LocalsWindow = 2555;
        public const int ProjectExplorer = 2557;
        public const int PropertiesWindow = 222;
        public const int References = 942;
        public const int ProjectProperties = 2578;
        public const int Options = 522;
    }

    /// <summary>
    /// Whether a command must run through the session's own command path rather than being executed
    /// where it was found. Toggling a breakpoint is bookkeeping as well as a command, and running,
    /// stepping and saving need the developer's edits and caret to reach the editor first. A menu
    /// item that executes one of these directly does the command and skips all of that, which is
    /// exactly how an invisible breakpoint was set once before.
    /// </summary>
    public static bool RoutesThroughSession(int id) => id is
        Command.Run or Command.Break or Command.Reset or Command.ToggleBreakpoint or
        Command.StepInto or Command.StepOver or Command.StepOut or Command.RunToCursor or
        Command.Save;

    /// <summary>
    /// What became of a command. Most callers want <see cref="CommandRun.Ran"/> and nothing else,
    /// but a script driving the editor through the xlide api needs the rest: "this host has no
    /// such command" and "the item was greyed" are different answers, and a caller handed only
    /// false cannot tell them apart, nor tell either from an exception.
    /// </summary>
    public readonly record struct CommandRun(bool Ran, string Detail)
    {
        public static CommandRun Ok(string detail = "executed") => new(true, detail);

        public static CommandRun No(string detail) => new(false, detail);
    }

    /// <summary>
    /// Executes a command by identifier.
    /// </summary>
    /// <returns>Whether it ran, and when it did not, why not.</returns>
    public static CommandRun Execute(DispatchObject editor, int commandId)
    {
        ArgumentNullException.ThrowIfNull(editor);

        try
        {
            using var bars = editor.GetObject("CommandBars");
            if (bars is null)
            {
                Log.Warn("command: the editor exposed no command bars");
                return CommandRun.No("the editor exposed no command bars");
            }

            // The bars are walked rather than searched.
            //
            // The collection has a search of its own, and it does not find any of these. It looks
            // only at the top level of each bar unless told otherwise, and every command here is a
            // menu item, which is one level further in. It also returned nothing rather than
            // failing, so using it looked exactly like a host that has no Run command.
            using var control = Find(bars, commandId);
            if (control is null)
            {
                Log.Info($"command: {commandId} is not present in this host");
                return CommandRun.No("not present in this host");
            }

            // A command that cannot run right now is not a failure worth reporting as one: Break is
            // disabled unless something is running, and Reset unless something is stopped.
            if (!control.GetBool("Enabled"))
            {
                Log.Info($"command: {commandId} is currently disabled");
                return CommandRun.No("currently disabled");
            }

            control.Invoke("Execute");
            Log.Info($"command: {commandId} executed");
            return CommandRun.Ok();
        }
        catch (Exception ex)
        {
            Log.Error($"command: {commandId} could not be executed", ex);
            return CommandRun.No($"raised {ex.GetType().Name}");
        }
    }

    /// <summary>Finds a control by identifier, looking at the likeliest bars first.</summary>
    private static DispatchObject? Find(DispatchObject bars, int commandId)
    {
        var count = bars.GetInt32("Count");

        // Two passes: the bars that carry these commands, then everything else. The first pass
        // almost always answers, and the second means a host that arranges its menus differently
        // still works rather than silently doing nothing.
        for (var pass = 0; pass < 2; pass++)
        {
            for (var i = 1; i <= count; i++)
            {
                using var bar = bars.GetItem(i);
                if (bar is null)
                {
                    continue;
                }

                var preferred = Array.IndexOf(PreferredBars, bar.GetString("Name")) >= 0;
                if (preferred != (pass == 0))
                {
                    continue;
                }

                if (FindOn(bar, commandId) is { } found)
                {
                    return found;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Finds a control by identifier on a bar, descending into its menus.
    ///
    /// The descent is what lets commands that live only inside a menu (References, Project
    /// Properties) be found at all. It is only sound for identifiers that are unique across the
    /// menus: several are shared by unrelated items, and looking one of those up would find
    /// whichever came first. Everything in <see cref="Command"/> is unique; menu replication,
    /// which cannot make that promise, addresses by position instead.
    /// </summary>
    private static DispatchObject? FindOn(DispatchObject bar, int commandId, int depth = 0)
    {
        const int popupControl = 10;
        const int deepestMenu = 3;

        using var controls = bar.GetObject("Controls");
        var count = controls?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            var control = controls!.GetItem(i);
            if (control is null)
            {
                continue;
            }

            if (control.GetInt32("Id") == commandId)
            {
                return control;
            }

            if (depth < deepestMenu)
            {
                int controlType;
                try
                {
                    controlType = control.GetInt32("Type");
                }
                catch (Exception)
                {
                    controlType = 0;
                }

                if (controlType == popupControl && FindOn(control, commandId, depth + 1) is { } found)
                {
                    control.Dispose();
                    return found;
                }
            }

            control.Dispose();
        }

        return null;
    }

    /// <summary>
    /// The editor command a toolbar name means, or zero when it is not one of ours.
    ///
    /// Names rather than numbers cross the boundary, so the page never has to know the host's
    /// identifiers and a host that numbers them differently changes nothing on the far side.
    /// </summary>
    public static int ForName(string name) => name switch
    {
        "save" => Command.Save,
        "objectBrowser" => Command.ObjectBrowser,
        "run" => Command.Run,
        "break" => Command.Break,
        "reset" => Command.Reset,
        "designMode" => Command.DesignMode,
        "stepInto" => Command.StepInto,
        "stepOver" => Command.StepOver,
        "stepOut" => Command.StepOut,
        "runToCursor" => Command.RunToCursor,
        "compile" => Command.Compile,
        "setNextStatement" => Command.SetNextStatement,
        "showNextStatement" => Command.ShowNextStatement,
        "toggleBreakpoint" => Command.ToggleBreakpoint,
        "quickWatch" => Command.QuickWatch,
        "addWatch" => Command.AddWatch,
        "editWatch" => Command.EditWatch,
        "callStack" => Command.CallStack,
        // No goToDefinition or lastPosition: navigating code is the surface's now, and the two
        // host commands they used to reach are gone with them (2026-08-06).
        "clearAllBreakpoints" => Command.ClearAllBreakpoints,
        "references" => Command.References,
        "projectProperties" => Command.ProjectProperties,
        _ => 0,
    };

    /// <summary>
    /// The surface command a keystroke means, or null when the surface does not own that key.
    ///
    /// These are keys that would otherwise be taken before the page could see them. F1 is the
    /// host's help, and Ctrl+PageDown belongs to the browser, which treats it as its own tab
    /// switching and swallows it whole; the only way the page's tabs can have the key is to claim
    /// it here and say what it meant.
    /// </summary>
    public static string? SurfaceCommandForKey(uint virtualKey, bool shift, bool control) => virtualKey switch
    {
        VirtualKey.F1 when !shift && !control => "editor.action.quickCommand",
        // Closing a tab is the PAGE's decision, because the page is what has tabs. The host used
        // to decide here, closing whatever module it believed was shown, and with two workbooks
        // open its belief drifted: it held a null module and a project from the other workbook,
        // so the key was claimed and nothing closed (the developer, 2026-08-07).
        VirtualKey.W when control && !shift => "xlide.tab.close",
        VirtualKey.F4 when control && !shift => "xlide.tab.close",
        VirtualKey.PageDown when control && !shift => "xlide.tab.next",
        VirtualKey.PageUp when control && !shift => "xlide.tab.previous",
        _ => null,
    };

    /// <summary>
    /// The editor command a keystroke means, or zero when the editor does not own that key.
    ///
    /// These are the editor's own shortcuts, and they are claimed because the surface covers the
    /// pane the editor would otherwise have received them through. A developer who presses F5
    /// expects the procedure to run, and it stopping working is not a trade anyone agreed to.
    ///
    /// Only keys the editor owns are taken. Everything else, including every editing key and every
    /// shortcut the surface itself defines, is left for the document.
    /// </summary>
    public static int ForKey(uint virtualKey, bool shift, bool control)
    {
        return virtualKey switch
        {
            // Shift+F2 and Ctrl+Shift+F2 are NOT claimed. They were the host's navigation pair
            // until the surface grew its own (2026-08-06), and the surface's knows more: it
            // crosses modules, resolves members reached through a receiver, and reads the text as
            // typed rather than as last written back. Claiming a key here takes it before the
            // page can see it, so the keys are left alone and the surface binds them. Plain F2 is
            // the Object Browser, as it always was.
            VirtualKey.F2 when shift => 0,
            VirtualKey.F2 => Command.ObjectBrowser,
            VirtualKey.F5 when control => Command.Break,
            VirtualKey.F5 when shift => Command.Reset,
            VirtualKey.F5 => Command.Run,
            VirtualKey.F8 when control => Command.RunToCursor,
            VirtualKey.F8 when shift => Command.StepOver,
            VirtualKey.F8 => Command.StepInto,
            VirtualKey.F9 => Command.ToggleBreakpoint,

            // Saving belongs to the host: the workbook is what gets written, not the module.
            VirtualKey.S when control => Command.Save,
            _ => 0,
        };
    }
}

/// <summary>The few virtual key codes this needs, so the numbers do not appear bare in a switch.</summary>
internal static class VirtualKey
{
    public const uint F1 = 0x70;
    public const uint F2 = 0x71;
    public const uint F4 = 0x73;
    public const uint F5 = 0x74;
    public const uint F8 = 0x77;
    public const uint F9 = 0x78;
    public const uint PageUp = 0x21;
    public const uint PageDown = 0x22;
    public const uint S = 0x53;
    public const uint W = 0x57;
}
