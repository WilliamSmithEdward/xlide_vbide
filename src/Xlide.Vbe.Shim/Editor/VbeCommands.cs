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
    /// <summary>Command bar control type for a plain button, which is what a menu item is.</summary>
    private const int ButtonControl = 1;

    /// <summary>
    /// Identifiers of the editor commands this drives.
    ///
    /// These are established by enumerating the editor's own menus, not taken from documentation:
    /// <see cref="Describe"/> writes each menu item's identifier and caption to the log, and the
    /// values here are what it reported. Re-run it if a host build ever disagrees.
    /// </summary>
    public static class Command
    {
        public const int Run = 186;
        public const int Break = 189;
        public const int Reset = 228;
        public const int ToggleBreakpoint = 51;
        public const int StepInto = 188;
        public const int StepOver = 194;
        public const int StepOut = 2559;
        public const int RunToCursor = 1811;
        public const int QuickWatch = 229;
        public const int AddWatch = 1820;
        public const int CallStack = 620;
        public const int CommentBlock = 192;
        public const int UncommentBlock = 2552;
    }

    /// <summary>
    /// Executes a command by identifier.
    /// </summary>
    /// <returns>False when the editor has no such command, or it is currently unavailable.</returns>
    public static bool Execute(DispatchObject editor, int commandId)
    {
        ArgumentNullException.ThrowIfNull(editor);

        try
        {
            using var bars = editor.GetObject("CommandBars");
            if (bars is null)
            {
                Log.Warn("command: the editor exposed no command bars");
                return false;
            }

            // Searched across every bar rather than a named one, because the same command appears on
            // a menu and on a toolbar and either will do.
            using var control = bars.CallObject("FindControl", ButtonControl, commandId);
            if (control is null)
            {
                Log.Info($"command: {commandId} is not present in this host");
                return false;
            }

            // A command that cannot run right now is not a failure worth reporting as one: Break is
            // disabled unless something is running, and Reset unless something is stopped.
            if (control.GetInt32("Enabled") == 0)
            {
                Log.Info($"command: {commandId} is currently disabled");
                return false;
            }

            control.Invoke("Execute");
            return true;
        }
        catch (Exception ex)
        {
            Log.Error($"command: {commandId} could not be executed", ex);
            return false;
        }
    }

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
            VirtualKey.F5 when control => Command.Break,
            VirtualKey.F5 when shift => Command.Reset,
            VirtualKey.F5 => Command.Run,
            VirtualKey.F8 when control => Command.RunToCursor,
            VirtualKey.F8 when shift => Command.StepOver,
            VirtualKey.F8 => Command.StepInto,
            VirtualKey.F9 => Command.ToggleBreakpoint,
            _ => 0,
        };
    }

    /// <summary>
    /// Writes every menu item's identifier and caption to the log.
    ///
    /// This is how the identifiers above were established, and it is the way to re-establish them
    /// against a host that disagrees. It is not called during normal operation.
    /// </summary>
    public static void Describe(DispatchObject editor)
    {
        ArgumentNullException.ThrowIfNull(editor);

        try
        {
            using var bars = editor.GetObject("CommandBars");
            var barCount = bars?.GetInt32("Count") ?? 0;

            for (var i = 1; i <= barCount; i++)
            {
                using var bar = bars!.GetItem(i);
                var name = bar?.GetString("Name");
                if (bar is null || name is null)
                {
                    continue;
                }

                Log.Info($"command bar {i}: '{name}'");

                using var controls = bar.GetObject("Controls");
                var controlCount = controls?.GetInt32("Count") ?? 0;

                for (var j = 1; j <= controlCount; j++)
                {
                    using var control = controls!.GetItem(j);
                    if (control is null)
                    {
                        continue;
                    }

                    Log.Info($"    {control.GetInt32("Id")} '{control.GetString("Caption")}'");
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error("command: the editor's menus could not be enumerated", ex);
        }
    }
}

/// <summary>The few virtual key codes this needs, so the numbers do not appear bare in a switch.</summary>
internal static class VirtualKey
{
    public const uint F5 = 0x74;
    public const uint F8 = 0x77;
    public const uint F9 = 0x78;
}
