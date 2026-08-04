using Xlide.Vbe.Shim.Com;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Reads and addresses the editor's menus, so the surface can draw its own menu bar without losing
/// anything the native one reaches.
///
/// The items are read live each time a menu opens rather than once at start-up, because the menus
/// are not static: enablement follows the execution state, checkmarks follow the layout, and the
/// Window menu lists whatever is open at the moment it drops down. Reading live is also what makes
/// the surface's menu complete by construction, which it must be: once the native bar is covered,
/// this menu is the only route to everything on it.
///
/// Items are addressed by their position chain, never by identifier. Identifiers are not unique
/// across the menus (several repeat), so an identifier does not name one item, but a chain of
/// positions does.
/// </summary>
internal static class VbeMenus
{
    /// <summary>msoBarTypeMenuBar: the one bar that is the menu bar.</summary>
    private const int MenuBarType = 1;

    /// <summary>msoControlPopup: a control that opens a submenu rather than doing something.</summary>
    private const int PopupControlType = 10;

    /// <summary>
    /// The keys the surface itself claims, by the command they run.
    ///
    /// Shown in the menu because the native items do not carry their shortcut text anywhere
    /// readable, and because where the surface's binding differs from the native one, the menu must
    /// teach the binding that actually works while the surface has focus.
    /// </summary>
    private static readonly Dictionary<int, string> ClaimedKeys = new()
    {
        [VbeCommands.Command.Run] = "F5",
        [VbeCommands.Command.Break] = "Ctrl+F5",
        [VbeCommands.Command.Reset] = "Shift+F5",
        [VbeCommands.Command.StepInto] = "F8",
        [VbeCommands.Command.StepOver] = "Shift+F8",
        [VbeCommands.Command.RunToCursor] = "Ctrl+F8",
        [VbeCommands.Command.ToggleBreakpoint] = "F9",
        [VbeCommands.Command.Save] = "Ctrl+S",
    };

    /// <summary>
    /// Items the surface has replaced outright, left out of every menu it serves — the
    /// developer's standing rule (2026-08-04): as a native window is ported fully, its menu
    /// item goes. 761 is every native toolbar's visibility toggle — the native toolbars are
    /// hidden and the surface's toolbar is the toolbar — and 830 is every entry of the native
    /// window list, whose job the tab strip does. 2554 Immediate, 2555 Locals, 2557 Project
    /// Explorer, and 222 Properties Window name windows whose surface replacements are always
    /// a panel tab away; the Watch window, Call Stack, and Object Browser keep their items
    /// until their ports land. Each id is shared across exactly its family and nothing else
    /// (measured; the handoff's command table), which is what makes suppression by id safe.
    /// </summary>
    private static readonly HashSet<int> Replaced = [761, 830, 2554, 2555, 2557, 222];

    /// <summary>
    /// Reads the items of one menu: the bar itself for an empty path, or the submenu the path leads
    /// to. Items the editor is not showing are left out, but every item keeps its real position, so
    /// a path built from what is returned still addresses the right control.
    /// </summary>
    public static SurfaceMenuItem[] Read(DispatchObject editor, ReadOnlySpan<int> path)
    {
        ArgumentNullException.ThrowIfNull(editor);

        using var controls = ControlsAt(editor, path);
        if (controls is null)
        {
            return [];
        }

        var count = controls.GetInt32("Count");
        var items = new List<SurfaceMenuItem>(count);

        for (var i = 1; i <= count; i++)
        {
            using var control = controls.GetItem(i);
            if (control is null || !TryBool(control, "Visible", fallback: true))
            {
                continue;
            }

            var caption = TryString(control, "Caption");
            if (string.IsNullOrEmpty(caption))
            {
                continue;
            }

            var id = TryInt(control, "Id", fallback: 0);
            if (Replaced.Contains(id))
            {
                continue;
            }

            var popup = TryInt(control, "Type", fallback: 0) == PopupControlType;
            var shortcut = ExtractShortcut(ref caption, control, id, popup);

            items.Add(new SurfaceMenuItem(
                i,
                caption,
                TryBool(control, "Enabled", fallback: true),
                TryBool(control, "BeginGroup", fallback: false),
                popup,
                // Checked is a button state; a popup has no state and the read is refused, which
                // the fallback turns into the right answer.
                !popup && TryInt(control, "State", fallback: 0) != 0,
                shortcut));
        }

        return [.. items];
    }

    /// <summary>The control a non-empty path leads to, or null when the path no longer exists.</summary>
    public static DispatchObject? ControlAt(DispatchObject editor, ReadOnlySpan<int> path)
    {
        ArgumentNullException.ThrowIfNull(editor);

        if (path.Length == 0)
        {
            return null;
        }

        using var controls = ControlsAt(editor, path[..^1]);
        return controls?.GetItem(path[^1]);
    }

    /// <summary>
    /// The control collection a path names: the menu bar's own controls for an empty path, or the
    /// controls of the popup the path walks to.
    /// </summary>
    private static DispatchObject? ControlsAt(DispatchObject editor, ReadOnlySpan<int> path)
    {
        using var bar = FindMenuBar(editor);
        var current = bar?.GetObject("Controls");

        foreach (var index in path)
        {
            if (current is null)
            {
                return null;
            }

            using var owner = current;
            using var control = owner.GetItem(index);
            current = control?.GetObject("Controls");
        }

        return current;
    }

    /// <summary>
    /// Finds the menu bar by what it is rather than what it is called. Its name is a fixed string
    /// today, but its type is what makes it the menu bar, and the type is not a string at all.
    /// Also consulted for its height, which is where the loader's ground begins.
    /// </summary>
    internal static DispatchObject? FindMenuBar(DispatchObject editor)
    {
        using var bars = editor.GetObject("CommandBars");
        var count = bars?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            var bar = bars!.GetItem(i);
            if (bar is not null && TryInt(bar, "Type", fallback: -1) == MenuBarType)
            {
                return bar;
            }

            bar?.Dispose();
        }

        return null;
    }

    /// <summary>
    /// Separates a shortcut from a caption, wherever the editor put it.
    ///
    /// A caption can embed its shortcut after a tab; a custom control can carry one as a property,
    /// which a built-in control refuses to answer; and the keys the surface claims for itself are
    /// known here and shown for the commands they run, because they are the ones that work.
    /// </summary>
    private static string? ExtractShortcut(ref string caption, DispatchObject control, int id, bool popup)
    {
        var tab = caption.IndexOf('\t');
        if (tab >= 0)
        {
            var embedded = caption[(tab + 1)..];
            caption = caption[..tab];
            return embedded.Length > 0 ? embedded : null;
        }

        if (popup)
        {
            return null;
        }

        if (TryString(control, "ShortcutText") is { Length: > 0 } declared)
        {
            return declared;
        }

        return ClaimedKeys.TryGetValue(id, out var claimed) ? claimed : null;
    }

    /*
     * Property reads that tolerate refusal. A control answers only the questions that make sense
     * for what it is — a popup has no State, a built-in button no ShortcutText — and a refusal is
     * an answer here, not a fault.
     */

    private static bool TryBool(DispatchObject control, string name, bool fallback)
    {
        try
        {
            return control.GetBool(name);
        }
        catch (Exception)
        {
            return fallback;
        }
    }

    private static int TryInt(DispatchObject control, string name, int fallback)
    {
        try
        {
            return control.GetInt32(name);
        }
        catch (Exception)
        {
            return fallback;
        }
    }

    private static string? TryString(DispatchObject control, string name)
    {
        try
        {
            return control.GetString(name);
        }
        catch (Exception)
        {
            return null;
        }
    }
}
