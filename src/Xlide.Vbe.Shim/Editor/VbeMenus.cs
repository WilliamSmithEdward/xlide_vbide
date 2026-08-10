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
    /// developer's standing rule (2026-08-04, extended 2026-08-05: the end goal is menus
    /// stripped to what only they can reach). Suppression is by id, which is safe exactly
    /// because every id here was MEASURED unique across the menu bar (2026-08-05 full
    /// enumeration; the one id that repeats everywhere, 746, is not here).
    ///
    /// The window replacements: 761 is every native toolbar's visibility toggle and 30045
    /// the Toolbars popup that held them — the surface's toolbar is the toolbar; 830 is the
    /// native window list, whose job the tab strip does; 2554 Immediate, 2555 Locals, 2557
    /// Project Explorer, 222 Properties Window are panels now. Watch (2556), Call Stack
    /// (620), and Object Browser (473) keep their items until their ports land — and the
    /// View menu goes ENTIRELY when Watch and Call Stack do (developer, 2026-08-05), with
    /// the UserForm designer's Toolbox (548) and Tab Order (469) moving to the designer
    /// backlog rather than surviving here.
    ///
    /// The Edit menu's editing half (developer, 2026-08-05: anything duplicative or already
    /// on the toolbar goes): Undo 128, Redo 129, Cut 21, Copy 19, Paste 22, Clear 478, and
    /// Select All 756 are the editor's own keys; Find 141, Find Next 570, and Replace 313
    /// are the find widget and the Search panel; Indent 15 and Outdent 14 are toolbar
    /// buttons and the Tab key itself; List Properties/Methods 2529, List Constants 2530,
    /// Quick Info 2531, Parameter Info 2532, and Complete Word 2533 are the engine's
    /// completions, hovers, and signature help. These also acted on the COVERED native
    /// pane, not on the surface — traps as much as duplicates. Bookmarks stays: nothing on
    /// the surface does its job yet.
    /// </summary>
    /// <remarks>
    /// 2026-08-05, the second half: the Edit and View menus are GONE ENTIRELY (30003, 30004 —
    /// the top-level popups themselves). Bookmarks became the surface's own (page-side, in the
    /// command palette and context menu); Watch became a panel fed by its own ghost palette;
    /// Call Stack sits on the toolbar's debug cluster; Definition and Last Position ride their
    /// native keys (Shift+F2, Ctrl+Shift+F2) and the editor's context menu; the Object Browser
    /// was already a toolbar button. The designer's entry points (View Object, Toolbox, Tab
    /// Order) return with the designer itself (#14), not with a menu.
    /// </remarks>
    /// <remarks>
    /// 2026-08-05, the Debug menu's watch items: Add Watch 1820, Edit Watch 940, and Quick
    /// Watch 229 are the Watch PANEL's own buttons now (developer: move the watch triggers
    /// into our UI and out of the menu bar). They still open the editor's dialogs — decision
    /// 11 keeps those, because driving a modal invisibly can hang the editor — but the panel
    /// is where a watch is read, so it is where the work belongs. Decision 11's "reachable
    /// forever" clause was about there being NO other route; the panel is that route, and
    /// these items may only be suppressed while it carries them.
    /// </remarks>
    /// <remarks>
    /// 2026-08-05, the last of that evening: the DEBUG menu (30165) and the FORMAT menu
    /// (30006) are gone entirely, and the bar reads File Insert Run Tools Add-Ins Window
    /// Help. Debug's commands all have homes — stepping, breakpoints, and the Call Stack
    /// were already toolbar buttons, the watches are the panel's, and Compile 578, Run To
    /// Cursor 1811, Set Next Statement 1812, and Show Next Statement 1813 joined the
    /// toolbar's debug cluster the same day (the last three greyed outside a break, the way
    /// the Call Stack button is). Format needed nothing: every item on it — Align 32787,
    /// Make Same Size 32790, Size to Fit 551, Size to Grid 550, Horizontal Spacing 32791,
    /// Vertical Spacing 32800, Center in Form 32789, Arrange Buttons 31215, Group 164,
    /// Ungroup 165, Order 32809 — arranges controls on a UserForm, and the designer that
    /// would give them meaning is backlog #14, which brings its own surfaces when it lands.
    /// </remarks>
    /// <remarks>
    /// 2026-08-06: the WINDOW menu (30009, from the same 2026-08-05 full enumeration) goes,
    /// and the bar reads File Insert Run Tools Add-Ins Help. Its window list (830) was
    /// already replaced by the tab strips; Split, Tile, Cascade, and Arrange Icons manage
    /// native MDI children the surface covers, and the surface's own editor groups — split
    /// right, split down, drag a tab between groups — are that job done where the developer
    /// actually looks (developer: "remove the window ribbon").
    /// </remarks>
    /// <remarks>
    /// 2026-08-09: the LAST FOUR top-level menus go, and what is left is folded into one. The bar
    /// read File Insert Run Tools Add-Ins Help; it reads `xlide` and nothing else.
    ///
    /// FILE (30002). Save is the toolbar's, Import File and Export File are the sync dialog and
    /// better than the native pair, Close and Return keeps its Alt+Q, Print the developer waved
    /// off. Remove was the blocker and it now lives on the module's own row in the tree, which is
    /// where it belonged: File's own Remove carries id 746, the id that repeats across every menu,
    /// so it could not have been suppressed on its own even if it had stayed.
    ///
    /// RUN (30012). Run 186, Break 189 and Reset 228 have been toolbar buttons for days. Design
    /// Mode 212 became one in the same change (the developer: "u can add a button for design mode
    /// to the bar"). Run Project 5415 is disabled in this host.
    ///
    /// HELP (30010). About xlide is the toolbar's. Native F1 goes with the menu, deliberately
    /// (the developer, 2026-08-09: "f1 native gone").
    ///
    /// INSERT (30005). Module 3039, Class Module 2579 and UserForm 512 are the plus button on
    /// every workbook row in the tree, and the workbook's own context menu. What genuinely goes
    /// with it is Procedure 559, the Add Procedure dialog; a developer types `Sub Name` instead.
    /// The designer entries wait for the designer, as the rest of them do.
    ///
    /// TOOLS (30007) and ADD-INS (30038) are suppressed as MENUS but not as items: the surface
    /// composes them into the one `xlide` menu below, which is the only thing left on the bar.
    /// </remarks>
    private static readonly HashSet<int> Replaced =
    [
        30003, 30004, 30006, 30165, 30009,
        30002, 30005, 30012, 30010,
        30007, 30038,
        // Additional Controls 642, off the xlide menu (the developer, 2026-08-09). It registers
        // ActiveX controls onto a UserForm's toolbox and is disabled anywhere else, which is
        // everywhere until the designer lands: an item that is grey every time it is seen teaches
        // a developer to stop reading the menu. It returns with the designer, backlog #14, with
        // the rest of that surface.
        642,
        // Options 522, off the xlide menu the same day: the toolbar's Settings is where a
        // developer goes to change how this editor behaves, and two dialogs answering that
        // question is one too many.
        //
        // WHAT IT ALSO HELD, recorded because a suppression that quietly drops a capability is
        // the failure this table is supposed to prevent. Most of the native Options dialog
        // governs the pane this surface covers — full module view, procedure separator, the
        // editor colours, docking — so it was already describing a window nobody looks at. Two
        // items are not that: Auto Syntax Check, which the analyzer's squiggles replace, and
        // REQUIRE VARIABLE DECLARATION, which puts `Option Explicit` at the top of every new
        // module and has no equivalent in xlide's settings. That one is a real gap, not a
        // rehoming, and it belongs in the settings dialog before anybody misses it.
        522,
        761, 830, 2554, 2555, 2557, 222, 30045,
        128, 129, 21, 19, 22, 478, 756, 141, 570, 313, 15, 14, 2529, 2530, 2531, 2532, 2533,
        1820, 940, 229,
    ];

    /// <summary>
    /// The one menu left, and it is not one of the editor's.
    ///
    /// Tools and Add-Ins between them hold six enabled items — References, Macros, Options,
    /// VBAProject Properties, Digital Signature, Add-In Manager — and none of them has anywhere
    /// else to be: they are the editor's own dialogs and this product does not reimplement them.
    /// Two menus for six items, next to nothing else, is a menu bar pretending to be a menu bar.
    /// So they are composed into one, and it carries the product's name because at that point it
    /// is the product's menu and not the editor's (the developer, 2026-08-09).
    ///
    /// HOW THE ADDRESSING SURVIVES IT. Items are addressed by their position chain and nothing
    /// else, which is the invariant this whole class rests on, so a synthetic menu has to have
    /// synthetic positions that lead back to real ones without ambiguity. A child of the xlide
    /// menu is numbered `rank * SourceStride + realPosition`, where rank is the source's place in
    /// XlideSources. That is arithmetic, not a remembered table: nothing is stored between the
    /// read that produced a path and the execute that uses it, so nothing can go stale in
    /// between, and the real position is preserved exactly as the rest of this class expects.
    /// </summary>
    private const int XlideMenuPosition = 900;

    /// <summary>One source's worth of positions. Far above any menu's item count.</summary>
    private const int SourceStride = 1000;

    /// <summary>Which menus the xlide menu is made of, in the order it shows them.</summary>
    private static readonly int[] XlideSources = [ToolsMenu, AddInsMenu];

    private const int ToolsMenu = 30007;
    private const int AddInsMenu = 30038;

    /// <summary>
    /// Reads the items of one menu: the bar itself for an empty path, or the submenu the path leads
    /// to. Items the editor is not showing are left out, but every item keeps its real position, so
    /// a path built from what is returned still addresses the right control.
    /// </summary>
    public static SurfaceMenuItem[] Read(DispatchObject editor, ReadOnlySpan<int> path)
    {
        ArgumentNullException.ThrowIfNull(editor);

        if (path.Length == 1 && path[0] == XlideMenuPosition)
        {
            return ReadXlideMenu(editor);
        }

        using var controls = ControlsAt(editor, Resolve(editor, path));
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

        // Everything the editor put on its bar is suppressed now, so the bar the surface draws is
        // this one item and nothing else. Appended rather than substituted for a real entry: it
        // stands for no single menu, and giving it one menu's position would make it look like it
        // did the next time somebody read this.
        if (path.Length == 0)
        {
            // A wrench, not a word (the developer, 2026-08-09). The caption is still carried: it
            // is the accessible name and the tooltip, and what the page falls back to if the icon
            // ever goes missing. VBA rather than xlide, because it says what is BEHIND the button
            // — the editor's own dialogs — and the whole surface is xlide already.
            items.Add(new SurfaceMenuItem(
                XlideMenuPosition, "VBA", Enabled: true, Separator: false, Popup: true,
                Checked: false, Shortcut: null, Icon: "wrench"));
        }

        return [.. items];
    }

    /// <summary>
    /// The xlide menu's items: every source menu's controls in turn, each numbered into its own
    /// band, with a divider where one source ends and the next begins.
    /// </summary>
    private static SurfaceMenuItem[] ReadXlideMenu(DispatchObject editor)
    {
        var items = new List<SurfaceMenuItem>();

        for (var rank = 0; rank < XlideSources.Length; rank++)
        {
            var at = PositionOf(editor, XlideSources[rank]);
            if (at == 0)
            {
                continue;
            }

            using var controls = ControlsAt(editor, [at]);
            var count = controls?.GetInt32("Count") ?? 0;
            var firstOfSource = true;

            for (var i = 1; i <= count; i++)
            {
                using var control = controls!.GetItem(i);
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
                    (rank * SourceStride) + i,
                    caption,
                    TryBool(control, "Enabled", fallback: true),
                    // The editor's own grouping within a source, and a divider between sources:
                    // the menu is a join, and a reader should be able to see the seam.
                    (firstOfSource && items.Count > 0) || TryBool(control, "BeginGroup", fallback: false),
                    popup,
                    !popup && TryInt(control, "State", fallback: 0) != 0,
                    shortcut));

                firstOfSource = false;
            }
        }

        return [.. items];
    }

    /// <summary>
    /// A path as the editor would address it. Chains that do not start in the xlide menu are
    /// already real and are handed back untouched.
    /// </summary>
    private static int[] Resolve(DispatchObject editor, ReadOnlySpan<int> path)
    {
        if (path.Length < 2 || path[0] != XlideMenuPosition)
        {
            return path.ToArray();
        }

        var rank = path[1] / SourceStride;
        var position = path[1] % SourceStride;
        if (rank < 0 || rank >= XlideSources.Length || position <= 0)
        {
            return path.ToArray();
        }

        var at = PositionOf(editor, XlideSources[rank]);
        if (at == 0)
        {
            return path.ToArray();
        }

        // [900, rank*1000 + i, rest...] becomes [realTopLevel, i, rest...]: the same depth, so
        // everything downstream counts levels the way it always did.
        var real = new int[path.Length];
        real[0] = at;
        real[1] = position;
        for (var i = 2; i < path.Length; i++)
        {
            real[i] = path[i];
        }

        return real;
    }

    /// <summary>
    /// Where a top-level menu sits on the bar right now, or 0 when it is not there.
    ///
    /// By id rather than by a remembered position: the bar's contents are the host's business and
    /// a hard-coded position is a bet on a layout nobody promised.
    /// </summary>
    private static int PositionOf(DispatchObject editor, int id)
    {
        using var controls = ControlsAt(editor, []);
        var count = controls?.GetInt32("Count") ?? 0;

        for (var i = 1; i <= count; i++)
        {
            using var control = controls!.GetItem(i);
            if (control is not null && TryInt(control, "Id", fallback: 0) == id)
            {
                return i;
            }
        }

        return 0;
    }

    /// <summary>
    /// Every control at a path, INCLUDING the ones the surface suppresses, each with its id.
    ///
    /// What Read answers is the product's menu; this is the editor's. The difference between the
    /// two is the only place a suppression can be checked, and until this existed there was no way
    /// to ask the running editor for a menu's ids at all — they were measured once by hand and
    /// written into a comment, which is how a table of numbers goes quietly out of date.
    ///
    /// Debug only, through the api. Nothing in the product reads it.
    /// </summary>
    public static (int Index, int Id, string Caption, bool Popup, bool Enabled, bool Suppressed)[]
        Describe(DispatchObject editor, ReadOnlySpan<int> path)
    {
        ArgumentNullException.ThrowIfNull(editor);

        // The xlide menu is not one of the editor's, so there is nothing at that position to walk
        // to and asking the bar for control 900 is a bare HRESULT. Answered from the composition
        // instead, which is the more useful answer anyway: it is the one place a reader can see
        // which real control each synthetic position leads to.
        if (path.Length == 1 && path[0] == XlideMenuPosition)
        {
            var composed = new List<(int, int, string, bool, bool, bool)>();
            foreach (var item in ReadXlideMenu(editor))
            {
                using var control = ControlAt(editor, [XlideMenuPosition, item.Index]);
                composed.Add((
                    item.Index,
                    control is null ? 0 : TryInt(control, "Id", fallback: 0),
                    item.Caption,
                    item.Popup,
                    item.Enabled,
                    false));
            }

            return [.. composed];
        }

        using var controls = ControlsAt(editor, Resolve(editor, path));
        if (controls is null)
        {
            return [];
        }

        var count = controls.GetInt32("Count");
        var rows = new List<(int, int, string, bool, bool, bool)>(count);

        for (var i = 1; i <= count; i++)
        {
            using var control = controls.GetItem(i);
            if (control is null)
            {
                continue;
            }

            var id = TryInt(control, "Id", fallback: 0);
            rows.Add((
                i,
                id,
                TryString(control, "Caption") ?? string.Empty,
                TryInt(control, "Type", fallback: 0) == PopupControlType,
                TryBool(control, "Enabled", fallback: true),
                Replaced.Contains(id)));
        }

        return [.. rows];
    }

    /// <summary>The control a non-empty path leads to, or null when the path no longer exists.</summary>
    public static DispatchObject? ControlAt(DispatchObject editor, ReadOnlySpan<int> path)
    {
        ArgumentNullException.ThrowIfNull(editor);

        if (path.Length == 0)
        {
            return null;
        }

        // Resolved WHOLE and before anything is walked. Doing it inside ControlsAt would translate
        // the parents and leave the leaf synthetic, which is the one position that names the item
        // about to be executed.
        var real = Resolve(editor, path);
        using var controls = ControlsAt(editor, real.AsSpan()[..^1]);
        return controls?.GetItem(real[^1]);
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
