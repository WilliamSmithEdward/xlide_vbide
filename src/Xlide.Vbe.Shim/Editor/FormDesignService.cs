using Xlide.Vbe.Core.Forms;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The PRODUCT side of the form markup: a UserForm walked into Core's FormSpec and printed as
/// text, for the markup tab. The debug door's designer routes carry the rich diagnostic walk
/// (every colour, every font) and the mutations; this walk is the narrow one the markup layer
/// speaks - identity, containment, geometry, caption - because an unspoken property is one an
/// apply can never erase (docs/userform-designer.md, the markup layer).
///
/// Some plumbing here mirrors the debug partial's (the item basis probe, the tolerant reads).
/// Duplicated for now rather than shared, because the debug side lives inside #if DEBUG and a
/// Release build must carry THIS: the markup tab is product surface. Unifying them is part of
/// finishing M2.
/// </summary>
internal static class FormDesignService
{
    /// <summary>The form's markup, or null with a reason when the component has no designer.</summary>
    public static string? MarkupOf(DispatchObject component, string module, out string? reason)
    {
        reason = null;

        if (component.GetInt32("Type") != 3)
        {
            reason = $"{module} is not a UserForm";
            return null;
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            reason = $"{module} has no designer";
            return null;
        }

        var controls = new List<ControlSpec>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Walk(designer, null, controls, seen, 0);

        var spec = new FormSpec(
            module,
            TryText(designer, "Caption"),
            TryNumber(designer, "Width") ?? PropertyNumber(component, "Width"),
            TryNumber(designer, "Height") ?? PropertyNumber(component, "Height"),
            [],
            controls);

        return FormMarkup.Print(spec);
    }

    /// <summary>
    /// The walk knows each control's container because it is standing in it, so the narrow
    /// rows carry exact parents without a COM read per control. Recursion plus the dedupe
    /// makes flat and hierarchical collections the same world, as the debug walk's does.
    /// </summary>
    private static void Walk(
        DispatchObject container, string? parentName, List<ControlSpec> rows, HashSet<string> seen, int depth)
    {
        if (depth > 8)
        {
            return;
        }

        using var controls = container.GetObject("Controls");
        if (controls is null)
        {
            return;
        }

        foreach (var control in ItemsOf(controls))
        {
            using (control)
            {
                var name = TryText(control, "Name");
                if (name is null)
                {
                    continue;
                }

                if (seen.Add(name))
                {
                    string type;
                    try
                    {
                        type = FriendlyType(control.TypeName() ?? "Control");
                    }
                    catch
                    {
                        type = "Control";
                    }

                    rows.Add(new ControlSpec(
                        type, name, TryText(control, "Caption"),
                        TryNumber(control, "Left"), TryNumber(control, "Top"),
                        TryNumber(control, "Width"), TryNumber(control, "Height"),
                        parentName, []));
                }

                if (control.GetDispId("Pages") != DispId.Unknown)
                {
                    WalkPages(control, name, rows, seen, depth + 1);
                }
                else if (control.GetDispId("Controls") != DispId.Unknown)
                {
                    Walk(control, name, rows, seen, depth + 1);
                }
            }
        }
    }

    private static void WalkPages(
        DispatchObject multiPage, string multiPageName, List<ControlSpec> rows, HashSet<string> seen, int depth)
    {
        if (depth > 8)
        {
            return;
        }

        using var pages = multiPage.GetObject("Pages");
        if (pages is null)
        {
            return;
        }

        foreach (var page in ItemsOf(pages))
        {
            using (page)
            {
                var name = TryText(page, "Name");
                if (name is null)
                {
                    continue;
                }

                if (seen.Add(name))
                {
                    rows.Add(new ControlSpec(
                        "Page", name, TryText(page, "Caption"),
                        null, null, null, null, multiPageName, []));
                }

                Walk(page, name, rows, seen, depth + 1);
            }
        }
    }

    private static IEnumerable<DispatchObject> ItemsOf(DispatchObject collection)
    {
        int count;
        try
        {
            count = collection.GetInt32("Count");
        }
        catch
        {
            yield break;
        }

        if (count <= 0)
        {
            yield break;
        }

        var basis = 0;
        try
        {
            using var probe = collection.GetItem(0);
            if (probe is null)
            {
                basis = 1;
            }
        }
        catch
        {
            basis = 1;
        }

        for (var i = 0; i < count; i++)
        {
            DispatchObject? item = null;
            try
            {
                item = collection.GetItem(i + basis);
            }
            catch
            {
                // One unreadable item does not end the walk.
            }

            if (item is not null)
            {
                yield return item;
            }
        }
    }

    private static string? TryText(DispatchObject target, string name)
    {
        if (target.GetDispId(name) == DispId.Unknown)
        {
            return null;
        }

        try
        {
            return target.GetString(name);
        }
        catch
        {
            return null;
        }
    }

    private static double? TryNumber(DispatchObject target, string name)
    {
        if (target.GetDispId(name) == DispId.Unknown)
        {
            return null;
        }

        try
        {
            return target.GetDouble(name);
        }
        catch
        {
            return null;
        }
    }

    private static double? PropertyNumber(DispatchObject component, string name)
    {
        try
        {
            using var properties = component.GetObject("Properties");
            using var row = properties?.CallObject("Item", name);
            return row?.GetDouble("Value");
        }
        catch
        {
            return null;
        }
    }

    /// <summary>The toolbox name for an extender's internal interface, unknown names untouched.</summary>
    private static string FriendlyType(string raw) => raw switch
    {
        "ILabelControl" => "Label",
        "IMdcText" => "TextBox",
        "IMdcCombo" => "ComboBox",
        "IMdcList" => "ListBox",
        "IMdcCheckBox" => "CheckBox",
        "IMdcOptionButton" => "OptionButton",
        "IMdcToggleButton" => "ToggleButton",
        "IOptionFrame" => "Frame",
        "ICommandButton" => "CommandButton",
        "ITabStrip" => "TabStrip",
        "IMultiPage" => "MultiPage",
        "IPage" => "Page",
        "IScrollbar" or "IScrollBar" => "ScrollBar",
        "ISpinbutton" or "ISpinButton" => "SpinButton",
        "IImage" => "Image",
        _ => raw,
    };
}
