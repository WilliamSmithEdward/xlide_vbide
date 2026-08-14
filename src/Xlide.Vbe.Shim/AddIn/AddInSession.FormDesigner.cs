#if DEBUG
using System.Globalization;
using System.Text.Json;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The designer half of the debug api: a UserForm's control tree read as data, and the three
/// mutations a fixture is made of. This is docs/userform-designer.md's M1 instrument - the
/// route ships before any canvas exists, so every later claim about the designer has something
/// to be measured against.
///
/// Everything here runs on the host thread, called out of the on-host route switch the way
/// every session-touching route is. Reads are tolerant PER PROPERTY: controls differ - a Label
/// has a Caption and no Font.Size surprises, a TextBox has no Caption at all - and a missing
/// member is an answer (null), not a failure. A control that cannot even say its Name is
/// skipped rather than invented.
///
/// The walk recurses into containers (a Frame's Controls, a MultiPage's Pages and their
/// Controls) AND dedupes by name, deliberately: whether the forms runtime hands out a flat
/// collection or a hierarchical one, both worlds reduce to the same rows, with each control's
/// own Parent naming its true container. Form-wide control names are unique, which is what
/// makes the name a safe key.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>The form and every control, as JSON. The GET side of the designer route.</summary>
    private string DesignerRead(string module, string? projectDisplay)
    {
        var projectId = ResolveNamedProject(projectDisplay, out var complaint);
        if (complaint is not null)
        {
            return HostError(complaint);
        }

        using var component = FindComponent(module, projectId, out var foundProject);
        if (component is null)
        {
            return HostError($"no component named {module}");
        }

        if (component.GetInt32("Type") != 3)
        {
            return HostError($"{module} is not a UserForm, so it has no designer");
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            return HostError($"{module} has no designer to read");
        }

        var (form, rows) = CollectDesigner(component, designer);
        FormDesignService.KeepDesignerDown(component);

        return JsonSerializer.Serialize(
            new DebugDesignerReply(module, DisplayFromProjectId(foundProject), form, [.. rows]),
            DebugJsonContext.Default.DebugDesignerReply);
    }

    /// <summary>One walk, shared by the JSON read, the markup projection, and the apply diff.</summary>
    private static (DebugDesignerForm Form, List<DebugDesignerControl> Rows) CollectDesigner(
        DispatchObject component, DispatchObject designer)
    {
        // Width and Height are not the designer's: they live on the component's own Properties
        // collection, the one the native Properties window edits. InsideWidth and InsideHeight
        // ARE the designer's, and the pair differs by the frame chrome.
        var form = new DebugDesignerForm(
            TryText(designer, "Caption"),
            TryNumber(designer, "Width") ?? ComponentPropertyNumber(component, "Width"),
            TryNumber(designer, "Height") ?? ComponentPropertyNumber(component, "Height"),
            TryNumber(designer, "InsideWidth"),
            TryNumber(designer, "InsideHeight"),
            TryInt(designer, "BackColor"),
            TryInt(designer, "ForeColor"),
            TryInt(designer, "Zoom"));

        var rows = new List<DebugDesignerControl>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        WalkDesignerControls(designer, rows, seen, 0);
        return (form, rows);
    }

    /// <summary>
    /// The form as markup: the same walk, projected through Core's printer. The print
    /// vocabulary is deliberately narrow - identity, containment, geometry, caption - per
    /// docs/userform-designer.md: an unspoken property is one an apply can never erase.
    /// </summary>
    private string DesignerMarkup(string module, string? projectDisplay)
    {
        var projectId = ResolveNamedProject(projectDisplay, out var complaint);
        if (complaint is not null)
        {
            return HostError(complaint);
        }

        using var component = FindComponent(module, projectId, out var foundProject);
        if (component is null || component.GetInt32("Type") != 3)
        {
            return HostError($"{module} is not a UserForm of this project");
        }

        // The PRODUCT projection, not a route-local one: the tab, the apply and this route
        // must all print the same document, form properties included, and two projections is
        // exactly the drift the service exists to end.
        var markup = FormDesignService.MarkupOf(component, module, out var markupReason);
        if (markup is null)
        {
            return HostError(markupReason ?? $"{module} could not be projected");
        }

        return JsonSerializer.Serialize(
            new DebugDesignerMarkupReply(module, DisplayFromProjectId(foundProject), markup),
            DebugJsonContext.Default.DebugDesignerMarkupReply);
    }

    /// <summary>
    /// Applies a markup document to the live form as a NAME-KEYED DIFF: controls only in the
    /// markup are added, controls only in the model are removed, matched controls take their
    /// header geometry, caption and property lines. A changed type or container is a
    /// remove-and-add, which is also the truth of what it means. A document that does not
    /// parse applies NOTHING; a refusal partway reports exactly what landed first.
    /// </summary>
    private string DesignerApplyMarkup(string module, string? projectDisplay, string body)
    {
        var projectId = ResolveNamedProject(projectDisplay, out var complaint);
        if (complaint is not null)
        {
            return HostError(complaint);
        }

        using var component = FindComponent(module, projectId, out _);
        if (component is null || component.GetInt32("Type") != 3)
        {
            return HostError($"{module} is not a UserForm of this project");
        }

        // The machinery lives product-side now (FormDesignService.ApplyMarkup), because the
        // markup tab's Ctrl+S ships in Release and shares it. This route is a wrapper: same
        // operation by construction, its own JSON.
        var outcome = FormDesignService.ApplyMarkup(component, module, body);
        if (!outcome.ParseFailed)
        {
            RefreshDesignerTabFor(module);
        }

        if (outcome.ParseFailed)
        {
            return HostError(outcome.Refused!);
        }

        return JsonSerializer.Serialize(
            new DebugDesignerApplyReply(outcome.Ok, [.. outcome.Added], [.. outcome.Removed],
                outcome.Set, outcome.Refused, [.. outcome.Notes]),
            DebugJsonContext.Default.DebugDesignerApplyReply);
    }

    /// <summary>
    /// Adds a control through the designer model - the same call the native toolbox makes - and
    /// puts it where it was asked to go. An add that then cannot be placed is taken back out, so
    /// this either produces the control that was asked for or produces nothing; the lesson is
    /// the component route's, and it holds here for the same reason.
    /// </summary>
    private string DesignerAdd(
        string module, string? projectDisplay, string type, string? name, string? parent,
        double? left, double? top, double? width, double? height)
    {
        var projectId = ResolveNamedProject(projectDisplay, out var complaint);
        if (complaint is not null)
        {
            return HostError(complaint);
        }

        using var component = FindComponent(module, projectId, out _);
        if (component is null || component.GetInt32("Type") != 3)
        {
            return HostError($"{module} is not a UserForm of this project");
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            return HostError($"{module} has no designer");
        }

        var progId = FormDesignService.ProgIdFor(type);
        if (progId is null)
        {
            return HostError(
                $"'{type}' is not a control kind; pass a ProgID or one of label, textBox, comboBox, " +
                "listBox, checkBox, optionButton, toggleButton, frame, commandButton, tabStrip, " +
                "multiPage, scrollBar, spinButton, image");
        }

        using var owner = parent is { Length: > 0 }
            ? FormDesignService.FindContainerControls(designer, parent, 0)
            : designer.GetObject("Controls");
        if (owner is null)
        {
            return HostError(parent is { Length: > 0 }
                ? $"no container named {parent} on {module}; a MultiPage takes controls on its Pages, by page name"
                : $"{module}'s designer has no controls collection");
        }

        try
        {
            var (actualName, actualType) = FormDesignService.AddControl(owner, progId, name, left, top, width, height);
            Log.Info($"designer: added {actualType} '{actualName}' to {module}{(parent is { Length: > 0 } ? $" in {parent}" : "")}");
            FormDesignService.KeepDesignerDown(component);
            RefreshDesignerTabFor(module);
            return JsonSerializer.Serialize(
                new DebugDesignerEditReply(true, "add", actualName, actualType,
                    $"added to {(parent is { Length: > 0 } ? parent : module)}"),
                DebugJsonContext.Default.DebugDesignerEditReply);
        }
        catch (Exception ex)
        {
            return HostError(ex.Message);
        }
    }

    /// <summary>
    /// Removes a control through the collection that owns it. The owner is searched containers
    /// first, so with a flat top-level collection the true container still answers before the
    /// form does - Remove belongs to the collection the control actually lives in.
    /// </summary>
    private string DesignerRemove(string module, string? projectDisplay, string name)
    {
        var projectId = ResolveNamedProject(projectDisplay, out var complaint);
        if (complaint is not null)
        {
            return HostError(complaint);
        }

        using var component = FindComponent(module, projectId, out _);
        if (component is null || component.GetInt32("Type") != 3)
        {
            return HostError($"{module} is not a UserForm of this project");
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            return HostError($"{module} has no designer");
        }

        if (!FormDesignService.RemoveControl(designer, name))
        {
            return HostError($"no control named {name} on {module}");
        }

        Log.Info($"designer: removed '{name}' from {module}");
        FormDesignService.KeepDesignerDown(component);
        RefreshDesignerTabFor(module);
        return JsonSerializer.Serialize(
            new DebugDesignerEditReply(true, "remove", name, null, $"removed from {module}"),
            DebugJsonContext.Default.DebugDesignerEditReply);
    }

    /// <summary>
    /// Writes one property of a control, or of the form itself when no control is named, and
    /// answers what the property READS BACK - the Properties panel's convention, because the
    /// write proving only that it was asked for is the lie that route existed to end. One level
    /// of dotting reaches an object-valued property's member: Font.Size, Font.Bold.
    /// </summary>
    private string DesignerSet(
        string module, string? projectDisplay, string? name, string property, string value, string? asKind)
    {
        var projectId = ResolveNamedProject(projectDisplay, out var complaint);
        if (complaint is not null)
        {
            return HostError(complaint);
        }

        using var component = FindComponent(module, projectId, out _);
        if (component is null || component.GetInt32("Type") != 3)
        {
            return HostError($"{module} is not a UserForm of this project");
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            return HostError($"{module} has no designer");
        }

        var targetLabel = name is { Length: > 0 } ? name : module;
        try
        {
            var display = FormDesignService.SetControlProperty(component, designer, name, property, value, asKind);
            FormDesignService.KeepDesignerDown(component);
            RefreshDesignerTabFor(module);
            return JsonSerializer.Serialize(
                new DebugDesignerEditReply(true, "set", targetLabel, null, $"{targetLabel}.{property} is {display}"),
                DebugJsonContext.Default.DebugDesignerEditReply);
        }
        catch (Exception ex)
        {
            return HostError($"{targetLabel}.{property} refused the write ({ex.Message.Trim()})");
        }
    }

    /// <summary>A query number, invariant, or null when absent or not a number.</summary>
    private static double? DesignerNumber(string? text) =>
        double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? value : null;

    // ------------------------------------------------------------------ the walk

    private static void WalkDesignerControls(
        DispatchObject container, List<DebugDesignerControl> rows, HashSet<string> seen, int depth)
    {
        // Containers cannot honestly nest this deep; a cycle would.
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
                    rows.Add(ReadDesignerControl(control, name));
                }

                if (control.GetDispId("Pages") != DispId.Unknown)
                {
                    WalkDesignerPages(control, rows, seen, depth + 1);
                }
                else if (control.GetDispId("Controls") != DispId.Unknown)
                {
                    WalkDesignerControls(control, rows, seen, depth + 1);
                }
            }
        }
    }

    private static void WalkDesignerPages(
        DispatchObject multiPage, List<DebugDesignerControl> rows, HashSet<string> seen, int depth)
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
                    rows.Add(ReadDesignerControl(page, name));
                }

                WalkDesignerControls(page, rows, seen, depth + 1);
            }
        }
    }

    /// <summary>
    /// A collection's items in order, whichever end it counts from. The forms runtime indexes
    /// from zero and the editor's own collections from one; probing the first item once decides,
    /// and a probe that fails both ways is an empty answer rather than an error.
    /// </summary>
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

    private static DebugDesignerControl ReadDesignerControl(DispatchObject control, string name)
    {
        string type;
        try
        {
            type = FriendlyControlType(control.TypeName() ?? "Control");
        }
        catch
        {
            type = "Control";
        }

        return new DebugDesignerControl(
            name,
            type,
            TryParentName(control),
            TryNumber(control, "Left"),
            TryNumber(control, "Top"),
            TryNumber(control, "Width"),
            TryNumber(control, "Height"),
            TryInt(control, "TabIndex"),
            TryFlag(control, "Visible"),
            TryFlag(control, "Enabled"),
            TryText(control, "Caption"),
            TryInt(control, "ForeColor"),
            TryInt(control, "BackColor"),
            TryInt(control, "BorderStyle"),
            TryInt(control, "SpecialEffect"),
            TryText(control, "GroupName"),
            TryFont(control));
    }

    // ------------------------------------------------------------------ finding

    /// <summary>The Controls collection of the named container: a Frame's own, or a Page's by page name.</summary>
    // ------------------------------------------------------------------ tolerant reads

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

    private static int? TryInt(DispatchObject target, string name)
    {
        if (target.GetDispId(name) == DispId.Unknown)
        {
            return null;
        }

        try
        {
            return target.GetInt32(name);
        }
        catch
        {
            return null;
        }
    }

    private static bool? TryFlag(DispatchObject target, string name)
    {
        if (target.GetDispId(name) == DispId.Unknown)
        {
            return null;
        }

        try
        {
            return target.GetBool(name);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>A number off the component's Properties collection, the native panel's source.</summary>
    private static double? ComponentPropertyNumber(DispatchObject component, string name)
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

    private static DebugDesignerFont? TryFont(DispatchObject control)
    {
        if (control.GetDispId("Font") == DispId.Unknown)
        {
            return null;
        }

        try
        {
            using var font = control.GetObject("Font");
            return font is null
                ? null
                : new DebugDesignerFont(
                    TryText(font, "Name"), TryNumber(font, "Size"), TryFlag(font, "Bold"), TryFlag(font, "Italic"));
        }
        catch
        {
            return null;
        }
    }

    private static string? TryParentName(DispatchObject control)
    {
        if (control.GetDispId("Parent") == DispId.Unknown)
        {
            return null;
        }

        try
        {
            using var parent = control.GetObject("Parent");
            return parent is null ? null : TryText(parent, "Name");
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// The name a developer uses for what the type info calls something else. The extenders
    /// describe themselves by their internal interfaces - IMdcText, IOptionFrame - which are
    /// stable, documented-by-observation, and useless on a canvas. A name not in this table
    /// passes through untouched, so a third-party control stays honestly itself.
    /// </summary>
    private static string FriendlyControlType(string raw) => raw switch
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
#endif
