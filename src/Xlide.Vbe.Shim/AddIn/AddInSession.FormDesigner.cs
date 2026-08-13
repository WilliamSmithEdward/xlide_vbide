#if DEBUG
using System.Globalization;
using System.Text.Json;
using Xlide.Vbe.Shim.Com;
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

        return JsonSerializer.Serialize(
            new DebugDesignerReply(module, DisplayFromProjectId(foundProject), form, [.. rows]),
            DebugJsonContext.Default.DebugDesignerReply);
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

        var progId = ProgIdFor(type);
        if (progId is null)
        {
            return HostError(
                $"'{type}' is not a control kind; pass a ProgID or one of label, textBox, comboBox, " +
                "listBox, checkBox, optionButton, toggleButton, frame, commandButton, tabStrip, " +
                "multiPage, scrollBar, spinButton, image");
        }

        using var owner = parent is { Length: > 0 }
            ? FindContainerControls(designer, parent, 0)
            : designer.GetObject("Controls");
        if (owner is null)
        {
            return HostError(parent is { Length: > 0 }
                ? $"no container named {parent} on {module}; a MultiPage takes controls on its Pages, by page name"
                : $"{module}'s designer has no controls collection");
        }

        DispatchObject? added;
        try
        {
            added = name is { Length: > 0 } ? owner.CallObject("Add", progId, name) : owner.CallObject("Add", progId);
        }
        catch (Exception ex)
        {
            return HostError($"the designer refused to add a {progId} ({ex.Message.Trim()})");
        }

        if (added is null)
        {
            return HostError($"the designer answered nothing for a {progId}");
        }

        using (added)
        {
            var actualName = TryText(added, "Name") ?? name ?? "";
            try
            {
                if (left is { } l) { added.SetDouble("Left", l); }
                if (top is { } t) { added.SetDouble("Top", t); }
                if (width is { } w) { added.SetDouble("Width", w); }
                if (height is { } h) { added.SetDouble("Height", h); }
            }
            catch (Exception ex)
            {
                try { owner.Invoke("Remove", actualName); }
                catch (Exception undo) { Log.Warn($"designer: could not undo the add ({undo.GetType().Name})"); }

                return HostError($"'{actualName}' would not take its geometry, so nothing was added ({ex.Message.Trim()})");
            }

            var actualType = FriendlyControlType(added.TypeName() ?? type);
            Log.Info($"designer: added {actualType} '{actualName}' to {module}{(parent is { Length: > 0 } ? $" in {parent}" : "")}");
            return JsonSerializer.Serialize(
                new DebugDesignerEditReply(true, "add", actualName, actualType,
                    $"added to {(parent is { Length: > 0 } ? parent : module)}"),
                DebugJsonContext.Default.DebugDesignerEditReply);
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

        using var owner = FindOwnerControls(designer, name, 0);
        if (owner is null)
        {
            return HostError($"no control named {name} on {module}");
        }

        try
        {
            owner.Invoke("Remove", name);
        }
        catch (Exception ex)
        {
            return HostError($"the designer refused to remove {name} ({ex.Message.Trim()})");
        }

        Log.Info($"designer: removed '{name}' from {module}");
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

        using var found = name is { Length: > 0 } ? FindControlNamed(designer, name, 0) : null;
        if (name is { Length: > 0 } && found is null)
        {
            return HostError($"no control named {name} on {module}");
        }

        var target = found ?? designer;
        var targetLabel = name is { Length: > 0 } ? name : module;

        var head = property;
        string? tail = null;
        var dot = property.IndexOf('.');
        if (dot > 0 && dot < property.Length - 1)
        {
            head = property[..dot];
            tail = property[(dot + 1)..];
        }

        try
        {
            // A form property the designer does not carry - Width, Height - lives on the
            // component's Properties collection instead, which is what the native Properties
            // window writes. Controls never take this path; their surface is the designer's.
            if (tail is null && found is null && target.GetDispId(property) == DispId.Unknown)
            {
                using var properties = component.GetObject("Properties");
                using var row = properties?.CallObject("Item", property);
                if (row is null)
                {
                    return HostError($"{module} has no property named {property}, on the designer or the component");
                }

                WriteDesignerProperty(row, "Value", value, asKind);
                var (_, rowDisplay) = row.ReadProperty("Value");
                return JsonSerializer.Serialize(
                    new DebugDesignerEditReply(true, "set", targetLabel, null, $"{targetLabel}.{property} is {rowDisplay}"),
                    DebugJsonContext.Default.DebugDesignerEditReply);
            }

            if (tail is null)
            {
                WriteDesignerProperty(target, property, value, asKind);
                var (_, display) = target.ReadProperty(property);
                return JsonSerializer.Serialize(
                    new DebugDesignerEditReply(true, "set", targetLabel, null, $"{targetLabel}.{property} is {display}"),
                    DebugJsonContext.Default.DebugDesignerEditReply);
            }

            using var inner = target.GetObject(head);
            if (inner is null)
            {
                return HostError($"{targetLabel}.{head} is not an object, so {property} cannot be reached");
            }

            WriteDesignerProperty(inner, tail, value, asKind);
            var (_, innerDisplay) = inner.ReadProperty(tail);
            return JsonSerializer.Serialize(
                new DebugDesignerEditReply(true, "set", targetLabel, null, $"{targetLabel}.{property} is {innerDisplay}"),
                DebugJsonContext.Default.DebugDesignerEditReply);
        }
        catch (Exception ex)
        {
            return HostError($"{targetLabel}.{property} refused the write ({ex.Message.Trim()})");
        }
    }

    /// <summary>
    /// Writes with the type the value spells, unless the caller names one with `as`. The
    /// heuristic order matters: "True" is a flag before it is text, "12" is a whole number
    /// before it is a double, and anything else is text - a caption of "123" wants
    /// as=text, which is why the override exists.
    /// </summary>
    private static void WriteDesignerProperty(DispatchObject target, string property, string value, string? asKind)
    {
        switch (asKind)
        {
            case "text":
                target.SetString(property, value);
                return;
            case "flag":
                target.SetBool(property, bool.Parse(value));
                return;
            case "number":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var asked))
                {
                    target.SetInt32(property, asked);
                }
                else
                {
                    target.SetDouble(property, double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture));
                }

                return;
        }

        if (bool.TryParse(value, out var flag))
        {
            target.SetBool(property, flag);
        }
        else if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var whole))
        {
            target.SetInt32(property, whole);
        }
        else if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var real))
        {
            target.SetDouble(property, real);
        }
        else
        {
            target.SetString(property, value);
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
    private static DispatchObject? FindContainerControls(DispatchObject container, string wanted, int depth)
    {
        if (depth > 8)
        {
            return null;
        }

        using var controls = container.GetObject("Controls");
        if (controls is null)
        {
            return null;
        }

        foreach (var control in ItemsOf(controls))
        {
            using (control)
            {
                var name = TryText(control, "Name");

                if (control.GetDispId("Pages") != DispId.Unknown)
                {
                    using var pages = control.GetObject("Pages");
                    if (pages is not null)
                    {
                        foreach (var page in ItemsOf(pages))
                        {
                            using (page)
                            {
                                if (string.Equals(TryText(page, "Name"), wanted, StringComparison.OrdinalIgnoreCase))
                                {
                                    return page.GetObject("Controls");
                                }

                                var below = FindContainerControls(page, wanted, depth + 1);
                                if (below is not null)
                                {
                                    return below;
                                }
                            }
                        }
                    }
                }
                else if (control.GetDispId("Controls") != DispId.Unknown)
                {
                    if (string.Equals(name, wanted, StringComparison.OrdinalIgnoreCase))
                    {
                        return control.GetObject("Controls");
                    }

                    var below = FindContainerControls(control, wanted, depth + 1);
                    if (below is not null)
                    {
                        return below;
                    }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// The collection that owns the named control. Containers are searched BEFORE this level's
    /// names, so a flat top-level collection cannot claim a control its Frame owns.
    /// </summary>
    private static DispatchObject? FindOwnerControls(DispatchObject container, string wanted, int depth)
    {
        if (depth > 8)
        {
            return null;
        }

        var controls = container.GetObject("Controls");
        if (controls is null)
        {
            return null;
        }

        var keep = false;
        try
        {
            foreach (var control in ItemsOf(controls))
            {
                using (control)
                {
                    if (control.GetDispId("Pages") != DispId.Unknown)
                    {
                        using var pages = control.GetObject("Pages");
                        if (pages is null)
                        {
                            continue;
                        }

                        foreach (var page in ItemsOf(pages))
                        {
                            using (page)
                            {
                                var below = FindOwnerControls(page, wanted, depth + 1);
                                if (below is not null)
                                {
                                    return below;
                                }
                            }
                        }
                    }
                    else if (control.GetDispId("Controls") != DispId.Unknown)
                    {
                        var below = FindOwnerControls(control, wanted, depth + 1);
                        if (below is not null)
                        {
                            return below;
                        }
                    }
                }
            }

            foreach (var control in ItemsOf(controls))
            {
                using (control)
                {
                    if (string.Equals(TryText(control, "Name"), wanted, StringComparison.OrdinalIgnoreCase))
                    {
                        keep = true;
                        return controls;
                    }
                }
            }
        }
        finally
        {
            if (!keep)
            {
                controls.Dispose();
            }
        }

        return null;
    }

    /// <summary>
    /// The named control itself, wherever it sits - or a Page by its name, since pages take
    /// property writes too. Match order does not matter: form-wide names are unique.
    /// </summary>
    private static DispatchObject? FindControlNamed(DispatchObject container, string wanted, int depth)
    {
        if (depth > 8)
        {
            return null;
        }

        using var controls = container.GetObject("Controls");
        if (controls is null)
        {
            return null;
        }

        foreach (var control in ItemsOf(controls))
        {
            var give = true;
            try
            {
                if (string.Equals(TryText(control, "Name"), wanted, StringComparison.OrdinalIgnoreCase))
                {
                    give = false;
                    return control;
                }

                if (control.GetDispId("Pages") != DispId.Unknown)
                {
                    using var pages = control.GetObject("Pages");
                    if (pages is null)
                    {
                        continue;
                    }

                    foreach (var page in ItemsOf(pages))
                    {
                        var keepPage = false;
                        try
                        {
                            if (string.Equals(TryText(page, "Name"), wanted, StringComparison.OrdinalIgnoreCase))
                            {
                                keepPage = true;
                                return page;
                            }

                            var below = FindControlNamed(page, wanted, depth + 1);
                            if (below is not null)
                            {
                                return below;
                            }
                        }
                        finally
                        {
                            if (!keepPage)
                            {
                                page.Dispose();
                            }
                        }
                    }
                }
                else if (control.GetDispId("Controls") != DispId.Unknown)
                {
                    var below = FindControlNamed(control, wanted, depth + 1);
                    if (below is not null)
                    {
                        return below;
                    }
                }
            }
            finally
            {
                if (give)
                {
                    control.Dispose();
                }
            }
        }

        return null;
    }

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

    /// <summary>The standard toolbox by its everyday names, or any ProgID a caller spells whole.</summary>
    private static string? ProgIdFor(string type) => type.Contains('.')
        ? type
        : type.ToLowerInvariant() switch
        {
            "label" => "Forms.Label.1",
            "textbox" => "Forms.TextBox.1",
            "combobox" => "Forms.ComboBox.1",
            "listbox" => "Forms.ListBox.1",
            "checkbox" => "Forms.CheckBox.1",
            "optionbutton" => "Forms.OptionButton.1",
            "togglebutton" => "Forms.ToggleButton.1",
            "frame" => "Forms.Frame.1",
            "commandbutton" => "Forms.CommandButton.1",
            "tabstrip" => "Forms.TabStrip.1",
            "multipage" => "Forms.MultiPage.1",
            "scrollbar" => "Forms.ScrollBar.1",
            "spinbutton" => "Forms.SpinButton.1",
            "image" => "Forms.Image.1",
            _ => null,
        };
}
#endif
