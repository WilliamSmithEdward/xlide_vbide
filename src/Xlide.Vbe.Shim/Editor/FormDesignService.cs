using Xlide.Vbe.Core.Forms;
using System.Runtime.InteropServices;
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
internal static partial class FormDesignService
{
    /// <summary>The form's markup, or null with a reason when the component has no designer.</summary>
    public static string? MarkupOf(DispatchObject component, string module, out string? reason)
    {
        var spec = SpecOf(component, module, out reason);
        return spec is null ? null : FormMarkup.Print(spec);
    }

    /// <summary>
    /// The projection itself, in ONE walk: what the designer tab's two halves both ride. The
    /// markup text is Print of this and the visual renders this, so the document and the
    /// canvas cannot disagree about a form they were read from at different moments.
    /// </summary>
    public static FormSpec? SpecOf(DispatchObject component, string module, out string? reason)
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

        var rows = new List<DesignRow>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Walk(designer, null, rows, seen, 0);
        var controls = rows.Select(row => row.Spec).ToList();

        // The form's own properties, as the native Properties window holds them - which is
        // what links the markup's lines to that panel: same source, same value. Printed only
        // when NOT the default, because the dialect's rule is that an unspoken property is
        // one an apply can never erase - so a document without the line leaves a custom
        // colour standing, and printing defaults on every form would bury the real choices.
        var properties = new List<PropertySpec>();
        if (TryInt(designer, "BackColor") is { } backColor && backColor != DefaultFormBackColor)
        {
            properties.Add(new PropertySpec("BackColor", backColor.ToString(System.Globalization.CultureInfo.InvariantCulture), PropertyValueKind.Number));
        }

        if (TryInt(designer, "ForeColor") is { } foreColor && foreColor != DefaultFormForeColor)
        {
            properties.Add(new PropertySpec("ForeColor", foreColor.ToString(System.Globalization.CultureInfo.InvariantCulture), PropertyValueKind.Number));
        }

        var spec = new FormSpec(
            module,
            TryText(designer, "Caption"),
            TryNumber(designer, "Width") ?? PropertyNumber(component, "Width"),
            TryNumber(designer, "Height") ?? PropertyNumber(component, "Height"),
            properties,
            controls);

        KeepDesignerDown(component);
        lastWalkRows = rows;
        lastWalkFormBack = TryInt(designer, "BackColor");
        lastWalkFormFore = TryInt(designer, "ForeColor");
        lastWalkFormInsideWidth = TryNumber(designer, "InsideWidth");
        lastWalkFormInsideHeight = TryNumber(designer, "InsideHeight");
        return spec;
    }

    /*
     * The display half of the LAST SpecOf, for the publisher that calls SpecOf and then
     * needs the rows: one walk serves both, and threading a tuple through every SpecOf
     * caller for the one that wants more would put the message's needs into the markup's
     * signature. Host-thread only, read immediately after the call, like the out reason.
     */
    [ThreadStatic] internal static List<DesignRow>? lastWalkRows;
    [ThreadStatic] internal static int? lastWalkFormBack;
    [ThreadStatic] internal static int? lastWalkFormFore;
    [ThreadStatic] internal static double? lastWalkFormInsideWidth;
    [ThreadStatic] internal static double? lastWalkFormInsideHeight;

    /// <summary>An OLE colour as CSS: system indexes through the live system palette, the
    /// rest as the BGR they are. The canvas paints what the machine would.</summary>
    internal static string OleColorToCss(int ole)
    {
        var colorRef = (ole & unchecked((int)0x80000000)) != 0 ? GetSysColor(ole & 0xFF) : ole;
        return $"#{colorRef & 0xFF:x2}{(colorRef >> 8) & 0xFF:x2}{(colorRef >> 16) & 0xFF:x2}";
    }

    [LibraryImport("user32.dll")]
    private static partial int GetSysColor(int index);

    /// <summary>COLOR_BTNFACE and COLOR_BTNTEXT as OLE colours: what a fresh form carries.</summary>
    private const int DefaultFormBackColor = unchecked((int)0x8000000F);
    private const int DefaultFormForeColor = unchecked((int)0x80000012);

    /// <summary>
    /// The controls as the analyzer's implicit members: name plus the type completion resolves
    /// it as. The standard toolbox spells as MSForms - the model the analyzer promoted for
    /// exactly this (xlide_vscode#17) - and a type outside the table passes through raw, which
    /// upstream takes unchanged rather than guessing at. Null when there is nothing to say.
    /// </summary>
    public static Xlide.Vbe.Core.Engine.EngineImplicitMember[]? ControlMembers(DispatchObject component)
    {
        if (component.GetInt32("Type") != 3)
        {
            return null;
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            return null;
        }

        var rows = new List<DesignRow>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Walk(designer, null, rows, seen, 0);
        if (rows.Count == 0)
        {
            return null;
        }

        KeepDesignerDown(component);
        return [.. rows.Select(row => new Xlide.Vbe.Core.Engine.EngineImplicitMember(
            row.Spec.Name,
            IsToolboxType(row.Spec.Type) ? $"MSForms.{row.Spec.Type}" : row.Spec.Type))];
    }

    /// <summary>
    /// Touching a Designer materialises its window, and a workbook saved while one existed
    /// RESTORES it on open - which is how the native Toolbox appeared floating over the
    /// surface the first time a saved form's project was walked (2026-08-13). The designer
    /// surface is the canvas milestone's; until it lands, a designer window this side wakes
    /// goes back down, and the Toolbox follows it.
    /// </summary>
    private static void KeepDesignerDown(DispatchObject component)
    {
        try
        {
            using var window = component.CallObject("DesignerWindow");
            if (window is not null && window.GetBool("Visible"))
            {
                window.SetBool("Visible", false);
            }
        }
        catch
        {
            // A designer without a window has nothing to put down.
        }
    }

    private static bool IsToolboxType(string type) => type is "Label" or "TextBox" or "ComboBox"
        or "ListBox" or "CheckBox" or "OptionButton" or "ToggleButton" or "Frame"
        or "CommandButton" or "TabStrip" or "MultiPage" or "Page" or "ScrollBar"
        or "SpinButton" or "Image";

    /// <summary>
    /// The walk knows each control's container because it is standing in it, so the narrow
    /// rows carry exact parents without a COM read per control. Recursion plus the dedupe
    /// makes flat and hierarchical collections the same world, as the debug walk's does.
    /// </summary>
    /// <summary>
    /// One control's row with what the CANVAS needs beside what the markup speaks: the spec
    /// (identity, containment, geometry, caption - the dialect), and the display truths the
    /// dialect deliberately does not (font, colours, and a container's REAL client area via
    /// InsideWidth/InsideHeight - the parity numbers, so the canvas derives its insets from
    /// the model instead of guessing them).
    /// </summary>
    internal sealed record DesignRow(
        ControlSpec Spec,
        double? InsideWidth,
        double? InsideHeight,
        string? FontName,
        double? FontSize,
        bool? FontBold,
        bool? FontItalic,
        int? BackColor,
        int? ForeColor);

    private static void Walk(
        DispatchObject container, string? parentName, List<DesignRow> rows, HashSet<string> seen, int depth)
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

                    rows.Add(RowOf(control, new ControlSpec(
                        type, name, TryText(control, "Caption"),
                        TryNumber(control, "Left"), TryNumber(control, "Top"),
                        TryNumber(control, "Width"), TryNumber(control, "Height"),
                        parentName, [])));
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

    /// <summary>The display truths, read tolerantly like everything else: a control without a
    /// Font is a row with null fonts, never a failure.</summary>
    private static DesignRow RowOf(DispatchObject control, ControlSpec spec)
    {
        string? fontName = null;
        double? fontSize = null;
        bool? fontBold = null;
        bool? fontItalic = null;
        try
        {
            using var font = control.GetDispId("Font") != DispId.Unknown ? control.GetObject("Font") : null;
            if (font is not null)
            {
                fontName = TryText(font, "Name");
                fontSize = TryNumber(font, "Size");
                fontBold = TryFlag(font, "Bold");
                fontItalic = TryFlag(font, "Italic");
            }
        }
        catch
        {
            // A control whose font will not answer renders in the form's own.
        }

        return new DesignRow(
            spec,
            TryNumber(control, "InsideWidth"),
            TryNumber(control, "InsideHeight"),
            fontName, fontSize, fontBold, fontItalic,
            TryInt(control, "BackColor"),
            TryInt(control, "ForeColor"));
    }

    private static void WalkPages(
        DispatchObject multiPage, string multiPageName, List<DesignRow> rows, HashSet<string> seen, int depth)
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
                    rows.Add(RowOf(page, new ControlSpec(
                        "Page", name, TryText(page, "Caption"),
                        null, null, null, null, multiPageName, [])));
                }

                Walk(page, name, rows, seen, depth + 1);
            }
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

    // ------------------------------------------------------------------ the apply
    //
    // Moved here from the debug partial 2026-08-13, which is the unification the header above
    // promised: the markup tab's Ctrl+S is product surface and a Release build must carry it,
    // so the machinery the #if DEBUG routes shared could no longer live behind the #if. The
    // debug routes now call THESE, which is also what makes the api's applyMarkup and the
    // tab's apply the same operation by construction rather than by promise.

    /// <summary>How an apply ended: what landed, and why it stopped if it did. A parse
    /// refusal applies NOTHING; a stop partway reports exactly what landed first.</summary>
    internal sealed record ApplyOutcome(
        bool Ok,
        IReadOnlyList<string> Added,
        IReadOnlyList<string> Removed,
        int Set,
        string? Refused,
        IReadOnlyList<string> Notes,
        bool ParseFailed);

    /// <summary>
    /// Applies a markup document to the live form as a NAME-KEYED DIFF: controls only in the
    /// markup are added, controls only in the model are removed, matched controls take their
    /// header geometry, caption and property lines. A changed type or container is a
    /// remove-and-add, which is also the truth of what it means.
    /// </summary>
    public static ApplyOutcome ApplyMarkup(DispatchObject component, string module, string markup)
    {
        FormSpec wanted;
        try
        {
            wanted = FormMarkup.Parse(markup);
        }
        catch (FormMarkupException refused)
        {
            return new ApplyOutcome(false, [], [], 0,
                $"the markup did not parse, so nothing was applied: {refused.Message}", [], true);
        }

        using var designer = component.GetObject("Designer");
        if (designer is null)
        {
            return new ApplyOutcome(false, [], [], 0, $"{module} has no designer", [], false);
        }

        var walked = new List<DesignRow>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Walk(designer, null, walked, seen, 0);
        var current = walked.Select(row => row.Spec).ToList();

        var currentByName = current.ToDictionary(row => row.Name, StringComparer.OrdinalIgnoreCase);
        var wantedByName = wanted.Controls.ToDictionary(spec => spec.Name, StringComparer.OrdinalIgnoreCase);

        var added = new List<string>();
        var removed = new List<string>();
        var setCount = 0;
        var notes = new List<string>();

        try
        {
            // The form's own header and properties first: caption, size, property lines.
            if (wanted.Caption is { } formCaption)
            {
                SetControlProperty(component, designer, null, "Caption", formCaption, "text");
                setCount++;
            }

            if (wanted.Width is { } formWidth)
            {
                SetControlProperty(component, designer, null, "Width", FormatNumber(formWidth), "number");
                setCount++;
            }

            if (wanted.Height is { } formHeight)
            {
                SetControlProperty(component, designer, null, "Height", FormatNumber(formHeight), "number");
                setCount++;
            }

            foreach (var property in wanted.Properties)
            {
                SetControlProperty(component, designer, null, property.Path, property.Value, KindWord(property.Kind));
                setCount++;
            }

            // Removals before additions, so a name moving container or kind frees itself first.
            // Pages are the MultiPage's own; the diff neither creates nor removes one (a page
            // set that disagrees is reported rather than half-applied).
            foreach (var row in current)
            {
                var isPage = string.Equals(row.Type, "Page", StringComparison.OrdinalIgnoreCase);
                var match = wantedByName.TryGetValue(row.Name, out var spec) ? spec : null;
                var stays = match is not null
                    && string.Equals(match.Type, row.Type, StringComparison.OrdinalIgnoreCase)
                    && SameParent(match, row, module);

                if (isPage)
                {
                    if (match is null)
                    {
                        notes.Add($"{row.Name}: pages belong to their MultiPage; the diff does not remove one");
                    }

                    continue;
                }

                if (!stays)
                {
                    if (RemoveControl(designer, row.Name))
                    {
                        removed.Add(row.Name);
                    }
                }
            }

            // Additions and updates, in document order so containers exist before their children.
            foreach (var spec in wanted.Controls)
            {
                var isPage = string.Equals(spec.Type, "Page", StringComparison.OrdinalIgnoreCase);
                var match = currentByName.TryGetValue(spec.Name, out var row) ? row : null;
                var survived = match is not null
                    && string.Equals(match.Type, spec.Type, StringComparison.OrdinalIgnoreCase)
                    && SameParent(spec, match, module)
                    && !removed.Contains(match.Name);

                if (isPage && match is null)
                {
                    notes.Add($"{spec.Name}: pages belong to their MultiPage; the diff does not add one");
                    continue;
                }

                if (!isPage && !survived)
                {
                    var progId = ProgIdFor(spec.Type)
                        ?? spec.Properties.FirstOrDefault(p => string.Equals(p.Path, "ProgId", StringComparison.OrdinalIgnoreCase))?.Value;
                    if (progId is null)
                    {
                        notes.Add($"{spec.Name}: '{spec.Type}' is not a toolbox kind and no ProgId line names one, so it was not added");
                        continue;
                    }

                    using var owner = spec.Parent is { Length: > 0 }
                        ? FindContainerControls(designer, spec.Parent, 0)
                        : designer.GetObject("Controls");
                    if (owner is null)
                    {
                        notes.Add($"{spec.Name}: no container named {spec.Parent}, so it was not added");
                        continue;
                    }

                    AddControl(owner, progId, spec.Name, spec.Left, spec.Top, spec.Width, spec.Height);
                    added.Add(spec.Name);
                }
                else if (survived || isPage)
                {
                    if (spec.Left is { } left) { SetControlProperty(component, designer, spec.Name, "Left", FormatNumber(left), "number"); setCount++; }
                    if (spec.Top is { } top) { SetControlProperty(component, designer, spec.Name, "Top", FormatNumber(top), "number"); setCount++; }
                    if (spec.Width is { } width) { SetControlProperty(component, designer, spec.Name, "Width", FormatNumber(width), "number"); setCount++; }
                    if (spec.Height is { } height) { SetControlProperty(component, designer, spec.Name, "Height", FormatNumber(height), "number"); setCount++; }
                }

                if (spec.Caption is { } caption)
                {
                    SetControlProperty(component, designer, spec.Name, "Caption", caption, "text");
                    setCount++;
                }

                foreach (var property in spec.Properties)
                {
                    if (string.Equals(property.Path, "ProgId", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    SetControlProperty(component, designer, spec.Name, property.Path, property.Value, KindWord(property.Kind));
                    setCount++;
                }
            }
        }
        catch (Exception ex)
        {
            Log.Warn($"designer: apply to {module} stopped partway ({ex.GetType().Name})");
            KeepDesignerDown(component);
            return new ApplyOutcome(false, [.. added], [.. removed], setCount,
                $"stopped at: {ex.Message.Trim()}. What is listed here landed before the refusal.",
                [.. notes], false);
        }

        Log.Info($"designer: markup applied to {module}: +{added.Count} -{removed.Count} set {setCount}");
        KeepDesignerDown(component);
        return new ApplyOutcome(true, [.. added], [.. removed], setCount, null, [.. notes], false);
    }

    /// <summary>Parents compare with the form's own name standing for "no parent", so a
    /// top-level control matches whether its row spells the form or nothing.</summary>
    private static bool SameParent(ControlSpec left, ControlSpec right, string module)
    {
        var leftParent = left.Parent ?? module;
        var rightParent = right.Parent ?? module;
        return string.Equals(leftParent, rightParent, StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatNumber(double value) =>
        value.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static string KindWord(PropertyValueKind kind) => kind switch
    {
        PropertyValueKind.Text => "text",
        PropertyValueKind.Flag => "flag",
        _ => "number",
    };

    /// <summary>
    /// Writes one property of a control, or of the form itself when no control is named, and
    /// answers what the property READS BACK - the Properties panel's convention, because a
    /// write proving only that it was asked for is a lie. One level of dotting reaches an
    /// object-valued property's member (Font.Size); a form property the designer does not
    /// carry - Width, Height - falls through to the component's Properties collection, which
    /// is what the native Properties window writes. Throws rather than answering, because
    /// both its callers - the route and the apply - turn refusals into their own shapes.
    /// </summary>
    internal static string SetControlProperty(
        DispatchObject component, DispatchObject designer,
        string? name, string property, string value, string? asKind)
    {
        using var found = name is { Length: > 0 } ? FindControlNamed(designer, name, 0) : null;
        if (name is { Length: > 0 } && found is null)
        {
            throw new InvalidOperationException($"no control named {name}");
        }

        var target = found ?? designer;

        var head = property;
        string? tail = null;
        var dot = property.IndexOf('.');
        if (dot > 0 && dot < property.Length - 1)
        {
            head = property[..dot];
            tail = property[(dot + 1)..];
        }

        if (tail is null && found is null && target.GetDispId(property) == DispId.Unknown)
        {
            using var properties = component.GetObject("Properties");
            using var row = properties?.CallObject("Item", property);
            if (row is null)
            {
                throw new InvalidOperationException(
                    $"no property named {property}, on the designer or the component");
            }

            WriteDesignerProperty(row, "Value", value, asKind);
            var (_, rowDisplay) = row.ReadProperty("Value");
            return rowDisplay;
        }

        if (tail is null)
        {
            WriteDesignerProperty(target, property, value, asKind);
            var (_, display) = target.ReadProperty(property);
            return display;
        }

        using var inner = target.GetObject(head);
        if (inner is null)
        {
            throw new InvalidOperationException($"{head} is not an object, so {property} cannot be reached");
        }

        WriteDesignerProperty(inner, tail, value, asKind);
        var (_, innerDisplay) = inner.ReadProperty(tail);
        return innerDisplay;
    }

    /// <summary>
    /// Writes with the type the value spells, unless the caller names one with `as`. The
    /// heuristic order matters: "True" is a flag before it is text, "12" is a whole number
    /// before it is a double, and anything else is text - a caption of "123" wants as=text,
    /// which is why the override exists.
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
                if (int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var asked))
                {
                    target.SetInt32(property, asked);
                }
                else
                {
                    target.SetDouble(property, double.Parse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture));
                }

                return;
        }

        if (bool.TryParse(value, out var flag))
        {
            target.SetBool(property, flag);
        }
        else if (int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var whole))
        {
            target.SetInt32(property, whole);
        }
        else if (double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var real))
        {
            target.SetDouble(property, real);
        }
        else
        {
            target.SetString(property, value);
        }
    }

    /// <summary>
    /// The add itself, against a resolved Controls collection: the control asked for, placed -
    /// or nothing, because an add that cannot then take its geometry is taken back out. Throws
    /// rather than answering, for the same reason the property write does.
    /// </summary>
    internal static (string Name, string Type) AddControl(
        DispatchObject owner, string progId, string? name,
        double? left, double? top, double? width, double? height)
    {
        DispatchObject? added;
        try
        {
            added = name is { Length: > 0 } ? owner.CallObject("Add", progId, name) : owner.CallObject("Add", progId);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"the designer refused to add a {progId} ({ex.Message.Trim()})");
        }

        if (added is null)
        {
            throw new InvalidOperationException($"the designer answered nothing for a {progId}");
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

                throw new InvalidOperationException(
                    $"'{actualName}' would not take its geometry, so nothing was added ({ex.Message.Trim()})");
            }

            return (actualName, FriendlyType(added.TypeName() ?? progId));
        }
    }

    /// <summary>
    /// The remove itself: through the collection that owns the control. False when no control
    /// carries the name - which for the markup diff is an answer, not a failure, because a
    /// child leaves with its removed container.
    /// </summary>
    internal static bool RemoveControl(DispatchObject designer, string name)
    {
        using var owner = FindOwnerControls(designer, name, 0);
        if (owner is null)
        {
            return false;
        }

        try
        {
            owner.Invoke("Remove", name);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"the designer refused to remove {name} ({ex.Message.Trim()})");
        }

        return true;
    }

    /// <summary>A container's Controls collection by the container's name: a Frame's, or a
    /// MultiPage PAGE's, searched recursively - a MultiPage itself takes controls on its
    /// pages, never directly.</summary>
    internal static DispatchObject? FindContainerControls(DispatchObject container, string wanted, int depth)
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

    /// <summary>The standard toolbox by its everyday names, or any ProgID a caller spells whole.</summary>
    internal static string? ProgIdFor(string type) => type.Contains('.')
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
