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
    public static string? MarkupOf(
        DispatchObject component, string module, out string? reason, ControlDefaults? defaults = null)
    {
        var spec = SpecOf(component, module, out reason, defaults);
        return spec is null ? null : FormMarkup.Print(spec);
    }

    /// <summary>
    /// The workbook a component lives in, by walking back up through its collection to the
    /// project - or null, which is the honest answer for a project never saved to disk. Two
    /// crossings, and the read behind it is cached by the file's own write time, so the cost of
    /// asking is paid once per save rather than once per projection.
    /// </summary>
    internal static string? WorkbookOf(DispatchObject component)
    {
        try
        {
            using var collection = component.GetObject("Collection");
            using var project = collection?.GetObject("Parent");
            var path = project?.GetString("FileName");
            return string.IsNullOrWhiteSpace(path) ? null : path;
        }
        catch
        {
            // A project with no file yet answers by throwing. That is a form with no saved
            // baseline, which is a state this walk already knows how to be.
            return null;
        }
    }

    /// <summary>
    /// The projection itself, in ONE walk: what the designer tab's two halves both ride. The
    /// markup text is Print of this and the visual renders this, so the document and the
    /// canvas cannot disagree about a form they were read from at different moments.
    /// </summary>
    public static FormSpec? SpecOf(
        DispatchObject component, string module, out string? reason, ControlDefaults? defaults = null)
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

        // A control's FONT is inherited from the form rather than defaulted per kind, so the
        // walk compares against the form's own: a bare Label answers "MS Sans Serif" and every
        // control on a Tahoma form would otherwise print a FontName it never chose.
        // The SAVED baseline narrows what the walk asks each control - see ChangedProperties.
        // Only worth fetching when there is an inventory to compare against, because a walk
        // without one prints identity alone and would pay for an answer it cannot use.
        var saved = defaults is null ? null : Core.Forms.SavedDesign.For(WorkbookOf(component));

        Walk(designer, null, string.Empty, rows, seen, 0,
            new WalkContext(defaults, FormBaselineOf(component, designer), saved, module));
        var controls = rows.Select(row => row.Spec).ToList();

        // The form's own properties, from the component's Properties collection FIRST - the
        // native Properties window's slot, and the one the real surface paints; the designer
        // dispatch holds a copy that reads back happily without ever reaching the form frame
        // (measured 2026-08-14). Printed only when NOT the default, because the dialect's
        // rule is that an unspoken property is one an apply can never erase - so a document
        // without the line leaves a custom colour standing, and printing defaults on every
        // form would bury the real choices.
        var properties = new List<PropertySpec>();
        if ((PropertyInt(component, "BackColor") ?? TryInt(designer, "BackColor")) is { } backColor
            && backColor != DefaultFormBackColor)
        {
            properties.Add(new PropertySpec("BackColor", backColor.ToString(System.Globalization.CultureInfo.InvariantCulture), PropertyValueKind.Colour));
        }

        if ((PropertyInt(component, "ForeColor") ?? TryInt(designer, "ForeColor")) is { } foreColor
            && foreColor != DefaultFormForeColor)
        {
            properties.Add(new PropertySpec("ForeColor", foreColor.ToString(System.Globalization.CultureInfo.InvariantCulture), PropertyValueKind.Colour));
        }

        var spec = new FormSpec(
            module,
            PropertyText(component, "Caption") ?? TryText(designer, "Caption"),
            PropertyNumber(component, "Width") ?? TryNumber(designer, "Width"),
            PropertyNumber(component, "Height") ?? TryNumber(designer, "Height"),
            properties,
            controls);

        KeepDesignerDown(component);
        lastWalkRows = rows;
        lastWalkFormBack = PropertyInt(component, "BackColor") ?? TryInt(designer, "BackColor");
        lastWalkFormFore = PropertyInt(component, "ForeColor") ?? TryInt(designer, "ForeColor");
        lastWalkFormInsideWidth = TryNumber(designer, "InsideWidth");
        lastWalkFormInsideHeight = TryNumber(designer, "InsideHeight");
        lastWalkFormPicture = PictureFaceOf(designer);
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
    [ThreadStatic] internal static PictureFace? lastWalkFormPicture;

    /// <summary>An OLE colour as CSS: system indexes through the live system palette, the
    /// rest as the BGR they are. The canvas paints what the machine would.</summary>
    internal static string OleColorToCss(int ole)
    {
        var colorRef = ColorRefOf(ole);
        return $"#{colorRef & 0xFF:x2}{(colorRef >> 8) & 0xFF:x2}{(colorRef >> 16) & 0xFF:x2}";
    }

    /// <summary>An OLE colour as plain 0x00bbggrr, system indexes resolved through the live
    /// palette. The one place that asks "what colour IS this, on this machine, now".</summary>
    internal static int ColorRefOf(int ole) =>
        (ole & unchecked((int)0x80000000)) != 0 ? GetSysColor(ole & 0xFF) : ole;

    /// <summary>The same colour as 0x00rrggbb, which is the order everything outside GDI wants -
    /// GDI+'s background argument, and the RGB half of an ARGB.</summary>
    internal static int ColorRefToRgb(int ole)
    {
        var colorRef = ColorRefOf(ole);
        return ((colorRef & 0xFF) << 16) | (colorRef & 0xFF00) | ((colorRef >> 16) & 0xFF);
    }

    [LibraryImport("user32.dll")]
    private static partial int GetSysColor(int index);

    /// <summary>COLOR_BTNFACE and COLOR_BTNTEXT as OLE colours: what a fresh form carries.</summary>
    private const int DefaultFormBackColor = unchecked((int)0x8000000F);
    private const int DefaultFormForeColor = unchecked((int)0x80000012);

    /// <summary>
    /// A CHEAP KEY for "has this form's design changed", for the one kind of edit the liveness
    /// funnel cannot see: one made outside this product, in the native designer underneath.
    ///
    /// Every mutation xlide makes re-projects the open tab, so the document and the canvas follow
    /// an api set, a panel edit and a canvas gesture. A native edit goes round all of that, and
    /// nothing in the object model announces one - a designer raises no event a code pane's
    /// revision counter is the equivalent of.
    ///
    /// So it is asked rather than announced, and the question has to be cheap enough to ask on
    /// the events that already fire. This reads NAME and the four bounds per control and nothing
    /// else: no fonts, no colours, no pictures, no containers walked twice. It therefore catches
    /// what a hand does in a designer - add, remove, rename, move, resize - and deliberately not
    /// a property nobody can change without one of those (a colour typed into the native
    /// Properties window is not caught, and that is a stated limit rather than an oversight).
    ///
    /// Null when the form cannot be read, which is not the same as "no controls": the caller must
    /// not treat a failed read as a change, or a form mid-teardown re-projects for ever.
    /// </summary>
    public static string? FingerprintOf(DispatchObject designer)
    {
        try
        {
            var key = new System.Text.StringBuilder();
            AppendFingerprint(designer, key, 0);
            return key.ToString();
        }
        catch (Exception why)
        {
            Log.Verbose($"designer: a form would not fingerprint ({why.Message.Trim()})");
            return null;
        }
    }

    private static void AppendFingerprint(DispatchObject container, System.Text.StringBuilder key, int depth)
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
                key.Append(TryText(control, "Name")).Append(' ')
                    .Append(TryNumber(control, "Left")).Append(',')
                    .Append(TryNumber(control, "Top")).Append(' ')
                    .Append(TryNumber(control, "Width")).Append('x')
                    .Append(TryNumber(control, "Height")).Append(';');

                // A container's children are part of its form's shape. Pages are walked through
                // their own collection, which is the only way to reach what sits on page two.
                if (control.GetDispId("Pages") != DispId.Unknown)
                {
                    using var pages = control.GetObject("Pages");
                    foreach (var page in pages is null ? [] : ItemsOf(pages))
                    {
                        using (page)
                        {
                            key.Append(TryText(page, "Name")).Append(';');
                            AppendFingerprint(page, key, depth + 1);
                        }
                    }
                }
                else if (control.GetDispId("Controls") != DispId.Unknown)
                {
                    AppendFingerprint(control, key, depth + 1);
                }
            }
        }
    }

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

        // Names and kinds only: the analyzer wants the members a form's code can reach, and
        // reading every control's whole property set to answer that would be a walk paid for
        // on every seed.
        Walk(designer, null, string.Empty, rows, seen, 0, default);
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
    /// goes back down - and the TOOLBOX goes down WITH it, explicitly, because the editor
    /// raises it whenever a designer window stirs and does not always lower it when the
    /// window merely turns invisible. With the liveness hooks re-projecting on every
    /// mutation, that gap showed as "the native toolbox pane keeps showing" (the developer,
    /// 2026-08-13): each refresh stirred the designer, the Toolbox stood back up, and
    /// nothing put it away.
    /// </summary>
    internal static void KeepDesignerDown(DispatchObject component)
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

        try
        {
            // vbext_wt_Toolbox = 10 in the editor's window-type enumeration.
            using var vbe = component.GetObject("VBE");
            using var windows = vbe?.GetObject("Windows");
            if (windows is null)
            {
                return;
            }

            foreach (var toolWindow in ItemsOf(windows))
            {
                using (toolWindow)
                {
                    if (toolWindow.GetInt32("Type") == 10 && toolWindow.GetBool("Visible"))
                    {
                        toolWindow.SetBool("Visible", false);
                    }
                }
            }
        }
        catch
        {
            // A toolbox that will not answer is not one standing over the surface.
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
        int? ForeColor,
        IReadOnlyList<string>? Tabs,
        /// <summary>Where this control sits in its container's TAB ORDER. Display truth rather
        /// than dialect: the markup does not print it, and the tab-order dialog reads it from
        /// here rather than walking the form a second time.</summary>
        int? TabIndex,
        /// <summary>The control's picture as a data URI, and how it sits. Display truth of the
        /// purest kind: the dialect cannot speak a picture at all - it is binary in the form's
        /// .frx and MSForms does not remember where it came from - so it rides here or the
        /// canvas draws a box where the developer put an image.</summary>
        PictureFace? Picture = null);

    /// <summary>
    /// A picture and its placement, as the canvas needs them. The two families of placement are
    /// both here because a control has one or the other, never both: a SURFACE picture (a form,
    /// an Image, a Frame, a Page) has a size mode, an alignment and a tiling flag, and a CAPTION
    /// picture (a button, a Label, a check box) has a position relative to its caption.
    /// </summary>
    internal sealed record PictureFace(
        string DataUri, int? SizeMode, int? Alignment, bool? Tiling, int? Position);

    /// <summary>
    /// What the walk needs beyond the controls themselves: the measured inventory it compares
    /// against to find CHANGED properties, and the form's own font, which is the baseline for a
    /// control's - a control inherits the form's font, so a bare instance's is the wrong
    /// question. A default context (no inventory) walks identity and geometry alone, which is
    /// what the analyzer's member seed wants.
    /// </summary>
    private readonly record struct WalkContext(
        ControlDefaults? Defaults,
        FormBaseline Form,
        Core.Forms.SavedDesign? Saved = null,
        string Module = "");

    /// <summary>
    /// What a control INHERITS from the form, and therefore what its own values are compared
    /// against. Measured, not assumed: a Toggle freshly added to a form reads BackColor
    /// -2147483633 where a bare instance of the same coclass reads -2147483643, and its font is
    /// the form's Tahoma rather than the bare "MS Sans Serif" (2026-08-16). Comparing an
    /// inherited value against a bare one prints a choice nobody made, on every control.
    /// </summary>
    private readonly record struct FormBaseline(
        string? FontName, double? FontSize, bool? FontBold, bool? FontItalic, int? BackColor, int? ForeColor);

    private static FormBaseline FormBaselineOf(DispatchObject component, DispatchObject designer)
    {
        string? name = null;
        double? size = null;
        bool? bold = null;
        bool? italic = null;
        try
        {
            using var font = designer.GetDispId("Font") != DispId.Unknown ? designer.GetObject("Font") : null;
            if (font is not null)
            {
                name = TryText(font, "Name");
                size = TryNumber(font, "Size");
                bold = TryFlag(font, "Bold");
                italic = TryFlag(font, "Italic");
            }
        }
        catch
        {
            // A form whose font will not answer compares its controls against their own kinds.
        }

        return new FormBaseline(
            name, size, bold, italic,
            PropertyInt(component, "BackColor") ?? TryInt(designer, "BackColor"),
            PropertyInt(component, "ForeColor") ?? TryInt(designer, "ForeColor"));
    }

    /// <summary>
    /// `path` is the dotted chain of container names above this one - empty at the form,
    /// `Options` inside a Frame, `Wizard.Page1` on a page. It exists for the saved baseline,
    /// which is keyed that way because a workbook's storage nests the same way the form does;
    /// `parentName` stays the immediate parent, which is what the dialect prints.
    /// </summary>
    private static void Walk(
        DispatchObject container, string? parentName, string path, List<DesignRow> rows,
        HashSet<string> seen, int depth, WalkContext context)
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

                var where = path.Length == 0 ? name : $"{path}.{name}";

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
                        parentName, ChangedProperties(control, type, where, context)), context));
                }

                if (control.GetDispId("Pages") != DispId.Unknown)
                {
                    WalkPages(control, name, where, rows, seen, depth + 1, context);
                }
                else if (control.GetDispId("Tabs") != DispId.Unknown)
                {
                    // A TabStrip: its tabs are rows, and it holds no controls of its own - what
                    // sits over its face belongs to the form.
                    WalkTabs(control, name, rows, seen, context);
                }
                else if (control.GetDispId("Controls") != DispId.Unknown)
                {
                    Walk(control, name, where, rows, seen, depth + 1, context);
                }
            }
        }
    }

    /*
     * WHAT THIS CONTROL HAS THAT AN UNTOUCHED ONE DOES NOT.
     *
     * The document had been carrying identity, containment, geometry and caption and nothing
     * else, so a colour set through the Properties panel lived only in the object model: absent
     * from the text, absent from the draft preview, outside the document's undo, and lost when a
     * Frame was copied between forms. The document calls itself the transaction log; this is the
     * first instalment of making that true of more than shape.
     *
     * COLOURS ONLY, and the measurement is why. The obvious baseline - a bare instance of the
     * same coclass, which ControlDefaults already measures - is the wrong one for a control that
     * has been SITED. Probed 2026-08-16 by adding a control to a real form and reading it back:
     *
     *     Frame.SpecialEffect      bare 0                 sited 3
     *     ToggleButton.BackColor   bare -2147483643       sited -2147483633
     *     any control's font       bare MS Sans Serif 8.25    sited Tahoma 8 (a Frame: 8.34)
     *
     * MSForms initialises a control differently when it joins a form, so a bare comparison prints
     * choices nobody made: on the fixture it put a font line under every control and a
     * SpecialEffect under the Frame. What a control INHERITS is comparable, because the form is
     * standing right there to be read - so BackColor and ForeColor are compared against the
     * form's own and print exactly when a developer has changed them.
     *
     * AND THE REST ARRIVES BY A THIRD ROAD, 2026-08-17, which neither of the two honest options
     * above is: the workbook Excel has ALREADY saved. MSForms persists only properties that
     * differ from the file format default and names them in a PropMask, so
     * Core.Forms.SavedDesign reads that list straight off the file - no export, nothing written,
     * and no probe control put on anybody's form.
     *
     * IT DOES NOT ANSWER THE QUESTION, IT NARROWS IT. A set bit means "differs from the FILE
     * FORMAT default", which is not "the developer chose this": every control on a Tahoma form
     * carries FontName because the file's default is MS Sans Serif, and every CheckBox,
     * OptionButton and ToggleButton carries BackColor whatever was done to it. So the mask is
     * used as a SHORT LIST of properties worth asking about, and each one still has to beat the
     * bare-coclass comparison to be printed. That is what keeps the file-format noise out, and it
     * is why reading fifty properties of every control on every projection is not necessary.
     */
    private static List<PropertySpec> ChangedProperties(
        DispatchObject control, string kind, string path, WalkContext context)
    {
        // A walk with no inventory is one that wants identity alone - the analyzer's member seed,
        // and the apply's own read of what is currently on the form.
        // The kind's own vocabulary, which is cached - `For` would rebuild a dictionary of forty
        // entries per control per walk to answer two questions about colour.
        if (context.Defaults is not { } defaults)
        {
            return [];
        }

        var vocabulary = defaults.Describe(kind);
        var lines = ColourDifferences(control, vocabulary, context.Form);

        if (context.Saved is { } saved && saved.Knows(context.Module))
        {
            foreach (var name in saved.ChangedOn(context.Module, path))
            {
                if (SavedNarrowing(control, vocabulary, name) is { } line)
                {
                    lines.Add(line);
                }
            }
        }

        return lines;
    }

    /// <summary>
    /// Properties the saved mask names that this walk deliberately does not take from it.
    /// The two colours have a BETTER comparison already - against what the form passes down as
    /// well as against the bare kind - and the fonts are inherited the same way, so both would be
    /// printed on every control of a form that is not MS Sans Serif. A picture is binary and
    /// rides its own face. The packed fields name many properties with one bit and cannot say
    /// which, so their members keep the comparison they have.
    /// </summary>
    private static readonly HashSet<string> NotFromTheSavedMask = new(StringComparer.OrdinalIgnoreCase)
    {
        "BackColor", "ForeColor", "Picture", "MouseIcon",
        "VariousPropertyBits", "BooleanProperties",
        "FontName", "FontHeight", "FontWeight", "FontEffects", "FontCharSet",
        "FontPitchAndFamily", "ParagraphAlign",
    };

    /// <summary>
    /// One property the saved file says is not the format's default, printed only if it also
    /// differs from what this KIND is born with. Answers null for everything else - a name this
    /// walk takes another way, a property the vocabulary does not offer or cannot write, one
    /// with no measured default to compare against, and one whose live value equals it.
    /// </summary>
    private static PropertySpec? SavedNarrowing(
        DispatchObject control, ControlDefaults.Vocabulary kind, string name)
    {
        if (NotFromTheSavedMask.Contains(name))
        {
            return null;
        }

        // The vocabulary has already dropped identity, geometry, caption, the runtime-only
        // members and anything object-valued, so a name it does not carry is a name the document
        // has no business printing.
        var known = kind.Properties.FirstOrDefault(
            one => string.Equals(one.Name, name, StringComparison.OrdinalIgnoreCase));
        if (known is not { Settable: true, Default: { } born })
        {
            return null;
        }

        var (variant, display) = ReadQuietly(control, name);
        if (display is null || !ControlDefaults.Printable(variant) || display == born)
        {
            return null;
        }

        return new PropertySpec(name, display, KindOfValue(known, variant));
    }

    /// <summary>A property read that answers rather than throwing, because a control that will
    /// not discuss one of its own members is an answer this walk already knows how to take.</summary>
    private static (System.Runtime.InteropServices.VarEnum Variant, string? Display) ReadQuietly(
        DispatchObject control, string name)
    {
        try
        {
            return control.GetDispId(name) == DispId.Unknown
                ? (System.Runtime.InteropServices.VarEnum.VT_EMPTY, null)
                : control.ReadProperty(name);
        }
        catch
        {
            return (System.Runtime.InteropServices.VarEnum.VT_EMPTY, null);
        }
    }

    /// <summary>How the dialect spells this value: a colour in hex, a flag as True/False, a
    /// number bare, everything else quoted.</summary>
    private static PropertyValueKind KindOfValue(
        ControlDefaults.Known known, System.Runtime.InteropServices.VarEnum variant) =>
        known.Colour ? PropertyValueKind.Colour
        : variant == System.Runtime.InteropServices.VarEnum.VT_BOOL ? PropertyValueKind.Flag
        : variant == System.Runtime.InteropServices.VarEnum.VT_BSTR ? PropertyValueKind.Text
        : PropertyValueKind.Number;

    /// <summary>
    /// The two colours, against BOTH baselines - and it takes both, measured 2026-08-16. A Label
    /// or a Toggle inherits the FORM's button face; a TextBox, a ComboBox and a ListBox take the
    /// WINDOW colours from their own kind and keep them on any form. Comparing against the form
    /// alone printed `BackColor = &amp;H80000005&amp;` under every entry control on the fixture, and
    /// comparing against the kind alone printed one under every button.
    ///
    /// So a colour is the developer's only when it matches neither: not what this form passes
    /// down, and not what this kind is born with. A value that equals either is unspoken, which
    /// is the dialect's own rule for a property nobody changed.
    /// </summary>
    private static List<PropertySpec> ColourDifferences(
        DispatchObject control, ControlDefaults.Vocabulary kind, FormBaseline form)
    {
        var lines = new List<PropertySpec>();

        void Colour(string name, int? inherited)
        {
            if (TryInt(control, name) is not { } value || value == inherited)
            {
                return;
            }

            var bare = kind.Properties
                .FirstOrDefault(one => string.Equals(one.Name, name, StringComparison.OrdinalIgnoreCase))?
                .Default;
            if (bare is not null
                && int.TryParse(bare, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var born)
                && born == value)
            {
                return;
            }

            lines.Add(new PropertySpec(name, Text(ColorRefOf(value)), PropertyValueKind.Colour));
        }

        Colour("BackColor", form.BackColor);
        Colour("ForeColor", form.ForeColor);
        return lines;
    }

    private static string Text(int value) =>
        value.ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>The display truths, read tolerantly like everything else: a control without a
    /// Font is a row with null fonts, never a failure.</summary>
    private static DesignRow RowOf(DispatchObject control, ControlSpec spec, WalkContext context)
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

        // A TabStrip's tabs are NOT controls - unlike a MultiPage's pages - so the canvas
        // cannot learn them from child rows. Their captions ride the row instead.
        List<string>? tabs = null;
        if (spec.Type == "TabStrip")
        {
            try
            {
                using var strip = control.GetDispId("Tabs") != DispId.Unknown ? control.GetObject("Tabs") : null;
                var count = strip?.GetInt32("Count") ?? 0;
                for (var index = 0; index < count; index++)
                {
                    using var tab = strip!.GetItem(index);
                    var caption = tab is null ? null : TryText(tab, "Caption");
                    if (caption is not null)
                    {
                        (tabs ??= []).Add(caption);
                    }
                }
            }
            catch
            {
                // A strip whose tabs will not answer renders as the bare box it was.
            }
        }

        return new DesignRow(
            spec,
            TryNumber(control, "InsideWidth"),
            TryNumber(control, "InsideHeight"),
            fontName, fontSize, fontBold, fontItalic,
            TryInt(control, "BackColor"),
            TryInt(control, "ForeColor"),
            tabs,
            TryInt(control, "TabIndex"),
            CanCarryPicture(spec.Type) ? PictureFaceOf(control) : null);
    }

    /// <summary>
    /// Which KINDS can hold a picture, asked before the control is - because asking the control
    /// costs two crossings and most controls on most forms cannot. A TextBox, a ComboBox, a
    /// ListBox, a ScrollBar, a SpinButton, a MultiPage and a TabStrip have no Picture at all.
    /// A type outside the toolbox is asked, because a third-party control might.
    /// </summary>
    /// <summary>The properties whose value is a picture, and which are therefore written from a
    /// FILE. Both panels, both write paths and the api all ask here, so there is one answer.</summary>
    internal static bool IsPictureSlot(string property) =>
        string.Equals(property, "Picture", StringComparison.OrdinalIgnoreCase)
        || string.Equals(property, "MouseIcon", StringComparison.OrdinalIgnoreCase);

    private static bool CanCarryPicture(string type) => type is not (
        "TextBox" or "ComboBox" or "ListBox" or "ScrollBar" or "SpinButton"
        or "MultiPage" or "TabStrip" or "Tab" or "RefEdit");

    /// <summary>
    /// The picture a control is wearing, or null - which is the answer for most controls, and
    /// the cheap one: an absent or empty Picture costs the read that found it out and nothing
    /// more. The placement properties are only asked once there is something to place.
    /// </summary>
    private static PictureFace? PictureFaceOf(DispatchObject control)
    {
        try
        {
            using var picture = PictureBytes.PictureOn(control, "Picture");
            if (picture is null || PictureBytes.DataUriOf(picture) is not { } uri)
            {
                return null;
            }

            return new PictureFace(
                uri,
                TryInt(control, "PictureSizeMode"),
                TryInt(control, "PictureAlignment"),
                TryFlag(control, "PictureTiling"),
                TryInt(control, "PicturePosition"));
        }
        catch
        {
            // A control whose picture will not answer draws the bounds it drew before.
            return null;
        }
    }

    /// <summary>
    /// A TabStrip's TABS as rows of their own, so the document can carry them.
    ///
    /// They are not controls - a tab holds nothing and has no geometry - but they are the only
    /// thing about a TabStrip a developer edits, and until 2026-08-16 the markup said nothing
    /// about them at all: the canvas drew two tabs and the document showed a bare line with
    /// nothing indented under it (the owner: "in the markdown, i dont see anything indented under
    /// the tab view"). A Tab line is shaped like a Page's, for the same reason - identity and a
    /// caption, nothing else to say.
    /// </summary>
    private static void WalkTabs(
        DispatchObject strip, string stripName, List<DesignRow> rows, HashSet<string> seen,
        WalkContext context)
    {
        using var tabs = strip.GetDispId("Tabs") != DispId.Unknown ? strip.GetObject("Tabs") : null;
        if (tabs is null)
        {
            return;
        }

        var count = 0;
        try
        {
            count = tabs.GetInt32("Count");
        }
        catch
        {
            // A strip whose tabs will not answer keeps the bare line it had.
            return;
        }

        for (var index = 0; index < count; index++)
        {
            using var tab = tabs.GetItem(index);
            var name = tab is null ? null : TryText(tab, "Name");
            if (tab is null || name is null || !seen.Add(name))
            {
                continue;
            }

            rows.Add(new DesignRow(
                new ControlSpec("Tab", name, TryText(tab, "Caption"), null, null, null, null, stripName, []),
                null, null, null, null, null, null, null, null, null, null, null));
        }
    }

    private static void WalkPages(
        DispatchObject multiPage, string multiPageName, string path, List<DesignRow> rows,
        HashSet<string> seen, int depth, WalkContext context)
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

                var where = $"{path}.{name}";

                if (seen.Add(name))
                {
                    rows.Add(RowOf(page, new ControlSpec(
                        "Page", name, TryText(page, "Caption"),
                        null, null, null, null, multiPageName,
                        ChangedProperties(page, "Page", where, context)), context));
                }

                Walk(page, name, where, rows, seen, depth + 1, context);
            }
        }
    }

    internal static bool? TryFlag(DispatchObject target, string name)
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

    internal static IEnumerable<DispatchObject> ItemsOf(DispatchObject collection)
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

    internal static string? TryText(DispatchObject target, string name)
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

    internal static double? TryNumber(DispatchObject target, string name)
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

    internal static int? TryInt(DispatchObject target, string name)
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

    internal static double? PropertyNumber(DispatchObject component, string name)
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

    internal static string? PropertyText(DispatchObject component, string name)
    {
        try
        {
            using var properties = component.GetObject("Properties");
            using var row = properties?.CallObject("Item", name);
            return row?.GetString("Value");
        }
        catch
        {
            return null;
        }
    }

    internal static int? PropertyInt(DispatchObject component, string name)
    {
        try
        {
            using var properties = component.GetObject("Properties");
            using var row = properties?.CallObject("Item", name);
            return row?.GetInt32("Value");
        }
        catch
        {
            return null;
        }
    }

    /// <summary>The toolbox name for an extender's internal interface, unknown names untouched.</summary>
    /// <summary>A control's toolbox kind, read off the live dispatch; "Control" when it
    /// will not answer.</summary>
    internal static string FriendlyTypeOf(DispatchObject control)
    {
        try
        {
            return FriendlyType(control.TypeName() ?? "Control");
        }
        catch
        {
            return "Control";
        }
    }

    /// <summary>
    /// The event a double-click writes, per toolbox kind - the native designer's own
    /// defaults: Change for the value-bearing kinds, Click for everything else, the form
    /// included.
    /// </summary>
    internal static string DefaultEventFor(string type) => type switch
    {
        "TextBox" or "ComboBox" or "ListBox" or "ScrollBar" or "SpinButton"
            or "MultiPage" or "TabStrip" => "Change",
        _ => "Click",
    };

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

        // Identity and containment only: the diff matches by NAME and writes the document's own
        // property lines, so reading what the form currently holds for each of them would be a
        // walk whose answer the apply is about to overwrite.
        var walked = new List<DesignRow>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Walk(designer, null, string.Empty, walked, seen, 0, default);
        var current = walked.Select(row => row.Spec).ToList();

        var currentByName = current.ToDictionary(row => row.Name, StringComparer.OrdinalIgnoreCase);

        var added = new List<string>();
        var removed = new List<string>();
        var setCount = 0;
        var notes = new List<string>();

        /*
         * A NAME IS ONE CONTROL, and a document that says otherwise is refused here rather than
         * by the dictionary underneath.
         *
         * The whole diff is keyed by name, so two controls of one name have no meaning: the
         * document cannot say which of them a later line describes, and MSForms would not take
         * the second anyway. The lint already marks it at the second mention ("the name 'X' is
         * already taken on this form"); before this the apply threw its dictionary's own words at
         * the developer instead - `An item with the same key has already been added. Key:
         * OkButton` reached the tab's error strip on Ctrl+S (measured 2026-08-16, the owner
         * asking whether the markup checks for collisions at all).
         *
         * Refused whole rather than applied past: an apply that dropped the duplicate would leave
         * a form that matches neither the document nor what the developer meant.
         */
        var clash = wanted.Controls
            .GroupBy(spec => spec.Name, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (clash is not null)
        {
            return new ApplyOutcome(false, [], [], 0,
                $"the name '{clash.Key}' is used {clash.Count()} times, so nothing was applied"
                + " - a name is one control on a form",
                [], true);
        }

        var wantedByName = wanted.Controls.ToDictionary(spec => spec.Name, StringComparer.OrdinalIgnoreCase);

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

                // A TAB leaves its strip's own Tabs collection - it is not in any Controls
                // collection, so the ordinary removal cannot find it.
                if (string.Equals(row.Type, "Tab", StringComparison.OrdinalIgnoreCase))
                {
                    if (!stays && RemoveTab(designer, row.Parent, row.Name))
                    {
                        removed.Add(row.Name);
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

                // A TAB is its strip's, and unlike a page the diff DOES make one: Tabs.Add takes
                // a name and a caption, which is the whole of what a tab is.
                if (string.Equals(spec.Type, "Tab", StringComparison.OrdinalIgnoreCase))
                {
                    var landed = ApplyTab(designer, spec, survived);
                    if (landed && !survived)
                    {
                        added.Add(spec.Name);
                    }
                    else if (!landed)
                    {
                        notes.Add($"{spec.Name}: no TabStrip named {spec.Parent}, so the tab was not added");
                    }
                    else
                    {
                        setCount++;
                    }

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

    internal static string FormatNumber(double value) =>
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

        // A PICTURE never goes through the property bag, whatever it belongs to. The bag hands
        // out a VBE `Property` wrapper whose own `Value` takes a variant, and a picture assigned
        // through one is a value copy of an interface pointer rather than the reference put the
        // control wants. The designer answers the picture the canvas paints, so it is also the
        // slot the write belongs in.
        if (tail is null && IsPictureSlot(property))
        {
            WriteDesignerProperty(target, property, value, asKind);
            using var written = PictureBytes.PictureOn(target, property);
            return PictureBytes.Describe(written);
        }

        if (tail is null && found is null)
        {
            // THE BAG FIRST for a form-level property, the designer dispatch only when the
            // bag lacks the name. The native Properties window writes the component's
            // Properties collection, and that is the slot the real surface paints: a
            // Caption written on the designer dispatch reads back happily and never
            // reaches the form frame - the running form and the design surface both said
            // "UserForm1" over a designer whose Caption read "Quarter Entry" (measured
            // 2026-08-14; the owner's run-beside-canvas is what made it visible).
            using var properties = component.GetObject("Properties");
            using var row = properties?.CallObject("Item", property);
            if (row is not null)
            {
                WriteDesignerProperty(row, "Value", value, asKind);
                var (_, rowDisplay) = row.ReadProperty("Value");
                return rowDisplay;
            }

            if (target.GetDispId(property) == DispId.Unknown)
            {
                throw new InvalidOperationException(
                    $"no property named {property}, on the component or the designer");
            }
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
        // A PICTURE takes a file, not a value, and it is the one property in the panel whose
        // written text is not what it ends up holding: the developer chooses `logo.png` and the
        // control holds a picture object. An empty string takes it off, which is the panel's
        // Clear and the api's way of saying Nothing.
        if (IsPictureSlot(property))
        {
            if (value.Trim().Length == 0)
            {
                target.ClearObject(property);
                return;
            }

            // The control's own BackColor goes with it, for what a transparent PNG's holes
            // become: an OLE picture is a bitmap and a bitmap has no alpha, so they become a
            // colour, and the colour that makes them disappear is the one behind them.
            using var picture = PictureBytes.FromPath(value.Trim(), TryInt(target, "BackColor"));
            target.SetObject(property, picture);
            return;
        }

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
    /// The SENTENCE AFTER THE REFUSAL, for the two failures a developer can do something about.
    ///
    /// Measured 2026-08-16, adding a dozen real third-party ProgIDs to the fixture form. Every
    /// one of them came back with one of two messages and neither says what to do:
    ///
    ///   - `Invalid class string` - the control is not registered on this machine at all. The
    ///     ProgID is a spelling, and a spelling nothing answers to is a typo or a missing OCX.
    ///   - `The subject is not trusted for the specified action` - TRUST_E_SUBJECT_NOT_TRUSTED,
    ///     and the control IS registered: `MSComctlLib.TreeCtrl.2`, `Shell.Explorer.2` and
    ///     `RefEdit.Ctrl` all resolve on this machine and all are refused. That is OFFICE
    ///     refusing, not MSForms and not this product - the Trust Center's ActiveX setting turns
    ///     every non-MSForms control off, and nothing in a form can create one while it is on.
    ///
    /// Naming the second one matters more than it looks. Without it the refusal reads as a defect
    /// in xlide, and the developer's next hour goes into the wrong place entirely.
    /// </summary>
    private static string WhyAddFailed(Exception ex, string? name)
    {
        const int SubjectNotTrusted = unchecked((int)0x800B0004);

        if (ex.HResult == SubjectNotTrusted
            || ex.Message.Contains("not trusted", StringComparison.OrdinalIgnoreCase))
        {
            return " - Office is refusing to create ActiveX controls at all, which is the Trust"
                + " Center's ActiveX Settings rather than anything about this control";
        }

        if (ex.Message.Contains("Invalid class string", StringComparison.OrdinalIgnoreCase))
        {
            return " - no control of that ProgID is registered on this machine";
        }

        // MSForms answers `error 800a9c6c` and nothing else for a NAME it will not take, which
        // teaches a developer nothing at all (measured 2026-08-16, adding a control called
        // `_Leading`). A control's name is a VBA identifier, and the rule is worth stating.
        // The same rule the markup's lint squiggles with, from the same place: one answer to
        // "is this a name MSForms will take", whichever end of the product asks.
        return name is { Length: > 0 } && !FormMarkup.IsIdentifier(name)
            ? $" - '{name}' is not a name MSForms will take: a control's name is a VBA identifier,"
                + " so it starts with a letter and holds only letters, digits and underscores"
            : string.Empty;
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
            throw new InvalidOperationException(
                $"the designer refused to add a {progId} ({ex.Message.Trim()}){WhyAddFailed(ex, name)}");
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

    /// <summary>A TabStrip's Tabs collection, by the strip's name. Null when the name is not a
    /// strip, which is how a Tab line under the wrong parent becomes a note rather than a
    /// failure.</summary>
    private static DispatchObject? TabsOf(DispatchObject designer, string? stripName)
    {
        if (stripName is not { Length: > 0 })
        {
            return null;
        }

        using var strip = FindControlNamed(designer, stripName, 0);
        return strip is not null && strip.GetDispId("Tabs") != DispId.Unknown
            ? strip.GetObject("Tabs")
            : null;
    }

    /// <summary>
    /// A tab the document names: made if the strip has no such tab, and its caption written
    /// either way. Answers false only when there is no strip of that name to work on.
    /// </summary>
    private static bool ApplyTab(DispatchObject designer, ControlSpec spec, bool survived)
    {
        using var tabs = TabsOf(designer, spec.Parent);
        if (tabs is null)
        {
            return false;
        }

        if (!survived)
        {
            // The NAME alone: MSForms takes a caption too, but the caption line below sets it
            // anyway and a tab created with only its name is captioned with it, which is the
            // same default the native New Page gives.
            tabs.Invoke("Add", spec.Name);
        }

        if (spec.Caption is { } caption)
        {
            using var tab = TabNamed(tabs, spec.Name);
            tab?.SetString("Caption", caption);
        }

        return true;
    }

    /// <summary>Takes a tab out of its strip, by name.</summary>
    private static bool RemoveTab(DispatchObject designer, string? stripName, string name)
    {
        using var tabs = TabsOf(designer, stripName);
        if (tabs is null)
        {
            return false;
        }

        try
        {
            tabs.Invoke("Remove", name);
            return true;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"the strip refused to remove {name} ({ex.Message.Trim()})");
        }
    }

    /// <summary>One tab of a strip, by name - the collection indexes by name as well as by
    /// number, but a walk is the honest way to answer "is there one".</summary>
    private static DispatchObject? TabNamed(DispatchObject tabs, string name)
    {
        var count = 0;
        try
        {
            count = tabs.GetInt32("Count");
        }
        catch
        {
            return null;
        }

        for (var index = 0; index < count; index++)
        {
            var tab = tabs.GetItem(index);
            if (tab is not null && string.Equals(TryText(tab, "Name"), name, StringComparison.OrdinalIgnoreCase))
            {
                return tab;
            }

            tab?.Dispose();
        }

        return null;
    }

    /// <summary>
    /// Brings a control to the front of its container or sends it to the back - MSForms' own
    /// ZOrder, which is a METHOD rather than a property and so is reachable through none of the
    /// property paths this service otherwise uses.
    /// </summary>
    internal static bool ZOrderControl(DispatchObject designer, string name, bool toFront)
    {
        using var control = FindControlNamed(designer, name, 0);
        if (control is null)
        {
            return false;
        }

        try
        {
            // fmZOrderFront is 0 and fmZOrderBack is 1, which is the one place in this file where
            // a bare number is the API rather than a value we chose.
            control.Invoke("ZOrder", toFront ? 0 : 1);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"the designer refused to reorder {name} ({ex.Message.Trim()})");
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
    internal static DispatchObject? FindControlNamed(DispatchObject container, string wanted, int depth)
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

    /// <summary>
    /// The standard toolbox: the everyday name a document spells, and the coclass an apply creates
    /// for it. ONE list - `ProgIdFor` answers from it, the defaults inventory measures each entry,
    /// and the language service offers exactly these kinds - so a kind a document can be given is
    /// a kind a completion can suggest, by construction rather than by two lists agreeing.
    /// </summary>
    internal static readonly IReadOnlyDictionary<string, string> Toolbox =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["Label"] = "Forms.Label.1",
            ["TextBox"] = "Forms.TextBox.1",
            ["ComboBox"] = "Forms.ComboBox.1",
            ["ListBox"] = "Forms.ListBox.1",
            ["CheckBox"] = "Forms.CheckBox.1",
            ["OptionButton"] = "Forms.OptionButton.1",
            ["ToggleButton"] = "Forms.ToggleButton.1",
            ["Frame"] = "Forms.Frame.1",
            ["CommandButton"] = "Forms.CommandButton.1",
            ["TabStrip"] = "Forms.TabStrip.1",
            ["MultiPage"] = "Forms.MultiPage.1",
            ["ScrollBar"] = "Forms.ScrollBar.1",
            ["SpinButton"] = "Forms.SpinButton.1",
            ["Image"] = "Forms.Image.1",
        };

    /// <summary>The standard toolbox by its everyday names, or any ProgID a caller spells whole.</summary>
    internal static string? ProgIdFor(string type) => type.Contains('.')
        ? type
        : Toolbox.TryGetValue(type, out var progId) ? progId : null;
}
