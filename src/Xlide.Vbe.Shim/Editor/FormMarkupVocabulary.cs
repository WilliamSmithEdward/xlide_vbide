using Xlide.Vbe.Shim.Com;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// What the markup language KNOWS, shaped for the page: every kind a document can name, and every
/// property each of them holds. The answer behind the designer tab's completions and hovers
/// (docs/userform-designer.md, the language service, step 3).
///
/// Nothing here is a table of the language's own. The kinds are the apply's toolbox, the
/// properties are measured from a bare instance of each coclass, and their meanings come from the
/// type library MSForms ships on this machine - so a completion offers what an apply can actually
/// write, and a hover repeats what the Object Browser would say rather than a sentence invented
/// here. A property that is not in the vocabulary is one this product cannot honestly describe.
///
/// The FORM is the exception that proves the rule: it has no coclass to instantiate bare, so it is
/// described from a live form when the page names one - the same designer the Properties panel
/// reads, put back down afterwards the way every walk here puts it down.
/// </summary>
internal static class FormMarkupVocabulary
{
    /// <summary>The containers, which the completions filter by: a Page belongs to a MultiPage,
    /// a control belongs to anything that holds one. The parser is the authority on this and
    /// refuses the rest; this is the same rule said early, as a suggestion.</summary>
    private static readonly HashSet<string> Containers =
        new(StringComparer.OrdinalIgnoreCase) { "Frame", "MultiPage", "Page" };

    /// <summary>
    /// What each kind IS, in a line - the sentence a hover leads with (the owner, 2026-08-16:
    /// "hovering the control class should describe the class").
    ///
    /// The WORDING is ours, and it is the second table in this product that has to be, for the
    /// same measured reason as the system colour names: MSForms ships no help strings at all -
    /// every `doc` came back empty across all fifteen kinds - so there is nothing to read this
    /// out of. Each line says what the control does rather than what it looks like, because the
    /// canvas beside the document is already showing what it looks like.
    /// </summary>
    private static readonly Dictionary<string, string> Descriptions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Form"] = "The UserForm itself. Everything indented under it sits on it.",
        ["Label"] = "Static text on the form.",
        ["TextBox"] = "A box the user types into, one line or many.",
        ["ComboBox"] = "A text box with a drop-down list beside it.",
        ["ListBox"] = "A list the user picks from, one row or several.",
        ["CheckBox"] = "A box the user ticks; its Value is True, False or Null.",
        ["OptionButton"] = "One of a group, where picking one clears the rest.",
        ["ToggleButton"] = "A button that stays pressed - a check box shaped like a button.",
        ["Frame"] = "A captioned box that groups controls and owns their coordinates.",
        ["CommandButton"] = "A push button; its Click is the usual reason for a handler.",
        ["TabStrip"] = "Tabs over ONE set of controls: the strip reports which tab is current and the Change event is where the code repoints them. Same widgets, different data.",
        ["MultiPage"] = "Tabs, each with a page of controls of its own - MSForms shows and hides them, so no code is needed. Different widgets per tab.",
        ["Page"] = "One page of a MultiPage, holding its own controls.",
        ["Tab"] = "One tab of a TabStrip. It holds nothing - the controls it appears to hold belong to the form, and the strip's Change event is what repoints them.",
        ["ScrollBar"] = "A bar the user drags between Min and Max.",
        ["SpinButton"] = "Two arrows that step a value, usually beside a TextBox.",
        ["Image"] = "A picture on the form.",
    };

    /// <summary>
    /// The whole vocabulary. <paramref name="form"/> is a live form component when the page has
    /// one open, which is what makes the Form entry possible; without it every other kind still
    /// answers, because those are measured from coclasses and need no workbook at all.
    /// </summary>
    public static FormMarkupKind[] Of(ControlDefaults defaults, PropertyTypes types, DispatchObject? form)
    {
        var kinds = new List<FormMarkupKind>();

        if (FormProperties(types, form) is { Length: > 0 } formProperties)
        {
            kinds.Add(new FormMarkupKind("Form", null, true, formProperties, Descriptions["Form"]));
        }

        foreach (var kind in FormDesignService.Toolbox.Keys)
        {
            kinds.Add(Describe(defaults, kind));
        }

        // A Page and a Tab are spelled and have properties, and NEITHER has a coclass to
        // instantiate bare - the same problem the Form has, with the same answer: describe one
        // that is standing on the live form. Without a form open they keep the empty vocabulary
        // the bare walk gives them, which is honest rather than invented.
        kinds.Add(DescribedLive(defaults, types, form, "Page"));
        kinds.Add(DescribedLive(defaults, types, form, "Tab"));

        return [.. kinds];
    }

    /// <summary>
    /// A kind with no coclass, described from one that is LIVE on the open form: a Page off the
    /// first MultiPage, a Tab off the first TabStrip. Falls back to the bare description - which
    /// for these two is empty - when no form is open or the form has none of them, because an
    /// empty list is the truth and a written-down one would be a guess.
    /// </summary>
    private static FormMarkupKind DescribedLive(
        ControlDefaults defaults, PropertyTypes types, DispatchObject? form, string kind)
    {
        var bare = Describe(defaults, kind);
        if (form is null)
        {
            return bare;
        }

        try
        {
            using var designer = form.GetObject("Designer");
            using var live = designer is null ? null : FirstOfKind(designer, kind);
            if (live is null)
            {
                return bare;
            }

            var properties = PropertiesOf(types, live);
            return properties.Length > 0
                ? bare with { Properties = properties }
                : bare;
        }
        catch (Exception)
        {
            return bare;
        }
        finally
        {
            FormDesignService.KeepDesignerDown(form);
        }
    }

    /// <summary>The first Page of any MultiPage on the form, or the first Tab of any TabStrip.</summary>
    private static DispatchObject? FirstOfKind(DispatchObject designer, string kind)
    {
        var collection = string.Equals(kind, "Page", StringComparison.OrdinalIgnoreCase) ? "Pages" : "Tabs";
        using var controls = designer.GetObject("Controls");
        if (controls is null)
        {
            return null;
        }

        foreach (var control in FormDesignService.ItemsOf(controls))
        {
            using (control)
            {
                if (control.GetDispId(collection) == DispId.Unknown)
                {
                    continue;
                }

                using var items = control.GetObject(collection);
                if ((items?.GetInt32("Count") ?? 0) > 0)
                {
                    return items!.GetItem(0);
                }
            }
        }

        return null;
    }

    private static FormMarkupKind Describe(ControlDefaults defaults, string kind)
    {
        var known = defaults.Describe(kind);
        return new FormMarkupKind(
            kind,
            known.ProgId,
            Containers.Contains(kind),
            // Only what an apply can WRITE: a completion is an offer to put a line in a document,
            // and a document line the model will not take is an offer of a refusal.
            [.. known.Properties.Where(one => one.Settable).Select(Property)],
            Descriptions.GetValueOrDefault(kind));
    }

    private static FormMarkupProperty Property(ControlDefaults.Known known) => new(
        known.Name,
        known.Type,
        known.Default,
        known.Doc,
        known.Members is { Count: > 0 } members
            ? [.. members.Select(member => new FormMarkupEnumMember(member.Name, member.Value))]
            : null,
        known.Colour,
        ValuesFor(known.Name));

    /// <summary>
    /// The values a property can be OFFERED, where they are not an enum's members but the machine
    /// can still be asked. A font face is the only one today, and it comes from the same walk the
    /// Properties panel's Font.Name row offers - one measurement, two surfaces, so the document
    /// and the panel cannot hold different ideas of what fonts this machine has.
    /// </summary>
    private static string[]? ValuesFor(string property) =>
        property.Equals("FontName", StringComparison.OrdinalIgnoreCase)
            ? [.. InstalledFonts.All]
            : null;

    /// <summary>
    /// The form's own properties, from the designer the panel reads and the component's Properties
    /// bag beside it - the two slots a form-level write goes through, in that order, so what is
    /// offered here is what an apply can land. Empty when no form is open, which is honest: the
    /// document's Form line still completes its geometry from the grammar.
    /// </summary>
    private static FormMarkupProperty[] FormProperties(PropertyTypes types, DispatchObject? component)
    {
        if (component is null)
        {
            return [];
        }

        try
        {
            using var designer = component.GetObject("Designer");
            if (designer is null)
            {
                return [];
            }

            return PropertiesOf(types, designer);
        }
        finally
        {
            // Touching Designer materialises the designer window; every walk in this product
            // puts it back down, and this one is no different.
            FormDesignService.KeepDesignerDown(component);
        }
    }

    /// <summary>
    /// One live object's properties, shaped for the page: what it holds now, what its type library
    /// says about it, and the clauses a header carries left out. The Form, a Page and a Tab all
    /// come through here - the three kinds with no coclass to measure bare.
    /// </summary>
    private static FormMarkupProperty[] PropertiesOf(PropertyTypes types, DispatchObject live)
    {
        var described = types.Describe(live);
        var properties = new List<FormMarkupProperty>();
        foreach (var name in PropertyTypes.PropertyNames(live))
        {
            // The header's own clauses, and the members no document line can carry: a form's
            // controls, its active control, the picture slots.
            if (NotAFormLine.Contains(name))
            {
                continue;
            }

            described.TryGetValue(name, out var meaning);
            string? value = null;
            try
            {
                var (variant, display) = live.ReadProperty(name);
                if (!ControlDefaults.Printable(variant))
                {
                    continue;
                }

                value = display;
            }
            catch (Exception)
            {
                // Offered without a value rather than dropped: the developer may set what this
                // object declines to read back.
            }

            properties.Add(new FormMarkupProperty(
                name, meaning?.Type, value, meaning?.Doc,
                meaning?.Members is { Count: > 0 } members
                    ? [.. members.Select(member => new FormMarkupEnumMember(member.Name, member.Value))]
                    : null,
                meaning?.Colour ?? false,
                ValuesFor(name)));
        }

        return [.. properties];
    }

    /// <summary>What a Form line never carries: the header's own clauses, and the members that
    /// name live state rather than design.</summary>
    private static readonly HashSet<string> NotAFormLine = new(StringComparer.OrdinalIgnoreCase)
    {
        "Name", "Caption", "Left", "Top", "Width", "Height", "Parent", "Object", "Application",
        "Controls", "ActiveControl", "Font", "Picture", "MouseIcon", "Value", "Tag",
        "InsideHeight", "InsideWidth", "OldHeight", "OldLeft", "OldTop", "OldWidth",
        "CurTargetX", "CurX", "Selected", "DesignMode", "CanPaste", "CanRedo", "CanUndo",
    };
}
