using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// What a control of each kind holds when nobody has touched it: the inventory the markup
/// projection needs to say which properties a developer has actually CHANGED.
///
/// The document had been describing identity, geometry and caption and nothing else, so a font,
/// a colour or an alignment set through the Properties panel lived only in the object model -
/// invisible in the text, absent from the draft preview, and outside the document's undo. To
/// print the rest, something has to know which values are worth printing, and that means knowing
/// the defaults (the owner, 2026-08-16: "i think u should inventory all defaults for every
/// control").
///
/// The inventory is MEASURED, never written down here. Each kind's ProgID is the one the add
/// route already accepts, and MSForms registers those coclasses in-process: a bare instance is
/// created, every property it will answer is read, and that reading IS the default for that kind.
/// No workbook is opened, no form is created, nothing appears on screen - the owner's line was
/// "i absolutely do not want xlide to have to open any files" - and no table of ours can rot
/// against a version of MSForms it was written before.
///
/// Two honest limits. A property a bare instance refuses without a container (MultiPage's Style
/// is one) has no default here, so it is never printed as changed: the projection stays silent
/// rather than guessing. And a FONT is inherited from the form rather than defaulted per control,
/// so it is compared against the form's own - that comparison belongs to the walk, not here.
///
/// The bare instance answers a second question while it is up (2026-08-16): what each of those
/// properties MEANS - the type it is declared as, its enum's members, its help string - which is
/// the vocabulary the markup's completions and hovers are built from. Same instance, same walk,
/// so the language service and the projection can never hold different ideas of what a Label has.
///
/// A property the instance REFUSES is still OFFERED, without a default: the developer may set what
/// a control outside a container will not read back, and a MultiPage - which answers one value of
/// its fifteen - would otherwise complete nothing at all. And a kind that answers badly does not
/// take the others down with it: a bare MultiPage throws out of its own type library rather than
/// declining, which silenced the whole vocabulary until each half of the read was guarded.
/// </summary>
internal sealed class ControlDefaults(PropertyTypes types)
{
    /// <summary>IID_IDispatch: the only interface this needs of a control it will not site.</summary>
    private static readonly Guid Dispatch = new("00020400-0000-0000-C000-000000000046");

    private readonly PropertyTypes _types = types;

    /// <summary>One property of one kind, as the language service offers it: what it is called,
    /// what it takes, what it holds untouched, and what the library says about it.</summary>
    public sealed record Known(
        string Name,
        string? Type,
        string? Default,
        string? Doc,
        IReadOnlyList<PropertyTypes.EnumMember>? Members,
        bool Colour,
        /// <summary>Whether the library declares a setter. A document line is an instruction to
        /// an apply and a completion is an offer to write one, so both consumers here keep to
        /// what can actually be written; the raw inventory still reports the rest.</summary>
        bool Settable = false);

    /// <summary>What one control kind offers a document: its ProgID, and every property a bare
    /// instance answers with a value the markup can spell.</summary>
    public sealed record Vocabulary(string Kind, string? ProgId, IReadOnlyList<Known> Properties);

    /// <summary>
    /// Never printed even when they differ. Identity and placement are the header's business,
    /// and the rest name things that cannot mean the same twice: a tab order in a form the
    /// document does not describe, a parent, a live value the developer typed at runtime.
    /// </summary>
    private static readonly HashSet<string> NotWorthPrinting = new(StringComparer.OrdinalIgnoreCase)
    {
        "Name", "Left", "Top", "Width", "Height", "Caption", "TabIndex", "TabStop", "Parent",
        "Object", "Application", "Value", "Text", "Tag", "Index", "OldHeight", "OldLeft",
        "OldTop", "OldWidth", "Font", "Picture", "MouseIcon", "Controls", "Pages", "Tabs",
        "ActiveControl", "Selected", "ListIndex", "ListCount", "TopIndex", "CurTargetX",
        "CurX", "InsideHeight", "InsideWidth", "ScrollHeight", "ScrollWidth", "ScrollLeft",
        "ScrollTop", "SelLength", "SelStart", "SelText", "LineCount", "TextLength",
    };

    private readonly Dictionary<string, Vocabulary> _byKind = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// One kind's untouched values, by property name. Empty for a kind with no ProgID of ours -
    /// a third-party control - and empty if the coclass will not come up, which is an ordinary
    /// answer: the projection then prints that kind's header alone, exactly as it did before.
    /// </summary>
    public IReadOnlyDictionary<string, string> For(string kind)
    {
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in Describe(kind).Properties)
        {
            if (property.Default is { } value)
            {
                defaults[property.Name] = value;
            }
        }

        return defaults;
    }

    /// <summary>
    /// One kind's whole vocabulary: every property a bare instance answers, what it takes, what it
    /// holds untouched, and what its library says about it. Measured once per kind and cached for
    /// the session, because a registered coclass does not change while Excel is up.
    /// </summary>
    public Vocabulary Describe(string kind)
    {
        if (_byKind.TryGetValue(kind, out var known))
        {
            return known;
        }

        var read = Read(kind);
        _byKind[kind] = read;
        return read;
    }

    private Vocabulary Read(string kind)
    {
        var progId = FormDesignService.ProgIdFor(kind);
        var nothing = new Vocabulary(kind, progId, []);
        if (progId is null)
        {
            return nothing;
        }

        var hr = Win32.CLSIDFromProgID(progId, out var classId);
        if (hr < 0)
        {
            Log.Info($"defaults: {kind} has no class registered for {progId}, 0x{hr:X8}");
            return nothing;
        }

        hr = Win32.CoCreateInstance(in classId, 0, Win32.ClassContextInProcessServer, in Dispatch, out var instance);
        if (hr < 0 || instance == 0)
        {
            Log.Info($"defaults: {kind} would not come up outside a form, 0x{hr:X8}");
            return nothing;
        }

        using var bare = DispatchObject.Attach(instance);
        if (bare is null)
        {
            return nothing;
        }

        // A bare control is a control MSForms never sited, and some of them answer badly: a
        // MultiPage outside a container throws from its own type library rather than declining
        // (measured 2026-08-16, where it took the whole vocabulary down with it). Each half of
        // the read is guarded on its own, and what survives is offered - the alternative is one
        // awkward kind silencing the completions for all fifteen.
        IReadOnlyDictionary<string, PropertyTypes.Described> described;
        try
        {
            described = _types.Describe(bare);
        }
        catch (Exception ex)
        {
            Log.Info($"defaults: {kind}'s type library would not be read ({ex.GetType().Name})");
            described = new Dictionary<string, PropertyTypes.Described>(StringComparer.OrdinalIgnoreCase);
        }

        var properties = new List<Known>();
        foreach (var name in PropertyTypes.PropertyNames(bare))
        {
            if (NotWorthPrinting.Contains(name))
            {
                continue;
            }

            described.TryGetValue(name, out var meaning);

            string? value = null;
            try
            {
                var (variant, display) = bare.ReadProperty(name);
                if (Printable(variant))
                {
                    value = display;
                }
                else
                {
                    // An object-valued property is not a value a document can carry, so it is
                    // not offered as one either - the language service lists what can be written.
                    continue;
                }
            }
            catch (Exception)
            {
                // A property a bare control will not answer has no default, and a projection
                // that cannot compare says nothing rather than printing a difference it invented.
                // It is still worth OFFERING: the developer may set what this instance refuses.
            }

            properties.Add(new Known(
                name, meaning?.Type, value, meaning?.Doc, meaning?.Members,
                meaning?.Colour ?? false, meaning?.Settable ?? false));
        }

        var withDefaults = properties.Count(one => one.Default is not null);
        Log.Info($"defaults: {kind} answers {withDefaults} property value(s) untouched, "
            + $"{properties.Count} property(s) in its vocabulary");
        return new Vocabulary(kind, progId, properties);
    }

    /*
     * NO DOTTED FONT PATHS, and the measurement is why.
     *
     * The dialect spells one level of dotting - `Font.Size = 12` is in this document's own example
     * - and the apply reaches it, so the first cut walked each bare control's font object and
     * offered `Font.Name`, `Font.Size` and the rest. It offered nothing at all, measured
     * 2026-08-16: an OLE font is a vtable interface whose getters declare their result as a retval
     * PARAMETER, and the property walk counts a parameter as "this is indexed, skip it".
     *
     * Chasing that turned out to be the wrong fix. A control's own extender already carries
     * `FontName`, `FontSize`, `FontBold` and `FontItalic` flat, they measure cleanly, and an apply
     * writes them - so the dotted spelling would have been a SECOND way to say what the list
     * already says. Two spellings of one property is a worse completion list, not a richer one.
     * The dotted form still parses and still applies for anyone who types it; it is simply not
     * offered.
     */

    /// <summary>The kinds the markup can spell: everything else is an object, and an object is
    /// not a value a document can carry. VT_CY is here because a FONT SIZE is currency on the
    /// wire - MSForms' own choice - and without the case `FontSize` had no measured default.</summary>
    internal static bool Printable(VarEnum kind) => kind is VarEnum.VT_BSTR or VarEnum.VT_BOOL
        or VarEnum.VT_I2 or VarEnum.VT_I4 or VarEnum.VT_INT or VarEnum.VT_R4 or VarEnum.VT_R8
        or VarEnum.VT_CY;
}
