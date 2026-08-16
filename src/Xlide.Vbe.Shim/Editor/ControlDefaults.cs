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
/// </summary>
internal sealed class ControlDefaults
{
    /// <summary>IID_IDispatch: the only interface this needs of a control it will not site.</summary>
    private static readonly Guid Dispatch = new("00020400-0000-0000-C000-000000000046");

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

    private readonly Dictionary<string, IReadOnlyDictionary<string, string>> _byKind =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// One kind's untouched values, by property name. Empty for a kind with no ProgID of ours -
    /// a third-party control - and empty if the coclass will not come up, which is an ordinary
    /// answer: the projection then prints that kind's header alone, exactly as it did before.
    /// </summary>
    public IReadOnlyDictionary<string, string> For(string kind)
    {
        if (_byKind.TryGetValue(kind, out var known))
        {
            return known;
        }

        var read = Read(kind);
        _byKind[kind] = read;
        return read;
    }

    private static Dictionary<string, string> Read(string kind)
    {
        var empty = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var progId = FormDesignService.ProgIdFor(kind);
        if (progId is null)
        {
            return empty;
        }

        var hr = Win32.CLSIDFromProgID(progId, out var classId);
        if (hr < 0)
        {
            Log.Info($"defaults: {kind} has no class registered for {progId}, 0x{hr:X8}");
            return empty;
        }

        hr = Win32.CoCreateInstance(in classId, 0, Win32.ClassContextInProcessServer, in Dispatch, out var instance);
        if (hr < 0 || instance == 0)
        {
            Log.Info($"defaults: {kind} would not come up outside a form, 0x{hr:X8}");
            return empty;
        }

        using var bare = DispatchObject.Attach(instance);
        if (bare is null)
        {
            return empty;
        }

        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in PropertyTypes.PropertyNames(bare))
        {
            if (NotWorthPrinting.Contains(name))
            {
                continue;
            }

            try
            {
                var (variant, display) = bare.ReadProperty(name);
                if (Printable(variant))
                {
                    defaults[name] = display;
                }
            }
            catch (Exception)
            {
                // A property a bare control will not answer has no default, and a projection
                // that cannot compare says nothing rather than printing a difference it invented.
            }
        }

        Log.Info($"defaults: {kind} answers {defaults.Count} property value(s) untouched");
        return defaults;
    }

    /// <summary>The kinds the markup can spell: everything else is an object, and an object is
    /// not a value a document can carry.</summary>
    private static bool Printable(VarEnum kind) => kind is VarEnum.VT_BSTR or VarEnum.VT_BOOL
        or VarEnum.VT_I2 or VarEnum.VT_I4 or VarEnum.VT_INT or VarEnum.VT_R4 or VarEnum.VT_R8;
}
