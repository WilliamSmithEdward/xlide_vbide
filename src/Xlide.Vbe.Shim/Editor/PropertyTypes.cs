using System.Globalization;
using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// What a property's value MEANS, read from the type library of the object holding it.
///
/// The Properties panel had been showing what the object model hands over: `BorderStyle 0`,
/// `Cycle 0`, `MousePointer 0`, `BackColor -2147483633`. Every one of those is a name in the
/// language the developer writes - `fmBorderStyleNone`, `fmCycleAllForms`, `fmMousePointerDefault`,
/// `&amp;H8000000F&amp;` - and the panel that shows the number instead is asking its reader to
/// keep the enum tables in their head (the owner, 2026-08-15: "can we get the actual enums
/// loaded? instead of the int representation").
///
/// The names come from the same place the Object Browser's do: the object's own ITypeInfo. For
/// each property getter, the return type is followed - through an alias if it is one - and where
/// it lands on an enum, that enum's members are read with their values. Nothing is guessed from
/// a property's NAME, which is the trap this could have been: `fmBorderStyle` for `BorderStyle`
/// is a convention until the day a control names one differently, and a panel that renames a
/// value wrongly is worse than one that shows the number.
///
/// Read once per CLASS and cached: a form's type does not change while it is loaded, and the
/// panel republishes on every selection.
///
/// The same walk answers the markup's language service (2026-08-16), which wants two things the
/// panel never needed: the type each property is DECLARED as, and the help string the library
/// carries for it. Collecting them costs nothing extra - both were already in hand and thrown
/// away - so there is one walk, one cache, and two views of it.
/// </summary>
internal sealed unsafe class PropertyTypes
{
    /// <summary>One member of an enum: what the developer writes, and what the model stores.</summary>
    public sealed record EnumMember(string Name, int Value);

    /// <summary>What is known about one property: its enum's members, or that it is a colour.</summary>
    public sealed record Shape(IReadOnlyList<EnumMember>? Members, bool Colour);

    /// <summary>
    /// Everything the library says about one property: the type it is declared as, its enum's
    /// members where it has one, whether it is a colour, and the help string the library carries.
    ///
    /// The panel needs the middle two and gets them as <see cref="Shape"/>, projected from this;
    /// the markup's language service needs all four, because a completion that offers a property
    /// name without saying what it takes is a list of words. One walk answers both.
    /// </summary>
    public sealed record Described(
        string? Type, IReadOnlyList<EnumMember>? Members, bool Colour, string? Doc,
        bool Hidden = false, bool Settable = false);

    private const int InvokePropertyGet = 2;
    private const int InvokePropertyPut = 4;
    private const int InvokePropertyPutRef = 8;
    private const int TypeKindEnum = 0;
    private const int TypeKindAlias = 6;
    private const short VtI2 = 2;
    private const short VtI4 = 3;
    private const short VtInt = 22;
    private const short VtPointer = 26;
    private const short VtUserDefined = 29;

    /// <summary>FUNCFLAG_FRESTRICTED and FUNCFLAG_FHIDDEN: what a library marks as not for the
    /// developer's eyes. The Object Browser filters on exactly these two.</summary>
    private const short FunctionFlagRestricted = 0x1;
    private const short FunctionFlagHidden = 0x40;

    /// <summary>How far an alias chain is followed before it is called a loop.</summary>
    private const int AliasDepth = 4;

    private static readonly IReadOnlyDictionary<string, Shape> Nothing =
        new Dictionary<string, Shape>(StringComparer.OrdinalIgnoreCase);

    private static readonly IReadOnlyDictionary<string, Described> NothingDescribed =
        new Dictionary<string, Described>(StringComparer.OrdinalIgnoreCase);

    private readonly Dictionary<string, IReadOnlyDictionary<string, Described>> _byClass =
        new(StringComparer.OrdinalIgnoreCase);

    private readonly Dictionary<string, IReadOnlyDictionary<string, Shape>> _shapesByClass =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The shapes of one object's properties, by property name. Empty when the object has no type
    /// information to offer, which is an ordinary answer rather than a failure: a panel that shows
    /// the raw value is what this product did until today, and it is still honest.
    ///
    /// Only the properties that MEAN something beyond their number appear here - an enum's, a
    /// colour's - because that is the question the panel asks: is there a name for this value.
    /// </summary>
    public IReadOnlyDictionary<string, Shape> Of(DispatchObject subject)
    {
        ArgumentNullException.ThrowIfNull(subject);

        var className = subject.TypeName();
        if (className is not { Length: > 0 })
        {
            return Nothing;
        }

        if (_shapesByClass.TryGetValue(className, out var known))
        {
            return known;
        }

        var shapes = new Dictionary<string, Shape>(StringComparer.OrdinalIgnoreCase);
        foreach (var (name, described) in Describe(subject))
        {
            if (described.Members is { Count: > 0 } || described.Colour)
            {
                shapes[name] = new Shape(described.Members, described.Colour);
            }
        }

        _shapesByClass[className] = shapes;
        return shapes;
    }

    /// <summary>
    /// Everything the library says about every property the object answers, by name. The richer
    /// read behind <see cref="Of"/>, and the vocabulary the markup's completions and hovers are
    /// built from. Cached per class beside the shapes, off the same walk.
    /// </summary>
    public IReadOnlyDictionary<string, Described> Describe(DispatchObject subject)
    {
        ArgumentNullException.ThrowIfNull(subject);

        var className = subject.TypeName();
        if (className is not { Length: > 0 })
        {
            return NothingDescribed;
        }

        if (_byClass.TryGetValue(className, out var known))
        {
            return known;
        }

        var read = Read(subject, className);
        _byClass[className] = read;
        return read;
    }

    /// <summary>
    /// Every property an object will answer with no arguments, by name - the list the defaults
    /// inventory reads and the projection compares against. Getters only, and no leading
    /// underscore: those are the library's own business rather than the developer's.
    /// </summary>
    public static IReadOnlyList<string> PropertyNames(DispatchObject subject)
    {
        ArgumentNullException.ThrowIfNull(subject);

        using var info = subject.TypeInfo();
        if (info is null)
        {
            return [];
        }

        var names = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            NamesInto(info, names, seen, depth: 0);
        }
        catch (Exception ex)
        {
            Log.Info($"property types: a property list could not be read ({ex.GetType().Name})");
        }

        return names;
    }

    private static void NamesInto(
        ComHandle<ITypeInfo> info, List<string> names, HashSet<string> seen, int depth)
    {
        if (info.Target.GetTypeAttr(out var attrPointer) < 0 || attrPointer == 0)
        {
            return;
        }

        var attr = (TypeAttr*)attrPointer;
        var kind = attr->TypeKind;
        var funcCount = attr->FuncCount;
        var implCount = attr->ImplTypeCount;
        info.Target.ReleaseTypeAttr(attrPointer);

        if (kind != 5)
        {
            for (var i = 0; i < funcCount; i++)
            {
                if (info.Target.GetFuncDesc(i, out var funcPointer) < 0 || funcPointer == 0)
                {
                    continue;
                }

                try
                {
                    var function = (FuncDesc*)funcPointer;
                    // Zero parameters: an indexed property is a collection reader, and reading
                    // one without an index says nothing about a default.
                    if (function->InvokeKind != InvokePropertyGet || function->ParameterCount != 0)
                    {
                        continue;
                    }

                    if (NameOf(info, function->MemberId) is { Length: > 0 } name
                        && !name.StartsWith('_') && seen.Add(name))
                    {
                        names.Add(name);
                    }
                }
                finally
                {
                    info.Target.ReleaseFuncDesc(funcPointer);
                }
            }
        }

        if (depth >= 3)
        {
            return;
        }

        for (var i = 0; i < implCount; i++)
        {
            if (info.Target.GetRefTypeOfImplType(i, out var refType) < 0
                || info.Target.GetRefTypeInfo(refType, out var otherPointer) < 0 || otherPointer == 0)
            {
                continue;
            }

            using var other = ComHandle<ITypeInfo>.Own(otherPointer);
            if (other is not null)
            {
                NamesInto(other, names, seen, depth + 1);
            }
        }
    }

    /// <summary>The value as the developer would write it: an enum member's name, a colour's
    /// VBA hex, or the number as it stands when neither applies.</summary>
    public static string Spell(Shape? shape, string raw)
    {
        if (shape is null || !int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
        {
            return raw;
        }

        if (shape.Members is { Count: > 0 } members)
        {
            return members.FirstOrDefault(member => member.Value == value)?.Name ?? raw;
        }

        // A colour reads the way a developer writes one - `#f0f0f0`, or a system colour's name
        // where there is no honest hex to give (SystemColours has the why).
        return shape.Colour
            ? SystemColours.Spell(value)
            : raw;
    }

    /// <summary>The number behind what the panel shows, for the write. Anything this cannot read
    /// as a member name or a hex literal goes to the model untouched, so a developer who types a
    /// number still gets a number.</summary>
    public static string Unspell(Shape? shape, string shown)
    {
        if (shape is null || shown.Length == 0)
        {
            return shown;
        }

        if (shape.Members is { Count: > 0 } members)
        {
            var named = members.FirstOrDefault(member =>
                string.Equals(member.Name, shown, StringComparison.OrdinalIgnoreCase));
            return named is null ? shown : named.Value.ToString(CultureInfo.InvariantCulture);
        }

        // Every spelling a colour row can show, and the two the VBE speaks besides. Anything
        // else goes to the model untouched, so a developer who types a number gets a number.
        if (shape.Colour && SystemColours.Unspell(shown) is { } colour)
        {
            return colour.ToString(CultureInfo.InvariantCulture);
        }

        return shown;
    }

    private static IReadOnlyDictionary<string, Described> Read(DispatchObject subject, string className)
    {
        using var info = subject.TypeInfo();
        if (info is null)
        {
            return NothingDescribed;
        }

        var described = new Dictionary<string, Described>(StringComparer.OrdinalIgnoreCase);
        var settable = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            ReadInto(info, described, settable, depth: 0);
        }
        catch (Exception ex)
        {
            Log.Info($"property types: '{className}' could not be read ({ex.GetType().Name})");
        }

        // A property is writable when the library declares a PUT for it, which is a fact about
        // the type rather than about the value that came back: `CanPaste` reads as a Boolean and
        // is no more settable for it.
        foreach (var name in described.Keys.ToArray())
        {
            described[name] = described[name] with { Settable = settable.Contains(name) };
        }

        var named = described.Count(one => one.Value.Members is { Count: > 0 } || one.Value.Colour);
        Log.Info($"property types: {className} carries {described.Count} property(s), {named} of them named value(s)");
        return described;
    }

    private static void ReadInto(
        ComHandle<ITypeInfo> info, Dictionary<string, Described> described,
        HashSet<string> settable, int depth)
    {
        if (info.Target.GetTypeAttr(out var attrPointer) < 0 || attrPointer == 0)
        {
            return;
        }

        var attr = (TypeAttr*)attrPointer;
        var kind = attr->TypeKind;
        var funcCount = attr->FuncCount;
        var implCount = attr->ImplTypeCount;
        info.Target.ReleaseTypeAttr(attrPointer);

        // A coclass keeps its members on its default interface, exactly as the Object Browser
        // finds them; the dispatch a UserForm hands over is usually the coclass.
        if (kind == 5 && depth < 3)
        {
            for (var i = 0; i < implCount; i++)
            {
                if (info.Target.GetRefTypeOfImplType(i, out var refType) < 0
                    || info.Target.GetRefTypeInfo(refType, out var implPointer) < 0 || implPointer == 0)
                {
                    continue;
                }

                using var implemented = ComHandle<ITypeInfo>.Own(implPointer);
                if (implemented is not null)
                {
                    ReadInto(implemented, described, settable, depth + 1);
                }
            }

            return;
        }

        for (var i = 0; i < funcCount; i++)
        {
            if (info.Target.GetFuncDesc(i, out var funcPointer) < 0 || funcPointer == 0)
            {
                continue;
            }

            try
            {
                var function = (FuncDesc*)funcPointer;
                if (function->InvokeKind is InvokePropertyPut or InvokePropertyPutRef)
                {
                    if (NameOf(info, function->MemberId) is { Length: > 0 } putName)
                    {
                        settable.Add(putName);
                    }

                    continue;
                }

                if (function->InvokeKind != InvokePropertyGet)
                {
                    continue;
                }

                var (name, doc) = DocumentedOf(info, function->MemberId);
                if (name is not { Length: > 0 } || described.ContainsKey(name))
                {
                    continue;
                }

                var (type, members, colour) = TypeOf(info, function->Return.Type, AliasDepth);
                described[name] = new Described(
                    type, members, colour, doc,
                    (function->Flags & (FunctionFlagHidden | FunctionFlagRestricted)) != 0);
            }
            finally
            {
                info.Target.ReleaseFuncDesc(funcPointer);
            }
        }

        // An interface's own base carries the rest - a control's properties are split across the
        // interface it declares and the ones it inherits.
        if (depth < 3)
        {
            for (var i = 0; i < implCount; i++)
            {
                if (info.Target.GetRefTypeOfImplType(i, out var refType) < 0
                    || info.Target.GetRefTypeInfo(refType, out var basePointer) < 0 || basePointer == 0)
                {
                    continue;
                }

                using var inherited = ComHandle<ITypeInfo>.Own(basePointer);
                if (inherited is not null)
                {
                    ReadInto(inherited, described, settable, depth + 1);
                }
            }
        }
    }

    /// <summary>
    /// What a return type IS and what it MEANS: the name it is declared as, an enum's members
    /// where it has them, and whether it is a colour. The name is the type's OWN - `fmBorderStyle`,
    /// `OLE_COLOR` - because that is what a developer looking for the property's vocabulary would
    /// search for; an alias is followed for its members but keeps its own spelling.
    /// </summary>
    private static (string? Type, IReadOnlyList<EnumMember>? Members, bool Colour) TypeOf(
        ComHandle<ITypeInfo> info, TypeDesc type, int depth)
    {
        if (depth <= 0)
        {
            return (null, null, false);
        }

        // A pointer to a type is that type as far as a property's spelling is concerned: an
        // object-valued getter declares IFontDisp* and the developer writes Font.
        if (type.VarType == VtPointer)
        {
            return type.Detail == 0
                ? (null, null, false)
                : TypeOf(info, *(TypeDesc*)type.Detail, depth - 1);
        }

        if (type.VarType != VtUserDefined)
        {
            return (Primitive(type.VarType), null, false);
        }

        if (info.Target.GetRefTypeInfo((int)type.Detail, out var referenced) < 0 || referenced == 0)
        {
            return (null, null, false);
        }

        using var target = ComHandle<ITypeInfo>.Own(referenced);
        if (target is null)
        {
            return (null, null, false);
        }

        if (target.Target.GetTypeAttr(out var attrPointer) < 0 || attrPointer == 0)
        {
            return (null, null, false);
        }

        var attr = (TypeAttr*)attrPointer;
        var kind = attr->TypeKind;
        var alias = attr->Alias;
        var varCount = attr->VarCount;
        target.Target.ReleaseTypeAttr(attrPointer);

        var name = NameOf(target, -1);

        if (kind == TypeKindEnum)
        {
            var members = MembersOf(target, varCount);
            return (name, members.Count == 0 ? null : members, false);
        }

        if (kind != TypeKindAlias)
        {
            return (name, null, false);
        }

        // OLE_COLOR is an alias for a plain unsigned long, so only its NAME says what it is -
        // and this is the one place a name decides anything, because the type system has
        // nothing else to offer here.
        if (string.Equals(name, "OLE_COLOR", StringComparison.OrdinalIgnoreCase))
        {
            return (name, null, true);
        }

        var (_, aliasMembers, aliasColour) = TypeOf(target, alias, depth - 1);
        return (name, aliasMembers, aliasColour);
    }

    /// <summary>A variant tag as the language spells it, for the types that have no name of
    /// their own. Null where a tag says nothing a developer would write.</summary>
    private static string? Primitive(short varType) => varType switch
    {
        VtI2 => "Integer",
        VtI4 or VtInt => "Long",
        4 => "Single",
        5 => "Double",
        6 => "Currency",
        7 => "Date",
        8 => "String",
        9 or 13 => "Object",
        11 => "Boolean",
        12 => "Variant",
        17 => "Byte",
        20 => "LongLong",
        _ => null,
    };

    private static List<EnumMember> MembersOf(ComHandle<ITypeInfo> info, short varCount)
    {
        var members = new List<EnumMember>(varCount);
        for (var i = 0; i < varCount; i++)
        {
            if (info.Target.GetVarDesc(i, out var varPointer) < 0 || varPointer == 0)
            {
                continue;
            }

            try
            {
                var variable = (VarDesc*)varPointer;
                if (variable->VarKind != 2 || variable->Value == 0)
                {
                    continue;
                }

                if (ConstantOf(variable->Value) is not { } value)
                {
                    continue;
                }

                if (NameOf(info, variable->MemberId) is { Length: > 0 } name)
                {
                    members.Add(new EnumMember(name, value));
                }
            }
            finally
            {
                info.Target.ReleaseVarDesc(varPointer);
            }
        }

        return members;
    }

    /// <summary>An enum member's value, for the integral kinds an enum can hold.</summary>
    private static int? ConstantOf(nint variantPointer)
    {
        var varType = *(short*)variantPointer;
        var data = variantPointer + 8;
        return varType switch
        {
            VtI2 => *(short*)data,
            VtI4 or VtInt => *(int*)data,
            _ => null,
        };
    }

    /// <summary>A member's name, or the type's own for MEMBERID_NIL.</summary>
    private static string? NameOf(ComHandle<ITypeInfo> info, int memberId) => DocumentedOf(info, memberId).Name;

    /// <summary>A member's name and the help string its library carries, which is the sentence the
    /// Object Browser's bottom strip shows. Often absent - MSForms describes some members and not
    /// others - and an absent one is simply not said rather than invented.</summary>
    private static (string? Name, string? Doc) DocumentedOf(ComHandle<ITypeInfo> info, int memberId)
    {
        if (info.Target.GetDocumentation(memberId, out var name, out var documentation, out _, out var helpFile) < 0)
        {
            return (null, null);
        }

        try
        {
            return (
                name == 0 ? null : Marshal.PtrToStringBSTR(name),
                documentation == 0 ? null : Marshal.PtrToStringBSTR(documentation));
        }
        finally
        {
            if (name != 0) { Marshal.FreeBSTR(name); }
            if (documentation != 0) { Marshal.FreeBSTR(documentation); }
            if (helpFile != 0) { Marshal.FreeBSTR(helpFile); }
        }
    }
}
