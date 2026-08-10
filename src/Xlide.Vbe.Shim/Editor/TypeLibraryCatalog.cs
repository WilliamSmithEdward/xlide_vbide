using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The Object Browser's data: every type library a project references, read into browsable
/// rows - libraries, their types, and each type's members with signatures rendered the way
/// VBA would spell them.
///
/// Libraries load once and stay loaded; types and members are read on first request and
/// cached, because a library's contents cannot change while it is loaded. Everything the
/// type machinery hands out is caller-released, and every release happens here.
/// </summary>
internal sealed unsafe class TypeLibraryCatalog : IDisposable
{
    public sealed record LibraryRow(string Name, string Description);
    public sealed record TypeRow(string Name, string Kind);
    public sealed record MemberRow(string Name, string Kind, string Signature, string Description);

    private sealed record LoadedType(int Index, string Name, string Kind);

    private sealed class LoadedLibrary
    {
        public required string Name;
        public required string Description;
        public required string Path;
        public required ComHandle<ITypeLib> Library;
        public List<LoadedType>? Types;
        public readonly Dictionary<string, IReadOnlyList<MemberRow>> Members = new(StringComparer.OrdinalIgnoreCase);
    }

    private readonly List<LoadedLibrary> _libraries = [];

    private const int TypeKindEnum = 0;
    private const int TypeKindRecord = 1;
    private const int TypeKindModule = 2;
    private const int TypeKindInterface = 3;
    private const int TypeKindDispatch = 4;
    private const int TypeKindCoClass = 5;

    private const short TypeFlagRestricted = 0x1;
    private const short TypeFlagHidden = 0x10;
    private const short FunctionFlagRestricted = 0x1;
    private const short FunctionFlagHidden = 0x40;
    private const short ParameterFlagOptional = 0x10;
    private const short ParameterFlagReturnValue = 0x8;

    /// <summary>Adds a referenced library by path; a repeat of the same path is a no-op.</summary>
    public void AddLibrary(string name, string description, string path)
    {
        if (path.Length == 0
            || _libraries.Any(held => string.Equals(held.Path, path, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var hr = TypeLibraryNative.LoadTypeLibEx(path, TypeLibraryNative.LoadWithoutRegistering, out var pointer);
        if (hr < 0 || pointer == 0)
        {
            Log.Info($"typelib: '{name}' would not load from {path} (0x{hr:X8})");
            return;
        }

        var library = ComHandle<ITypeLib>.Own(pointer);
        if (library is null)
        {
            return;
        }

        _libraries.Add(new LoadedLibrary
        {
            Name = name,
            Description = description,
            Path = path,
            Library = library,
        });

        Log.Info($"typelib: '{name}' loaded ({description})");
    }

    public IReadOnlyList<LibraryRow> Libraries() =>
        [.. _libraries.Select(held => new LibraryRow(held.Name, held.Description))];

    /// <summary>The browsable types of one library, or null when it is not loaded.</summary>
    public IReadOnlyList<TypeRow>? TypesOf(string libraryName)
    {
        var held = Find(libraryName);
        if (held is null)
        {
            return null;
        }

        held.Types ??= ReadTypes(held);
        return [.. held.Types.Select(type => new TypeRow(type.Name, type.Kind))];
    }

    /// <summary>One type's members, or null when the library or type is unknown.</summary>
    public IReadOnlyList<MemberRow>? MembersOf(string libraryName, string typeName)
    {
        var held = Find(libraryName);
        if (held is null)
        {
            return null;
        }

        if (held.Members.TryGetValue(typeName, out var known))
        {
            return known;
        }

        held.Types ??= ReadTypes(held);
        var type = held.Types.FirstOrDefault(
            candidate => string.Equals(candidate.Name, typeName, StringComparison.OrdinalIgnoreCase));
        if (type is null)
        {
            return null;
        }

        var members = ReadMembers(held, type);
        held.Members[typeName] = members;
        return members;
    }

    private LoadedLibrary? Find(string libraryName) =>
        _libraries.FirstOrDefault(held => string.Equals(held.Name, libraryName, StringComparison.OrdinalIgnoreCase));

    private static List<LoadedType> ReadTypes(LoadedLibrary held)
    {
        var types = new List<LoadedType>();

        try
        {
            var count = held.Library.Target.GetTypeInfoCount();

            for (var i = 0; i < count; i++)
            {
                if (held.Library.Target.GetTypeInfo(i, out var infoPointer) < 0 || infoPointer == 0)
                {
                    continue;
                }

                using var info = ComHandle<ITypeInfo>.Own(infoPointer);
                if (info is null || info.Target.GetTypeAttr(out var attrPointer) < 0 || attrPointer == 0)
                {
                    continue;
                }

                var attr = (TypeAttr*)attrPointer;
                var kind = attr->TypeKind;
                var flags = attr->TypeFlags;
                info.Target.ReleaseTypeAttr(attrPointer);

                if ((flags & (TypeFlagHidden | TypeFlagRestricted)) != 0)
                {
                    continue;
                }

                var spelled = kind switch
                {
                    TypeKindEnum => "enum",
                    TypeKindRecord => "type",
                    TypeKindModule => "module",
                    TypeKindInterface or TypeKindDispatch or TypeKindCoClass => "class",
                    _ => null,
                };

                if (spelled is null)
                {
                    continue;
                }

                var name = TakeName(held.Library.Target, i);
                if (name is not { Length: > 0 } || name.StartsWith('_'))
                {
                    continue;
                }

                // Coclasses and their default interfaces usually travel in pairs with the
                // interface underscore-hidden; where both are plainly named, the class row
                // wins and the duplicate is dropped.
                if (types.Any(existing => string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase)))
                {
                    if (spelled == "class" && kind == TypeKindCoClass)
                    {
                        types.RemoveAll(existing => string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase));
                    }
                    else
                    {
                        continue;
                    }
                }

                types.Add(new LoadedType(i, name, spelled));
            }
        }
        catch (Exception ex)
        {
            Log.Info($"typelib: '{held.Name}' types could not be read ({ex.GetType().Name})");
        }

        types.Sort(static (a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
        return types;
    }

    private static List<MemberRow> ReadMembers(LoadedLibrary held, LoadedType type)
    {
        var members = new List<MemberRow>();

        try
        {
            if (held.Library.Target.GetTypeInfo(type.Index, out var infoPointer) < 0 || infoPointer == 0)
            {
                return members;
            }

            var info = ComHandle<ITypeInfo>.Own(infoPointer);
            if (info is null)
            {
                return members;
            }

            try
            {
                ReadMembersOf(info, members, depth: 0);
            }
            finally
            {
                info.Dispose();
            }
        }
        catch (Exception ex)
        {
            Log.Info($"typelib: '{held.Name}.{type.Name}' members could not be read ({ex.GetType().Name})");
        }

        members.Sort(static (a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
        return members;
    }

    private static void ReadMembersOf(ComHandle<ITypeInfo> info, List<MemberRow> members, int depth)
    {
        if (info.Target.GetTypeAttr(out var attrPointer) < 0 || attrPointer == 0)
        {
            return;
        }

        var attr = (TypeAttr*)attrPointer;
        var kind = attr->TypeKind;
        var funcCount = attr->FuncCount;
        var varCount = attr->VarCount;
        var implCount = attr->ImplTypeCount;
        info.Target.ReleaseTypeAttr(attrPointer);

        // A coclass has no members of its own; they live on its default interface.
        if (kind == TypeKindCoClass)
        {
            if (depth < 3)
            {
                for (var i = 0; i < implCount; i++)
                {
                    if (info.Target.GetImplTypeFlags(i, out var implFlags) < 0
                        || (implFlags & 0x1) == 0 || (implFlags & 0x2) != 0)
                    {
                        continue;
                    }

                    if (info.Target.GetRefTypeOfImplType(i, out var refType) < 0
                        || info.Target.GetRefTypeInfo(refType, out var innerPointer) < 0 || innerPointer == 0)
                    {
                        continue;
                    }

                    using var inner = ComHandle<ITypeInfo>.Own(innerPointer);
                    if (inner is not null)
                    {
                        ReadMembersOf(inner, members, depth + 1);
                    }
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
                AppendFunction(info, (FuncDesc*)funcPointer, kind, members);
            }
            finally
            {
                info.Target.ReleaseFuncDesc(funcPointer);
            }
        }

        for (var i = 0; i < varCount; i++)
        {
            if (info.Target.GetVarDesc(i, out var varPointer) < 0 || varPointer == 0)
            {
                continue;
            }

            try
            {
                AppendVariable(info, (VarDesc*)varPointer, kind, members);
            }
            finally
            {
                info.Target.ReleaseVarDesc(varPointer);
            }
        }
    }

    private static void AppendFunction(ComHandle<ITypeInfo> info, FuncDesc* func, int typeKind, List<MemberRow> members)
    {
        if ((func->Flags & (FunctionFlagHidden | FunctionFlagRestricted)) != 0 || func->MemberId < 0)
        {
            return;
        }

        // A plain interface carries IUnknown and IDispatch in its first seven slots.
        if (typeKind == TypeKindInterface && func->VtableOffset < 7 * sizeof(nint))
        {
            return;
        }

        var names = NamesOf(info, func->MemberId, func->ParameterCount + 1);
        if (names.Count == 0 || names[0].StartsWith('_'))
        {
            return;
        }

        var name = names[0];

        // Dual interfaces return HRESULT and carry the real result in a retval parameter.
        var parameters = (ElemDesc*)func->Parameters;
        var returnType = RenderType(info, func->Return.Type);
        var lastIsReturn = false;

        if (func->Return.Type.VarType is VtHResult or VtVoid)
        {
            returnType = null;

            if (func->ParameterCount > 0
                && (parameters[func->ParameterCount - 1].Parameter.Flags & ParameterFlagReturnValue) != 0)
            {
                returnType = RenderType(info, parameters[func->ParameterCount - 1].Type);
                lastIsReturn = true;
            }
        }

        var rendered = new List<string>();
        var count = func->ParameterCount - (lastIsReturn ? 1 : 0);
        for (var i = 0; i < count; i++)
        {
            var parameterName = i + 1 < names.Count ? names[i + 1] : $"arg{i + 1}";
            var optional = (parameters[i].Parameter.Flags & ParameterFlagOptional) != 0;
            var part = $"{parameterName} As {RenderType(info, parameters[i].Type)}";
            rendered.Add(optional ? $"[{part}]" : part);
        }

        var (verb, isSub) = func->InvokeKind switch
        {
            2 => ("Property Get", false),
            4 => ("Property Let", true),
            8 => ("Property Set", true),
            _ => returnType is null ? ("Sub", true) : ("Function", false),
        };

        var signature = $"{verb} {name}({string.Join(", ", rendered)})"
            + (!isSub && returnType is not null ? $" As {returnType}" : string.Empty);

        // Property Get/Let pairs and overloaded dispatch entries collapse to one row per
        // name; the Get's signature is the one worth showing.
        var existing = members.FindIndex(m => string.Equals(m.Name, name, StringComparison.Ordinal));
        if (existing >= 0)
        {
            if (func->InvokeKind == 2)
            {
                members[existing] = new MemberRow(name, KindOf(func->InvokeKind, returnType), signature, DocumentationOf(info, func->MemberId));
            }

            return;
        }

        members.Add(new MemberRow(name, KindOf(func->InvokeKind, returnType), signature, DocumentationOf(info, func->MemberId)));
    }

    private static string KindOf(int invokeKind, string? returnType) => invokeKind switch
    {
        2 or 4 or 8 => "Property",
        _ => returnType is null ? "Sub" : "Function",
    };

    private static void AppendVariable(ComHandle<ITypeInfo> info, VarDesc* variable, int typeKind, List<MemberRow> members)
    {
        var names = NamesOf(info, variable->MemberId, 1);
        if (names.Count == 0 || names[0].StartsWith('_'))
        {
            return;
        }

        var name = names[0];

        // An enum member or a module constant carries its value; a record field its type.
        if (variable->VarKind == 2 && variable->Value != 0)
        {
            var value = RenderConstant(variable->Value);
            members.Add(new MemberRow(
                name,
                "Const",
                $"Const {name}{(value is null ? string.Empty : $" = {value}")}",
                DocumentationOf(info, variable->MemberId)));
            return;
        }

        members.Add(new MemberRow(
            name,
            typeKind == TypeKindRecord ? "Field" : "Property",
            $"{name} As {RenderType(info, variable->Variable.Type)}",
            DocumentationOf(info, variable->MemberId)));
    }

    /// <summary>A constant's value, for the integral kinds worth printing.</summary>
    private static string? RenderConstant(nint variantPointer)
    {
        var varType = *(short*)variantPointer;
        var data = variantPointer + 8;

        return varType switch
        {
            VtI2 => (*(short*)data).ToString(System.Globalization.CultureInfo.InvariantCulture),
            VtI4 or VtInt => (*(int*)data).ToString(System.Globalization.CultureInfo.InvariantCulture),
            VtBool => *(short*)data == 0 ? "False" : "True",
            VtBstr => $"\"{System.Runtime.InteropServices.Marshal.PtrToStringBSTR(*(nint*)data)}\"",
            _ => null,
        };
    }

    private const short VtI2 = 2;
    private const short VtI4 = 3;
    private const short VtR4 = 4;
    private const short VtR8 = 5;
    private const short VtCy = 6;
    private const short VtDate = 7;
    private const short VtBstr = 8;
    private const short VtDispatch = 9;
    private const short VtError = 10;
    private const short VtBool = 11;
    private const short VtVariant = 12;
    private const short VtUnknown = 13;
    private const short VtDecimal = 14;
    private const short VtI1 = 16;
    private const short VtUi1 = 17;
    private const short VtUi2 = 18;
    private const short VtUi4 = 19;
    private const short VtI8 = 20;
    private const short VtInt = 22;
    private const short VtUint = 23;
    private const short VtVoid = 24;
    private const short VtHResult = 25;
    private const short VtPtr = 26;
    private const short VtSafeArray = 27;
    private const short VtUserDefined = 29;
    private const short VtLpWStr = 31;

    /// <summary>A TYPEDESC, spelled the way VBA would spell it.</summary>
    private static string RenderType(ComHandle<ITypeInfo> info, in TypeDesc type)
    {
        switch (type.VarType)
        {
            case VtPtr:
                return RenderType(info, in *(TypeDesc*)type.Detail);

            case VtSafeArray:
                return RenderType(info, in *(TypeDesc*)type.Detail) + "()";

            case VtUserDefined:
            {
                if (info.Target.GetRefTypeInfo((int)type.Detail, out var refPointer) >= 0 && refPointer != 0)
                {
                    using var referenced = ComHandle<ITypeInfo>.Own(refPointer);
                    if (referenced is not null
                        && referenced.Target.GetDocumentation(-1, out var namePointer, out var docPointer, out _, out var filePointer) >= 0)
                    {
                        FreeString(docPointer);
                        FreeString(filePointer);
                        var name = TakeString(namePointer);
                        if (name is { Length: > 0 })
                        {
                            return name.TrimStart('_');
                        }
                    }
                }

                return "Object";
            }

            case VtI2:
                return "Integer";
            case VtI4 or VtInt or VtError or VtUi4 or VtUint:
                return "Long";
            case VtR4:
                return "Single";
            case VtR8:
                return "Double";
            case VtCy:
                return "Currency";
            case VtDate:
                return "Date";
            case VtBstr or VtLpWStr:
                return "String";
            case VtDispatch or VtUnknown:
                return "Object";
            case VtBool:
                return "Boolean";
            case VtI1 or VtUi1:
                return "Byte";
            case VtUi2:
                return "Integer";
            case VtI8:
                return "LongLong";
            case VtDecimal:
                return "Variant";
            default:
                return "Variant";
        }
    }

    private static List<string> NamesOf(ComHandle<ITypeInfo> info, int memberId, int expected)
    {
        var names = new List<string>(expected);
        var slots = stackalloc nint[64];
        var wanted = Math.Min(expected, 64);

        if (info.Target.GetNames(memberId, (nint)slots, wanted, out var got) >= 0)
        {
            for (var i = 0; i < got; i++)
            {
                var name = TakeString(slots[i]);
                if (name is not null)
                {
                    names.Add(name);
                }
            }
        }

        return names;
    }

    private static string DocumentationOf(ComHandle<ITypeInfo> info, int memberId)
    {
        if (info.Target.GetDocumentation(memberId, out var namePointer, out var docPointer, out _, out var filePointer) < 0)
        {
            return string.Empty;
        }

        FreeString(namePointer);
        FreeString(filePointer);
        return TakeString(docPointer) ?? string.Empty;
    }

    private static string? TakeName(ITypeLib library, int index)
    {
        if (library.GetDocumentation(index, out var namePointer, out var docPointer, out _, out var filePointer) < 0)
        {
            return null;
        }

        FreeString(docPointer);
        FreeString(filePointer);
        return TakeString(namePointer);
    }

    private static string? TakeString(nint bstr)
    {
        if (bstr == 0)
        {
            return null;
        }

        var text = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(bstr);
        FreeString(bstr);
        return text;
    }

    private static void FreeString(nint bstr)
    {
        if (bstr != 0)
        {
            System.Runtime.InteropServices.Marshal.FreeBSTR(bstr);
        }
    }

    public void Dispose()
    {
        foreach (var held in _libraries)
        {
            held.Library.Dispose();
        }

        _libraries.Clear();
    }
}
