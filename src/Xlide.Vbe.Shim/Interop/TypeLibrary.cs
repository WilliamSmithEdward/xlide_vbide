using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Interop;

/*
 * Just enough of the type-library machinery to browse what a project references.
 *
 * A workbook's project references real type libraries — Excel's own, the VBA runtime,
 * Office, stdole — and those libraries are the Object Browser's subject matter. The
 * libraries describe themselves through ITypeLib and ITypeInfo, whose data lives in
 * caller-released structures reached through raw pointers.
 *
 * Every interface here is declared flat, with a placeholder for each method that is not
 * used, so the slots line up with the real vtable. The counts in the comments are what
 * keeps that honest.
 */

/// <summary>ITypeLib. Enumeration and names; the attributes travel as raw structs.</summary>
[GeneratedComInterface]
[Guid("00020402-0000-0000-C000-000000000046")]
internal partial interface ITypeLib
{
    // 1: returns the count itself, not an HRESULT.
    [PreserveSig] int GetTypeInfoCount();

    // 2..4
    [PreserveSig] int GetTypeInfo(int index, out nint typeInfo);
    [PreserveSig] int GetTypeInfoType(int index, out int typeKind);
    [PreserveSig] int GetTypeInfoOfGuid(in Guid guid, out nint typeInfo);

    // 5..6
    [PreserveSig] int GetLibAttr(out nint attributes);
    [PreserveSig] int GetTypeComp(out nint typeComp);

    // 7: index -1 names the library itself.
    [PreserveSig] int GetDocumentation(int index, out nint name, out nint documentation, out int helpContext, out nint helpFile);

    // 8..9
    [PreserveSig] int IsName(nint nameBuffer, int hash, out int found);
    [PreserveSig] int FindName(nint nameBuffer, int hash, nint typeInfos, nint memberIds, nint found);

    // 10: returns nothing.
    [PreserveSig] void ReleaseTLibAttr(nint attributes);
}

/// <summary>ITypeInfo. Attributes, functions, variables, names, and referenced types.</summary>
[GeneratedComInterface]
[Guid("00020401-0000-0000-C000-000000000046")]
internal partial interface ITypeInfo
{
    // 1..2
    [PreserveSig] int GetTypeAttr(out nint attributes);
    [PreserveSig] int GetTypeComp(out nint typeComp);

    // 3..4
    [PreserveSig] int GetFuncDesc(int index, out nint funcDesc);
    [PreserveSig] int GetVarDesc(int index, out nint varDesc);

    // 5: nameBuffer is an array of BSTR slots the caller owns.
    [PreserveSig] int GetNames(int memberId, nint nameBuffer, int maxNames, out int nameCount);

    // 6..7
    [PreserveSig] int GetRefTypeOfImplType(int index, out int refType);
    [PreserveSig] int GetImplTypeFlags(int index, out int flags);

    // 8..9
    [PreserveSig] int GetIDsOfNames(nint names, int count, nint memberIds);
    [PreserveSig] int Invoke(nint instance, int memberId, ushort flags, nint parameters, nint result, nint exceptions, nint argumentError);

    // 10
    [PreserveSig] int GetDocumentation(int memberId, out nint name, out nint documentation, out int helpContext, out nint helpFile);

    // 11
    [PreserveSig] int GetDllEntry(int memberId, int invokeKind, nint dllName, nint entryName, nint ordinal);

    // 12
    [PreserveSig] int GetRefTypeInfo(int refType, out nint typeInfo);

    // 13..16
    [PreserveSig] int AddressOfMember(int memberId, int invokeKind, out nint address);
    [PreserveSig] int CreateInstance(nint outer, in Guid interfaceId, out nint instance);
    [PreserveSig] int GetMops(int memberId, out nint mops);
    [PreserveSig] int GetContainingTypeLib(out nint typeLib, out int index);

    // 17..19: return nothing.
    [PreserveSig] void ReleaseTypeAttr(nint attributes);
    [PreserveSig] void ReleaseFuncDesc(nint funcDesc);
    [PreserveSig] void ReleaseVarDesc(nint varDesc);
}

/// <summary>TYPEDESC: a type, as a variant tag plus a pointer whose meaning the tag picks.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct TypeDesc
{
    /// <summary>Points at another TypeDesc for VT_PTR and VT_SAFEARRAY; holds the reference
    /// handle for VT_USERDEFINED.</summary>
    public nint Detail;
    public short VarType;
}

/// <summary>PARAMDESC: a parameter's flags, and its default when it declares one.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct ParamDesc
{
    public nint DefaultValue;
    public short Flags;
}

/// <summary>ELEMDESC: one element — a parameter or a return — as its type plus flags.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct ElemDesc
{
    public TypeDesc Type;
    public ParamDesc Parameter;
}

/// <summary>FUNCDESC, the parts the browser reads.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct FuncDesc
{
    public int MemberId;
    public nint Codes;
    public nint Parameters;
    public int FuncKind;
    public int InvokeKind;
    public int CallConv;
    public short ParameterCount;
    public short OptionalParameterCount;
    public short VtableOffset;
    public short CodeCount;
    public ElemDesc Return;
    public short Flags;
}

/// <summary>VARDESC, the parts the browser reads: enum members and module constants.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct VarDesc
{
    public int MemberId;
    public nint Schema;
    /// <summary>A VARIANT* for constants; an instance offset otherwise.</summary>
    public nint Value;
    public ElemDesc Variable;
    public short Flags;
    public int VarKind;
}

/// <summary>TYPEATTR, the parts the browser reads.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct TypeAttr
{
    public Guid Guid;
    public int Lcid;
    public int Reserved;
    public int ConstructorId;
    public int DestructorId;
    public nint Schema;
    public int InstanceSize;
    public int TypeKind;
    public short FuncCount;
    public short VarCount;
    public short ImplTypeCount;
    public short VtableSize;
    public short Alignment;
    public short TypeFlags;
    public short MajorVersion;
    public short MinorVersion;
    public TypeDesc Alias;
    public nint IdlReserved;
    public short IdlFlags;
}

internal static partial class TypeLibraryNative
{
    /// <summary>REGKIND_NONE: load without touching the registry.</summary>
    public const int LoadWithoutRegistering = 2;

    [LibraryImport("oleaut32.dll", StringMarshalling = StringMarshalling.Utf16)]
    public static partial int LoadTypeLibEx(string path, int registerKind, out nint typeLib);
}
