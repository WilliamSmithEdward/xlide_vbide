using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>Automation dispatch interface, used to call the editor object model by name.</summary>
[GeneratedComInterface]
[Guid("00020400-0000-0000-C000-000000000046")]
internal partial interface IDispatch
{
    [PreserveSig]
    int GetTypeInfoCount(out uint count);

    [PreserveSig]
    int GetTypeInfo(uint typeInfoIndex, uint lcid, out nint typeInfo);

    [PreserveSig]
    int GetIDsOfNames(in Guid riid, nint names, uint nameCount, uint lcid, nint dispIds);

    [PreserveSig]
    int Invoke(
        int dispIdMember,
        in Guid riid,
        uint lcid,
        ushort flags,
        nint dispParams,
        nint result,
        nint exceptionInfo,
        nint argumentError);
}

/// <summary>Argument block passed to <see cref="IDispatch.Invoke"/>.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct DispatchParameters
{
    public nint Arguments;
    public nint NamedArguments;
    public uint ArgumentCount;
    public uint NamedArgumentCount;
}

/// <summary>Invocation kinds accepted by <see cref="IDispatch.Invoke"/>.</summary>
[Flags]
internal enum InvokeKind : ushort
{
    Method = 1,
    PropertyGet = 2,
    PropertyPut = 4,
    PropertyPutRef = 8,
}

/// <summary>Well-known dispatch identifiers.</summary>
internal static class DispId
{
    public const int Value = 0;
    public const int Unknown = -1;
    public const int PropertyPut = -3;
}
