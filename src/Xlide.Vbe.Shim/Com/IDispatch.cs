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

/// <summary>
/// EXCEPINFO. Filled in by the callee when a call fails because it raised an error, and the only
/// place the error's own text is available: the HRESULT alone says nothing more than that
/// something went wrong.
///
/// The strings are allocated by the callee and owned by the caller once the call returns.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
internal struct ExcepInfo
{
    public ushort Code;
    public ushort Reserved;
    public nint Source;
    public nint Description;
    public nint HelpFile;
    public uint HelpContext;
    public nint ReservedPointer;
    public nint DeferredFillIn;
    public int ErrorCode;
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

/// <summary>
/// The parts of a dual interface every implementation answers the same way.
///
/// Two objects here are called BY name from outside - the add-in the editor binds to, and the
/// inside door a workbook's VBA reaches - and each carried its own copy of this, including the
/// pointer walk below. Two copies of an unsafe loop is two places for a pointer mistake to
/// live, and the only thing that ever differed between them is which names they know, which
/// is what <paramref name="resolve"/> carries.
/// </summary>
internal static class Dispatch
{
    /// <summary>No type information: neither object ships a type library, and saying so plainly
    /// is what keeps a late-binding host from waiting on one.</summary>
    public static int NoTypeInfoCount(out uint count)
    {
        count = 0;
        return HResult.Ok;
    }

    /// <summary>The other half of the same answer.</summary>
    public static int NoTypeInfo(out nint typeInfo)
    {
        typeInfo = 0;
        return HResult.Fail;
    }

    /// <summary>
    /// GetIDsOfNames over the caller's own name table: every name is resolved, the whole block
    /// is written back, and the reply is Ok only when ALL of them landed - a partial answer is
    /// DISP_E_UNKNOWNNAME with the unknown ones marked, which is what a caller reads to find
    /// out which name it got wrong.
    /// </summary>
    public static unsafe int ResolveNames(
        nint names, uint nameCount, nint dispIds, Func<string?, int> resolve)
    {
        ArgumentNullException.ThrowIfNull(resolve);

        if (names == 0 || dispIds == 0 || nameCount == 0)
        {
            return HResult.InvalidArg;
        }

        var namePointers = (char**)names;
        var results = (int*)dispIds;
        var resolvedAll = true;

        for (var i = 0u; i < nameCount; i++)
        {
            var dispId = resolve(Marshal.PtrToStringUni((nint)namePointers[i]));
            results[i] = dispId;
            if (dispId == DispId.Unknown)
            {
                resolvedAll = false;
            }
        }

        return resolvedAll ? HResult.Ok : HResult.DispUnknownName;
    }
}
