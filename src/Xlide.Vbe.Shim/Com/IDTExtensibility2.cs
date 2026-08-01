using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>How the host connected the add-in.</summary>
internal enum ExtConnectMode
{
    AfterStartup = 0,
    Startup = 1,
    External = 2,
    CommandLine = 3,
}

/// <summary>Why the host disconnected the add-in.</summary>
internal enum ExtDisconnectMode
{
    HostShutdown = 0,
    UserClosed = 1,
}

/// <summary>
/// The interface the Visual Basic Editor calls to drive an add-in's lifetime.
///
/// This is a dual interface, so its vtable is IUnknown, then the four IDispatch slots, then the
/// five extensibility methods. The declaration order below reproduces that layout exactly and
/// must not be reordered: source-generated COM assigns vtable slots in declaration order, and the
/// host binds through the vtable.
///
/// The IDispatch slots are declared so the layout is correct. They are also implemented, because a
/// host is free to late-bind instead, and a half-built dual interface fails in a way that is hard
/// to diagnose from inside another process.
/// </summary>
[GeneratedComInterface]
[Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
internal partial interface IDTExtensibility2
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

    [PreserveSig]
    int OnConnection(nint application, ExtConnectMode connectMode, nint addInInstance, nint custom);

    [PreserveSig]
    int OnDisconnection(ExtDisconnectMode disconnectMode, nint custom);

    [PreserveSig]
    int OnAddInsUpdate(nint custom);

    [PreserveSig]
    int OnStartupComplete(nint custom);

    [PreserveSig]
    int OnBeginShutdown(nint custom);
}
