using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// The object the Visual Basic Editor creates and calls to start and stop the add-in.
///
/// Everything here runs on the host user interface thread. Work that is not immediate is handed to
/// the session, which owns the lifetime of hooks, windows, and the engine connection.
/// </summary>
/// <remarks>
/// The class implements IDispatch as well as the extensibility interface. The editor treats an
/// add-in as an automation object and asks for both; source-generated interop supplies only what a
/// class declares, so an object that offers just the extensibility interface is refused with no
/// diagnostic beyond the add-in silently staying disconnected. The two interfaces declare the same
/// four dispatch methods, so one implementation satisfies both.
/// </remarks>
[GeneratedComClass]
internal sealed partial class XlideAddIn : IDTExtensibility2, IDispatch, IDisposable
{
    private AddInSession? _session;

    /// <summary>
    /// Releases the session if the host never called OnDisconnection. COM controls this object's
    /// lifetime, so the normal path is OnDisconnection and this is the safety net.
    /// </summary>
    public void Dispose()
    {
        var session = _session;
        _session = null;
        session?.Dispose();
    }

    public int OnConnection(nint application, ExtConnectMode connectMode, nint addInInstance, nint custom)
    {
        try
        {
            Log.Initialize();
            Log.Info($"OnConnection, mode {connectMode}");

            // The pointers belong to the caller for the duration of the call, so take our own
            // references rather than assuming they outlive it.
            var editor = DispatchObject.AttachBorrowed(application);
            if (editor is null)
            {
                Log.Error("OnConnection received an editor pointer that is not an automation object");
                return HResult.InvalidArg;
            }

            var addIn = DispatchObject.AttachBorrowed(addInInstance);

            _session = new AddInSession(editor, addIn);
            _session.Start();

            return HResult.Ok;
        }
        catch (Exception ex)
        {
            Log.Error("OnConnection failed", ex);

            // Returning a failure would make the editor disable the add-in for the user. Report
            // success and stay inert instead, so the log is reachable and the host is unaffected.
            return HResult.Ok;
        }
    }

    public int OnStartupComplete(nint custom)
    {
        try
        {
            Log.Info("OnStartupComplete");
            _session?.HostStartupComplete();
            return HResult.Ok;
        }
        catch (Exception ex)
        {
            Log.Error("OnStartupComplete failed", ex);
            return HResult.Ok;
        }
    }

    public int OnAddInsUpdate(nint custom) => HResult.Ok;

    public int OnBeginShutdown(nint custom)
    {
        try
        {
            Log.Info("OnBeginShutdown");

            // Every hook, subclass, and event sink must be gone before the host starts tearing
            // itself down. This is the single most important line in the file.
            _session?.Stop();
            return HResult.Ok;
        }
        catch (Exception ex)
        {
            Log.Error("OnBeginShutdown failed", ex);
            return HResult.Ok;
        }
    }

    public int OnDisconnection(ExtDisconnectMode disconnectMode, nint custom)
    {
        try
        {
            Log.Info($"OnDisconnection, mode {disconnectMode}");

            var session = _session;
            _session = null;

            session?.Stop();
            session?.Dispose();

            Log.Info("disconnected cleanly");
            return HResult.Ok;
        }
        catch (Exception ex)
        {
            Log.Error("OnDisconnection failed", ex);
            return HResult.Ok;
        }
    }

    // The remaining members complete the dual interface. The editor binds through the vtable, but
    // a host is free to late-bind, and an incomplete dual interface fails obscurely when it does.

    public int GetTypeInfoCount(out uint count)
    {
        count = 0;
        return HResult.Ok;
    }

    public int GetTypeInfo(uint typeInfoIndex, uint lcid, out nint typeInfo)
    {
        typeInfo = 0;
        return HResult.Fail;
    }

    public unsafe int GetIDsOfNames(in Guid riid, nint names, uint nameCount, uint lcid, nint dispIds)
    {
        if (names == 0 || dispIds == 0 || nameCount == 0)
        {
            return HResult.InvalidArg;
        }

        var namePointers = (char**)names;
        var results = (int*)dispIds;
        var resolvedAll = true;

        for (var i = 0u; i < nameCount; i++)
        {
            var name = Marshal.PtrToStringUni((nint)namePointers[i]);
            var dispId = DispIdForName(name);
            results[i] = dispId;

            if (dispId == DispId.Unknown)
            {
                resolvedAll = false;
            }
        }

        return resolvedAll ? HResult.Ok : HResult.DispUnknownName;
    }

    public unsafe int Invoke(
        int dispIdMember,
        in Guid riid,
        uint lcid,
        ushort flags,
        nint dispParams,
        nint result,
        nint exceptionInfo,
        nint argumentError)
    {
        if (dispParams == 0)
        {
            return HResult.InvalidArg;
        }

        var parameters = (DispatchParameters*)dispParams;
        var arguments = (ComVariant*)parameters->Arguments;
        var count = parameters->ArgumentCount;

        // Automation passes arguments in reverse order.
        ComVariant* Argument(uint indexFromLeft, uint total) =>
            arguments is null || indexFromLeft >= total ? null : &arguments[total - 1 - indexFromLeft];

        static nint InterfaceOf(ComVariant* variant) =>
            variant is null || variant->VarType is not (VarEnum.VT_DISPATCH or VarEnum.VT_UNKNOWN)
                ? 0
                : variant->GetRawDataRef<nint>();

        static int Int32Of(ComVariant* variant) =>
            variant is null || variant->VarType != VarEnum.VT_I4 ? 0 : variant->GetRawDataRef<int>();

        switch (dispIdMember)
        {
            case ExtensibilityDispId.OnConnection when count >= 3:
                return OnConnection(
                    InterfaceOf(Argument(0, count)),
                    (ExtConnectMode)Int32Of(Argument(1, count)),
                    InterfaceOf(Argument(2, count)),
                    0);

            case ExtensibilityDispId.OnDisconnection when count >= 1:
                return OnDisconnection((ExtDisconnectMode)Int32Of(Argument(0, count)), 0);

            case ExtensibilityDispId.OnAddInsUpdate:
                return OnAddInsUpdate(0);

            case ExtensibilityDispId.OnStartupComplete:
                return OnStartupComplete(0);

            case ExtensibilityDispId.OnBeginShutdown:
                return OnBeginShutdown(0);

            default:
                return HResult.DispMemberNotFound;
        }
    }

    private static int DispIdForName(string? name) => name switch
    {
        "OnConnection" => ExtensibilityDispId.OnConnection,
        "OnDisconnection" => ExtensibilityDispId.OnDisconnection,
        "OnAddInsUpdate" => ExtensibilityDispId.OnAddInsUpdate,
        "OnStartupComplete" => ExtensibilityDispId.OnStartupComplete,
        "OnBeginShutdown" => ExtensibilityDispId.OnBeginShutdown,
        _ => DispId.Unknown,
    };
}

/// <summary>Dispatch identifiers defined by the extensibility interface.</summary>
internal static class ExtensibilityDispId
{
    public const int OnConnection = 1;
    public const int OnDisconnection = 2;
    public const int OnAddInsUpdate = 3;
    public const int OnStartupComplete = 4;
    public const int OnBeginShutdown = 5;
}
