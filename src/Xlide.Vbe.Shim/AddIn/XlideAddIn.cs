using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Interop;

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
    /// The host pointers, retained past OnConnection so a shutdown the developer cancels can be
    /// recovered from: OnBeginShutdown has already stopped the session by the time the save
    /// prompt appears, and standing it up again needs what OnConnection was given.
    /// </summary>
    private nint _application;
    private nint _addInInstance;

    /// <summary>Alive between OnBeginShutdown and whatever the shutdown turns out to be.</summary>
    private ShutdownWatchdog? _watchdog;
    private int _watchdogTicks;
    private int _watchdogEnabledTicks;

    /// <summary>
    /// Releases the session if the host never called OnDisconnection. COM controls this object's
    /// lifetime, so the normal path is OnDisconnection and this is the safety net.
    /// </summary>
    public void Dispose()
    {
        _watchdog?.Dispose();
        _watchdog = null;

        var session = _session;
        _session = null;
        session?.Dispose();

        ReleaseHostPointers();
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

            // Kept for a revival, with references of their own.
            _application = application;
            Marshal.AddRef(application);
            if (addInInstance != 0)
            {
                _addInInstance = addInInstance;
                Marshal.AddRef(addInInstance);
            }

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

            // But the shutdown is not a done deal: the host asks about unsaved changes AFTER
            // this, and Cancel abandons the whole thing with no callback to say so. The watchdog
            // is the one thing left listening; its ticks only arrive if the host thread is still
            // pumping, which is itself the news.
            _watchdogTicks = 0;
            _watchdogEnabledTicks = 0;
            _watchdog ??= ShutdownWatchdog.Create(OnWatchdogTick);
            return HResult.Ok;
        }
        catch (Exception ex)
        {
            Log.Error("OnBeginShutdown failed", ex);
            return HResult.Ok;
        }
    }

    /// <summary>
    /// A tick after OnBeginShutdown, on a host thread that is evidently still pumping. If the
    /// editor's frame is standing, the developer cancelled the shutdown, and the session is
    /// stood up again from the retained pointers, exactly as OnConnection stood it up. A real
    /// shutdown never gets here: the thread stops pumping, and OnDisconnection retires the
    /// watchdog anyway.
    /// </summary>
    private void OnWatchdogTick()
    {
        try
        {
            if (_application == 0)
            {
                RetireWatchdog();
                return;
            }

            var frame = CodePaneTracker.FindFrame();
            if (frame == 0)
            {
                // No editor frame yet: either the teardown is mid-flight after all, or the
                // frame is briefly gone. A few patient ticks tell the difference.
                _watchdogEnabledTicks = 0;
                if (++_watchdogTicks >= 8)
                {
                    Log.Info("watchdog: no editor frame returned, standing down");
                    RetireWatchdog();
                }

                return;
            }

            // A standing frame is not yet the answer. The host asks about unsaved changes with
            // an app-modal dialog AFTER OnBeginShutdown, and a modal loop pumps timers, so a
            // tick during the dialog proves only that the developer is still deciding. Reviving
            // there painted the surface over an undecided shutdown, and when the developer then
            // chose Save, the real teardown ripped through a seconds-old session mid-start and
            // took the host down with it. What the dialog does do is DISABLE the frame, the way
            // every app-modal dialog disables its application's windows; an enabled frame held
            // across consecutive ticks is the cancellation. The wait costs nothing — while the
            // dialog is up the developer is looking at the dialog — and it does not spend the
            // patience budget above, because a dialog can sit unanswered for minutes.
            if (!Win32.IsWindowEnabled(frame))
            {
                _watchdogEnabledTicks = 0;
                return;
            }

            if (++_watchdogEnabledTicks < 2)
            {
                return;
            }

            Log.Info("watchdog: the shutdown was cancelled, reviving the session");
            RetireWatchdog();

            var stopped = _session;
            _session = null;
            stopped?.Dispose();

            var editor = DispatchObject.AttachBorrowed(_application);
            if (editor is null)
            {
                Log.Error("watchdog: the retained editor pointer no longer answers");
                return;
            }

            var addIn = _addInInstance != 0 ? DispatchObject.AttachBorrowed(_addInInstance) : null;

            _session = new AddInSession(editor, addIn);
            _session.Start();
            _session.HostStartupComplete();
        }
        catch (Exception ex)
        {
            Log.Error("watchdog: revival failed", ex);
            RetireWatchdog();
        }
    }

    private void RetireWatchdog()
    {
        _watchdog?.Dispose();
        _watchdog = null;
    }

    public int OnDisconnection(ExtDisconnectMode disconnectMode, nint custom)
    {
        try
        {
            Log.Info($"OnDisconnection, mode {disconnectMode}");

            RetireWatchdog();

            var session = _session;
            _session = null;

            session?.Stop();
            session?.Dispose();

            ReleaseHostPointers();

            Log.Info("disconnected cleanly");
            return HResult.Ok;
        }
        catch (Exception ex)
        {
            Log.Error("OnDisconnection failed", ex);
            return HResult.Ok;
        }
    }

    private void ReleaseHostPointers()
    {
        if (_application != 0)
        {
            Marshal.Release(_application);
            _application = 0;
        }

        if (_addInInstance != 0)
        {
            Marshal.Release(_addInInstance);
            _addInInstance = 0;
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
