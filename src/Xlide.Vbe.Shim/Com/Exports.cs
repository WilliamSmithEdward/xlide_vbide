using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Xlide.Vbe.Core;
using Xlide.Vbe.Shim.AddIn;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.ToolWindow;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// The functions COM looks for by name when it loads this library. Compiling ahead of time means
/// these are real native exports and no runtime has to start before they can be called.
/// </summary>
public static unsafe class Exports
{
    private static readonly Guid AddInClsid = new(ProductIdentity.AddInClsid);
    private static readonly Guid ToolWindowHostClsid = new(ProductIdentity.ToolWindowHostClsid);

    /// <summary>Hands COM a factory for one of our coclasses.</summary>
    [UnmanagedCallersOnly(EntryPoint = "DllGetClassObject")]
    public static int DllGetClassObject(Guid* classId, Guid* interfaceId, nint* factory)
    {
        if (factory is null)
        {
            return HResult.InvalidArg;
        }

        *factory = 0;

        if (classId is null || interfaceId is null)
        {
            return HResult.InvalidArg;
        }

        try
        {
            // Logged unconditionally. When a host declines to load the add-in, the first question
            // is always whether it got as far as asking for the class, and this is the only place
            // that can answer it.
            Log.Initialize();
            Log.Info($"DllGetClassObject for {*classId:B}");

            Func<object>? create = null;

            if (*classId == AddInClsid)
            {
                create = static () => new XlideAddIn();
            }
            else if (*classId == ToolWindowHostClsid)
            {
                create = static () => new ToolWindowHost();
            }

            if (create is null)
            {
                Log.Warn($"no class registered for {*classId:B}");
                return HResult.ClassNotAvailable;
            }

            var instance = new ClassFactory(create);
            var unknown = ComRuntime.Wrappers.GetOrCreateComInterfaceForObject(instance, CreateComInterfaceFlags.None);

            try
            {
                return Marshal.QueryInterface(unknown, in Unsafe.AsRef<Guid>(interfaceId), out *factory);
            }
            finally
            {
                Marshal.Release(unknown);
            }
        }
        catch (Exception ex)
        {
            Log.Initialize();
            Log.Error("DllGetClassObject failed", ex);
            return HResult.Fail;
        }
    }

    /// <summary>
    /// Reports whether the library may be unloaded. It answers no for the life of the process.
    ///
    /// Unloading an add-in library out from under a host that still holds window procedures, hook
    /// callbacks, and event sinks is a class of crash with no upside. The library is small and the
    /// host is going to exit soon enough.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "DllCanUnloadNow")]
    public static int DllCanUnloadNow() => HResult.False;
}
