using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Xlide.Vbe.Core;
using Xlide.Vbe.Shim.AddIn;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// The functions COM looks for by name when it loads this library. Compiling ahead of time means
/// these are real native exports and no runtime has to start before they can be called.
/// </summary>
public static unsafe class Exports
{
    private static readonly Guid AddInClsid = new(ProductIdentity.AddInClsid);

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

            Func<object>? create = *classId == AddInClsid
                ? static () => new XlideAddIn()
                : null;

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

#if DEBUG
    /// <summary>
    /// Runs <see cref="SelfCheck"/> in the calling process: the named stages, comma separated, or
    /// all of them for null or empty, each finding appended to the file at
    /// <paramref name="reportPath"/>. Answers the number of checks that failed, or minus one
    /// when it could not run at all. The caller's thread must be a single-threaded apartment,
    /// as the editor's is.
    ///
    /// Debug builds only: a shipped library has no such export.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "XlideSelfCheck")]
    public static int XlideSelfCheck(char* reportPath, char* stages)
    {
        if (reportPath is null)
        {
            return -1;
        }

        try
        {
            Log.Initialize();
            return SelfCheck.Run(new string(reportPath), stages is null ? string.Empty : new string(stages));
        }
        catch (Exception ex)
        {
            Log.Error("the self-check could not run", ex);
            return -1;
        }
    }
#endif
}
