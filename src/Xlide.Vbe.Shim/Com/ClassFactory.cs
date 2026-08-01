using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// Creates instances of one coclass on demand. COM obtains this from DllGetClassObject and then
/// asks it for the object itself.
/// </summary>
[GeneratedComClass]
internal sealed partial class ClassFactory : IClassFactory
{
    private readonly Func<object> _create;

    public ClassFactory(Func<object> create) => _create = create;

    public int CreateInstance(nint outerUnknown, in Guid riid, out nint instance)
    {
        instance = 0;

        // Aggregation would require this object to delegate its identity to an outer object. The
        // editor never asks for it, and supporting it silently would be worse than refusing.
        if (outerUnknown != 0)
        {
            return HResult.NoAggregation;
        }

        try
        {
            var managed = _create();
            var unknown = ComRuntime.Wrappers.GetOrCreateComInterfaceForObject(managed, CreateComInterfaceFlags.None);

            try
            {
                var hr = Marshal.QueryInterface(unknown, in riid, out instance);

                // The interface the host asks for, and whether we can supply it, decides whether
                // the add-in is ever connected. A refusal here is silent from the outside.
                Log.Info($"CreateInstance for {riid:B} returned 0x{hr:X8}");
                return hr;
            }
            finally
            {
                // QueryInterface took its own reference on success, so release ours either way.
                Marshal.Release(unknown);
            }
        }
        catch (Exception ex)
        {
            Log.Error("class activation failed", ex);
            return HResult.Fail;
        }
    }

    public int LockServer(bool @lock) => HResult.Ok;
}
