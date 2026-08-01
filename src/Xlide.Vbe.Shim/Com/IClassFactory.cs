using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// Standard COM activation interface. COM calls this to create instances of our coclasses after
/// DllGetClassObject hands it a factory.
/// </summary>
[GeneratedComInterface]
[Guid("00000001-0000-0000-C000-000000000046")]
internal partial interface IClassFactory
{
    [PreserveSig]
    int CreateInstance(nint outerUnknown, in Guid riid, out nint instance);

    [PreserveSig]
    int LockServer([MarshalAs(UnmanagedType.Bool)] bool @lock);
}
