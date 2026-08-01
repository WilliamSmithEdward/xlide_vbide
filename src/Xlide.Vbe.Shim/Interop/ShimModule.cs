using System.Runtime.InteropServices;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>
/// Where this library actually is on disk, and the module handle that owns anything it registers
/// with the window manager.
///
/// The usual base-directory notion answers with the host executable's folder. Inside Excel that is
/// the Office install directory, which contains none of our files. Everything the shim loads at run
/// time - the browser loader, the shell document - sits beside the shim itself, so the module has
/// to locate itself from an address known to be inside it.
/// </summary>
internal static unsafe class ShimModule
{
    private static nint _handle;
    private static string? _directory;
    private static bool _resolved;

    /// <summary>
    /// Module handle of this library. Window classes are registered against it, which ties their
    /// lifetime to a library that is never unloaded.
    /// </summary>
    public static nint Handle
    {
        get
        {
            Resolve();
            return _handle;
        }
    }

    /// <summary>Directory containing this library, or null when it cannot be determined.</summary>
    public static string? Directory
    {
        get
        {
            Resolve();
            return _directory;
        }
    }

    private static void Resolve()
    {
        if (_resolved)
        {
            return;
        }

        _resolved = true;

        // The address has to be one the loader can attribute to this image. A native entry point
        // compiled into the library is such an address; a managed method or a static field is not
        // guaranteed to be.
        var anchor = (nint)(delegate* unmanaged<void>)&ModuleAnchor;

        nint module = 0;
        var flags = Win32.GetModuleHandleFromAddress | Win32.GetModuleHandleUnchangedRefCount;
        if (!Win32.GetModuleHandleEx(flags, anchor, &module) || module == 0)
        {
            return;
        }

        _handle = module;

        const int capacity = 1024;
        var buffer = stackalloc char[capacity];
        var length = Win32.GetModuleFileName(module, buffer, capacity);
        if (length is 0 or >= capacity)
        {
            return;
        }

        _directory = Path.GetDirectoryName(new string(buffer, 0, (int)length));
    }

    /// <summary>
    /// Exists only so its address can be taken. It is never called, by us or by anything else.
    /// </summary>
    [UnmanagedCallersOnly]
    private static void ModuleAnchor()
    {
    }
}
