using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// Owns one early-bound interface pointer and the wrapper built over it.
///
/// Two references exist for every wrapped pointer and both have to come back. The raw pointer we
/// were handed is one. The wrapper the marshalling layer builds takes its own, which it would
/// otherwise only give up when the finalizer runs, so it is released explicitly here. Disposing
/// twice, or disposing a handle whose pointer was never ours, is the failure this type exists to
/// make impossible.
/// </summary>
/// <typeparam name="TInterface">A source-generated COM interface the pointer supports.</typeparam>
internal sealed class ComHandle<TInterface> : IDisposable
    where TInterface : class
{
    private nint _pointer;
    private object? _wrapper;

    private ComHandle(nint pointer, object wrapper, TInterface target)
    {
        _pointer = pointer;
        _wrapper = wrapper;
        Target = target;
    }

    /// <summary>The interface to call. Valid until this handle is disposed.</summary>
    public TInterface Target { get; }

    /// <summary>
    /// The raw pointer, borrowed. Callers pass it on to COM but never release it.
    /// </summary>
    public nint Pointer => _pointer;

    /// <summary>
    /// Takes ownership of a pointer that is already counted for us, which is the case for anything
    /// returned from a COM call as an out parameter. Releases the pointer and returns null when the
    /// object does not support the interface.
    /// </summary>
    public static ComHandle<TInterface>? Own(nint pointer)
    {
        if (pointer == 0)
        {
            return null;
        }

        var wrapper = ComRuntime.Wrappers.GetOrCreateObjectForComInstance(pointer, CreateObjectFlags.UniqueInstance);
        if (wrapper is not TInterface target)
        {
            (wrapper as IDisposable)?.Dispose();
            Marshal.Release(pointer);
            return null;
        }

        return new ComHandle<TInterface>(pointer, wrapper, target);
    }

    /// <summary>
    /// Takes its own reference on a pointer the caller keeps, which is the case for anything
    /// received as an argument to an inbound call.
    /// </summary>
    public static ComHandle<TInterface>? Borrow(nint pointer)
    {
        if (pointer == 0)
        {
            return null;
        }

        Marshal.AddRef(pointer);
        return Own(pointer);
    }

    /// <summary>
    /// Asks the object for another interface and wraps the result. The new handle is independent of
    /// this one and outlives it.
    ///
    /// The identifier is passed rather than read from the interface type. Reading it would mean
    /// reflecting over an attribute, and ahead-of-time compilation gives no guarantee that an
    /// attribute nothing else consumes is still present in the image.
    /// </summary>
    public ComHandle<TOther>? As<TOther>(in Guid interfaceId)
        where TOther : class
    {
        var pointer = _pointer;
        if (pointer == 0)
        {
            return null;
        }

        return Marshal.QueryInterface(pointer, in interfaceId, out var other) < 0
            ? null
            : ComHandle<TOther>.Own(other);
    }

    public void Dispose()
    {
        var pointer = Interlocked.Exchange(ref _pointer, 0);
        var wrapper = _wrapper;
        _wrapper = null;

        // The wrapper's reference goes first: releasing ours while the wrapper still holds one is
        // safe, but the reverse ordering reads as if the wrapper could outlive the object.
        (wrapper as IDisposable)?.Dispose();

        if (pointer != 0)
        {
            Marshal.Release(pointer);
        }
    }
}
