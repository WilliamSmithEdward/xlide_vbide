using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// The single wrapper table for the whole shim.
///
/// This is deliberately one instance and not one per call site. A ComWrappers instance caches the
/// mapping between a managed object and its COM identity, so two instances hand out two different
/// unknowns for the same object. COM defines identity as pointer equality on IUnknown, and a
/// container that receives our control through one path and compares it against the same control
/// received through another would decide they are different objects. That failure is silent and
/// arbitrarily delayed, which is the worst combination.
/// </summary>
internal static class ComRuntime
{
    /// <summary>Wrapper table used for every managed-to-COM and COM-to-managed transition here.</summary>
    public static readonly StrategyBasedComWrappers Wrappers = new();

    private static long _wrappersTaken;
    private static long _wrappersGivenBack;

    /// <summary>
    /// Wrappers this shim has built over the editor's objects, and given back.
    ///
    /// COUNTED BECAUSE THE ALTERNATIVE IS A CRASH REPORT. A wrapper holds its own reference on an
    /// apartment-threaded object, and one that is never disposed is given back by the finalizer
    /// thread instead, where releasing it is an access violation that ahead-of-time compilation
    /// cannot throw and so ends the host process. The damage also surfaces late and somewhere
    /// else: on 2026-08-07 the same leak was reported once against this library, once against
    /// VBE7.DLL, and twice as heap corruption inside ntdll.
    ///
    /// So the leak is made visible while it is still only a leak. Live should return to its
    /// resting level after any operation; a figure that only ever climbs is the crash, early.
    /// Reported by the `stats` route.
    /// </summary>
    public static long WrappersTaken => Interlocked.Read(ref _wrappersTaken);

    public static long WrappersGivenBack => Interlocked.Read(ref _wrappersGivenBack);

    public static long WrappersLive => WrappersTaken - WrappersGivenBack;

    /// <summary>
    /// Builds the wrapper and counts it, in one call, because the two must not be separable.
    ///
    /// THE FIRST VERSION OF THIS COUNTER COUNTED THE WRONG THING. It was incremented beside the
    /// disposal rather than by it, so a build with the disposal deliberately removed still counted
    /// balanced and the suite passed. An instrument that agrees with the code and disagrees with
    /// the product is worse than no instrument, because it is believed.
    ///
    /// Taking and giving back are therefore the only two doors, and each does its own counting.
    /// A caller cannot dispose without counting, or count without disposing.
    /// </summary>
    public static object? TakeWrapper(nint pointer)
    {
        var wrapper = Wrappers.GetOrCreateObjectForComInstance(
            pointer,
            System.Runtime.InteropServices.CreateObjectFlags.UniqueInstance);

        Interlocked.Increment(ref _wrappersTaken);
        return wrapper;
    }

    /// <summary>
    /// Gives a wrapper's own reference back, on the caller's thread, and counts it.
    ///
    /// The thread is the whole point. A `UniqueInstance` wrapper is not cached, so nothing else
    /// will ever release it; left alone it is released by the FINALIZER THREAD, and the editor's
    /// objects are apartment-threaded, where that is an access violation the runtime cannot throw
    /// and so ends the host process, minutes later, in a stack that names nothing useful.
    /// </summary>
    public static void GiveBackWrapper(object? wrapper)
    {
        if (wrapper is null)
        {
            return;
        }

        (wrapper as IDisposable)?.Dispose();
        Interlocked.Increment(ref _wrappersGivenBack);
    }
}
