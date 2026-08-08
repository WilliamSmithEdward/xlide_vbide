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
    private static long _wrappersDisposed;

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

    /// <summary>
    /// How many of those were actually DISPOSED, which is the only thing that releases the
    /// wrapper's own reference. Equal to given-back on a healthy build; any gap means wrappers
    /// were counted home without being released, and the finalizer thread has them.
    /// </summary>
    public static long WrappersDisposed => Interlocked.Read(ref _wrappersDisposed);

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
    public static object? TakeWrapper(
        nint pointer,
        [System.Runtime.CompilerServices.CallerMemberName] string? member = null,
        [System.Runtime.CompilerServices.CallerFilePath] string? file = null,
        [System.Runtime.CompilerServices.CallerLineNumber] int line = 0)
    {
        var wrapper = Wrappers.GetOrCreateObjectForComInstance(
            pointer,
            System.Runtime.InteropServices.CreateObjectFlags.UniqueInstance);

        Interlocked.Increment(ref _wrappersTaken);

#if DEBUG
        // WATCHED, so a wrapper that reaches the finalizer NAMES THE CALL THAT TOOK IT.
        //
        // Five crashes across two days share one stack and none of them says anything about what
        // created the wrapper - by the time the finalizer runs, the call that leaked it is minutes
        // gone. Worse, two of the leaks were invisible to the counter above, which can only see
        // wrappers taken through this door; and the counter reads balanced RIGHT NOW while the
        // host still dies, so something it counts is not actually being released.
        //
        // This closes both gaps. The sentinel is attached to the wrapper, so it becomes garbage at
        // the same moment the wrapper does, and its finalizer writes the file and line that took
        // it. It touches no COM at all, so it cannot itself fault: it only writes a line, and that
        // line is the answer the last five crash reports could not give.
        if (wrapper is not null)
        {
            Watched.Add(wrapper, new WrapperWatch($"{System.IO.Path.GetFileName(file)}:{line} {member}"));
        }
#endif

        return wrapper;
    }

#if DEBUG
    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<object, WrapperWatch> Watched = new();

    /// <summary>
    /// Rides along with one wrapper and reports it if it is collected without being given back.
    ///
    /// Deliberately holds nothing but a string. A finalizer that touched the wrapper would keep it
    /// alive, and one that touched COM would be the very fault this is trying to name.
    /// </summary>
    private sealed class WrapperWatch(string where)
    {
        public bool GivenBack;

        ~WrapperWatch()
        {
            if (!GivenBack)
            {
                Diagnostics.Log.Error(
                    $"COM WRAPPER LEAKED: taken at {where} and finalized without being given back. "
                    + "This is what ends the host.");
            }
        }
    }
#endif

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

#if DEBUG
        if (Watched.TryGetValue(wrapper, out var watch))
        {
            watch.GivenBack = true;
        }
#endif

        /*
         * FinalRelease, NOT `as IDisposable`.
         *
         * THIS IS THE ONE. `(wrapper as IDisposable)?.Dispose()` stood here since the type was
         * written and it never released anything: the wrapper `StrategyBasedComWrappers` hands
         * back is a `ComObject`, and `ComObject` DOES NOT IMPLEMENT IDisposable. The cast failed
         * silently, `?.` swallowed it, and the line read as a disposal to everyone who looked at
         * it - including three separate hunts for this exact crash.
         *
         * So every wrapper this product has ever taken kept its reference and waited for the
         * FINALIZER thread, which releases it in a context where an access violation cannot be
         * thrown: the runtime FailFasts and Excel goes with it. That is the stack behind five
         * crashes across 2026-08-07 and 2026-08-08, and behind the ones before them that were
         * filed against ntdll and VBE7. The two earlier fixes were real - they stopped wrappers
         * being taken that nobody needed - but they only lowered the rate, because the release
         * itself had never worked.
         *
         * It was invisible to the counter for the oldest reason there is: the counter incremented
         * whether or not the disposal did anything, so 1,882 wrappers taken and 1,882 given back
         * read as perfectly balanced with 1,882 of them pending finalization. Counting the
         * INTENTION rather than the EFFECT. Found 2026-08-08 by counting them separately and
         * reading `disposed: 0`.
         *
         * The counters stay separate for that reason. They must always agree, and the day they do
         * not is the day this is broken again.
         */
        var released = false;

        if (wrapper is IDisposable disposable)
        {
            disposable.Dispose();
            released = true;
        }
        else if (wrapper is ComObject comObject)
        {
            comObject.FinalRelease();
            released = true;
        }

        if (released)
        {
            Interlocked.Increment(ref _wrappersDisposed);
        }
        else
        {
            // Loud, because the alternative is silence and a dead host later. There is no third
            // kind of wrapper today; if one appears, this says so before it kills anything.
            Diagnostics.Log.Error(
                $"COM WRAPPER NOT RELEASED: {wrapper.GetType().FullName} is neither IDisposable "
                + "nor a ComObject, so its reference is going to the finalizer thread.");
        }

        Interlocked.Increment(ref _wrappersGivenBack);
    }
}
