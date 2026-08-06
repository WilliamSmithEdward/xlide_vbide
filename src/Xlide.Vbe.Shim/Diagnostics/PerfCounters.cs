#if DEBUG
namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// Cheap counters behind the stats route, measuring the regression classes this product
/// has actually lived through: placement-pass storms (sixty full passes per resize drag,
/// 2026-08-05), host-thread stalls (the marshal every api request already performs is the
/// probe), log-write storms ("resizing is slippery", 2026-08-04), and poll tiers left
/// running. Interlocked all the way; nothing here takes a lock or throws. Debug only.
/// </summary>
internal static class PerfCounters
{
    public static readonly long StartedAt = Environment.TickCount64;

    private static long _placementFullPasses;
    private static long _placementFastPasses;
    private static long _placementLastMs;
    private static long _placementMaxMs;
    private static long _marshalCount;
    private static long _marshalLastMs;
    private static long _marshalMaxMs;
    private static long _logLines;
    private static long _pollIntervalMs;

    public static void PlacementFull(long milliseconds)
    {
        Interlocked.Increment(ref _placementFullPasses);
        Interlocked.Exchange(ref _placementLastMs, milliseconds);
        RaiseToAtLeast(ref _placementMaxMs, milliseconds);
    }

    public static void PlacementFast() => Interlocked.Increment(ref _placementFastPasses);

    public static void Marshal(long milliseconds)
    {
        Interlocked.Increment(ref _marshalCount);
        Interlocked.Exchange(ref _marshalLastMs, milliseconds);
        RaiseToAtLeast(ref _marshalMaxMs, milliseconds);
    }

    public static void LogLine() => Interlocked.Increment(ref _logLines);

    public static void Poll(long intervalMilliseconds) => Interlocked.Exchange(ref _pollIntervalMs, intervalMilliseconds);

    public static (long FullPasses, long FastPasses, long LastMs, long MaxMs) PlacementSnapshot() =>
        (Interlocked.Read(ref _placementFullPasses), Interlocked.Read(ref _placementFastPasses),
            Interlocked.Read(ref _placementLastMs), Interlocked.Read(ref _placementMaxMs));

    public static (long Count, long LastMs, long MaxMs) MarshalSnapshot() =>
        (Interlocked.Read(ref _marshalCount), Interlocked.Read(ref _marshalLastMs), Interlocked.Read(ref _marshalMaxMs));

    public static long LogLineCount => Interlocked.Read(ref _logLines);

    public static long PollIntervalMs => Interlocked.Read(ref _pollIntervalMs);

    private static void RaiseToAtLeast(ref long slot, long value)
    {
        long seen;
        while (value > (seen = Interlocked.Read(ref slot))
            && Interlocked.CompareExchange(ref slot, value, seen) != seen)
        {
        }
    }
}
#endif
