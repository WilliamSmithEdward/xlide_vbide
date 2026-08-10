#if DEBUG
using System.Collections.Concurrent;
using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// What the analyzer costs, per method, split into QUEUED and CALLED.
///
/// Every language feature this product has goes down one pipe - completions, hover, signature
/// help, diagnostics, navigation, rename, semantic tokens, outline - and that pipe serves ONE
/// request at a time behind a semaphore. So a method's total latency is two different things
/// added together, and only one of them is the analyzer's fault: the wait to get on the pipe,
/// and the round trip once on it. A diagnostics pass over a large module delays every keystroke's
/// completion request behind it, and a total-time figure reports that as slow completions.
///
/// Nothing measured this before today. It is the largest perf blind spot in the product: the
/// features a developer feels most are the ones with no instrument at all.
///
/// Per method rather than in aggregate, because the methods have nothing in common - hover is
/// called on a mouse move and diagnostics on a debounce, and averaging them describes neither.
/// Debug only, allocation-light, never throws.
/// </summary>
internal static class EngineCounters
{
    /// <summary>Kept small: the tail is what a perf hunt reads, and the aggregate is per method.</summary>
    private const int SlowestKept = 24;

    /// <summary>
    /// Above this, a call is worth keeping individually rather than only counting.
    ///
    /// Sixty, which is roughly four frames: past that a keystroke's answer is late enough to
    /// feel late. A first pass at 120 kept nothing at all through a load test whose p95 was 107
    /// and whose completions were visibly waiting, which is a threshold measuring its own
    /// optimism (2026-08-07).
    /// </summary>
    private const long SlowMs = 60;

    private sealed class MethodTally
    {
        public long Calls;
        public long WaitTotalMs;
        public long CallTotalMs;
        public long WaitMaxMs;
        public long CallMaxMs;
        public long Refused;

        /// <summary>The last durations, raw, so a median and a p95 are computable. See PerfCounters.</summary>
        public readonly long[] Recent = new long[32];
        public int Cursor;
    }

    private static readonly ConcurrentDictionary<string, MethodTally> Tallies = new(StringComparer.Ordinal);
    private static readonly Lock Gate = new();
    private static readonly List<EngineSlowCall> Slowest = [];
    private static long _started = Environment.TickCount64;

    /// <summary>One completed call. `waitMs` is time spent queued, `callMs` the round trip.</summary>
    public static void Record(string method, long waitMs, long callMs, bool refused)
    {
        var tally = Tallies.GetOrAdd(method, static _ => new MethodTally());

        Interlocked.Increment(ref tally.Calls);
        Interlocked.Add(ref tally.WaitTotalMs, waitMs);
        Interlocked.Add(ref tally.CallTotalMs, callMs);
        RaiseToAtLeast(ref tally.WaitMaxMs, waitMs);
        RaiseToAtLeast(ref tally.CallMaxMs, callMs);

        if (refused)
        {
            Interlocked.Increment(ref tally.Refused);
        }

        lock (Gate)
        {
            tally.Recent[tally.Cursor] = waitMs + callMs;
            tally.Cursor = (tally.Cursor + 1) % tally.Recent.Length;

            if (waitMs + callMs >= SlowMs)
            {
                // Newest first, oldest dropped. A perf hunt reads the top of this list and
                // wants the most recent evidence, not the record holder from an hour ago.
                Slowest.Insert(0, new EngineSlowCall(
                    method, waitMs, callMs, Environment.TickCount64 - PerfCounters.StartedAt));
                if (Slowest.Count > SlowestKept)
                {
                    Slowest.RemoveAt(Slowest.Count - 1);
                }
            }
        }
    }

    /// <summary>Forgets everything, so a measurement can be scoped to what happens next.</summary>
    public static void Reset()
    {
        lock (Gate)
        {
            Tallies.Clear();
            Slowest.Clear();
            Interlocked.Exchange(ref _started, Environment.TickCount64);
        }
    }

    public static (EngineMethodCost[] Methods, EngineSlowCall[] Slowest, long WindowMs) Snapshot()
    {
        lock (Gate)
        {
            var rows = Tallies
                .Select(entry =>
                {
                    var tally = entry.Value;
                    var recent = Ordered(tally.Recent, tally.Cursor);
                    var sorted = recent.OrderBy(one => one).ToArray();
                    return new EngineMethodCost(
                        entry.Key,
                        tally.Calls,
                        tally.Refused,
                        tally.WaitTotalMs,
                        tally.CallTotalMs,
                        tally.WaitMaxMs,
                        tally.CallMaxMs,
                        sorted.Length == 0 ? 0 : sorted[sorted.Length / 2],
                        sorted.Length == 0 ? 0 : sorted[Math.Min(sorted.Length - 1, (int)(sorted.Length * 0.95))],
                        recent);
                })
                // By TOTAL time, not by count or by worst case: the method to look at first is
                // the one the session actually spent its seconds in.
                .OrderByDescending(row => row.WaitTotalMs + row.CallTotalMs)
                .ToArray();

            return (rows, [.. Slowest], Environment.TickCount64 - Interlocked.Read(ref _started));
        }
    }

    private static long[] Ordered(long[] ring, int cursor)
    {
        var taken = new List<long>(ring.Length);
        for (var i = 0; i < ring.Length; i++)
        {
            var value = ring[(cursor + i) % ring.Length];
            if (value > 0)
            {
                taken.Add(value);
            }
        }

        return [.. taken];
    }

    private static void RaiseToAtLeast(ref long slot, long value)
    {
        long seen;
        while (value > (seen = Interlocked.Read(ref slot))
            && Interlocked.CompareExchange(ref slot, value, seen) != seen)
        {
        }
    }
}

/// <summary>One analyzer method's share of the session.</summary>
public sealed record EngineMethodCost(
    [property: JsonPropertyName("method")] string Method,
    [property: JsonPropertyName("calls")] long Calls,
    [property: JsonPropertyName("refused")] long Refused,
    /// <summary>Time spent QUEUED behind another call. Not the analyzer's doing.</summary>
    [property: JsonPropertyName("waitTotalMs")] long WaitTotalMs,
    /// <summary>Time spent ON the pipe. This one is the analyzer's.</summary>
    [property: JsonPropertyName("callTotalMs")] long CallTotalMs,
    [property: JsonPropertyName("waitMaxMs")] long WaitMaxMs,
    [property: JsonPropertyName("callMaxMs")] long CallMaxMs,
    [property: JsonPropertyName("medianMs")] long MedianMs,
    [property: JsonPropertyName("p95Ms")] long P95Ms,
    [property: JsonPropertyName("recentMs")] long[] RecentMs);

/// <summary>One call worth keeping on its own. `atMs` is milliseconds since the session started.</summary>
public sealed record EngineSlowCall(
    [property: JsonPropertyName("method")] string Method,
    [property: JsonPropertyName("waitMs")] long WaitMs,
    [property: JsonPropertyName("callMs")] long CallMs,
    [property: JsonPropertyName("atMs")] long AtMs);
#endif
