using System;
using System.Diagnostics;
using System.Linq;
using Xlide.Vbe.Core.Changes;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// What recording a change costs, at the size where it could hurt.
///
/// RECORDING IS ON THE WRITE PATH. Every write hands the log the module's whole text before and
/// after, and the log keeps texts by content - which means hashing both. At VBA's per-module
/// ceiling that is about 3.8 MB hashed per write, and this product has twice learned what a
/// whole-text operation costs when it lands somewhere that runs often.
///
/// It is fine, measured: about 2.7 ms at the ceiling, scaling linearly with the text, and writes
/// are flushed when typing stops rather than per keystroke - so this is paid at the rate a
/// developer finishes thoughts. The bound below is loose enough not to fail on a slow machine and
/// tight enough to notice the shape changing, which is what it is for.
/// </summary>
public sealed class RecordCostTests
{
    private static readonly DateTimeOffset Noon = new(2026, 8, 22, 12, 0, 0, TimeSpan.Zero);

    private static string ModuleOf(int lines) => string.Join("\r\n", Enumerable.Range(0, lines)
        .Select(at => $"    Debug.Print \"line {at} of a module that is {lines} long\""));

    [Fact]
    public void RecordingACeilingSizedModuleStaysOffTheWritePath()
    {
        // 64,802 lines is VBA's own per-module ceiling, so this is the worst case that exists.
        var text = ModuleOf(64_802);
        var next = $"{text}\r\n    ' one more";
        var log = new ChangeLog();

        // Warm, so the figure is the work rather than the first call through it.
        log.Record("Massive", ChangeKind.Written, text, next, "claude", Noon);

        var clock = Stopwatch.StartNew();
        for (var at = 0; at < 20; at++)
        {
            log.Record("Massive", ChangeKind.Written, text, next + at, "claude", Noon.AddSeconds(at));
        }

        clock.Stop();
        var each = clock.Elapsed.TotalMilliseconds / 20;

        Assert.True(each < 25, $"recording a 3.8 MB module took {each:F2} ms per write");
    }

    [Fact]
    public void TheCostFollowsTheTextRatherThanTheHistory()
    {
        // THE SHAPE THAT WOULD HURT is a cost that grows with how much the log already holds,
        // because that turns a long session into a slow editor. Keeping texts by content means
        // hashing one text, not walking the ones already kept - and this is what says so: two
        // hundred rounds in, a write costs what it cost at the start.
        var text = ModuleOf(2_000);
        var log = new ChangeLog();

        var early = Stopwatch.StartNew();
        for (var at = 0; at < 20; at++)
        {
            log.Record("M", ChangeKind.Written, text, $"{text}{at}", "claude", Noon.AddSeconds(at));
        }

        early.Stop();

        for (var round = 0; round < 200; round++)
        {
            log.Record("M", ChangeKind.Written, text, $"{text}filler{round}", "claude", Noon.AddMinutes(round));
            log.Close(null, Noon.AddMinutes(round).AddSeconds(1));
        }

        var late = Stopwatch.StartNew();
        for (var at = 0; at < 20; at++)
        {
            log.Record("M", ChangeKind.Written, text, $"{text}late{at}", "claude", Noon.AddHours(9).AddSeconds(at));
        }

        late.Stop();

        // Generous, because these are milliseconds on a shared machine. What it refuses is the
        // late writes costing several times the early ones, which is what growth with history
        // would look like.
        Assert.True(
            late.Elapsed.TotalMilliseconds < Math.Max(50, early.Elapsed.TotalMilliseconds * 6),
            $"twenty writes cost {early.Elapsed.TotalMilliseconds:F1} ms at the start and "
            + $"{late.Elapsed.TotalMilliseconds:F1} ms after 200 rounds");
    }
}
