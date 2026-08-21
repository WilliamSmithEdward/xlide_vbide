using System.Diagnostics;
using System.Linq;
using System.Text;
using Xlide.Vbe.Core.Sync;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// A side-by-side comparison of two large modules must not take the process down with it.
///
/// THE DEFECT THIS EXISTS FOR. The comparison lines two texts up with a longest common
/// subsequence, and the table for that is `new int[left + 1, right + 1]` - quadratic in time and
/// in MEMORY. Measured on generated modules whose middles share nothing: 803 lines 5ms and 3 MB,
/// 1,603 lines 21ms and 10 MB, 3,203 lines 83ms and 40 MB, 6,403 lines 353ms and 158 MB. Four
/// times the cost for twice the lines, all the way up.
///
/// VBA holds 65,534 lines in one module and this product ships a fixture at 64,802 of them, where
/// that extrapolates to roughly 36 seconds and 16 GB - an OutOfMemoryException in the middle of an
/// import dialog somebody was about to accept, not a slow redraw. Nothing capped it (2026-08-21).
///
/// The head and the tail are taken off before the table is built, so this is only reached by two
/// texts that genuinely disagree over thousands of lines. Those are exactly the ones where lining
/// up line by line tells a reader nothing, so past the cap the answer is the cheap honest one:
/// this went, that came.
/// </summary>
public sealed class ModuleSyncDiffCostTests
{
    private static string Body(int procedures, string prefix)
    {
        var builder = new StringBuilder();
        builder.Append("Option Explicit\r\n\r\n");
        for (var n = 0; n < procedures; n++)
        {
            builder.Append($"Public Function {prefix}{n}(ByVal seed As Long) As Long\r\n");
            builder.Append($"    {prefix}{n} = seed * {n + 1}\r\n");
            builder.Append("End Function\r\n\r\n");
        }

        return builder.ToString();
    }

    [Fact]
    public void TwoLargeModulesThatShareNothingCompareInBoundedTime()
    {
        // 12,000 procedures is about 48,000 lines: past the cap several times over, and the size
        // at which the unbounded table would have been tens of gigabytes.
        var left = Body(12_000, "A");
        var right = Body(12_000, "B");

        var watch = Stopwatch.StartNew();
        var diff = ModuleSync.Diff(left, right);
        watch.Stop();

        Assert.True(
            watch.ElapsedMilliseconds < 4_000,
            $"comparing two 48,000-line modules took {watch.ElapsedMilliseconds}ms; the middle is "
            + "being lined up with a table again rather than answered as a block.");

        // Every line of both is accounted for, which is what makes the cheap answer honest.
        Assert.Equal(left.Split("\r\n").Length, diff.Count(one => one.LeftNumber is not null));
        Assert.Equal(right.Split("\r\n").Length, diff.Count(one => one.RightNumber is not null));
    }

    [Fact]
    public void ALargeModuleWithOneChangedLineIsStillLinedUpExactly()
    {
        // The head and tail come off first, so a real edit never reaches the cap however large the
        // module is. This is the case the cap must not touch.
        var left = Body(12_000, "A");
        var right = left.Replace("    A6000 = seed * 6001\r\n", "    A6000 = seed * 6001 + 1\r\n");
        Assert.NotEqual(left, right);

        var diff = ModuleSync.Diff(left, right);

        Assert.Equal(1, diff.Count(one => one.Kind != DiffKind.Equal));
        Assert.Equal(DiffKind.Changed, diff.Single(one => one.Kind != DiffKind.Equal).Kind);
    }

    [Fact]
    public void ASmallRewriteIsStillLinedUpLineByLine()
    {
        // Under the cap nothing changed: the same table, the same answer as before.
        var left = "Option Explicit\r\nDim a As Long\r\nDim b As Long\r\nEnd\r\n";
        var right = "Option Explicit\r\nDim a As Long\r\nDim c As Long\r\nDim d As Long\r\nEnd\r\n";

        var diff = ModuleSync.Diff(left, right);

        Assert.Equal(DiffKind.Changed, diff[2].Kind);
        Assert.Equal(DiffKind.Added, diff[3].Kind);
        Assert.Equal("Dim d As Long", diff[3].Right);
    }
}
