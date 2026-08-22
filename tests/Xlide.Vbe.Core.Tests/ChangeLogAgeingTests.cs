using System;
using System.Diagnostics;
using Xlide.Vbe.Core.Changes;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// What ageing COSTS, and what it does when it cannot help.
///
/// The log holds whole module texts, so it ages: past its budget the oldest rounds let theirs go.
/// That is tested elsewhere and works. What is tested here is the two cases where the budget
/// cannot be met - a single round holding more than the whole budget - and where meeting it means
/// dropping a great many rounds at once. Both are reachable from ordinary use at VBA's per-module
/// ceiling, where one copy of a module is about 1.5 MB and forty rounds is an afternoon.
/// </summary>
public sealed class ChangeLogAgeingTests
{
    private static readonly DateTimeOffset Noon = new(2026, 8, 22, 12, 0, 0, TimeSpan.Zero);

    /// <summary>A megabyte of distinct text, so no two are the same content.</summary>
    private static string Big(int number) => new string((char)('a' + (number % 26)), 1024 * 1024) + number;

    [Fact]
    public void AgeingManyRoundsAtOnceStaysCheap()
    {
        var log = new ChangeLog();

        // Eighty megabytes of distinct text across eighty rounds, which is past the budget and so
        // makes ageing do real work rather than return at the first check.
        var clock = Stopwatch.StartNew();
        for (var round = 0; round < 80; round++)
        {
            log.Record("Massive", ChangeKind.Written, Big(round), Big(round + 1), "claude", Noon.AddMinutes(round));
            log.Close(null, Noon.AddMinutes(round).AddSeconds(1));
        }

        clock.Stop();

        Assert.True(
            log.HeldBytes <= ChangeLog.LargestHeldBytes,
            $"the log is holding {log.HeldBytes} bytes, past its own budget");

        // A BOUND RATHER THAN A BENCHMARK. This runs on whatever machine builds the product, so
        // the number is loose on purpose: what it is there to catch is the shape going quadratic,
        // which at this size costs seconds rather than the tens of milliseconds it should.
        Assert.True(
            clock.ElapsedMilliseconds < 4000,
            $"eighty rounds took {clock.ElapsedMilliseconds}ms to record and age");
    }

    [Fact]
    public void ARoundBiggerThanTheWholeBudgetDoesNotSpin()
    {
        // THE CASE AGEING CANNOT FIX. One round holding more than the budget has nothing older to
        // drop, so the log stays over budget - which is correct, because the alternative is
        // throwing away the round the developer is looking at. What must NOT happen is the
        // attempt costing something every time: the loop walks the remaining rounds and every
        // held text on each pass, so a case it can never satisfy is a case it must not keep
        // retrying.
        var log = new ChangeLog();

        for (var write = 0; write < 70; write++)
        {
            log.Record("Massive", ChangeKind.Written, Big(write), Big(write + 1), "claude", Noon);
        }

        var clock = Stopwatch.StartNew();
        log.Close("one enormous round", Noon.AddSeconds(1));
        clock.Stop();

        // It is over budget and says so honestly rather than having dropped the only round it has.
        Assert.Single(log.Rounds());

        Assert.True(
            clock.ElapsedMilliseconds < 2000,
            $"closing an over-budget round took {clock.ElapsedMilliseconds}ms");
    }

    [Fact]
    public void CloseIsCheapWhenNothingNeedsAgeing()
    {
        // The common case by far: a session well inside the budget closing a round. Ageing should
        // cost nothing measurable here, and a shape that walks everything on every close would
        // show up as this getting slower the longer a session runs.
        var log = new ChangeLog();
        var text = new string('x', 4096);

        for (var round = 0; round < 400; round++)
        {
            log.Record("Small", ChangeKind.Written, $"{text}{round}", $"{text}{round + 1}", "claude",
                Noon.AddMinutes(round));
            log.Close(null, Noon.AddMinutes(round).AddSeconds(1));
        }

        var clock = Stopwatch.StartNew();
        log.Record("Small", ChangeKind.Written, "a", "b", "claude", Noon.AddHours(9));
        log.Close(null, Noon.AddHours(9).AddSeconds(1));
        clock.Stop();

        // ASKED OF THE STORE, not of the default view. `Rounds()` stops at its limit, so counting
        // what it returns measures the limit rather than the log - which is how this check first
        // read 200 out of a log holding 401, and how the silent truncation got noticed at all.
        Assert.Equal(401, log.RoundCount);
        Assert.Equal(200, log.Rounds().Count);
        Assert.Equal(401, log.Rounds(int.MaxValue).Count);

        Assert.True(
            clock.ElapsedMilliseconds < 250,
            $"the 401st close took {clock.ElapsedMilliseconds}ms, with the log well inside its budget");
    }
}
