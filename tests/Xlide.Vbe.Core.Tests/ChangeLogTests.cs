using System;
using System.IO;
using System.Linq;
using Xlide.Vbe.Core.Changes;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The change log: what happened to a project's module code, by whom, in rounds.
///
/// The clock is a parameter throughout, so the rules that depend on it - a silence ending a round,
/// a round's own start and end - are tested rather than waited for.
/// </summary>
public sealed class ChangeLogTests : IDisposable
{
    private readonly string _home = Path.Combine(
        Path.GetTempPath(), $"xlide-changelog-{Guid.NewGuid():N}");

    private static readonly DateTimeOffset Noon =
        new(2026, 8, 21, 12, 0, 0, TimeSpan.Zero);

    private ChangeLog Open() => ChangeLog.For(_home, @"C:\books\Ledger.xlsm");

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_home))
            {
                Directory.Delete(_home, recursive: true);
            }
        }
        catch (IOException)
        {
        }
    }

    [Fact]
    public void AWriteOpensARoundAndTheRoundHoldsWhatItDid()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "old", "new", "claude", Noon);

        var round = Assert.Single(log.Rounds());
        Assert.True(round.Open);
        Assert.Equal("claude", round.Author);
        Assert.Equal(1, round.Number);

        var entry = Assert.Single(round.Entries);
        Assert.Equal("Parser", entry.Module);
        Assert.Equal(ChangeKind.Written, entry.Kind);
        Assert.Equal("old", log.TextOf(entry.Before));
        Assert.Equal("new", log.TextOf(entry.After));
    }

    [Fact]
    public void WritingAModuleFiveTimesInOneRoundKeepsTheFirstBeforeAndTheLastAfter()
    {
        // The whole point of a round: it says what it did to a module ONCE, however many times an
        // agent rewrote it on the way to the answer.
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "v1", "v2", "claude", Noon);
        log.Record("Parser", ChangeKind.Written, "v2", "v3", "claude", Noon.AddSeconds(2));
        log.Record("Parser", ChangeKind.Written, "v3", "v4", "claude", Noon.AddSeconds(4));

        var entry = Assert.Single(Assert.Single(log.Rounds()).Entries);
        Assert.Equal("v1", log.TextOf(entry.Before));
        Assert.Equal("v4", log.TextOf(entry.After));
    }

    [Fact]
    public void ADifferentHandIsADifferentRound()
    {
        // Agent, then the developer, then the agent again is three rounds - which is what lets
        // someone read their own edit sitting between two of an agent's rather than folded in.
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        log.Record("Parser", ChangeKind.Written, "b", "c", "developer", Noon.AddSeconds(10));
        log.Record("Parser", ChangeKind.Written, "c", "d", "claude", Noon.AddSeconds(20));

        var rounds = log.Rounds();
        Assert.Equal(3, rounds.Count);
        Assert.Equal(["claude", "developer", "claude"], rounds.Select(one => one.Author).Reverse());
    }

    [Fact]
    public void ALongSilenceEndsARoundEvenWhenNobodySaysSo()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        log.Record("Parser", ChangeKind.Written, "b", "c", "claude", Noon + ChangeLog.Silence.Add(TimeSpan.FromMinutes(1)));

        Assert.Equal(2, log.Rounds().Count);
    }

    [Fact]
    public void ASnapshotEndsTheRoundAndLabelsIt()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        log.Close("taught Parser about dates", Noon.AddSeconds(30));

        var round = Assert.Single(log.Rounds());
        Assert.False(round.Open);
        Assert.Equal("taught Parser about dates", round.Label);
    }

    [Fact]
    public void ClosingWhenNothingIsRunningDoesNothing()
    {
        // A caller may say "that was a round" as often as it likes without inventing empty ones.
        var log = Open();
        log.Close("nothing happened", Noon);
        log.Close("still nothing", Noon.AddSeconds(1));

        Assert.Empty(log.Rounds());
    }

    [Fact]
    public void AnAddedModuleStaysAnAddEvenAfterItIsWrittenAgain()
    {
        var log = Open();
        log.Record("Fresh", ChangeKind.Added, null, "first", "claude", Noon);
        log.Record("Fresh", ChangeKind.Written, "first", "second", "claude", Noon.AddSeconds(1));

        var entry = Assert.Single(Assert.Single(log.Rounds()).Entries);
        Assert.Equal(ChangeKind.Added, entry.Kind);
        Assert.Null(entry.Before);
        Assert.Equal("second", log.TextOf(entry.After));
    }

    [Fact]
    public void ARemovedModuleIsARemovalWhateverCameBefore()
    {
        var log = Open();
        log.Record("Doomed", ChangeKind.Written, "a", "b", "claude", Noon);
        log.Record("Doomed", ChangeKind.Removed, "b", null, "claude", Noon.AddSeconds(1));

        var entry = Assert.Single(Assert.Single(log.Rounds()).Entries);
        Assert.Equal(ChangeKind.Removed, entry.Kind);
        Assert.Null(entry.After);
    }

    [Fact]
    public void TheLogSurvivesTheSessionThatWroteIt()
    {
        var first = Open();
        first.Record("Parser", ChangeKind.Written, "old", "new", "claude", Noon);
        first.Close("a round", Noon.AddSeconds(5));

        var second = Open();
        var round = Assert.Single(second.Rounds());
        Assert.Equal("a round", round.Label);
        Assert.Equal("claude", round.Author);
        Assert.Equal("old", second.TextOf(Assert.Single(round.Entries).Before));
    }

    [Fact]
    public void ARoundLeftRunningWhenTheSessionEndedIsStillReadAsARound()
    {
        var first = Open();
        first.Record("Parser", ChangeKind.Written, "old", "new", "claude", Noon);
        // No Close: the host went away.

        var second = Open();
        var round = Assert.Single(second.Rounds());
        Assert.False(round.Open);
        Assert.Equal("Parser", Assert.Single(round.Entries).Module);
    }

    [Fact]
    public void TheNextSessionCountsOnFromWhereTheLastOneStopped()
    {
        var first = Open();
        first.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        first.Close(null, Noon.AddSeconds(1));

        var second = Open();
        second.Record("Parser", ChangeKind.Written, "b", "c", "claude", Noon.AddMinutes(30));

        Assert.Equal([2, 1], second.Rounds().Select(one => one.Number));
    }

    [Fact]
    public void AcceptingDrawsALineWithoutDestroyingAnything()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        log.Accept(Noon.AddSeconds(5));

        Assert.Equal(1, log.AcceptedAt);
        Assert.Single(log.Rounds());
        Assert.Equal("a", log.TextOf(Assert.Single(Assert.Single(log.Rounds()).Entries).Before));

        var later = Open();
        Assert.Equal(1, later.AcceptedAt);
        Assert.Single(later.Rounds());
    }

    [Fact]
    public void ATextIsKeptOnceHoweverManyRoundsNameIt()
    {
        // Content-addressed: the after of one round is the before of the next, and that is one
        // file. This is what keeps a long session from costing a copy per write.
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "one", "two", "claude", Noon);
        log.Close(null, Noon.AddSeconds(1));
        log.Record("Parser", ChangeKind.Written, "two", "three", "developer", Noon.AddSeconds(2));

        var texts = Directory.GetFiles(Path.Combine(log.Directory, "texts"));
        Assert.Equal(3, texts.Length);
    }

    [Fact]
    public void TwoWorkbooksWithTheSameNameAreTwoLogs()
    {
        var left = ChangeLog.For(_home, @"C:\one\Ledger.xlsm");
        var right = ChangeLog.For(_home, @"C:\two\Ledger.xlsm");

        Assert.NotEqual(left.Directory, right.Directory);

        left.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        Assert.Empty(right.Rounds());
    }

    [Fact]
    public void ACallerThatDoesNotSayWhoItIsIsRecordedAsUnattributed()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", null, Noon);

        Assert.Equal(ChangeLog.Unattributed, Assert.Single(log.Rounds()).Author);
    }
}
