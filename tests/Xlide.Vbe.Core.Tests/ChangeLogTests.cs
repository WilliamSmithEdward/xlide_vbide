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
public sealed class ChangeLogTests
{
    private static readonly DateTimeOffset Noon =
        new(2026, 8, 21, 12, 0, 0, TimeSpan.Zero);

    private static ChangeLog Open() => new();

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
    public void AcceptingMarksTheRoundsWithoutDestroyingAnything()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", "claude", Noon);
        log.Accept(Noon.AddSeconds(5));

        Assert.Equal(1, log.AcceptedAt);

        // The round is still there, and so is what the module held before it. Accepting moves
        // where a reader starts counting; it does not remove the past.
        Assert.Single(log.Rounds());
        Assert.Equal("a", log.TextOf(Assert.Single(Assert.Single(log.Rounds()).Entries).Before));

        // It also ends the round that was running, so what comes next is counted apart.
        log.Record("Parser", ChangeKind.Written, "b", "c", "claude", Noon.AddSeconds(6));
        Assert.Equal(2, log.Rounds().Count);
        Assert.Equal(1, log.AcceptedAt);
    }

    [Fact]
    public void ATextIsHeldOnceHoweverManyRoundsNameIt()
    {
        // Kept by content: the after of one round is the before of the next, and that is one copy.
        // This is what keeps a long session over a large module from costing a copy per write.
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "one", "two", "claude", Noon);
        log.Close(null, Noon.AddSeconds(1));
        log.Record("Parser", ChangeKind.Written, "two", "three", "developer", Noon.AddSeconds(2));

        // "one", "two" and "three" - and "two" only once, though two rounds name it.
        Assert.Equal(("one".Length + "two".Length + "three".Length) * 2L, log.HeldBytes);
    }

    [Fact]
    public void TheOldestRoundsLetTheirTextsGoRatherThanGrowingWithoutBound()
    {
        // Nothing evicts this now that it is memory rather than disk, so the log ages itself. The
        // entries stay - they are the record - and what goes is the text, which says so.
        var log = Open();
        var big = new string('x', 4 * 1024 * 1024);

        for (var round = 0; round < 12; round++)
        {
            log.Record("Big", ChangeKind.Written, $"{big}{round}", $"{big}{round + 1}", "claude", Noon.AddMinutes(round));
            log.Close(null, Noon.AddMinutes(round).AddSeconds(1));
        }

        Assert.True(
            log.HeldBytes <= ChangeLog.LargestHeldBytes,
            $"the log is holding {log.HeldBytes} bytes, past its own budget");

        // Every round is still there, and the newest can still show what it did.
        Assert.Equal(12, log.Rounds().Count);
        Assert.NotNull(log.TextOf(log.Rounds()[0].Entries[0].After));

        // The oldest has let its text go, and answering null is how it says so.
        Assert.Null(log.TextOf(log.Rounds()[^1].Entries[0].Before));
    }

    [Fact]
    public void ARenameMovesTheEntryRatherThanStartingAnother()
    {
        // A round's entries are keyed by the module's name, and a rename changes it - so without
        // this a module renamed and then written reads as two modules, one of which is gone.
        var log = Open();
        log.Record("Ledger", ChangeKind.Written, "before", "after", "claude", Noon);
        log.Record("Accounts", ChangeKind.Renamed, null, null, "claude", Noon.AddSeconds(1), from: "Ledger");
        log.Record("Accounts", ChangeKind.Written, "after", "after and more", "claude", Noon.AddSeconds(2));

        var entry = Assert.Single(Assert.Single(log.Rounds()).Entries);
        Assert.Equal("Accounts", entry.Module);
        Assert.Equal("Ledger", entry.From);

        // Written, not Renamed: this module was renamed AND its text changed, and `From` is what
        // says the first while the kind says the second. Collapsing both into Renamed would lose
        // the more interesting half. A bare rename, with nothing written, is Renamed.
        Assert.Equal(ChangeKind.Written, entry.Kind);

        // And it still knows what the module held before any of it.
        Assert.Equal("before", log.TextOf(entry.Before));
        Assert.Equal("after and more", log.TextOf(entry.After));
    }

    [Fact]
    public void ACallerThatDoesNotSayWhoItIsIsRecordedAsUnattributed()
    {
        var log = Open();
        log.Record("Parser", ChangeKind.Written, "a", "b", null, Noon);

        Assert.Equal(ChangeLog.Unattributed, Assert.Single(log.Rounds()).Author);
    }
}
