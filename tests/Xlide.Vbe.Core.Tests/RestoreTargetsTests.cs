using System;
using System.Linq;
using Xlide.Vbe.Core.Changes;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The restore planner: what each module held at a boundary, worked out from the rounds after it.
///
/// The planner is the half of restore that can be wrong quietly - the applying half either writes
/// or refuses in words - so every shape that carries risk is pinned here without a host: rename
/// chains, the pure-rename entry that holds no text, a name freed by a removal and taken by a new
/// module, and the boundary cases at zero and at the newest round.
/// </summary>
public sealed class RestoreTargetsTests
{
    private static readonly DateTimeOffset Noon = new(2026, 8, 30, 12, 0, 0, TimeSpan.Zero);

    private static DateTimeOffset At(int minute) => Noon.AddMinutes(minute);

    /// <summary>One closed round doing one thing, so tests read as the story they stage.</summary>
    private static void Round(
        ChangeLog log, int minute, string module, ChangeKind kind,
        string? before, string? after, string? from = null, string? componentKind = null)
    {
        log.Record(module, kind, before, after, "claude", At(minute), from, componentKind);
        log.Close(null, At(minute));
    }

    [Fact]
    public void AModuleWrittenAfterTheBoundaryRestoresToItsFirstBefore()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "v1", "v2");
        Round(log, 1, "Ledger", ChangeKind.Written, "v2", "v3");
        Round(log, 2, "Ledger", ChangeKind.Written, "v3", "v4");

        var target = Assert.Single(log.RestoreTargets(1));
        Assert.Equal("Ledger", target.Module);
        Assert.Equal("Ledger", target.NameNow);
        Assert.True(target.ExistedAtBoundary);

        // After round 1 (v1 -> v2) the module held v2, which is also round 2's before - the
        // planner reads it off the first touch after the boundary.
        Assert.Equal("v2", log.TextOf(target.TextKey));
    }

    [Fact]
    public void BoundaryZeroMeansBeforeEverything()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "original", "changed");

        var target = Assert.Single(log.RestoreTargets(0));
        Assert.Equal("original", log.TextOf(target.TextKey));
    }

    [Fact]
    public void TheNewestBoundaryPlansNothing()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "v1", "v2");

        Assert.Empty(log.RestoreTargets(1));
    }

    [Fact]
    public void AModuleUntouchedAfterTheBoundaryIsNotInThePlan()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "a", "b");
        Round(log, 1, "Other", ChangeKind.Written, "x", "y");

        var target = Assert.Single(log.RestoreTargets(1));
        Assert.Equal("Other", target.Module);
    }

    [Fact]
    public void AModuleAddedAfterTheBoundaryDidNotExistThere()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "a", "b");
        Round(log, 1, "Fresh", ChangeKind.Added, null, "body", componentKind: "class");

        var target = Assert.Single(log.RestoreTargets(1));
        Assert.Equal("Fresh", target.Module);
        Assert.False(target.ExistedAtBoundary);
        Assert.Null(target.TextKey);
        Assert.Equal("class", target.ComponentKind);
    }

    [Fact]
    public void ARemovedModuleRestoresWithItsLastWordsAndItsKind()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "a", "b");
        Round(log, 1, "Ledger", ChangeKind.Removed, "b", null, componentKind: "standard");

        var target = Assert.Single(log.RestoreTargets(0));
        Assert.True(target.ExistedAtBoundary);
        Assert.Equal("a", log.TextOf(target.TextKey));
        Assert.Equal("standard", target.ComponentKind);
    }

    [Fact]
    public void ARenameChainIsFollowedToTheCurrentName()
    {
        var log = new ChangeLog();
        Round(log, 0, "B", ChangeKind.Renamed, null, null, from: "A");
        Round(log, 1, "C", ChangeKind.Renamed, null, null, from: "B");
        Round(log, 2, "C", ChangeKind.Written, "v1", "v2");

        var target = Assert.Single(log.RestoreTargets(0));
        Assert.Equal("A", target.Module);
        Assert.Equal("C", target.NameNow);
        Assert.True(target.ExistedAtBoundary);
        Assert.Equal("v1", log.TextOf(target.TextKey));
    }

    [Fact]
    public void APureRenameSettlesTheNameAndNotTheText()
    {
        var log = new ChangeLog();
        Round(log, 0, "Accounts", ChangeKind.Renamed, null, null, from: "Ledger");

        var target = Assert.Single(log.RestoreTargets(0));
        Assert.Equal("Ledger", target.Module);
        Assert.Equal("Accounts", target.NameNow);
        Assert.True(target.ExistedAtBoundary);

        // A rename carries no text, so there is nothing to write - only the name to carry back.
        Assert.Null(target.TextKey);
    }

    [Fact]
    public void ANameFreedByARemovalAndTakenByANewModuleIsTwoIdentities()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "old text", "doomed");
        Round(log, 1, "Ledger", ChangeKind.Removed, "doomed", null, componentKind: "standard");
        Round(log, 2, "Ledger", ChangeKind.Added, null, "reborn", componentKind: "class");

        var targets = log.RestoreTargets(0);
        Assert.Equal(2, targets.Count);

        // The original: existed, restores to its first before, a standard module.
        var original = Assert.Single(targets, one => one.ExistedAtBoundary);
        Assert.Equal("old text", log.TextOf(original.TextKey));
        Assert.Equal("standard", original.ComponentKind);

        // The reborn one: did not exist at the boundary, so restoring removes it - and it is a
        // class, which is exactly why one name must not be one identity.
        var reborn = Assert.Single(targets, one => !one.ExistedAtBoundary);
        Assert.Equal("class", reborn.ComponentKind);
    }

    [Fact]
    public void AModuleAddedAndRemovedAfterTheBoundaryNeverExistedThere()
    {
        var log = new ChangeLog();
        Round(log, 0, "Fleeting", ChangeKind.Added, null, "hello");
        Round(log, 1, "Fleeting", ChangeKind.Removed, "hello", null);

        var target = Assert.Single(log.RestoreTargets(0));
        Assert.False(target.ExistedAtBoundary);
        Assert.Null(target.TextKey);
    }

    [Fact]
    public void OnlyClosedRoundsArePlannedFrom()
    {
        var log = new ChangeLog();
        Round(log, 0, "Ledger", ChangeKind.Written, "v1", "v2");
        log.Record("Ledger", ChangeKind.Written, "v2", "v3", "claude", At(1));

        // The running round is invisible to the plan; the caller closes it first.
        var target = Assert.Single(log.RestoreTargets(0));
        Assert.Equal("v1", log.TextOf(target.TextKey));
    }

    [Fact]
    public void TheKindSurvivesARenameWithinTheRound()
    {
        var log = new ChangeLog();
        log.Record("Ledger", ChangeKind.Written, "a", "b", "claude", At(0), componentKind: "class");
        log.Record("Accounts", ChangeKind.Renamed, null, null, "claude", At(0), from: "Ledger");
        log.Close(null, At(0));

        var entry = Assert.Single(log.Rounds().Single().Entries);
        Assert.Equal("class", entry.ComponentKind);
        Assert.Equal("Accounts", entry.Module);
    }
}
