using Xlide.Vbe.Core.Engine;
using Xunit;

// The turns taken below are deliberately UNCANCELLABLE, which is why the token the analyzer asks
// for is left off. A cancellable turn is handed over through a wrapper that completes one
// continuation later, so "this one has not been granted yet" would be an assertion about how fast
// a continuation runs rather than about the order the gate chose. The one test that is about
// giving up brings its own token, explicitly.
#pragma warning disable xUnit1051

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The engine's pipe serves one call at a time, and the gate decides whose turn is next. The
/// rule it exists for - a query the developer is waiting for goes in front of background work -
/// is only safe because of the rule beside it: nothing goes in front of a barrier. Both are
/// pinned here, along with the two ways a gate breaks for good: losing a turn, and starving the
/// work it keeps passing.
/// </summary>
public class CallGateTests
{
    /// <summary>Takes the turn and holds it, so the tests below queue behind something.</summary>
    private static async Task<CallGate> HeldGate()
    {
        var gate = new CallGate();
        await gate.EnterAsync(CallKind.Background);
        return gate;
    }

    /// <summary>The order the gate hands turns out, one Leave at a time.</summary>
    private static async Task<List<string>> Served(CallGate gate, params (string Name, CallKind Kind)[] queued)
    {
        var waiting = new List<(string Name, Task Turn)>();
        foreach (var (name, kind) in queued)
        {
            waiting.Add((name, gate.EnterAsync(kind)));
        }

        // Everyone is queued behind the held turn before any of them is let go, which is the
        // state the choice is made in.
        //
        // The wait is over everyone NOT YET SERVED, which is not the same as everyone not yet
        // completed: Leave hands the turn out before the wait begins, so filtering on completion
        // drops the one that just took it and waits on tasks nothing will finish - the helper
        // hung, and read as the gate deadlocking.
        var order = new List<string>();
        while (order.Count < queued.Length)
        {
            gate.Leave();
            var next = await Task.WhenAny(waiting.Where(one => !order.Contains(one.Name)).Select(one => one.Turn));
            await next;
            order.Add(waiting.First(one => one.Turn == next).Name);
        }

        gate.Leave();
        return order;
    }

    [Fact]
    public void AnUncontestedTurnIsTakenWithoutWaiting()
    {
        var gate = new CallGate();
        var turn = gate.EnterAsync(CallKind.Interactive);

        Assert.True(turn.IsCompletedSuccessfully);
        gate.Leave();
    }

    [Fact]
    public async Task AQueryAPersonIsWaitingForGoesBeforeQueuedBackgroundWork()
    {
        var gate = await HeldGate();

        var order = await Served(
            gate,
            ("diagnostics", CallKind.Background),
            ("colouring", CallKind.Background),
            ("completion", CallKind.Interactive));

        Assert.Equal(["completion", "diagnostics", "colouring"], order);
    }

    [Fact]
    public async Task NothingGoesInFrontOfABarrier()
    {
        var gate = await HeldGate();

        // The didChange carries the text the completion is about. Moving the completion in
        // front of it would answer about the line as it was before the keystroke.
        var order = await Served(
            gate,
            ("diagnostics", CallKind.Background),
            ("didChange", CallKind.Barrier),
            ("completion", CallKind.Interactive));

        Assert.Equal(["diagnostics", "didChange", "completion"], order);
    }

    [Fact]
    public async Task BackgroundWorkPassedOverEnoughTimesGoesNext()
    {
        var gate = await HeldGate();

        var passedOver = gate.EnterAsync(CallKind.Background);
        for (var round = 0; round < CallGate.Patience; round++)
        {
            var jumper = gate.EnterAsync(CallKind.Interactive);
            gate.Leave();
            await jumper;
            Assert.False(passedOver.IsCompleted, $"the squiggles were passed over {round + 1} times");
        }

        // Patience spent: the next turn is its own, even with another interactive query queued
        // behind it.
        var later = gate.EnterAsync(CallKind.Interactive);
        gate.Leave();
        await passedOver;

        Assert.False(later.IsCompleted);
        gate.Leave();
        await later;
        gate.Leave();
    }

    [Fact]
    public async Task EqualKindsKeepTheirArrivalOrder()
    {
        var gate = await HeldGate();

        var order = await Served(
            gate,
            ("first", CallKind.Interactive),
            ("second", CallKind.Interactive),
            ("third", CallKind.Interactive));

        Assert.Equal(["first", "second", "third"], order);
    }

    [Fact]
    public async Task ACallerThatGivesUpLeavesTheQueueAndTheTurnStillArrives()
    {
        var gate = await HeldGate();

        using var gaveUp = new CancellationTokenSource();
        var abandoned = gate.EnterAsync(CallKind.Background, gaveUp.Token);
        var following = gate.EnterAsync(CallKind.Background);

        await gaveUp.CancelAsync();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => abandoned);

        // The turn must not be handed to the caller that has gone: that would be a pipe nobody
        // ever releases, and every language feature dead for the rest of the session.
        gate.Leave();
        await following.WaitAsync(TimeSpan.FromSeconds(5));
        gate.Leave();
    }

    [Fact]
    public async Task ATurnTakenByOneCallerIsNotGivenToAnother()
    {
        var gate = new CallGate();
        await gate.EnterAsync(CallKind.Interactive);

        var second = gate.EnterAsync(CallKind.Interactive);
        Assert.False(second.IsCompleted);

        gate.Leave();
        await second.WaitAsync(TimeSpan.FromSeconds(5));
        gate.Leave();
    }
}
