using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The VBE validates a line when the caret leaves it. The hold is the publish-side half of that
/// contract, and these pin its edges: a click hides nothing, typing hides the typed line's
/// verdicts everywhere they span, and leaving the line is the one and only reveal.
/// </summary>
public class ActiveLineHoldTests
{
    [Fact]
    public void NothingIsHiddenBeforeTypingBeginsAHold()
    {
        var hold = new ActiveLineHold();

        Assert.False(hold.IsHolding);
        Assert.False(hold.Hides("Module1", 40, 40));
    }

    [Fact]
    public void TypingHoldsTheTypedLineOfTheTypedModuleOnly()
    {
        var hold = new ActiveLineHold();

        Assert.True(hold.Begin("Module1", 40));

        Assert.True(hold.Hides("Module1", 40, 40));
        Assert.True(hold.Hides("module1", 40, 40));
        Assert.False(hold.Hides("Module1", 39, 39));
        Assert.False(hold.Hides("Other", 40, 40));
    }

    [Fact]
    public void AFindingSpanningTheHeldLineIsHidden()
    {
        var hold = new ActiveLineHold();
        hold.Begin("Module1", 40);

        Assert.True(hold.Hides("Module1", 38, 42));
        Assert.False(hold.Hides("Module1", 41, 42));
    }

    [Fact]
    public void RepeatedTypingOnTheSameLineChangesNothing()
    {
        var hold = new ActiveLineHold();

        Assert.True(hold.Begin("Module1", 40));
        Assert.False(hold.Begin("Module1", 40));
    }

    [Fact]
    public void TypingOnAnotherLineMovesTheHoldThere()
    {
        var hold = new ActiveLineHold();
        hold.Begin("Module1", 40);

        Assert.True(hold.Begin("Module1", 41));
        Assert.False(hold.Hides("Module1", 40, 40));
        Assert.True(hold.Hides("Module1", 41, 41));
    }

    [Fact]
    public void TheCaretSettlingOnTheHeldLineKeepsTheHold()
    {
        var hold = new ActiveLineHold();
        hold.Begin("Module1", 40);

        Assert.False(hold.Release("Module1", 40));
        Assert.True(hold.IsHolding);
    }

    [Fact]
    public void TheCaretLeavingTheLineReleasesAndAsksForARepublish()
    {
        var hold = new ActiveLineHold();
        hold.Begin("Module1", 40);

        Assert.True(hold.Release("Module1", 41));
        Assert.False(hold.IsHolding);
        Assert.False(hold.Hides("Module1", 40, 40));
    }

    [Fact]
    public void ReleasingWithoutAHoldAsksForNothing()
    {
        var hold = new ActiveLineHold();

        Assert.False(hold.Release("Module1", 41));
        Assert.False(hold.Release());
    }

    [Fact]
    public void AModuleSwitchReleasesUnconditionally()
    {
        var hold = new ActiveLineHold();
        hold.Begin("Module1", 40);

        Assert.True(hold.Release());
        Assert.False(hold.IsHolding);
    }
}
