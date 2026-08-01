using Xlide.Vbe.Core.Hosting;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// A sited control is told where to sit by a position rectangle and a clipping rectangle at once,
/// and has to reconcile them itself. Getting that wrong produces a window in the wrong place rather
/// than an error, so the arithmetic is pinned here where no container is needed to check it.
/// </summary>
public class PixelRectTests
{
    [Fact]
    public void SizeIsTheDistanceBetweenEdges()
    {
        var rect = new PixelRect(10, 20, 110, 220);

        Assert.Equal(100, rect.Width);
        Assert.Equal(200, rect.Height);
        Assert.False(rect.IsEmpty);
    }

    [Fact]
    public void FromSizeAnchorsAtTheOrigin()
    {
        var rect = PixelRect.FromSize(5, 7, 40, 60);

        Assert.Equal(new PixelRect(5, 7, 45, 67), rect);
    }

    [Fact]
    public void AnInvertedRectangleReportsNoSizeRatherThanANegativeOne()
    {
        // A container is entitled to hand over a degenerate rectangle while a pane is being
        // dragged. Passing a negative width to a sizing call is accepted and produces nonsense.
        var rect = new PixelRect(100, 100, 40, 40);

        Assert.Equal(0, rect.Width);
        Assert.Equal(0, rect.Height);
        Assert.True(rect.IsEmpty);
    }

    [Fact]
    public void IntersectionIsTheOverlap()
    {
        var position = new PixelRect(0, 0, 200, 200);
        var clip = new PixelRect(50, 60, 300, 120);

        Assert.Equal(new PixelRect(50, 60, 200, 120), position.Intersect(clip));
    }

    [Fact]
    public void IntersectionWithAnEnclosingRectangleChangesNothing()
    {
        var position = new PixelRect(10, 10, 100, 100);
        var clip = new PixelRect(0, 0, 1000, 1000);

        Assert.Equal(position, position.Intersect(clip));
    }

    [Fact]
    public void FullyClippedProducesAnEmptyRectangleAtTheOriginalOrigin()
    {
        // Being scrolled entirely out of view is normal, and the answer has to be an empty window
        // where the control belongs, not a window somewhere else.
        var position = new PixelRect(10, 10, 100, 100);
        var clip = new PixelRect(500, 500, 600, 600);

        var result = position.Intersect(clip);

        Assert.True(result.IsEmpty);
        Assert.Equal(10, result.Left);
        Assert.Equal(10, result.Top);
    }

    [Fact]
    public void TouchingEdgesDoNotOverlap()
    {
        var position = new PixelRect(0, 0, 100, 100);
        var clip = new PixelRect(100, 0, 200, 100);

        Assert.True(position.Intersect(clip).IsEmpty);
    }

    [Fact]
    public void IntersectionIsCommutativeInSize()
    {
        var a = new PixelRect(0, 0, 200, 100);
        var b = new PixelRect(50, 25, 300, 75);

        var first = a.Intersect(b);
        var second = b.Intersect(a);

        Assert.Equal(first.Width, second.Width);
        Assert.Equal(first.Height, second.Height);
    }

    [Fact]
    public void AtOriginKeepsTheSizeAndDropsThePosition()
    {
        // A hosted browser is positioned in its parent's client coordinates, which start at zero
        // however far the parent is from the container's own origin.
        var rect = new PixelRect(120, 340, 320, 540);

        Assert.Equal(new PixelRect(0, 0, 200, 200), rect.AtOrigin());
    }
}
