using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Where a module's breakpoints are once its text has been written back to the editor.
///
/// The editor keeps the real breakpoints and will not say where they are, so the record the
/// margin draws from has to follow what the write did. Issue #84 is what following the text
/// instead looks like: a statement deleted under a breakpoint left its dot on the next line, and
/// F9 there asked the editor to toggle a breakpoint it did not have - which SET one, real and
/// undrawn. Each case below is one shape of write and where the editor leaves breakpoints after
/// it, measured on 2026-10-08 against a live Excel.
/// </summary>
public class BreakpointWriteTests
{
    private static readonly string[] Written =
    [
        /* 1 */ "Option Explicit",
        /* 2 */ "",
        /* 3 */ "Public Sub Walk()",
        /* 4 */ "    Debug.Print 1",
        /* 5 */ "    Debug.Print 2",
        /* 6 */ "    Debug.Print 3",
        /* 7 */ "End Sub",
    ];

    private static Breakpoints.WrittenWindow[] One(int at, int removing, int inserting, bool inPlace) =>
        [new(at, removing, inserting, inPlace)];

    [Fact]
    public void ALineBelowAGrownWindowMovesDown()
    {
        // Two lines inserted before line 4: the breakpoint on 4 is now on 6.
        Assert.Equal([6], Breakpoints.AfterWrite([4], One(4, 0, 2, false), Written));
    }

    [Fact]
    public void ALineBelowAShrunkWindowMovesUp()
    {
        // Line 4 deleted: the breakpoint on 6 is now on 5.
        Assert.Equal([5], Breakpoints.AfterWrite([6], One(4, 1, 0, false), Written));
    }

    [Fact]
    public void ALineAboveAWindowStaysPut()
    {
        Assert.Equal([4], Breakpoints.AfterWrite([4], One(5, 1, 3, false), Written));
    }

    [Fact]
    public void ALineReplacedInPlaceKeepsItsBreakpoint()
    {
        // Typing on the line: the editor replaced it one for one and kept the breakpoint.
        Assert.Equal([5], Breakpoints.AfterWrite([5], One(5, 1, 1, true), Written));
    }

    [Fact]
    public void ALineReplacedInPlaceByABlankLosesIt()
    {
        // The text was cleared: the editor cannot hold a breakpoint on a blank line and drops it.
        string[] cleared = [.. Written];
        cleared[4] = "";
        Assert.Empty(Breakpoints.AfterWrite([5], One(5, 1, 1, true), cleared));
    }

    [Fact]
    public void ADeletedLineTakesItsBreakpointWithIt()
    {
        // Lines 4 and 5 deleted and one line inserted in their place. Even where the inserted
        // text repeats a deleted line, the editor forgot the breakpoint with the deletion.
        Assert.Empty(Breakpoints.AfterWrite([4, 5], One(4, 2, 1, false), Written));
    }

    [Fact]
    public void BreakpointsOutsideAWindowAreUnmoved()
    {
        Assert.Equal([3, 6], Breakpoints.AfterWrite([3, 6], One(4, 1, 1, false), Written));
    }

    [Fact]
    public void SeveralWindowsAccumulateTheirShifts()
    {
        // +1 line at 2, then -1 line at 5 (the old 5): a breakpoint on 4 moves to 5, one on 6 stays 6.
        Breakpoints.WrittenWindow[] windows = [new(2, 0, 1, false), new(5, 1, 0, false)];
        Assert.Equal([5, 6], Breakpoints.AfterWrite([4, 6], windows, Written));
    }

    [Fact]
    public void AModuleReplacedWholeKeepsNothing()
    {
        Assert.Empty(Breakpoints.AfterWrite([3, 4, 5], One(1, 7, 7, false), Written));
    }

    [Fact]
    public void NoWindowsChangeNothing()
    {
        Assert.Equal([4, 5], Breakpoints.AfterWrite([4, 5], [], Written));
    }
}
