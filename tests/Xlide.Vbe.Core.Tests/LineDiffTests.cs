using Xlide.Vbe.Core.Editor;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The arithmetic behind a write-back. An off-by-one here does not throw: it removes a line the
/// developer wrote, or inserts one where it does not belong, and the module keeps whatever came
/// out. So the window is checked here, on strings, rather than only against a live editor.
/// </summary>
public class LineDiffTests
{
    private const int Window = 400;

    private static string Module(params string[] lines) => string.Join("\r\n", lines);

    /// <summary>Applies a window the way the editor does, so a case can be checked end to end.</summary>
    private static string Apply(string baseline, LineDiff diff)
    {
        var lines = baseline.Split(["\r\n", "\n"], StringSplitOptions.None).ToList();
        lines.RemoveRange(diff.At - 1, diff.Removing);

        // On the COUNT, the way the editor is driven: an empty window text is one empty line when
        // the diff says one line goes in, and nothing at all when it says none.
        if (diff.Inserting > 0)
        {
            lines.InsertRange(diff.At - 1, diff.Text.Split("\r\n"));
        }

        return string.Join("\r\n", lines);
    }

    [Fact]
    public void TextThatHasNotMovedIsNotAWrite()
    {
        var held = Module("Option Explicit", "", "Public Sub One()", "End Sub");
        var diff = LineDiff.Between(held, held, Window);

        Assert.Equal(LineChange.Identical, diff.Change);
        Assert.Equal(4, diff.TotalLines);
    }

    [Fact]
    public void OneChangedLineIsAWindowOfOne()
    {
        var was = Module("Option Explicit", "", "Public Sub One()", "    x = 1", "End Sub");
        var now = Module("Option Explicit", "", "Public Sub One()", "    x = 2", "End Sub");

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(4, diff.At);
        Assert.Equal(1, diff.Removing);
        Assert.Equal("    x = 2", diff.Text);
        Assert.Equal("    x = 1", diff.Removed);
        Assert.Equal(5, diff.TotalLines);
        Assert.Equal(now, Apply(was, diff));
    }

    [Fact]
    public void ALineAddedBetweenTwoIdenticalOnesLandsInTheRightPlace()
    {
        // The case a character-wise comparison gets wrong if it is not backed off to whole
        // lines: the shared tail reaches back through text identical to the inserted line.
        var was = Module("Public Sub One()", "    x = 1", "    x = 1", "End Sub");
        var now = Module("Public Sub One()", "    x = 1", "    x = 1", "    x = 1", "End Sub");

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(5, diff.TotalLines);
    }

    [Fact]
    public void ALineRemovedIsAWindowThatInsertsNothing()
    {
        var was = Module("Public Sub One()", "    x = 1", "    y = 2", "End Sub");
        var now = Module("Public Sub One()", "    y = 2", "End Sub");

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(3, diff.TotalLines);
    }

    [Fact]
    public void AppendingAtTheEndDoesNotTouchWhatCameBefore()
    {
        var was = Module("Public Sub One()", "End Sub");
        var now = Module("Public Sub One()", "End Sub", "", "Public Sub Two()", "End Sub");

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(5, diff.TotalLines);
    }

    [Fact]
    public void InsertingAtTheTopDoesNotTouchWhatComesAfter()
    {
        var was = Module("Public Sub One()", "End Sub");
        var now = Module("Option Explicit", "", "Public Sub One()", "End Sub");

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(1, diff.At);
        Assert.Equal(now, Apply(was, diff));
    }

    [Fact]
    public void TheTwoSpellingsOfALineEndingAreTheSameText()
    {
        // The surface sends LF and the editor stores CRLF. Reading that as a change would
        // rewrite every line of every module on the first write.
        var stored = "Public Sub One()\r\n    x = 1\r\nEnd Sub";
        var typed = "Public Sub One()\n    x = 1\nEnd Sub";

        Assert.Equal(LineChange.Identical, LineDiff.Between(stored, stored, Window).Change);

        var diff = LineDiff.Between(stored, typed, Window);
        Assert.NotEqual(LineChange.Wholesale, diff.Change);
        if (diff.Change == LineChange.Window)
        {
            // Whatever window it chooses, what lands must be the text that was typed.
            Assert.Equal(typed.ReplaceLineEndings("\r\n"), Apply(stored, diff).ReplaceLineEndings("\r\n"));
        }
    }

    [Fact]
    public void AChangeTooBigToBeAnEditIsNotOne()
    {
        var was = Module([.. Enumerable.Range(0, 900).Select(one => $"    x = {one}")]);
        var now = Module([.. Enumerable.Range(0, 900).Select(one => $"    y = {one}")]);

        Assert.Equal(LineChange.Wholesale, LineDiff.Between(was, now, Window).Change);
    }

    /// <summary>Applies windows the way the shim does: from the last to the first.</summary>
    private static string ApplyAll(string baseline, IReadOnlyList<LineDiff> windows)
    {
        var text = baseline;
        for (var i = windows.Count - 1; i >= 0; i--)
        {
            text = Apply(text, windows[i]);
        }

        return text;
    }

    [Fact]
    public void TwoEditsFarApartAreTwoWindowsNotAReplacement()
    {
        // The shape a rename, or a declaration and its use, produces: 64,802 lines, two of them
        // changed, 64,600 lines apart. One window cannot hold that; two can.
        var lines = Enumerable.Range(0, 64_802).Select(one => $"    x = {one}").ToArray();
        var was = Module(lines);
        lines[100] = "    x = near the top";
        lines[64_700] = "    x = near the end";
        var now = Module(lines);

        Assert.Equal(LineChange.Wholesale, LineDiff.Between(was, now, Window).Change);

        var windows = LineDiff.Windows(was, now, Window, 32);
        Assert.NotNull(windows);
        Assert.Equal(2, windows.Count);
        Assert.Equal((101, 1, 1, "    x = near the top", "    x = 100"), (windows[0].At, windows[0].Removing, windows[0].Inserting, windows[0].Text, windows[0].Removed));
        Assert.Equal((64_701, 1, 1, "    x = near the end", "    x = 64700"), (windows[1].At, windows[1].Removing, windows[1].Inserting, windows[1].Text, windows[1].Removed));
        Assert.All(windows, one => Assert.Equal(64_802, one.TotalLines));
        Assert.Equal(now, ApplyAll(was, windows));
    }

    [Fact]
    public void ADeclarationAtTheTopAndItsUseAtTheBottomAreTwoInsertions()
    {
        var lines = Enumerable.Range(0, 2_000).Select(one => $"    x = {one}").ToList();
        var was = Module([.. lines]);
        lines.Insert(1, "    Dim total As Long");
        lines.Add("    total = x");
        var now = Module([.. lines]);

        var windows = LineDiff.Windows(was, now, Window, 32);
        Assert.NotNull(windows);
        Assert.Equal(2, windows.Count);
        Assert.Equal((2, 0, 1, "    Dim total As Long"), (windows[0].At, windows[0].Removing, windows[0].Inserting, windows[0].Text));
        Assert.Equal((2_001, 0, 1, "    total = x"), (windows[1].At, windows[1].Removing, windows[1].Inserting, windows[1].Text));
        Assert.All(windows, one => Assert.Equal(2_002, one.TotalLines));
        Assert.Equal(now, ApplyAll(was, windows));
    }

    [Fact]
    public void RepeatedLinesBetweenTheEditsDoNotAnchorAndDoNotConfuse()
    {
        // Real modules are mostly "End If", blank lines and "End Sub": none of those occurs once,
        // so none anchors, and the procedure headers do the work instead.
        var lines = new List<string>();
        for (var procedure = 0; procedure < 300; procedure++)
        {
            lines.Add($"Public Sub Proc{procedure}()");
            lines.Add("    If x > 0 Then");
            lines.Add("        y = 1");
            lines.Add("    End If");
            lines.Add("End Sub");
            lines.Add("");
        }

        var was = Module([.. lines]);
        lines[2] = "        y = 2";
        lines[lines.Count - 4] = "        y = 3";
        lines.RemoveAt(900);
        var now = Module([.. lines]);

        var windows = LineDiff.Windows(was, now, Window, 32);
        Assert.NotNull(windows);
        Assert.Equal(3, windows.Count);
        Assert.Equal(now, ApplyAll(was, windows));
        Assert.All(windows, one => Assert.True(one.Removing <= 1 && one.Inserting <= 1));
    }

    [Fact]
    public void TooManyScatteredEditsAreStillAReplacement()
    {
        var lines = Enumerable.Range(0, 900).Select(one => $"    x = {one}").ToArray();
        var was = Module(lines);
        for (var i = 0; i < 900; i += 10)
        {
            lines[i] = $"    y = {i}";
        }

        var now = Module(lines);

        Assert.Null(LineDiff.Windows(was, now, Window, 32));
        Assert.NotNull(LineDiff.Windows(was, now, Window, 100));
    }

    [Fact]
    public void AnUnsplittableMiddleIsStillAReplacement()
    {
        // Nothing occurs once on both sides, so no anchor exists and the one honest answer stands.
        var was = Module([.. Enumerable.Range(0, 900).Select(one => $"    x = {one}")]);
        var now = Module([.. Enumerable.Range(0, 900).Select(one => $"    y = {one}")]);

        Assert.Null(LineDiff.Windows(was, now, Window, 32));
    }

    [Fact]
    public void OneWindowAndNoChangeAnswerAsBetweenDoes()
    {
        var was = Module("Option Explicit", "", "Public Sub One()", "    x = 1", "End Sub");
        var now = Module("Option Explicit", "", "Public Sub One()", "    x = 2", "End Sub");

        Assert.Empty(LineDiff.Windows(was, was, Window, 32)!);
        var one = Assert.Single(LineDiff.Windows(was, now, Window, 32)!);
        Assert.Equal((4, 1, 1, "    x = 2"), (one.At, one.Removing, one.Inserting, one.Text));
    }

    [Fact]
    public void APasteOfThousandsOfLinesIsOneWindowUnderTheShimsCap()
    {
        // The shim's cap is 5,000 lines (2026-09-09): a 3,000-line paste into the middle of a
        // large module is one insert, where at 400 it was the module replaced.
        var lines = Enumerable.Range(0, 20_000).Select(one => $"    x = {one}").ToList();
        var was = Module([.. lines]);
        lines.InsertRange(10_000, Enumerable.Range(0, 3_000).Select(one => $"    ' pasted {one}"));
        var now = Module([.. lines]);

        var diff = LineDiff.Between(was, now, 5_000);

        // One line wider than the least window, at most: the character scan's head and tail
        // overlap inside the baseline when the paste lands between two similar lines, and the
        // line the tail could not claim is rewritten with itself. Correct end to end, and one
        // line is nothing to the editor.
        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(10_001, diff.At);
        Assert.InRange(diff.Removing, 0, 1);
        Assert.Equal(3_000 + diff.Removing, diff.Inserting);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(LineChange.Wholesale, LineDiff.Between(was, now, 400).Change);
    }

    [Fact]
    public void AnEditDeepInALargeModuleIsStillOneLine()
    {
        // The shape this exists for: 64,802 lines, one of them changed.
        var lines = Enumerable.Range(0, 64_802).Select(one => $"    x = {one}").ToArray();
        var was = Module(lines);
        lines[40_000] = "    x = touched";
        var now = Module(lines);

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(40_001, diff.At);
        Assert.Equal(1, diff.Removing);
        Assert.Equal("    x = touched", diff.Text);
        Assert.Equal(64_802, diff.TotalLines);
    }

    [Fact]
    public void ABlankLineAtTheEndIsALineToInsert()
    {
        // A file on disk ends with a line ending, so importing one adds an empty last line. The
        // window's TEXT for that is the empty string, which is indistinguishable from "insert
        // nothing" unless the count says otherwise - and reading it as nothing lost the line:
        // the same import through the dialog and through the route left different modules,
        // which module-sync compares byte for byte (2026-08-21).
        var was = Module("Public Sub One()", "End Sub");
        var now = was + "\r\n";

        var diff = LineDiff.Between(was, now, Window);

        // WHAT LANDS, first and always. The window's exact shape is not pinned here: finding it
        // by characters and finding it by lines pick different but equally valid windows when the
        // two texts disagree about the ending on the boundary line, and either is correct as long
        // as applying it gives the text that was asked for. What must not happen is the line
        // going missing.
        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(3, diff.TotalLines);
        Assert.True(diff.Inserting >= 1, $"an empty line is still a line to insert, got {diff.Inserting}");
        Assert.True(diff.Inserting <= 2, $"one blank line is not a reason to rewrite {diff.Inserting} lines");
    }

    [Fact]
    public void AWindowRunningToTheEndKeepsTheBlankLineAfterIt()
    {
        // The one the two tests around it missed, and the one that actually bit: the window holds
        // content lines AND the empty last line, so the slice's trailing ending separates two
        // lines that are BOTH in the window. Trimming it there dropped the blank line an imported
        // file ends with (2026-08-21).
        var was = Module("Public Sub One()", "End Sub");
        var now = Module("Public Sub One()", "End Sub Renamed") + "\r\n";

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(2, diff.Inserting);

        // Two lines - the renamed one and the empty one after it - joined with CRLF, which is
        // one trailing ending and no more.
        Assert.Equal("End Sub Renamed\r\n", diff.Text);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(3, diff.TotalLines);
    }

    [Fact]
    public void RemovingTheBlankLineAtTheEndInsertsNothing()
    {
        var was = Module("Public Sub One()", "End Sub") + "\r\n";
        var now = Module("Public Sub One()", "End Sub");

        var diff = LineDiff.Between(was, now, Window);

        Assert.Equal(LineChange.Window, diff.Change);
        Assert.Equal(now, Apply(was, diff));
        Assert.Equal(2, diff.TotalLines);
        Assert.True(diff.Removing >= 1, "the blank line has to be removed by something");
        Assert.True(diff.Removing <= 2, $"removing {diff.Removing} lines to drop one blank one");
    }

    [Fact]
    public void EmptyOnEitherSideIsHandledRatherThanAssumed()
    {
        var some = Module("Public Sub One()", "End Sub");

        var filled = LineDiff.Between(string.Empty, some, Window);
        Assert.Equal(LineChange.Window, filled.Change);
        Assert.Equal(1, filled.At);
        Assert.Equal(2, filled.TotalLines);

        var emptied = LineDiff.Between(some, string.Empty, Window);
        Assert.Equal(LineChange.Window, emptied.Change);
        Assert.Equal(1, emptied.At);
    }
}
