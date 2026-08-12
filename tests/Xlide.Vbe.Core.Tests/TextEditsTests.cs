using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Engine;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// Rebuilding a module's text from the page's edits.
///
/// This is the arithmetic standing between a developer's typing and a write-back computed against
/// text nobody has: above 64,000 characters the page stops sending the whole document, and this is
/// the only thing that keeps the shim's copy right. An off-by-one here does not throw, it writes a
/// corrupted module into the workbook.
///
/// The single-pass implementation replaced a loop that spliced the whole string per edit, so most
/// of what these check is that the two agree - including on the sets the fast path refuses and
/// hands back.
/// </summary>
public class TextEditsTests
{
    private static EngineTextEdit Edit(int start, int end, string text) => new(start, end, text);

    /// <summary>The old implementation, kept here as the thing the new one has to agree with.</summary>
    private static string? Spliced(string text, IReadOnlyList<EngineTextEdit> edits)
    {
        var updated = text;

        foreach (var edit in edits)
        {
            if (edit.Start < 0 || edit.End < edit.Start || edit.End > updated.Length)
            {
                return null;
            }

            updated = string.Concat(updated.AsSpan(0, edit.Start), edit.Text, updated.AsSpan(edit.End));
        }

        return updated;
    }

    [Fact]
    public void NoEditsLeavesTheTextAlone() =>
        Assert.Equal("Option Explicit", TextEdits.Apply("Option Explicit", []));

    [Fact]
    public void OneReplacementLandsWhereItWasAsked()
    {
        // "Explicit" -> "Implicit", the shape a single keystroke or one match produces.
        Assert.Equal("Option Implicit", TextEdits.Apply("Option Explicit", [Edit(7, 15, "Implicit")]));
    }

    [Fact]
    public void AnInsertionHasStartEqualToEnd() =>
        Assert.Equal("ab!c", TextEdits.Apply("abc", [Edit(2, 2, "!")]));

    [Fact]
    public void ADeletionCarriesEmptyText() =>
        Assert.Equal("ac", TextEdits.Apply("abc", [Edit(1, 2, "")]));

    [Fact]
    public void AnEditAtTheVeryEndIsInBounds() =>
        Assert.Equal("abcd", TextEdits.Apply("abc", [Edit(3, 3, "d")]));

    [Fact]
    public void DescendingEditsAllApply()
    {
        // Bottom-up, which is the order Monaco reports and the order the page forwards.
        //
        //            a b c d _ e f g h _ i
        // index      0 1 2 3 4 5 6 7 8 9 10
        // replaced   X       Y       Z
        var edits = new[] { Edit(8, 9, "Z"), Edit(4, 5, "Y"), Edit(0, 1, "X") };
        Assert.Equal("XbcdYefgZ i", TextEdits.Apply("abcd efgh i", edits));
    }

    [Fact]
    public void AdjacentEditsAreNotAnOverlap()
    {
        // One edit ending exactly where the next begins is legal and must take the fast path.
        var edits = new[] { Edit(3, 6, "ZZZ"), Edit(0, 3, "AAA") };
        Assert.Equal("AAAZZZ", TextEdits.Apply("abcdef", edits));
    }

    [Theory]
    [InlineData(-1, 2)]
    [InlineData(0, 99)]
    [InlineData(3, 1)]
    public void AnEditOutOfBoundsRefusesTheWholeSet(int start, int end) =>
        Assert.Null(TextEdits.Apply("abcdef", [Edit(start, end, "x")]));

    [Fact]
    public void AnAscendingSetStillProducesWhatItAlwaysDid()
    {
        // Not the order the page sends, so the fast path declines it and the old splice answers.
        // What matters is that the answer is unchanged, because refusing would leave the caller
        // holding text the page does not have.
        var edits = new[] { Edit(0, 1, "X"), Edit(4, 5, "Y") };
        Assert.Equal(Spliced("abcd efgh", edits), TextEdits.Apply("abcd efgh", edits));
    }

    [Fact]
    public void OverlappingEditsStillProduceWhatTheyAlwaysDid()
    {
        var edits = new[] { Edit(2, 6, "LONGER"), Edit(4, 5, "x") };
        Assert.Equal(Spliced("abcdefgh", edits), TextEdits.Apply("abcdefgh", edits));
    }

    /// <summary>
    /// The case the change exists for: one Replace All over a large module, thousands of matches in
    /// a single change. The old path copied the whole text once per match.
    /// </summary>
    [Fact]
    public void AReplaceAllOverAThousandMatchesAgreesWithTheOldSplice()
    {
        const int matches = 1000;
        var source = string.Join("\r\n", Enumerable.Range(0, matches).Select(n => $"    Debug.Print seed{n}"));

        // Every occurrence of "seed", bottom-up, exactly as the page reports them.
        var edits = new List<EngineTextEdit>();
        for (var at = source.LastIndexOf("seed", StringComparison.Ordinal);
             at >= 0;
             at = at == 0 ? -1 : source.LastIndexOf("seed", at - 1, StringComparison.Ordinal))
        {
            edits.Add(Edit(at, at + 4, "sown"));
        }

        Assert.Equal(matches, edits.Count);

        var fast = TextEdits.Apply(source, edits);
        Assert.Equal(Spliced(source, edits), fast);
        Assert.DoesNotContain("seed", fast, StringComparison.Ordinal);
        Assert.Equal(matches, fast!.Split("sown").Length - 1);
    }

    /// <summary>
    /// Random well-formed sets against the old implementation. The single pass computes the final
    /// length up front and writes into it back to front, so a wrong span or a wrong running offset
    /// is a class of bug no hand-written example reliably catches.
    /// </summary>
    [Fact]
    public void RandomDescendingSetsAgreeWithTheOldSplice()
    {
        // Fixed seed: a failure has to be reproducible, and this runs on every gate.
        var random = new Random(20260811);

        for (var round = 0; round < 400; round++)
        {
            var source = new string([.. Enumerable.Range(0, random.Next(1, 120))
                .Select(_ => (char)('a' + random.Next(0, 26)))]);

            var edits = new List<EngineTextEdit>();
            var cursor = source.Length;

            while (cursor > 0)
            {
                var end = random.Next(0, cursor + 1);
                var start = random.Next(0, end + 1);
                if (random.Next(0, 3) == 0)
                {
                    // A gap with no edit in it, so untouched runs are exercised too.
                    cursor = start;
                    continue;
                }

                var inserted = new string('#', random.Next(0, 4));
                edits.Add(Edit(start, end, inserted));
                cursor = start;
            }

            Assert.Equal(Spliced(source, edits), TextEdits.Apply(source, edits));
        }
    }
}
