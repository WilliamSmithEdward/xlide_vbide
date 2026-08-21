namespace Xlide.Vbe.Core.Editor;

/// <summary>What a write turned out to be, once the two texts have been compared.</summary>
public enum LineChange
{
    /// <summary>The module already holds this text. Nothing to write, nothing to check.</summary>
    Identical,

    /// <summary>One run of lines differs; <see cref="LineDiff"/> says which and what goes there.</summary>
    Window,

    /// <summary>Too much moved to call it an edit. The caller replaces the module wholesale.</summary>
    Wholesale,
}

/// <summary>
/// The run of lines that differs between two versions of a module's text, expressed the way the
/// host's editor takes it: a one-based line to start at, how many lines to remove, and the text
/// to put there.
///
/// WHY THIS IS NOT A SPLIT. The obvious way to find this is to split both texts into lines and
/// walk the arrays, and that is what the shim did. On a 64,802-line module each split is 64,802
/// strings and a 518KB array - two of them per write, both past the large object heap threshold,
/// so every pause in typing pushed the host's heap toward a generation-2 collection. Twelve
/// write-backs of one line moved the managed heap from 171MB to 273MB and cost two full
/// collections in six seconds (measured 2026-08-21).
///
/// So the comparison walks the two strings line by line over spans, and allocates only the
/// window it is going to write - a few hundred characters for the edit a person actually made.
///
/// LINE BY LINE, NOT CHARACTER BY CHARACTER, and that is not a detail: the editing surface sends
/// LF and the host's editor stores CRLF, so a comparison that read the ending as content would
/// find every line changed and rewrite whole modules - which is the ten to seventeen second
/// freeze this is all trying to avoid.
/// </summary>
/// <param name="Change">Whether there is anything to write, and in what shape.</param>
/// <param name="At">One-based line the window starts at.</param>
/// <param name="Removing">How many of the module's lines the window replaces.</param>
/// <param name="Inserting">
/// How many lines go there. NOT the same question as whether <paramref name="Text"/> is empty:
/// one empty line is a line, and reading it as nothing to insert loses the blank line at the end
/// of a text that ends with an ending - which is how most files on disk end (2026-08-21, caught
/// by module-sync comparing an import through the dialog against the same import through the
/// route, byte for byte).
/// </param>
/// <param name="Text">The lines to insert, joined with CRLF. Empty when the window only removes.</param>
/// <param name="Removed">What those lines held before, so a refused insert can be put back.</param>
/// <param name="TotalLines">How many lines the module should hold once the window is written.</param>
public readonly record struct LineDiff(
    LineChange Change,
    int At,
    int Removing,
    int Inserting,
    string Text,
    string Removed,
    int TotalLines)
{
    /// <summary>
    /// Compares two versions of a module's text without splitting either.
    /// </summary>
    /// <param name="largestWindow">
    /// Past this many lines on either side the change is not an edit any more, and replacing the
    /// module outright is both simpler and no slower.
    /// </param>
    public static LineDiff Between(string baseline, string text, int largestWindow)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(text);

        // THE COMMON CASE FIRST, and it is common: a write-back whose text has not moved is what
        // every pause in typing produces once the module is already up to date. One vectorised
        // comparison answers it, where walking the lines to reach the same conclusion costs two
        // passes over both texts - 1.49MB each - and put 35ms back on every write (measured
        // 2026-08-21, after the walk replaced the split).
        if (string.Equals(baseline, text, StringComparison.Ordinal))
        {
            return new LineDiff(LineChange.Identical, 0, 0, 0, string.Empty, string.Empty, CountLines(text));
        }

        // MOST WRITES ARE ONE LINE IN A LARGE MODULE, and finding that by walking lines is a
        // step per unchanged line: 64,700 of them for an edit near the top, which measured
        // 123-187ms of host thread per write. The scan below finds the same window with
        // vectorised comparisons over the whole string and no per-line step at all.
        //
        // It can be fooled, and that is why it is allowed to decline: the two texts may spell
        // their line endings differently - the surface sends LF, the editor stores CRLF - and a
        // character comparison then finds nothing in common. It answers null in that case, and
        // the line walk below, which reads endings as endings, gives the real answer.
        if (ByCharacters(baseline, text, largestWindow) is { } quick)
        {
            return quick;
        }

        var oldTotal = CountLines(baseline);
        var newTotal = CountLines(text);

        // The shared head, a line at a time. `oldAt` and `newAt` come out holding the offset of
        // the first line that differs, which is where the window begins.
        var oldAt = 0;
        var newAt = 0;
        var prefix = 0;
        while (prefix < oldTotal && prefix < newTotal
            && TryLine(baseline, oldAt, out var oldStart, out var oldEnd, out var oldNext)
            && TryLine(text, newAt, out var newStart, out var newEnd, out var newNext)
            && baseline.AsSpan(oldStart, oldEnd - oldStart).SequenceEqual(text.AsSpan(newStart, newEnd - newStart)))
        {
            oldAt = oldNext;
            newAt = newNext;
            prefix++;
        }

        if (prefix == oldTotal && prefix == newTotal)
        {
            return new LineDiff(LineChange.Identical, 0, 0, 0, string.Empty, string.Empty, newTotal);
        }

        // And the shared tail, from the back, never crossing into the head.
        var oldBlock = baseline.Length;
        var newBlock = text.Length;
        var oldTailStart = baseline.Length;
        var newTailStart = text.Length;
        var suffix = 0;
        while (prefix + suffix < oldTotal && prefix + suffix < newTotal
            && TryLineBack(baseline, oldBlock, out var oldStart, out var oldEnd, out var oldPrevious)
            && TryLineBack(text, newBlock, out var newStart, out var newEnd, out var newPrevious)
            && baseline.AsSpan(oldStart, oldEnd - oldStart).SequenceEqual(text.AsSpan(newStart, newEnd - newStart)))
        {
            oldBlock = oldPrevious;
            newBlock = newPrevious;
            oldTailStart = oldStart;
            newTailStart = newStart;
            suffix++;
        }

        var removing = oldTotal - prefix - suffix;
        var inserting = newTotal - prefix - suffix;

        if (removing > largestWindow || inserting > largestWindow)
        {
            return new LineDiff(LineChange.Wholesale, 0, 0, 0, string.Empty, string.Empty, newTotal);
        }

        return new LineDiff(
            LineChange.Window,
            prefix + 1,
            removing,
            inserting,
            inserting == 0 ? string.Empty : Crlf(text[newAt..newTailStart], suffix > 0),
            removing == 0 ? string.Empty : Crlf(baseline[oldAt..oldTailStart], suffix > 0),
            newTotal);
    }


    /// <summary>
    /// The window found by comparing CHARACTERS, backed off to whole lines - or null when what
    /// that finds is too big to be an edit, which is both the genuinely-wholesale case and the
    /// case where the two texts disagree about how a line ends.
    /// </summary>
    private static LineDiff? ByCharacters(string baseline, string text, int largestWindow)
    {
        var head = baseline.AsSpan().CommonPrefixLength(text.AsSpan());

        // The shared tail, by binary search over vectorised comparisons: "the last k characters
        // match" is true for every k below the answer and false above it, so this costs a
        // handful of passes rather than a step per character.
        var lo = 0;
        var hi = Math.Min(baseline.Length, text.Length) - head;
        while (lo < hi)
        {
            var mid = lo + ((hi - lo + 1) / 2);
            if (baseline.AsSpan(baseline.Length - mid).SequenceEqual(text.AsSpan(text.Length - mid)))
            {
                lo = mid;
            }
            else
            {
                hi = mid - 1;
            }
        }

        var tail = lo;

        // Back off to whole lines. Everything before the head is shared, so the line it lands in
        // starts the same distance back in both - the shorter of the two, when one of them began
        // at the very start.
        var oldStart = LineStartAt(baseline, head);
        var newStart = LineStartAt(text, head);
        var shared = Math.Min(head - oldStart, head - newStart);
        oldStart = head - shared;
        newStart = head - shared;

        // And forward to the line boundary the shared tail begins on: a tail that starts inside a
        // line leaves that line partly changed, so the whole of it belongs to the window.
        var oldEnd = Math.Max(oldStart, LineEndAt(baseline, baseline.Length - tail));
        var newEnd = Math.Max(newStart, LineEndAt(text, text.Length - tail));

        var removing = LinesIn(baseline, oldStart, oldEnd);
        var inserting = LinesIn(text, newStart, newEnd);
        if (removing > largestWindow || inserting > largestWindow)
        {
            return null;
        }

        var separated = newEnd < text.Length;
        return new LineDiff(
            LineChange.Window,
            NewlinesIn(text, 0, newStart) + 1,
            removing,
            inserting,
            inserting == 0 ? string.Empty : Crlf(text[newStart..newEnd], separated),
            removing == 0 ? string.Empty : Crlf(baseline[oldStart..oldEnd], oldEnd < baseline.Length),
            CountLines(text));
    }

    /// <summary>Where the line containing an offset begins.</summary>
    private static int LineStartAt(string text, int at)
    {
        if (at <= 0)
        {
            return 0;
        }

        var ending = text.LastIndexOf('\n', Math.Min(at, text.Length) - 1);
        return ending < 0 ? 0 : ending + 1;
    }

    /// <summary>Where the line containing an offset ends, meaning where the next one begins.</summary>
    private static int LineEndAt(string text, int at)
    {
        if (at >= text.Length)
        {
            return text.Length;
        }

        var ending = text.IndexOf('\n', Math.Max(at, 0));
        return ending < 0 ? text.Length : ending + 1;
    }

    /// <summary>
    /// Lines between two line boundaries. A run that reaches the end of the text carries one more
    /// than its endings, because the text's last line has no ending after it - the same
    /// convention <see cref="CountLines"/> counts by.
    /// </summary>
    private static int LinesIn(string text, int from, int to) =>
        NewlinesIn(text, from, to) + (to == text.Length ? 1 : 0);

    private static int NewlinesIn(string text, int from, int to) =>
        to <= from ? 0 : text.AsSpan(from, to - from).Count('\n');
    /// <summary>
    /// The line beginning at <paramref name="at"/>, as the half-open range of its content with
    /// the ending left off, and where the next line begins. False when the offset is past the
    /// end - a text ending in a newline has one more line after it, and that line is empty.
    /// </summary>
    private static bool TryLine(string text, int at, out int start, out int end, out int next)
    {
        start = end = next = -1;
        if (at > text.Length)
        {
            return false;
        }

        start = at;
        var ending = text.IndexOf('\n', at);
        if (ending < 0)
        {
            end = text.Length;
            next = text.Length + 1;
            return true;
        }

        end = ending > at && text[ending - 1] == '\r' ? ending - 1 : ending;
        next = ending + 1;
        return true;
    }

    /// <summary>
    /// The last line of <c>text[0..blockEnd)</c>, and where the block before it ends - with that
    /// line's ending left off, so the next call reads the line before rather than an empty one.
    /// </summary>
    private static bool TryLineBack(string text, int blockEnd, out int start, out int end, out int previous)
    {
        start = end = previous = -1;
        if (blockEnd <= 0)
        {
            return false;
        }

        if (text[blockEnd - 1] == '\n')
        {
            // The block ends with an ending, so its last line is the empty one after it.
            start = blockEnd;
            end = blockEnd;
            previous = blockEnd - 1;
            if (previous > 0 && text[previous - 1] == '\r')
            {
                previous--;
            }

            return true;
        }

        var ending = text.LastIndexOf('\n', blockEnd - 1);
        start = ending + 1;
        end = blockEnd;
        previous = ending < 0 ? 0 : ending;
        if (previous > 0 && text[previous - 1] == '\r')
        {
            previous--;
        }

        return true;
    }

    /// <summary>
    /// How many lines a text is, counted the way splitting it would: a trailing ending leaves an
    /// empty line after it, and an empty text is one empty line.
    ///
    /// Public because it is also how a whole-module write knows what the editor added to it.
    /// </summary>
    public static int CountLines(string text)
    {
        var lines = 1;
        foreach (var one in text)
        {
            if (one == '\n')
            {
                lines++;
            }
        }

        return lines;
    }

    /// <summary>
    /// A window as the host's editor takes it: CRLF between lines, and nothing after the last
    /// one - InsertLines reads a trailing ending as one more empty line.
    /// </summary>
    /// <param name="separated">
    /// Whether a line follows the window in the text it was cut from. When one does, the slice
    /// ends with the ending that SEPARATES them, and that ending is not part of the window. When
    /// none does, the slice runs to the end of the text - and a trailing ending there is the
    /// separator before the empty LAST line, which IS part of the window. Trimming it in that
    /// case cost the blank line at the end of every imported file: the same import through the
    /// dialog and through the route left modules differing by two characters, which is what
    /// module-sync compares (2026-08-21).
    /// </param>
    private static string Crlf(string window, bool separated)
    {
        var trimmed = window;
        if (separated && trimmed.EndsWith('\n'))
        {
            trimmed = trimmed[..^1];
            if (trimmed.EndsWith('\r'))
            {
                trimmed = trimmed[..^1];
            }
        }

        return trimmed.ReplaceLineEndings("\r\n");
    }
}
