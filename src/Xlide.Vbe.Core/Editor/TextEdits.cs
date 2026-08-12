using Xlide.Vbe.Core.Engine;

namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// Rebuilding a document's text from the edits the page reports, rather than from the whole text.
///
/// The page sends the full text with every change while a module is small, and stops above 64,000
/// characters because the message itself becomes the cost. Past that gate this is the only thing
/// that keeps the shim's copy of a module correct, which makes it the thing standing between a
/// developer's typing and a write-back computed against text nobody has.
/// </summary>
public static class TextEdits
{
    /// <summary>
    /// Applies a set of edits, all addressing the SAME original text, and answers the result.
    /// Null when an edit is out of bounds.
    ///
    /// ONE PASS, ONE BUFFER, and the reason is Replace All rather than typing. A keystroke carries
    /// one edit and any implementation is fine for it. A module-scope Replace All carries EVERY
    /// match in a single change - up to the page's cap of 10,000 - and the obvious loop splices the
    /// whole module once per edit. On this repo's Massive fixture (1,493,254 characters, about 3 MB
    /// a copy) that is thousands of full-module copies inside one synchronous handler on the thread
    /// that draws Excel, and the same again on the engine's single thread.
    ///
    /// The edits arrive strictly descending and non-overlapping. The per-edit splice depended on
    /// that too - silently, to be correct at all - so nothing is being assumed here that was not
    /// already assumed; it is only checked now, and a set that breaks the rule takes the old path
    /// instead of being refused. Refusing is the tempting change and the dangerous one: the caller
    /// treats null as "leave the document alone", so a refusal would leave it holding text the page
    /// no longer has, and every later offset would address the wrong string.
    /// </summary>
    public static string? Apply(string text, IReadOnlyList<EngineTextEdit> edits)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(edits);

        if (edits.Count == 0)
        {
            return text;
        }

        // Measured before anything is written: bounds, descending order, no overlap, and the
        // finished length. `previous` is the start of the edit below this one, so an edit may end
        // exactly where that one began and not one character later.
        var length = text.Length;
        var previous = text.Length;
        var orderly = true;

        foreach (var edit in edits)
        {
            ArgumentNullException.ThrowIfNull(edit);

            if (edit.Start < 0 || edit.End < edit.Start || edit.End > text.Length)
            {
                return null;
            }

            if (edit.End > previous)
            {
                orderly = false;
                break;
            }

            length += edit.Text.Length - (edit.End - edit.Start);
            previous = edit.Start;
        }

        if (!orderly || length < 0)
        {
            return OneAtATime(text, edits);
        }

        // Written back to front, the order the edits already arrive in, so every source offset
        // still means what it meant and each span is copied exactly once.
        return string.Create(length, (text, edits), static (span, state) =>
        {
            var (source, set) = state;
            var read = source.Length;
            var write = span.Length;

            foreach (var edit in set)
            {
                var tail = read - edit.End;
                source.AsSpan(edit.End, tail).CopyTo(span[(write - tail)..]);
                write -= tail;

                edit.Text.AsSpan().CopyTo(span[(write - edit.Text.Length)..]);
                write -= edit.Text.Length;

                read = edit.Start;
            }

            source.AsSpan(0, read).CopyTo(span[..read]);
        });
    }

    /// <summary>
    /// The original splice, one whole-text copy per edit, kept for sets the single pass will not
    /// vouch for. Quadratic in the document's size, which is why it is no longer the normal path.
    /// </summary>
    private static string? OneAtATime(string text, IReadOnlyList<EngineTextEdit> edits)
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
}
