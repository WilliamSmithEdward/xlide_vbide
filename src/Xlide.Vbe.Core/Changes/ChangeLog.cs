using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Xlide.Vbe.Core.Changes;

/// <summary>What happened to a module.</summary>
public enum ChangeKind
{
    /// <summary>Its code was written.</summary>
    Written,

    /// <summary>The module came into the project.</summary>
    Added,

    /// <summary>The module left it.</summary>
    Removed,

    /// <summary>The module is the same one under a different name.</summary>
    Renamed,
}

/// <summary>One module's before and after within a round.</summary>
/// <param name="Module">The module's name at the time.</param>
/// <param name="Kind">What happened to it.</param>
/// <param name="Before">Key of the text it held when the round began, or null when it did not exist.</param>
/// <param name="After">Key of the text it held when the round ended, or null when it no longer exists.</param>
/// <param name="From">What it was called when the round began, when that is not what it is called now.</param>
public sealed record ChangeEntry(
    string Module, ChangeKind Kind, string? Before, string? After, string? From = null);

/// <summary>
/// One round: a stretch of writes by one author, ended by a snapshot, by the author changing, by
/// a save, or by a long enough silence.
/// </summary>
public sealed record ChangeRound(
    int Number,
    DateTimeOffset Started,
    DateTimeOffset Ended,
    string Author,
    string? Label,
    bool Open,
    bool AcceptedHere,
    IReadOnlyList<ChangeEntry> Entries);

/// <summary>
/// A record of what happened to a project's module code, by whom, in rounds.
///
/// IN MEMORY, AND IT WRITES NOTHING ANYWHERE. It began as a file under the product's own data
/// directory, which was wrong for a reason worth keeping written down: nothing in production may
/// depend on an external log file (the owner, 2026-08-22). A pane that only works while a file it
/// wrote is still on disk has a second way to be wrong, and a shipped feature reading what is
/// otherwise a developer artifact is a category error however tidy the directory looks. So the log
/// lives in the session and goes when the session does - which is the span it is for: reviewing
/// what an agent is doing while it does it.
///
/// IT NEVER WRITES TO THE WORKBOOK EITHER. It says what changed and holds the text from before so
/// it can be read; every operation that could put text back belongs to whoever is reading it - the
/// developer through the editor, an agent through the write route. That is the whole reason it can
/// be trusted: nothing in here can lose work, because nothing in here writes any.
///
/// WHAT IT COSTS is one copy of a module's text the first time a round touches it, and nothing at
/// all for modules nobody edited. Texts are kept by content, so a module written five times in one
/// round costs one before and one after, and an after that becomes the next round's before is one
/// copy between them. Past <see cref="LargestHeldBytes"/> the oldest rounds let their texts go and
/// say so, which is a log ageing rather than a session growing without bound.
/// </summary>
public sealed class ChangeLog
{
    /// <summary>The author recorded when a caller did not say who it was.</summary>
    public const string Unattributed = "unattributed";

    /// <summary>
    /// How long a silence ends a round when nobody says so. Long enough that a developer thinking
    /// between edits stays in one round, short enough that a morning's work is not one round.
    /// </summary>
    public static readonly TimeSpan Silence = TimeSpan.FromMinutes(5);

    /// <summary>
    /// How much module text one project's log holds before the oldest rounds let theirs go.
    ///
    /// Nothing evicts this for us now that it is memory rather than disk, and the texts are the
    /// only part with any size to them: at VBA's per-module ceiling one copy is about 1.5 MB, so
    /// 64 covers a long session over a large module without being able to run a host out of
    /// memory. A round whose texts have gone keeps its entries and says the text is no longer
    /// held, which the pane draws rather than hiding.
    /// </summary>
    public const long LargestHeldBytes = 64L * 1024 * 1024;

    private readonly List<ChangeRound> _rounds = [];
    private readonly Dictionary<string, ChangeEntry> _open = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _texts = new(StringComparer.Ordinal);

    private long _held;
    private int _number;
    private string _author = Unattributed;
    private string? _label;
    private DateTimeOffset _started;
    private DateTimeOffset _touched;
    private bool _running;

    /// <summary>The newest round marked reviewed, or zero when none has been.</summary>
    public int AcceptedAt { get; private set; }

    /// <summary>How much module text this log is holding, in bytes.</summary>
    public long HeldBytes => _held;

    /// <summary>
    /// Records what a write did to a module, opening a round when one is not already running.
    ///
    /// The FIRST time a round touches a module its previous text is kept, and later writes in the
    /// same round only move the after. So a round says what it did to a module, once, however many
    /// times an agent rewrote it while thinking.
    /// </summary>
    public void Record(
        string module, ChangeKind kind, string? before, string? after, string? by, DateTimeOffset now,
        string? from = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(module);

        var author = string.IsNullOrWhiteSpace(by) ? Unattributed : by.Trim();

        // A DIFFERENT HAND IS A DIFFERENT ROUND. Agent, then the developer by hand, then the agent
        // again is three rounds and not two - which is what lets someone read the log and see
        // their own edit sitting between two of an agent's, rather than folded into either.
        if (_running && (!string.Equals(_author, author, StringComparison.OrdinalIgnoreCase)
            || now - _touched > Silence))
        {
            Close(null, now);
        }

        if (!_running)
        {
            _number++;
            _author = author;
            _label = null;
            _started = now;
            _running = true;
            _open.Clear();
        }

        _touched = now;

        // A RENAME MOVES THE ENTRY RATHER THAN STARTING ANOTHER. The round's entries are keyed by
        // the module's name, and a rename changes it - so without this a module renamed and then
        // written reads as two modules, one of which no longer exists. The entry moves to the new
        // name and remembers what it was called when the round began, which is the honest answer
        // to "what happened to Ledger": it is called Accounts now.
        if (kind == ChangeKind.Renamed && from is { Length: > 0 } && _open.Remove(from, out var moved))
        {
            _open[module] = moved with { Module = module, From = moved.From ?? from };
            return;
        }

        // THE FIRST BEFORE THAT IS KNOWN - and knowing the difference between an absence and a
        // gap. A module that ARRIVED in this round had no text before it, and that null is the
        // answer rather than something to fill in. A RENAME's null is a gap: a rename carries no
        // text of its own, and a removal later in the same round used to inherit that emptiness
        // instead of recording what the module actually held - which is the one text a removal
        // can never get back afterwards.
        _open.TryGetValue(module, out var already);
        var beforeKey = already switch
        {
            null => Keep(before),
            { Kind: ChangeKind.Added } => null,
            _ => already.Before ?? Keep(before),
        };
        var afterKey = Keep(after);
        var was = already?.Kind ?? kind;

        // Added then written in one round is still an add; written then removed is a removal; and
        // a module renamed and then written is still the one that was renamed.
        var settled = kind == ChangeKind.Removed ? ChangeKind.Removed
            : was == ChangeKind.Added ? ChangeKind.Added
            : was == ChangeKind.Renamed ? ChangeKind.Renamed
            : kind;

        _open[module] = new ChangeEntry(
            module, settled, beforeKey, afterKey,
            already?.From ?? (kind == ChangeKind.Renamed ? from : null));
    }

    /// <summary>
    /// Ends the round that is running, with a label if the caller has one. Does nothing when no
    /// round is running, so a caller may say "that was one round" as often as it likes.
    /// </summary>
    public void Close(string? label, DateTimeOffset now)
    {
        if (!_running)
        {
            return;
        }

        _label = string.IsNullOrWhiteSpace(label) ? _label : label.Trim();
        _rounds.Add(new ChangeRound(
            _number, _started, now, _author, _label, false, false, [.. _open.Values]));

        _running = false;
        _open.Clear();
        Age();
    }

    /// <summary>
    /// Marks every round so far as reviewed. None of them are removed - a log that deletes its own
    /// past is not one - so what this changes is where a reader starts counting from.
    /// </summary>
    public void Accept(DateTimeOffset now)
    {
        Close(null, now);
        AcceptedAt = _number;
    }

    /// <summary>Every round, newest first, with the one still running - if any - at the front.</summary>
    public IReadOnlyList<ChangeRound> Rounds(int limit = 200)
    {
        var all = new List<ChangeRound>();
        if (_running)
        {
            all.Add(new ChangeRound(
                _number, _started, _touched, _author, _label, true, false, [.. _open.Values]));
        }

        for (var at = _rounds.Count - 1; at >= 0 && all.Count < limit; at--)
        {
            all.Add(_rounds[at]);
        }

        return all;
    }

    /// <summary>A text this log holds, or null when it never held one or has let it go.</summary>
    public string? TextOf(string? key) =>
        key is { Length: > 0 } && _texts.TryGetValue(key, out var held) ? held : null;

    /// <summary>Keeps a text under a key of its own content, and answers that key.</summary>
    private string? Keep(string? text)
    {
        if (text is null)
        {
            return null;
        }

        var key = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();
        if (_texts.TryAdd(key, text))
        {
            _held += text.Length * 2L;
        }

        return key;
    }

    /// <summary>
    /// The oldest rounds let their texts go, until what is held is back inside the budget.
    ///
    /// The ROUNDS stay: they are the record, and they cost almost nothing. What goes is the text,
    /// which is the only part with size - and an entry whose text has gone says so rather than
    /// drawing an empty comparison, because a log that quietly stops being able to show something
    /// is worse than one that says it cannot.
    /// </summary>
    private void Age()
    {
        var oldest = 0;
        while (_held > LargestHeldBytes && oldest < _rounds.Count - 1)
        {
            oldest++;

            var live = new HashSet<string>(StringComparer.Ordinal);
            for (var at = oldest; at < _rounds.Count; at++)
            {
                foreach (var entry in _rounds[at].Entries)
                {
                    if (entry.Before is { Length: > 0 } before) { live.Add(before); }
                    if (entry.After is { Length: > 0 } after) { live.Add(after); }
                }
            }

            foreach (var entry in _open.Values)
            {
                if (entry.Before is { Length: > 0 } before) { live.Add(before); }
                if (entry.After is { Length: > 0 } after) { live.Add(after); }
            }

            foreach (var key in _texts.Keys.Where(key => !live.Contains(key)).ToList())
            {
                _held -= _texts[key].Length * 2L;
                _texts.Remove(key);
            }
        }
    }
}
