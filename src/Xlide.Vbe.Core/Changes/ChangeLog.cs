using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

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
}

/// <summary>One module's before and after within a round.</summary>
/// <param name="Module">The module's name at the time.</param>
/// <param name="Kind">What happened to it.</param>
/// <param name="Before">Hash of the text it held when the round began, or null when it did not exist.</param>
/// <param name="After">Hash of the text it held when the round ended, or null when it no longer exists.</param>
public sealed record ChangeEntry(string Module, ChangeKind Kind, string? Before, string? After);

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
/// APPEND-ONLY, AND IT NEVER WRITES TO THE WORKBOOK. It is a log, not a version control system:
/// it says what changed and holds the text from before so it can be read, and every operation
/// that could put text back belongs to whoever is reading it - the developer through the editor,
/// an agent through the write route. That is the whole reason it can be trusted: nothing in here
/// can lose work, because nothing in here writes any.
///
/// WHERE IT LIVES follows from what it is. A record of what happened on THIS machine belongs
/// beside the diagnostic log, not inside the workbook (which it would bloat, and which VBA can
/// see) and not in a folder next to it (which a network share or a rename makes a mess of). It
/// survives closing the workbook because a log that forgets is not a log.
///
/// WHAT IT COSTS is one copy of a module's text the first time a round touches it, and nothing at
/// all for modules nobody edited. Texts are stored by content hash, so a module written five
/// times in one round costs one before and one after, and an after that becomes the next round's
/// before is stored once between them.
/// </summary>
public sealed class ChangeLog
{
    /// <summary>The author recorded when a caller did not say who it was.</summary>
    public const string Unattributed = "unattributed";

    /// <summary>
    /// How long a silence ends a round when nobody says so. Long enough that a developer thinking
    /// between edits stays in one round, short enough that yesterday's work is not in today's.
    /// </summary>
    public static readonly TimeSpan Silence = TimeSpan.FromMinutes(5);

    private readonly string _root;
    private readonly string _texts;
    private readonly string _path;

    private readonly List<ChangeRound> _rounds = [];
    private readonly Dictionary<string, ChangeEntry> _open = new(StringComparer.OrdinalIgnoreCase);

    private int _number;
    private string _author = Unattributed;
    private string? _label;
    private DateTimeOffset _started;
    private DateTimeOffset _touched;
    private bool _running;

    private ChangeLog(string root)
    {
        _root = root;
        _texts = Path.Combine(root, "texts");
        _path = Path.Combine(root, "log.jsonl");
    }

    /// <summary>
    /// The log for one project, under <paramref name="home"/>.
    ///
    /// The directory is named for the workbook so a developer can find it, and carries a hash of
    /// its full path so two workbooks with the same file name in different folders are two logs.
    /// </summary>
    public static ChangeLog For(string home, string projectId)
    {
        ArgumentException.ThrowIfNullOrEmpty(home);
        ArgumentException.ThrowIfNullOrEmpty(projectId);

        var name = Path.GetFileNameWithoutExtension(projectId);
        if (string.IsNullOrEmpty(name))
        {
            name = "project";
        }

        foreach (var bad in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(bad, '-');
        }

        var stamp = Convert.ToHexString(
            SHA256.HashData(Encoding.UTF8.GetBytes(projectId.ToLowerInvariant())))[..8].ToLowerInvariant();

        var log = new ChangeLog(Path.Combine(home, "changes", $"{name}-{stamp}"));
        log.Load();
        return log;
    }

    /// <summary>Where this project's log is kept, for anyone who wants to look at the file.</summary>
    public string Directory => _root;

    /// <summary>
    /// Records what a write did to a module, opening a round when one is not already running.
    ///
    /// The FIRST time a round touches a module its previous text is kept, and later writes in the
    /// same round only move the after. So a round says what it did to a module, once, however
    /// many times an agent rewrote it while thinking.
    /// </summary>
    public void Record(string module, ChangeKind kind, string? before, string? after, string? by, DateTimeOffset now)
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
            Append(writer =>
            {
                writer.WriteString("k", "open");
                writer.WriteNumber("n", _number);
                writer.WriteString("at", Iso(now));
                writer.WriteString("by", author);
            });
        }

        _touched = now;

        var beforeHash = _open.TryGetValue(module, out var already) ? already.Before : Keep(before);
        var afterHash = Keep(after);
        var was = already?.Kind ?? kind;

        // Added then written in one round is still an add; written then removed is a removal.
        var settled = kind == ChangeKind.Removed ? ChangeKind.Removed
            : was == ChangeKind.Added ? ChangeKind.Added
            : kind;

        _open[module] = new ChangeEntry(module, settled, beforeHash, afterHash);

        Append(writer =>
        {
            writer.WriteString("k", "write");
            writer.WriteNumber("n", _number);
            writer.WriteString("at", Iso(now));
            writer.WriteString("module", module);
            writer.WriteString("kind", settled.ToString().ToLowerInvariant());
            WriteHash(writer, "before", beforeHash);
            WriteHash(writer, "after", afterHash);
        });
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

        Append(writer =>
        {
            writer.WriteString("k", "close");
            writer.WriteNumber("n", _number);
            writer.WriteString("at", Iso(now));
            if (_label is { Length: > 0 } said)
            {
                writer.WriteString("label", said);
            }
        });
    }

    /// <summary>
    /// Marks every round so far as reviewed. None of them are removed - a log that deletes its
    /// own past is not one - so what this changes is where a reader starts counting from.
    /// </summary>
    public void Accept(DateTimeOffset now)
    {
        Close(null, now);
        Append(writer =>
        {
            writer.WriteString("k", "accept");
            writer.WriteNumber("n", _number);
            writer.WriteString("at", Iso(now));
        });

        AcceptedAt = _number;
    }

    /// <summary>The newest round marked reviewed, or zero when none has been.</summary>
    public int AcceptedAt { get; private set; }

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

    /// <summary>A text this log kept, or null when it was never held or has been let go.</summary>
    public string? TextOf(string? hash)
    {
        if (string.IsNullOrEmpty(hash))
        {
            return null;
        }

        try
        {
            var file = Path.Combine(_texts, $"{hash}.vba");
            return File.Exists(file) ? File.ReadAllText(file, Encoding.UTF8) : null;
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
    }

    // ---- keeping the texts ------------------------------------------------------------------

    /// <summary>Stores a text under its own hash and answers the hash. Null text, null hash.</summary>
    private string? Keep(string? text)
    {
        if (text is null)
        {
            return null;
        }

        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();

        try
        {
            System.IO.Directory.CreateDirectory(_texts);
            var file = Path.Combine(_texts, $"{hash}.vba");
            if (!File.Exists(file))
            {
                File.WriteAllText(file, text, Encoding.UTF8);
            }
        }
        catch (IOException)
        {
            // The hash is still the truth about what the text WAS; only the copy is missing, and
            // TextOf says so by answering null. A log that cannot store a text must not take the
            // write down with it.
        }
        catch (UnauthorizedAccessException)
        {
        }

        return hash;
    }

    private static void WriteHash(Utf8JsonWriter writer, string name, string? hash)
    {
        if (hash is { Length: > 0 })
        {
            writer.WriteString(name, hash);
        }
    }

    private static string Iso(DateTimeOffset when) =>
        when.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);

    /// <summary>
    /// Adds one line to the log.
    ///
    /// Every line stands alone and is written the moment the thing happened, so a session that
    /// ends badly still leaves a readable record: a round with an `open` and no `close` reads as
    /// the round that was running, which is exactly what it was.
    /// </summary>
    private void Append(Action<Utf8JsonWriter> body)
    {
        try
        {
            System.IO.Directory.CreateDirectory(_root);

            using var buffer = new MemoryStream();
            using (var writer = new Utf8JsonWriter(buffer))
            {
                writer.WriteStartObject();
                body(writer);
                writer.WriteEndObject();
            }

            buffer.WriteByte((byte)'\n');

            using var file = new FileStream(
                _path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            buffer.Position = 0;
            buffer.CopyTo(file);
        }
        catch (IOException)
        {
            // Recording is never worth failing a write over.
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    // ---- reading it back --------------------------------------------------------------------

    private void Load()
    {
        if (!File.Exists(_path))
        {
            return;
        }

        var entries = new Dictionary<int, Dictionary<string, ChangeEntry>>();
        var opened = new Dictionary<int, (DateTimeOffset At, string By)>();
        var closed = new Dictionary<int, (DateTimeOffset At, string? Label)>();
        var order = new List<int>();

        try
        {
            foreach (var line in File.ReadLines(_path))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                JsonDocument held;
                try
                {
                    held = JsonDocument.Parse(line);
                }
                catch (JsonException)
                {
                    // A line torn by a crash mid-write. Everything around it is still readable,
                    // which is the point of one line per event.
                    continue;
                }

                using (held)
                {
                    var root = held.RootElement;
                    if (!root.TryGetProperty("k", out var kind)
                        || !root.TryGetProperty("n", out var numbered)
                        || !numbered.TryGetInt32(out var number))
                    {
                        continue;
                    }

                    var at = root.TryGetProperty("at", out var when)
                        && DateTimeOffset.TryParse(
                            when.GetString(), CultureInfo.InvariantCulture,
                            DateTimeStyles.RoundtripKind, out var parsed)
                        ? parsed
                        : DateTimeOffset.MinValue;

                    switch (kind.GetString())
                    {
                        case "open":
                            opened[number] = (at, root.TryGetProperty("by", out var by)
                                ? by.GetString() ?? Unattributed
                                : Unattributed);
                            if (!order.Contains(number))
                            {
                                order.Add(number);
                            }

                            break;

                        case "write":
                        {
                            if (!root.TryGetProperty("module", out var module))
                            {
                                break;
                            }

                            var name = module.GetString() ?? string.Empty;
                            var was = entries.TryGetValue(number, out var held2) ? held2 : [];
                            entries[number] = was;

                            var kindName = root.TryGetProperty("kind", out var said)
                                ? said.GetString() ?? "written"
                                : "written";
                            var settled = kindName switch
                            {
                                "added" => ChangeKind.Added,
                                "removed" => ChangeKind.Removed,
                                _ => ChangeKind.Written,
                            };

                            var before = was.TryGetValue(name, out var already)
                                ? already.Before
                                : root.TryGetProperty("before", out var first) ? first.GetString() : null;
                            var after = root.TryGetProperty("after", out var last) ? last.GetString() : null;

                            was[name] = new ChangeEntry(name, settled, before, after);
                            break;
                        }

                        case "close":
                            closed[number] = (at, root.TryGetProperty("label", out var label)
                                ? label.GetString()
                                : null);
                            break;

                        case "accept":
                            AcceptedAt = number;
                            break;
                    }
                }
            }
        }
        catch (IOException)
        {
            return;
        }
        catch (UnauthorizedAccessException)
        {
            return;
        }

        foreach (var number in order)
        {
            if (!closed.TryGetValue(number, out var end))
            {
                // Opened and never closed: the session ended while it was running. It is a real
                // round and it is over, whatever the file says.
                end = (opened.TryGetValue(number, out var only) ? only.At : DateTimeOffset.MinValue, null);
            }

            _rounds.Add(new ChangeRound(
                number,
                opened.TryGetValue(number, out var start) ? start.At : end.At,
                end.At,
                opened.TryGetValue(number, out var who) ? who.By : Unattributed,
                end.Label,
                false,
                number == AcceptedAt,
                entries.TryGetValue(number, out var mine) ? [.. mine.Values] : []));
        }

        _number = _rounds.Count == 0 ? 0 : _rounds.Max(one => one.Number);
    }
}
