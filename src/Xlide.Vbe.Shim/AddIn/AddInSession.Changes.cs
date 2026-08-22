using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Changes;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Changes;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.AddIn;

/*
 * THE CHANGE LOG: what happened to this project's module code, by whom, in rounds.
 *
 * It exists because of a question with no good answer before it: an agent has been editing this
 * workbook through the door for twenty minutes - what did it change? The log answers that at a
 * glance, and holds the text from before so the answer can be read rather than described.
 *
 * IT NEVER WRITES TO THE WORKBOOK. Not a revert button, not an undo stack: a log. Putting text
 * back is the reader's to do - the developer through the editor, where normal undo protects them,
 * or an agent through the write route, where its work lands in this log like anything else. That
 * is what makes it safe to trust: nothing here can lose work, because nothing here writes any.
 *
 * WHAT IT COVERS is module code and nothing else. Not references, not project properties, not a
 * form's design. A log that is quietly partial about what it watches is worse than no log, so it
 * says so in its own answer (`covers`) and the pane repeats it where a reader will see it.
 *
 * WHAT IT COSTS is one copy of a module's text the first time a round touches it - the text the
 * write path had already read to diff against - and nothing at all for modules nobody edited.
 */
internal sealed partial class AddInSession
{
    /// <summary>In words, what the log watches. Said in its own answer, and drawn in the pane.</summary>
    private const string ChangeLogCovers =
        "Module code written through xlide. References, project properties and form designs are "
        + "not recorded, and edits made directly in the VBE are not either.";

    /// <summary>One log per project, opened when a project is first written to.</summary>
    private readonly Dictionary<string, ChangeLog> _changeLogs = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Who the write in flight belongs to, or null for the developer at the keyboard.
    ///
    /// A write that arrives through the debug api is never the developer typing - it is a program,
    /// whether or not it says which - so the door sets this for the length of a request and clears
    /// it afterwards. `by=` on the request names the caller; without it the honest answer is that
    /// something wrote and did not say who.
    /// </summary>
    private string? _writingAs;

    /// <summary>Runs a request with its writes attributed to <paramref name="by"/>.</summary>
    private T AttributedTo<T>(string? by, Func<T> body)
    {
        var was = _writingAs;
        _writingAs = string.IsNullOrWhiteSpace(by) ? ChangeLog.Unattributed : by.Trim();
        try
        {
            return body();
        }
        finally
        {
            _writingAs = was;
        }
    }

    /// <summary>
    /// Which project a caller meant, whatever they called it.
    ///
    /// The pane names a workbook the way the tree shows it and a script names the identity the
    /// `projects` route hands out, and the log is keyed by one thing - so without this they are
    /// two logs for one workbook, and the pane reads an empty one while the route reads a full
    /// one. Exactly the defect the sync route carried until this morning, in a new place
    /// (2026-08-21). Named nothing at all means the project on screen.
    /// </summary>
    private string? ChangeLogProject(string? asked) =>
        asked is { Length: > 0 }
            ? ProjectIdFromDisplay(asked) ?? asked
            : _shownProject;

    /// <summary>The log for a project, or null when there is nowhere to keep one.</summary>
    private ChangeLog? ChangeLogFor(string? projectId)
    {
        if (string.IsNullOrEmpty(projectId))
        {
            return null;
        }

        if (_changeLogs.TryGetValue(projectId, out var already))
        {
            return already;
        }

        try
        {
            var home = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ProductIdentity.DataFolderName);

            var opened = ChangeLog.For(home, projectId);
            _changeLogs[projectId] = opened;
            return opened;
        }
        catch (Exception ex)
        {
            Log.Warn($"changes: no log could be opened for {projectId}, {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Records what a write did, if anything did.
    ///
    /// NEVER THROWS INTO THE WRITE. A log that can take a developer's write down with it is worse
    /// than no log, so everything here is inside the catch: the write has already happened and is
    /// none of this method's business.
    /// </summary>
    private void RecordChange(string module, string? projectId, ChangeKind kind, string? before, string? after)
    {
        try
        {
            if (before == after)
            {
                return;
            }

            ChangeLogFor(projectId)?.Record(
                module, kind, before, after, _writingAs ?? "developer", DateTimeOffset.UtcNow);
        }
        catch (Exception ex)
        {
            Log.Warn($"changes: {module} was not recorded, {ex.Message}");
        }
    }

    /// <summary>
    /// Ends the round that is running on every project, with a label when the caller has one.
    ///
    /// A SAVE IS A BOUNDARY, which is why this is called from there as well as from the route: the
    /// state a developer last saved is already how they think about a last-good point, and drawing
    /// the line there costs nothing.
    /// </summary>
    private void CloseChangeRounds(string? label)
    {
        try
        {
            var now = DateTimeOffset.UtcNow;
            foreach (var log in _changeLogs.Values)
            {
                log.Close(label, now);
            }
        }
        catch (Exception ex)
        {
            Log.Warn($"changes: a round would not close, {ex.Message}");
        }
    }

    /// <summary>
    /// The Changes pane asking the change log something.
    ///
    /// ON THIS THREAD, unlike a sync plan, and deliberately: the work is reading a few files and
    /// comparing texts that are already in hand, and it only happens when somebody is LOOKING at
    /// the pane. Nothing here reads the object model, and nothing here writes.
    /// </summary>
    private void OnChangesRequested(int requestId, IReadOnlyDictionary<string, string> arguments)
    {
        string answer;
        try
        {
            arguments.TryGetValue("action", out var action);
            arguments.TryGetValue("project", out var project);
            arguments.TryGetValue("module", out var module);
            arguments.TryGetValue("which", out var which);
            arguments.TryGetValue("label", out var label);
            arguments.TryGetValue("round", out var round);

            var at = int.TryParse(round, out var parsed) ? parsed : 0;
            var wanted = project is { Length: > 0 } ? project : _shownProject;

            answer = action switch
            {
                "text" => ChangeTextReply(wanted, at, module, which),
                "diff" => ChangeDiffReply(wanted, at, module),
                "snapshot" => Snapshotted(wanted, label),
                "accept" => Accepted(wanted),
                _ => ChangesReply("listed", wanted, 200),
            };
        }
        catch (Exception ex)
        {
            Log.Error("changes: the pane's question could not be answered", ex);
            answer = System.Text.Json.JsonSerializer.Serialize(
                new ChangeLogReply(ex.Message.Trim(), string.Empty, string.Empty, 0, ChangeLogCovers, []),
                ChangeJsonContext.Default.ChangeLogReply);
        }

        _editorSurface?.ShowChangesResult(requestId, answer);
    }

    private string Snapshotted(string? projectId, string? label)
    {
        CloseChangeRounds(label);
        return ChangesReply(
            label is { Length: > 0 } named ? $"the round is closed: {named}" : "the round is closed",
            projectId, 200);
    }

    private string Accepted(string? projectId)
    {
        ChangeLogFor(ChangeLogProject(projectId))?.Accept(DateTimeOffset.UtcNow);
        return ChangesReply("accepted", projectId, 200);
    }

    /// <summary>The change log's answer for one project, rounds newest first.</summary>
    private string ChangesReply(string detail, string? projectId, int limit)
    {
        var wanted = ChangeLogProject(projectId);
        var log = ChangeLogFor(wanted);
        if (log is null)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new ChangeLogReply(
                    "no project is shown, and none was named", string.Empty, string.Empty, 0,
                    ChangeLogCovers, []),
                ChangeJsonContext.Default.ChangeLogReply);
        }

        var rounds = log.Rounds(limit <= 0 ? 200 : limit)
            .Select(round => new ChangeRoundReply(
                round.Number,
                round.Started.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                round.Ended.ToString("o", System.Globalization.CultureInfo.InvariantCulture),
                round.Author,
                round.Label,
                round.Open,
                round.Number == log.AcceptedAt,
                [.. round.Entries
                    .OrderBy(entry => entry.Module, StringComparer.OrdinalIgnoreCase)
                    .Select(entry => EntryReply(log, entry))]))
            .ToArray();

        return System.Text.Json.JsonSerializer.Serialize(
            new ChangeLogReply(
                detail, DisplayFromProjectId(wanted) ?? wanted ?? string.Empty,
                log.Directory, log.AcceptedAt, ChangeLogCovers, rounds),
            ChangeJsonContext.Default.ChangeLogReply);
    }

    /// <summary>
    /// One entry, with the line counts worked out here rather than stored.
    ///
    /// The log keeps facts - which module, which texts - and a count is not a fact about the
    /// write, it is an answer about two texts that can always be worked out again. Deriving it
    /// keeps the count off the write path, where this product has learned twice what a whole-text
    /// comparison costs when it happens per keystroke.
    /// </summary>
    private static ChangeEntryReply EntryReply(ChangeLog log, ChangeEntry entry)
    {
        var before = log.TextOf(entry.Before);
        var after = log.TextOf(entry.After);
        var held = (entry.Before is null || before is not null) && (entry.After is null || after is not null);

        var added = 0;
        var removed = 0;
        if (held)
        {
            foreach (var line in ModuleSync.Diff(before ?? string.Empty, after ?? string.Empty))
            {
                if (line.Kind == DiffKind.Added) { added++; }
                else if (line.Kind == DiffKind.Removed) { removed++; }
                else if (line.Kind == DiffKind.Changed) { added++; removed++; }
            }
        }

        return new ChangeEntryReply(
            entry.Module, entry.Kind.ToString().ToLowerInvariant(),
            added, removed, entry.Before, entry.After, held);
    }

    /// <summary>
    /// One module's change, lined up. The comparison is the host's own and the same one the
    /// import dialog uses - bounded there, so a module whose middle disagrees over thousands of
    /// lines is answered as a block rather than by a table nothing can afford.
    /// </summary>
    private string ChangeDiffReply(string? projectId, int round, string? module)
    {
        var wanted = ChangeLogProject(projectId);
        var log = ChangeLogFor(wanted);
        var found = log?.Rounds(int.MaxValue).FirstOrDefault(one => one.Number == round);
        var entry = found?.Entries.FirstOrDefault(
            one => string.Equals(one.Module, module, StringComparison.OrdinalIgnoreCase));

        if (log is null || entry is null)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new Changes.ChangeDiffReply(
                    found is null ? $"the log holds no round {round}" : $"round {round} did not touch {module}",
                    round, module ?? string.Empty, false, []),
                ChangeJsonContext.Default.ChangeDiffReply);
        }

        var before = log.TextOf(entry.Before);
        var after = log.TextOf(entry.After);
        var held = (entry.Before is null || before is not null) && (entry.After is null || after is not null);

        var rows = held
            ? ModuleSync.Diff(before ?? string.Empty, after ?? string.Empty)
                .Select(line => new Sync.SyncDiffRow(
                    line.LeftNumber, line.RightNumber, line.Left, line.Right,
                    line.Kind.ToString().ToLowerInvariant()))
                .ToArray()
            : [];

        return System.Text.Json.JsonSerializer.Serialize(
            new Changes.ChangeDiffReply(
                held ? "held" : "the text is no longer held", round, entry.Module, held, rows),
            ChangeJsonContext.Default.ChangeDiffReply);
    }

    /// <summary>A text the log kept, named by round and module.</summary>
    private string ChangeTextReply(string? projectId, int round, string? module, string? which)
    {
        var wanted = ChangeLogProject(projectId);
        var log = ChangeLogFor(wanted);
        var side = string.Equals(which, "after", StringComparison.OrdinalIgnoreCase) ? "after" : "before";

        var found = log?.Rounds(int.MaxValue).FirstOrDefault(one => one.Number == round);
        var entry = found?.Entries.FirstOrDefault(
            one => string.Equals(one.Module, module, StringComparison.OrdinalIgnoreCase));

        if (log is null || entry is null)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new Changes.ChangeTextReply(
                    found is null
                        ? $"the log holds no round {round}"
                        : $"round {round} did not touch {module}",
                    round, module ?? string.Empty, side, false, null),
                ChangeJsonContext.Default.ChangeTextReply);
        }

        var hash = side == "after" ? entry.After : entry.Before;
        var text = log.TextOf(hash);

        return System.Text.Json.JsonSerializer.Serialize(
            new Changes.ChangeTextReply(
                hash is null
                    ? $"{module} had no text {side} round {round}"
                    : text is null ? "the text is no longer held" : "held",
                round, entry.Module, side, text is not null, text),
            ChangeJsonContext.Default.ChangeTextReply);
    }
}
