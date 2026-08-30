using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Changes;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Shim.Changes;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.AddIn;

/*
 * THE CHANGE LOG: what happened to this project's module code, by whom, in rounds.
 *
 * It exists because of a question with no good answer before it: an agent has been editing this
 * workbook through the door for twenty minutes - what did it change? The log answers that at a
 * glance, and holds the text from before so the answer can be read rather than described.
 *
 * THE LOG NEVER WRITES TO THE WORKBOOK; since 2026-08-30 the SESSION restores from it. Those are
 * different claims and both matter. The log keeps facts and holds texts - it still cannot lose
 * work, because it still writes none. Restore (the owner: "full ability to restore from any
 * arbitrary snapshot, and revert to last accepted") is the session acting on the log's answers,
 * through the same WriteModule every write takes, and every restore lands as a round of its own -
 * so the log can always take back its own restores, which is what preserves the original
 * principle's substance now that the button exists.
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
    /// <summary>
    /// In words, what the log watches. Said in its own answer, and drawn in the pane.
    ///
    /// IT SAYS WHAT IS RECORDED FIRST, because that is the part a reader needs and the part they
    /// are most likely to doubt: a developer's own typing IS in here, under their own name. The
    /// first wording led with the exclusions and called the uncovered case "edits made directly in
    /// the VBE", which reads as "anything I type" to somebody sitting inside the VBE looking at
    /// this editor - and their typing is recorded, so the sentence said the opposite of the truth
    /// to the person most likely to read it (the owner, 2026-08-22).
    ///
    /// AND THEN IT STOPPED NAMING THE EDITOR'S OWN WINDOW AT ALL. "Written through this editor"
    /// already carries that exclusion, so spelling it out was a second clause that added nothing
    /// and invited the same misreading twice.
    /// </summary>
    private const string ChangeLogCovers =
        "Records module code written through this editor, yours and an agent's alike. Not "
        + "recorded: references, project properties, or form designs.";

    /// <summary>One log per project, opened when a project is first written to.</summary>
    private readonly Dictionary<string, ChangeLog> _changeLogs = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Who the write in flight belongs to, or null for the developer at the keyboard.
    ///
    /// A write that arrives through the xlide api is never the developer typing - it is a program,
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

        var opened = new ChangeLog();
        _changeLogs[projectId] = opened;
        return opened;
    }

    /// <summary>
    /// Records what a write did, if anything did.
    ///
    /// NEVER THROWS INTO THE WRITE. A log that can take a developer's write down with it is worse
    /// than no log, so everything here is inside the catch: the write has already happened and is
    /// none of this method's business.
    /// </summary>
    private void RecordChange(
        string module, string? projectId, ChangeKind kind, string? before, string? after,
        string? from = null, string? componentKind = null)
    {
        try
        {
            // A write that changed nothing is not a change. A RENAME is, though, and it carries no
            // text of its own - so it is never the case this turns away.
            if (kind != ChangeKind.Renamed && before == after)
            {
                return;
            }

            ChangeLogFor(ChangeLogProject(projectId))?.Record(
                module, kind, before, after, _writingAs ?? "developer", DateTimeOffset.UtcNow, from,
                componentKind);

            // A TAP ON THE SHOULDER, CARRYING NO COUNTS. The pane reads the log when it is opened
            // and when the developer asks again, never on the write path, because every count in
            // it is a comparison of two whole texts and this product has learned twice what that
            // costs per keystroke. That rule stands - and it left the pane showing a reading from
            // minutes ago beside live code, with nothing on screen to say which was which (the
            // owner, 2026-08-22, reading +54 next to a module that had grown to 61 lines).
            //
            // So the host says only THAT something landed. No module, no text, no arithmetic: an
            // integer the pane compares with the one it last drew. Writes are flushed when typing
            // stops rather than per keystroke, so this fires at the rate a developer finishes
            // thoughts, not the rate they press keys.
            _editorSurface?.ShowChangesStamp(++_changeStamp);
        }
        catch (Exception ex)
        {
            Log.Warn($"changes: {module} was not recorded, {ex.Message}");
        }
    }

    /// <summary>How many changes this session has recorded. Only ever compared, never read as a
    /// count of anything the developer would recognise.</summary>
    private int _changeStamp;

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
            arguments.TryGetValue("since", out var since);

            var at = int.TryParse(round, out var parsed) ? parsed : 0;
            var wanted = project is { Length: > 0 } ? project : _shownProject;

            answer = action switch
            {
                "text" => ChangeTextReply(wanted, at, module, which),
                "diff" when since == "accept" => ChangeSinceDiffReply(wanted, module),
                "diff" => ChangeDiffReply(wanted, at, module),
                "snapshot" => Snapshotted(wanted, label),
                "accept" => Accepted(wanted),
                // The pane's own restore is the developer at the keyboard, so no `by`: the round
                // it lands as says "developer", exactly as their typing does.
                "restore" => ChangeRestoreReply(wanted, at, module, null),
                "reject" => ChangeRestoreReply(
                    wanted, ChangeLogFor(ChangeLogProject(wanted))?.AcceptedAt ?? 0, module, null,
                    acceptAfter: true),
                _ => ChangesReply("listed", wanted, 200),
            };
        }
        catch (Exception ex)
        {
            Log.Error("changes: the pane's question could not be answered", ex);
            answer = System.Text.Json.JsonSerializer.Serialize(
                new ChangeLogReply(ex.Message.Trim(), string.Empty, 0, ChangeLogCovers, []),
                ChangeJsonContext.Default.ChangeLogReply);
        }

        _editorSurface?.ShowChangesResult(requestId, answer);
    }

    private string Snapshotted(string? projectId, string? label)
    {
        CloseChangeRounds(label);

        // The stamp moves for the same reason a write moves it: the pane follows stamps, and a
        // snapshot or an accept through the DOOR would otherwise leave every open pane - and a
        // tab's painted highlights - describing a mark that has moved (the owner, 2026-08-30).
        _editorSurface?.ShowChangesStamp(++_changeStamp);
        return ChangesReply(
            label is { Length: > 0 } named ? $"the round is closed: {named}" : "the round is closed",
            projectId, 200);
    }

    private string Accepted(string? projectId)
    {
        ChangeLogFor(ChangeLogProject(projectId))?.Accept(DateTimeOffset.UtcNow);
        _editorSurface?.ShowChangesStamp(++_changeStamp);
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
                    "no project is shown, and none was named", string.Empty, 0, ChangeLogCovers, []),
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
                log.AcceptedAt, ChangeLogCovers, rounds, log.RoundCount,
                SinceAccept(log, wanted)),
            ChangeJsonContext.Default.ChangeLogReply);
    }

    /// <summary>
    /// Everything since the accept mark, summed per module and in total - what the Reject button
    /// would take back, worn as a number.
    ///
    /// TEXT AGAINST TEXT, NOT ROUNDS ADDED UP. A module written +5 in one round and -5 back in
    /// the next has changed nothing, and a sum of round counts would call it ten lines. So each
    /// touched module's LIVE text is compared against its text at the accept mark - the same
    /// boundary walk the restore plans from, the open round included, because a count that goes
    /// quiet while somebody is typing under-reports exactly when it is looked at. Cost: one read
    /// and one bounded diff per module touched since the mark, only when the pane asks.
    ///
    /// Null on any failure and null when nothing changed: the button it feeds has nothing to say
    /// then, and a zero row would be furniture.
    /// </summary>
    private ChangeSinceReply? SinceAccept(ChangeLog log, string? projectId)
    {
        try
        {
            var entries = new List<ChangeSinceEntry>();
            foreach (var target in log.RestoreTargets(log.AcceptedAt, includeOpen: true))
            {
                string? before;
                if (target.TextKey is not null)
                {
                    // Aged out means the count would be a guess; the module is listed as changed
                    // with no line arithmetic rather than with digits made up.
                    before = log.TextOf(target.TextKey);
                    if (before is null)
                    {
                        entries.Add(new ChangeSinceEntry(target.NameNow, 0, 0));
                        continue;
                    }
                }
                else
                {
                    // No text at the boundary: it did not exist, or only its name moved.
                    before = string.Empty;
                }

                using var component = FindComponent(target.NameNow, projectId, out _);
                var after = component is null
                    ? string.Empty
                    : ProjectReader.ReadSource(component) ?? string.Empty;
                if (string.Equals(before, after, StringComparison.Ordinal))
                {
                    continue;
                }

                var added = 0;
                var removed = 0;
                foreach (var line in ModuleSync.Diff(before, after))
                {
                    if (line.Kind == DiffKind.Added) { added++; }
                    else if (line.Kind == DiffKind.Removed) { removed++; }
                    else if (line.Kind == DiffKind.Changed) { added++; removed++; }
                }

                entries.Add(new ChangeSinceEntry(target.NameNow, added, removed));
            }

            if (entries.Count == 0)
            {
                return null;
            }

            entries.Sort((left, right) =>
                string.Compare(left.Module, right.Module, StringComparison.OrdinalIgnoreCase));
            return new ChangeSinceReply(
                entries.Count,
                entries.Sum(one => one.Added),
                entries.Sum(one => one.Removed),
                [.. entries]);
        }
        catch (Exception ex)
        {
            Log.Warn($"changes: the since-accept summary could not be built, {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// One module's changes since the accept mark, lined up: its text at the mark against its
    /// LIVE text - the rows the editor's green-and-red highlights are built from.
    /// </summary>
    private string ChangeSinceDiffReply(string? projectId, string? module)
    {
        var wanted = ChangeLogProject(projectId);
        var log = ChangeLogFor(wanted);
        if (log is null || module is not { Length: > 0 })
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new Changes.ChangeDiffReply(
                    log is null ? "no project is shown, and none was named" : "diff since=accept needs module=",
                    0, module ?? string.Empty, false, []),
                ChangeJsonContext.Default.ChangeDiffReply);
        }

        var target = log.RestoreTargets(log.AcceptedAt, includeOpen: true).FirstOrDefault(one =>
            string.Equals(one.NameNow, module, StringComparison.OrdinalIgnoreCase)
            || string.Equals(one.Module, module, StringComparison.OrdinalIgnoreCase));

        var before = target is null
            ? null
            : target.TextKey is null ? string.Empty : log.TextOf(target.TextKey);
        if (target is null || before is null)
        {
            return System.Text.Json.JsonSerializer.Serialize(
                new Changes.ChangeDiffReply(
                    target is null
                        ? $"nothing about {module} has changed since the accept mark"
                        : "the text at the mark is no longer held",
                    log.AcceptedAt, module, false, []),
                ChangeJsonContext.Default.ChangeDiffReply);
        }

        using var component = FindComponent(target.NameNow, wanted, out _);
        var after = component is null
            ? string.Empty
            : ProjectReader.ReadSource(component) ?? string.Empty;

        var rows = ModuleSync.Diff(before, after)
            .Select(line => new Sync.SyncDiffRow(
                line.LeftNumber, line.RightNumber, line.Left, line.Right,
                line.Kind.ToString().ToLowerInvariant()))
            .ToArray();

        return System.Text.Json.JsonSerializer.Serialize(
            new Changes.ChangeDiffReply("held", log.AcceptedAt, target.NameNow, true, rows),
            ChangeJsonContext.Default.ChangeDiffReply);
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
            added, removed, entry.From, held);
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

    /*
     * RESTORE: make what the log remembers true again.
     *
     * THE DESIGN REVERSAL, SAID PLAINLY. The pane shipped show-only, on the principle that a pane
     * that writes is a pane that can lose work - and the owner reversed it (2026-08-30: "full
     * ability to restore from any arbitrary snapshot, and revert to last accepted"). What keeps
     * the principle's substance is that a restore is ITSELF A ROUND: every text it replaces is
     * recorded on the way, so the log can always take back its own restores, and nothing this
     * does is out of reach of the next restore. The log still writes nothing; the session does,
     * through the same WriteModule every other write takes - the stopped-debugger refusal, the
     * long-line refusal, the put-back on a refused replace, the engine and surface sync, all of
     * it, for free.
     *
     * WHAT "TO ROUND N" MEANS: the state when round N ended. A module's text then is the BEFORE
     * of the first round after N that touched it - the one fact the log was built to keep - so
     * only modules touched since the boundary are visited at all, and each is written once with
     * its final text rather than replayed through every round between. That is the whole
     * performance story: a restore costs one write per module that actually differs, nothing
     * per module that does not, and no walk of any history at apply time.
     *
     * ORDER: removes, then renames, then adds, then writes. Removes first because a name freed
     * by a removal may be needed by a rename-back or a re-add; renames before adds so a re-add
     * cannot land on a name a rename-back is about to reclaim; writes last, to modules that by
     * then all exist under their boundary names.
     */
    private string ChangeRestoreReply(
        string? projectId, int boundary, string? onlyModule, string? by, bool acceptAfter = false)
    {
        var wanted = ChangeLogProject(projectId);
        var log = ChangeLogFor(wanted);
        if (log is null)
        {
            return RestoreRefused(boundary, "No project is shown, and none was named.");
        }

        // The same refusal every write path opens with, asked ONCE here so a stopped project
        // answers one sentence rather than one per module.
        if (ProjectModeNow() != DesignMode)
        {
            return RestoreRefused(boundary,
                "The project is stopped in the debugger. Restoring now would reset it and lose "
                + "the run, so nothing was touched. Press Reset in the editor, or POST "
                + "command?name=reset, and restore again.");
        }

        // The running round closes first, so work still open counts as "after the boundary" and
        // the restore lands as a round of its own rather than folding into somebody's edits.
        log.Close(null, DateTimeOffset.UtcNow);

        if (boundary < 0 || boundary > log.NewestClosedRound)
        {
            return RestoreRefused(boundary,
                $"The log holds no round {boundary}; it runs 1 to {log.NewestClosedRound}.");
        }

        var targets = log.RestoreTargets(boundary);
        if (onlyModule is { Length: > 0 })
        {
            targets = [.. targets.Where(one =>
                string.Equals(one.Module, onlyModule, StringComparison.OrdinalIgnoreCase)
                || string.Equals(one.NameNow, onlyModule, StringComparison.OrdinalIgnoreCase))];
            if (targets.Count == 0)
            {
                return RestoreRefused(boundary,
                    $"Nothing about {onlyModule} has changed since round {boundary}, so there is "
                    + "nothing to restore.");
            }
        }

        var display = DisplayFromProjectId(wanted);
        var outcomes = new List<ChangeRestoreOutcome>();

        AttributedTo(by ?? "developer", () =>
        {
            ApplyRestore(log, targets, display, wanted, outcomes);
            return 0;
        });

        log.Close(
            onlyModule is { Length: > 0 }
                ? $"restore {onlyModule} to round {boundary}"
                : $"restore to round {boundary}",
            DateTimeOffset.UtcNow);

        // A CLEAN reject lands reviewed. The developer has just returned everything to the
        // accepted state - the reject round records the going-back, and leaving it above the
        // mark would show unreviewed work that is, by construction, exactly what was already
        // accepted. So the mark comes forward over it (the owner, 2026-08-30: "accept/reject
        // should reset accepted mark - bring it to current"). A plain restore does not: aiming
        // at an arbitrary round says nothing about having reviewed the result.
        //
        // ONLY when nothing was skipped or failed. A reject that could not touch a module - it
        // held unwritten edits, its text had aged out - has NOT returned that module to the
        // accepted state, and stamping the mark over the difference would review work sight
        // unseen. The mark stays, the hybrid button keeps counting the remainder, and the
        // outcome card names the module and the way out - which is the button being right, not
        // stale (the owner's screenshot after a reject that skipped two).
        var landedClean = outcomes.All(one => one.Did is not "skipped" and not "failed");
        if (acceptAfter && landedClean)
        {
            log.Accept(DateTimeOffset.UtcNow);
        }

        // A name that lived twice - added, removed, added again - is two identities in the plan,
        // and both can answer "already absent". True twice, worth reading once.
        var told = outcomes.Distinct().ToList();

        var restored = told.Count(one => one.Did is "written" or "added" or "removed" or "renamed");
        var skipped = told.Count(one => one.Did == "skipped");
        var failed = told.Count(one => one.Did == "failed");
        var summary = $"{restored} restored, {skipped} skipped, "
            + (failed > 0 ? $"{failed} failed, " : string.Empty)
            + $"{told.Count - restored - skipped - failed} already right, to round {boundary}";

        // A reject that could not finish says WHY the counts will keep counting, in its own
        // summary line - the one sentence a glance reads. Without it, a held-back mark and a
        // still-lit button read as the pane failing to update (the owner's screenshot, twice),
        // when they are the pane declining to call unreviewed work reviewed.
        if (acceptAfter && !landedClean)
        {
            summary += ". The accept mark stays where it was: the modules listed below were not "
                + "returned to it - settle them, then reject again";
        }

        Log.Info($"changes: {summary}");
        _editorSurface?.ShowChangesStamp(++_changeStamp);
        return System.Text.Json.JsonSerializer.Serialize(
            new ChangeRestoreReply(
                summary, boundary, restored, skipped + failed, [.. told], log.NewestClosedRound),
            ChangeJsonContext.Default.ChangeRestoreReply);
    }

    private static string RestoreRefused(int boundary, string why) =>
        System.Text.Json.JsonSerializer.Serialize(
            new ChangeRestoreReply(why, boundary, 0, 0, [], 0),
            ChangeJsonContext.Default.ChangeRestoreReply);

    /// <summary>Applies one restore plan, module by module, recording each outcome.</summary>
    private void ApplyRestore(
        ChangeLog log, IReadOnlyList<RestoreTarget> targets,
        string? display, string? projectId, List<ChangeRestoreOutcome> outcomes)
    {
        // REMOVES FIRST: a module that did not exist at the boundary goes, freeing its name.
        foreach (var target in targets.Where(one => !one.ExistedAtBoundary))
        {
            using var standing = FindComponent(target.NameNow, projectId, out _);
            if (standing is null)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.NameNow, "unchanged", "Already absent."));
                continue;
            }

            if (_editorSurface?.HasUnwritten(target.NameNow, display) == true)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.NameNow, "skipped",
                    "It holds edits you have not written yet. Save or discard them, then restore again."));
                continue;
            }

            var gone = RemoveComponent(target.NameNow, display);
            outcomes.Add(gone is null
                ? new ChangeRestoreOutcome(target.NameNow, "removed", null)
                : new ChangeRestoreOutcome(target.NameNow, "failed", gone));
        }

        // RENAMES BACK, in passes: a rename whose target name is still occupied waits for the
        // pass that frees it. Two modules that swapped names deadlock any single order, so a
        // pass that gets nowhere breaks the tie through a temporary name - each temp guarantees
        // the next pass progresses, which bounds the loop.
        var renames = targets
            .Where(one => one.ExistedAtBoundary
                && !string.Equals(one.Module, one.NameNow, StringComparison.OrdinalIgnoreCase))
            .Select(one => (From: one.NameNow, To: one.Module))
            .ToList();
        var tempCount = 0;
        while (renames.Count > 0)
        {
            var progressed = false;
            for (var at = renames.Count - 1; at >= 0; at--)
            {
                var (from, to) = renames[at];

                // RENAMED AND THEN REMOVED since the boundary: there is nothing standing to
                // rename, and the re-add pass below already restores this identity under its
                // boundary name. Planning a rename anyway painted a red "failed - nothing named
                // X to rename back" over a restore that then succeeded (the owner's live test,
                // 2026-08-30). The step stands down without a row; the add speaks for it.
                using var source = FindComponent(from, projectId, out _);
                if (source is null)
                {
                    renames.RemoveAt(at);
                    progressed = true;
                    continue;
                }

                using var taken = FindComponent(to, projectId, out _);
                if (taken is not null)
                {
                    continue;
                }

                var moved = RestoreRename(from, to, projectId);
                outcomes.Add(moved is null
                    ? new ChangeRestoreOutcome(to, "renamed", $"Was {from}.")
                    : new ChangeRestoreOutcome(to, "failed", moved));
                renames.RemoveAt(at);
                progressed = true;
            }

            if (!progressed && renames.Count > 0)
            {
                var (from, to) = renames[0];
                var temp = $"XlideRestoreTmp{++tempCount}";
                var moved = RestoreRename(from, temp, projectId);
                if (moved is not null)
                {
                    outcomes.Add(new ChangeRestoreOutcome(to, "failed", moved));
                    renames.RemoveAt(0);
                    continue;
                }

                renames[0] = (temp, to);
            }
        }

        // RE-ADDS: a module removed since the boundary comes back, with its recorded kind.
        foreach (var target in targets.Where(one => one.ExistedAtBoundary))
        {
            using var standing = FindComponent(target.Module, projectId, out _);
            if (standing is not null)
            {
                continue;
            }

            if (target.ComponentKind == "form")
            {
                // The log records module CODE and says so in `covers`: a form's design is not in
                // it, and a re-added form holding the right code over an empty canvas is a trap,
                // not a restore.
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "skipped",
                    "It is a form, and a form's design is not recorded; only its code would come back."));
                continue;
            }

            var addKind = target.ComponentKind switch
            {
                "class" => 2,
                "module" or null => 1,
                _ => 0,
            };
            if (addKind == 0)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "skipped",
                    $"A {target.ComponentKind} cannot be re-added."));
                continue;
            }

            var added = AddComponentCore(addKind, target.Module, display, projectId, out _);
            if (added is not null)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "failed", added));
                continue;
            }

            var body = log.TextOf(target.TextKey);
            if (body is null)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "added",
                    target.TextKey is null ? null : "Its text is no longer held, so it came back empty."));
                continue;
            }

            var filled = WriteModule(target.Module, body, projectId, hostRewrite: true);
            outcomes.Add(filled is null
                ? new ChangeRestoreOutcome(target.Module, "added", null)
                : new ChangeRestoreOutcome(target.Module, "failed", filled));
        }

        // WRITES: everything else that existed and changed text since the boundary.
        foreach (var target in targets.Where(one => one.ExistedAtBoundary && one.TextKey is not null))
        {
            using var standing = FindComponent(target.Module, projectId, out _);
            if (standing is null)
            {
                // Re-added above, or its add failed - either way the outcome is already listed.
                continue;
            }

            var body = log.TextOf(target.TextKey);
            if (body is null)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "skipped",
                    "The text is no longer held; the log has aged it out."));
                continue;
            }

            if (outcomes.Any(one => one.Module.Equals(target.Module, StringComparison.OrdinalIgnoreCase)
                && one.Did == "added"))
            {
                continue;
            }

            if (_editorSurface?.HasUnwritten(target.Module, display) == true)
            {
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "skipped",
                    "It holds edits you have not written yet. Save or discard them, then restore again."));
                continue;
            }

            // ALREADY HOLDING THE BOUNDARY TEXT - touched since and then brought back by hand, or
            // by an earlier per-module restore. Without this read the write path answered null
            // for its own had-nothing-to-write case and the outcome said "written, 1 restored"
            // about a no-op (found by probing, 2026-08-30). One read per restored module buys the
            // honest word AND skips the write machinery entirely for the module that needs none.
            if (string.Equals(ProjectReader.ReadSource(standing), body, StringComparison.Ordinal))
            {
                outcomes.Add(new ChangeRestoreOutcome(target.Module, "unchanged",
                    "Already holds its text from this boundary."));
                continue;
            }

            var wrote = WriteModule(target.Module, body, projectId, hostRewrite: true);
            outcomes.Add(wrote is null
                ? new ChangeRestoreOutcome(target.Module, "written", null)
                : new ChangeRestoreOutcome(target.Module, "failed", wrote));
        }
    }

    /// <summary>A rename on the restore path: the component's own name, then the session's adoption.</summary>
    private string? RestoreRename(string from, string to, string? projectId)
    {
        try
        {
            using var component = FindComponent(from, projectId, out _);
            if (component is null)
            {
                return $"Nothing named {from} to rename back.";
            }

            component.SetString("Name", to);
            AdoptRename(from, to);
            ComponentsChanged();
            return null;
        }
        catch (Exception ex)
        {
            return $"{from} would not take the name {to} back: {ex.Message.Trim()}";
        }
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
