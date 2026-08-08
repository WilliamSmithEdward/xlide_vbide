using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Engine;

/// <summary>One finding, positioned the way the editing surface wants it.</summary>
/// <param name="Project">
/// Identity of the project the module belongs to, so a finding about Book1's Module1 never
/// decorates Book2's. Null only from paths that predate the qualification.
/// </param>
internal sealed record Finding(
    string Module,
    string? Code,
    string Message,
    string Severity,
    int StartLine,
    int StartColumn,
    int EndLine,
    int EndColumn,
    string? Project = null);

/// <summary>
/// Keeps the engine supplied with the editor's current sources and hands back findings.
///
/// Everything asynchronous lives here rather than in the add-in's lifetime code, because the two
/// have opposite constraints: the add-in must never block the host thread, and the engine must
/// never be asked two things at once. This owns the boundary between them.
/// </summary>
internal sealed class AnalysisService : IAsyncDisposable
{
    private readonly DispatchObject _editor;
    private readonly CancellationTokenSource _stopping = new();

    private EngineClient? _engine;
    private int _generation;

    /// <summary>The generation the engine was last seeded with, which live analysis must name.</summary>
    private int _lastSeededGeneration;

    /// <summary>
    /// Projects the engine has been seeded with. Concurrent because pool threads write it
    /// during a pass while feature calls read it, and the session's host thread now asks it
    /// whether a shown project is known at all.
    /// </summary>
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, byte> _openProjects =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Which projects each module name belonged to at the last pass, and what kind of module
    /// each was. Requests arrive with a module name and nothing else; this is what turns the
    /// name back into the engine's addressing. A name can live in several open workbooks at
    /// once, so the value is every home it has, and resolution prefers the shown project.
    /// Swapped wholesale per pass — readers snapshot the reference.
    /// </summary>
    private Dictionary<string, List<(string ProjectId, string ModuleType)>> _moduleHomes =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Identity of the project whose module the surface is showing, kept by the session. The
    /// tie-break when a bare module name has homes in more than one workbook.
    /// </summary>
    public string? PreferredProject { get; set; }

    /// <summary>
    /// The engine address a bare module name means right now.
    ///
    /// Almost every request is about the module on the surface — live text, diagnostics,
    /// completion, hover, everything the editor itself asks — and for those the shown project
    /// always wins: its home when the name has one there, and the shown project directly when
    /// the pass has not seeded it yet, because answering with a same-named module elsewhere is
    /// exactly how one workbook's live text overwrote another's. Only the outline serves
    /// modules that are NOT on the surface (the tree's rows), and it says so: for those a
    /// name's own home stands even when another project is shown. Null when nothing can say.
    /// </summary>
    private (string ProjectId, string ModuleType)? ResolveHome(string moduleName, bool aboutShownModule = true)
    {
        var homes = _moduleHomes;
        var preferred = PreferredProject;

        if (homes.TryGetValue(moduleName, out var candidates) && candidates.Count > 0)
        {
            if (preferred is not null)
            {
                foreach (var candidate in candidates)
                {
                    if (string.Equals(candidate.ProjectId, preferred, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                }
            }

            if (!aboutShownModule || preferred is null)
            {
                return candidates[0];
            }

            return (preferred, "standard");
        }

        if (aboutShownModule && preferred is not null)
        {
            return (preferred, "standard");
        }

        // A module no pass has seen, with nothing shown: a single open project is still a safe
        // address for it.
        return _openProjects.Count == 1 ? (_openProjects.Keys.First(), "standard") : null;
    }

    public AnalysisService(DispatchObject editor) => _editor = editor;

    /// <summary>
    /// Runs an action on the host thread, which owns the object model, answering false while
    /// there is nothing to carry it. The full pass runs on pool threads — the pipe conversation
    /// belongs there — but the read of the projects does not: the editor's objects live on the
    /// host's thread, and calling them from anywhere else worked only by luck. Wired by the
    /// session to the overlay's action timer.
    /// </summary>
    public Func<Action, bool>? HostMarshal { get; set; }

    /// <summary>Raised when a module has been analysed.</summary>
    public event Action<IReadOnlyList<Finding>>? FindingsReady;

    /// <summary>
    /// Raised after a pass has seeded every project, with the union of the projects' words:
    /// names that denote types and names that denote procedures, for the surface's tokenizer.
    /// </summary>
    public event Action<IReadOnlyList<string>, IReadOnlyList<string>>? LanguageFactsReady;

    /// <summary>True once an engine is running and answering.</summary>
    public bool IsReady => _engine is { IsRunning: true };

    /// <summary>True when a pass has seeded this project into the engine.</summary>
    public bool KnowsProject(string projectId) => _openProjects.ContainsKey(projectId);

    /// <summary>
    /// Starts the engine and analyses everything currently open.
    ///
    /// Started rather than awaited: this is called while the host is still bringing itself up, and
    /// nothing here is worth delaying that for.
    /// </summary>
    public void Start()
    {
        var enginePath = EnginePath();
        Log.Info($"engine: looking for {enginePath}");

        _ = Task.Run(async () =>
        {
            try
            {
                _engine = await EngineClient.StartAsync(enginePath, _stopping.Token).ConfigureAwait(false);
                if (_engine is null)
                {
                    return;
                }

                await AnalyseEverythingAsync().ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                // The session ended while the engine was starting.
            }
            catch (Exception ex)
            {
                Log.Error("engine: start-up failed", ex);
            }
        });
    }

    /// <summary>
    /// Reads every project from the editor and analyses each module.
    ///
    /// Reading crosses into the host, so it happens on the caller's thread and produces plain data;
    /// everything after that is off it. Mixing the two is what turns an analysis pass into a stall.
    /// </summary>
    /// <summary>
    /// Re-runs analysis over everything the editor currently holds.
    ///
    /// Started rather than awaited, and called from the host thread after a module has been
    /// written. Reading the projects has to happen on that thread because the object model is
    /// apartment bound, but the engine round trip must not block it, so the read happens first and
    /// the waiting happens elsewhere.
    /// </summary>
    public void Reanalyse()
    {
        if (_engine is not { IsRunning: true })
        {
            return;
        }

        _ = Task.Run(async () =>
        {
            try
            {
                await AnalyseEverythingAsync().ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                // Shutting down.
            }
            catch (Exception ex)
            {
                Log.Error("engine: re-analysis failed", ex);
            }
        });
    }

    /// <summary>
    /// Streams a module's live text to the engine, whole or as edits. Fire-and-forget, but its
    /// place in the pipe's order is registered before this returns, so a request made a moment
    /// later is always about text the engine already holds. A caller who knows the module's
    /// project says so; the rest resolve by name, shown project first.
    /// </summary>
    public void NotifyLiveText(string moduleName, string? source, EngineTextEdit[]? edits, string? projectId = null)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return;
        }

        var home = projectId is { Length: > 0 }
            ? projectId
            : ResolveHome(moduleName) is { } found ? found.ProjectId : null;
        if (home is null)
        {
            return;
        }

        engine.NotifyDidChange(home, moduleName, source, edits);
    }

    /// <summary>
    /// The engine's own copy of a module, for comparing against the surface's. Null when there
    /// is no engine or no address for the module.
    /// </summary>
    public async Task<System.Text.Json.JsonElement?> LiveSourceAsync(
        string moduleName,
        bool includeText,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        return await engine.LiveSourceAsync(home.ProjectId, moduleName, includeText, cancellation)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Asks the engine what can be typed at an offset into a module's live text, or null when
    /// there is no engine or it has not yet seen the module's project.
    /// </summary>
    public async Task<EngineCompletionItem[]?> CompleteAsync(
        string moduleName,
        string source,
        int offset,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        var result = await engine.CompleteAsync(home.ProjectId, moduleName, home.ModuleType, null, offset, cancellation)
            .ConfigureAwait(false);

        return result?.Items;
    }

    /// <summary>
    /// Asks the engine to describe the identifier at an offset into a module's live text, or null
    /// when there is no engine, no address for the module, or nothing under the cursor.
    /// </summary>
    public async Task<EngineHoverPayload?> HoverAsync(
        string moduleName,
        string source,
        int offset,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        var result = await engine.HoverAsync(home.ProjectId, moduleName, home.ModuleType, null, offset, cancellation)
            .ConfigureAwait(false);

        return result?.Hover;
    }

    /// <summary>
    /// Asks the engine for the quick fixes over a span of a module's live text. Empty when there
    /// is no engine, no address for the module, or nothing on that span can be fixed.
    /// </summary>
    public async Task<EngineCodeAction[]> CodeActionsAsync(
        string moduleName,
        int start,
        int end,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return [];
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return [];
        }

        var result = await engine.CodeActionsAsync(home.ProjectId, moduleName, home.ModuleType, null, start, end, cancellation)
            .ConfigureAwait(false);

        return result?.Actions ?? [];
    }

    /// <summary>
    /// Asks the engine where the identifier at an offset is declared, or everywhere in the
    /// workbook it is used. Empty when there is no engine, no address for the module, or nothing
    /// at that offset resolves.
    /// </summary>
    public async Task<EngineLocation[]> NavigateAsync(
        string moduleName,
        int offset,
        bool references,
        bool includeDeclaration,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return [];
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return [];
        }

        var result = await engine.NavigateAsync(
            references ? "textDocument/references" : "textDocument/definition",
            home.ProjectId,
            moduleName,
            home.ModuleType,
            offset,
            includeDeclaration,
            cancellation).ConfigureAwait(false);

        return result?.Locations ?? [];
    }

    /// <summary>
    /// Asks the engine what a rename would make of every module in the workbook it touches. Null
    /// when there is no engine or no address for the module, which is a different fact from a
    /// rename the engine refused and said why.
    /// </summary>
    public async Task<(EngineRename Answer, string ProjectId)?> RenameAsync(
        string moduleName,
        int offset,
        string newName,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        var result = await engine
            .RenameAsync(home.ProjectId, moduleName, home.ModuleType, offset, newName, cancellation)
            .ConfigureAwait(false);

        return result is null ? null : (result, home.ProjectId);
    }

    /// <summary>
    /// Asks the engine what renaming a module would make of every module that mentions it. Null
    /// when there is no engine or no address for the module.
    /// </summary>
    public async Task<(EngineRename Answer, string ProjectId)?> RenameModuleAsync(
        string moduleName,
        string? projectId,
        string newName,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveNamedHome(moduleName, projectId) is not { } home)
        {
            return null;
        }

        var result = await engine
            .RenameModuleAsync(home.ProjectId, moduleName, newName, cancellation)
            .ConfigureAwait(false);

        return result is null ? null : (result, home.ProjectId);
    }

    /// <summary>
    /// Asks the engine for the call tip at an offset into a module's live text, or null when
    /// there is no engine, no address for the module, or the caret is not inside a call.
    /// </summary>
    public async Task<EngineSignatureInfo?> SignatureHelpAsync(
        string moduleName,
        string source,
        int offset,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        var result = await engine.SignatureHelpAsync(home.ProjectId, moduleName, home.ModuleType, null, offset, cancellation)
            .ConfigureAwait(false);

        return result?.Signature;
    }

    /// <summary>
    /// Asks the engine what Enter should leave behind, or null when there is no engine, no
    /// address for the module, or nothing is owed.
    /// </summary>
    public async Task<EngineSmartEnter?> SmartEnterAsync(
        string moduleName,
        string source,
        int offset,
        ProductSettings settings,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        return await engine.SmartEnterAsync(
                home.ProjectId, moduleName, home.ModuleType, null, offset, settings, cancellation)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Asks the engine for the case corrections over a span of a module's live text, or null
    /// when there is no engine or no address for the module.
    /// </summary>
    public async Task<EngineTextEdit[]?> CanonicalCaseAsync(
        string moduleName,
        string source,
        int start,
        int end,
        bool single,
        bool completeHeader,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        var result = await engine.CanonicalCaseAsync(
                home.ProjectId, moduleName, home.ModuleType, null, start, end, single, completeHeader, cancellation)
            .ConfigureAwait(false);

        return result?.Edits;
    }

    /// <summary>
    /// Asks the engine for the paired loop-iterator rename after an edit at an offset, or null
    /// when there is no engine or no address for the module.
    /// </summary>
    public async Task<EngineTextEdit[]?> LoopSyncAsync(
        string moduleName,
        string source,
        int offset,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        var result = await engine.LoopSyncAsync(home.ProjectId, moduleName, home.ModuleType, null, offset, cancellation)
            .ConfigureAwait(false);

        return result?.Edits;
    }

    /// <summary>
    /// Asks the engine for a module's procedures, or null when there is no engine or no address
    /// for the module. The source is given for the module being edited, null otherwise.
    /// </summary>
    public async Task<EngineOutlineProcedure[]?> OutlineAsync(
        string moduleName,
        string? projectId,
        string? source,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        if (ResolveNamedHome(moduleName, projectId) is not { } resolved)
        {
            return null;
        }

        var result = await engine.OutlineAsync(resolved.ProjectId, moduleName, resolved.ModuleType, null, cancellation)
            .ConfigureAwait(false);

        return result?.Procedures;
    }

    /// <summary>
    /// Asks the engine to colour a module: which identifiers name types, which kind of type each
    /// names, and which are host globals nothing has shadowed. Empty when there is no engine or
    /// no address for the module.
    /// </summary>
    public async Task<EngineSemanticToken[]> SemanticTokensAsync(
        string moduleName,
        string? projectId,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return [];
        }

        // Addressed by name rather than taken from the shown module, the way the outline is: a
        // split shows two modules at once and the editor asks about both, so answering only for
        // whichever is host-active would leave one half of the split plainly coloured.
        if (ResolveNamedHome(moduleName, projectId) is not { } resolved)
        {
            return [];
        }

        var result = await engine.SemanticTokensAsync(resolved.ProjectId, moduleName, resolved.ModuleType, null, cancellation)
            .ConfigureAwait(false);

        return result?.Tokens ?? [];
    }

    /// <summary>
    /// Addresses a module the surface may not be showing. An explicit project is the whole
    /// answer; a bare name lets the module's own home stand even while another project is shown.
    /// </summary>
    private (string ProjectId, string ModuleType)? ResolveNamedHome(string moduleName, string? projectId)
    {
        if (projectId is null)
        {
            return ResolveHome(moduleName, aboutShownModule: false);
        }

        if (_moduleHomes.TryGetValue(moduleName, out var candidates))
        {
            foreach (var candidate in candidates)
            {
                if (string.Equals(candidate.ProjectId, projectId, StringComparison.OrdinalIgnoreCase))
                {
                    return candidate;
                }
            }
        }

        return (projectId, "standard");
    }

    /// <summary>
    /// Finds text across the engine's modules. Null when there is no engine to ask; an engine
    /// that answers nothing answers an empty result, which is a different fact.
    /// </summary>
    public async Task<EngineSearchResult?> SearchAsync(
        string scope,
        string? projectId,
        string? module,
        string query,
        bool matchCase,
        bool wholeWord,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine)
        {
            return null;
        }

        return await engine.SearchAsync(scope, projectId, module, query, matchCase, wholeWord, cancellation)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Reads the projects on the host thread, which owns them, and hands the snapshots back to
    /// this pool thread. Null when the read could not be run — the session is between surfaces,
    /// or the host thread never answered — and the pass is abandoned rather than served stale:
    /// whatever prompted it will prompt again.
    /// </summary>
    private async Task<List<ProjectSnapshot>?> ReadProjectsAsync(int generation)
    {
        var read = new TaskCompletionSource<List<ProjectSnapshot>>(
            TaskCreationOptions.RunContinuationsAsynchronously);

        void ReadOnHost()
        {
            try
            {
                read.TrySetResult(ProjectReader.ReadAll(_editor, generation));
            }
            catch (Exception ex)
            {
                read.TrySetException(ex);
            }
        }

        // At start-up the engine can finish connecting a beat before the surface — and its
        // host-thread door — exists, so a declined marshal is retried rather than answered by
        // reading from the wrong thread, which is exactly the call this method exists to end.
        for (var attempt = 0; ; attempt++)
        {
            if (HostMarshal is { } marshal && marshal(ReadOnHost))
            {
                break;
            }

            if (attempt >= 40)
            {
                return null;
            }

            await Task.Delay(250, _stopping.Token).ConfigureAwait(false);
        }

        // Bounded, because the door can also drop its action silently when the surface dies
        // between accept and tick, and a pass that waits forever would sit on nothing.
        var settled = await Task.WhenAny(read.Task, Task.Delay(TimeSpan.FromSeconds(10), _stopping.Token))
            .ConfigureAwait(false);

        if (settled != read.Task)
        {
            return null;
        }

        return await read.Task.ConfigureAwait(false);
    }

    private async Task AnalyseEverythingAsync()
    {
        var engine = _engine;
        if (engine is null)
        {
            return;
        }

        var generation = Interlocked.Increment(ref _generation);
        var snapshots = await ReadProjectsAsync(generation).ConfigureAwait(false);
        if (snapshots is null)
        {
            Log.Info($"engine: generation {generation} could not read the projects, pass abandoned");
            return;
        }

        Log.Info($"engine: analysing {snapshots.Count} project(s) at generation {generation}");

        var factTypes = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        var factProcedures = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var snapshot in snapshots)
        {
            var opened = await engine.OpenProjectAsync(snapshot.ProjectId, snapshot.Generation, snapshot.Modules, _stopping.Token)
                .ConfigureAwait(false);

            if (opened is not null)
            {
                factTypes.UnionWith(opened.Types);
                factProcedures.UnionWith(opened.Procedures);
            }

            _openProjects.TryAdd(snapshot.ProjectId, 0);
            _lastSeededGeneration = snapshot.Generation;

            // This project's homes, rebuilt into a fresh map and swapped in whole: readers
            // snapshot the reference, and a name shared across workbooks keeps every home it
            // has rather than the last writer's.
            var rehomed = new Dictionary<string, List<(string ProjectId, string ModuleType)>>(
                _moduleHomes.Count,
                StringComparer.OrdinalIgnoreCase);

            foreach (var (name, homes) in _moduleHomes)
            {
                var kept = homes
                    .Where(h => !string.Equals(h.ProjectId, snapshot.ProjectId, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (kept.Count > 0)
                {
                    rehomed[name] = kept;
                }
            }

            foreach (var module in snapshot.Modules)
            {
                if (!rehomed.TryGetValue(module.ModuleName, out var list))
                {
                    rehomed[module.ModuleName] = list = [];
                }

                list.Add((snapshot.ProjectId, module.Type));
            }

            _moduleHomes = rehomed;

            var findings = new List<Finding>();

            foreach (var module in snapshot.Modules)
            {
                var result = await engine.DiagnoseAsync(
                    snapshot.ProjectId,
                    snapshot.Generation,
                    module.ModuleName,
                    module.Type,
                    null,
                    _stopping.Token).ConfigureAwait(false);

                if (result is null || result.Diagnostics.Length == 0)
                {
                    continue;
                }

                findings.AddRange(Convert(snapshot.ProjectId, module.ModuleName, module.Source, result.Diagnostics));
            }

            Log.Info($"engine: {snapshot.ProjectId} produced {findings.Count} finding(s)");

            if (findings.Count > 0)
            {
                FindingsReady?.Invoke(findings);
            }
        }

        // Projects the pass no longer saw are gone — closed workbooks, or a save-as that gave
        // the workbook a new identity. The engine forgets them so their modules stop answering,
        // and the homes map drops their entries so a shared name stops offering a dead address.
        var present = new HashSet<string>(snapshots.Select(s => s.ProjectId), StringComparer.OrdinalIgnoreCase);
        foreach (var known in _openProjects.Keys)
        {
            if (!present.Contains(known) && _openProjects.TryRemove(known, out _))
            {
                try
                {
                    await engine.CloseProjectAsync(known, _stopping.Token).ConfigureAwait(false);
                    Log.Info($"engine: {known} is no longer open, closed");
                }
                catch (Exception ex)
                {
                    Log.Info($"engine: {known} could not be closed ({ex.GetType().Name})");
                }

                var pruned = new Dictionary<string, List<(string ProjectId, string ModuleType)>>(
                    _moduleHomes.Count,
                    StringComparer.OrdinalIgnoreCase);

                foreach (var (name, homes) in _moduleHomes)
                {
                    var kept = homes
                        .Where(h => !string.Equals(h.ProjectId, known, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    if (kept.Count > 0)
                    {
                        pruned[name] = kept;
                    }
                }

                _moduleHomes = pruned;
            }
        }

        // After every project has been seeded, so the tokenizer's word lists describe them all.
        LanguageFactsReady?.Invoke([.. factTypes], [.. factProcedures]);
    }

    /// <summary>
    /// Turns the engine's findings into lines and columns.
    ///
    /// The engine sends both: the span in offsets, and the same span already converted against the
    /// text it analysed. Its conversion is the one to trust. A request that carries no source
    /// leaves the engine to choose between its live copy of the module and its seeded one, and
    /// that choice is invisible from here, so converting the offsets against <paramref name="source"/>
    /// is right only when the two texts happen to agree. Formatting a module was the reliable way
    /// to prove they need not: the page had the formatted text and the editor still held the
    /// original, so the engine measured a finding in one and this measured it in the other, and a
    /// squiggle six columns left of its word stayed there until the module was written back.
    ///
    /// <paramref name="source"/> is still the fallback, for an engine too old to send positions.
    /// The line index behind it is built once per module rather than per finding: converting an
    /// offset needs to know where every line starts, and rebuilding that for each of several
    /// hundred findings turns a linear pass into a quadratic one.
    /// </summary>
    private static IEnumerable<Finding> Convert(
        string projectId,
        string moduleName,
        string source,
        EngineDiagnostic[] diagnostics)
    {
        int[]? lineStarts = null;

        foreach (var diagnostic in diagnostics)
        {
            int startLine, startColumn, endLine, endColumn;

            if (diagnostic.At is { } at)
            {
                (startLine, startColumn) = (at.StartLine, at.StartColumn);
                (endLine, endColumn) = (at.EndLine, at.EndColumn);
            }
            else
            {
                lineStarts ??= TextPositions.LineStarts(source);
                (startLine, startColumn) = TextPositions.ToLineColumn(lineStarts, diagnostic.Span.Start);
                (endLine, endColumn) = TextPositions.ToLineColumn(lineStarts, diagnostic.Span.End);
            }

            yield return new Finding(
                moduleName,
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Severity,
                startLine,
                startColumn,
                endLine,
                endColumn,
                projectId);
        }
    }

    /// <summary>
    /// Analyses one module's live text, with the caret so the engine holds back the transient
    /// complaints of an expression mid-edit. Null when there is no engine, no address for the
    /// module, or the engine has not been seeded yet.
    /// </summary>
    public async Task<IReadOnlyList<Finding>?> DiagnoseLiveAsync(
        string moduleName,
        string source,
        int caretOffset,
        CancellationToken cancellation)
    {
        if (_engine is not { IsRunning: true } engine || _lastSeededGeneration == 0)
        {
            return null;
        }

        if (ResolveHome(moduleName) is not { } home)
        {
            return null;
        }

        try
        {
            var result = await engine.DiagnoseAsync(
                    home.ProjectId,
                    _lastSeededGeneration,
                    moduleName,
                    home.ModuleType,
                    null,
                    cancellation,
                    caretOffset)
                .ConfigureAwait(false);

            return result is null ? null : [.. Convert(home.ProjectId, moduleName, source, result.Diagnostics)];
        }
        catch (Exception ex)
        {
            // A reseed can race this and change the generation; the next pause asks again.
            Log.Info($"live: {moduleName} could not be analysed ({ex.GetType().Name})");
            return null;
        }
    }

    /// <summary>
    /// Where the engine lives. Beside the shim in an installation; in the build tree during
    /// development, so a developer does not have to install to try a change.
    /// </summary>
    private static string EnginePath()
    {
        var directory = ShimModule.Directory;

        if (directory is null)
        {
            return ProductIdentity.EngineFileName;
        }

        var installed = Path.Combine(directory, ProductIdentity.EngineFileName);
        if (File.Exists(installed))
        {
            return installed;
        }

        // In a build tree the engine sits in its own output folder, and how deep the shim's output
        // is below the repository root depends on configuration and target. Walking up to find it
        // is correct for every layout; counting directory separators is correct for exactly one.
        var probe = new DirectoryInfo(directory);

        for (var depth = 0; depth < 8 && probe is not null; depth++)
        {
            var candidate = Path.Combine(probe.FullName, "engine", "dist", ProductIdentity.EngineFileName);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            probe = probe.Parent;
        }

        return installed;
    }

    public async ValueTask DisposeAsync()
    {
        await _stopping.CancelAsync().ConfigureAwait(false);

        if (_engine is not null)
        {
            await _engine.DisposeAsync().ConfigureAwait(false);
            _engine = null;
        }

        _stopping.Dispose();
        _openProjects.Clear();
    }
}
