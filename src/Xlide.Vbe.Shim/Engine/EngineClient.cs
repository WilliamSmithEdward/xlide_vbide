using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Editor;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Engine;

/// <summary>
/// Owns the analysis engine: starts it, talks to it, and makes sure it dies with us.
///
/// The engine runs in its own process so a slow or failing analysis cannot stall the thread the
/// user is typing on. That isolation is only real if the process is genuinely disposable, so
/// nothing here waits on it from the host user interface thread and every call has a deadline.
///
/// A missing or broken engine is not an error the user should see. The add-in works without
/// analysis; it simply has less to say.
/// </summary>
internal sealed class EngineClient : IAsyncDisposable
{
    private readonly string _executablePath;
    private readonly string _pipeName;

    private Process? _process;
    private NamedPipeClientStream? _pipe;
    private StreamWriter? _writer;
    private StreamReader? _reader;
    private KillOnCloseJob? _job;
    private int _nextId = 1;

    /// <summary>
    /// One request on the pipe at a time. The protocol pairs each answer with the last question
    /// by position, so two concurrent calls would take each other's answers.
    ///
    /// The gate takes them in arrival order with ONE exception, which is its whole reason for
    /// being: a query the developer is waiting for may go in front of background work that is
    /// only queued, never in front of anything that changes what the engine holds. See CallGate
    /// for the measurements that put it there.
    /// </summary>
    private readonly CallGate _oneCall = new();

    private EngineClient(string executablePath)
    {
        _executablePath = executablePath;

        // Unique per add-in instance: two hosts open at once must not share a pipe.
        _pipeName = $"xlide-{Environment.ProcessId}-{Guid.NewGuid():N}";
    }

    public bool IsRunning => _process is { HasExited: false } && _pipe is { IsConnected: true };

    /// <summary>
    /// Starts the engine and connects to it, or returns null when it cannot be started. A null
    /// result is an ordinary outcome and is logged, not thrown.
    /// </summary>
    public static async Task<EngineClient?> StartAsync(string executablePath, CancellationToken cancellation)
    {
        if (!File.Exists(executablePath))
        {
            Log.Warn($"engine: not present at {executablePath}, continuing without analysis");
            return null;
        }

        var client = new EngineClient(executablePath);

        try
        {
            await client.LaunchAsync(cancellation).ConfigureAwait(false);
            Log.Info($"engine: connected on {client._pipeName}");
            return client;
        }
        catch (Exception ex)
        {
            Log.Error("engine: could not be started", ex);
            await client.DisposeAsync().ConfigureAwait(false);
            return null;
        }
    }

    private async Task LaunchAsync(CancellationToken cancellation)
    {
        var startInfo = new ProcessStartInfo(_executablePath)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };

        startInfo.ArgumentList.Add("--pipe");
        startInfo.ArgumentList.Add(_pipeName);

        _process = Process.Start(startInfo) ?? throw new InvalidOperationException("The engine did not start.");

        // Tie the engine to this process at the operating system level. If the host is terminated
        // rather than closed, and nothing here gets to run, the engine still goes with it: a
        // background process outliving its only client is a process nobody will ever clean up.
        _job = KillOnCloseJob.Create();
        _job?.Assign(_process);

        DrainAsync(_process.StandardError, "engine stderr");
        DrainAsync(_process.StandardOutput, "engine");

        _pipe = new NamedPipeClientStream(".", _pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);

        using var connectTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        connectTimeout.CancelAfter(TimeSpan.FromSeconds(20));

        await _pipe.ConnectAsync(connectTimeout.Token).ConfigureAwait(false);

        _writer = new StreamWriter(_pipe, new UTF8Encoding(false)) { AutoFlush = true };
        _reader = new StreamReader(_pipe, new UTF8Encoding(false));

        // THE HANDSHAKE IS READ, not merely performed.
        //
        // The engine answers `{engine, protocol}` and this discarded it, which made the exchange a
        // liveness check and nothing more. The two sides are built from separate trees - the
        // analyzer comes out of the neighbouring checkout - and the failure mode of a mismatch is
        // not a crash but a method answered wrongly or refused as unknown, hours later, with
        // nothing pointing here. A number that costs one comparison is worth reading.
        //
        // Logged rather than fatal. The shim is the newer half in practice and refusing to start
        // would turn a warning into a dead editor; the line is what turns "hover stopped working"
        // into one grep. Bump ExpectedEngineProtocol on the shim side and `protocol` in
        // engine/src/dispatcher.ts together whenever a method or a shape changes.
        var hello = await CallAsync("initialize", new Dictionary<string, object>(), cancellation)
            .ConfigureAwait(false);

        var spoken = hello is { } greeting
            && greeting.TryGetProperty("protocol", out var protocol)
            && protocol.TryGetInt32(out var version)
                ? version
                : -1;

        if (spoken != ExpectedEngineProtocol)
        {
            Log.Warn($"engine: speaks protocol {(spoken < 0 ? "(none reported)" : spoken.ToString())}"
                + $" and this build expects {ExpectedEngineProtocol}. The executable at"
                + $" {_executablePath} may be older than the shim; repackage it with"
                + " `npm run package --prefix engine`.");
        }
    }

    /// <summary>
    /// The engine protocol this build of the shim is written against, answered by `initialize`.
    /// Raise it here and in engine/src/dispatcher.ts together when a method or a shape changes.
    /// </summary>
    private const int ExpectedEngineProtocol = 1;

    /// <summary>
    /// How long an answer may take before the log says so. Above the honest cost of the slowest
    /// legitimate call - a first analysis of a module at VBA's 65,534-line ceiling is about a
    /// second - and far below the pipe's own thirty-second deadline, so it catches a call that
    /// has gone wrong without narrating the ones that are merely large.
    /// </summary>
    private const int SlowAnswerMs = 5000;

    /// <summary>Reads a child stream to the log so it cannot fill its buffer and block the child.</summary>
    private static void DrainAsync(StreamReader stream, string label) =>
        _ = Task.Run(async () =>
        {
            try
            {
                while (await stream.ReadLineAsync().ConfigureAwait(false) is { } line)
                {
                    if (line.Length > 0)
                    {
                        Log.Info($"{label}: {line}");
                    }
                }
            }
            catch (IOException)
            {
                // The child exited. Nothing to report.
            }
        });

    /// <summary>Replaces everything the engine knows about a project.</summary>
    public async Task<EngineProjectOpened?> OpenProjectAsync(
        string projectId,
        int generation,
        EngineModule[] modules,
        CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["generation"] = generation,
            ["modules"] = modules,
            // WHICH OFFICE APPLICATION, so the engine stops asserting Excel in the ones that are
            // not it. Sent with the project rather than at startup because the engine is a child
            // process that may outlive a project and be re-seeded; the host cannot change under
            // it, but the seeding is the one message guaranteed to carry the fact.
            ["host"] = HostApp.Name,
        };

        var result = await CallAsync("project/open", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineProjectOpened);
    }

    /// <summary>
    /// What the engine is HOLDING for a module, which nothing could see until now.
    ///
    /// Every finding is computed against that copy, and it is maintained incrementally by
    /// didChange rather than re-sent whole. So a squiggle on the wrong line is always the same
    /// question - does the engine's copy match the surface's? - and there was no way to ask it.
    /// </summary>
    public async Task<JsonElement?> LiveSourceAsync(
        string projectId,
        string moduleName,
        bool includeText,
        CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["includeText"] = includeText,
        };

        return await CallAsync("debug/liveSource", payload, cancellation).ConfigureAwait(false);
    }

    /// <summary>
    /// What the language service knows about the host's object model: the type inventory bare,
    /// or one type's members when a name is given. Product knowledge, not project state - the
    /// engine answers before any project opens, and for a host with no model it says so.
    /// </summary>
    public async Task<JsonElement?> KnowledgeModelAsync(string? type, CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>();
        if (!string.IsNullOrEmpty(type))
        {
            payload["type"] = type;
        }

        return await CallAsync("knowledge/objectModel", payload, cancellation).ConfigureAwait(false);
    }

    /// <summary>The analyzer's rule catalogue: every diagnostic it can raise, classified.</summary>
    public async Task<JsonElement?> KnowledgeAnalyzerAsync(CancellationToken cancellation)
    {
        return await CallAsync("knowledge/analyzer", new Dictionary<string, object>(), cancellation)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Asks the engine what an import or export would do.
    ///
    /// The engine answers with the COMPANION EDITOR'S OWN planner, which is why the modules go
    /// with the request: their code would otherwise open the .xlsm to read them, and this workbook
    /// is open in Excel, where the file on disk is stale by definition.
    ///
    /// Generously timed. It reads every file in the folder and compares each against a module, and
    /// a large project against a cold folder is slower than anything else asked of the engine.
    /// </summary>
    public async Task<JsonElement?> SyncPlanAsync(
        string direction,
        string workbookPath,
        string folder,
        string? mode,
        IReadOnlyList<Dictionary<string, object>> modules,
        CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>
        {
            ["direction"] = direction,
            ["workbookPath"] = workbookPath,
            ["folder"] = folder,
            ["modules"] = modules,
        };

        if (mode is { Length: > 0 })
        {
            payload["mode"] = mode;
        }

        return await CallAsync("sync/plan", payload, cancellation).ConfigureAwait(false);
    }

    /// <summary>Tells the engine a project is gone, so its modules stop answering for it.</summary>
    public async Task CloseProjectAsync(string projectId, CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
        };

        _ = await CallAsync("project/close", payload, cancellation).ConfigureAwait(false);
    }

    /// <summary>
    /// The analyzer's rule catalog, with each rule's legal severity moves. The truth for the
    /// rules modal and the api both: enumerated from the analyzer actually bundled, so what is
    /// offered is exactly what this build can enforce.
    /// </summary>
    public async Task<EngineAnalysisRules?> RulesAsync(CancellationToken cancellation)
    {
        var result = await CallAsync("analysis/rules", [], cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineAnalysisRules);
    }

    /// <summary>Analyses one module and returns its findings.</summary>
    public async Task<EngineDiagnostics?> DiagnoseAsync(
        string projectId,
        int generation,
        string moduleName,
        string moduleType,
        string? source,
        CancellationToken cancellation,
        int? activeIncompleteExpressionOffset = null,
        IReadOnlyDictionary<string, string>? severityOverrides = null)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["documentKey"] = $"{projectId}/{moduleName}";
        payload["generation"] = generation;

        if (activeIncompleteExpressionOffset is { } activeOffset)
        {
            payload["activeIncompleteExpressionOffset"] = activeOffset;
        }

        // The developer's machine-wide rule choices. Omitted when empty rather than sent as an
        // empty object, the same manners as `source`: the engine keys its memo on the request's
        // shape, and two spellings of "no overrides" would be two cache entries for one answer.
        if (severityOverrides is { Count: > 0 })
        {
            payload["severityOverrides"] = severityOverrides;
        }

        var result = await CallAsync("textDocument/diagnostics", payload, cancellation).ConfigureAwait(false);
        if (result is null)
        {
            return null;
        }

        return result.Value.Deserialize(EngineJsonContext.Default.EngineDiagnostics);
    }

    /// <summary>
    /// The payload every module-scoped call starts from: which project, which module, what kind of
    /// module it is, and the live source when the caller has one.
    ///
    /// Written once because it was written twelve times. The rule that matters is the last one -
    /// source is OMITTED rather than sent as null, because the engine reads a missing source as
    /// "use the seeded copy" - and a rule spelled out twelve times is a rule that will one day be
    /// spelled eleven ways.
    /// </summary>
    private static Dictionary<string, object> ModulePayload(
        string projectId,
        string moduleName,
        string moduleType,
        string? source)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

        return payload;
    }

    /// <summary>Asks what can be typed at an offset into a module's live source.</summary>
    public async Task<EngineCompletions?> CompleteAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int offset,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["offset"] = offset;

        var result = await CallAsync("textDocument/completion", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineCompletions);
    }

    /// <summary>Asks what the identifier at an offset into a module's live source is.</summary>
    public async Task<EngineHover?> HoverAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int offset,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["offset"] = offset;

        var result = await CallAsync("textDocument/hover", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineHover);
    }

    /// <summary>Asks for the call tip at an offset into a module's live source.</summary>
    public async Task<EngineSignatureHelp?> SignatureHelpAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int offset,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["offset"] = offset;

        var result = await CallAsync("textDocument/signatureHelp", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineSignatureHelp);
    }

    /// <summary>Asks what Enter should leave behind, given the text just after the newline.</summary>
    /// <param name="settings">
    /// The developer's typing choices, sent with the request. The engine used to hold these as
    /// constants and the dialog's switches did nothing at all: block layout and both comment
    /// options were offered, persisted, echoed to this shim and reported by the api, and never
    /// reached the code that acts on them (2026-08-08).
    /// </param>
    public async Task<EngineSmartEnter?> SmartEnterAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int offset,
        ProductSettings settings,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        var typing = new Dictionary<string, object>
        {
            ["offset"] = offset,
            ["blockLayout"] = settings.BlockLayout,
            ["continueCommentOnNewline"] = settings.ContinueCommentOnNewline,
            ["mirrorCommentSpacing"] = settings.MirrorCommentSpacing,

            // One indent level as the EDITOR writes it, so smart Enter and the developer's own
            // typing agree. Pressing Enter after a plain line and after `If ... Then` used to
            // produce different whitespace in the same file. Spaces, always: VBA's code store
            // will not hold a tab, and expands any it is handed.
            ["indentUnit"] = new string(' ', Math.Clamp(settings.FormatIndentSize, 1, 16)),
        };

        foreach (var (key, value) in typing)
        {
            payload[key] = value;
        }

        // The typing settings on the wire, because they are the half nobody can see: the engine
        // holds no state, so an answer that ignores a setting is indistinguishable from a
        // setting that never arrived.
        Log.Verbose($"engine: smartEnter layout {settings.BlockLayout}"
            + $", indentUnit {settings.FormatIndentSize} spaces");

        var result = await CallAsync("textDocument/smartEnter", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineSmartEnter);
    }

    /// <summary>
    /// Asks for the quick fixes offered over a span of a module's live source. No findings travel
    /// with the request: the engine resolves fixes from the analysis it holds, which carries fix
    /// data the surface never saw.
    /// </summary>
    public async Task<EngineCodeActions?> CodeActionsAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int start,
        int end,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["start"] = start;
        payload["end"] = end;

        var result = await CallAsync("textDocument/codeAction", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineCodeActions);
    }

    /// <summary>
    /// Asks where the identifier at an offset is declared, or everywhere it is used. One method
    /// for both because the request and the answer are the same shape; the engine names which.
    /// </summary>
    public async Task<EngineLocations?> NavigateAsync(
        string method,
        string projectId,
        string moduleName,
        string moduleType,
        int offset,
        bool includeDeclaration,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source: null);
        payload["offset"] = offset;
        payload["includeDeclaration"] = includeDeclaration;

        var result = await CallAsync(method, payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineLocations);
    }

    /// <summary>
    /// Asks what a rename would make of every module it touches. Whole module texts come back,
    /// because a module with no tab open has no editor to apply an edit list to.
    /// </summary>
    public async Task<EngineRename?> RenameAsync(
        string projectId,
        string moduleName,
        string moduleType,
        int offset,
        string newName,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source: null);
        payload["offset"] = offset;
        payload["newName"] = newName;

        var result = await CallAsync("textDocument/rename", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineRename);
    }

    /// <summary>
    /// Asks what lifting the selected lines into their own procedure would make of the module.
    ///
    /// Lines rather than an offset span, because extraction is a statement operation: the surface
    /// selects text, and half a statement cannot become a procedure. Whole module text back, for
    /// the reason rename gives.
    /// </summary>
    public async Task<EngineExtractMethod?> ExtractMethodAsync(
        string projectId,
        string moduleName,
        string moduleType,
        int startLine,
        int endLine,
        string newName,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source: null);
        payload["startLine"] = startLine;
        payload["endLine"] = endLine;
        payload["newName"] = newName;

        var result = await CallAsync("textDocument/extractMethod", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineExtractMethod);
    }

    /// <summary>
    /// Asks for the stubs a class owes the interfaces it declares. Whole module text back, for the
    /// reason rename gives.
    /// </summary>
    public async Task<EngineImplementInterface?> ImplementInterfaceAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? interfaceName,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source: null);
        if (interfaceName is { Length: > 0 })
        {
            payload["interfaceName"] = interfaceName;
        }

        var result = await CallAsync("textDocument/implementInterface", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineImplementInterface);
    }

    /// <summary>
    /// Asks what renaming a MODULE would make of every module that mentions it. Whole texts, and
    /// nothing written: the caller can still refuse on a name the host will not accept.
    /// </summary>
    public async Task<EngineRename?> RenameModuleAsync(
        string projectId,
        string moduleName,
        string newName,
        CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["newName"] = newName,
        };

        var result = await CallAsync("workspace/renameModule", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineRename);
    }

    /// <summary>Asks for a module's colouring: the type references and host globals it holds.</summary>
    public async Task<EngineSemanticTokens?> SemanticTokensAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);

        var result = await CallAsync("textDocument/semanticTokens", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineSemanticTokens);
    }

    /// <summary>Asks for the case corrections over a span of a module's live source.</summary>
    public async Task<EngineTextEdits?> CanonicalCaseAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int start,
        int end,
        bool single,
        bool completeHeader,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["start"] = start;
        payload["end"] = end;
        payload["single"] = single;
        payload["completeHeader"] = completeHeader;

        var result = await CallAsync("textDocument/canonicalCase", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineTextEdits);
    }

    /// <summary>Asks for the paired loop-iterator rename after an edit at an offset.</summary>
    public async Task<EngineTextEdits?> LoopSyncAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int offset,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);
        payload["offset"] = offset;

        var result = await CallAsync("textDocument/loopSync", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineTextEdits);
    }

    /// <summary>
    /// Asks for a module's procedures. The source is optional: given for the module being
    /// edited, omitted to answer from the engine's seeded copy.
    /// </summary>
    public async Task<EngineOutline?> OutlineAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        CancellationToken cancellation)
    {
        var payload = ModulePayload(projectId, moduleName, moduleType, source);

        var result = await CallAsync("textDocument/outline", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineOutline);
    }

    /// <summary>
    /// Finds text across the modules the engine holds, scoped to one module, one project, or
    /// everything open. The engine searches its live copies, so the results describe the text
    /// as it stands, unsaved edits included.
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
        var payload = new Dictionary<string, object>
        {
            ["scope"] = scope,
            ["query"] = query,
            ["matchCase"] = matchCase,
            ["wholeWord"] = wholeWord,
        };

        if (projectId is not null)
        {
            payload["projectId"] = projectId;
        }

        if (module is not null)
        {
            payload["module"] = module;
        }

        var result = await CallAsync("workspace/search", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineSearchResult);
    }

    /// <summary>
    /// Streams a module's live text, whole or as edits, as a notification: no id, no answer,
    /// and its place in the pipe's order is its meaning. The same one-at-a-time gate the calls
    /// use keeps it ordered among them, and the wait is registered before this returns, so a
    /// request made a moment later cannot overtake the text it is about.
    /// </summary>
    public void NotifyDidChange(string projectId, string moduleName, string? source, EngineTextEdit[]? edits)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }
        else if (edits is not null && edits.Length > 0)
        {
            payload["edits"] = edits;
        }
        else
        {
            return;
        }

        var request = new Dictionary<string, object>
        {
            ["jsonrpc"] = "2.0",
            ["method"] = "textDocument/didChange",
            ["params"] = payload,
        };

        var line = JsonSerializer.Serialize(request, EngineJsonContext.Default.DictionaryStringObject);
        _ = SendNotificationAsync(line);
    }

    private async Task SendNotificationAsync(string line)
    {
        var writer = _writer;
        if (writer is null)
        {
            return;
        }

        try
        {
            // A BARRIER, not a message in a queue: everything asked after this must be answered
            // about the text it carries, so nothing may be let past it.
            await _oneCall.EnterAsync(CallKind.Barrier).ConfigureAwait(false);
            try
            {
                await writer.WriteLineAsync(line.AsMemory()).ConfigureAwait(false);
            }
            finally
            {
                _oneCall.Leave();
            }
        }
        catch (Exception ex)
        {
            Log.Info($"engine: didChange could not be sent ({ex.GetType().Name})");
        }
    }

    /// <summary>
    /// What each method is to the gate. Named one by one rather than by pattern, and anything
    /// unrecognised is a BARRIER: a method nobody has classified is one whose effect on the
    /// engine nobody has thought about, and the safe reading of that is "do not move it".
    ///
    /// Interactive means a person is waiting with their hands on the keyboard. Background means
    /// the answer paints itself whenever it arrives. Everything that seeds, closes, renames or
    /// otherwise changes what the engine holds is a barrier, because a query moved in front of
    /// one would be answered about a project that no longer exists or does not exist yet.
    /// </summary>
    private static CallKind KindOf(string method) => method switch
    {
        "textDocument/completion" => CallKind.Interactive,
        "textDocument/hover" => CallKind.Interactive,
        "textDocument/signatureHelp" => CallKind.Interactive,
        "textDocument/canonicalCase" => CallKind.Interactive,
        "textDocument/smartEnter" => CallKind.Interactive,
        "textDocument/codeAction" => CallKind.Interactive,
        // A pure read of a build-time table; it depends on no seeded state and moves none.
        "analysis/rules" => CallKind.Interactive,
        "textDocument/loopSync" => CallKind.Interactive,

        "textDocument/diagnostics" => CallKind.Background,
        "textDocument/semanticTokens" => CallKind.Background,
        "textDocument/outline" => CallKind.Background,

        // Everything else - initialize, project/open, project/close, debug/liveSource, the
        // renames, workspace/search, the knowledge routes, shutdown - keeps its place.
        _ => CallKind.Barrier,
    };

    private async Task<JsonElement?> CallAsync(string method, Dictionary<string, object> parameters, CancellationToken cancellation)
    {
        var writer = _writer;
        var reader = _reader;

        if (writer is null || reader is null)
        {
            return null;
        }

        var id = Interlocked.Increment(ref _nextId);

        var request = new Dictionary<string, object>
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["method"] = method,
            ["params"] = parameters,
        };

        var line = JsonSerializer.Serialize(request, EngineJsonContext.Default.DictionaryStringObject);

        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        deadline.CancelAfter(TimeSpan.FromSeconds(30));

        // Timed on both sides of the semaphore, deliberately. One request is served at a time,
        // so a method's latency is the wait to get on the pipe PLUS the round trip, and only
        // the second is the analyzer's doing: a diagnostics pass over a large module delays
        // every keystroke's completion request behind it, and one combined figure reports that
        // as slow completions. See EngineCounters.
        var queued = System.Diagnostics.Stopwatch.StartNew();
        await _oneCall.EnterAsync(KindOf(method), deadline.Token).ConfigureAwait(false);
        queued.Stop();
        var refused = false;

        // TIMED IN EVERY BUILD, not only in a debug one, because the failure this had no words for
        // was a call that stopped answering. A diagnosis of a 64,802-line module cost fifteen
        // seconds for a while (xlide_vscode#45), and nothing said so: the caller gave up, the pass
        // was abandoned before it published, and the only trace was a later pass reporting that
        // nothing had changed. It was found by inference from that, not from anything that named
        // it. A stopwatch is nothing against a round trip over a pipe, and now a slow call says
        // which method it was and how long it took.
        var served = System.Diagnostics.Stopwatch.StartNew();

        try
        {
            await writer.WriteLineAsync(line.AsMemory(), deadline.Token).ConfigureAwait(false);

            /*
             * ONCE THE REQUEST IS WRITTEN, ITS ANSWER IS READ. The caller's deadline stops
             * applying here, and the answer is matched to the request by identifier.
             *
             * THIS IS THE ONE THAT KILLED A SESSION. One request is outstanding at a time and the
             * next line used to be taken as this call's answer, on the caller's own token. A
             * caller whose deadline expired after its request went out - a live diagnosis at five
             * seconds, a page request at eight, both reachable on a large module where a single
             * call is over a second - abandoned an answer that was still coming. The next call
             * then read THAT one, and the next read the one before it: the pipe was off by one
             * for the rest of the session, permanently.
             *
             * What that looks like is not a stall. Every answer belongs to the previous question,
             * so `project/open` is told "No current sources for this project, send project/open
             * first", the seed's facts arrive somewhere that expected diagnostics, and the pass
             * dies on a null. The Problems pane keeps whatever it last had and never changes
             * again - "the analyzer stopped working", reported by the owner while typing in a
             * 64,802-line module (2026-08-21, log lines at 20:20:49 onward).
             *
             * So: the read finishes on the PIPE's deadline, not the caller's, and a line whose id
             * is not ours is a leftover from a call that gave up - skipped, said once, and the
             * pipe is back in step instead of broken for good.
             */
            using var pipeDeadline = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            JsonDocument? document = null;

            while (document is null)
            {
                var response = await reader.ReadLineAsync(pipeDeadline.Token).ConfigureAwait(false);
                if (response is null)
                {
                    throw new IOException("The engine closed the connection.");
                }

                var candidate = JsonDocument.Parse(response);
                var answered = candidate.RootElement.TryGetProperty("id", out var answeredId)
                    && answeredId.ValueKind == JsonValueKind.Number
                    && answeredId.TryGetInt32(out var number)
                        ? number
                        : (int?)null;

                if (answered is null || answered == id)
                {
                    document = candidate;
                    continue;
                }

                candidate.Dispose();
                Log.Warn($"engine: an answer to request {answered} arrived while waiting for {id} ({method}); "
                         + "it belonged to a call that gave up, and was dropped rather than read as this one's");
            }

            using (document)
            {
                if (document.RootElement.TryGetProperty("error", out var error))
                {
                    var message = error.TryGetProperty("message", out var text) ? text.GetString() : "unknown";
                    Log.Warn($"engine: {method} refused: {message}");
                    refused = true;
                    return null;
                }

                return document.RootElement.TryGetProperty("result", out var result) ? result.Clone() : null;
            }
        }
        finally
        {
            _oneCall.Leave();

            // In the finally, so a cancelled or thrown call is counted too. A perf hunt that
            // silently drops the calls that went wrong is reading the healthy half only.
            served.Stop();

            if (served.ElapsedMilliseconds >= SlowAnswerMs)
            {
                Log.Warn($"engine: {method} took {served.ElapsedMilliseconds}ms to answer"
                         + (cancellation.IsCancellationRequested
                             ? ", and whoever asked had already given up on it"
                             : string.Empty));
            }
            Diagnostics.EngineCounters.Record(
                method, queued.ElapsedMilliseconds, served.ElapsedMilliseconds, refused);
        }
    }

    public async ValueTask DisposeAsync()
    {
        try
        {
            if (IsRunning)
            {
                using var quick = new CancellationTokenSource(TimeSpan.FromSeconds(2));
                await CallAsync("shutdown", new Dictionary<string, object>(), quick.Token).ConfigureAwait(false);
            }
        }
        catch (Exception)
        {
            // Shutting down politely is a courtesy. The job object below is the guarantee.
        }

        _writer?.Dispose();
        _reader?.Dispose();
        _pipe?.Dispose();

        try
        {
            if (_process is { HasExited: false })
            {
                _process.Kill(entireProcessTree: true);
            }
        }
        catch (Exception)
        {
            // Already gone.
        }

        _process?.Dispose();
        _job?.Dispose();

        _writer = null;
        _reader = null;
        _pipe = null;
        _process = null;
        _job = null;
    }
}
