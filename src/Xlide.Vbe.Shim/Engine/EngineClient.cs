using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;
using Xlide.Vbe.Core;
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
    /// by position, so two concurrent calls would take each other's answers; a completion asked
    /// during an analysis pass waits its turn instead.
    /// </summary>
    private readonly SemaphoreSlim _oneCall = new(1, 1);

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

        await CallAsync("initialize", new Dictionary<string, object>(), cancellation).ConfigureAwait(false);
    }

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
        };

        var result = await CallAsync("project/open", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineProjectOpened);
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

    /// <summary>Analyses one module and returns its findings.</summary>
    public async Task<EngineDiagnostics?> DiagnoseAsync(
        string projectId,
        int generation,
        string moduleName,
        string moduleType,
        string? source,
        CancellationToken cancellation,
        int? activeIncompleteExpressionOffset = null)
    {
        var payload = new Dictionary<string, object>
        {
            ["documentKey"] = $"{projectId}/{moduleName}",
            ["projectId"] = projectId,
            ["generation"] = generation,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

        if (activeIncompleteExpressionOffset is { } activeOffset)
        {
            payload["activeIncompleteExpressionOffset"] = activeOffset;
        }

        var result = await CallAsync("textDocument/diagnostics", payload, cancellation).ConfigureAwait(false);
        if (result is null)
        {
            return null;
        }

        return result.Value.Deserialize(EngineJsonContext.Default.EngineDiagnostics);
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
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["offset"] = offset,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

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
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["offset"] = offset,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

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
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["offset"] = offset,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

        var result = await CallAsync("textDocument/signatureHelp", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineSignatureHelp);
    }

    /// <summary>Asks what Enter should leave behind, given the text just after the newline.</summary>
    public async Task<EngineSmartEnter?> SmartEnterAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        int offset,
        CancellationToken cancellation)
    {
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["offset"] = offset,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

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
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["start"] = start,
            ["end"] = end,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

        var result = await CallAsync("textDocument/codeAction", payload, cancellation).ConfigureAwait(false);
        return result?.Deserialize(EngineJsonContext.Default.EngineCodeActions);
    }

    /// <summary>Asks for a module's colouring: the type references and host globals it holds.</summary>
    public async Task<EngineSemanticTokens?> SemanticTokensAsync(
        string projectId,
        string moduleName,
        string moduleType,
        string? source,
        CancellationToken cancellation)
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
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["start"] = start,
            ["end"] = end,
            ["single"] = single,
            ["completeHeader"] = completeHeader,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

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
        var payload = new Dictionary<string, object>
        {
            ["projectId"] = projectId,
            ["moduleName"] = moduleName,
            ["moduleType"] = moduleType,
            ["offset"] = offset,
        };

        if (source is not null)
        {
            payload["source"] = source;
        }

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
            await _oneCall.WaitAsync().ConfigureAwait(false);
            try
            {
                await writer.WriteLineAsync(line.AsMemory()).ConfigureAwait(false);
            }
            finally
            {
                _oneCall.Release();
            }
        }
        catch (Exception ex)
        {
            Log.Info($"engine: didChange could not be sent ({ex.GetType().Name})");
        }
    }

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

        await _oneCall.WaitAsync(deadline.Token).ConfigureAwait(false);

        try
        {
            await writer.WriteLineAsync(line.AsMemory(), deadline.Token).ConfigureAwait(false);

            // One request is outstanding at a time, so the next line is this call's answer. A
            // pipeline would need correlation by identifier; nothing here benefits from one.
            var response = await reader.ReadLineAsync(deadline.Token).ConfigureAwait(false);
            if (response is null)
            {
                throw new IOException("The engine closed the connection.");
            }

            using var document = JsonDocument.Parse(response);

            if (document.RootElement.TryGetProperty("error", out var error))
            {
                var message = error.TryGetProperty("message", out var text) ? text.GetString() : "unknown";
                Log.Warn($"engine: {method} refused: {message}");
                return null;
            }

            return document.RootElement.TryGetProperty("result", out var result) ? result.Clone() : null;
        }
        finally
        {
            _oneCall.Release();
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
