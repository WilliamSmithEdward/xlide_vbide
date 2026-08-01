using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Engine;

/// <summary>One finding, positioned the way the editing surface wants it.</summary>
internal sealed record Finding(
    string Module,
    string? Code,
    string Message,
    string Severity,
    int StartLine,
    int StartColumn,
    int EndLine,
    int EndColumn);

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
    private readonly HashSet<string> _openProjects = new(StringComparer.OrdinalIgnoreCase);

    public AnalysisService(DispatchObject editor) => _editor = editor;

    /// <summary>Raised when a module has been analysed.</summary>
    public event Action<IReadOnlyList<Finding>>? FindingsReady;

    /// <summary>True once an engine is running and answering.</summary>
    public bool IsReady => _engine is { IsRunning: true };

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
    private async Task AnalyseEverythingAsync()
    {
        var engine = _engine;
        if (engine is null)
        {
            return;
        }

        var generation = Interlocked.Increment(ref _generation);
        var snapshots = ProjectReader.ReadAll(_editor, generation);

        Log.Info($"engine: analysing {snapshots.Count} project(s) at generation {generation}");

        foreach (var snapshot in snapshots)
        {
            await engine.OpenProjectAsync(snapshot.ProjectId, snapshot.Generation, snapshot.Modules, _stopping.Token)
                .ConfigureAwait(false);

            _openProjects.Add(snapshot.ProjectId);

            var findings = new List<Finding>();

            foreach (var module in snapshot.Modules)
            {
                var result = await engine.DiagnoseAsync(
                    snapshot.ProjectId,
                    snapshot.Generation,
                    module.ModuleName,
                    module.Type,
                    module.Source,
                    _stopping.Token).ConfigureAwait(false);

                if (result is null || result.Diagnostics.Length == 0)
                {
                    continue;
                }

                findings.AddRange(Convert(module, result.Diagnostics));
            }

            Log.Info($"engine: {snapshot.ProjectId} produced {findings.Count} finding(s)");

            if (findings.Count > 0)
            {
                FindingsReady?.Invoke(findings);
            }
        }
    }

    /// <summary>
    /// Turns the engine's character offsets into lines and columns.
    ///
    /// The line index is built once per module rather than per finding: converting an offset needs
    /// to know where every line starts, and rebuilding that for each of several hundred findings
    /// turns a linear pass into a quadratic one.
    /// </summary>
    private static IEnumerable<Finding> Convert(EngineModule module, EngineDiagnostic[] diagnostics)
    {
        var lineStarts = TextPositions.LineStarts(module.Source);

        foreach (var diagnostic in diagnostics)
        {
            var (startLine, startColumn) = TextPositions.ToLineColumn(lineStarts, diagnostic.Span.Start);
            var (endLine, endColumn) = TextPositions.ToLineColumn(lineStarts, diagnostic.Span.End);

            yield return new Finding(
                module.ModuleName,
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Severity,
                startLine,
                startColumn,
                endLine,
                endColumn);
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
