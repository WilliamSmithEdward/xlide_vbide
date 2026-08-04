using System.Diagnostics;
using System.Globalization;
using System.Text;
using Xlide.Vbe.Core;

namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// Append-only log written from inside the host process.
///
/// Two rules govern everything here. It never throws, because an exception escaping into a COM
/// callback inside Excel is far worse than a lost log line. It never blocks for long, because it
/// runs on the host UI thread.
/// </summary>
internal static class Log
{
    private static readonly Lock Gate = new();
    private static string? _path;
    private static bool _failed;

    /// <summary>
    /// The file, held open. Opening per line was the whole cost of logging: a fresh
    /// CreateFile invites the antivirus to every append, and verbose logging writes thousands
    /// of lines during a resize storm — the host's UI thread was visibly dragging under it
    /// ("resizing is slippery", 2026-08-04). A kept-open stream flushed per line is two
    /// orders of magnitude cheaper and loses nothing in a crash, because every line is
    /// flushed before the call returns.
    /// </summary>
    private static StreamWriter? _writer;

    /// <summary>The thread the log opened on, which is the host's own. See Write.</summary>
    private static readonly int _hostThread = Environment.CurrentManagedThreadId;

    /// <summary>Full path of the current log file, or null if logging could not start.</summary>
    public static string? Path => _path;

    public static void Initialize()
    {
        if (_path is not null || _failed)
        {
            return;
        }

        try
        {
            var directory = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ProductIdentity.DataFolderName,
                "logs");

            Directory.CreateDirectory(directory);

            var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
            _path = System.IO.Path.Combine(directory, $"shim-{stamp}-{Environment.ProcessId}.log");

            _writer = new StreamWriter(
                new FileStream(_path, FileMode.Append, FileAccess.Write, FileShare.Read),
                Encoding.UTF8);

            Write("info", $"log opened, pid {Environment.ProcessId}, {RuntimeDescription()}");
            PruneOldLogs(directory);
        }
        catch
        {
            _failed = true;
            _path = null;
        }
    }

    public static void Info(string message) => Write("info", message);

    public static void Warn(string message) => Write("warn", message);

    public static void Error(string message) => Write("error", message);

    public static void Error(string message, Exception exception) =>
        Write("error", $"{message}: {exception.GetType().Name}: {exception.Message}");

    /// <summary>
    /// Whether verbose lines are written. ON by default while the product is in its development
    /// phase: the developer runs live tests, and the log is the only witness anyone can read
    /// afterwards, so it errs towards telling everything. Set XLIDE_VERBOSE=0 to quiet it;
    /// release engineering (#15) owns flipping the default.
    /// </summary>
    public static bool VerboseEnabled { get; } =
        Environment.GetEnvironmentVariable("XLIDE_VERBOSE") != "0";

    /// <summary>
    /// A line for the development log: state transitions, event traffic, decisions taken.
    /// Callers building an expensive message should check <see cref="VerboseEnabled"/> first.
    /// </summary>
    public static void Verbose(string message)
    {
        if (VerboseEnabled)
        {
            Write("verb", message);
        }
    }

    /*
     * Consecutive duplicate collapsing. Verbose logging turns window-event storms into log
     * storms — a resize drag repeats one identical line hundreds of times — and a log that is
     * mostly repetition hides the line that matters. An identical consecutive line is counted
     * instead of written; the count is flushed when a different line arrives, so the shape of
     * the burst is preserved in one line.
     */
    private static string? _lastKey;
    private static int _suppressed;

    private static void Write(string level, string message)
    {
        if (_failed)
        {
            return;
        }

        try
        {
            // The thread is part of every line because a whole class of defect is "right call,
            // wrong thread", and a log that hides the thread hides the defect. The thread the
            // log was opened on is the host's; everything else is named by number.
            var thread = Environment.CurrentManagedThreadId;
            var origin = thread == _hostThread ? "host" : $"t{thread}";

            var key = string.Concat(level, "|", origin, "|", message);

            lock (Gate)
            {
                if (key == _lastKey)
                {
                    _suppressed++;
                    return;
                }

                var text = new StringBuilder(message.Length + 64);

                if (_suppressed > 0)
                {
                    text.Append(CultureInfo.InvariantCulture,
                        $"{DateTime.Now:HH:mm:ss.fff} [info] [host] … last line repeated {_suppressed} more time(s){Environment.NewLine}");
                    _suppressed = 0;
                }

                _lastKey = key;

                text.Append(CultureInfo.InvariantCulture,
                    $"{DateTime.Now:HH:mm:ss.fff} [{level}] [{origin}] {message}{Environment.NewLine}");

                var line = text.ToString();
                Debug.Write(line);

                if (_writer is { } writer)
                {
                    // Flushed per line: a crash must not eat the line that names it.
                    writer.Write(line);
                    writer.Flush();
                }
            }
        }
        catch
        {
            // A log that cannot be written is not a reason to destabilise the host.
        }
    }

    private static string RuntimeDescription()
    {
        var bitness = Environment.Is64BitProcess ? "x64" : "x86";
        return $"{bitness}, host {Environment.ProcessPath ?? "unknown"}";
    }

    /// <summary>Keeps the log directory from growing without bound across sessions.</summary>
    private static void PruneOldLogs(string directory)
    {
        try
        {
            const int keep = 20;
            var files = new DirectoryInfo(directory).GetFiles("shim-*.log");
            if (files.Length <= keep)
            {
                return;
            }

            Array.Sort(files, static (a, b) => b.LastWriteTimeUtc.CompareTo(a.LastWriteTimeUtc));
            for (var i = keep; i < files.Length; i++)
            {
                files[i].Delete();
            }
        }
        catch
        {
            // Pruning is housekeeping. Failing it changes nothing that matters.
        }
    }
}
