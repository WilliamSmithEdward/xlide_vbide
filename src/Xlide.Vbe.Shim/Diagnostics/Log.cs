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

            var line = string.Create(CultureInfo.InvariantCulture,
                $"{DateTime.Now:HH:mm:ss.fff} [{level}] [{origin}] {message}{Environment.NewLine}");

            Debug.Write(line);

            var path = _path;
            if (path is null)
            {
                return;
            }

            lock (Gate)
            {
                File.AppendAllText(path, line, Encoding.UTF8);
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
