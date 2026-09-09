using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Scm;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Scm;

/// <summary>What one git invocation answered.</summary>
/// <param name="StdOutBytes">
/// Standard output as bytes, for `cat-file --batch`, whose sizes count bytes and cannot be walked
/// through a decoded string.
/// </param>
/// <param name="TimedOut">The deadline passed and the process was killed with its tree.</param>
internal sealed record GitResult(
    int ExitCode, string StdOut, string StdErr, byte[] StdOutBytes, long ElapsedMs, bool TimedOut)
{
    public bool Ok => ExitCode == 0 && !TimedOut;

    /// <summary>
    /// What git said, for a reply: stderr when it wrote any, else stdout, trimmed. Git puts its
    /// refusals on stderr and its progress there too, so a failure's words are nearly always the
    /// stderr text.
    /// </summary>
    public string Words =>
        TimedOut ? "git did not finish in time and was stopped"
        : StdErr.Trim().Length > 0 ? StdErr.Trim()
        : StdOut.Trim();
}

/// <summary>
/// The one way git.exe is run. Hidden window, both streams drained, UTF-8, no terminal prompt, a
/// deadline per call, and the process tree killed when it passes: a credential prompt in a hidden
/// console would otherwise wait forever for a keyboard nobody can reach. Never on the host thread
/// - the callers run it from the pool - and never throws for a non-zero exit, only when git cannot
/// be started at all.
/// </summary>
internal sealed class GitClient
{
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private GitClient(string executablePath, string version, IReadOnlyList<string> pathPrefix)
    {
        ExecutablePath = executablePath;
        Version = version;
        PathPrefix = pathPrefix;
    }

    public string ExecutablePath { get; }

    /// <summary>What `git --version` printed, e.g. "git version 2.55.0.windows.5".</summary>
    public string Version { get; }

    /// <summary>
    /// The folders put at the head of PATH for every call: what Git for Windows' launcher would
    /// have done before starting the binary this client starts directly. Empty when the git
    /// found is not that launcher.
    /// </summary>
    public IReadOnlyList<string> PathPrefix { get; }

    /// <summary>
    /// Finds git.exe once per session: every PATH entry, then Git for Windows' three standard
    /// places. Null when none holds one, which the pane reports as noGit; nothing else in the
    /// product changes.
    /// </summary>
    public static GitClient? Locate()
    {
        var candidates = GitLocator.Candidates(
            Environment.GetEnvironmentVariable("PATH"),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));

        var found = candidates.FirstOrDefault(File.Exists);
        if (found is null)
        {
            Log.Info($"scm: git was not found ({candidates.Count} places looked)");
            return null;
        }

        // THE BINARY, NOT THE LAUNCHER. Git for Windows' cmd\git.exe starts a process that
        // starts git, and this product runs git ten times per status: the launcher's process was
        // 7 of every 18 milliseconds (measured 2026-09-09). The real binary is started with the
        // PATH the launcher would have given it, so ssh and the credential manager are found
        // exactly as before.
        var real = GitLocator.RealBinary(found, File.Exists);
        var exe = real ?? found;
        var prefix = real is null ? [] : GitLocator.WrapperPath(real);

        var probe = new GitClient(exe, string.Empty, prefix);
        try
        {
            var version = probe.RunAsync(Path.GetTempPath(), ["--version"], TimeSpan.FromSeconds(20))
                .GetAwaiter().GetResult();
            var words = version.StdOut.Trim();
            Log.Info($"scm: git at {exe} ({(words.Length > 0 ? words : "version unknown")})"
                + (real is null ? string.Empty : $", started directly rather than through {found}"));
            return new GitClient(exe, words.Length > 0 ? words : "git", prefix);
        }
        catch (Exception ex)
        {
            Log.Warn($"scm: git at {exe} would not run ({ex.Message.Trim()})");
            return null;
        }
    }

    /// <summary>
    /// Runs git with <see cref="GitArguments.Prefix"/> first, then <paramref name="args"/>, in
    /// <paramref name="workingDirectory"/>. Standard input, when given, is written whole and
    /// closed before the output is read, which is what `cat-file --batch` needs.
    /// </summary>
    /// <exception cref="Win32Exception">git.exe could not be started.</exception>
    /// <exception cref="InvalidOperationException">Process.Start answered nothing.</exception>
    public async Task<GitResult> RunAsync(
        string workingDirectory, IReadOnlyList<string> args, TimeSpan deadline, byte[]? standardInput = null)
    {
        ArgumentNullException.ThrowIfNull(args);

        var startInfo = new ProcessStartInfo(ExecutablePath)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = standardInput is not null,
            WorkingDirectory = workingDirectory,
            StandardErrorEncoding = Utf8NoBom,
        };

        foreach (var prefix in GitArguments.Prefix)
        {
            startInfo.ArgumentList.Add(prefix);
        }

        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        // No terminal to answer a prompt in: a missing credential fails in words instead of
        // hanging. The language is pinned so the words a reply carries are the ones the docs and
        // the harness match.
        startInfo.Environment["GIT_TERMINAL_PROMPT"] = "0";
        startInfo.Environment["LC_ALL"] = "C";

        if (PathPrefix.Count > 0)
        {
            // What the launcher does before starting this binary: its bin folders first on PATH,
            // for ssh and the other tools git may start, and the MSYS system name those read.
            startInfo.Environment["PATH"] = string.Join(
                Path.PathSeparator, [.. PathPrefix, Environment.GetEnvironmentVariable("PATH") ?? string.Empty]);
            startInfo.Environment["MSYSTEM"] =
                ExecutablePath.Contains("mingw32", StringComparison.OrdinalIgnoreCase) ? "MINGW32" : "MINGW64";
        }

        var began = Stopwatch.GetTimestamp();
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"{ProductIdentity.GitFileName} did not start");

        if (standardInput is not null)
        {
            try
            {
                await process.StandardInput.BaseStream.WriteAsync(standardInput).ConfigureAwait(false);
            }
            catch (IOException)
            {
                // Git exited before reading its input; its exit code and stderr say why.
            }

            process.StandardInput.Close();
        }

        // BOTH STREAMS AT ONCE, or a child that fills the one nobody is reading blocks forever on
        // the other. Stdout is read as bytes: cat-file's sizes count bytes, and a string cannot
        // be walked by them.
        var stdoutTask = ReadAllAsync(process.StandardOutput.BaseStream);
        var stderrTask = process.StandardError.ReadToEndAsync();

        var timedOut = false;
        using (var timeout = new CancellationTokenSource(deadline))
        {
            try
            {
                await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                timedOut = true;
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch (Exception ex) when (ex is InvalidOperationException or Win32Exception)
                {
                    // Already gone between the deadline and the kill.
                }

                await process.WaitForExitAsync().ConfigureAwait(false);
            }
        }

        var stdoutBytes = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);
        var elapsed = (long)Stopwatch.GetElapsedTime(began).TotalMilliseconds;
        var exitCode = timedOut ? -1 : process.ExitCode;

        // Decoded once here for every caller that reads text, BOM tolerated.
        var stdout = Utf8NoBom.GetString(stdoutBytes);
        if (stdout.Length > 0 && stdout[0] == (char)0xFEFF)
        {
            stdout = stdout[1..];
        }

        Log.Info($"scm: git {Spelled(args)} -> exit {exitCode}{(timedOut ? " (killed at the deadline)" : string.Empty)} in {elapsed}ms");
        if (Log.VerboseEnabled && stderr.Trim().Length > 0)
        {
            Log.Verbose($"scm: git stderr: {stderr.Trim()}");
        }

        return new GitResult(exitCode, stdout, stderr, stdoutBytes, elapsed, timedOut);
    }

    private static async Task<byte[]> ReadAllAsync(Stream stream)
    {
        using var held = new MemoryStream();
        await stream.CopyToAsync(held).ConfigureAwait(false);
        return held.ToArray();
    }

    /// <summary>The arguments as a shell would read them, for the log.</summary>
    private static string Spelled(IReadOnlyList<string> args) =>
        string.Join(' ', args.Select(arg =>
            arg.Length == 0 || arg.Any(char.IsWhiteSpace) ? $"\"{arg.Replace("\"", "\\\"", StringComparison.Ordinal)}\"" : arg));
}
