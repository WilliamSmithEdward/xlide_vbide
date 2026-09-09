namespace Xlide.Vbe.Core.Scm;

/// <summary>
/// Where git.exe may be. The ordering rule lives here, away from the file system, so it can be
/// tested against a PATH string; the shim probes each candidate with File.Exists and keeps the
/// first that answers.
/// </summary>
public static class GitLocator
{
    /// <summary>
    /// Candidate full paths to git.exe in the order to try them: every PATH entry joined with
    /// git.exe, then the three places Git for Windows installs to - Program Files, Program Files
    /// (x86), and the per-user install under Local AppData, which is where a developer without
    /// administrator rights has it. Empty or null inputs contribute nothing, and a folder named
    /// twice is offered once, because probing it again cannot answer differently.
    /// </summary>
    public static IReadOnlyList<string> Candidates(
        string? pathVariable, string? programFiles, string? programFilesX86, string? localAppData)
    {
        var candidates = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void Offer(string? folder, params string[] below)
        {
            if (string.IsNullOrWhiteSpace(folder))
            {
                return;
            }

            // A PATH entry may be quoted, which cmd.exe tolerates and Path.Combine does not.
            var root = folder.Trim().Trim('"');
            if (root.Length == 0)
            {
                return;
            }

            var path = Path.Combine([root, .. below, ProductIdentity.GitFileName]);
            if (seen.Add(path))
            {
                candidates.Add(path);
            }
        }

        var entries = (pathVariable ?? string.Empty).Split(
            Path.PathSeparator,
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var entry in entries)
        {
            Offer(entry);
        }

        Offer(programFiles, "Git", "cmd");
        Offer(programFilesX86, "Git", "cmd");
        Offer(localAppData, "Programs", "Git", "cmd");

        return candidates;
    }

    /// <summary>
    /// The binary Git for Windows' launchers hand off to, or null when the git.exe found is not
    /// one of them. `cmd\git.exe` and `bin\git.exe` are git-wrapper.exe: every call starts it,
    /// it sets the environment up and starts `mingw64\bin\git.exe` (mingw32 on a 32-bit install),
    /// so each git call was two processes, and the launcher's share was 7 of the 18 milliseconds
    /// a call cost when it was measured (2026-09-09). The real binary is started directly, with
    /// the environment the launcher would have given it - see <see cref="WrapperPath"/>.
    /// </summary>
    public static string? RealBinary(string found, Func<string, bool> exists)
    {
        ArgumentNullException.ThrowIfNull(found);
        ArgumentNullException.ThrowIfNull(exists);

        var directory = Path.GetDirectoryName(found);
        var root = directory is null ? null : Path.GetDirectoryName(directory);
        var launcher = directory is null ? string.Empty : Path.GetFileName(directory);
        if (root is null
            || !(launcher.Equals("cmd", StringComparison.OrdinalIgnoreCase) || launcher.Equals("bin", StringComparison.OrdinalIgnoreCase)))
        {
            return null;
        }

        foreach (var arch in new[] { "mingw64", "mingw32" })
        {
            var real = Path.Combine(root, arch, "bin", ProductIdentity.GitFileName);
            if (exists(real))
            {
                return real;
            }
        }

        return null;
    }

    /// <summary>
    /// The folders the launcher puts at the head of PATH before starting the real binary: its
    /// own, and `usr\bin` beside it, which holds ssh and the other tools git may start.
    /// </summary>
    public static IReadOnlyList<string> WrapperPath(string realBinary)
    {
        ArgumentNullException.ThrowIfNull(realBinary);

        var bin = Path.GetDirectoryName(realBinary) ?? string.Empty;
        var arch = Path.GetDirectoryName(bin) ?? string.Empty;
        var root = Path.GetDirectoryName(arch) ?? string.Empty;
        return [bin, Path.Combine(root, "usr", "bin")];
    }
}
