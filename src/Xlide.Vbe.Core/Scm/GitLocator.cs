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
}
