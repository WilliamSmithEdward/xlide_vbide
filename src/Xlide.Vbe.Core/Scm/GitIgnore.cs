namespace Xlide.Vbe.Core.Scm;

/// <summary>What Initialize writes beside the modules.</summary>
public static class GitIgnore
{
    /// <summary>
    /// The export's own leavings: the folder lock (FolderLock, DeleteOnClose, but visible for the
    /// length of an export) and the partial file every module is written through before its move.
    /// Neither is source, and either showing as untracked would invite somebody to commit it.
    /// </summary>
    public const string Content = ".xlide-sync.lock\n.*.xlide-partial\n";
}
