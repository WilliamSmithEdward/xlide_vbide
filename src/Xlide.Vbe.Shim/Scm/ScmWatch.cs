using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Scm;

/// <summary>
/// Watches one repository for the pane: the module folder for files arriving, changing or going,
/// and the repository's own HEAD, ORIG_HEAD, MERGE_HEAD and refs for a checkout, a pull, a commit
/// made elsewhere. Every event is folded into one debounced tap on the pane's shoulder - the pane
/// re-reads a moment later if it is showing - because a checkout touches dozens of files in a
/// burst and a status read per file would be a status read per file.
///
/// The callback runs on a pool thread; the session hops it to the host thread itself.
/// </summary>
internal sealed class ScmWatch : IDisposable
{
    private static readonly TimeSpan Settle = TimeSpan.FromMilliseconds(750);

    private readonly FileSystemWatcher? _folder;
    private readonly FileSystemWatcher? _git;
    private readonly Timer _timer;
    private readonly Action _changed;
    private bool _disposed;

    public ScmWatch(string root, string folder, Action changed)
    {
        ArgumentNullException.ThrowIfNull(changed);

        Root = root;
        Folder = folder;
        _changed = changed;
        _timer = new Timer(_ => Fire(), null, Timeout.Infinite, Timeout.Infinite);

        _folder = TryWatch(folder, includeSubdirectories: false, OnFolderEvent);
        _git = TryWatch(Path.Combine(root, ".git"), includeSubdirectories: true, OnGitEvent);
    }

    public string Root { get; }

    public string Folder { get; }

    private static FileSystemWatcher? TryWatch(string path, bool includeSubdirectories, FileSystemEventHandler handler)
    {
        try
        {
            if (!Directory.Exists(path))
            {
                Log.Info($"scm: not watching {path}, it does not exist");
                return null;
            }

            var watcher = new FileSystemWatcher(path)
            {
                IncludeSubdirectories = includeSubdirectories,
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size,
            };
            watcher.Changed += handler;
            watcher.Created += handler;
            watcher.Deleted += handler;
            watcher.Renamed += (sender, e) => handler(sender, e);
            watcher.Error += (_, e) => Log.Warn($"scm: the watcher on {path} stopped ({e.GetException().Message.Trim()})");
            watcher.EnableRaisingEvents = true;
            return watcher;
        }
        catch (Exception ex)
        {
            Log.Warn($"scm: could not watch {path} ({ex.Message.Trim()})");
            return null;
        }
    }

    private void OnFolderEvent(object sender, FileSystemEventArgs e)
    {
        // The export's own lock and partial files, and .gitignore: not modules, and the export
        // that writes them also writes the module files, which are the events that matter.
        if (e.Name is null || e.Name.StartsWith('.'))
        {
            return;
        }

        Touch();
    }

    private void OnGitEvent(object sender, FileSystemEventArgs e)
    {
        // Only what moves the branch head or a merge's state. The index, the object store and
        // the logs churn on every git command this product runs itself, and none of that changes
        // what the pane shows.
        var name = e.Name?.Replace('\\', '/');
        if (name is null)
        {
            return;
        }

        // packed-refs too: `git gc`, and a fetch or a push on a busy repository, move branches
        // there rather than under refs/, and a branch moved that way was a checkout the pane
        // never heard about.
        if (name is "HEAD" or "ORIG_HEAD" or "MERGE_HEAD" or "FETCH_HEAD" or "packed-refs"
            || name.StartsWith("refs/", StringComparison.Ordinal))
        {
            Touch();
        }
    }

    private void Touch()
    {
        if (_disposed)
        {
            return;
        }

        try
        {
            _timer.Change(Settle, Timeout.InfiniteTimeSpan);
        }
        catch (ObjectDisposedException)
        {
            // Disposed between the check and the change; nothing to tap.
        }
    }

    private void Fire()
    {
        if (_disposed)
        {
            return;
        }

        try
        {
            _changed();
        }
        catch (Exception ex)
        {
            Log.Warn($"scm: the watcher's tap failed ({ex.Message.Trim()})");
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _folder?.Dispose();
        _git?.Dispose();
        _timer.Dispose();
    }
}
