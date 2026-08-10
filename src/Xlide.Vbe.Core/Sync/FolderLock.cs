namespace Xlide.Vbe.Core.Sync;

/// <summary>
/// One export folder, held by one Excel at a time.
///
/// WHY A FOLDER AND NOT A FILE. Every file this product writes is written whole: to a temporary
/// neighbour and then moved, so a reader never sees half a module. That says nothing about two
/// instances running a whole export against one folder. Each works out its own plan, and the
/// deletions in a plan are derived from the modules THAT instance holds, so the second instance
/// deletes what the first has just written. Every individual step is correct and the outcome is a
/// folder that matches neither workbook.
///
/// WHY A LOCK FILE AND NOT A MUTEX. A named mutex is per machine and this contention is per
/// folder: two developers exporting different projects to different folders have nothing to say
/// to each other, and a machine-wide lock would make one of them wait. A file in the folder is
/// also visible, which a kernel object is not: somebody looking at a stuck export can see what is
/// holding it and which process to blame.
///
/// The handle is opened without sharing WRITE access, so the operating system enforces it rather
/// than this code checking for a file and then creating one, which is a race with a wide window.
/// Reads are shared on purpose: the file says which process holds it, and a lock nobody can read
/// can only ever refuse with "another process". `DeleteOnClose`
/// means a killed Excel leaves nothing behind for the next run to clear up or, worse, to respect.
/// </summary>
public sealed class FolderLock : IDisposable
{
    /// <summary>Deliberately dotted, so it does not look like a module beside the .bas files.</summary>
    private const string LockFileName = ".xlide-sync.lock";

    private readonly FileStream? _held;

    private FolderLock(FileStream? held, string? complaint)
    {
        _held = held;
        Complaint = complaint;
    }

    /// <summary>Null when the lock is ours. What to tell the developer otherwise.</summary>
    public string? Complaint { get; }

    /// <summary>True when somebody else has this folder and nothing should be written to it.</summary>
    public bool HeldByAnother => Complaint is not null;

    /// <summary>
    /// Takes the folder, or reports who has it.
    ///
    /// A folder that cannot be created or written to does NOT block the export: that is the case
    /// where there is no contention to protect against, and refusing there would turn a read-only
    /// or not-yet-created folder into a sync failure with a message about locking, which explains
    /// nothing. The export itself will fail on its own terms and say something true.
    /// </summary>
    public static FolderLock Take(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            return new FolderLock(null, null);
        }

        var path = Path.Combine(folder, LockFileName);

        try
        {
            Directory.CreateDirectory(folder);

            // READ IS SHARED, WRITE IS NOT, and the difference is the whole point of writing who
            // we are into the file. With no sharing at all the identifying line can never be read
            // by the instance that needs it, and every refusal says "another process" (caught by
            // its own test, which is what the test was for). Sharing reads still refuses a second
            // Take, because that asks for write access and this handle does not share it.
            var held = new FileStream(
                path,
                FileMode.Create,
                FileAccess.Write,
                FileShare.Read,
                bufferSize: 64,
                FileOptions.DeleteOnClose);

            // WHO holds it, for the other instance to read out of the message. Written and flushed
            // rather than left in a buffer: the whole value of this line is that a second process
            // can read it while this one is still holding the handle open.
            using (var writer = new StreamWriter(held, System.Text.Encoding.UTF8, 64, leaveOpen: true))
            {
                writer.Write($"xlide, process {Environment.ProcessId}, since {DateTime.Now:HH:mm:ss}");
                writer.Flush();
            }

            held.Flush(true);
            return new FolderLock(held, null);
        }
        catch (IOException)
        {
            // The expected refusal: another instance has it open with no sharing.
            // Not logged from here: this assembly has no logger and should not grow one for a
            // sentence the caller is already being handed.
            var who = Describe(path);
            return new FolderLock(
                null,
                $"Another Excel is syncing this folder ({who}). "
                + "Wait for it to finish, or choose a different folder.");
        }
        catch (UnauthorizedAccessException)
        {
            // Not contention: a folder this account cannot write to. Let the export say so itself.
            return new FolderLock(null, null);
        }
    }

    /// <summary>Whatever the holder wrote about itself, or a shrug that is still true.</summary>
    private static string Describe(string path)
    {
        try
        {
            // Read-only and sharing everything, so this never becomes the thing that blocks the
            // holder it is asking about.
            using var reading = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(reading);
            var said = reader.ReadToEnd().Trim();
            return said.Length > 0 ? said : "another process";
        }
        catch (IOException)
        {
            return "another process";
        }
        catch (UnauthorizedAccessException)
        {
            return "another process";
        }
    }

    public void Dispose() => _held?.Dispose();
}
