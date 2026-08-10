using Xlide.Vbe.Core.Sync;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// One export folder, held by one instance at a time.
///
/// The defect this prevents is not a torn file: every file is written whole already. It is two
/// Excels exporting to one folder, each with a plan whose deletions were worked out from the
/// modules THAT instance holds, so the second deletes what the first has just written. Every step
/// correct, and a folder matching neither workbook.
/// </summary>
public sealed class FolderLockTests : IDisposable
{
    private readonly string _folder = Path.Combine(
        Path.GetTempPath(), $"xlide-folderlock-{Guid.NewGuid():N}");

    public FolderLockTests() => Directory.CreateDirectory(_folder);

    public void Dispose()
    {
        try
        {
            Directory.Delete(_folder, recursive: true);
        }
        catch (IOException)
        {
            // A test that leaves a temp folder behind has still tested what it came for.
        }
    }

    [Fact]
    public void TheFirstCallerGetsTheFolder()
    {
        using var held = FolderLock.Take(_folder);

        Assert.False(held.HeldByAnother);
        Assert.Null(held.Complaint);
    }

    [Fact]
    public void ASecondCallerIsRefusedWhileTheFirstHoldsIt()
    {
        using var first = FolderLock.Take(_folder);
        using var second = FolderLock.Take(_folder);

        Assert.True(second.HeldByAnother);
        Assert.NotNull(second.Complaint);
    }

    /// <summary>
    /// The refusal names the holder, because "it is locked" sends a developer looking for a lock
    /// and "process 21916 since 11:04" sends them to the other Excel.
    /// </summary>
    [Fact]
    public void TheRefusalSaysWhoHasIt()
    {
        using var first = FolderLock.Take(_folder);
        using var second = FolderLock.Take(_folder);

        Assert.Contains(Environment.ProcessId.ToString(), second.Complaint);
        Assert.Contains("Another Excel", second.Complaint);
    }

    [Fact]
    public void TheFolderIsFreeAgainOnceTheHolderLetsGo()
    {
        using (var first = FolderLock.Take(_folder))
        {
            Assert.False(first.HeldByAnother);
        }

        using var next = FolderLock.Take(_folder);
        Assert.False(next.HeldByAnother);
    }

    /// <summary>
    /// A killed process must not leave a folder locked forever. DeleteOnClose is what guarantees
    /// it, and the operating system honours that even when nothing runs on the way out, so the
    /// nearest thing a test can assert is that no file survives a normal release.
    /// </summary>
    [Fact]
    public void NothingIsLeftBehindToRespectLater()
    {
        using (var held = FolderLock.Take(_folder))
        {
            Assert.False(held.HeldByAnother);
        }

        Assert.Empty(Directory.GetFiles(_folder));
    }

    /// <summary>
    /// TWO DIFFERENT FOLDERS ARE NOT CONTENTION. This is the reason the lock is a file in the
    /// folder rather than a named mutex: two developers exporting different projects to different
    /// places have nothing to say to each other, and a machine-wide lock would make one wait.
    /// </summary>
    [Fact]
    public void ADifferentFolderIsNotBlocked()
    {
        var other = Path.Combine(Path.GetTempPath(), $"xlide-folderlock-{Guid.NewGuid():N}");

        try
        {
            using var here = FolderLock.Take(_folder);
            using var there = FolderLock.Take(other);

            Assert.False(here.HeldByAnother);
            Assert.False(there.HeldByAnother);
        }
        finally
        {
            try { Directory.Delete(other, recursive: true); } catch (IOException) { /* as above */ }
        }
    }

    /// <summary>
    /// A folder that does not exist yet is created rather than refused: the first export to a new
    /// folder is an ordinary thing to do, and it has no contention by definition.
    /// </summary>
    [Fact]
    public void AFolderThatDoesNotExistYetIsMadeRatherThanRefused()
    {
        var fresh = Path.Combine(_folder, "not-yet");

        using var held = FolderLock.Take(fresh);

        Assert.False(held.HeldByAnother);
        Assert.True(Directory.Exists(fresh));
    }

    /// <summary>
    /// No folder at all is not a lock failure. An export with nowhere to go fails on its own terms
    /// and says something true; refusing here would report it as a locking problem.
    /// </summary>
    [Fact]
    public void NoFolderIsNotTreatedAsContention()
    {
        using var held = FolderLock.Take("   ");

        Assert.False(held.HeldByAnother);
    }
}
