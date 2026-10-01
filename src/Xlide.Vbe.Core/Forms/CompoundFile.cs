using System.Buffers.Binary;

namespace Xlide.Vbe.Core.Forms;

/// <summary>
/// A compound file (MS-CFB), read from bytes already in hand rather than off a handle - because
/// the bytes this product wants are `xl/vbaProject.bin`, which arrives out of a ZIP.
///
/// This is the mechanism and nothing above it: sectors, the FAT, the mini stream, and the
/// directory tree flattened to paths. What the streams inside MEAN is <see cref="SavedDesign"/>.
///
/// It exists because a workbook's saved VBA storage is the only place that says which properties a
/// developer actually changed, and reading it costs no COM, no export and no running Excel. The
/// harness twin is tools\harness\vba-storage.mjs, which is where the road was proven first.
/// </summary>
public sealed class CompoundFile
{
    private const ulong Signature = 0xE11AB1A1E011CFD0;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;

    /// <summary>A runaway chain is a corrupt file, not a big one. No real VBA storage is anywhere
    /// near this, and a cycle in the FAT would otherwise spin for ever inside the editor.</summary>
    private const int MostSectors = 100_000;

    private readonly byte[] bytes;
    private readonly int sectorSize;
    private readonly int miniSize;
    private readonly uint miniCutoff;
    private readonly uint[] fat;
    private readonly uint[] miniFat;
    private readonly byte[] miniStream;

    /// <summary>Every stream and storage by path, `/EntryForm/f` style.</summary>
    public Dictionary<string, Entry> Paths { get; } = new(StringComparer.Ordinal);

    public readonly record struct Entry(byte Type, uint Start, int Size);

    /// <summary>A storage rather than a stream: 1 is a storage, 2 a stream, 5 the root.</summary>
    public const byte StorageType = 1;

    private CompoundFile(byte[] source)
    {
        bytes = source;

        // MS-CFB 2.2: 512-byte sectors in version 3, 4096 in version 4, 64-byte mini sectors in
        // both. Any other shift is not a compound file, and `1 << shift` of an arbitrary one can
        // be 1, negative, or wrap.
        var version = Read16(26);
        var shift = Read16(30);
        if (shift != (version == 3 ? 9 : version == 4 ? 12 : -1) || Read16(32) != 6)
        {
            throw new InvalidDataException($"Not a version 3 or 4 compound file (version {version}, shift {shift}).");
        }

        sectorSize = 1 << shift;
        miniSize = 64;
        miniCutoff = Read32(56);

        // The FAT reaches through the DIFAT: 109 entries live in the header, the rest in a chain
        // of sectors each ending with the next one's number.
        var fatSectors = new List<uint>();
        for (var i = 0; i < 109; i++)
        {
            var sector = Read32(76 + (i * 4));
            if (sector == FreeSector)
            {
                break;
            }

            fatSectors.Add(sector);
        }

        // The header's count is only a claim, up to four billion: a chain that loops back on
        // itself would add 127 sectors a step for that long, inside Excel's process.
        var difat = Read32(68);
        var difatCount = Read32(72);
        var visited = new HashSet<uint>();
        for (var n = 0; n < difatCount && difat != FreeSector && difat != EndOfChain; n++)
        {
            if (!visited.Add(difat) || visited.Count > MostSectors)
            {
                throw new InvalidDataException($"The DIFAT chain returns to sector {difat}.");
            }

            var here = Offset(difat);
            for (var i = 0; i < (sectorSize / 4) - 1; i++)
            {
                var sector = Read32(here + (i * 4));
                if (sector != FreeSector)
                {
                    fatSectors.Add(sector);
                }
            }

            difat = Read32(here + sectorSize - 4);
        }

        // Every FAT sector is in the file, so there are no more of them than the file has room for;
        // checked before the table is sized from the count.
        if (fatSectors.Count > bytes.Length / sectorSize)
        {
            throw new InvalidDataException($"{fatSectors.Count} FAT sectors do not fit a file of {bytes.Length} bytes.");
        }

        var entries = new uint[fatSectors.Count * (sectorSize / 4)];
        var written = 0;
        foreach (var sector in fatSectors)
        {
            var here = Offset(sector);
            for (var i = 0; i < sectorSize / 4; i++)
            {
                entries[written++] = Read32(here + (i * 4));
            }
        }

        fat = entries;

        var directory = ReadChain(Read32(48), -1);
        var found = new List<DirEntry>();
        for (var i = 0; (i + 1) * 128 <= directory.Length; i++)
        {
            var at = i * 128;
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(directory.AsSpan(at + 64));
            if (nameLength is 0 or > 64)
            {
                found.Add(default);
                continue;
            }

            found.Add(new DirEntry(
                System.Text.Encoding.Unicode.GetString(directory, at, nameLength - 2),
                directory[at + 66],
                BinaryPrimitives.ReadUInt32LittleEndian(directory.AsSpan(at + 68)),
                BinaryPrimitives.ReadUInt32LittleEndian(directory.AsSpan(at + 72)),
                BinaryPrimitives.ReadUInt32LittleEndian(directory.AsSpan(at + 76)),
                BinaryPrimitives.ReadUInt32LittleEndian(directory.AsSpan(at + 116)),
                (int)BinaryPrimitives.ReadUInt64LittleEndian(directory.AsSpan(at + 120))));
        }

        // Small streams are cut out of the mini stream, which hangs off the root entry.
        var root = found.FirstOrDefault(one => one.Type == 5);
        miniStream = root.Size > 0 ? ReadChain(root.Start, -1) : [];
        byte[] miniFatBytes = Read32(60) == EndOfChain ? [] : ReadChain(Read32(60), -1);
        miniFat = new uint[miniFatBytes.Length / 4];
        for (var i = 0; i < miniFat.Length; i++)
        {
            miniFat[i] = BinaryPrimitives.ReadUInt32LittleEndian(miniFatBytes.AsSpan(i * 4));
        }

        Walk(found, root.Child, string.Empty, 0);
    }

    private readonly record struct DirEntry(
        string Name, byte Type, uint Left, uint Right, uint Child, uint Start, int Size);

    /// <summary>
    /// The reader, or null when the bytes are not a compound file or are too damaged to walk.
    /// Never throws: a workbook this cannot read is a form without a saved baseline, which the
    /// projection already knows how to be.
    /// </summary>
    public static CompoundFile? TryRead(byte[] source)
    {
        try
        {
            if (source.Length < 512
                || BinaryPrimitives.ReadUInt64LittleEndian(source) != Signature)
            {
                return null;
            }

            return new CompoundFile(source);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>One stream's bytes, or an empty array when the path names nothing.</summary>
    public byte[] Read(string path)
    {
        if (!Paths.TryGetValue(path, out var entry) || entry.Size <= 0)
        {
            return [];
        }

        if (entry.Size >= miniCutoff)
        {
            return ReadChain(entry.Start, entry.Size);
        }

        // A small stream lives in the mini stream, so it is no longer than the mini stream.
        var whole = new byte[Math.Min(entry.Size, miniStream.Length)];
        var sector = entry.Start;
        var written = 0;
        for (var n = 0; n < MostSectors && written < whole.Length && sector != EndOfChain && sector != FreeSector; n++)
        {
            var from = (long)sector * miniSize;
            if (from >= miniStream.Length)
            {
                break;
            }

            var take = (int)Math.Min(miniSize, Math.Min(miniStream.Length - from, whole.Length - written));
            Array.Copy(miniStream, from, whole, written, take);
            written += take;
            if (written >= whole.Length || sector >= miniFat.Length)
            {
                break;
            }

            sector = miniFat[sector];
        }

        return whole;
    }

    /// <summary>
    /// Where a sector starts. The header takes the whole first sector: 512 bytes in version 3,
    /// and 512 bytes padded to 4096 in version 4 (MS-CFB 2.2), so sector n starts at
    /// (n + 1) * sectorSize, not after the 512 header bytes. Long, because a damaged file's
    /// sector number times 4096 does not fit an int.
    /// </summary>
    private long At(uint sector) => ((long)sector + 1) * sectorSize;

    /// <summary>A sector the reader must have: one past the end of the file is a damaged file.</summary>
    private int Offset(uint sector)
    {
        var at = At(sector);
        if (at + sectorSize > bytes.Length)
        {
            throw new InvalidDataException($"Sector {sector} is past the end of the file.");
        }

        return (int)at;
    }

    private ushort Read16(int at) => BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(at));

    private uint Read32(int at) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(at));

    private byte[] ReadChain(uint start, int size)
    {
        // Each sector once: a chain that loops is cut where it returns, rather than read round
        // and round to the cap.
        var sectors = new List<uint>();
        var seen = new HashSet<uint>();
        var sector = start;
        for (var n = 0; n < MostSectors && sector != EndOfChain && sector != FreeSector && seen.Add(sector); n++)
        {
            sectors.Add(sector);
            if (sector >= fat.Length)
            {
                break;
            }

            sector = fat[sector];
        }

        // No stream is longer than the file that holds it, whatever its directory entry claims.
        var length = size >= 0 ? (long)size : (long)sectors.Count * sectorSize;
        var whole = new byte[Math.Min(length, bytes.Length)];
        var written = 0;
        foreach (var one in sectors)
        {
            var from = At(one);
            if (from >= bytes.Length || written >= whole.Length)
            {
                break;
            }

            var take = (int)Math.Min(sectorSize, Math.Min(bytes.Length - from, whole.Length - written));
            Array.Copy(bytes, from, whole, written, take);
            written += take;
        }

        return whole;
    }

    /// <summary>
    /// Paths, by walking the red-black sibling trees under each storage. Depth is bounded because
    /// a cycle in the tree of a damaged file would otherwise recurse until the stack gives out,
    /// and a form nests containers a handful deep at most.
    /// </summary>
    private void Walk(List<DirEntry> found, uint index, string prefix, int depth)
    {
        if (index == FreeSector || index >= found.Count || depth > 64)
        {
            return;
        }

        var entry = found[(int)index];
        if (entry.Name is null)
        {
            return;
        }

        Walk(found, entry.Left, prefix, depth + 1);

        var path = $"{prefix}/{entry.Name}";
        Paths[path] = new Entry(entry.Type, entry.Start, entry.Size);
        if (entry.Type == StorageType)
        {
            Walk(found, entry.Child, path, depth + 1);
        }

        Walk(found, entry.Right, prefix, depth + 1);
    }
}
