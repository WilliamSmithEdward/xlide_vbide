using System.Buffers.Binary;
using Xlide.Vbe.Core.Forms;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The compound file reader on both versions of the format, and on damaged files.
///
/// Both fixtures were written by Windows itself, StgCreateStorageEx through pywin32 with
/// STGOPTIONS.SectorSize 512 and 4096, so they are the reference implementation's layout rather
/// than a reading of the specification. Each holds a stream below the mini-stream cutoff, one
/// above it, and a storage with two more; each stream holds <see cref="Pattern"/>. In version 4
/// the 512-byte header takes the whole first 4096-byte sector, so sector n starts at
/// (n + 1) * 4096; the reader used 512 + n * 4096 and read nothing.
/// </summary>
public class CompoundFileTests
{
    private static byte[] Fixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));

    /// <summary>The bytes the fixtures' streams were written with; they differ sector to sector.</summary>
    private static byte[] Pattern(int seed, int length)
    {
        var bytes = new byte[length];
        for (var i = 0; i < length; i++)
        {
            bytes[i] = (byte)((seed * 31) + (i * 7) + (i >> 9));
        }

        return bytes;
    }

    private static readonly (string Path, byte[] Data)[] Streams =
    [
        ("/Small", Pattern(1, 100)),
        ("/Big", Pattern(2, 10000)),
        ("/VBA/dir", Pattern(3, 5000)),
        ("/VBA/Module1", Pattern(4, 300)),
    ];

    [Theory]
    [InlineData("Version3Storage.cfb", 3)]
    [InlineData("Version4Storage.cfb", 4)]
    public void EveryStreamReadsBackInEitherVersion(string fixture, int version)
    {
        var bytes = Fixture(fixture);
        Assert.Equal(version, BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(26)));

        var cfb = CompoundFile.TryRead(bytes);
        Assert.NotNull(cfb);
        foreach (var (path, data) in Streams)
        {
            Assert.Equal(data, cfb.Read(path));
        }
    }

    [Theory]
    [InlineData(3, 12)]
    [InlineData(4, 9)]
    [InlineData(3, 0)]
    [InlineData(3, 31)]
    public void ASectorSizeTheVersionDoesNotHaveIsRefused(int version, int shift)
    {
        var bytes = Fixture("Version4Storage.cfb");
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(26), (ushort)version);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(30), (ushort)shift);
        Assert.Null(CompoundFile.TryRead(bytes));
    }

    [Fact]
    public void ADifatChainThatLoopsIsRefusedWhateverCountTheHeaderGives()
    {
        // A sector whose next-DIFAT pointer is itself, and a header asking for four billion.
        var original = Fixture("Version3Storage.cfb");
        var bytes = new byte[original.Length + 512];
        original.CopyTo(bytes, 0);
        var self = (uint)((original.Length / 512) - 1);
        bytes.AsSpan(original.Length, 512).Fill(0xFF);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(bytes.Length - 4), self);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(68), self);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(72), uint.MaxValue);
        Assert.Null(CompoundFile.TryRead(bytes));
    }

    [Fact]
    public void AStreamIsNoLongerThanTheFileWhateverItsEntryClaims()
    {
        // The directory is sector 1 in both fixtures; claim two gigabytes for /Big.
        var bytes = Fixture("Version4Storage.cfb");
        var directory = 2 * 4096;
        var big = System.Text.Encoding.Unicode.GetBytes("Big");
        var at = -1;
        for (var entry = directory; entry < directory + 4096; entry += 128)
        {
            if (bytes.AsSpan(entry, big.Length).SequenceEqual(big) && bytes[entry + big.Length] == 0)
            {
                at = entry;
                break;
            }
        }

        Assert.True(at >= 0, "the fixture's directory has no /Big");
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(at + 120), int.MaxValue);

        var cfb = CompoundFile.TryRead(bytes);
        Assert.NotNull(cfb);
        var read = cfb.Read("/Big");
        Assert.True(read.Length <= bytes.Length);
        Assert.Equal(Pattern(2, 10000), read.AsSpan(0, 10000).ToArray());
    }
}
