using System.Buffers.Binary;

namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// The run-length container VBA wraps its own streams in ([MS-OVBA] 2.4.1).
///
/// Everything a VBA project keeps as text - the directory of modules, and each module's source
/// with its attribute header - is stored compressed with this one scheme, so nothing inside a
/// `vbaProject.bin` can be read without it. It is not deflate and no framework call does it: a
/// one-byte signature, then chunks of at most 4096 decompressed bytes, each either raw or a
/// sequence of literal bytes and back-references whose length/offset split WIDENS as the chunk
/// fills. That last part is the whole of the difficulty, and getting it wrong does not throw - it
/// quietly produces plausible rubbish.
///
/// Proven against the fixtures before it was written here, in the harness twin
/// tools\harness\vba-storage.mjs, which is the thing to check this against when either changes.
/// </summary>
public static class VbaCompression
{
    /// <summary>Every chunk decompresses to at most this, which the format fixes.</summary>
    private const int ChunkBytes = 4096;

    /// <summary>
    /// The decompressed bytes, or null when this is not a container at all.
    ///
    /// <paramref name="mostBytes"/> stops the walk once enough has been produced, which is not an
    /// optimisation detail but the usual case: a caller after a module's attribute header wants
    /// the first few hundred bytes of a stream that may be a megabyte of code. Stopping early is
    /// safe because a back-reference can only reach backwards inside the chunk it sits in.
    /// </summary>
    public static byte[]? Decompress(byte[] bytes, int start, int mostBytes)
    {
        if (start < 0 || start >= bytes.Length || bytes[start] != 0x01 || mostBytes <= 0)
        {
            return null;
        }

        var into = new List<byte>(Math.Min(mostBytes, ChunkBytes));
        var at = start + 1;

        while (at + 1 < bytes.Length && into.Count < mostBytes)
        {
            var header = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(at));
            at += 2;

            // Bits 0-11 are the chunk's own byte count less three; 12-14 a fixed signature; 15
            // says whether the chunk was worth compressing at all.
            var size = (header & 0x0FFF) + 3;
            var signature = (header >> 12) & 0x07;
            var compressed = (header >> 15) & 0x01;
            if (signature != 0b011)
            {
                // Not a chunk header. Whatever this stream is, it stopped being one here, and
                // what was decoded so far is still the truth about the part that was.
                break;
            }

            var end = Math.Min(at + size - 2, bytes.Length);

            if (compressed == 0)
            {
                for (var i = at; i < end && into.Count < mostBytes; i++)
                {
                    into.Add(bytes[i]);
                }

                at = end;
                continue;
            }

            var chunkStart = into.Count;
            while (at < end && into.Count < mostBytes)
            {
                var flags = bytes[at];
                at++;

                for (var bit = 0; bit < 8 && at < end && into.Count < mostBytes; bit++)
                {
                    if (((flags >> bit) & 1) == 0)
                    {
                        into.Add(bytes[at]);
                        at++;
                        continue;
                    }

                    if (at + 1 >= bytes.Length)
                    {
                        break;
                    }

                    var token = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(at));
                    at += 2;

                    // THE SPLIT MOVES. Early in a chunk there is little to point back at, so most
                    // of the token is length; as the chunk fills, more of it becomes offset. Get
                    // this boundary wrong by one bit and the output is still bytes, just the
                    // wrong ones - which is why this is measured against real fixtures rather
                    // than reasoned about.
                    var filled = into.Count - chunkStart;
                    var bitCount = 4;
                    while (bitCount < 12 && (1 << bitCount) < filled)
                    {
                        bitCount++;
                    }

                    var lengthMask = 0xFFFF >> bitCount;
                    var length = (token & lengthMask) + 3;
                    var offset = ((token & ~lengthMask & 0xFFFF) >> (16 - bitCount)) + 1;

                    var from = into.Count - offset;
                    if (from < 0)
                    {
                        // A reference pointing before the output is a corrupt stream, not a long
                        // one. Stop rather than invent bytes to satisfy it.
                        return [.. into];
                    }

                    for (var k = 0; k < length && into.Count < mostBytes; k++)
                    {
                        into.Add(into[from + k]);
                    }
                }
            }

            at = Math.Max(at, end);
        }

        return [.. into];
    }
}
