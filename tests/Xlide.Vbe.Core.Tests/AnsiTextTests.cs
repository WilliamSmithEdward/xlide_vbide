using Xlide.Vbe.Core.Vba;
using Xunit;

namespace Xlide.Vbe.Core.Tests;

/// <summary>
/// The code page the editor's files are in, and the fallback that carries bytes through when
/// the page is unknown.
/// </summary>
public class AnsiTextTests
{
    [Fact]
    public void AWesternCodePageSpellsTheCharactersLatin1CannotRead()
    {
        // 0x92 is a right single quotation mark on a Western machine and a C1 control in Latin-1;
        // 0x80 is the euro sign. Both sat in descriptions that never matched under Latin-1.
        var western = AnsiText.For(1252);

        Assert.Equal("’€", western.GetString([0x92, 0x80]));
        Assert.Equal(new byte[] { 0x92, 0x80 }, western.GetBytes("’€"));
    }

    [Fact]
    public void AMultiByteCodePageIsAvailable()
    {
        // Japanese: a two-byte character comes back as the one character, not two.
        var japanese = AnsiText.For(932);

        Assert.Equal("あ", japanese.GetString([0x82, 0xa0]));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(99999)]
    public void AnUnknownOrAbsentCodePageFallsBackToLatin1(int codePage)
    {
        var fallback = AnsiText.For(codePage);

        // Every byte to a character and back, so an unedited file is carried through unchanged.
        var bytes = new byte[256];
        for (var at = 0; at < bytes.Length; at++)
        {
            bytes[at] = (byte)at;
        }
        Assert.Equal(bytes, fallback.GetBytes(fallback.GetString(bytes)));
    }
}
