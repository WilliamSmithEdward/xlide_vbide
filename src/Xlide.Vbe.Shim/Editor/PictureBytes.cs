using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// A control's PICTURE, both ways: out as bytes the canvas can draw, and in from a file the
/// developer chose.
///
/// The standing risk in docs/userform-designer.md said this might need an OLE round trip through a
/// temporary stream. It does not. The way in is the one part of an IPictureDisp that IS on its
/// dispatch interface - `Handle`, a GDI HBITMAP or HICON - and from there it is Win32 rather than
/// COM: GetDIBits for a bitmap, DrawIconEx onto a DIB section for an icon, and a BMP is
/// fifty-four bytes of header around the pixels. No IPicture vtable, no OLE stream, no temp file.
///
/// What it does NOT do is guess. A picture whose handle will not answer, a metafile (a drawing
/// rather than pixels), or a bitmap too large to be worth inlining all come back null, and the
/// canvas keeps drawing the honest empty box it drew before - which is the architecture's rule
/// for anything the round trip refuses.
/// </summary>
internal static unsafe partial class PictureBytes
{
    /// <summary>What a picture answers for its Type: 1 is a bitmap, 2 a metafile, 3 an icon,
    /// 4 an enhanced metafile. 0 is a picture object holding nothing.</summary>
    private const int PictureTypeBitmap = 1;
    private const int PictureTypeMetafile = 2;
    private const int PictureTypeIcon = 3;
    private const int PictureTypeEnhancedMetafile = 4;

    /// <summary>The ceiling on what rides the wire as a data URI. A form's picture is decoration
    /// at 96dpi; past this it is a photograph somebody pasted, and the canvas is better off
    /// drawing honest bounds than the page holding megabytes of base64 through every redraw.</summary>
    private const int MostPixels = 4_000_000;

    /// <summary>
    /// A picture as a `data:image/bmp;base64,...` URI, or null when there is nothing to read -
    /// which includes a kind this cannot turn into pixels.
    ///
    /// Takes the PICTURE rather than an owner and a property name, which was the first cut's
    /// shape: every caller holds the object already, because it wants the kind as well as the
    /// pixels, and the convenience overload was never called from outside this file.
    /// </summary>
    public static string? DataUriOf(DispatchObject picture)
    {
        try
        {
            var kind = FormDesignService.TryInt(picture, "Type") ?? 0;
            var handle = HandleOf(picture);
            if (handle == 0)
            {
                return null;
            }

            var bytes = kind switch
            {
                PictureTypeBitmap => BitmapBytes(handle),
                PictureTypeIcon => IconBytes(handle),
                // A metafile is a drawing rather than pixels, so there is nothing to read out of
                // it here. Not a failure: the row still says what it holds, and the canvas still
                // draws the bounds.
                _ => null,
            };

            return bytes is null ? null : $"data:image/bmp;base64,{Convert.ToBase64String(bytes)}";
        }
        catch (Exception why)
        {
            Log.Verbose($"picture: a picture would not read ({why.Message.Trim()})");
            return null;
        }
    }

    /// <summary>
    /// What the property HOLDS, in the words the panel shows: the native designer's own
    /// vocabulary, which names the kind rather than pretending a picture has a value.
    /// </summary>
    public static string Describe(DispatchObject? picture)
    {
        if (picture is null)
        {
            return "(None)";
        }

        return (FormDesignService.TryInt(picture, "Type") ?? 0) switch
        {
            PictureTypeBitmap => "(Bitmap)",
            PictureTypeMetafile or PictureTypeEnhancedMetafile => "(Metafile)",
            PictureTypeIcon => "(Icon)",
            _ => "(None)",
        };
    }

    /// <summary>
    /// The picture's own size in PIXELS, for a caller that wants to say how big it is without
    /// carrying it: the api's designer read describes a picture, it does not ship one.
    ///
    /// Measured off the GDI handle rather than off the picture's Width and Height, which are
    /// HIMETRIC hundredths of a millimetre and would need this machine's dpi to become pixels
    /// again - a conversion that answers a different question on every monitor.
    /// </summary>
    public static (int Width, int Height)? SizeOf(DispatchObject picture)
    {
        try
        {
            var kind = FormDesignService.TryInt(picture, "Type") ?? 0;
            var handle = HandleOf(picture);
            if (handle == 0)
            {
                return null;
            }

            if (kind == PictureTypeBitmap)
            {
                return BitmapSize(handle);
            }

            if (kind != PictureTypeIcon || !GetIconInfo(handle, out var parts))
            {
                return null;
            }

            try
            {
                return IconSize(parts, out var width, out var height) ? (width, height) : null;
            }
            finally
            {
                if (parts.ColourBitmap != 0)
                {
                    _ = DeleteObject(parts.ColourBitmap);
                }

                if (parts.MaskBitmap != 0)
                {
                    _ = DeleteObject(parts.MaskBitmap);
                }
            }
        }
        catch (Exception why)
        {
            Log.Verbose($"picture: a picture would not measure ({why.Message.Trim()})");
            return null;
        }
    }

    /// <summary>
    /// The picture object on a property, or null when the property is absent or empty. The
    /// caller owns what comes back.
    /// </summary>
    public static DispatchObject? PictureOn(DispatchObject owner, string property)
    {
        try
        {
            return owner.GetDispId(property) == DispId.Unknown ? null : owner.GetObject(property);
        }
        catch (Exception why)
        {
            Log.Verbose($"picture: {property} would not answer ({why.Message.Trim()})");
            return null;
        }
    }

    /// <summary>
    /// A picture LOADED from a file, ready to assign. The caller owns it.
    ///
    /// TWO ROADS, and the first one is not enough - measured 2026-08-16, when a PNG came back
    /// 0x80004005 from the obvious call. OleLoadPicturePath is OLE's own loader and it reads what
    /// OLE has always read: BMP, GIF, JPEG, ICO, CUR, WMF, EMF. It does not know PNG, which is
    /// most of the pictures anybody has now. So anything it refuses goes to GDI+, which decodes
    /// every format this machine has a codec for, and the bitmap that comes out is wrapped as a
    /// picture with the same OLE call the first road ends in.
    ///
    /// The order matters and is not arbitrary: OLE's loader keeps an ICON an icon and a METAFILE
    /// a drawing, where GDI+ would flatten both to pixels. So the old road runs first and the new
    /// one catches what it drops.
    /// </summary>
    /// <param name="onto">
    /// The colour behind the picture, as an OLE colour. Only GDI+'s road uses it, and only for a
    /// picture with an ALPHA CHANNEL: an OLE picture is a bitmap, a bitmap has no alpha, so the
    /// transparent parts have to become some colour at load time. They become the colour of the
    /// control the picture is going onto, which is the one that makes them disappear.
    /// </param>
    public static DispatchObject FromPath(string path, int? onto = null)
    {
        var full = Path.GetFullPath(path);
        if (!File.Exists(full))
        {
            // A FOLDER is not a missing file, and saying so saves the developer looking for a
            // file that is sitting right there (measured 2026-08-16, passing a folder by mistake).
            throw new FileNotFoundException(
                Directory.Exists(full)
                    ? $"{full} is a folder, not a picture file"
                    : $"there is no file at {full}",
                full);
        }

        var iid = PictureDispId;
        var hr = OleLoadPicturePath(full, 0, 0, 0, ref iid, out var pointer);
        if (hr >= 0 && pointer != 0)
        {
            Log.Verbose($"picture: {Path.GetFileName(full)} loaded by OLE");
            return DispatchObject.Attach(pointer)
                ?? throw new InvalidOperationException($"{Path.GetFileName(full)} loaded as something unreadable");
        }

        Log.Verbose($"picture: OLE would not read {Path.GetFileName(full)} (0x{hr:x8}), trying GDI+");

        return FromPathThroughGdiPlus(full, onto)
            ?? throw new InvalidOperationException(
                $"{Path.GetFileName(full)} is not a picture this machine can load (0x{hr:x8})");
    }

    /// <summary>
    /// The second road: GDI+ decodes the file, hands over an HBITMAP, and OLE wraps it. Null when
    /// GDI+ will not read it either, which is the point at which the file really is not a picture.
    /// </summary>
    private static DispatchObject? FromPathThroughGdiPlus(string path, int? onto)
    {
        var startup = new GdiPlusStartup { Version = 1 };
        if (GdiplusStartup(out var token, ref startup, 0) != 0)
        {
            return null;
        }

        var image = nint.Zero;
        var bitmap = nint.Zero;
        try
        {
            if (GdipLoadImageFromFile(path, out image) != 0 || image == 0)
            {
                return null;
            }

            // ARGB, opaque, over the control's own colour - see the parameter's note. Without a
            // background a transparent PNG blends onto black, which on a grey form is a black box.
            var background = 0xFF000000u | (uint)(onto is { } ole ? FormDesignService.ColorRefToRgb(ole) : 0xFFFFFF);
            if (GdipCreateHBITMAPFromBitmap(image, out bitmap, background) != 0 || bitmap == 0)
            {
                return null;
            }

            var described = new PictureDescription
            {
                Size = sizeof(PictureDescription),
                Type = PictureTypeBitmap,
                Handle = bitmap,
                Palette = 0,
            };

            var iid = PictureDispId;
            // fOwn: the picture takes the bitmap over and frees it, so nothing here does.
            if (OleCreatePictureIndirect(ref described, ref iid, 1, out var pointer) < 0 || pointer == 0)
            {
                return null;
            }

            bitmap = nint.Zero;
            return DispatchObject.Attach(pointer);
        }
        finally
        {
            if (bitmap != 0)
            {
                _ = DeleteObject(bitmap);
            }

            if (image != 0)
            {
                _ = GdipDisposeImage(image);
            }

            // The HBITMAP outlives this: it is a GDI object, not a GDI+ one.
            GdiplusShutdown(token);
        }
    }

    /// <summary>IID_IPictureDisp - the dispatch face of a picture, which is what a control's
    /// Picture property takes and what this side can then read `Handle` off.</summary>
    private static readonly Guid PictureDispId = new("7bf80981-bf32-101a-8bbb-00aa00300cab");

    /// <summary>
    /// A GDI handle out of the picture. OLE_HANDLE is a 32-bit unsigned value even on 64-bit
    /// Windows - handles are guaranteed to fit - so it widens UNSIGNED: sign-extending a handle
    /// with its top bit set would hand GDI an address in the wrong half of the world.
    /// </summary>
    private static nint HandleOf(DispatchObject picture) =>
        FormDesignService.TryInt(picture, "Handle") is { } handle ? (nint)(uint)handle : 0;

    /// <summary>
    /// A whole BMP for an HBITMAP: the file header, the info header, and the pixels GetDIBits
    /// hands back. Top-down 32-bit, which needs no row-padding arithmetic - a 24-bit BMP pads
    /// every row to four bytes, and that is where hand-written BMPs go wrong.
    /// </summary>
    private static byte[]? BitmapBytes(nint bitmap)
    {
        var screen = GetDC(0);
        try
        {
            if (BitmapSize(bitmap) is not { } shape || !WorthInlining(shape.Width, shape.Height))
            {
                return null;
            }

            var (width, height) = shape;
            var pixels = new byte[checked(width * height * 4)];
            var info = TopDown32(width, height);
            fixed (byte* into = pixels)
            {
                if (GetDIBits(screen, bitmap, 0, (uint)height, (nint)into, ref info, 0) == 0)
                {
                    return null;
                }
            }

            return Wrap(pixels, width, height);
        }
        finally
        {
            _ = ReleaseDC(0, screen);
        }
    }

    /// <summary>
    /// A bitmap's dimensions. Called with no buffer and no bit count, GetDIBits FILLS IN the
    /// header rather than reading anything - the one call that asks a bitmap to describe itself.
    /// </summary>
    private static (int Width, int Height)? BitmapSize(nint bitmap)
    {
        var screen = GetDC(0);
        try
        {
            var info = new BitmapInfoHeader { Size = 40 };
            return GetDIBits(screen, bitmap, 0, 0, 0, ref info, 0) == 0 || info.Width <= 0 || info.Height == 0
                ? null
                : (info.Width, Math.Abs(info.Height));
        }
        finally
        {
            _ = ReleaseDC(0, screen);
        }
    }

    /// <summary>
    /// A whole BMP for an HICON, which GetDIBits cannot read: an icon is a colour bitmap and a
    /// mask together rather than one surface. So it is DRAWN - onto a 32-bit DIB section, by the
    /// same call the shell uses - and the section's own bits are the answer.
    /// </summary>
    private static byte[]? IconBytes(nint icon)
    {
        if (!GetIconInfo(icon, out var parts))
        {
            return null;
        }

        var screen = GetDC(0);
        var memory = CreateCompatibleDC(screen);
        var section = nint.Zero;
        var previous = nint.Zero;
        try
        {
            if (!IconSize(parts, out var width, out var height) || !WorthInlining(width, height))
            {
                return null;
            }

            var info = TopDown32(width, height);
            section = CreateDIBSection(memory, ref info, 0, out var bits, 0, 0);
            if (section == 0 || bits == 0)
            {
                return null;
            }

            previous = SelectObject(memory, section);
            // DI_NORMAL: the colour drawn through the mask, which is the icon as anything else
            // would show it. The section starts transparent, so what the mask leaves out stays out.
            if (!DrawIconEx(memory, 0, 0, icon, width, height, 0, 0, 0x0003))
            {
                return null;
            }

            // The drawing is finished into the section before its bits are read: GDI batches.
            _ = GdiFlush();

            var pixels = new byte[checked(width * height * 4)];
            Marshal.Copy(bits, pixels, 0, pixels.Length);
            return Wrap(pixels, width, height);
        }
        finally
        {
            if (previous != 0)
            {
                _ = SelectObject(memory, previous);
            }

            if (section != 0)
            {
                _ = DeleteObject(section);
            }

            _ = DeleteDC(memory);
            _ = ReleaseDC(0, screen);

            // GetIconInfo hands over two bitmaps and they are the caller's to free.
            if (parts.ColourBitmap != 0)
            {
                _ = DeleteObject(parts.ColourBitmap);
            }

            if (parts.MaskBitmap != 0)
            {
                _ = DeleteObject(parts.MaskBitmap);
            }
        }
    }

    /// <summary>
    /// An icon's size, from the bitmaps it is made of. A colour icon's is its colour bitmap; a
    /// MONOCHROME icon has no colour bitmap at all and its mask holds the image above the mask,
    /// stacked - so that one is half as tall as it measures.
    /// </summary>
    private static bool IconSize(IconInfo parts, out int width, out int height)
    {
        width = 0;
        height = 0;

        var source = parts.ColourBitmap != 0 ? parts.ColourBitmap : parts.MaskBitmap;
        if (source == 0)
        {
            return false;
        }

        var shape = default(BitmapHeader);
        if (GetObjectW(source, sizeof(BitmapHeader), (nint)(&shape)) == 0 || shape.Width <= 0 || shape.Height <= 0)
        {
            return false;
        }

        width = shape.Width;
        height = parts.ColourBitmap != 0 ? shape.Height : shape.Height / 2;
        return height > 0;
    }

    private static bool WorthInlining(int width, int height)
    {
        if ((long)width * height <= MostPixels)
        {
            return true;
        }

        Log.Verbose($"picture: {width}x{height} is past what the canvas inlines");
        return false;
    }

    /// <summary>The header that asks for TOP-DOWN 32-bit rows - the order a reader expects, and
    /// the one format where a row is exactly four bytes a pixel with nothing padding it.</summary>
    private static BitmapInfoHeader TopDown32(int width, int height) => new()
    {
        Size = 40,
        Width = width,
        Height = -height,
        Planes = 1,
        BitCount = 32,
        Compression = 0,
        SizeImage = width * height * 4,
    };

    /// <summary>
    /// The pixels wrapped in a BMP file - and made OPAQUE when nothing set an alpha.
    ///
    /// A 32-bit BMP's fourth byte is officially unused, and a browser given one whose alpha is
    /// zero throughout may draw nothing at all. GDI leaves it zero for every bitmap that has no
    /// alpha of its own, which is most of them. So a picture with no alpha ANYWHERE gets a solid
    /// one, and a picture that has real alpha - a 32-bit icon - keeps every byte of it.
    /// </summary>
    private static byte[] Wrap(byte[] pixels, int width, int height)
    {
        var transparent = true;
        for (var at = 3; at < pixels.Length; at += 4)
        {
            if (pixels[at] != 0)
            {
                transparent = false;
                break;
            }
        }

        if (transparent)
        {
            for (var at = 3; at < pixels.Length; at += 4)
            {
                pixels[at] = 0xFF;
            }
        }

        const int headers = 14 + 40;
        var file = new byte[headers + pixels.Length];
        file[0] = (byte)'B';
        file[1] = (byte)'M';
        WriteInt32(file, 2, file.Length);
        WriteInt32(file, 10, headers);
        WriteInt32(file, 14, 40);
        WriteInt32(file, 18, width);
        WriteInt32(file, 22, -height);
        file[26] = 1;
        file[28] = 32;
        WriteInt32(file, 34, pixels.Length);
        pixels.CopyTo(file, headers);
        return file;
    }

    private static void WriteInt32(byte[] into, int at, int value)
    {
        into[at] = (byte)value;
        into[at + 1] = (byte)(value >> 8);
        into[at + 2] = (byte)(value >> 16);
        into[at + 3] = (byte)(value >> 24);
    }

    /// <summary>BITMAPINFOHEADER, the forty bytes GDI reads and writes.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct BitmapInfoHeader
    {
        public int Size;
        public int Width;
        public int Height;
        public short Planes;
        public short BitCount;
        public int Compression;
        public int SizeImage;
        public int XPixelsPerMetre;
        public int YPixelsPerMetre;
        public int ColoursUsed;
        public int ColoursImportant;
    }

    /// <summary>BITMAP, what GetObject answers about a bitmap handle.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct BitmapHeader
    {
        public int Kind;
        public int Width;
        public int Height;
        public int BytesPerLine;
        public short Planes;
        public short BitCount;
        public nint Bits;
    }

    /// <summary>ICONINFO. The two bitmaps are the caller's to free, which is the part that leaks.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct IconInfo
    {
        public int IsIcon;
        public int HotspotX;
        public int HotspotY;
        public nint MaskBitmap;
        public nint ColourBitmap;
    }

    /// <summary>PICTDESC for a bitmap: the four fields OleCreatePictureIndirect reads.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct PictureDescription
    {
        public int Size;
        public int Type;
        public nint Handle;
        public nint Palette;
    }

    /// <summary>GdiplusStartupInput. Version 1 is the only one there has ever been.</summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct GdiPlusStartup
    {
        public uint Version;
        public nint DebugEventCallback;
        public int SuppressBackgroundThread;
        public int SuppressExternalCodecs;
    }

    [LibraryImport("oleaut32.dll", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int OleLoadPicturePath(
        string path, nint caller, uint reserved, uint reservedColour, ref Guid iid, out nint picture);

    [LibraryImport("oleaut32.dll")]
    private static partial int OleCreatePictureIndirect(
        ref PictureDescription described, ref Guid iid, int own, out nint picture);

    [LibraryImport("gdiplus.dll")]
    private static partial int GdiplusStartup(out nint token, ref GdiPlusStartup input, nint output);

    [LibraryImport("gdiplus.dll")]
    private static partial void GdiplusShutdown(nint token);

    [LibraryImport("gdiplus.dll", StringMarshalling = StringMarshalling.Utf16)]
    private static partial int GdipLoadImageFromFile(string path, out nint image);

    [LibraryImport("gdiplus.dll")]
    private static partial int GdipCreateHBITMAPFromBitmap(nint bitmap, out nint handle, uint background);

    [LibraryImport("gdiplus.dll")]
    private static partial int GdipDisposeImage(nint image);

    [LibraryImport("gdi32.dll")]
    private static partial int GetDIBits(
        nint dc, nint bitmap, uint start, uint lines, nint bits, ref BitmapInfoHeader info, uint usage);

    [LibraryImport("gdi32.dll")]
    private static partial nint CreateDIBSection(
        nint dc, ref BitmapInfoHeader info, uint usage, out nint bits, nint section, uint offset);

    [LibraryImport("gdi32.dll")]
    private static partial nint CreateCompatibleDC(nint dc);

    [LibraryImport("gdi32.dll")]
    private static partial nint SelectObject(nint dc, nint handle);

    [LibraryImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool DeleteObject(nint handle);

    [LibraryImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool DeleteDC(nint dc);

    [LibraryImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GdiFlush();

    [LibraryImport("gdi32.dll")]
    private static partial int GetObjectW(nint handle, int size, nint into);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetIconInfo(nint icon, out IconInfo parts);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool DrawIconEx(
        nint dc, int x, int y, nint icon, int width, int height, uint step, nint brush, uint flags);

    [LibraryImport("user32.dll")]
    private static partial nint GetDC(nint window);

    [LibraryImport("user32.dll")]
    private static partial int ReleaseDC(nint window, nint dc);
}
