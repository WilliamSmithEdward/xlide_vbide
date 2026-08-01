namespace Xlide.Vbe.Core.Hosting;

/// <summary>
/// A rectangle in device pixels, with the same edge convention Win32 uses: left and top are
/// inclusive, right and bottom are exclusive.
///
/// This exists so the arithmetic an OLE container forces on us can be tested without a container.
/// A sited control is told where to sit by two rectangles at once, and getting their relationship
/// wrong produces a window that is subtly misplaced rather than an error.
/// </summary>
public readonly record struct PixelRect(int Left, int Top, int Right, int Bottom)
{
    /// <summary>Width in pixels, clamped at zero so an inverted rectangle cannot produce a negative size.</summary>
    public int Width => Right > Left ? Right - Left : 0;

    /// <summary>Height in pixels, clamped at zero.</summary>
    public int Height => Bottom > Top ? Bottom - Top : 0;

    /// <summary>True when the rectangle encloses no pixels.</summary>
    public bool IsEmpty => Width == 0 || Height == 0;

    /// <summary>Builds a rectangle from an origin and a size.</summary>
    public static PixelRect FromSize(int left, int top, int width, int height) =>
        new(left, top, left + width, top + height);

    /// <summary>
    /// The overlap of two rectangles, or an empty rectangle anchored at this one's origin when they
    /// do not overlap. Never returns an inverted rectangle, because Win32 sizing calls accept one
    /// and produce nonsense rather than failing.
    /// </summary>
    public PixelRect Intersect(PixelRect other)
    {
        var left = Math.Max(Left, other.Left);
        var top = Math.Max(Top, other.Top);
        var right = Math.Min(Right, other.Right);
        var bottom = Math.Min(Bottom, other.Bottom);

        return right <= left || bottom <= top
            ? new PixelRect(Left, Top, Left, Top)
            : new PixelRect(left, top, right, bottom);
    }

    /// <summary>
    /// The same size anchored at the origin. A hosted browser is positioned in its parent's client
    /// coordinates, which always start at zero regardless of where the parent sits in the container.
    /// </summary>
    public PixelRect AtOrigin() => new(0, 0, Width, Height);
}
