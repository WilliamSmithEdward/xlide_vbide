using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The font faces THIS machine has, so a Font.Name row can offer them instead of asking the
/// developer to spell one (the owner, 2026-08-16: "should font be a drop down selector for all
/// properties?").
///
/// Measured, for the same reason the system colours are: no list written down here could be true
/// on somebody else's machine, and a picker offering a font that is not installed is worse than a
/// text box. GDI's own enumeration is the answer - the same call every native font picker makes -
/// and it is read once per session, because installing a font mid-session is not a thing this
/// product needs to follow.
///
/// The row stays TYPEABLE around the list (the panel's own rule, from the enum rows): MSForms
/// stores a face as a plain string, and a form written on another machine may legitimately name a
/// font this one does not have. Offering the list must not turn that value into an error.
/// </summary>
internal static unsafe partial class InstalledFonts
{
    /// <summary>Every charset, which is what enumerates the families rather than one script's.</summary>
    private const byte DefaultCharSet = 1;

    private static readonly object Gate = new();

    /// <summary>The answer, once. Null until the first ask.</summary>
    private static string[]? _all;

    /// <summary>Where the callback puts what it finds: an UnmanagedCallersOnly method captures
    /// nothing, so the collection is static and the lock below is what keeps it one walk at a
    /// time.</summary>
    private static readonly SortedSet<string> Found = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The installed families, sorted the way a list is read. Cached: the walk costs a few
    /// milliseconds and the answer cannot change usefully within a session.
    /// </summary>
    public static IReadOnlyList<string> All
    {
        get
        {
            lock (Gate)
            {
                if (_all is not null)
                {
                    return _all;
                }

                var dc = GetDC(0);
                try
                {
                    var wanted = new LogFont { CharacterSet = DefaultCharSet };
                    _ = EnumFontFamiliesExW(dc, &wanted, &OnFamily, 0, 0);
                }
                catch (Exception why)
                {
                    // A panel that cannot list the fonts still shows the row: the field takes any
                    // face name typed into it, which is what it did before this existed.
                    Log.Error($"fonts: the family walk failed ({why.Message})");
                }
                finally
                {
                    _ = ReleaseDC(0, dc);
                }

                _all = [.. Found];
                Found.Clear();
                Log.Verbose($"fonts: {_all.Length} families on this machine");
                return _all;
            }
        }
    }

    /// <summary>
    /// One family from GDI. The face name is the first field of the LOGFONTW the callback is
    /// handed, NUL-terminated inside its 32-character room.
    ///
    /// A leading '@' is a VERTICAL variant of a CJK family - the same font rotated for vertical
    /// writing - which every native font picker hides, and which would double the length of this
    /// list on a machine with East Asian fonts installed.
    /// </summary>
    [UnmanagedCallersOnly]
    private static int OnFamily(LogFont* font, void* metrics, uint type, nint lParam)
    {
        try
        {
            var name = new string(font->FaceName);
            if (name.Length > 0 && name[0] != '@')
            {
                Found.Add(name);
            }
        }
        catch
        {
            // Never throw across a native callback: GDI is walking its own list and an exception
            // unwinding through it is a process-level fault, not a missing font.
        }

        // Anything but zero means "keep going".
        return 1;
    }

    /// <summary>LOGFONTW, of which only the charset is set going in and only the face name is read
    /// coming out. The layout must match exactly all the same, because GDI walks it by offset.</summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct LogFont
    {
        public int Height;
        public int Width;
        public int Escapement;
        public int Orientation;
        public int Weight;
        public byte Italic;
        public byte Underline;
        public byte StrikeOut;
        public byte CharacterSet;
        public byte OutPrecision;
        public byte ClipPrecision;
        public byte Quality;
        public byte PitchAndFamily;
        public fixed char FaceName[32];
    }

    [LibraryImport("gdi32.dll")]
    private static partial int EnumFontFamiliesExW(
        nint dc, LogFont* wanted, delegate* unmanaged<LogFont*, void*, uint, nint, int> callback,
        nint param, uint flags);

    [LibraryImport("user32.dll")]
    private static partial nint GetDC(nint window);

    [LibraryImport("user32.dll")]
    private static partial int ReleaseDC(nint window, nint dc);
}
