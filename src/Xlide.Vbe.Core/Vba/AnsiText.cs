using System;
using System.Text;

namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// The encodings VBA text arrives in outside the editor.
///
/// The Visual Basic Editor is an ANSI program. A module it exports, and the module streams it
/// writes into a saved package, are bytes in a code page: the system's active one for an export,
/// and the one PROJECTCODEPAGE records for a file ([MS-OVBA] 2.3.4.2.1.5). Read as Latin-1 the
/// bytes come back as the wrong characters for anything past ASCII - a curly quote in a
/// description is 0x92 on a Western machine and Latin-1 makes that a control character - so the
/// text this product compared with the live module never matched, and a description typed in a
/// non-Latin script was written into the file as question marks. Everything that touches such
/// bytes decodes and encodes with the code page they were written in, through here.
///
/// Latin-1 stays the fallback: every byte maps to a character and back, so a file whose code page
/// the machine cannot name is still carried through byte for byte where it is not edited.
/// </summary>
public static class AnsiText
{
    private static readonly object Gate = new();
    private static bool registered;

    /// <summary>The encoding for a Windows code page, or Latin-1 when it is zero or unknown here.</summary>
    public static Encoding For(int codePage)
    {
        if (codePage <= 0)
        {
            return Encoding.Latin1;
        }

        lock (Gate)
        {
            if (!registered)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                registered = true;
            }
        }

        try
        {
            return Encoding.GetEncoding(codePage);
        }
        catch (ArgumentException)
        {
            return Encoding.Latin1;
        }
        catch (NotSupportedException)
        {
            return Encoding.Latin1;
        }
    }
}
