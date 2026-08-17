using System.Globalization;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// How a colour is SPELLED for the developer, and read back from them.
///
/// The Properties panel had been showing OLE_COLOR the way the VBE spells it - `&amp;H8000000F&amp;`,
/// eight hex digits, blue-green-red order, the high bit meaning "ask the system" - which is the
/// language a form's binary speaks and not one anybody reads (the owner, 2026-08-16: "for color
/// pickers in properties, can you use #f0f0f0 format?").
///
/// So a plain colour reads `#rrggbb`, the spelling every developer already knows, byte order and
/// all. A SYSTEM colour does not: `&amp;H8000000F&amp;` is not a colour at all but a question - what
/// does this machine call a button face today - and answering it with `#f0f0f0` would freeze
/// today's answer into the form. Those rows read the name instead (`Button Face`), and the swatch
/// beside them paints what the machine currently answers.
///
/// The names are Win32's own COLOR_ constants in plain words, which is the one table here that
/// could not be measured: there is no call that hands back a display name for COLOR_BTNFACE. The
/// VALUE is always measured - GetSysColor, through the same path the canvas paints with - so the
/// only thing written down is the wording, and it is the wording the native picker's System tab
/// uses.
/// </summary>
internal static class SystemColours
{
    /// <summary>The bit that turns an OLE_COLOR from a colour into a question for the system.</summary>
    private const int SystemBit = unchecked((int)0x80000000);

    private static readonly (int Index, string Name)[] Named =
    [
        (0, "Scroll Bars"),
        (1, "Desktop"),
        (2, "Active Title Bar"),
        (3, "Inactive Title Bar"),
        (4, "Menu Bar"),
        (5, "Window Background"),
        (6, "Window Frame"),
        (7, "Menu Text"),
        (8, "Window Text"),
        (9, "Active Title Bar Text"),
        (10, "Active Border"),
        (11, "Inactive Border"),
        (12, "Application Workspace"),
        (13, "Highlight"),
        (14, "Highlight Text"),
        (15, "Button Face"),
        (16, "Button Shadow"),
        (17, "Disabled Text"),
        (18, "Button Text"),
        (19, "Inactive Title Bar Text"),
        (20, "Button Highlight"),
        (21, "Button Dark Shadow"),
        (22, "Button Light Shadow"),
        (23, "Tooltip Text"),
        (24, "Tooltip"),
        (26, "Hot-Tracked Item"),
        (27, "Gradient Active Title Bar"),
        (28, "Gradient Inactive Title Bar"),
        (29, "Menu Highlight"),
        (30, "Menu Bar Background"),
    ];

    /// <summary>Every system colour, with the value that ASKS for it and what the machine answers
    /// today - the System half of the picker, and the reason it can show colours at all.</summary>
    public static IReadOnlyList<(string Name, string Value, string Css)> All =>
        [.. Named.Select(one => (
            one.Name,
            Spell(SystemBit | one.Index),
            FormDesignService.OleColorToCss(SystemBit | one.Index)))];

    /// <summary>An OLE colour as the panel shows it: a system colour's name, or `#rrggbb`. A
    /// system index this does not name keeps the VBE's own hex, which is honest - a row that
    /// cannot say what a value means says what it is.</summary>
    public static string Spell(int ole)
    {
        if ((ole & SystemBit) != 0)
        {
            var index = ole & 0xFF;
            var named = Named.FirstOrDefault(one => one.Index == index);
            return named.Name ?? Core.Forms.FormMarkup.SpellColour(ole);
        }

        // The document's own spelling, from the document's own arithmetic: the panel and the
        // markup say `#c0dcc0` for the same colour because they say it through one function.
        return Core.Forms.FormMarkup.SpellColour(ole);
    }

    /// <summary>
    /// What the developer typed or picked, back as the number the model stores. Every spelling
    /// the panel can show plus the two the VBE speaks: `#rrggbb`, a system colour's name,
    /// `&amp;Hbbggrr&amp;`, or a plain number. Null when it is none of them, which leaves the
    /// caller to pass the text through untouched rather than write a colour nobody asked for.
    /// </summary>
    public static int? Unspell(string shown)
    {
        var text = shown.Trim();
        if (text.Length == 0)
        {
            return null;
        }

        if (text[0] == '#')
        {
            // The markup's own reader: one conversion, so a colour typed into a row and the same
            // colour typed into a document cannot disagree about which byte is red.
            return Core.Forms.FormMarkup.ReadColour(text);
        }

        var byName = Named.FirstOrDefault(one =>
            string.Equals(one.Name, text, StringComparison.OrdinalIgnoreCase));
        if (byName.Name is not null)
        {
            return SystemBit | byName.Index;
        }

        if (text.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            var digits = text.TrimEnd('&')[2..];
            return uint.TryParse(digits, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var hex)
                ? (int)hex
                : null;
        }

        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var plain)
            ? plain
            : null;
    }
}
