using System.Buffers.Binary;
using System.IO.Compression;

namespace Xlide.Vbe.Core.Forms;

/// <summary>
/// Which properties a form's controls could possibly have had changed, read out of the workbook
/// Excel has ALREADY saved. No export, no COM, no running Excel: the `.xlsm` is a ZIP,
/// `xl/vbaProject.bin` inside it is a compound file, and MSForms persists only properties that
/// differ from the file format default - saying which ones in a PropMask at a fixed offset in
/// every control's block ([MS-OFORMS] 2.2.10).
///
/// A SET BIT DOES NOT MEAN THE DEVELOPER CHANGED IT, and the whole design turns on that. The mask
/// measures against the FILE FORMAT default, so where a control KIND is born with something else
/// the bit is set on controls nobody touched: every control on a Tahoma form carries `FontName`,
/// because the file's default is MS Sans Serif, and every CheckBox, OptionButton and ToggleButton
/// carries `BackColor` and `ForeColor` whatever the developer did.
///
/// So this NARROWS the walk rather than answering for it. What comes back is the short list of
/// properties worth asking a control about; <see cref="FormDesignService"/> then reads only those
/// and compares against the bare coclass exactly as it already does, which filters the
/// file-format noise back out. Reading fifty properties of every control on every projection is
/// the cost that comparison was avoiding.
///
/// The harness twin is tools\harness\saved-design.mjs, where the road was proven against the
/// fixture first, and it is the thing to check this against when either changes.
/// </summary>
public sealed class SavedDesign
{
    /// <summary>Where a workbook keeps its VBA project inside the package.</summary>
    private const string ProjectPath = "xl/vbaProject.bin";

    /// <summary>A ceiling on the project stream, so a hostile or corrupt package cannot ask this
    /// to allocate the machine. Real projects are kilobytes; a form full of pictures is a few
    /// megabytes.</summary>
    private const int MostProjectBytes = 64 * 1024 * 1024;

    private static readonly Dictionary<string, (DateTime Stamp, SavedDesign? Design)> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Form name to control path (`Options.PickGround`) to the properties its masks
    /// name. A control absent from the map has nothing recorded, which is a real answer.</summary>
    private readonly Dictionary<string, Dictionary<string, string[]>> forms =
        new(StringComparer.OrdinalIgnoreCase);

    private SavedDesign()
    {
    }

    /// <summary>
    /// The baseline for a saved workbook, or null when there is none to have - a workbook never
    /// saved, a path that will not open, a package with no VBA project, or bytes this cannot
    /// walk. Every one of those is a form that keeps today's bare-coclass answer.
    ///
    /// Cached per path and invalidated by the file's own write time, which is what makes Ctrl+S
    /// re-read without anything having to tell it to.
    /// </summary>
    public static SavedDesign? For(string? workbookPath)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
        {
            return null;
        }

        try
        {
            if (!File.Exists(workbookPath))
            {
                return null;
            }

            var stamp = File.GetLastWriteTimeUtc(workbookPath);
            if (Cache.TryGetValue(workbookPath, out var held) && held.Stamp == stamp)
            {
                return held.Design;
            }

            var design = Read(workbookPath);
            Cache[workbookPath] = (stamp, design);
            return design;
        }
        catch
        {
            // A workbook held, moved, encrypted or half-written is a form with no saved baseline,
            // which the projection already knows how to be. The caller does the reporting: this
            // half of the product has no host to report to.
            return null;
        }
    }

    /// <summary>Whether anything was recorded for this form at all. A form added since the last
    /// save is not in the file, and asking about its controls would answer "nothing changed" for
    /// every one of them - which is the opposite of the truth.</summary>
    public bool Knows(string formName) => forms.ContainsKey(formName);

    /// <summary>
    /// The properties worth asking about for one control, by its path under the form:
    /// `PickGround` on the form itself, `Options.PickGround` inside a Frame. Empty means the file
    /// recorded nothing, which is a control left entirely alone.
    /// </summary>
    public IReadOnlyList<string> ChangedOn(string formName, string controlPath) =>
        forms.TryGetValue(formName, out var controls)
            && controls.TryGetValue(controlPath, out var names)
                ? names
                : [];

    /// <summary>Every control path the file recorded anything for, `Options.PickGround` style.
    /// The api reads this to show a whole form's baseline at once, and it is what a test walks to
    /// assert that nothing structural leaked into the answer.</summary>
    public IEnumerable<string> Controls(string formName) =>
        forms.TryGetValue(formName, out var controls) ? controls.Keys : [];

    private static SavedDesign? Read(string workbookPath)
    {
        var project = ProjectBytes(workbookPath);
        if (project is null || CompoundFile.TryRead(project) is not { } cfb)
        {
            return null;
        }

        var design = new SavedDesign();
        foreach (var (path, entry) in cfb.Paths)
        {
            // A form's storage sits directly under the root, beside /VBA, and holds an `f`.
            if (entry.Type != CompoundFile.StorageType
                || path.LastIndexOf('/') != 0
                || !cfb.Paths.ContainsKey($"{path}/f"))
            {
                continue;
            }

            var changed = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            ReadStorage(cfb, path, changed, string.Empty, 0);
            design.forms[path[1..]] = changed;
        }

        return design;
    }

    /// <summary>xl/vbaProject.bin out of the package, or null when the workbook has no VBA in
    /// it. Opened read-only and shared, because Excel has the file open.</summary>
    private static byte[]? ProjectBytes(string workbookPath)
    {
        using var file = new FileStream(
            workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var zip = new ZipArchive(file, ZipArchiveMode.Read);
        var entry = zip.GetEntry(ProjectPath);
        if (entry is null || entry.Length is <= 0 or > MostProjectBytes)
        {
            return null;
        }

        var bytes = new byte[entry.Length];
        using var stream = entry.Open();
        stream.ReadExactly(bytes);
        return bytes;
    }

    /* ---- the walk, [MS-OFORMS] 2.2.10 ------------------------------------------------------ */

    /// <summary>
    /// Every control under one storage, and then the storages its containers own. A container is
    /// named `i` plus its site ID - `Options` with ID 6 lives in `i06` - which is measured rather
    /// than assumed: the fixture's four containers land on their four storages exactly.
    /// </summary>
    private static void ReadStorage(
        CompoundFile cfb, string prefix, Dictionary<string, string[]> into, string path, int depth)
    {
        if (depth > 16)
        {
            return;
        }

        var f = cfb.Read($"{prefix}/f");
        if (f.Length < 8)
        {
            return;
        }

        var blocks = cfb.Read($"{prefix}/o");
        var (formNames, sites) = ReadSites(f);

        void Record(string where, IEnumerable<string> names)
        {
            // A container is described TWICE - once as a site in its parent, once by the
            // FormPropMask in its own storage - so the two merge rather than replacing.
            var all = new SortedSet<string>(
                into.TryGetValue(where, out var held) ? held : [], StringComparer.Ordinal);
            foreach (var name in names)
            {
                all.Add(name);
            }

            if (all.Count > 0)
            {
                into[where] = [.. all];
            }
        }

        if (path.Length > 0)
        {
            Record(path, formNames);
        }

        var at = 0;
        foreach (var site in sites)
        {
            // A MultiPage keeps its own TabStrip as a NAMELESS site: the container's chrome
            // rather than anything a developer put there, and nothing can ask about a control
            // with no name.
            if (string.IsNullOrEmpty(site.Name))
            {
                at += site.ObjectStreamSize;
                continue;
            }

            var where = path.Length > 0 ? $"{path}.{site.Name}" : site.Name;
            var changed = new SortedSet<string>(site.Changed, StringComparer.Ordinal);

            if (site.ObjectStreamSize > 0
                && site.Kind is { } kind
                && Masks.TryGetValue(kind, out var shape)
                && at + 8 <= blocks.Length)
            {
                var cb = BinaryPrimitives.ReadUInt16LittleEndian(blocks.AsSpan(at + 2));
                var mask = shape.Wide && at + 12 <= blocks.Length
                    ? BinaryPrimitives.ReadUInt64LittleEndian(blocks.AsSpan(at + 4))
                    : BinaryPrimitives.ReadUInt32LittleEndian(blocks.AsSpan(at + 4));
                foreach (var name in Named(mask, shape.Bits))
                {
                    changed.Add(name);
                }

                // TextProps trails the block and carries the font. Only reach for it when nothing
                // of unmeasured length lies between: a picture's bytes are not walked here, and a
                // font read at the wrong offset is worse than one not read at all.
                if (shape.Text && !changed.Contains("Picture") && !changed.Contains("MouseIcon"))
                {
                    var props = at + 4 + cb;
                    if (props + 8 <= Math.Min(blocks.Length, at + site.ObjectStreamSize)
                        && blocks[props + 1] == 2)
                    {
                        var text = BinaryPrimitives.ReadUInt32LittleEndian(blocks.AsSpan(props + 4));
                        foreach (var name in Named(text, TextPropsBits))
                        {
                            changed.Add(name);
                        }
                    }
                }
            }

            Record(where, changed);
            at += site.ObjectStreamSize;

            if (site.Kind is "Form" or "Frame" or "MultiPage")
            {
                ReadStorage(cfb, $"{prefix}/i{site.Id:D2}", into, where, depth + 1);
            }
        }
    }

    private readonly record struct Site(
        string? Name, uint Id, int ObjectStreamSize, string? Kind, string[] Changed, int Length);

    private static (string[] Form, List<Site> Sites) ReadSites(byte[] f)
    {
        var cbForm = BinaryPrimitives.ReadUInt16LittleEndian(f.AsSpan(2));
        var formNames = Named(BinaryPrimitives.ReadUInt32LittleEndian(f.AsSpan(4)), FormBits);

        // StreamData holds the form's mouse icon, font and picture as GUID-prefixed blobs. Rather
        // than decode three OLE structures to skip over them, find FormSiteData by its own
        // arithmetic: the pair (CountOfSites, CountOfBytes) whose count covers what follows.
        var at = -1;
        for (var probe = 4 + cbForm; probe + 8 < f.Length; probe++)
        {
            var count = BinaryPrimitives.ReadUInt32LittleEndian(f.AsSpan(probe));
            var size = BinaryPrimitives.ReadUInt32LittleEndian(f.AsSpan(probe + 4));
            if (count is > 0 and < 1000 && probe + 8 + size <= f.Length)
            {
                at = probe;
                break;
            }
        }

        if (at < 0)
        {
            return (formNames, []);
        }

        var sites = (int)BinaryPrimitives.ReadUInt32LittleEndian(f.AsSpan(at));
        var p = at + 8;

        // SiteDepthsAndTypes: a FormObjectDepthTypeCount per run, two bytes, or three when the
        // high bit of the second says the byte after it is a type shared by that many sites.
        var start = p;
        for (var seen = 0; seen < sites && p + 1 < f.Length;)
        {
            var typeOrCount = f[p + 1];
            if ((typeOrCount & 0x80) != 0)
            {
                seen += typeOrCount & 0x7F;
                p += 3;
            }
            else
            {
                seen += 1;
                p += 2;
            }
        }

        p += (4 - ((p - start) % 4)) % 4;                       // ArrayPadding

        var found = new List<Site>(sites);
        for (var i = 0; i < sites && p + 4 <= f.Length; i++)
        {
            var site = ReadSite(f, p);
            found.Add(site);
            p += site.Length;
        }

        return (formNames, found);
    }

    /// <summary>
    /// One OleSiteConcreteControl: `Version(2) cbSite(2) PropMask(4)` and two blocks. Alignment
    /// inside it is measured from the START of the record, which is what the spec's "from the
    /// beginning of the version number" means, so every read goes through the same cursor.
    /// </summary>
    private static Site ReadSite(byte[] f, int start)
    {
        var cbSite = BinaryPrimitives.ReadUInt16LittleEndian(f.AsSpan(start + 2));
        var mask = BinaryPrimitives.ReadUInt32LittleEndian(f.AsSpan(start + 4));
        var p = start + 8;
        var end = Math.Min(f.Length, start + 4 + cbSite);

        bool Has(int bit) => (mask & (1u << bit)) != 0;

        uint Read(int width)
        {
            var over = (p - start) % width;
            if (over != 0)
            {
                p += width - over;
            }

            if (p + width > end)
            {
                p = end;
                return 0;
            }

            var value = width == 2
                ? BinaryPrimitives.ReadUInt16LittleEndian(f.AsSpan(p))
                : BinaryPrimitives.ReadUInt32LittleEndian(f.AsSpan(p));
            p += width;
            return value;
        }

        var nameData = Has(0) ? Read(4) : 0;
        var tagData = Has(1) ? Read(4) : 0;
        var id = Has(2) ? Read(4) : 0;
        if (Has(3)) { Read(4); }                                // HelpContextID
        if (Has(4)) { Read(4); }                                // BitFlags
        var objectStreamSize = Has(5) ? (int)Read(4) : 0;
        if (Has(6)) { Read(2); }                                // TabIndex
        var clsid = Has(7) ? (int)Read(2) : -1;
        if (Has(9)) { Read(2); }                                // GroupID
        var tipData = Has(11) ? Read(4) : 0;
        var licData = Has(12) ? Read(4) : 0;
        var sourceData = Has(13) ? Read(4) : 0;
        var rowData = Has(14) ? Read(4) : 0;

        var tail = (p - (start + 8)) % 4;                       // the DataBlock's own tail
        if (tail != 0)
        {
            p += 4 - tail;
        }

        string? Text(uint data)
        {
            if (data == 0)
            {
                return null;
            }

            var cb = (int)(data & 0x7FFFFFFF);
            if (cb <= 0 || p + cb > end)
            {
                p = end;
                return null;
            }

            var text = (data >> 31) == 1
                ? System.Text.Encoding.Latin1.GetString(f, p, cb)
                : System.Text.Encoding.Unicode.GetString(f, p, cb);
            p += cb + ((4 - (cb % 4)) % 4);                     // strings pad to a multiple of 4
            return text;
        }

        var name = Text(nameData);
        Text(tagData);
        if (Has(8)) { p += 8; }                                 // SitePosition
        Text(tipData);
        Text(licData);
        Text(sourceData);
        Text(rowData);

        return new Site(
            name, id, objectStreamSize,
            Kinds.TryGetValue(clsid, out var kind) ? kind : null,
            Named(mask, SiteBits),
            4 + cbSite);
    }

    private static string[] Named(ulong mask, (int Bit, string Name)[] bits) =>
        [.. bits
            .Where(one => (mask & (1UL << one.Bit)) != 0 && !Structural.Contains(one.Name))
            .Select(one => one.Name)];

    /* ---- the tables, [MS-OFORMS] 2.2.x and 2.4.2 -------------------------------------------
     *
     * Bit numbering is the spec's own: the lettered fields run upwards from bit 0, and a
     * multi-bit unused run shifts what follows it - FormPropMask's Unused2 is two bits wide,
     * which is why fBooleanProperties is bit 6 rather than bit 5.
     */

    /// <summary>The names a developer never sets: bookkeeping, the MUST-be-1 bits, the fields
    /// carrying a container's own children, and TabIndex - which is set on every control but the
    /// first, so it distinguishes nothing, and already rides its own row.</summary>
    private static readonly HashSet<string> Structural = new(StringComparer.Ordinal)
    {
        "Size", "NewVersion", "Reserved", "DrawBuffer", "NextAvailableID", "ShapeCookie",
        "GroupCnt", "LogicalSize", "DisplayedSize", "ScrollPosition", "ObjectStreamSize",
        "ID", "ClsidCacheIndex", "BitFlags", "Position", "GroupID", "Name", "DisplayStyle",
        "cColumnInfo", "TabsAllocated", "TabData", "Items", "Names", "Tags", "TipStrings",
        "Accelerators", "ListIndex", "PrevEnabled", "NextEnabled", "RuntimeLicKey", "TabIndex",
    };

    private static readonly (int, string)[] FormBits =
    [
        (1, "BackColor"), (2, "ForeColor"), (3, "NextAvailableID"), (6, "BooleanProperties"),
        (7, "BorderStyle"), (8, "MousePointer"), (9, "ScrollBars"), (10, "DisplayedSize"),
        (11, "LogicalSize"), (12, "ScrollPosition"), (13, "GroupCnt"), (15, "MouseIcon"),
        (16, "Cycle"), (17, "SpecialEffect"), (18, "BorderColor"), (19, "Caption"), (20, "Font"),
        (21, "Picture"), (22, "Zoom"), (23, "PictureAlignment"), (24, "PictureTiling"),
        (25, "PictureSizeMode"), (26, "ShapeCookie"), (27, "DrawBuffer"),
    ];

    private static readonly (int, string)[] SiteBits =
    [
        (0, "Name"), (1, "Tag"), (2, "ID"), (3, "HelpContextID"), (4, "BitFlags"),
        (5, "ObjectStreamSize"), (6, "TabIndex"), (7, "ClsidCacheIndex"), (8, "Position"),
        (9, "GroupID"), (11, "ControlTipText"), (12, "RuntimeLicKey"), (13, "ControlSource"),
        (14, "RowSource"),
    ];

    private static readonly (int, string)[] MorphDataBits =
    [
        (0, "VariousPropertyBits"), (1, "BackColor"), (2, "ForeColor"), (3, "MaxLength"),
        (4, "BorderStyle"), (5, "ScrollBars"), (6, "DisplayStyle"), (7, "MousePointer"),
        (8, "Size"), (9, "PasswordChar"), (10, "ListWidth"), (11, "BoundColumn"),
        (12, "TextColumn"), (13, "ColumnCount"), (14, "ListRows"), (15, "cColumnInfo"),
        (16, "MatchEntry"), (17, "ListStyle"), (18, "ShowDropButtonWhen"), (20, "DropButtonStyle"),
        (21, "MultiSelect"), (22, "Value"), (23, "Caption"), (24, "PicturePosition"),
        (25, "BorderColor"), (26, "SpecialEffect"), (27, "MouseIcon"), (28, "Picture"),
        (29, "Accelerator"), (31, "Reserved"), (32, "GroupName"),
    ];

    private static readonly (int, string)[] LabelBits =
    [
        (0, "ForeColor"), (1, "BackColor"), (2, "VariousPropertyBits"), (3, "Caption"),
        (4, "PicturePosition"), (5, "Size"), (6, "MousePointer"), (7, "BorderColor"),
        (8, "BorderStyle"), (9, "SpecialEffect"), (10, "Picture"), (11, "Accelerator"),
        (12, "MouseIcon"),
    ];

    private static readonly (int, string)[] CommandButtonBits =
    [
        (0, "ForeColor"), (1, "BackColor"), (2, "VariousPropertyBits"), (3, "Caption"),
        (4, "PicturePosition"), (5, "Size"), (6, "MousePointer"), (7, "Picture"),
        (8, "Accelerator"), (9, "TakeFocusOnClick"), (10, "MouseIcon"),
    ];

    private static readonly (int, string)[] ImageBits =
    [
        (2, "AutoSize"), (3, "BorderColor"), (4, "BackColor"), (5, "BorderStyle"),
        (6, "MousePointer"), (7, "PictureSizeMode"), (8, "SpecialEffect"), (9, "Size"),
        (10, "Picture"), (11, "PictureAlignment"), (12, "PictureTiling"),
        (13, "VariousPropertyBits"), (14, "MouseIcon"),
    ];

    private static readonly (int, string)[] TabStripBits =
    [
        (0, "ListIndex"), (1, "BackColor"), (2, "ForeColor"), (4, "Size"), (5, "Items"),
        (6, "MousePointer"), (8, "TabOrientation"), (9, "TabStyle"), (10, "MultiRow"),
        (11, "TabFixedWidth"), (12, "TabFixedHeight"), (13, "Tooltips"), (15, "TipStrings"),
        (17, "Names"), (18, "VariousPropertyBits"), (19, "NewVersion"), (20, "TabsAllocated"),
        (21, "Tags"), (22, "TabData"), (23, "Accelerators"), (24, "MouseIcon"),
    ];

    private static readonly (int, string)[] ScrollBarBits =
    [
        (0, "ForeColor"), (1, "BackColor"), (2, "VariousPropertyBits"), (3, "Size"),
        (4, "MousePointer"), (5, "Min"), (6, "Max"), (7, "Position"), (9, "PrevEnabled"),
        (10, "NextEnabled"), (11, "SmallChange"), (12, "LargeChange"), (13, "Orientation"),
        (14, "ProportionalThumb"), (15, "Delay"), (16, "MouseIcon"),
    ];

    private static readonly (int, string)[] SpinButtonBits =
    [
        (0, "ForeColor"), (1, "BackColor"), (2, "VariousPropertyBits"), (3, "Size"),
        (5, "Min"), (6, "Max"), (7, "Position"), (8, "PrevEnabled"), (9, "NextEnabled"),
        (10, "SmallChange"), (11, "Orientation"), (12, "Delay"), (13, "MouseIcon"),
        (14, "MousePointer"),
    ];

    private static readonly (int, string)[] TextPropsBits =
    [
        (0, "FontName"), (1, "FontEffects"), (2, "FontHeight"), (4, "FontCharSet"),
        (5, "FontPitchAndFamily"), (6, "ParagraphAlign"), (7, "FontWeight"),
    ];

    /// <summary>What a ClsidCacheIndex means, [MS-OFORMS] 2.4.2.</summary>
    private static readonly Dictionary<int, string> Kinds = new()
    {
        [7] = "Form", [12] = "Image", [14] = "Frame", [15] = "MorphData", [16] = "SpinButton",
        [17] = "CommandButton", [18] = "TabStrip", [21] = "Label", [23] = "TextBox",
        [24] = "ListBox", [25] = "ComboBox", [26] = "CheckBox", [27] = "OptionButton",
        [28] = "ToggleButton", [47] = "ScrollBar", [57] = "MultiPage",
    };

    /// <summary>Six kinds share one structure and therefore one mask, which is eight bytes wide
    /// rather than four - the only one that is.</summary>
    private static readonly Dictionary<string, ((int, string)[] Bits, bool Wide, bool Text)> Masks =
        new(StringComparer.Ordinal)
        {
            ["Image"] = (ImageBits, false, false),
            ["SpinButton"] = (SpinButtonBits, false, false),
            ["CommandButton"] = (CommandButtonBits, false, true),
            ["TabStrip"] = (TabStripBits, false, true),
            ["Label"] = (LabelBits, false, true),
            ["ScrollBar"] = (ScrollBarBits, false, false),
            ["TextBox"] = (MorphDataBits, true, true),
            ["ListBox"] = (MorphDataBits, true, true),
            ["ComboBox"] = (MorphDataBits, true, true),
            ["CheckBox"] = (MorphDataBits, true, true),
            ["OptionButton"] = (MorphDataBits, true, true),
            ["ToggleButton"] = (MorphDataBits, true, true),
        };
}
