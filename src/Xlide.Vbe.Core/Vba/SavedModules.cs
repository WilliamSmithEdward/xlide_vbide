using System.Buffers.Binary;
using System.Text;
using Xlide.Vbe.Core.Forms;

namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// The attribute header of every module, read out of the document already SAVED on disk.
///
/// WHY THIS EXISTS. `Attribute VB_PredeclaredId = True` gives a class module a default instance,
/// which makes its own name usable as a value: `Ticket.ChangeTest` compiles against a predeclared
/// class and is `Variable not defined` against a plain one. The analyzer needs the flag to tell
/// those apart (xlide_vscode#47) and CANNOT SEE IT: a `CodeModule` returns code, and every
/// `Attribute VB_` line lives outside the code pane. The object model has no property for it
/// either. The only other way to get at it is to Export the module to a temporary file and read
/// the header back, and writing files to read them is not something this product does outside the
/// import/export feature the developer asked for (owner, 2026-08-23).
///
/// So it is read where it already sits: the package's `vbaProject.bin` is a compound file, each
/// module is a stream inside it, and the stream holds the module's source WITH its attribute
/// header, compressed ([MS-OVBA]). No export, no COM, nothing written.
///
///   the .xlsm is a ZIP        ->  xl/vbaProject.bin
///   vbaProject.bin is a CFB   ->  /VBA/dir          the module table, compressed
///                             ->  /VBA/&lt;module&gt;     the source, compressed from an offset
///                                                   the table gives
///
/// WHAT IT CANNOT SEE, and the reason every answer here is nullable: the file is the last SAVE. A
/// class module created since is not in it, and answers unknown - which is the honest answer and
/// the safe one, because the analyzer stays silent on unknown and only a vouched-for `false`
/// reports.
///
/// Staleness is only DANGEROUS where the saved answer contradicts the live module, and the flag
/// cannot be edited: the editor gives every new class module `False` and nothing in the interface
/// changes it. It can only arrive another way, by importing a `.cls` that carries it - so a name
/// removed and then imported again between saves is the one case where this file would answer
/// about a module that is no longer the one it describes. <see cref="Doubt"/> is how the import
/// says so, and a doubted name answers unknown until a save makes the file true again.
/// </summary>
public sealed class SavedModules
{
    /// <summary>Enough of a module stream to carry its attribute header. The header is a handful
    /// of lines at the very top, so the rest of a module - which may be a megabyte of code - is
    /// never decompressed at all.</summary>
    private const int HeaderBytes = 8 * 1024;

    /// <summary>A corrupt table cannot ask this to walk for ever. No real project is close.</summary>
    private const int MostModules = 10_000;

    /// <summary>The module table decompresses to kilobytes in a real project; this is the ceiling
    /// a corrupt one is held to.</summary>
    private const int MostDirBytes = 4 * 1024 * 1024;

    /// <summary>
    /// Read documents by path, with the write time they were read at.
    ///
    /// GUARDED, and the guard is not theoretical. Every caller in this product is on the host's
    /// user interface thread - the seed walk and the api route that reports the flag both reach
    /// COM, so both are marshalled there - but this is a public static on a type any future
    /// caller can reach from a pool thread, and an unsynchronised Dictionary under a concurrent
    /// write does not merely lose an entry: it can corrupt its buckets and spin. Two test classes
    /// running in parallel were enough to make one read go the long way round (2026-08-23), which
    /// is the cheap version of the same warning.
    /// </summary>
    private static readonly Dictionary<string, (DateTime Stamp, SavedModules? Modules)> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Per document, the module names the saved file can no longer be trusted about.
    /// Emptied when that document is read again, because a read only happens once its write time
    /// has moved, and a save is what makes the file true.</summary>
    private static readonly Dictionary<string, HashSet<string>> Doubted =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Module name to the attribute lines its saved stream carries.</summary>
    private readonly Dictionary<string, string[]> attributes =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Module name to where its whole source sits in the package, for the on-demand read.</summary>
    private readonly Dictionary<string, (string Stream, int Offset)> streams =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The whole attribute set of a module, read once per save and only when asked for.</summary>
    private readonly Dictionary<string, AttributeSet?> full =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Per document, the attribute sets this product itself wrote into modules since the document
    /// was saved. An applied attribute is true of the module the moment the import lands, and the
    /// saved file will not say so until the next save; the assertion stands in for the file until
    /// then, and is emptied when the file is read again, because the save is what makes the file
    /// true. Guarded like the other statics, for the same reason.
    /// </summary>
    private static readonly Dictionary<string, Dictionary<string, AttributeSet>> Asserted =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The whole of a module stream may be decompressed for its member attributes.</summary>
    private const int MostModuleBytes = 16 * 1024 * 1024;

    private readonly string path;

    private SavedModules(string path) => this.path = path;

    /// <summary>
    /// Says that a module has been created or removed since the document was saved, so whatever
    /// the file holds under that name is no longer about the module the editor has.
    ///
    /// Called by the import, which is the only thing that can put a module carrying
    /// `VB_PredeclaredId = True` into a project - the editor's own New Class Module always writes
    /// False, and no interface changes it afterwards.
    /// </summary>
    public static void Doubt(string? documentPath, string moduleName)
    {
        if (string.IsNullOrWhiteSpace(documentPath) || string.IsNullOrWhiteSpace(moduleName))
        {
            return;
        }

        lock (Doubted)
        {
            if (!Doubted.TryGetValue(documentPath, out var names))
            {
                names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Doubted[documentPath] = names;
            }

            names.Add(moduleName);
        }
    }

    private bool IsDoubted(string moduleName)
    {
        lock (Doubted)
        {
            return Doubted.TryGetValue(path, out var names) && names.Contains(moduleName);
        }
    }

    /// <summary>
    /// What the saved document says about its modules, or null when there is nothing to say - a
    /// document never saved, a path that will not open, a package with no VBA project, or bytes
    /// this cannot walk. Every one of those leaves the caller with no answer, which is a state
    /// the caller must already have.
    ///
    /// Cached per path and invalidated by the file's own write time, so Ctrl+S re-reads without
    /// anything having to tell it to, and a pass that changes nothing costs one stat.
    /// </summary>
    public static SavedModules? For(string? documentPath)
    {
        if (string.IsNullOrWhiteSpace(documentPath))
        {
            return null;
        }

        try
        {
            if (!File.Exists(documentPath))
            {
                return null;
            }

            var stamp = File.GetLastWriteTimeUtc(documentPath);
            lock (Cache)
            {
                if (Cache.TryGetValue(documentPath, out var held) && held.Stamp == stamp)
                {
                    return held.Modules;
                }
            }

            // The file has moved on, so whatever was doubted about it is settled: a save is what
            // makes the file describe the project again. What was asserted about it is settled
            // the same way - the file now carries what was applied, or the developer saved
            // without it, and either way the file is the truth again.
            lock (Doubted)
            {
                Doubted.Remove(documentPath);
            }
            lock (Asserted)
            {
                Asserted.Remove(documentPath);
            }

            // Walked OUTSIDE the lock, because it is the expensive half and two callers arriving
            // together should wait on each other for the dictionary, not for the disk. The worst
            // that costs is one duplicated walk, where holding the lock across it could park the
            // user interface thread behind somebody else's file read.
            var read = Read(documentPath);
            lock (Cache)
            {
                Cache[documentPath] = (stamp, read);
            }

            return read;
        }
        catch
        {
            // A document held, moved, encrypted or half-written is a project with no saved
            // header to read, which every caller here already knows how to be.
            return null;
        }
    }

    /// <summary>Whether the saved file carried this module at all. A module added since the last
    /// save is not in it, and that is different from one that is in it and says nothing.</summary>
    public bool Knows(string moduleName) =>
        AssertedOf(moduleName) is not null || (!IsDoubted(moduleName) && attributes.ContainsKey(moduleName));

    /// <summary>
    /// Says that this product wrote these attributes into a module of the document, so they are
    /// true of the module now whatever the saved file says, until the file is saved and read again.
    /// </summary>
    public static void Assert(string? documentPath, string moduleName, AttributeSet attributes)
    {
        ArgumentNullException.ThrowIfNull(attributes);
        if (string.IsNullOrWhiteSpace(documentPath) || string.IsNullOrWhiteSpace(moduleName))
        {
            return;
        }
        lock (Asserted)
        {
            if (!Asserted.TryGetValue(documentPath, out var byModule))
            {
                byModule = new Dictionary<string, AttributeSet>(StringComparer.OrdinalIgnoreCase);
                Asserted[documentPath] = byModule;
            }
            byModule[moduleName] = attributes;
        }
    }

    /// <summary>The attributes this product itself last wrote into the module, if the file has not been read since.</summary>
    public static AttributeSet? AssertedFor(string? documentPath, string moduleName)
    {
        if (string.IsNullOrWhiteSpace(documentPath))
        {
            return null;
        }
        lock (Asserted)
        {
            return Asserted.TryGetValue(documentPath, out var byModule) && byModule.TryGetValue(moduleName, out var set)
                ? set
                : null;
        }
    }

    private AttributeSet? AssertedOf(string moduleName) => AssertedFor(path, moduleName);

    /// <summary>
    /// Every attribute the module carries - module-level and per member - or null when the saved
    /// file cannot answer for it. The whole stream is decompressed for this, once per save, and
    /// only for the module asked about: member attributes sit beside their procedures, anywhere
    /// in a module that may be a megabyte of code, where the header read above stops at eight
    /// kilobytes on purpose.
    /// </summary>
    public AttributeSet? AttributesOf(string moduleName)
    {
        if (AssertedOf(moduleName) is { } asserted)
        {
            return asserted;
        }
        if (IsDoubted(moduleName))
        {
            return null;
        }
        lock (full)
        {
            if (full.TryGetValue(moduleName, out var held))
            {
                return held;
            }
        }

        AttributeSet? read = null;
        try
        {
            if (streams.TryGetValue(moduleName, out var where)
                && OfficePackage.ProjectBytes(path) is { } project
                && CompoundFile.TryRead(project) is { } cfb
                && VbaCompression.Decompress(cfb.Read($"/VBA/{where.Stream}"), where.Offset, MostModuleBytes) is { } text)
            {
                read = ModuleAttributes.Read(Encoding.Latin1.GetString(text));
            }
        }
        catch
        {
            // Unreadable is unknown, the same answer as a module the file never held.
        }

        lock (full)
        {
            full[moduleName] = read;
        }
        return read;
    }

    /// <summary>
    /// Whether the module has a default instance: true, false, or null when the saved file does
    /// not carry the module or its header does not name the attribute.
    ///
    /// NULL IS A REAL ANSWER and must be passed on as one. A caller that turns it into false puts
    /// `Variable not defined` under every use of a legitimately predeclared singleton.
    /// </summary>
    public bool? PredeclaredIdOf(string moduleName) =>
        AssertedOf(moduleName) is { } asserted ? asserted.PredeclaredId
            : IsDoubted(moduleName) ? null : Flag(moduleName, "VB_PredeclaredId");

    /// <summary>One boolean attribute of one module, or null when it is not there to read.</summary>
    private bool? Flag(string moduleName, string attribute)
    {
        if (!attributes.TryGetValue(moduleName, out var lines))
        {
            return null;
        }

        foreach (var line in lines)
        {
            var trimmed = line.AsSpan().Trim();
            if (!trimmed.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var equals = trimmed.IndexOf('=');
            if (equals < 0)
            {
                continue;
            }

            var name = trimmed[10..equals].Trim();
            if (!name.Equals(attribute, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = trimmed[(equals + 1)..].Trim();
            if (value.Equals("True", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (value.Equals("False", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Present but not a boolean. Unknown beats a guess.
            return null;
        }

        return null;
    }

    private static SavedModules? Read(string documentPath)
    {
        var project = OfficePackage.ProjectBytes(documentPath);
        if (project is null || CompoundFile.TryRead(project) is not { } cfb)
        {
            return null;
        }

        // A path that names nothing reads as an empty array, and an empty array is not a
        // container, so both steps answer null without a presence check of their own.
        var dir = VbaCompression.Decompress(cfb.Read("/VBA/dir"), 0, MostDirBytes);
        if (dir is null || dir.Length == 0)
        {
            return null;
        }

        var read = new SavedModules(documentPath)
        {
            ConditionalConstants = ConstantsIn(dir),
        };

        foreach (var (name, stream, offset) in ModulesIn(dir))
        {
            var text = VbaCompression.Decompress(cfb.Read($"/VBA/{stream}"), offset, HeaderBytes);
            if (text is null)
            {
                continue;
            }

            read.attributes[name] = HeaderLinesOf(text);
            read.streams[name] = (stream, offset);
        }

        return read;
    }

    /// <summary>The attribute lines at the top of a module's source, and nothing after them.</summary>
    private static string[] HeaderLinesOf(byte[] text)
    {
        // The header is ASCII whatever the project's code page, and the code below it may not be,
        // so it is read as Latin-1: every byte maps to a character, nothing can fail to decode,
        // and a line that is not an attribute is discarded anyway.
        var lines = Encoding.Latin1.GetString(text).Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');

        var header = new List<string>();
        foreach (var line in lines)
        {
            if (line.AsSpan().TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
            {
                header.Add(line);
                continue;
            }

            // A module's attributes are contiguous at the top, but a form's are preceded by its
            // VERSION/BEGIN block, so a non-attribute line only ends the header once one has
            // started.
            if (header.Count > 0)
            {
                break;
            }
        }

        return [.. header];
    }

    /// <summary>
    /// The project's own conditional compilation arguments, as the VBE's Project Properties box
    /// spells them: `Name = Value : Name = Value`. Null when the file names none.
    ///
    /// WHY IT IS HERE and not asked of the object model: `VBProject` has no property for it. The
    /// analyzer knows the compiler's own constants - VBA7, Win64, Mac - but a project's are its
    /// own, and without them every `#If MY_FLAG Then` is undecidable, so BOTH arms are analyzed
    /// and a finding can be reported from an arm the compiler never sees.
    /// </summary>
    public string? ConditionalConstants { get; private init; }

    /* ---- the module table, [MS-OVBA] 2.3.4.2 ----------------------------------------------- */

    private const ushort ProjectVersion = 0x0009;
    private const ushort ProjectConstants = 0x000C;
    private const ushort ProjectConstantsUnicode = 0x003C;
    private const ushort ProjectModules = 0x000F;
    private const ushort ProjectCookie = 0x0013;
    private const ushort ModuleName = 0x0019;
    private const ushort ModuleNameUnicode = 0x0047;
    private const ushort ModuleStreamName = 0x001A;
    private const ushort ModuleStreamNameUnicode = 0x0032;
    private const ushort ModuleOffset = 0x0031;
    private const ushort ModuleEnd = 0x002B;

    /// <summary>
    /// The PROJECTCONSTANTS record, read by walking FORWARD from the top of the stream.
    ///
    /// The opposite direction from <see cref="ModulesIn"/>, and safe for the reason that one is
    /// not: PROJECTINFORMATION comes first and every record in it follows the id/size/data shape,
    /// with the constants as its last entry, immediately before the references that break the
    /// shape. So the walk reaches it and stops - at the constants, at the first reference record,
    /// or at a bounded count - and never enters the stretch that derails.
    ///
    /// The Unicode twin wins where both are present: it says the same text with no code page in
    /// the way, which matters for a constant named in a script the machine cannot spell.
    /// </summary>
    private static string? ConstantsIn(byte[] dir)
    {
        // PROJECTINFORMATION is a dozen records. Anything past that is the references.
        const int MostRecords = 32;

        string? mbcs = null;
        var at = 0;

        for (var seen = 0; seen < MostRecords && at + 6 <= dir.Length; seen++)
        {
            var id = BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(at));

            // PROJECTVERSION IS NOT SHAPED LIKE THE REST, and it sits between the walk's start and
            // what the walk is for. Its four bytes after the id are a RESERVED constant rather
            // than a size, and its body is six: a major and a minor. Read as a size it says 4, so
            // the walk resumes two bytes early, reads the tail of this record as the next one's
            // id, and every record after it is garbage - which is why the constants were never
            // found in a file that has them ([MS-OVBA] 2.3.4.2.1.9).
            if (id == ProjectVersion)
            {
                at += 2 + 4 + 6;
                continue;
            }

            var size = (int)BinaryPrimitives.ReadUInt32LittleEndian(dir.AsSpan(at + 2));
            if (size < 0 || at + 6 + size > dir.Length)
            {
                break;
            }

            var data = dir.AsSpan(at + 6, size);
            switch (id)
            {
                case ProjectConstants:
                    mbcs = size > 0 ? Encoding.Latin1.GetString(data) : null;
                    break;

                // The twin follows the MBCS record directly, so reading it ends the walk.
                case ProjectConstantsUnicode:
                    return size > 0 ? Encoding.Unicode.GetString(data) : mbcs;

                // The first reference record: past here the shape no longer holds, and the
                // constants would have been seen already if the project had any.
                case ReferenceName or ReferenceRegistered or ReferenceProject or ReferenceControl:
                    return mbcs;
            }

            at += 6 + size;
        }

        return mbcs;
    }

    private const ushort ReferenceName = 0x0016;
    private const ushort ReferenceRegistered = 0x000D;
    private const ushort ReferenceProject = 0x000E;
    private const ushort ReferenceControl = 0x002F;

    /// <summary>
    /// Every module the table names, with the stream it lives in and where its source starts.
    ///
    /// STARTED FROM THE MODULE TABLE, not from the top of the stream, and deliberately. The
    /// records before it are the project's references, and those do NOT all follow the
    /// id/size/data shape the rest of the stream does - `REFERENCECONTROL` carries a size that
    /// covers only part of itself - so a walk from byte zero derails partway through and reads
    /// whatever follows as a record. Measured on LanguageFixture.xlsm: the walk survives to
    /// offset 136 of 1886 and then asks for a 1.8GB record. The table is found instead, and only
    /// what comes after it is walked, where the shape does hold.
    /// </summary>
    private static IEnumerable<(string Name, string Stream, int Offset)> ModulesIn(byte[] dir)
    {
        var at = TableIn(dir);
        if (at < 0)
        {
            yield break;
        }

        at += 6 + 2;
        if (at + 6 <= dir.Length && BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(at)) == ProjectCookie)
        {
            at += 6 + (int)BinaryPrimitives.ReadUInt32LittleEndian(dir.AsSpan(at + 2));
        }

        string? name = null;
        string? stream = null;
        var offset = -1;
        var seen = 0;

        while (at + 6 <= dir.Length && seen < MostModules)
        {
            var id = BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(at));
            var size = (int)BinaryPrimitives.ReadUInt32LittleEndian(dir.AsSpan(at + 2));
            if (size < 0 || at + 6 + size > dir.Length)
            {
                yield break;
            }

            var data = dir.AsSpan(at + 6, size);

            switch (id)
            {
                // The MBCS forms are read as a fallback only. The Unicode records that follow
                // them say the same names without a code page in the way, so a module named in a
                // script the machine's code page cannot spell still matches the live one.
                case ModuleName when name is null:
                    name = Encoding.Latin1.GetString(data);
                    break;
                case ModuleNameUnicode when size > 0:
                    name = Encoding.Unicode.GetString(data);
                    break;
                case ModuleStreamName when stream is null:
                    stream = Encoding.Latin1.GetString(data);
                    break;
                case ModuleStreamNameUnicode when size > 0:
                    stream = Encoding.Unicode.GetString(data);
                    break;
                case ModuleOffset when size >= 4:
                    offset = (int)BinaryPrimitives.ReadUInt32LittleEndian(data);
                    break;
                case ModuleEnd:
                    if (name is { Length: > 0 } && offset >= 0)
                    {
                        seen++;
                        yield return (name, stream is { Length: > 0 } ? stream : name, offset);
                    }

                    name = null;
                    stream = null;
                    offset = -1;
                    break;
                default:
                    break;
            }

            at += 6 + size;
        }
    }

    /// <summary>
    /// Where the module table starts, or -1.
    ///
    /// Found by scanning rather than by walking, for the reason above, and VALIDATED rather than
    /// taken on the first hit: the two bytes that spell this record can occur inside any earlier
    /// record's data. A real table is the record, its cookie, and then a module's name, so all
    /// three are checked before the position is believed.
    /// </summary>
    private static int TableIn(byte[] dir)
    {
        for (var i = 0; i + 6 <= dir.Length; i++)
        {
            if (BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(i)) != ProjectModules
                || BinaryPrimitives.ReadUInt32LittleEndian(dir.AsSpan(i + 2)) != 2)
            {
                continue;
            }

            var cookie = i + 6 + 2;
            if (cookie + 6 > dir.Length
                || BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(cookie)) != ProjectCookie)
            {
                continue;
            }

            var first = cookie + 6 + (int)BinaryPrimitives.ReadUInt32LittleEndian(dir.AsSpan(cookie + 2));
            if (first + 6 > dir.Length
                || BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(first)) != ModuleName)
            {
                continue;
            }

            return i;
        }

        return -1;
    }
}
