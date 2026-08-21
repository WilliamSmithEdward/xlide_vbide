using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Sync;

/// <summary>Which way the modules are moving.</summary>
public enum SyncDirection
{
    /// <summary>Workbook to folder.</summary>
    Export,

    /// <summary>Folder to workbook.</summary>
    Import,
}

/// <summary>
/// What an export does about files in the folder that no longer name a live module.
/// </summary>
public enum ExportMode
{
    /// <summary>Write the live modules and leave everything else alone.</summary>
    ExportAll,

    /// <summary>Make the folder match the project, which means removing what is no longer there.</summary>
    TrueUp,
}

/// <summary>
/// What an import does about modules in the project that the folder does not contain.
/// </summary>
public enum ImportMode
{
    /// <summary>Bring in what the folder has and leave the rest of the project alone.</summary>
    UpdateOnly,

    /// <summary>
    /// Make the project match the folder, which means deleting modules the folder does not have.
    /// Only standard and class modules: a document module cannot be removed from its workbook, and
    /// a UserForm's designer is not in the file.
    /// </summary>
    TrueUpStandardClass,
}

/// <summary>What one row of the plan will do if it is applied.</summary>
public enum SyncStatus
{
    /// <summary>Both sides already agree.</summary>
    Unchanged,

    /// <summary>The file is not there yet.</summary>
    WillCreate,

    /// <summary>The file is there and will be overwritten.</summary>
    WillWrite,

    /// <summary>The module is there and its code will be replaced.</summary>
    WillUpdate,

    /// <summary>The file or the module goes away.</summary>
    WillRemove,

    /// <summary>The file names a kind that cannot be brought into existence from source.</summary>
    SkippingImport,

    /// <summary>The file is there and could not be read.</summary>
    ReadError,
}

/// <summary>How one line pairs up against the other side.</summary>
public enum DiffKind
{
    Equal,
    Changed,
    Added,
    Removed,

    /// <summary>
    /// A run of identical lines left out, with how many. Not a line of either side: a marker the
    /// dialog draws as a break, so a comparison carries the changes and not the whole file.
    /// </summary>
    Gap,
}

/// <summary>One line of a side-by-side comparison. A number is absent where that side has no line.</summary>
public sealed record SyncDiffLine(int? LeftNumber, int? RightNumber, string Left, string Right, DiffKind Kind);

/// <summary>A module as the project currently holds it.</summary>
/// <param name="Name">The module's name in the project.</param>
/// <param name="Kind">standard, class, document or userform.</param>
/// <param name="Source">The full text INCLUDING the attribute header, so a file round-trips.</param>
public sealed record LiveModule(string Name, string Kind, string Source, string? Design = null);

/// <summary>A module file as the folder currently holds it.</summary>
/// <param name="FileName">The file's own name, with its extension.</param>
/// <param name="Source">The file's whole text, or empty when <paramref name="ReadError"/> is set.</param>
/// <param name="ReadError">Why the file could not be read, when it could not be.</param>
public sealed record RepoFile(string FileName, string Source, string? ReadError = null);

/// <summary>One row of the plan.</summary>
public sealed record SyncItem
{
    /// <summary>Stable within one plan, and what a caller names when it selects rows.</summary>
    public required string Id { get; init; }

    public required string ModuleName { get; init; }

    /// <summary>standard, class, document, userform, or "stale" for a file with no module.</summary>
    public required string ModuleKind { get; init; }

    /// <summary>The file this row is about, relative to the folder.</summary>
    public required string FileName { get; init; }

    public required SyncStatus Status { get; init; }

    /// <summary>Ticked when the row does something. An unchanged row is offered but not ticked.</summary>
    public required bool Checked { get; init; }

    /// <summary>What the row will do, in words.</summary>
    public required string Detail { get; init; }

    /// <summary>Set when the row needs the developer to understand something before applying it.</summary>
    public string? Warning { get; init; }

    public required bool ExistsInProject { get; init; }

    public required bool ExistsInFolder { get; init; }

    /// <summary>A document or UserForm that the folder has and the project does not.</summary>
    public required bool CannotBeCreated { get; init; }

    /// <summary>This row is a form's DESIGN rather than its code: on import it goes through the
    /// markup apply, not through a module write.</summary>
    public bool IsDesign { get; init; }

    public required string LeftTitle { get; init; }

    public required string RightTitle { get; init; }

    /// <summary>The comparison as it is shown: attribute headers taken off.</summary>
    public required IReadOnlyList<SyncDiffLine> Diff { get; init; }

    /// <summary>The same comparison over the raw text, headers and all.</summary>
    public required IReadOnlyList<SyncDiffLine> DiffWithHeaders { get; init; }

    /// <summary>The text this row would write, which is what an apply actually uses.</summary>
    public required string PayloadSource { get; init; }
}

/// <summary>Everything a developer needs to decide whether to press Apply.</summary>
public sealed record SyncPlan
{
    public required SyncDirection Direction { get; init; }

    public required string ProjectId { get; init; }

    public required string ProjectName { get; init; }

    public required string Folder { get; init; }

    public ExportMode ExportMode { get; init; }

    public ImportMode ImportMode { get; init; }

    public required IReadOnlyList<SyncItem> Items { get; init; }

    public required IReadOnlyList<string> Warnings { get; init; }
}

/// <summary>
/// Works out what an import or an export would do, without doing any of it.
///
/// Everything here is pure: it is handed the live modules and the folder's files and answers a
/// plan. The COM and the file system live in the shim's service, which is what makes this half
/// testable without an Excel, and what lets the dialog and the debug api share one answer rather
/// than each deciding for itself what an import means.
///
/// The strategy follows the companion editor's, because the two products write the same files into
/// the same folders and a developer moves between them:
///
///   - A standard module is a .bas; everything else is a .cls.
///   - Files carry their VBA attribute header, so what is written round-trips.
///   - The comparison shown hides that header, because nobody edits it, but the raw comparison is
///     kept beside it and the dialog can show either.
///   - A document module and a UserForm cannot be created from source. They can be updated when
///     they already exist. The row says so rather than failing at apply time.
///   - Removing is opt-in, per direction, and never touches a document module.
/// </summary>
public static class ModuleSync
{
    /// <summary>A UserForm's VB_Base carries two of these; a class or document carries one.</summary>
    private static readonly Regex GuidPattern = new(
        @"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}",
        RegexOptions.Compiled);

    private static readonly Regex VbBasePattern = new(
        @"^\s*Attribute\s+VB_Base\s*=\s*""([^""]*)""",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Multiline);

    /// <summary>
    /// Excel's document class ids: Worksheet, Chart and Workbook. A .cls whose VB_Base names one of
    /// these is a sheet or the workbook, not a class the developer can new up.
    /// </summary>
    private static readonly HashSet<string> DocumentClassIds = new(StringComparer.OrdinalIgnoreCase)
    {
        "{00020819-0000-0000-C000-000000000046}",
        "{00020820-0000-0000-C000-000000000046}",
        "{00020821-0000-0000-C000-000000000046}",
    };

    /// <summary>
    /// What a sheet module is called in the Excel versions a developer is likely to meet. The name
    /// is a fallback for a file with no usable VB_Base, and it is a fallback rather than the rule
    /// because a developer may rename a sheet's module to anything.
    /// </summary>
    private static readonly Regex DocumentNamePattern = new(
        @"^(Sheet|Feuil|Hoja|Tabelle|Foglio|Planilha|Blad|Ark|List)\d*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>Characters a file name cannot carry, replaced rather than rejected.</summary>
    private static readonly Regex UnsafeFileNameCharacters = new(
        @"[<>:""/\\|?*\x00-\x1F]",
        RegexOptions.Compiled);

    /// <summary>
    /// The extension a module of this kind is written as.
    ///
    /// A UserForm is a `.frm`, which is what the VBE's own exporter writes and the only file that
    /// can carry a form's controls: the binary sidecar beside it holds them, and the `.frm` names
    /// it. Writing a form as a `.cls` - which this did until 2026-08-16 - produced a file whose
    /// `OleObjectBlob` line pointed at the temporary export path, so it named a `.frx` that was
    /// not there and changed on every export.
    /// </summary>
    public static string ExtensionFor(string moduleKind) => moduleKind.ToLowerInvariant() switch
    {
        "standard" => "bas",
        "userform" => "frm",
        _ => "cls",
    };

    /*
     * A FORM'S DESIGN IS A FILE OF ITS OWN, and it is xlide's markup rather than anything Office
     * writes (docs/userform-designer.md, the markup layer).
     *
     * A UserForm exported as code alone cannot be put back: Excel's own export writes the
     * controls into a binary .frx and names it in an `OleObjectBlob` line, so a form in source
     * control has always been half a form - the code diffable, the design a blob or nothing. The
     * markup is the design as TEXT, in the same dialect the designer tab edits, so a form now
     * round-trips through a folder the way a module does.
     *
     * `.form` rather than `.frm`: a `.frm` is the VBE's own format and a developer must not be
     * able to hand this file to Excel expecting it to be read. Nothing but xlide reads a `.form`,
     * and that is what the extension says.
     */
    public const string DesignExtension = "form";

    /// <summary>The file a form's design is written to, beside the file its code goes to.</summary>
    public static string DesignFileNameFor(string moduleName)
    {
        var safe = UnsafeFileNameCharacters.Replace(moduleName, "_").TrimEnd('.', ' ');
        return $"{(safe.Length == 0 ? moduleName : safe)}.{DesignExtension}";
    }

    /// <summary>True for a file this product reads as a form's DESIGN rather than as code.</summary>
    public static bool IsDesignFileName(string fileName) =>
        fileName.EndsWith($".{DesignExtension}", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The file a module is written to. A module may be named things a file cannot be, so the unsafe
    /// characters are replaced and a trailing dot or space is dropped, because Windows accepts neither at
    /// the end of a name, and silently trims them, which would make the written file and the name we
    /// looked for disagree.
    /// </summary>
    public static string FileNameFor(string moduleName, string moduleKind)
    {
        var safe = UnsafeFileNameCharacters.Replace(moduleName, "_").TrimEnd('.', ' ');
        if (safe.Length == 0)
        {
            safe = moduleName;
        }

        return $"{safe}.{ExtensionFor(moduleKind)}";
    }

    /// <summary>True for a file this product will read as a module. A `.frx` is not one: it is a
    /// form's binary sidecar, read by the VBE's importer and by nothing here.</summary>
    public static bool IsModuleFileName(string fileName) =>
        fileName.EndsWith(".bas", StringComparison.OrdinalIgnoreCase)
        || fileName.EndsWith(".cls", StringComparison.OrdinalIgnoreCase)
        || fileName.EndsWith(".frm", StringComparison.OrdinalIgnoreCase);

    /// <summary>The extension of a form's binary sidecar - the controls, and everything the
    /// markup does not print. Listed in the folder so the planner can see whether the pair is
    /// whole; never read here, because it is binary and only the VBE's importer reads it.</summary>
    public const string SidecarExtension = "frx";

    public static bool IsSidecarFileName(string fileName) =>
        fileName.EndsWith($".{SidecarExtension}", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The MSForms UserForm coclass, which is what the first `Begin` of an exported form names.
    ///
    /// CONTENT DECIDES WHAT A FILE IS, and it has to: the shared planner writes a form's code as
    /// `.cls` (xlide_vscode#21) while this product writes `.frm`, so the same form arrives under
    /// either name and the text is identical either way. Measured 2026-08-16 on this product's own
    /// export: `EntryForm.cls` holding `VERSION 5.00` and this GUID, beside `EntryForm.frx`.
    /// </summary>
    private const string UserFormCoclass = "{C62A69F0-16DC-11CE-9E98-00AA00574A4F}";

    /// <summary>True when the text is a FORM's exported file, whatever the file is called.</summary>
    public static bool IsFormText(string source) =>
        source.StartsWith("VERSION", StringComparison.OrdinalIgnoreCase)
        && source.Contains(UserFormCoclass, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// The sidecar an exported form NAMES, or null when it names none.
    ///
    /// A form's header carries `OleObjectBlob = "EntryForm.frx":0000`, and that name is the
    /// contract between the two files: import the pair under any other pair of names and the
    /// blob is not found. A form with nothing in it names no blob, and needs no sidecar.
    /// </summary>
    public static string? SidecarNamedBy(string source)
    {
        var match = OleObjectBlobPattern.Match(source);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static readonly Regex OleObjectBlobPattern = new(
        @"OleObjectBlob\s*=\s*""([^""]+)""",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// What kind of module a file holds.
    ///
    /// A .bas is always standard. A .cls is a class unless its header says otherwise: two GUIDs in
    /// VB_Base means a UserForm (the type library and the instance), and one that names an Excel
    /// document class means a sheet or the workbook. VB_PredeclaredId is deliberately NOT used to
    /// decide this: document modules set it, but so does any class written in the singleton style,
    /// and treating it as proof would file those as documents and refuse to create them.
    /// </summary>
    public static string ClassifyFile(string fileName, string source)
    {
        if (fileName.EndsWith(".bas", StringComparison.OrdinalIgnoreCase))
        {
            return "standard";
        }

        // A .frm says what it is without being read: the VBE writes one only for a form. So does
        // the TEXT, whatever the file is called, and that matters because the shared planner
        // writes a form's code as `.cls` where this product writes `.frm` - same bytes, two names.
        if (fileName.EndsWith(".frm", StringComparison.OrdinalIgnoreCase) || IsFormText(source))
        {
            return "userform";
        }

        var moduleName = ModuleNameFromFileName(fileName);
        var vbBase = VbBasePattern.Match(source);
        if (vbBase.Success)
        {
            var guids = GuidPattern.Matches(vbBase.Groups[1].Value);
            if (guids.Count >= 2)
            {
                return "userform";
            }

            foreach (Match guid in guids)
            {
                if (DocumentClassIds.Contains(guid.Value))
                {
                    return "document";
                }
            }
        }

        if (string.Equals(moduleName, "ThisWorkbook", StringComparison.OrdinalIgnoreCase)
            || DocumentNamePattern.IsMatch(moduleName))
        {
            return "document";
        }

        return "class";
    }

    /// <summary>The module a file is about, which is its name without the extension.</summary>
    public static string ModuleNameFromFileName(string fileName)
    {
        var dot = fileName.LastIndexOf('.');
        return dot > 0 ? fileName[..dot] : fileName;
    }

    /// <summary>True for a line the VBA editor keeps out of sight and the developer never edits.</summary>
    public static bool IsAttributeLine(string line) =>
        line.TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// True for a line of the block a file carries ahead of its code: VERSION, the BEGIN/END
    /// block, and the designer properties a UserForm keeps inside it.
    ///
    /// None of it is code and none of it is edited, so the shown comparison drops it along with
    /// the attributes, and the shim reads the same lines back off an export to learn what kind of
    /// module it is looking at. One predicate for both: they were written separately, they
    /// disagreed about the UserForm properties, and the half that draws the diff was the half
    /// without them - so a form's comparison opened with a screenful of designer noise.
    ///
    /// Only ever consulted BEFORE the first line of code, so a Caption or a Client inside a
    /// procedure is code like anything else.
    /// </summary>
    public static bool IsHeaderPreamble(string line)
    {
        var trimmed = line.TrimStart();
        return trimmed.StartsWith("VERSION", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("BEGIN", StringComparison.OrdinalIgnoreCase)
            || trimmed.Equals("END", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("MultiUse", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("Caption", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("Client", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("OleObjectBlob", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("StartUpPosition", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Line endings differ between the code store and a file; neither side means anything by it.</summary>
    public static string NormaliseEol(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');

    /// <summary>
    /// The text as the comparison shows it: no header, and no blank lines left at the top where the
    /// header used to be.
    /// </summary>
    public static string CodeWithoutHeader(string source)
    {
        var lines = NormaliseEol(source).Split('\n');
        var kept = new List<string>(lines.Length);
        var stillInHeader = true;

        foreach (var line in lines)
        {
            if (stillInHeader && (IsAttributeLine(line) || IsHeaderPreamble(line)))
            {
                continue;
            }

            // An Attribute line can also appear INSIDE a procedure (VB_Description), and those are
            // dropped wherever they are, the way the editor hides them.
            if (IsAttributeLine(line))
            {
                continue;
            }

            if (stillInHeader && line.Trim().Length == 0)
            {
                continue;
            }

            stillInHeader = false;
            kept.Add(line);
        }

        return string.Join("\n", kept);
    }

    /// <summary>True when the two texts are the same code, whatever their line endings.</summary>
    public static bool SameText(string left, string right) =>
        string.Equals(NormaliseEol(left), NormaliseEol(right), StringComparison.Ordinal);

    /// <summary>
    /// A side-by-side comparison, longest common subsequence over whole lines.
    ///
    /// The common head and tail are taken off first. Two texts that agree, which is most rows in a
    /// project that is nearly in sync, then need no table at all, and the table that remains is
    /// over the part that actually differs rather than over both files.
    /// </summary>
    public static IReadOnlyList<SyncDiffLine> Diff(string leftText, string rightText)
    {
        var left = SplitLines(leftText);
        var right = SplitLines(rightText);
        var output = new List<SyncDiffLine>();

        var head = 0;
        while (head < left.Length && head < right.Length && left[head] == right[head])
        {
            output.Add(new SyncDiffLine(head + 1, head + 1, left[head], right[head], DiffKind.Equal));
            head++;
        }

        var leftEnd = left.Length;
        var rightEnd = right.Length;
        var tail = 0;
        while (leftEnd > head && rightEnd > head && left[leftEnd - 1] == right[rightEnd - 1])
        {
            leftEnd--;
            rightEnd--;
            tail++;
        }

        var midLeft = left[head..leftEnd];
        var midRight = right[head..rightEnd];

        // BOUNDED, because the table below is the whole cost and it is quadratic in BOTH time and
        // memory. Measured on generated modules whose middles share nothing: 803 lines 5ms and
        // 3 MB, 1,603 lines 21ms and 10 MB, 3,203 lines 83ms and 40 MB, 6,403 lines 353ms and
        // 158 MB - four times the cost for twice the lines, exactly as `new int[n+1, m+1]` says
        // it must be. At VBA's own per-module ceiling of 65,534 lines that extrapolates to about
        // 36 seconds and 16 GB, which is not a slow dialog: it is an OutOfMemoryException while
        // a developer is looking at an import they were about to accept (2026-08-21).
        //
        // The head and tail are already off, so this cap is only reached by two texts that
        // genuinely disagree over thousands of lines - a module regenerated, or an export edited
        // wholesale elsewhere. Lining those up line by line tells a reader nothing they can use
        // anyway, so the honest answer is the one that costs nothing: this went, that came.
        if (midLeft.Length > LargestAlignableMiddle || midRight.Length > LargestAlignableMiddle)
        {
            for (var gone = 0; gone < midLeft.Length; gone++)
            {
                output.Add(new SyncDiffLine(head + gone + 1, null, midLeft[gone], string.Empty, DiffKind.Removed));
            }

            for (var came = 0; came < midRight.Length; came++)
            {
                output.Add(new SyncDiffLine(null, head + came + 1, string.Empty, midRight[came], DiffKind.Added));
            }

            for (var k = 0; k < tail; k++)
            {
                output.Add(new SyncDiffLine(
                    leftEnd + k + 1, rightEnd + k + 1, left[leftEnd + k], right[rightEnd + k], DiffKind.Equal));
            }

            return output;
        }

        var table = LongestCommonSubsequence(midLeft, midRight);

        int i = 0, j = 0;
        while (i < midLeft.Length && j < midRight.Length)
        {
            if (midLeft[i] == midRight[j])
            {
                output.Add(new SyncDiffLine(head + i + 1, head + j + 1, midLeft[i], midRight[j], DiffKind.Equal));
                i++;
                j++;
            }
            else if (table[i + 1, j] >= table[i, j + 1])
            {
                if (table[i + 1, j] == table[i, j + 1])
                {
                    // Neither side is longer, so this is one line rewritten rather than a line
                    // taken away and a different one put back.
                    output.Add(new SyncDiffLine(head + i + 1, head + j + 1, midLeft[i], midRight[j], DiffKind.Changed));
                    i++;
                    j++;
                }
                else
                {
                    output.Add(new SyncDiffLine(head + i + 1, null, midLeft[i], string.Empty, DiffKind.Removed));
                    i++;
                }
            }
            else
            {
                output.Add(new SyncDiffLine(null, head + j + 1, string.Empty, midRight[j], DiffKind.Added));
                j++;
            }
        }

        while (i < midLeft.Length)
        {
            output.Add(new SyncDiffLine(head + i + 1, null, midLeft[i], string.Empty, DiffKind.Removed));
            i++;
        }

        while (j < midRight.Length)
        {
            output.Add(new SyncDiffLine(null, head + j + 1, string.Empty, midRight[j], DiffKind.Added));
            j++;
        }

        for (var k = 0; k < tail; k++)
        {
            output.Add(new SyncDiffLine(leftEnd + k + 1, rightEnd + k + 1, left[leftEnd + k], right[rightEnd + k], DiffKind.Equal));
        }

        return output;
    }

    /// <summary>
    /// The comparison with its long stretches of agreement left out, keeping <paramref name="context"/>
    /// identical lines either side of every change.
    ///
    /// WHY THIS EXISTS, measured. A plan carried every line of every module twice, as objects: for
    /// a project of 81,795 lines that was a 15MB answer of which 100% was comparison lines, for a
    /// dialog that shows one row at a time and a module that is usually identical on both sides.
    /// Condensed, a row that has not changed carries a single gap, and a row that has carries its
    /// changes (2026-08-09).
    ///
    /// Nothing about the DECISION passes through here. The statuses were already settled; this is
    /// only what the developer is shown.
    /// </summary>
    /// <param name="most">
    /// The most lines to answer with. Condensing alone is not enough: a FIRST export compares every
    /// module against a file that is not there, so every line is a change and there is no agreement
    /// to leave out: 163,627 lines and 15MB for a project of 81,795, measured. Nobody reads a
    /// comparison that long, so it stops and says how much it did not show.
    /// </param>
    public static IReadOnlyList<SyncDiffLine> Condense(
        IReadOnlyList<SyncDiffLine> diff,
        int context = 3,
        int most = 400)
    {
        ArgumentNullException.ThrowIfNull(diff);

        // Which lines are worth keeping: every change, and `context` lines around each.
        var keep = new bool[diff.Count];
        for (var i = 0; i < diff.Count; i++)
        {
            if (diff[i].Kind == DiffKind.Equal)
            {
                continue;
            }

            for (var near = Math.Max(0, i - context); near <= Math.Min(diff.Count - 1, i + context); near++)
            {
                keep[near] = true;
            }
        }

        var condensed = new List<SyncDiffLine>();
        var skipped = 0;

        for (var i = 0; i < diff.Count; i++)
        {
            if (keep[i])
            {
                if (skipped > 0)
                {
                    condensed.Add(GapOf(skipped));
                    skipped = 0;
                }

                condensed.Add(diff[i]);
                continue;
            }

            skipped++;
        }

        if (skipped > 0)
        {
            condensed.Add(GapOf(skipped));
        }

        if (condensed.Count <= most)
        {
            return condensed;
        }

        var capped = condensed.Take(most).ToList();
        capped.Add(GapOf(condensed.Count - most, "not shown"));
        return capped;
    }

    private static SyncDiffLine GapOf(int lines, string what = "identical") => new(
        null,
        null,
        $"{lines:N0} {what} line{(lines == 1 ? string.Empty : "s")}",
        string.Empty,
        DiffKind.Gap);

    /// <summary>
    /// The comparison of a text with itself, without comparing it.
    ///
    /// A row whose two sides are already known to agree still ran a longest-common-subsequence
    /// over both of them, which for a module of 64,802 lines is the most expensive thing in the
    /// plan and produces a screenful of agreement that condenses back down to this one line
    /// anyway. The STATUS never came from the comparison - it comes from SameText, which is a
    /// string equality - so nothing about what the row does changes (2026-08-09).
    ///
    /// Measured on a project of 81,795 lines already in sync: 414ms of planning became 5ms.
    ///
    /// Public because the SHARED planner's rows want the same treatment for the same reason, one
    /// pipe further out. It stops drawing an unchanged row line by line and this writes the line
    /// that would have survived, so both planners' unchanged rows are the same object.
    /// </summary>
    public static IReadOnlyList<SyncDiffLine> Identical(string text)
    {
        var lines = text.Length == 0 ? 0 : NormaliseEol(text).Split('\n').Length;
        return lines == 0 ? [] : [GapOf(lines)];
    }

    /// <summary>Works out what writing the project into the folder would do.</summary>
    public static SyncPlan PlanExport(
        string projectId,
        string projectName,
        string folder,
        IReadOnlyList<LiveModule> modules,
        IReadOnlyList<RepoFile> folderFiles,
        ExportMode mode)
    {
        var byFileName = folderFiles.ToDictionary(f => f.FileName, StringComparer.OrdinalIgnoreCase);
        var live = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var module in modules)
        {
            live.Add(FileNameFor(module.Name, module.Kind));
        }

        // ONE MODULE AT A TIME IS ONE CORE AT A TIME.
        //
        // Every row is worked out from its own module and its own file and touches nothing else,
        // and the expensive part of each is a longest-common-subsequence over the whole text. On
        // a project of 81,795 lines that was 221ms of one core while the rest idled (2026-08-09).
        //
        // ORDERED on purpose. A plan that lists its rows differently on consecutive runs is a plan
        // nobody can trust, and the sort afterwards is by status first, so the order within a
        // status is the order the modules arrived in.
        var items = modules.AsParallel().AsOrdered().Select(module =>
        {
            var fileName = FileNameFor(module.Name, module.Kind);
            var existing = byFileName.GetValueOrDefault(fileName);
            var onDisk = existing?.Source ?? string.Empty;
            var unchanged = existing is not null && SameText(module.Source, onDisk);
            var status = unchanged
                ? SyncStatus.Unchanged
                : existing is not null ? SyncStatus.WillWrite : SyncStatus.WillCreate;

            return new SyncItem
            {
                Id = $"export:{module.Name}",
                ModuleName = module.Name,
                ModuleKind = module.Kind,
                FileName = fileName,
                Status = status,
                Checked = status is SyncStatus.WillWrite or SyncStatus.WillCreate,
                Detail = DetailFor(status),
                ExistsInProject = true,
                ExistsInFolder = existing is not null,
                CannotBeCreated = false,
                LeftTitle = $"{module.Name} in the project",
                RightTitle = status switch
                {
                    SyncStatus.WillCreate => $"{fileName} (new file)",
                    SyncStatus.WillWrite => $"{fileName} (overwritten)",
                    _ => fileName,
                },
                Diff = unchanged
                    ? Identical(CodeWithoutHeader(module.Source))
                    : Diff(CodeWithoutHeader(module.Source), CodeWithoutHeader(onDisk)),
                DiffWithHeaders = unchanged
                    ? Identical(module.Source)
                    : Diff(module.Source, onDisk),
                PayloadSource = module.Source,
            };
        }).ToList();

        if (mode == ExportMode.TrueUp)
        {
            foreach (var file in folderFiles)
            {
                if (live.Contains(file.FileName) || !IsModuleFileName(file.FileName))
                {
                    continue;
                }

                items.Add(new SyncItem
                {
                    Id = $"export-stale:{file.FileName}",
                    ModuleName = ModuleNameFromFileName(file.FileName),
                    ModuleKind = "stale",
                    FileName = file.FileName,
                    Status = SyncStatus.WillRemove,
                    Checked = true,
                    Detail = "Delete the file",
                    Warning = "No module of this name is in the project. Making the folder match will delete this file.",
                    ExistsInProject = false,
                    ExistsInFolder = true,
                    CannotBeCreated = false,
                    LeftTitle = "Not in the project",
                    RightTitle = $"{file.FileName} (deleted)",
                    Diff = Diff(string.Empty, CodeWithoutHeader(file.Source)),
                    DiffWithHeaders = Diff(string.Empty, file.Source),
                    PayloadSource = string.Empty,
                });
            }
        }

        items.AddRange(DesignRowsForExport(modules, byFileName, live));

        return new SyncPlan
        {
            Direction = SyncDirection.Export,
            ProjectId = projectId,
            ProjectName = projectName,
            Folder = folder,
            ExportMode = mode,
            Items = Sort(items),
            Warnings = [],
        };
    }


    /*
     * A FORM'S DESIGN ROWS, for whichever planner drew the rest.
     *
     * Sync has two planners - the companion editor's, shared through the engine, and this one -
     * and the shared one is the default because both products write into the same folders. It has
     * no idea a form's design can be a file, and it should not: the markup is xlide_vbide's own.
     * So the rows are built here and added to whichever plan came back, which keeps one
     * implementation for both paths rather than a copy that agrees today.
     */
    public static IReadOnlyList<SyncItem> DesignRowsForExport(
        IReadOnlyList<LiveModule> modules,
        IReadOnlyDictionary<string, RepoFile> byFileName,
        HashSet<string>? live = null)
    {
        var rows = new List<SyncItem>();
        foreach (var module in modules)
        {
            if (module.Design is not { } design)
            {
                continue;
            }

            var designFile = DesignFileNameFor(module.Name);
            live?.Add(designFile);
            var onDisk = byFileName.GetValueOrDefault(designFile);
            var same = onDisk is not null && SameText(design, onDisk.Source);
            var status = same
                ? SyncStatus.Unchanged
                : onDisk is not null ? SyncStatus.WillWrite : SyncStatus.WillCreate;

            rows.Add(new SyncItem
            {
                Id = $"export-design:{module.Name}",
                ModuleName = module.Name,
                ModuleKind = module.Kind,
                FileName = designFile,
                Status = status,
                Checked = status is SyncStatus.WillWrite or SyncStatus.WillCreate,
                Detail = DetailFor(status),
                ExistsInProject = true,
                ExistsInFolder = onDisk is not null,
                CannotBeCreated = false,
                IsDesign = true,
                LeftTitle = $"{module.Name}'s design in the project",
                RightTitle = status switch
                {
                    SyncStatus.WillCreate => $"{designFile} (new file)",
                    SyncStatus.WillWrite => $"{designFile} (overwritten)",
                    _ => designFile,
                },
                Diff = same ? Identical(design) : Diff(design, onDisk?.Source ?? string.Empty),
                DiffWithHeaders = same ? Identical(design) : Diff(design, onDisk?.Source ?? string.Empty),
                PayloadSource = design,
            });
        }

        return rows;
    }

    /// <summary>The same rows the other way: a `.form` in the folder applies to the form it names,
    /// through the markup's own name-keyed diff rather than through a module write.</summary>
    public static IReadOnlyList<SyncItem> DesignRowsForImport(
        IReadOnlyList<LiveModule> modules,
        IReadOnlyList<RepoFile> folderFiles,
        List<string>? warnings = null)
    {
        var byModuleName = modules.ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);
        var rows = new List<SyncItem>();

        foreach (var file in folderFiles.Where(one => IsDesignFileName(one.FileName)))
        {
            var moduleName = ModuleNameFromFileName(file.FileName);
            var existing = byModuleName.GetValueOrDefault(moduleName);

            // THE FORM MAY BE ARRIVING IN THE SAME IMPORT. A design row used to be skipped
            // outright when the project held no such form; now that a form can be created from
            // its own pair, the row that follows the create is a real row - the creates run
            // first, because the module rows come first in the plan.
            var arriving = existing is null && CreatableFormFor(moduleName, folderFiles);
            var missing = existing is null && !arriving;
            var status = file.ReadError is not null
                ? SyncStatus.ReadError
                : missing
                    ? SyncStatus.SkippingImport
                    : existing is not null && SameText(file.Source, existing.Design ?? string.Empty)
                        ? SyncStatus.Unchanged
                        : SyncStatus.WillUpdate;

            if (missing)
            {
                warnings?.Add($"{moduleName}: skipped, because the form itself is not in the project.");
            }

            rows.Add(new SyncItem
            {
                Id = $"import-design:{file.FileName}",
                ModuleName = moduleName,
                ModuleKind = existing?.Kind ?? "userform",
                FileName = file.FileName,
                Status = status,
                // UNTICKED ON A CREATE, and that is the foot-gun guard rather than an oversight.
                // The binary sidecar the create imports is AUTHORITATIVE: it carries the whole
                // form, including everything the markup does not print. Applying the text on top
                // can therefore only take things away - a `.form` out of date with its `.frx`
                // prunes controls the binary brought in, because the markup's control list is
                // TOTAL. So the row is offered and left for the developer to tick, which makes
                // pruning a choice rather than a surprise.
                Checked = status == SyncStatus.WillUpdate && !arriving,
                Detail = status == SyncStatus.WillUpdate ? "Apply the design" : DetailFor(status),
                Warning = file.ReadError
                    ?? (missing
                        ? "This is a form's design. Add the form first and this will build its controls."
                        : arriving
                            ? "The form is being created from its own binary in this import, which already "
                                + "carries every control. Tick this only to make the form match this file "
                                + "exactly - a control the file does not name would be removed."
                            : status == SyncStatus.WillUpdate
                                ? "The controls this names are added, moved and set on the form itself. "
                                    + "A control the file does not name is removed."
                                : null),
                ExistsInProject = existing is not null,
                ExistsInFolder = true,
                CannotBeCreated = missing,
                IsDesign = true,
                LeftTitle = file.FileName,
                RightTitle = status switch
                {
                    SyncStatus.WillUpdate => $"{moduleName}'s design (applied)",
                    SyncStatus.SkippingImport => $"{moduleName} (no such form)",
                    _ => moduleName,
                },
                Diff = status == SyncStatus.Unchanged
                    ? Identical(file.Source)
                    : Diff(file.Source, existing?.Design ?? string.Empty),
                DiffWithHeaders = status == SyncStatus.Unchanged
                    ? Identical(file.Source)
                    : Diff(file.Source, existing?.Design ?? string.Empty),
                PayloadSource = file.Source,
            });
        }

        return rows;
    }

    /// <summary>
    /// True when the folder holds a whole form PAIR for this name - the exported text and the
    /// sidecar it points at - which is what makes a design row for a form the project does not
    /// hold a real row rather than a skip.
    /// </summary>
    private static bool CreatableFormFor(string moduleName, IReadOnlyList<RepoFile> folderFiles)
    {
        foreach (var file in folderFiles)
        {
            if (!IsModuleFileName(file.FileName)
                || !string.Equals(ModuleNameFromFileName(file.FileName), moduleName, StringComparison.OrdinalIgnoreCase)
                || !IsFormText(file.Source))
            {
                continue;
            }

            return SidecarMissingFor(file, folderFiles) is null;
        }

        return false;
    }

    /// <summary>
    /// The sidecar a form's file names and the folder does not hold, or null when the pair is
    /// whole - which includes a form that names no sidecar at all, because a form with nothing
    /// in it has no blob to point at.
    /// </summary>
    private static string? SidecarMissingFor(RepoFile file, IReadOnlyList<RepoFile> folderFiles)
    {
        if (SidecarNamedBy(file.Source) is not { Length: > 0 } named)
        {
            return null;
        }

        return folderFiles.Any(one => string.Equals(one.FileName, named, StringComparison.OrdinalIgnoreCase))
            ? null
            : named;
    }

    /// <summary>Works out what reading the folder into the project would do.</summary>
    public static SyncPlan PlanImport(
        string projectId,
        string projectName,
        string folder,
        IReadOnlyList<LiveModule> modules,
        IReadOnlyList<RepoFile> folderFiles,
        ImportMode mode)
    {
        var warnings = new List<string>();
        var byModuleName = modules.ToDictionary(m => m.Name, StringComparer.OrdinalIgnoreCase);
        var fromFolder = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var incoming = folderFiles.Where(f => IsModuleFileName(f.FileName)).ToList();
        foreach (var file in incoming)
        {
            fromFolder.Add(ModuleNameFromFileName(file.FileName));
        }

        // As above: one row per file, each independent, each dominated by its own comparison.
        var items = incoming.AsParallel().AsOrdered().Select(file =>
        {
            var moduleName = ModuleNameFromFileName(file.FileName);
            var existing = byModuleName.GetValueOrDefault(moduleName);
            var kind = existing?.Kind ?? ClassifyFile(file.FileName, file.Source);
            var missingSidecar = existing is null && kind == "userform"
                ? SidecarMissingFor(file, folderFiles)
                : null;
            // A FORM CAN BE CREATED NOW, which is what changed on 2026-08-16. The pair the VBE's
            // own exporter writes - the header text and the binary beside it - is what its
            // importer reads, so a form round-trips through a folder the way a module does. What
            // still cannot be created is a form whose SIDECAR is not there (the header names it
            // and the import would find nothing) and a document module, which belongs to the
            // workbook rather than to the project.
            var cannotBeCreated = existing is null
                && (kind == "document"
                    || (kind == "userform" && (!IsFormText(file.Source) || missingSidecar is not null)));
            var projectSource = existing?.Source ?? string.Empty;

            var status = file.ReadError is not null
                ? SyncStatus.ReadError
                : cannotBeCreated
                    ? SyncStatus.SkippingImport
                    : existing is not null && SameText(file.Source, projectSource)
                        ? SyncStatus.Unchanged
                        : existing is not null ? SyncStatus.WillUpdate : SyncStatus.WillCreate;

            string? warning = null;
            if (file.ReadError is not null)
            {
                warning = file.ReadError;
            }
            else if (cannotBeCreated)
            {
                warning = kind == "document"
                    ? "A worksheet or workbook module belongs to the workbook and cannot be created from a file. "
                        + "Add the sheet first and this will update its code."
                    : missingSidecar is not null
                        ? $"This form names {missingSidecar}, which is not in the folder. That file holds its "
                            + "controls, so the form cannot be created without it. Add the form first and this "
                            + "will update its code."
                        : "A UserForm's designer is not in this file, so the form cannot be created from it. "
                            + "Add the form first and this will update its code.";
                lock (warnings)
                {
                    warnings.Add(missingSidecar is not null
                        ? $"{moduleName}: skipped, because {missingSidecar} is not in the folder."
                        : $"{moduleName}: skipped, because a {kind} cannot be created from source.");
                }
            }
            else if (existing is null && kind == "userform")
            {
                warning = "The form is created from the pair the VBE's own exporter writes - this file and its "
                    + $"binary sidecar - so it arrives whole: controls, fonts, pictures and all.";
            }
            else if (existing is not null && kind is "document" or "userform")
            {
                warning = "The module already exists, so its code will be replaced. The sheet or form itself is untouched.";
            }

            return new SyncItem
            {
                Id = $"import:{file.FileName}",
                ModuleName = moduleName,
                ModuleKind = kind,
                FileName = file.FileName,
                Status = status,
                Checked = status is SyncStatus.WillUpdate or SyncStatus.WillCreate,
                Detail = DetailFor(status),
                Warning = warning,
                ExistsInProject = existing is not null,
                ExistsInFolder = true,
                CannotBeCreated = cannotBeCreated,
                LeftTitle = file.FileName,
                RightTitle = status switch
                {
                    SyncStatus.WillCreate => $"{moduleName} (new module)",
                    SyncStatus.WillUpdate => $"{moduleName} (replaced)",
                    SyncStatus.SkippingImport => $"{moduleName} (cannot be created)",
                    _ => moduleName,
                },
                Diff = status == SyncStatus.Unchanged
                    ? Identical(CodeWithoutHeader(file.Source))
                    : Diff(CodeWithoutHeader(file.Source), CodeWithoutHeader(projectSource)),
                DiffWithHeaders = status == SyncStatus.Unchanged
                    ? Identical(file.Source)
                    : Diff(file.Source, projectSource),
                PayloadSource = file.Source,
            };
        }).ToList();

        items.AddRange(DesignRowsForImport(modules, folderFiles, warnings));

        if (mode == ImportMode.TrueUpStandardClass)
        {
            foreach (var module in modules)
            {
                if (fromFolder.Contains(module.Name) || module.Kind is not ("standard" or "class"))
                {
                    continue;
                }

                items.Add(new SyncItem
                {
                    Id = $"import-stale:{module.Name}",
                    ModuleName = module.Name,
                    ModuleKind = module.Kind,
                    FileName = FileNameFor(module.Name, module.Kind),
                    Status = SyncStatus.WillRemove,
                    Checked = true,
                    Detail = "Delete the module",
                    Warning = "The folder has no file for this module. Making the project match will delete it from the workbook.",
                    ExistsInProject = true,
                    ExistsInFolder = false,
                    CannotBeCreated = false,
                    LeftTitle = "Not in the folder",
                    RightTitle = $"{module.Name} (deleted)",
                    Diff = Diff(string.Empty, CodeWithoutHeader(module.Source)),
                    DiffWithHeaders = Diff(string.Empty, module.Source),
                    PayloadSource = string.Empty,
                });
            }
        }

        return new SyncPlan
        {
            Direction = SyncDirection.Import,
            ProjectId = projectId,
            ProjectName = projectName,
            Folder = folder,
            ImportMode = mode,
            Items = Sort(items),
            Warnings = warnings,
        };
    }

    /// <summary>What the row says it will do, in the words the dialog shows.</summary>
    public static string DetailFor(SyncStatus status) => status switch
    {
        SyncStatus.Unchanged => "Already the same",
        SyncStatus.WillCreate => "Create",
        SyncStatus.WillWrite => "Overwrite the file",
        SyncStatus.WillUpdate => "Replace the module's code",
        SyncStatus.WillRemove => "Delete",
        SyncStatus.SkippingImport => "Skipped",
        SyncStatus.ReadError => "Could not be read",
        _ => string.Empty,
    };

    /// <summary>The word a caller spells a status with, over the wire and in the api.</summary>
    public static string NameOf(SyncStatus status) => status switch
    {
        SyncStatus.Unchanged => "unchanged",
        SyncStatus.WillCreate => "will-create",
        SyncStatus.WillWrite => "will-write",
        SyncStatus.WillUpdate => "will-update",
        SyncStatus.WillRemove => "will-remove",
        SyncStatus.SkippingImport => "skipping-import",
        SyncStatus.ReadError => "read-error",
        _ => "unchanged",
    };

    /// <summary>The word for one line's relationship to the other side.</summary>
    public static string NameOf(DiffKind kind) => kind switch
    {
        DiffKind.Equal => "equal",
        DiffKind.Changed => "changed",
        DiffKind.Added => "added",
        DiffKind.Removed => "removed",
        DiffKind.Gap => "gap",
        _ => "equal",
    };

    /// <summary>Reads an export mode the way a caller spelled it, defaulting to the safe one.</summary>
    public static ExportMode ExportModeFrom(string? text) =>
        string.Equals(text, "trueUp", StringComparison.OrdinalIgnoreCase)
        || string.Equals(text, "true-up", StringComparison.OrdinalIgnoreCase)
            ? ExportMode.TrueUp
            : ExportMode.ExportAll;

    /// <summary>Reads an import mode the way a caller spelled it, defaulting to the safe one.</summary>
    public static ImportMode ImportModeFrom(string? text) =>
        string.Equals(text, "trueUpStandardClass", StringComparison.OrdinalIgnoreCase)
        || string.Equals(text, "true-up", StringComparison.OrdinalIgnoreCase)
            ? ImportMode.TrueUpStandardClass
            : ImportMode.UpdateOnly;

    /// <summary>
    /// Rows that do something come first, then rows that need attention, then the ones already in
    /// agreement. Within a group, by name, so a plan reads the same twice running.
    /// </summary>
    private static List<SyncItem> Sort(List<SyncItem> items) =>
        [.. items
            .OrderBy(SortRank)
            .ThenBy(item => item.ModuleName, StringComparer.OrdinalIgnoreCase)];

    private static int SortRank(SyncItem item) => item.Status switch
    {
        SyncStatus.WillCreate or SyncStatus.WillUpdate or SyncStatus.WillWrite => 0,
        SyncStatus.WillRemove => 1,
        SyncStatus.SkippingImport => 2,
        SyncStatus.ReadError => 3,
        _ => 4,
    };

    private static string[] SplitLines(string text) =>
        text.Length == 0 ? [] : NormaliseEol(text).Split('\n');

    /// <summary>
    /// How many differing lines either side may have before the comparison stops trying to line
    /// them up. 2,000 is a table of 4 million ints - 16 MB and about 150ms at the worst - which is
    /// the most a dialog should spend deciding how to draw something.
    /// </summary>
    private const int LargestAlignableMiddle = 2000;

    private static int[,] LongestCommonSubsequence(string[] left, string[] right)
    {
        var table = new int[left.Length + 1, right.Length + 1];
        for (var i = left.Length - 1; i >= 0; i--)
        {
            for (var j = right.Length - 1; j >= 0; j--)
            {
                table[i, j] = left[i] == right[j]
                    ? table[i + 1, j + 1] + 1
                    : Math.Max(table[i + 1, j], table[i, j + 1]);
            }
        }

        return table;
    }
}
