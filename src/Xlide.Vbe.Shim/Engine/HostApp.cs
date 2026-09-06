namespace Xlide.Vbe.Shim.Engine;

/// <summary>
/// WHICH OFFICE APPLICATION THIS ADD-IN IS LOADED INTO.
///
/// The VBE is shared: the same add-in loads in Word, PowerPoint, Access and Outlook as readily as
/// in Excel, and until 2026-08-18 the language service assumed Excel in all of them - Word's
/// ThisDocument was told it was an Excel.Worksheet and offered a worksheet's members (the owner,
/// having run it in Word).
///
/// A NORMALISED WORD RATHER THAN THE PATH, because the engine should not be parsing executable
/// names to decide what it knows. The path is the shim's business; the answer is a token.
///
/// Read from the process rather than from the object model. `Application.Name` would need a
/// crossing and a host that is willing to answer during startup, and the executable already says
/// it: the add-in is loaded BY the host, so the process it is running in IS the host.
/// </summary>
internal static class HostApp
{
    /// <summary>
    /// `excel`, `word`, `powerpoint`, `access`, `outlook`, `visio`, `project`, or `other`.
    ///
    /// `other` is a real answer and the important one: it is what a host this product has never
    /// been run in returns, and the engine's rule for it is to assert nothing about document
    /// modules rather than to guess. Being silent in an unknown host beats being wrong in it.
    /// </summary>
    public static string Name { get; } = FromProcess();

    /// <summary>
    /// Whether this host's VBE carries the MSForms designer at all. Access's does not - Access
    /// VBA has its own Forms and no UserForms, its VBE offers no Insert > UserForm, and
    /// `VBComponents.Add(3)` can only fail there - so every surface that would CREATE a
    /// userform consults this and refuses plainly instead of relaying a COM error (the owner,
    /// 2026-08-19). Reading forms is not gated: a form component cannot exist in such a host.
    /// </summary>
    public static bool CarriesMsForms => Name is not "access";

    /// <summary>
    /// The name this host's <c>Application.Run</c> understands for a procedure, given the file
    /// the project belongs to (null when it has none yet).
    ///
    /// THREE HOSTS, THREE SPELLINGS, each measured against a live one rather than assumed:
    ///
    ///   Excel   <c>'Book1.xlsm'!Module.Proc</c>  Run resolves an unqualified name against the
    ///           ACTIVE workbook, so once two are open the name has to say which file it means,
    ///           or the host answers that the macro may not be available (2026-08-07).
    ///   Word    <c>Module.Proc</c>               Word's Run resolves across the open projects
    ///           and refuses Excel's form outright (2026-08-19, the day the Immediate window
    ///           first ran in Word).
    ///   Access  <c>Proc</c>                      Access has ONE database per application and
    ///           reads a dotted prefix as a LIBRARY DATABASE rather than a module, so both of
    ///           the other forms are refused. Measured 2026-09-06 against an open database:
    ///           `DiscountRate` answered 0.1, while `Pricing.DiscountRate` and the bang form
    ///           both came back "Microsoft Access cannot find the procedure". This is why the
    ///           test runner and the Immediate window could not run anything at all there.
    ///
    /// A host nobody has run this in keeps Excel's form, which is where every host started.
    /// The bare form is why a generated entry point has to carry a name of this product's own:
    /// in Access there is no module to disambiguate it with.
    /// </summary>
    public static string RunTarget(string? file, string module, string procedure) => Name switch
    {
        "access" => procedure,
        "word" => $"{module}.{procedure}",
        _ => file is null ? $"{module}.{procedure}" : $"'{file}'!{module}.{procedure}",
    };

    private static string FromProcess()
    {
        var image = Path.GetFileNameWithoutExtension(Environment.ProcessPath ?? string.Empty);
        return image.ToUpperInvariant() switch
        {
            "EXCEL" => "excel",
            "WINWORD" => "word",
            "POWERPNT" => "powerpoint",
            "MSACCESS" => "access",
            "OUTLOOK" => "outlook",
            "VISIO" => "visio",
            "WINPROJ" => "project",
            _ => "other",
        };
    }
}
