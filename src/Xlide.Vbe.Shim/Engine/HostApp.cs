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
