using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Engine;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// Giving a project a reference to another application's type library.
///
/// `Dim wd As Word.Application` in a workbook with no reference to Word does not compile - the
/// VBE stops the whole project and says "User-defined type not defined" - and the analyzer
/// reports it as `missing-library-reference`. The fix is one tick in Tools > References, so the
/// finding carries it as a quick fix, and this is what that fix runs.
///
/// A TEXT EDIT CANNOT DO IT. Every other quick fix rewrites the module; this one writes a record
/// into the project, so it travels as a command the host performs rather than as edits the
/// surface applies. The editor extension reaches the same conclusion from the other side: it
/// writes the reference into the saved file's `dir` stream, because it has no open host. Here
/// the project IS open, so the reference goes on through the object model the add-in already
/// holds - which also means it needs no "Trust access to the VBA project object model", the
/// setting the add-in is already past by the time the host hands it the VBE.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>
    /// Adds the library to the project by its GUID, or answers in words why it did not.
    ///
    /// NO VERSION TRAVELS. `AddFromGuid(guid, 0, 0)` binds whatever the machine has registered:
    /// it bound Word 8.7 from MSWORD.OLB on a machine with Office 16 (measured 2026-09-21).
    /// Sending a major.minor would write this machine's Office version into the product and
    /// break on the next one.
    /// </summary>
    /// <param name="library">The library's token, for the log: excel, word, powerpoint, access.</param>
    /// <param name="displayName">What to call it in a sentence the developer reads.</param>
    /// <param name="guid">The identity to add, braced as the type library registers it.</param>
    /// <param name="projectDisplay">Which project, or null for the one the surface is showing.</param>
    /// <param name="added">False when the project already had it, which is not a failure.</param>
    private string? AddLibraryReference(
        string library,
        string displayName,
        string guid,
        string? projectDisplay,
        out bool added)
    {
        added = false;

        if (guid.Length == 0)
        {
            return $"'{library}' is not a library this editor can add.";
        }

        using var project = FindProjectByDisplayName(projectDisplay)
            ?? _editor.GetObject("ActiveVBProject");
        if (project is null)
        {
            return "No VBA project is active, so there is nothing to add a reference to.";
        }

        // ALREADY THERE IS NOT A FAILURE, and it has to be asked before the add rather than
        // after: adding a second time raises "Name conflicts with existing module, project, or
        // object library" (error 32813, measured), which is a true statement about COM and says
        // nothing a developer would act on.
        if (ProjectReader.ReferencesOf(project)
            .Any(one => string.Equals(one.LibraryId, guid, StringComparison.OrdinalIgnoreCase)))
        {
            return null;
        }

        try
        {
            using var references = project.GetObject("References");
            if (references is null)
            {
                return $"This project does not expose a reference list, so {displayName} cannot be added.";
            }

            references.Invoke("AddFromGuid", guid, 0, 0);
        }
        catch (Exception ex)
        {
            // The host's own words, not an HRESULT: a library that is not installed, or one the
            // project cannot take, is a real answer the developer can act on.
            Log.Info($"reference: {displayName} could not be added ({ex.GetType().Name}: {ex.Message})");
            return $"{displayName} could not be added: {ex.Message}";
        }

        added = true;
        Log.Info($"reference: added {displayName} ({guid}) to {projectDisplay ?? "the active project"}");

        // THE FINDING HAS TO GO NOW. Nothing in any module's text changed, so no pass would be
        // provoked and none that was would re-seed - both sameness gates compare the reference
        // set for exactly this reason, and this is what starts the pass they answer.
        _analysis?.Reanalyse();
        return null;
    }

    /// <summary>
    /// An engine-named action's arguments with the project the request was about on the end.
    ///
    /// The engine knows a project by the id it was seeded under and the host addresses one by
    /// the name it shows, so the last argument is added on this side, where the translation
    /// lives. Null when the id names no open project, which the handler reads as "the active
    /// one" - the same answer it would have had anyway.
    /// </summary>
    private static string?[] WithProject(string[]? arguments, string? project) =>
        [.. arguments ?? [], project];

    /// <summary>What to say once the add has happened, or has turned out to be unnecessary.</summary>
    private static string AddedReferenceNotice(string displayName, bool added) => added
        ? $"{displayName} is referenced by this project now. Save the workbook to keep it."
        : $"This project already references {displayName}.";
}
