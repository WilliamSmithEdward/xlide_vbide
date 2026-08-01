namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// Reads what a code pane's window caption can tell us, and is explicit about what it cannot.
///
/// The editor gives code panes and the Immediate window the same window class, so the caption is
/// the only thing that separates them by inspection. It is also localised, so it is treated as a
/// hint that narrows candidates and never as proof. The object model's pane collection remains the
/// authority for which components actually have panes open; this only helps match a window to one
/// of them.
/// </summary>
public static class CodePaneCaption
{
    /// <summary>
    /// The suffix an English editor appends to a code pane's caption. Other languages use their own,
    /// which is exactly why a failed match here means "unknown", not "not a code pane".
    /// </summary>
    private const string EnglishSuffix = " (Code)";

    /// <summary>
    /// Extracts the component name a caption appears to name, or null when the caption does not
    /// have the shape of a code pane caption in a language we recognise.
    /// </summary>
    public static string? ComponentName(string? caption)
    {
        if (string.IsNullOrWhiteSpace(caption))
        {
            return null;
        }

        var trimmed = caption.Trim();
        if (!trimmed.EndsWith(EnglishSuffix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var name = trimmed[..^EnglishSuffix.Length].Trim();
        return name.Length == 0 ? null : name;
    }

    /// <summary>
    /// True when the caption is one the editor uses for a window that is not a code pane, even
    /// though it carries the same window class. Only the cases we can recognise are listed; an
    /// unrecognised caption is not evidence either way.
    /// </summary>
    public static bool IsKnownNonCodePane(string? caption) =>
        caption is not null && caption.Trim().Equals("Immediate", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Chooses which of the components known to have a pane open corresponds to a caption.
    ///
    /// Matching on the name the caption carries is what makes this work in a localised editor: the
    /// suffix may be unrecognisable while the component name, which the user chose, is not.
    /// </summary>
    public static string? MatchComponent(string? caption, IReadOnlyCollection<string> openComponents)
    {
        ArgumentNullException.ThrowIfNull(openComponents);

        if (string.IsNullOrWhiteSpace(caption) || openComponents.Count == 0)
        {
            return null;
        }

        var fromSuffix = ComponentName(caption);
        if (fromSuffix is not null)
        {
            foreach (var component in openComponents)
            {
                if (string.Equals(component, fromSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    return component;
                }
            }
        }

        // The suffix was not recognised, so fall back to the component name appearing at the start
        // of the caption. Longest first, so a component named "Sheet1" cannot claim a caption that
        // belongs to "Sheet10".
        var trimmed = caption.Trim();
        string? best = null;

        foreach (var component in openComponents)
        {
            if (component.Length == 0 || !trimmed.StartsWith(component, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (best is null || component.Length > best.Length)
            {
                best = component;
            }
        }

        return best;
    }
}
