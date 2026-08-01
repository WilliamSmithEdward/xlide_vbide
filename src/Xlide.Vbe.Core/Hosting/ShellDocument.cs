namespace Xlide.Vbe.Core.Hosting;

/// <summary>
/// Builds the markup the tool window renders.
///
/// The document is handed to the browser as a string rather than fetched from a file URL. A file
/// URL would put the surface in a security context that blocks module scripts, storage, and fetch,
/// which is the opposite of where this is going. The eventual editor surface will be served from a
/// virtual host name mapped to the UI directory, because that gives a normal HTTPS origin with none
/// of the file-scheme restrictions. That mapping lives on a later revision of the browser
/// interface, so until the surface actually needs it the placeholder is composed here and pushed
/// straight into the browser.
/// </summary>
public static class ShellDocument
{
    /// <summary>Token in the shell document replaced with the browser version at load time.</summary>
    public const string VersionToken = "{{WEBVIEW2_VERSION}}";

    /// <summary>Substitutes the runtime facts the static document cannot know.</summary>
    public static string Compose(string template, string? browserVersion)
    {
        ArgumentNullException.ThrowIfNull(template);

        return template.Replace(
            VersionToken,
            string.IsNullOrWhiteSpace(browserVersion) ? "unknown" : browserVersion,
            StringComparison.Ordinal);
    }

    /// <summary>
    /// Markup used when the shell document is missing from the install. It states the fact rather
    /// than showing an empty pane, because an empty pane is indistinguishable from a browser that
    /// never started.
    /// </summary>
    public static string Missing(string expectedPath)
    {
        ArgumentNullException.ThrowIfNull(expectedPath);

        var escaped = expectedPath
            .Replace("&", "&amp;", StringComparison.Ordinal)
            .Replace("<", "&lt;", StringComparison.Ordinal)
            .Replace(">", "&gt;", StringComparison.Ordinal);

        // Double the interpolation markers so the style rules keep their own braces.
        return $$"""
            <!doctype html>
            <meta charset="utf-8">
            <title>xlide</title>
            <style>
              :root { color-scheme: light dark; }
              body { font: 13px/1.5 "Segoe UI", system-ui, sans-serif; margin: 0; padding: 24px; }
              code { font-family: "Cascadia Mono", Consolas, monospace; }
            </style>
            <h1>xlide</h1>
            <p>The shell document is not installed.</p>
            <p>Expected at <code>{{escaped}}</code>.</p>
            """;
    }
}
