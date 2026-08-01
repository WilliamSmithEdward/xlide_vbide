namespace Xlide.Vbe.Shim.WebView;

/// <summary>
/// What a browser surface is for.
///
/// The two surfaces show unrelated documents and are created in unrelated places, so which one a
/// surface shows is stated by whoever creates it. An earlier version inferred it from what happened
/// to exist on disk, which put the editing surface inside the docked panel; the result looked like a
/// rendering fault rather than the wrong document, which is exactly why the choice is explicit now.
/// </summary>
internal enum SurfaceContent
{
    /// <summary>A docked panel: the problems list and the views beside it.</summary>
    Panel,

    /// <summary>The editing surface that sits over a code pane.</summary>
    Editor,
}
