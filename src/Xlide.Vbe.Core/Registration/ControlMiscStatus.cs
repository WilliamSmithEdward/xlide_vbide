namespace Xlide.Vbe.Core.Registration;

/// <summary>
/// Status bits that tell a container how to treat an embedded control.
///
/// They live here rather than beside the rest of the embedding code because a container may read
/// them from two places: from the class registration before the object exists, or from the object
/// itself once it does. The two answers have to agree, and the only way to guarantee that is for
/// both to come from the same constant.
/// </summary>
public static class ControlMiscStatus
{
    /// <summary>The control redraws itself when resized rather than scaling its last rendering.</summary>
    public const int RecomposeOnResize = 0x00000001;

    /// <summary>The control cannot be the target of a link. It has no persistent location to link to.</summary>
    public const int CannotLinkInside = 0x00000010;

    /// <summary>
    /// The control is active whenever it is visible, rather than showing a static picture until the
    /// user activates it. Every control that hosts live content is inside-out.
    /// </summary>
    public const int InsideOut = 0x00000080;

    /// <summary>The container should activate the control as soon as it becomes visible.</summary>
    public const int ActivateWhenVisible = 0x00000100;

    /// <summary>
    /// The container must supply the client site before asking the control to load or initialise
    /// its state. Without it, the control is asked to initialise while it has no way to reach the
    /// container.
    /// </summary>
    public const int SetClientSiteFirst = 0x00020000;

    /// <summary>
    /// What the tool window host reports, and what its class registration advertises.
    /// </summary>
    public const int ToolWindowHost =
        RecomposeOnResize | CannotLinkInside | InsideOut | ActivateWhenVisible | SetClientSiteFirst;
}
