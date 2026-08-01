using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Com;

/// <summary>
/// The single wrapper table for the whole shim.
///
/// This is deliberately one instance and not one per call site. A ComWrappers instance caches the
/// mapping between a managed object and its COM identity, so two instances hand out two different
/// unknowns for the same object. COM defines identity as pointer equality on IUnknown, and a
/// container that receives our control through one path and compares it against the same control
/// received through another would decide they are different objects. That failure is silent and
/// arbitrarily delayed, which is the worst combination.
/// </summary>
internal static class ComRuntime
{
    /// <summary>Wrapper table used for every managed-to-COM and COM-to-managed transition here.</summary>
    public static readonly StrategyBasedComWrappers Wrappers = new();
}
