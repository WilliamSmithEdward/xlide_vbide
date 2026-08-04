using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Sits in the editor frame's message chain to re-place the surface INSIDE a resize, before the
/// native layout paints.
///
/// The window-event route is correct but late: events are posted, so on every tick of a resize
/// drag the frame relayouts and paints its native chrome first, and the surface catches up a
/// beat later — the old editor flickering through the new one, which is exactly what it looked
/// like (developer's report, 2026-08-04). A subclass receives WM_SIZE synchronously, before the
/// frame's own handler moves a single native band, so the surface is already covering the new
/// rectangle when anything under it paints.
///
/// The subclass must be removed before the host tears down, same as every hook: a message chain
/// entry pointing into an unloaded library is a crash on the next message.
/// </summary>
internal sealed unsafe class FrameSubclass : IDisposable
{
    /// <summary>One per session, reached from the static callback the way the event hook is.</summary>
    private static FrameSubclass? _current;

    private readonly Action _sized;
    private nint _frame;

    private FrameSubclass(nint frame, Action sized)
    {
        _frame = frame;
        _sized = sized;
    }

    /// <summary>
    /// Installs over the frame, or returns null and lets the event route carry on alone.
    /// </summary>
    public static FrameSubclass? Install(nint frame, Action sized)
    {
        ArgumentNullException.ThrowIfNull(sized);

        if (frame == 0)
        {
            return null;
        }

        if (_current is not null)
        {
            // One editor frame per session; a second install is a lifecycle bug worth hearing about.
            Log.Warn("frame subclass: already installed, refusing a second");
            return null;
        }

        var subclass = new FrameSubclass(frame, sized);
        _current = subclass;

        if (!Win32.SetWindowSubclass(
            frame,
            (nint)(delegate* unmanaged<nint, uint, nint, nint, nuint, nuint, nint>)&OnMessage,
            SubclassId,
            0))
        {
            Log.Warn("frame subclass: could not be installed; resize follows events only");
            _current = null;
            return null;
        }

        Log.Info($"frame subclass: installed on {frame:X}");
        return subclass;
    }

    private const nuint SubclassId = 1;

    private const uint WmSize = 0x0005;

    [UnmanagedCallersOnly]
    private static nint OnMessage(nint window, uint message, nint wParam, nint lParam, nuint id, nuint reference)
    {
        // Nothing may escape into the host's message dispatch, and nothing here may be slow.
        try
        {
            if (message == WmSize && _current is { } current && window == current._frame)
            {
                // Before the frame's own handler runs: the client rectangle already has its new
                // size, and none of the native children have been laid out or painted yet, so
                // the surface placed here is covering them from the first pixel.
                Log.Verbose("frame subclass: WM_SIZE, placing the surface before the native layout");
                current._sized();
            }
        }
        catch (Exception ex)
        {
            Log.Error("frame subclass: the resize callback failed", ex);
        }

        return Win32.DefSubclassProc(window, message, wParam, lParam);
    }

    public void Dispose()
    {
        var frame = _frame;
        _frame = 0;

        if (frame != 0)
        {
            Win32.RemoveWindowSubclass(
                frame,
                (nint)(delegate* unmanaged<nint, uint, nint, nint, nuint, nuint, nint>)&OnMessage,
                SubclassId);
            Log.Info("frame subclass: removed");
        }

        if (ReferenceEquals(_current, this))
        {
            _current = null;
        }
    }
}
