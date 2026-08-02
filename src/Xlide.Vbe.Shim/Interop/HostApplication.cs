using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>
/// Reaches the host application's own automation object from inside its process.
///
/// The editor does not expose it. Its object model describes projects, components and windows, and
/// stops at the boundary of the application hosting it: there is no property anywhere on it that
/// answers "the program I am part of". Running a procedure by name is the host's job, so it has to
/// be found another way.
///
/// A worksheet window answers with it. Asking a window for the object model behind it is a
/// documented accessibility call, and the answer is the same object automation would have handed
/// out. This does not go through the running object table, which publishes lazily and would make
/// the first evaluation of a session wait tens of seconds for a table entry that has nothing to do
/// with the question.
/// </summary>
internal static class HostApplication
{
    /// <summary>The host's own object model, rather than an accessibility view of the window.</summary>
    private const uint ObjectIdNativeObjectModel = 0xFFFFFFF0;

    /// <summary>IDispatch, which is what the call is asked to hand back.</summary>
    private static readonly Guid DispatchInterface = new("00020400-0000-0000-C000-000000000046");

    /// <summary>
    /// The application object, or null when no worksheet window can be found in this process.
    /// The caller owns the result.
    /// </summary>
    public static DispatchObject? Find()
    {
        var sheet = FindWorksheetWindow();
        if (sheet == 0)
        {
            Log.Warn("host: no worksheet window in this process, so its object model cannot be reached");
            return null;
        }

        var hr = Win32.AccessibleObjectFromWindow(sheet, ObjectIdNativeObjectModel, in DispatchInterface, out var pointer);
        if (hr < 0 || pointer == 0)
        {
            Log.Error($"host: the worksheet window did not answer with an object model, 0x{hr:X8}");
            return null;
        }

        using var window = DispatchObject.Attach(pointer);
        return window?.GetObject("Application");
    }

    /// <summary>
    /// Finds a worksheet window belonging to this process.
    ///
    /// Walked by class rather than enumerated, because the shape is fixed and known: the frame
    /// holds a desktop which holds the worksheet windows. Only windows of this process count; a
    /// second copy of the host running alongside must not be the one answering.
    /// </summary>
    private static nint FindWorksheetWindow()
    {
        var ours = Win32.GetCurrentProcessId();

        nint frame = 0;
        while ((frame = Win32.FindWindowEx(0, frame, "XLMAIN", null)) != 0)
        {
            Win32.GetWindowThreadProcessId(frame, out var owner);
            if (owner != ours)
            {
                continue;
            }

            var desk = Win32.FindWindowEx(frame, 0, "XLDESK", null);
            if (desk == 0)
            {
                continue;
            }

            var sheet = Win32.FindWindowEx(desk, 0, "EXCEL7", null);
            if (sheet != 0)
            {
                return sheet;
            }
        }

        return 0;
    }
}
