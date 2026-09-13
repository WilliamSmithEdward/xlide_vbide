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
    /// The application object, or null when no document window can be found in this process.
    /// The caller owns the result.
    /// </summary>
    public static DispatchObject? Find()
    {
        var pane = FindDocumentWindow();
        if (pane == 0)
        {
            Log.Warn($"host: no document window for {Engine.HostApp.Name} in this process, so its object model cannot be reached");
            return null;
        }

        var hr = Win32.AccessibleObjectFromWindow(pane, ObjectIdNativeObjectModel, in DispatchInterface, out var pointer);
        if (hr < 0 || pointer == 0)
        {
            Log.Error($"host: the document window did not answer with an object model, 0x{hr:X8}");
            return null;
        }

        using var window = DispatchObject.Attach(pointer);
        return window?.GetObject("Application");
    }

    /// <summary>
    /// Finds the host's document window in this process: the one OBJID_NATIVEOM answers on.
    ///
    /// PER HOST, because the classes are (2026-08-19, the day `?1+1` in Word answered "the host
    /// application could not be reached": this walked XLMAIN/XLDESK/EXCEL7 in every host).
    /// Excel keeps its fixed-shape walk; Word's document pane (`_WwG`) sits deeper and varies,
    /// so it is searched. A host with no classes wired here answers 0 and its callers report
    /// the same honest failure they always did - silence in an unwired host beats a guess.
    /// Only windows of this process count either way; a second copy of the host running
    /// alongside must not be the one answering.
    /// </summary>
    private static nint FindDocumentWindow() => Engine.HostApp.Name switch
    {
        "excel" => FindExcelWorksheetWindow(),
        "word" => FindWordDocumentPane(),

        // ACCESS ANSWERS ON ITS FRAME, and answers with the Application ITSELF rather than with a
        // window object - `Name` on it reads "Microsoft Access" (measured 2026-09-06 against a
        // live database). The `Application` hop below still holds, because an Office Application
        // answers that property with itself, so nothing above needs to know the difference.
        // Access has no document window of its own worth walking to: its MDI children are forms
        // and reports, and a database with none open would then have no object model at all.
        "access" => FindProcessFrame("OMain"),
        _ => 0,
    };

    /// <summary>
    /// Word's document pane, from whichever of this process's frames holds one.
    ///
    /// EVERY FRAME, the way the Excel walk below does. Word keeps more than one top-level
    /// `OpusApp` window, so asking only the first - which this did - makes the answer depend on
    /// window order. The gate's Word session answered "the host application could not be reached"
    /// for the Immediate window, the test runs and the dirty check on three runs in a row, while
    /// the same suite passed by hand on the same build (2026-09-13).
    ///
    /// A frame with no pane under it is still worth answering with when NO frame has one: Word's
    /// frame carries the object model itself, as Access's does, so a session whose document
    /// window has not been built yet reaches the Application rather than nothing.
    /// </summary>
    private static nint FindWordDocumentPane()
    {
        nint frame = 0;
        nint first = 0;
        var frames = 0;

        while ((frame = NextProcessFrame("OpusApp", frame)) != 0)
        {
            frames += 1;
            if (first == 0)
            {
                first = frame;
            }

            var pane = FindDescendantByClass(frame, "_WwG", 0);
            if (pane != 0)
            {
                return pane;
            }
        }

        if (first != 0)
        {
            Log.Verbose($"host: no _WwG under {frames} Word frame(s); asking the frame itself");
        }

        return first;
    }

    private static nint FindExcelWorksheetWindow()
    {
        nint frame = 0;
        while ((frame = NextProcessFrame("XLMAIN", frame)) != 0)
        {
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

    /// <summary>This process's first top-level window of the class, or 0.</summary>
    private static nint FindProcessFrame(string frameClass) => NextProcessFrame(frameClass, 0);

    private static nint NextProcessFrame(string frameClass, nint after)
    {
        var ours = Win32.GetCurrentProcessId();
        var frame = after;
        while ((frame = Win32.FindWindowEx(0, frame, frameClass, null)) != 0)
        {
            Win32.GetWindowThreadProcessId(frame, out var owner);
            if (owner == ours)
            {
                return frame;
            }
        }

        return 0;
    }

    /// <summary>
    /// Depth-first search for a descendant of the class, bounded: Word puts `_WwG` three
    /// levels under the frame today, and a bound keeps a strange window tree from becoming
    /// a hang rather than a miss.
    /// </summary>
    private static nint FindDescendantByClass(nint parent, string wantedClass, int depth)
    {
        if (parent == 0 || depth > 5)
        {
            return 0;
        }

        nint child = 0;
        while ((child = Win32.FindWindowEx(parent, child, null, null)) != 0)
        {
            if (Win32.ReadClassName(child) == wantedClass)
            {
                return child;
            }

            var below = FindDescendantByClass(child, wantedClass, depth + 1);
            if (below != 0)
            {
                return below;
            }
        }

        return 0;
    }
}
