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
    /// The application object, or null when no window in this process answers with one.
    /// The caller owns the result.
    ///
    /// EVERY CANDIDATE IS TRIED, not just the first. Finding a window of the right class is not
    /// the same as finding the one that answers: in Word the first `_WwG` under the first
    /// `OpusApp` frame returned E_FAIL to OBJID_NATIVEOM a hundred and six times in one gate
    /// session while a later one would have answered, and every feature that needs the host -
    /// the Immediate window, every test run, the dirty check - reported "The host application
    /// could not be reached" for the whole run (2026-09-13). The window search below now offers
    /// its candidates in order and this asks each in turn, which is the same correction the
    /// harness launcher needed the same day (lessons.md 85): an empty answer from COM is a
    /// failure to see, never a fact about the world.
    /// </summary>
    public static DispatchObject? Find()
    {
        var tried = 0;
        var lastFailure = 0;

        foreach (var candidate in DocumentWindows())
        {
            tried += 1;

            var hr = Win32.AccessibleObjectFromWindow(candidate, ObjectIdNativeObjectModel, in DispatchInterface, out var pointer);
            if (hr < 0 || pointer == 0)
            {
                lastFailure = hr;
                continue;
            }

            using var window = DispatchObject.Attach(pointer);
            var application = window?.GetObject("Application");
            if (application is not null)
            {
                if (tried > 1)
                {
                    Log.Verbose($"host: the object model came from candidate window {tried}");
                }

                return application;
            }
        }

        // ONE LINE PER LOOKUP, naming what was tried. The old code logged an error per attempt
        // and a failing session wrote the same sentence a hundred times, which buries the first
        // one - the only one whose timestamp says when the host stopped answering.
        if (tried == 0)
        {
            Log.Warn($"host: no document window for {Engine.HostApp.Name} in this process, so its object model cannot be reached");
        }
        else
        {
            Log.Error($"host: none of {tried} window(s) in this process answered with an object model, last 0x{lastFailure:X8}");
        }

        return null;
    }

    /// <summary>
    /// The windows in this process worth asking for the host's object model, best first.
    ///
    /// PER HOST, because the classes are (2026-08-19, the day `?1+1` in Word answered "the host
    /// application could not be reached": this walked XLMAIN/XLDESK/EXCEL7 in every host).
    /// A host with no classes wired here offers nothing and its callers report the same honest
    /// failure they always did - silence in an unwired host beats a guess. Only windows of this
    /// process are offered either way; a second copy of the host running alongside must not be
    /// the one answering.
    ///
    /// SEVERAL, because which one answers cannot be known from its class. The caller asks each
    /// in turn and keeps the first that hands back an object model.
    /// </summary>
    private static IEnumerable<nint> DocumentWindows() => Engine.HostApp.Name switch
    {
        "excel" => ExcelWorksheetWindows(),
        "word" => WordDocumentPanes(),

        // ACCESS ANSWERS ON ITS FRAME, and answers with the Application ITSELF rather than with a
        // window object - `Name` on it reads "Microsoft Access" (measured 2026-09-06 against a
        // live database). The `Application` hop above still holds, because an Office Application
        // answers that property with itself, so nothing there needs to know the difference.
        // Access has no document window of its own worth walking to: its MDI children are forms
        // and reports, and a database with none open would then have no object model at all.
        "access" => ProcessFrames("OMain"),
        _ => [],
    };

    /// <summary>
    /// Word's document panes, every `_WwG` under every `OpusApp` frame of this process, then the
    /// frames themselves.
    ///
    /// EVERY PANE, not the first one found. Word keeps more than one top-level `OpusApp` window
    /// and more than one pane under them, and which one answers OBJID_NATIVEOM is not something
    /// its class says: the first pane returned E_FAIL a hundred and six times in one gate
    /// session, and returning it and stopping meant the Immediate window, every test run and the
    /// dirty check all reported "The host application could not be reached" for the whole run
    /// (2026-09-13). Asking only the first frame had already caused the same outage a different
    /// way, which is what made walking every frame look like the whole fix.
    ///
    /// The frames come last, and on the Word measured here they are not what rescues anything:
    /// of the four candidates a two-frame session offers - two panes, two frames - exactly ONE
    /// answers and both frames refuse with E_FAIL (measured 2026-09-13 against a live session,
    /// then again with the frame order deliberately reversed). So the earlier fix for this, a
    /// fallback to the frame itself on the theory that a Word frame carries the object model as
    /// an Access frame does, could never have worked; the second PANE is what answers. They stay
    /// last because they cost one call each and a different Word build may differ.
    /// </summary>
    private static IEnumerable<nint> WordDocumentPanes()
    {
        var frames = new List<nint>();

        nint frame = 0;
        while ((frame = NextProcessFrame("OpusApp", frame)) != 0)
        {
            frames.Add(frame);
            foreach (var pane in DescendantsByClass(frame, "_WwG", 0))
            {
                yield return pane;
            }
        }

        foreach (var bare in frames)
        {
            yield return bare;
        }
    }

    /// <summary>
    /// Excel's worksheet windows, every `EXCEL7` under every `XLDESK` under every `XLMAIN` of
    /// this process. Same reason as Word's: a window of the right class is a candidate, not an
    /// answer.
    /// </summary>
    private static IEnumerable<nint> ExcelWorksheetWindows()
    {
        nint frame = 0;
        while ((frame = NextProcessFrame("XLMAIN", frame)) != 0)
        {
            nint desk = 0;
            while ((desk = Win32.FindWindowEx(frame, desk, "XLDESK", null)) != 0)
            {
                nint sheet = 0;
                while ((sheet = Win32.FindWindowEx(desk, sheet, "EXCEL7", null)) != 0)
                {
                    yield return sheet;
                }
            }
        }
    }

    /// <summary>Every top-level window of the class in this process.</summary>
    private static IEnumerable<nint> ProcessFrames(string frameClass)
    {
        nint frame = 0;
        while ((frame = NextProcessFrame(frameClass, frame)) != 0)
        {
            yield return frame;
        }
    }

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
    /// Depth-first walk of the descendants of the class, bounded: Word puts `_WwG` three
    /// levels under the frame today, and a bound keeps a strange window tree from becoming
    /// a hang rather than a miss.
    ///
    /// ALL OF THEM, in the order the tree gives them. The caller wants the first that ANSWERS,
    /// which is not always the first that exists, and stopping at one match is what hid a
    /// second, working pane behind a failing one.
    /// </summary>
    private static IEnumerable<nint> DescendantsByClass(nint parent, string wantedClass, int depth)
    {
        if (parent == 0 || depth > 5)
        {
            yield break;
        }

        nint child = 0;
        while ((child = Win32.FindWindowEx(parent, child, null, null)) != 0)
        {
            if (Win32.ReadClassName(child) == wantedClass)
            {
                yield return child;
                continue;
            }

            foreach (var below in DescendantsByClass(child, wantedClass, depth + 1))
            {
                yield return below;
            }
        }
    }
}
