using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// Finds the native dialogs standing in this process, and dismisses one by button.
///
/// Every route that reads the session marshals to the host thread and gives up after three
/// seconds. That is honest but blind: a modal dialog owns the host thread until something
/// answers it, so the one moment a person most wants to know what is happening is the moment
/// every other route can only say "the host thread did not answer in time". Twice in one day
/// a probe left a dialog standing - an Add Watch whose filler mis-parsed its arguments, a
/// Macros dialog raised by a Run with the caret on line one - and both times the editor
/// simply stopped, with nothing able to say why.
///
/// Nothing here touches the host thread. Window enumeration, class and caption reads, and a
/// posted button click all work from any thread, which is exactly why this can answer while
/// the editor is stuck. That also makes it the one instrument that stays useful during the
/// failure it was built for.
/// </summary>
internal static unsafe class DialogWatch
{
    /// <summary>The window class every common dialog and VBA dialog uses.</summary>
    private const string DialogClass = "#32770";

    /// <summary>BM_CLICK, posted rather than sent: a click that opens another modal must not
    /// block the api's own thread inside the nested loop it starts.</summary>
    private const uint BmClick = 0x00F5;

    /// <summary>One dialog: what it is, what it says, and which buttons would answer it.</summary>
    public sealed record DialogRow(
        string Window,
        string Caption,
        string Text,
        string[] Buttons,
        bool Enabled);

    /// <summary>
    /// Buttons that only acknowledge. A dialog offering nothing else is a NOTICE, not a question:
    /// it has already happened, and pressing OK changes nothing but the fact that it is on screen.
    /// </summary>
    private static readonly string[] Acknowledgements = ["OK", "Help", "Close", "Continue"];

    /// <summary>
    /// Buttons that decline. Preferred on a real question, because declining is the answer that
    /// cannot destroy anything.
    /// </summary>
    private static readonly string[] Declines = ["Cancel", "No", "Close"];

    /// <summary>Buttons that agree. Only ever pressed inside <see cref="ExpectingConfirmation"/>,
    /// where the thing being agreed to is the thing this product just asked for.</summary>
    private static readonly string[] Confirmations = ["OK", "Yes", "Continue"];

    private static int _expecting;
    private static long _expectingUntil;

    /// <summary>
    /// Says that a confirmation raised in the next moment is the confirmation of a command THIS
    /// PRODUCT deliberately issued, so agreeing to it is finishing what we started rather than
    /// deciding something on the developer's behalf.
    ///
    /// WHY THE SCOPE RATHER THAN THE CAPTION. The obvious version matches the dialog's text -
    /// "This action will reset your project" - and that text is localised, so it would work here
    /// and quietly stop working on a German or Japanese Office, leaving exactly the deadlock this
    /// exists to prevent on the machines least able to report it. A scope held around our own call
    /// needs no words at all.
    ///
    /// TIME-BOUNDED, because the dialog is answered by another thread and a scope that leaked
    /// would turn every later question into a yes. Five seconds is far longer than the gap
    /// between issuing a command and its confirmation appearing, and far shorter than the gap to
    /// anything a developer would raise next.
    /// </summary>
    /// <remarks>
    /// IT ALSO ANSWERS, and that is not a convenience. Until now this only decided WHAT an
    /// answerer would press, and the only answerer is the api's rescue for a request stuck behind
    /// a dialog - so the mechanism worked exactly where a request happened to be in flight and
    /// nowhere else. A caller with no request behind it - a timer, a poll, anything reacting to
    /// the editor rather than to somebody asking - would have issued its command, raised a
    /// confirmation, and left the host thread parked on a modal nobody was going to press.
    ///
    /// So the scope carries its own short-lived watcher on a pool thread. What it will press is
    /// tightly bounded, because pressing OK on the wrong dialog is the one thing worse than
    /// pressing nothing:
    ///
    ///   only a dialog that was NOT already standing when the scope opened, so a question the
    ///     developer raised a moment earlier is never touched;
    ///   only ONE, after which the watcher stops, so a scope cannot become a policy;
    ///   only until the scope is disposed or its deadline passes.
    /// </remarks>
    public static IDisposable ExpectingConfirmation(int forMs = 5000)
    {
        Interlocked.Increment(ref _expecting);
        Volatile.Write(ref _expectingUntil, Environment.TickCount64 + forMs);

        var expectation = new Expectation();

        // What was ALREADY on screen. Anything in here is somebody else's question.
        var standing = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            foreach (var dialog in Dialogs())
            {
                standing.Add(dialog.Window);
            }
        }
        catch
        {
            // A window list that will not read leaves the set empty, which only makes the
            // watcher more cautious: every dialog then looks pre-existing and none is pressed.
            return expectation;
        }

        // A plain blocking loop rather than an async one: this file is compiled unsafe, so it
        // cannot await, and a pool thread sleeping 60ms at a time for a few seconds is what the
        // pool is for.
        _ = Task.Run(() =>
        {
            var until = Environment.TickCount64 + forMs;
            while (Environment.TickCount64 < until && !expectation.Closed)
            {
                try
                {
                    foreach (var dialog in Dialogs())
                    {
                        if (standing.Contains(dialog.Window) || SafeAnswerFor(dialog) is not { } press)
                        {
                            continue;
                        }

                        Log.Info($"dialog watch: \"{dialog.Caption}\" arrived inside an expected "
                            + $"confirmation; answering with {press}");
                        Dismiss(dialog.Caption, press);
                        return;
                    }
                }
                catch
                {
                    // The window list is a best effort; a tick that cannot read it tries again.
                }

                Thread.Sleep(60);
            }
        });

        return expectation;
    }

    private sealed class Expectation : IDisposable
    {
        private int _closed;

        /// <summary>Whether the scope has been let go, so its watcher can stop.</summary>
        public bool Closed => Volatile.Read(ref _closed) == 1;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _closed, 1) == 0)
            {
                Interlocked.Decrement(ref _expecting);
            }
        }
    }

    /// <summary>Whether a confirmation standing right now would be one we asked for.</summary>
    private static bool Expecting =>
        Volatile.Read(ref _expecting) > 0
        && Environment.TickCount64 <= Volatile.Read(ref _expectingUntil);

    /// <summary>Every visible dialog this process owns.</summary>
    public static DialogRow[] Dialogs()
    {
        var rows = new List<DialogRow>();
        foreach (var dialog in TopLevelDialogs())
        {
            var buttons = new List<string>();
            var said = new List<string>();

            foreach (var child in ChildrenOf(dialog))
            {
                var kind = Win32.ReadClassName(child);

                if (kind.Equals("Button", StringComparison.OrdinalIgnoreCase)
                    && Plain(Win32.ReadWindowText(child)) is { Length: > 0 } caption
                    && buttons.Count < 16)
                {
                    buttons.Add(caption);
                    continue;
                }

                // What the dialog SAYS, not just what it is called. Every VBA compile error wears
                // the same caption - "Microsoft Visual Basic for Applications" - so the caption
                // alone cannot tell "Ambiguous name detected: Recalculate" from anything else,
                // and a harness that only reads captions learns nothing about what went wrong
                // (2026-08-07).
                if (kind.Equals("Static", StringComparison.OrdinalIgnoreCase)
                    && Win32.ReadWindowText(child) is { Length: > 0 } line
                    && said.Count < 8)
                {
                    said.Add(line.Trim());
                }
            }

            rows.Add(new DialogRow(
                $"0x{dialog:X}",
                Win32.ReadWindowText(dialog),
                string.Join(" ", said).Trim(),
                [.. buttons],
                Win32.IsWindowEnabled(dialog)));
        }

        return [.. rows];
    }

    /// <summary>
    /// Whether this dialog only reports. Every button an acknowledgement means nothing is being
    /// asked: the thing has already happened, and the dialog is the record of it.
    /// </summary>
    public static bool IsNotice(DialogRow dialog) =>
        dialog.Buttons.Length > 0
        && dialog.Buttons.All(button => Acknowledgements.Contains(button, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// The button that answers a dialog without deciding anything, or null when every button it
    /// offers commits to something.
    ///
    /// A NOTICE - every button an acknowledgement - is always safe: it is reporting, not asking.
    /// That is the case the old policy missed. It would press only Cancel, Close or No, so a
    /// compile error offering OK and Help matched nothing and stood for six minutes with the host
    /// thread behind it (2026-08-07). A real question is still only ever declined - UNLESS it is
    /// the confirmation of a command this product itself issued a moment ago, which is what
    /// <see cref="ExpectingConfirmation"/> declares.
    ///
    /// That exception exists because declining without it is not neutral either. The immediate
    /// window recovers a stopped project by issuing Reset; Reset asks "proceed anyway?"; this
    /// answered Cancel, so the recovery could never complete, and every later evaluation tried
    /// again until the editor faulted and took Excel with it (issue #6, found by chaos.mjs).
    /// Declining a question nobody here asked is safe. Declining our own is just a way of never
    /// finishing anything.
    /// </summary>
    public static string? SafeAnswerFor(DialogRow dialog)
    {
        if (IsNotice(dialog))
        {
            return dialog.Buttons.FirstOrDefault(button =>
                button.Equals("OK", StringComparison.OrdinalIgnoreCase))
                ?? dialog.Buttons[0];
        }

        if (Expecting
            && Confirmations.FirstOrDefault(button =>
                dialog.Buttons.Contains(button, StringComparer.OrdinalIgnoreCase)) is { } agree)
        {
            return agree;
        }

        return Declines.FirstOrDefault(button =>
            dialog.Buttons.Contains(button, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Answers a dialog by clicking one of its buttons, matched on its caption without the
    /// ampersand the accelerator uses ("&amp;Cancel" answers to "Cancel"). Posted, not sent,
    /// so a dialog that raises another one cannot hold this thread inside it.
    /// </summary>
    /// <summary>
    /// What this watch last took off the screen, and the tick it did it on.
    ///
    /// Evidence, for callers that have to tell one kind of stop from another and have nothing
    /// else to go on. Empty until something has been answered.
    /// </summary>
    public static (string Said, long At) LastAnswered { get; private set; } = (string.Empty, 0);

    /// <summary>Whether this watch answered something matching `phrase` within the last `withinMs`.</summary>
    public static bool AnsweredRecently(string phrase, int withinMs = 15000)
    {
        var (said, at) = LastAnswered;
        return at > 0
            && Environment.TickCount64 - at <= withinMs
            && said.Contains(phrase, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>What a dialog's static text says, joined - the same reading `Dialogs` reports.</summary>
    private static string Said(nint dialog)
    {
        var said = new List<string>();
        foreach (var child in ChildrenOf(dialog))
        {
            if (Win32.ReadClassName(child).Equals("Static", StringComparison.OrdinalIgnoreCase)
                && Win32.ReadWindowText(child) is { Length: > 0 } line
                && said.Count < 8)
            {
                said.Add(line.Trim());
            }
        }

        return string.Join(" ", said).Trim();
    }

    public static bool Dismiss(string? caption, string button)
    {
        foreach (var dialog in TopLevelDialogs())
        {
            var title = Win32.ReadWindowText(dialog);
            if (caption is { Length: > 0 }
                && !title.Contains(caption, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            foreach (var child in ChildrenOf(dialog))
            {
                if (!Win32.ReadClassName(child).Equals("Button", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (Plain(Win32.ReadWindowText(child)).Equals(Plain(button), StringComparison.OrdinalIgnoreCase))
                {
                    // WHAT IT SAID, KEPT, and read BEFORE the click - a dismissed dialog cannot
                    // be asked afterwards. Every VBA compile error wears the same caption, so
                    // the caption alone tells a caller nothing; the body is the only thing that
                    // separates "Compile error" from any other box this answers. A test run
                    // stopped behind a compile error and a developer's own breakpoint inside a
                    // test look identical from every other angle, and one of them must never be
                    // cleared (#10).
                    LastAnswered = (Said(dialog), Environment.TickCount64);
                    Log.Info($"dialog watch: answering \"{title}\" with \"{button}\"");
                    Win32.PostMessage(child, BmClick, 0, 0);
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>The window class a RUNNING form wears - the forms runtime's own frame.</summary>
    private const string RunFormClass = "ThunderDFrame";

    /// <summary>WM_CLOSE: what clicking a form's X posts, QueryClose and all.</summary>
    private const uint WmClose = 0x0010;

    /// <summary>
    /// The captions of every form RUNNING in this process right now. Off the host thread on
    /// purpose, like everything here: a modally running form holds the host thread inside the
    /// Run command until it closes, and this is the instrument that watches it do so.
    /// </summary>
    public static string[] RunningForms() =>
        [.. TopLevelRunForms().Select(Win32.ReadWindowText)];

    /// <summary>
    /// The window handle of a running form, matched on caption the way the close is, or the
    /// first one standing when no caption is given. Zero when nothing matches.
    ///
    /// What it is FOR: a running form is the only picture of a designer's work that nothing in
    /// the object model can answer for. MSForms draws its controls windowless, so there are no
    /// child handles to enumerate and the designer's own collection describes the STORED form,
    /// not the one on screen. Capturing the window is the only way to check the two agree - and
    /// that is exactly the question F5 raised (the owner, 2026-08-16: "are you checking against
    /// the live form?").
    /// </summary>
    public static nint RunningFormHandle(string? caption)
    {
        foreach (var form in TopLevelRunForms())
        {
            if (caption is not { Length: > 0 }
                || Win32.ReadWindowText(form).Contains(caption, StringComparison.OrdinalIgnoreCase))
            {
                return form;
            }
        }

        return 0;
    }

    /// <summary>
    /// Closes a running form the way its X would - WM_CLOSE, posted - matched on caption, or
    /// the first one standing when no caption is given. False when nothing matched.
    /// </summary>
    public static bool CloseRunningForm(string? caption)
    {
        foreach (var form in TopLevelRunForms())
        {
            var title = Win32.ReadWindowText(form);
            if (caption is { Length: > 0 }
                && !title.Contains(caption, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            Log.Info($"dialog watch: closing the running form \"{title}\"");
            Win32.PostMessage(form, WmClose, 0, 0);
            return true;
        }

        return false;
    }

    /*
     * The enumeration callbacks must be static and unmanaged, so the window each pass
     * collects goes into a static list behind a gate. Collection is the only thing done
     * under it: the callers walk what it returns with nothing held.
     */
    private static readonly Lock Gate = new();
    private static readonly List<nint> Collected = [];

    private static nint[] TopLevelDialogs()
    {
        lock (Gate)
        {
            Collected.Clear();
            Win32.EnumWindows((nint)(delegate* unmanaged<nint, nint, int>)&OnTopLevel, 0);
            return [.. Collected];
        }
    }

    private static nint[] TopLevelRunForms()
    {
        lock (Gate)
        {
            Collected.Clear();
            Win32.EnumWindows((nint)(delegate* unmanaged<nint, nint, int>)&OnTopLevelForm, 0);
            return [.. Collected];
        }
    }

    private static nint[] ChildrenOf(nint parent)
    {
        lock (Gate)
        {
            Collected.Clear();
            Win32.EnumChildWindows(parent, (nint)(delegate* unmanaged<nint, nint, int>)&OnChild, 0);
            return [.. Collected];
        }
    }

    [UnmanagedCallersOnly]
    private static int OnTopLevel(nint window, nint parameter)
    {
        uint owner;
        Win32.GetWindowThreadProcessId(window, &owner);
        if (owner == Win32.GetCurrentProcessId()
            && Win32.IsWindowVisible(window)
            && Win32.ReadClassName(window) == DialogClass)
        {
            Collected.Add(window);
        }

        return Collected.Count < 32 ? 1 : 0;
    }

    [UnmanagedCallersOnly]
    private static int OnTopLevelForm(nint window, nint parameter)
    {
        uint owner;
        Win32.GetWindowThreadProcessId(window, &owner);
        if (owner == Win32.GetCurrentProcessId()
            && Win32.IsWindowVisible(window)
            && Win32.ReadClassName(window) == RunFormClass)
        {
            Collected.Add(window);
        }

        return Collected.Count < 32 ? 1 : 0;
    }

    [UnmanagedCallersOnly]
    private static int OnChild(nint window, nint parameter)
    {
        Collected.Add(window);
        return Collected.Count < 128 ? 1 : 0;
    }

    private static string Plain(string text) =>
        text.Replace("&", string.Empty, StringComparison.Ordinal).Trim();

}
