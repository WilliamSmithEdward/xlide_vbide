#if DEBUG
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
    /// thread behind it (2026-08-07). A real question is still only ever declined.
    /// </summary>
    public static string? SafeAnswerFor(DialogRow dialog)
    {
        if (IsNotice(dialog))
        {
            return dialog.Buttons.FirstOrDefault(button =>
                button.Equals("OK", StringComparison.OrdinalIgnoreCase))
                ?? dialog.Buttons[0];
        }

        return Declines.FirstOrDefault(button =>
            dialog.Buttons.Contains(button, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Answers a dialog by clicking one of its buttons, matched on its caption without the
    /// ampersand the accelerator uses ("&amp;Cancel" answers to "Cancel"). Posted, not sent,
    /// so a dialog that raises another one cannot hold this thread inside it.
    /// </summary>
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
#endif
