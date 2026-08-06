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

    /// <summary>One dialog: what it is, and which buttons would answer it.</summary>
    public sealed record DialogRow(string Window, string Caption, string[] Buttons, bool Enabled);

    /// <summary>Every visible dialog this process owns.</summary>
    public static DialogRow[] Dialogs()
    {
        var rows = new List<DialogRow>();
        foreach (var dialog in TopLevelDialogs())
        {
            var buttons = new List<string>();
            foreach (var child in ChildrenOf(dialog))
            {
                if (ReadClass(child).Equals("Button", StringComparison.OrdinalIgnoreCase)
                    && Plain(ReadText(child)) is { Length: > 0 } caption
                    && buttons.Count < 16)
                {
                    buttons.Add(caption);
                }
            }

            rows.Add(new DialogRow(
                $"0x{dialog:X}",
                ReadText(dialog),
                [.. buttons],
                Win32.IsWindowEnabled(dialog)));
        }

        return [.. rows];
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
            var title = ReadText(dialog);
            if (caption is { Length: > 0 }
                && !title.Contains(caption, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            foreach (var child in ChildrenOf(dialog))
            {
                if (!ReadClass(child).Equals("Button", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (Plain(ReadText(child)).Equals(Plain(button), StringComparison.OrdinalIgnoreCase))
                {
                    Log.Info($"dialog watch: answering \"{title}\" with \"{button}\"");
                    Win32.PostMessage(child, BmClick, 0, 0);
                    return true;
                }
            }
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
            && ReadClass(window) == DialogClass)
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

    private static string ReadText(nint window)
    {
        const int capacity = 256;
        var buffer = stackalloc char[capacity];
        var length = Win32.GetWindowText(window, buffer, capacity);
        return length <= 0 ? string.Empty : new string(buffer, 0, length);
    }

    private static string ReadClass(nint window)
    {
        const int capacity = 64;
        var buffer = stackalloc char[capacity];
        var length = Win32.GetClassName(window, buffer, capacity);
        return length <= 0 ? string.Empty : new string(buffer, 0, length);
    }
}
#endif
