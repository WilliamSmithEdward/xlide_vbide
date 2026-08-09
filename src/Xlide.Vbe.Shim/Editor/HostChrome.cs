using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The title bar of the editor's own window: its caption and its icon.
///
/// Everything inside that window is already ours, so a title bar still announcing the product it
/// replaced is the one place the illusion breaks, and it is the place a person looks when they are
/// alt-tabbing back to what they were doing.
///
/// Only the product name is replaced. The editor's caption also carries the workbook and the module
/// being edited, and that half is the useful half: "xlide - Book1.xlsm" tells someone which window
/// they are looking at in a way that "xlide" alone does not.
///
/// The editor rewrites its own caption whenever the active project or module changes, so this is
/// re-applied on every name change rather than set once. Both the caption and the icons are put
/// back at shutdown: the window belongs to the host, and an add-in that unloads should leave it as
/// it found it.
/// </summary>
internal sealed class HostChrome : IDisposable
{
    private const string HostName = "Microsoft Visual Basic for Applications";
    // Capitalised here and nowhere else. In a title bar it is a wordmark sitting beside the
    // workbook's name, which is how the surface's own corner spells it; in prose it is a product
    // called xlide.
    private const string OurName = "XLIDE";

    private readonly nint _window;
    private readonly string? _originalCaption;
    private readonly nint _previousSmall;
    private readonly nint _previousBig;

    private nint _small;
    private nint _big;
    private bool _disposed;

    private HostChrome(nint window, string? originalCaption, nint previousSmall, nint previousBig)
    {
        _window = window;
        _originalCaption = originalCaption;
        _previousSmall = previousSmall;
        _previousBig = previousBig;
    }

    /// <summary>
    /// Takes over the window's caption and icon. Answers null when there is no window to take over
    /// or nothing to put on it, because a title bar left alone is a far better outcome than one
    /// left half-changed.
    /// </summary>
    public static HostChrome? Install(nint window, string? shimDirectory)
    {
        if (window == 0)
        {
            return null;
        }

        var caption = ReadCaption(window);
        var previousSmall = Win32.SendMessage(window, Win32.WmGetIcon, Win32.IconSmall, 0);
        var previousBig = Win32.SendMessage(window, Win32.WmGetIcon, Win32.IconBig, 0);

        var chrome = new HostChrome(window, caption, previousSmall, previousBig);
        chrome.LoadIcons(shimDirectory);
        chrome.Apply();

        Log.Info($"host chrome: caption and icon taken over on window {window:X}");
        return chrome;
    }

    /// <summary>
    /// Puts our name and icon back on the window. Idempotent, and cheap when nothing has changed:
    /// setting the caption raises the very name change that brings us back here, so a version that
    /// wrote unconditionally would chase its own tail.
    /// </summary>
    public void Apply()
    {
        if (_disposed || _window == 0)
        {
            return;
        }

        var ours = Compose();
        if (ReadCaption(_window) != ours)
        {
            Win32.SetWindowText(_window, ours);
        }

        // The icons go on regardless of the caption. The caption is compared against what the
        // window actually reads rather than against what was written last time, so a caption
        // changed by anything else is corrected on the next pass - and the icons are cheap enough
        // that a comparison to skip them would cost more than the messages do.
        if (_small != 0)
        {
            Win32.SendMessage(_window, Win32.WmSetIcon, Win32.IconSmall, _small);
        }

        if (_big != 0)
        {
            Win32.SendMessage(_window, Win32.WmSetIcon, Win32.IconBig, _big);
        }
    }

    /// <summary>
    /// The mode the editor is in, as the title bar should say it: "design", "running", "break".
    ///
    /// ALWAYS SAID, including design. It was left blank there at first, on the reasoning that not
    /// debugging is not news; the developer's answer was that a mode you only see sometimes is one
    /// you cannot rely on reading, and a window that says nothing is indistinguishable from one
    /// that has stopped reporting. Design is therefore stated too.
    ///
    /// Defaulted rather than left null, so the very first caption is right: the mode is set by the
    /// session as it changes, and the first change may be some way after the window appears.
    ///
    /// The editor puts its own "[break]" on the caption and takes it off again, but only sometimes
    /// and never for design, so it is stripped and restated from what this product actually knows
    /// rather than parsed back out of a string.
    /// </summary>
    public string? Mode { get; set; } = "design";

    /// <summary>The workbook on screen, as the tab strip and the tree name it.</summary>
    public string? Workbook { get; set; }

    /// <summary>The module on screen. Null when the workspace is empty.</summary>
    public string? Module { get; set; }

    /// <summary>
    /// The caption this product wants on the window, BUILT FROM WHAT THIS PRODUCT KNOWS rather
    /// than parsed out of what the editor last wrote.
    ///
    /// WHY NOTHING IS PARSED ANY MORE. Two versions tried to keep the editor's own caption and
    /// edit it, and both grew: the frame is an MDI frame, and Windows appends " - [Child (Code)]"
    /// to whatever the caption CURRENTLY says whenever the maximised child changes -- to ours, not
    /// to the editor's. Keeping the tail grew a segment per tab click; keeping only the last
    /// segment fixed that and still showed the module twice, because the assumption about where
    /// the editor puts the module was wrong to begin with.
    ///
    /// The session already knows the workbook and the module: they are the same values it puts on
    /// the tab strip and the tree. Composing from those cannot accumulate, cannot double, and
    /// cannot drift with whatever the editor decides its own caption should say. The one thing
    /// given up is that the caption no longer echoes the editor's wording, which was never the
    /// point of taking it over.
    /// </summary>
    private string Compose()
    {
        var built = new System.Text.StringBuilder(OurName);

        if (Mode is { Length: > 0 } mode)
        {
            built.Append(" [").Append(mode).Append(']');
        }

        if (Workbook is { Length: > 0 } workbook)
        {
            built.Append(" - ").Append(workbook);
        }

        if (Module is { Length: > 0 } module)
        {
            built.Append(" - ").Append(module);
        }

        return built.ToString();
    }

    private void LoadIcons(string? shimDirectory)
    {
        if (shimDirectory is null)
        {
            return;
        }

        var file = Path.Combine(shimDirectory, "xlide.ico");
        if (!File.Exists(file))
        {
            Log.Warn($"host chrome: no icon at {file}, leaving the editor's own");
            return;
        }

        // Asked for at the sizes the shell asks for, so the right image in the file is chosen
        // rather than one being scaled from whichever happens to come first.
        _small = Win32.LoadImage(0, file, Win32.ImageIcon,
            Win32.GetSystemMetrics(Win32.SmCxSmIcon), Win32.GetSystemMetrics(Win32.SmCxSmIcon), Win32.LrLoadFromFile);
        _big = Win32.LoadImage(0, file, Win32.ImageIcon,
            Win32.GetSystemMetrics(Win32.SmCxIcon), Win32.GetSystemMetrics(Win32.SmCxIcon), Win32.LrLoadFromFile);

        if (_small == 0 && _big == 0)
        {
            Log.Warn($"host chrome: {file} could not be loaded as an icon");
        }
    }

    private static unsafe string? ReadCaption(nint window)
    {
        const int Capacity = 512;
        var buffer = stackalloc char[Capacity];
        var length = Win32.GetWindowText(window, buffer, Capacity);
        return length > 0 ? new string(buffer, 0, length) : null;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // The caption is restored to what the window had when we arrived rather than to a computed
        // one: by now it names a different module, and the editor will correct it the moment
        // anything moves. Leaving OUR name on a window we no longer draw in is the failure worth
        // avoiding.
        if (_originalCaption is not null && _window != 0)
        {
            Win32.SetWindowText(_window, _originalCaption);
        }

        if (_window != 0)
        {
            Win32.SendMessage(_window, Win32.WmSetIcon, Win32.IconSmall, _previousSmall);
            Win32.SendMessage(_window, Win32.WmSetIcon, Win32.IconBig, _previousBig);
        }

        // Ours to destroy: LoadImage without LR_SHARED hands over ownership.
        if (_small != 0) Win32.DestroyIcon(_small);
        if (_big != 0) Win32.DestroyIcon(_big);
        _small = 0;
        _big = 0;

        Log.Info("host chrome: the editor's own caption and icon are back");
    }
}
