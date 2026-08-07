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
    private const string OurName = "xlide";

    private readonly nint _window;
    private readonly string? _originalCaption;
    private readonly nint _previousSmall;
    private readonly nint _previousBig;

    private nint _small;
    private nint _big;
    private string? _applied;
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

        var caption = ReadCaption(_window);
        if (caption is not null && caption != _applied)
        {
            var ours = Rename(caption);
            if (ours != caption)
            {
                Win32.SetWindowText(_window, ours);
                _applied = ours;
            }
            else
            {
                _applied = caption;
            }
        }

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
    /// Swaps the host's product name for ours and leaves the rest of the caption alone, so the
    /// workbook and module the editor names are still named.
    /// </summary>
    private static string Rename(string caption) =>
        caption.StartsWith(HostName, StringComparison.Ordinal)
            ? string.Concat(OurName, caption.AsSpan(HostName.Length))
            : caption;

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
