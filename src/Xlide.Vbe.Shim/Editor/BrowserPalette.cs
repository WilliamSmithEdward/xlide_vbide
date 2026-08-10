using System.Runtime.InteropServices;
using System.Text.Json;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.WebView;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// The Object Browser as the developer chose it (2026-08-05): a real top-level window of our
/// own, dark, floating beside the editor - hosting a second browser surface that boots the
/// editor bundle into its browser page (?view=objbrowser). Nothing native is involved, so
/// there is nothing the editor can blank, close, or reclaim.
///
/// The palette answers the page's questions itself - libraries, types, members, navigation -
/// through delegates its owner wires to the typelib catalog and the engine. Closing hides
/// the window; a second summons presents it again with its state intact.
/// </summary>
internal sealed unsafe class BrowserPalette : IDisposable
{
    private const string ClassName = "XlidePalette";

    private static readonly object ClassGate = new();
    private static bool _classRegistered;

    private GCHandle _self;
    private nint _handle;
    private WebView2Surface? _browser;

#if DEBUG
    /// <summary>The palette's page, for the debug api's eval route. Debug builds only.</summary>
    internal WebView2Surface? Browser => _browser;
#endif

    /// <summary>The libraries to list, projects included.</summary>
    public Func<ObLibraryRow[]>? LibrariesRequested { get; set; }

    /// <summary>A library's types; null when the library is unknown.</summary>
    public Func<string, ObTypeRow[]?>? TypesRequested { get; set; }

    /// <summary>A type's members, answered through the reply - some answers wait on the engine.</summary>
    public Action<string, string, Action<ObMemberRow[]>>? MembersRequested { get; set; }

    /// <summary>The developer picked a project member; the editor goes to it.</summary>
    public Action<string, int, string?>? NavigateRequested { get; set; }

    private BrowserPalette()
    {
    }

    public nint Handle => _handle;

    /// <summary>Creates the palette over <paramref name="owner"/> and starts its page.</summary>
    public static BrowserPalette? Open(nint owner)
    {
        if (owner == 0 || !EnsureClassRegistered())
        {
            return null;
        }

        var palette = new BrowserPalette();
        palette._self = GCHandle.Alloc(palette);

        Rect ownerRect;
        Win32.GetWindowRect(owner, &ownerRect);
        var width = Math.Min(960, Math.Max(620, ownerRect.Right - ownerRect.Left - 200));
        var height = Math.Min(700, Math.Max(460, ownerRect.Bottom - ownerRect.Top - 180));
        var left = ownerRect.Left + ((ownerRect.Right - ownerRect.Left) - width) / 2;
        var top = ownerRect.Top + ((ownerRect.Bottom - ownerRect.Top) - height) / 2;

        var handle = Win32.CreateWindowEx(
            0,
            ClassName,
            "Object Browser",
            Win32.WsOverlappedWindow | Win32.WsClipChildren,
            left,
            top,
            width,
            height,
            owner,
            0,
            ShimModule.Handle,
            GCHandle.ToIntPtr(palette._self));

        if (handle == 0)
        {
            palette._self.Free();
            return null;
        }

        palette._handle = handle;

        Rect client;
        Win32.GetClientRect(handle, &client);
        palette._browser = WebView2Surface.Start(
            handle,
            new PixelRect(0, 0, client.Right - client.Left, client.Bottom - client.Top),
            "?view=objbrowser");

        if (palette._browser is null)
        {
            Log.Error("palette: the browser surface could not be started");
            palette.Dispose();
            return null;
        }

#if DEBUG
        palette._browser.DebugName = "palette";
#endif

        palette._browser.MessageReceived = palette.OnMessage;
        AdoptOwnerIcon(handle, owner);
        Win32.ShowWindow(handle, Win32.SwShow);
        Win32.SetForegroundWindow(handle);
        Log.Info("palette: the Object Browser window is open");
        return palette;
    }

    /// <summary>
    /// Stamps the editor's own icon onto the window, so an xlide window reads as part of
    /// the editor (developer, 2026-08-05). The running editor is the source: its window is
    /// asked via WM_GETICON, its class is the fallback, and the handles are shared rather
    /// than copied - the palette lives strictly inside the editor's lifetime, so borrowing
    /// is safe, and no icon resource is taken from anyone's files.
    /// </summary>
    private static void AdoptOwnerIcon(nint window, nint owner)
    {
        var small = Win32.SendMessage(owner, Win32.WmGetIcon, Win32.IconSmall2, 0);
        if (small == 0)
        {
            small = Win32.SendMessage(owner, Win32.WmGetIcon, Win32.IconSmall, 0);
        }

        if (small == 0)
        {
            small = Win32.GetClassLongPtr(owner, Win32.GclpHIconSm);
        }

        var big = Win32.SendMessage(owner, Win32.WmGetIcon, Win32.IconBig, 0);
        if (big == 0)
        {
            big = Win32.GetClassLongPtr(owner, Win32.GclpHIcon);
        }

        if (small != 0)
        {
            Win32.SendMessage(window, Win32.WmSetIcon, Win32.IconSmall, small);
        }

        if (big != 0)
        {
            Win32.SendMessage(window, Win32.WmSetIcon, Win32.IconBig, big);
        }
    }

    /// <summary>
    /// Hides the palette. An owned window does not follow its owner out of sight, so the
    /// editor closing must say this explicitly - otherwise the palette outlives the editor
    /// window and is standing there, ownerless-looking, when the editor next opens
    /// ("this pops up when I first open xlide", 2026-08-05). Pure user32, so it is safe on
    /// the editor's close path (lesson 27).
    /// </summary>
    public void Hide()
    {
        if (_handle != 0)
        {
            Win32.ShowWindow(_handle, Win32.SwHide);
        }
    }

    /// <summary>Brings the palette back, shown, forward, and focused for its search box.</summary>
    public void Present()
    {
        if (_handle != 0)
        {
            Win32.ShowWindow(_handle, Win32.SwShow);
            Win32.SetForegroundWindow(_handle);
            _browser?.Focus();
        }
    }

    private void OnMessage(string json)
    {
        // The browser is created invisible and shown on the page's word (the editor's loader
        // discipline, lesson 27). This window has no loader, so the page's FIRST message is
        // its word: revealing any earlier is a call on a controller that does not exist yet -
        // the white nothing of 09:18 - and the browser's idle ground is the theme's, so there
        // is no flash to hide.
        _browser?.Reveal();

        try
        {
            using var document = JsonDocument.Parse(json);
            if (!document.RootElement.TryGetProperty("type", out var typeElement)
                || typeElement.GetString() is not { Length: > 0 } type)
            {
                return;
            }

            switch (type)
            {
                case "obLibraries":
                    if (TryId(document, out var librariesId))
                    {
                        var rows = LibrariesRequested?.Invoke() ?? [];
                        Post(JsonSerializer.Serialize(
                            new ObLibrariesResultMessage("obLibrariesResult", librariesId, rows),
                            EditorMessageContext.Default.ObLibrariesResultMessage));
                    }

                    break;

                case "obTypes":
                    if (TryId(document, out var typesId)
                        && document.RootElement.TryGetProperty("library", out var typesLibrary)
                        && typesLibrary.GetString() is { Length: > 0 } typesLibraryName)
                    {
                        var rows = TypesRequested?.Invoke(typesLibraryName) ?? [];
                        Post(JsonSerializer.Serialize(
                            new ObTypesResultMessage("obTypesResult", typesId, rows),
                            EditorMessageContext.Default.ObTypesResultMessage));
                    }

                    break;

                case "obMembers":
                    if (TryId(document, out var membersId)
                        && document.RootElement.TryGetProperty("library", out var membersLibrary)
                        && membersLibrary.GetString() is { Length: > 0 } membersLibraryName
                        && document.RootElement.TryGetProperty("typeName", out var membersType)
                        && membersType.GetString() is { Length: > 0 } membersTypeName)
                    {
                        if (MembersRequested is { } ask)
                        {
                            ask(membersLibraryName, membersTypeName, rows =>
                                Post(JsonSerializer.Serialize(
                                    new ObMembersResultMessage("obMembersResult", membersId, rows),
                                    EditorMessageContext.Default.ObMembersResultMessage)));
                        }
                    }

                    break;

                case "navigate":
                    if (document.RootElement.TryGetProperty("module", out var moduleElement)
                        && moduleElement.GetString() is { Length: > 0 } module
                        && document.RootElement.TryGetProperty("line", out var lineElement)
                        && lineElement.TryGetInt32(out var line))
                    {
                        var project = document.RootElement.TryGetProperty("project", out var projectElement)
                            ? projectElement.GetString()
                            : null;
                        NavigateRequested?.Invoke(module, line, project);
                    }

                    break;

                case "close":
                    // Escape from inside the page: the same hide the close box performs.
                    if (_handle != 0)
                    {
                        Win32.ShowWindow(_handle, Win32.SwHide);
                    }

                    break;
            }
        }
        catch (Exception ex)
        {
            Log.Error("palette: a page message could not be handled", ex);
        }
    }

    private static bool TryId(JsonDocument document, out int id)
    {
        id = 0;
        return document.RootElement.TryGetProperty("id", out var element) && element.TryGetInt32(out id);
    }

    private void Post(string json) => _browser?.PostMessage(json);

    private void FitBrowser()
    {
        if (_handle == 0 || _browser is null)
        {
            return;
        }

        Rect client;
        if (Win32.GetClientRect(_handle, &client))
        {
            _browser.SetBounds(new PixelRect(0, 0, client.Right - client.Left, client.Bottom - client.Top));
        }
    }

    private static BrowserPalette? FromHandle(nint window)
    {
        var data = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        return data != 0 ? GCHandle.FromIntPtr(data).Target as BrowserPalette : null;
    }

    private static bool EnsureClassRegistered()
    {
        lock (ClassGate)
        {
            if (_classRegistered)
            {
                return true;
            }

            fixed (char* className = ClassName)
            {
                var windowClass = new WndClassExW
                {
                    Size = (uint)sizeof(WndClassExW),
                    Style = 0,
                    WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&WindowProc,
                    Instance = ShimModule.Handle,
                    Cursor = Win32.LoadCursor(0, Win32.IdcArrow),
                    Background = 0,
                    ClassName = className,
                };

                if (Win32.RegisterClassEx(&windowClass) != 0
                    || Marshal.GetLastWin32Error() == Win32.ErrorClassAlreadyExists)
                {
                    _classRegistered = true;
                    return true;
                }

                Log.Error($"palette: the window class could not be registered, error {Marshal.GetLastWin32Error()}");
                return false;
            }
        }
    }

    [UnmanagedCallersOnly]
    private static nint WindowProc(nint window, uint message, nint wParam, nint lParam)
    {
        try
        {
            switch (message)
            {
                case Win32.WmNcCreate:
                {
                    var create = (CreateStructW*)lParam;
                    Win32.SetWindowLongPtr(window, Win32.GwlpUserData, create->CreateParams);
                    break;
                }

                case Win32.WmSize:
                    FromHandle(window)?.FitBrowser();
                    return 0;

                case Win32.WmClose:
                    // Ours alone - nothing native sends this window anything. Closing hides,
                    // so the next summons presents the same page, state intact.
                    Win32.ShowWindow(window, Win32.SwHide);
                    return 0;
            }
        }
        catch (Exception ex)
        {
            Log.Error("palette: the window procedure failed", ex);
        }

        return Win32.DefWindowProc(window, message, wParam, lParam);
    }

    public void Dispose()
    {
        _browser?.Dispose();
        _browser = null;

        var handle = _handle;
        _handle = 0;

        if (handle != 0)
        {
            Win32.DestroyWindow(handle);
        }

        if (_self.IsAllocated)
        {
            _self.Free();
        }
    }
}
