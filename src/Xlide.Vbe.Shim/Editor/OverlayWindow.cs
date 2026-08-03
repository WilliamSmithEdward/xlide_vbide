using System.Runtime.InteropServices;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// A plain window that sits over a code pane and hosts the editing surface.
///
/// It is a child of the editor's own frame rather than of the pane. A child of the pane would be
/// clipped, moved, and destroyed by code that has no idea it exists, and the pane is recreated more
/// often than it looks. A sibling positioned over it survives all of that and is ours to place.
///
/// While the browser is coming up, this window paints the loader: the wordmark and a pulse on the
/// editor's dark ground, because the alternative is the compositor's blank rectangle. From the
/// moment the page reports ready, nothing is painted here again; every pixel comes from the
/// surface.
/// </summary>
internal sealed unsafe class OverlayWindow : IDisposable
{
    private const string ClassName = "XlideEditorOverlay";

    /// <summary>Timer identifiers, scoped to this window.</summary>
    private const nuint WriteTimerId = 1;
    private const nuint PollTimerId = 2;
    private const nuint ActionTimerId = 3;
    private const nuint LoaderTimerId = 4;
    private const nuint AnalyseTimerId = 5;

    /// <summary>One pulse step of the loader; three steps make its cycle.</summary>
    private const uint LoaderTickMilliseconds = 240;

    /// <summary>
    /// Ticks before the loader admits something is wrong and says where to look. Covers every
    /// way the browser can fail to arrive with one message instead of a spinner that never ends.
    /// </summary>
    private const int LoaderStalledAfterTicks = 75;

    // The loader's palette, as COLORREF values. They are the page's own dark theme, so the
    // hand-off from loader to surface reads as one screen coming into focus.
    private const uint LoaderBackground = 0x001E1E1E;
    private const uint LoaderForeground = 0x00CCCCCC;
    private const uint LoaderDimmed = 0x00404040;
    private const uint LoaderHint = 0x00808080;

    private static bool _classRegistered;
    private static readonly Lock ClassGate = new();

    private nint _handle;
    private GCHandle _self;

    /// <summary>
    /// Where the overlay was last put, so an unchanged position costs nothing.
    ///
    /// Window events arrive for anything that moves anywhere in the editor, and most of them leave
    /// the pane exactly where it was. Repositioning anyway would resize the browser and force a
    /// relayout of the document on every one of them.
    /// </summary>
    private PixelRect _placed;
    private bool _shown;

    /// <summary>True from creation until the surface reports ready, while the loader paints.</summary>
    private bool _loading = true;

    /// <summary>Which of the loader's dots is lit, advanced by its timer.</summary>
    private int _loaderPhase;

    /// <summary>
    /// True once the loader has been showing implausibly long. Placement consults this: a
    /// stalled loader retreats below the native menu bar, so a page that never arrives cannot
    /// keep every menu covered.
    /// </summary>
    public bool LoaderStalled => _loading && _loaderPhase >= LoaderStalledAfterTicks;

    private OverlayWindow()
    {
    }

    /// <summary>
    /// Retires the loader: the browser is about to be seen, so this window goes back to painting
    /// nothing. Idempotent, and only meaningful on the host thread, which owns the window.
    /// </summary>
    public void HideLoader()
    {
        if (_handle == 0 || !_loading)
        {
            return;
        }

        _loading = false;
        Win32.KillTimer(_handle, LoaderTimerId);
        Win32.InvalidateRect(_handle, null, false);
    }

    public nint Handle => _handle;

    /// <summary>Client area, in the coordinates the browser surface expects.</summary>
    public PixelRect ClientBounds()
    {
        Rect rect;
        if (_handle == 0 || !Win32.GetClientRect(_handle, &rect))
        {
            return default;
        }

        return new PixelRect(0, 0, rect.Right - rect.Left, rect.Bottom - rect.Top);
    }

    /// <summary>Raised when the window has been resized, so the surface can follow.</summary>
    public Action<PixelRect>? Resized { get; set; }

    /// <summary>Raised once, on the host thread, after <see cref="StartWriteTimer"/> elapses.</summary>
    public Action? Elapsed { get; set; }

    /// <summary>Raised repeatedly, on the host thread, while the poll timer runs.</summary>
    public Action? Polled { get; set; }

    /// <summary>
    /// Raised on the host thread at each loader animation step, while the loader is showing.
    ///
    /// The loading phase is event-starved: no pane exists yet, so nothing else re-places the
    /// surface, while the editor is still arranging itself underneath — restoring its size,
    /// raising its own bands. Whoever owns placement listens here and re-asserts it, which is
    /// what keeps the loader covering the window it was placed over rather than the window as
    /// it was a moment ago.
    /// </summary>
    public Action? LoaderTicked { get; set; }

    /// <summary>
    /// Starts or restarts a one-shot timer. Restarting is what makes it a debounce: each call
    /// pushes the deadline out, so a burst of keystrokes produces one callback rather than one per
    /// key.
    /// </summary>
    public void StartWriteTimer(uint milliseconds)
    {
        if (_handle != 0)
        {
            Win32.SetTimer(_handle, WriteTimerId, milliseconds, 0);
        }
    }

    /// <summary>Cancels the one-shot timer if it is running.</summary>
    public void StopWriteTimer()
    {
        if (_handle != 0)
        {
            Win32.KillTimer(_handle, WriteTimerId);
        }
    }

    /// <summary>Raised once, on the host thread, after <see cref="StartAnalyseTimer"/> elapses.</summary>
    public Action? AnalysisDue { get; set; }

    /// <summary>
    /// Starts or restarts the live-analysis debounce: each keystroke pushes the deadline out,
    /// and the pass runs once per pause rather than once per key.
    /// </summary>
    public void StartAnalyseTimer(uint milliseconds)
    {
        if (_handle != 0)
        {
            Win32.SetTimer(_handle, AnalyseTimerId, milliseconds, 0);
        }
    }

    /// <summary>Starts a repeating timer, or changes its interval.</summary>
    public void StartPollTimer(uint milliseconds)
    {
        if (_handle != 0)
        {
            Win32.SetTimer(_handle, PollTimerId, milliseconds, 0);
        }
    }

    /// <summary>Stops the repeating timer.</summary>
    public void StopPollTimer()
    {
        if (_handle != 0)
        {
            Win32.KillTimer(_handle, PollTimerId);
        }
    }

    /// <summary>Actions waiting to run on the host thread, drained by the action timer.</summary>
    private readonly System.Collections.Concurrent.ConcurrentQueue<Action> _actions = new();

    /// <summary>
    /// Runs an action on the thread that owns this window, which is the only thread the editor's
    /// object model and the browser may be touched from. An engine answer arrives on a pool
    /// thread, and this is its way back.
    ///
    /// The hop rides a window timer, not a posted message. The host's message loop swallows
    /// app-range messages posted to windows it does not manage — across every recorded session,
    /// not one posted action was ever dispatched — while WM_TIMER demonstrably always arrives:
    /// the write debounce and the poll run on it in those same sessions.
    /// </summary>
    public void RunOnHostThread(Action action)
    {
        ArgumentNullException.ThrowIfNull(action);

        if (_handle == 0)
        {
            // Dropped, and said so: a marshal that vanishes silently once cost two days of
            // completions that never arrived.
            Log.Info("overlay: an action for the host thread was dropped, the window is gone");
            return;
        }

        _actions.Enqueue(action);
        Win32.SetTimer(_handle, ActionTimerId, 0, 0);
    }

    /// <summary>Runs everything queued for the host thread. Only the window's thread calls this.</summary>
    private void DrainActions()
    {
        while (_actions.TryDequeue(out var action))
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                // One failing action must not take the rest of the queue with it.
                Log.Error("overlay: a marshalled action failed", ex);
            }
        }
    }

    /// <summary>Creates the overlay as a child of <paramref name="parent"/>.</summary>
    public static OverlayWindow? Create(nint parent, PixelRect bounds)
    {
        if (parent == 0 || !EnsureClassRegistered())
        {
            return null;
        }

        var overlay = new OverlayWindow();
        overlay._self = GCHandle.Alloc(overlay);

        var handle = Win32.CreateWindowEx(
            0,
            ClassName,
            null,
            Win32.WsChild | Win32.WsClipChildren | Win32.WsClipSiblings,
            bounds.Left,
            bounds.Top,
            bounds.Width,
            bounds.Height,
            parent,
            0,
            ShimModule.Handle,
            GCHandle.ToIntPtr(overlay._self));

        if (handle == 0)
        {
            Log.Error($"overlay: could not be created, error {Marshal.GetLastWin32Error()}");
            overlay._self.Free();
            return null;
        }

        overlay._handle = handle;

        // The loader animates from the first frame; the browser retires it when its page is up.
        Win32.SetTimer(handle, LoaderTimerId, LoaderTickMilliseconds, 0);
        return overlay;
    }

    /// <summary>Moves and sizes the overlay, and shows or hides it.</summary>
    public void Place(PixelRect bounds, bool visible)
    {
        if (_handle == 0)
        {
            return;
        }

        if (!visible)
        {
            if (_shown)
            {
                _shown = false;
                Win32.ShowWindow(_handle, Win32.SwHide);
            }

            return;
        }

        // Raising and moving are separated deliberately.
        //
        // The overlay has to be raised on every call, whether or not it has moved. Its siblings are
        // the code panes, and the editor puts a pane on top of its siblings whenever it activates
        // one, which happens every time the user picks a different module. Skipping the raise
        // because the rectangle had not changed let the newly activated pane cover the surface, and
        // the surface looked like it had vanished.
        //
        // Moving is the expensive half: it resizes the browser and relayouts the document. Window
        // events arrive for anything that moves anywhere in the editor and most leave the pane
        // exactly where it was, so that half is skipped when nothing changed.
        var moved = !_shown || !bounds.Equals(_placed);

        var flags = Win32.SwpNoActivate | (moved ? 0 : Win32.SwpNoMove | Win32.SwpNoSize);

        Win32.SetWindowPos(
            _handle,
            Win32.HwndTop,
            bounds.Left,
            bounds.Top,
            bounds.Width,
            bounds.Height,
            flags);

        if (moved)
        {
            _placed = bounds;
            Log.Info($"overlay: placed at {bounds.Left},{bounds.Top} {bounds.Width}x{bounds.Height} in its parent");
        }

        if (!_shown)
        {
            _shown = true;
            Win32.ShowWindow(_handle, Win32.SwShowNoActivate);
        }
    }

    /// <summary>The cutouts last applied, in the parent's client space, and the placement they were computed against.</summary>
    private PixelRect[]? _cutouts;
    private PixelRect _cutoutsPlacedAt;

    /// <summary>
    /// Punches holes in the overlay where native windows must show through.
    ///
    /// The surface never retreats to make room for a native tool window; it covers the whole
    /// frame and cuts a hole exactly where each one sits. The hole is real absence, not
    /// transparency: painting, hit-testing, and the browser's own child windows all stop at the
    /// region boundary, so the native window underneath is fully live inside it.
    ///
    /// Rectangles arrive in the parent's client space, the same space <see cref="Place"/> uses;
    /// the region wants them window-relative, so the current placement is subtracted. An empty
    /// set restores the whole rectangle.
    /// </summary>
    public void SetCutouts(ReadOnlySpan<PixelRect> holes)
    {
        if (_handle == 0)
        {
            return;
        }

        // A region is sized to the window; the same holes against a different placement are a
        // different region. Skipping only truly identical states matters because rebuilding the
        // region forces a repaint.
        if (_cutouts is not null && _cutoutsPlacedAt.Equals(_placed) && holes.SequenceEqual(_cutouts))
        {
            return;
        }

        _cutouts = holes.ToArray();
        _cutoutsPlacedAt = _placed;

        if (holes.IsEmpty)
        {
            _ = Win32.SetWindowRgn(_handle, 0, true);
            return;
        }

        var region = Win32.CreateRectRgn(0, 0, _placed.Width, _placed.Height);
        if (region == 0)
        {
            return;
        }

        foreach (var hole in holes)
        {
            var cut = Win32.CreateRectRgn(
                hole.Left - _placed.Left,
                hole.Top - _placed.Top,
                hole.Right - _placed.Left,
                hole.Bottom - _placed.Top);

            if (cut == 0)
            {
                continue;
            }

            _ = Win32.CombineRgn(region, region, cut, Win32.RgnDiff);
            Win32.DeleteObject(cut);
        }

        if (Win32.SetWindowRgn(_handle, region, true) == 0)
        {
            // The window only owns the region on success.
            Win32.DeleteObject(region);
        }
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
                    Style = Win32.CsHRedraw | Win32.CsVRedraw,
                    WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&WindowProc,
                    Instance = ShimModule.Handle,
                    Cursor = Win32.LoadCursor(0, Win32.IdcArrow),
                    Background = 0,
                    ClassName = className,
                };

                if (Win32.RegisterClassEx(&windowClass) != 0)
                {
                    _classRegistered = true;
                    return true;
                }

                // Registering twice in one process is not a failure; the class is already there.
                var error = Marshal.GetLastWin32Error();
                if (error == Win32.ErrorClassAlreadyExists)
                {
                    _classRegistered = true;
                    return true;
                }

                Log.Error($"overlay: the window class could not be registered, error {error}");
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
                    // The only chance to associate the window with its managed object before any
                    // other message arrives.
                    var create = (CreateStructW*)lParam;
                    Win32.SetWindowLongPtr(window, Win32.GwlpUserData, create->CreateParams);
                    break;
                }

                case Win32.WmSize:
                {
                    var overlay = FromHandle(window);
                    if (overlay is not null)
                    {
                        overlay.Resized?.Invoke(overlay.ClientBounds());
                    }

                    return 0;
                }

                case Win32.WmTimer:
                {
                    var overlay = FromHandle(window);
                    if (overlay is null)
                    {
                        return 0;
                    }

                    if ((nuint)wParam == WriteTimerId)
                    {
                        // One shot. A window timer repeats until it is killed, and the work behind
                        // this only needs doing once per burst.
                        overlay.StopWriteTimer();
                        overlay.Elapsed?.Invoke();
                    }
                    else if ((nuint)wParam == PollTimerId)
                    {
                        overlay.Polled?.Invoke();
                    }
                    else if ((nuint)wParam == ActionTimerId)
                    {
                        // One shot per burst of queued actions.
                        Win32.KillTimer(window, ActionTimerId);
                        overlay.DrainActions();
                    }
                    else if ((nuint)wParam == LoaderTimerId)
                    {
                        overlay._loaderPhase++;
                        overlay.LoaderTicked?.Invoke();
                        Win32.InvalidateRect(window, null, false);
                    }
                    else if ((nuint)wParam == AnalyseTimerId)
                    {
                        // One shot per pause in the typing.
                        Win32.KillTimer(window, AnalyseTimerId);
                        overlay.AnalysisDue?.Invoke();
                    }

                    return 0;
                }

                case Win32.WmPaint:
                {
                    var overlay = FromHandle(window);
                    if (overlay is not null && overlay._loading)
                    {
                        overlay.PaintLoader(window);
                        return 0;
                    }

                    // The browser owns the pixels; the default handling just validates.
                    break;
                }

                case Win32.WmEraseBackground:
                    // The browser paints every pixel, and so does the loader. Erasing first
                    // would flash.
                    return 1;

                case Win32.WmNcDestroy:
                {
                    var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
                    if (stored != 0)
                    {
                        var handle = GCHandle.FromIntPtr(stored);
                        if (handle.IsAllocated)
                        {
                            if (handle.Target is OverlayWindow overlay)
                            {
                                overlay._handle = 0;
                            }

                            handle.Free();
                        }

                        Win32.SetWindowLongPtr(window, Win32.GwlpUserData, 0);
                    }

                    break;
                }
            }
        }
        catch (Exception)
        {
            // Nothing may escape into the window procedure.
        }

        return Win32.DefWindowProc(window, message, wParam, lParam);
    }

    private static OverlayWindow? FromHandle(nint window)
    {
        var stored = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        if (stored == 0)
        {
            return null;
        }

        var handle = GCHandle.FromIntPtr(stored);
        return handle.IsAllocated ? handle.Target as OverlayWindow : null;
    }

    /// <summary>
    /// The loader frame: the wordmark over the editor's dark ground, three dots pulsing beneath
    /// it, and past the stall threshold a line saying where the log is. Plain GDI, because this
    /// paints before anything richer exists — being before everything else is its whole point.
    /// </summary>
    private void PaintLoader(nint window)
    {
        PaintStruct paint;
        var dc = Win32.BeginPaint(window, &paint);
        if (dc == 0)
        {
            return;
        }

        try
        {
            Rect client;
            if (!Win32.GetClientRect(window, &client))
            {
                return;
            }

            var background = Win32.CreateSolidBrush(LoaderBackground);
            if (background != 0)
            {
                _ = Win32.FillRect(dc, &client, background);
                Win32.DeleteObject(background);
            }

            _ = Win32.SetBkMode(dc, Win32.BackgroundTransparent);

            var centerX = (client.Left + client.Right) / 2;
            var centerY = (client.Top + client.Bottom) / 2;

            // The wordmark, a little above centre so the dots sit at the optical middle.
            var wordmarkFont = Win32.CreateFont(
                -30, 0, 0, 0,
                Win32.FontWeightSemibold,
                0, 0, 0,
                Win32.FontDefaultCharset,
                0, 0,
                Win32.FontClearTypeQuality,
                0,
                "Segoe UI");

            if (wordmarkFont != 0)
            {
                var previousFont = Win32.SelectObject(dc, wordmarkFont);
                _ = Win32.SetTextColor(dc, LoaderForeground);

                var wordmarkRect = client;
                wordmarkRect.Bottom = centerY;
                Win32.DrawText(
                    dc,
                    "xlide",
                    -1,
                    &wordmarkRect,
                    Win32.DtCenter | Win32.DtVCenter | Win32.DtSingleLine | Win32.DtNoPrefix);

                Win32.SelectObject(dc, previousFont);
                Win32.DeleteObject(wordmarkFont);
            }

            // The pulse: three dots, one lit, walking left to right.
            var previousPen = Win32.SelectObject(dc, Win32.GetStockObject(Win32.NullPen));
            for (var dot = 0; dot < 3; dot++)
            {
                var lit = _loaderPhase % 3 == dot;
                var brush = Win32.CreateSolidBrush(lit ? LoaderForeground : LoaderDimmed);
                if (brush == 0)
                {
                    continue;
                }

                var previousBrush = Win32.SelectObject(dc, brush);
                var dotX = centerX + (dot - 1) * 18;
                var dotY = centerY + 18;
                Win32.Ellipse(dc, dotX - 4, dotY - 4, dotX + 5, dotY + 5);
                Win32.SelectObject(dc, previousBrush);
                Win32.DeleteObject(brush);
            }

            Win32.SelectObject(dc, previousPen);

            if (_loaderPhase >= LoaderStalledAfterTicks)
            {
                var hintFont = Win32.CreateFont(
                    -13, 0, 0, 0,
                    0, 0, 0, 0,
                    Win32.FontDefaultCharset,
                    0, 0,
                    Win32.FontClearTypeQuality,
                    0,
                    "Segoe UI");

                if (hintFont != 0)
                {
                    var previousFont = Win32.SelectObject(dc, hintFont);
                    _ = Win32.SetTextColor(dc, LoaderHint);

                    var hintRect = client;
                    hintRect.Top = centerY + 40;
                    hintRect.Bottom = centerY + 70;
                    Win32.DrawText(
                        dc,
                        "still starting — the xlide log has the story",
                        -1,
                        &hintRect,
                        Win32.DtCenter | Win32.DtSingleLine | Win32.DtNoPrefix);

                    Win32.SelectObject(dc, previousFont);
                    Win32.DeleteObject(hintFont);
                }
            }
        }
        finally
        {
            Win32.EndPaint(window, &paint);
        }
    }

    public void Dispose()
    {
        var handle = _handle;

        if (handle != 0)
        {
            Win32.KillTimer(handle, WriteTimerId);
            Win32.KillTimer(handle, PollTimerId);
            Win32.KillTimer(handle, ActionTimerId);
            Win32.KillTimer(handle, LoaderTimerId);
            Win32.KillTimer(handle, AnalyseTimerId);
        }

        // Anything still queued was for a surface that no longer exists.
        while (_actions.TryDequeue(out _))
        {
        }

        _handle = 0;
        Elapsed = null;
        Polled = null;
        LoaderTicked = null;

        if (handle != 0 && Win32.IsWindow(handle))
        {
            // Destroying the window frees the handle to this object from its own message handler.
            Win32.DestroyWindow(handle);
            return;
        }

        if (_self.IsAllocated)
        {
            _self.Free();
        }
    }
}
