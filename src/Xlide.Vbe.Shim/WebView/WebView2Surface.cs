using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.WebView;

/// <summary>
/// One browser surface hosted in one window.
///
/// Creation is asynchronous in two stages and neither stage can be waited on: the environment
/// arrives on a callback, and the controller arrives on a second callback raised from the first.
/// Both callbacks come back on the thread that started the work, which is the host user interface
/// thread, delivered through the message loop Excel is already running. Nothing here may block
/// waiting for them, because the loop that delivers them is the loop that would be blocked.
///
/// Everything on this type runs on that one thread, including the message callbacks. There is no
/// synchronisation here and none is wanted: an apartment-bound object protected by a lock hides
/// the fact that a call from another thread was already wrong before the lock was reached.
/// </summary>
internal sealed class WebView2Surface : IDisposable
{
    private static bool _loaderLoaded;

    private readonly nint _parentWindow;

    private ComHandle<ICoreWebView2Environment>? _environment;
    private ComHandle<ICoreWebView2Controller>? _controller;
    private ComHandle<ICoreWebView2>? _webView;

    /// <summary>
    /// The same content surface asked for a later revision of its interface, or null on a runtime
    /// that predates it. Everything the shell document does not need goes through this: the folder
    /// mapping and the message channel both arrived after the first revision.
    /// </summary>
    private ComHandle<ICoreWebView2_3>? _view;

    private EnvironmentCompletedHandler? _environmentHandler;
    private ControllerCompletedHandler? _controllerHandler;
    private NavigationCompletedHandler? _navigationHandler;
    private WebMessageReceivedHandler? _messageHandler;
    private AcceleratorKeyPressedHandler? _acceleratorHandler;

    /// <summary>Serves the bundle over loopback; the folder mapping is the fallback behind it.</summary>
    private LoopbackPageServer? _pageServer;
    private EventRegistrationToken _navigationCompletedToken;
    private EventRegistrationToken _webMessageReceivedToken;
    private EventRegistrationToken _acceleratorKeyPressedToken;

    private PixelRect _bounds;
    private bool _disposed;
    private bool _reportedUnhandledMessage;

    private WebView2Surface(nint parentWindow, PixelRect bounds)
    {
        _parentWindow = parentWindow;
        _bounds = bounds;
    }

    /// <summary>
    /// Begins creating a browser inside <paramref name="parentWindow"/>. Returns as soon as the
    /// work is started; nothing is rendered until the two callbacks have run.
    /// </summary>
    public static WebView2Surface? Start(nint parentWindow, PixelRect bounds)
    {
        if (parentWindow == 0)
        {
            return null;
        }

        var surface = new WebView2Surface(parentWindow, bounds);
        return surface.BeginEnvironment() ? surface : null;
    }

    /// <summary>
    /// Raised with the text of each message the page posts through
    /// window.chrome.webview.postMessage. Invoked on the host user interface thread, inside the
    /// browser's own callback, so the handler does only enough work to hand the message on.
    /// </summary>
    public Action<string>? MessageReceived { get; set; }

    /// <summary>
    /// Asked about each accelerator key before the page sees it. Return true to claim the key, which
    /// stops the document acting on it as well.
    /// </summary>
    public Func<uint, bool>? AcceleratorPressed { get; set; }

    /// <summary>
    /// Sizes the browser to the window's client area. Idempotent: an unchanged size is not
    /// re-sent, so callers may assert the bounds freely — placement does, on every pass, which
    /// is what keeps a resize message that never arrived from leaving the page laid out for a
    /// width the window no longer has (the clipped minimap of 2026-08-05).
    /// </summary>
    public void SetBounds(PixelRect bounds)
    {
        _bounds = bounds;

        var controller = _controller;
        if (controller is null || _appliedBounds == bounds)
        {
            return;
        }

        var rect = new Rect
        {
            Left = bounds.Left,
            Top = bounds.Top,
            Right = bounds.Right,
            Bottom = bounds.Bottom,
        };

        controller.Target.PutBounds(rect);
        _appliedBounds = bounds;
        Log.Verbose($"webview: bounds {bounds.Width}x{bounds.Height}");
    }

    /// <summary>The bounds the controller last accepted; null until it exists to accept any.</summary>
    private PixelRect? _appliedBounds;

    /// <summary>Gives keyboard focus to the browser.</summary>
    public void Focus() => _controller?.Target.MoveFocus(MoveFocusReason.Programmatic);

    /// <summary>
    /// Serves <paramref name="folderPath"/> under <paramref name="hostName"/>, so its contents are
    /// reachable at https://hostName/... inside this browser and nowhere else.
    ///
    /// The access kind is DENY_CORS rather than ALLOW. Both serve the folder to the document loaded
    /// from the mapped host; ALLOW additionally serves it to requests from any other origin the
    /// browser ever loads, which would make a directory inside the install readable by an arbitrary
    /// page. Nothing the editor surface does needs that.
    ///
    /// The mapping is a property of the content surface and survives navigation, so it is set once
    /// rather than before each navigation.
    /// </summary>
    /// <returns>False when the runtime is too old to map folders, or when the browser refused.</returns>
    public bool SetContentRoot(string hostName, string folderPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hostName);
        ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);

        var view = _view;
        if (view is null)
        {
            Log.Warn($"webview: cannot map {hostName}, this runtime has no folder mapping");
            return false;
        }

        var hr = view.Target.SetVirtualHostNameToFolderMapping(hostName, folderPath, HostResourceAccessKind.DenyCors);
        if (hr < 0)
        {
            Log.Error($"webview: SetVirtualHostNameToFolderMapping({hostName}) returned 0x{hr:X8}");
            return false;
        }

        Log.Info($"webview: {hostName} maps to {folderPath}");
        return true;
    }

    /// <summary>
    /// Navigates the surface to an address.
    ///
    /// Goes through the first revision of the interface, which every runtime has. Only the mapping
    /// that makes a virtual host name resolve needs the later one, so a caller that has already
    /// mapped a root can navigate to it here without a second capability check.
    /// </summary>
    /// <returns>False when the browser is not up yet, or when it refused the address.</returns>
    public bool NavigateToUrl(string url)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(url);

        var view = _webView;
        if (view is null)
        {
            Log.Warn($"webview: cannot navigate to {url}, no content surface");
            return false;
        }

        var hr = view.Target.Navigate(url);
        if (hr < 0)
        {
            Log.Error($"webview: Navigate({url}) returned 0x{hr:X8}");
            return false;
        }

        Log.Info($"webview: navigating to {url}");
        return true;
    }

    /// <summary>
    /// Posts a message to the page, where it arrives as the data of a message event on
    /// window.chrome.webview.
    ///
    /// Sent as a string, not as JSON, which is what makes the channel symmetric: the page posts
    /// text and receives text, and the JSON lives inside that text in both directions. Posting as
    /// JSON would deliver a parsed object one way and a string the other, and would also fail the
    /// call outright on malformed input rather than letting the far side reject it.
    ///
    /// Nothing is queued. A message posted before the page exists is discarded by the browser, so
    /// callers post in response to something the page did.
    /// </summary>
    /// <returns>False when the browser is not up yet, or when it refused the message.</returns>
    public bool PostMessage(string json)
    {
        ArgumentNullException.ThrowIfNull(json);

        var view = _view;
        if (view is null)
        {
            Log.Warn("webview: a message was posted before the content surface existed, discarding it");
            return false;
        }

        var hr = view.Target.PostWebMessageAsString(json);
        if (hr < 0)
        {
            Log.Error($"webview: PostWebMessageAsString returned 0x{hr:X8}{KnownHresult(hr)}");
            return false;
        }

        return true;
    }

    /// <summary>
    /// Names the failure codes this integration has actually met, so a support log reads as a
    /// diagnosis rather than a puzzle. An unknown code stays a bare number honestly.
    /// </summary>
    private static string KnownHresult(int hr) => hr switch
    {
        unchecked((int)0x802A000C) => " (UI_E_WRONG_THREAD: called off the browser's own thread)",
        unchecked((int)0x8001010E) => " (RPC_E_WRONG_THREAD)",
        unchecked((int)0x80070005) => " (E_ACCESSDENIED)",
        unchecked((int)0x8007139F) => " (ERROR_INVALID_STATE: the browser is gone or not up yet)",
        _ => string.Empty,
    };

    private bool BeginEnvironment()
    {
        if (!EnsureLoaderLoaded())
        {
            return false;
        }

        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var userDataFolder = WebViewPaths.UserDataFolder(localAppData);

        try
        {
            // The browser writes a profile, a cache, and lock files here on first use. The install
            // directory would be the obvious place and is the wrong one: it is not writable by the
            // user the host runs as.
            System.IO.Directory.CreateDirectory(userDataFolder);
        }
        catch (Exception ex)
        {
            Log.Error($"webview: could not create the user data folder at {userDataFolder}", ex);
            return false;
        }

        _environmentHandler = new EnvironmentCompletedHandler(OnEnvironmentCreated);
        var callback = CreateCallback(_environmentHandler, WebViewIid.EnvironmentCompletedHandler);
        if (callback == 0)
        {
            Log.Error("webview: could not expose the environment callback to the browser");
            return false;
        }

        try
        {
            Log.Info($"webview: creating environment, user data folder {userDataFolder}");
            var hr = WebView2Loader.CreateCoreWebView2EnvironmentWithOptions(null, userDataFolder, 0, callback);
            if (hr < 0)
            {
                Log.Error($"webview: CreateCoreWebView2EnvironmentWithOptions returned 0x{hr:X8}");
                return false;
            }

            return true;
        }
        catch (DllNotFoundException ex)
        {
            Log.Error("webview: the loader library could not be bound", ex);
            return false;
        }
        catch (EntryPointNotFoundException ex)
        {
            Log.Error("webview: the loader library has no environment entry point", ex);
            return false;
        }
        finally
        {
            // The browser took its own reference for the duration of the asynchronous call.
            Marshal.Release(callback);
        }
    }

    private void OnEnvironmentCreated(int errorCode, nint environmentPointer)
    {
        _environmentHandler = null;

        if (errorCode < 0 || environmentPointer == 0)
        {
            Log.Error($"webview: environment creation failed with 0x{errorCode:X8}");
            return;
        }

        if (_disposed)
        {
            Log.Info("webview: environment arrived after the surface was closed, discarding it");
            return;
        }

        _environment = ComHandle<ICoreWebView2Environment>.Borrow(environmentPointer);
        if (_environment is null)
        {
            Log.Error("webview: the created environment does not answer to the environment interface");
            return;
        }

        Log.Info($"webview: environment created, browser version {ReadBrowserVersion() ?? "unknown"}");

        _controllerHandler = new ControllerCompletedHandler(OnControllerCreated);
        var callback = CreateCallback(_controllerHandler, WebViewIid.ControllerCompletedHandler);
        if (callback == 0)
        {
            Log.Error("webview: could not expose the controller callback to the browser");
            return;
        }

        try
        {
            var hr = _environment.Target.CreateCoreWebView2Controller(_parentWindow, callback);
            if (hr < 0)
            {
                Log.Error($"webview: CreateCoreWebView2Controller returned 0x{hr:X8}");
            }
        }
        finally
        {
            Marshal.Release(callback);
        }
    }

    private void OnControllerCreated(int errorCode, nint controllerPointer)
    {
        _controllerHandler = null;

        if (errorCode < 0 || controllerPointer == 0)
        {
            Log.Error($"webview: controller creation failed with 0x{errorCode:X8}");
            return;
        }

        if (_disposed)
        {
            // The window is already gone. Take the controller only to shut it down: leaving it
            // alive leaves browser processes parented to a destroyed window.
            using var orphan = ComHandle<ICoreWebView2Controller>.Borrow(controllerPointer);
            orphan?.Target.Close();
            Log.Info("webview: controller arrived after the surface was closed, closed it");
            return;
        }

        _controller = ComHandle<ICoreWebView2Controller>.Borrow(controllerPointer);
        if (_controller is null)
        {
            Log.Error("webview: the created controller does not answer to the controller interface");
            return;
        }

        // A fresh controller has accepted nothing yet, whatever an earlier one was told.
        _appliedBounds = null;

        Log.Info("webview: controller created");

        // The browser's own idle colour, shown wherever the renderer has nothing newer: behind
        // a loading page and in the fringe of a resize. Its default is white, and a white flash
        // through a dark editor reads as a defect on its own. Asked for by interface revision,
        // and skipped without complaint on a runtime old enough to lack it.
        if (Marshal.QueryInterface(controllerPointer, in WebViewIid.Controller2, out var controller2Pointer) >= 0
            && controller2Pointer != 0)
        {
            using var controller2 = ComHandle<ICoreWebView2Controller2>.Own(controller2Pointer);
            controller2?.Target.PutDefaultBackgroundColor(
                new WebViewColor { A = 255, R = 0x1E, G = 0x1E, B = 0x1E });
            Log.Verbose("webview: idle background set to the theme ground");
        }

        SetBounds(_bounds);

        // Held invisible until the page reports ready. A controller shown now is the
        // compositor's blank rectangle for as long as the page takes to arrive; the overlay's
        // loader owns those pixels instead, and Reveal retires it.
        _controller.Target.PutIsVisible(0);

        if (_controller.Target.GetCoreWebView2(out var viewPointer) < 0 || viewPointer == 0)
        {
            Log.Error("webview: the controller exposed no content surface");
            return;
        }

        _webView = ComHandle<ICoreWebView2>.Own(viewPointer);
        if (_webView is null)
        {
            Log.Error("webview: the content surface does not answer to the view interface");
            return;
        }

        // The later revision is asked for once and kept. A runtime that predates it refuses the
        // query, which is not a failure: the shell document still renders, and the two things that
        // interface carries are the two things the shell document does not use.
        _view = _webView.As<ICoreWebView2_3>(WebViewIid.WebView3);
        if (_view is null)
        {
            Log.Warn(
                "webview: this runtime is older than the folder mapping and the message channel, " +
                "so the editor bundle cannot be served and the page cannot talk to the shim");
        }

        SubscribeNavigationCompleted();
        SubscribeWebMessageReceived();
        SubscribeAcceleratorKeyPressed();
        Navigate();
    }

    /// <summary>
    /// Makes the browser visible. Held back until the page says ready, so the first thing seen
    /// over the pane is the styled surface rather than a blank coming into being.
    /// </summary>
    public void Reveal()
    {
        _controller?.Target.PutIsVisible(1);
    }

    /// <summary>
    /// Subscribes to accelerator keys. This is on the controller rather than the content surface,
    /// because it is about the browser's place in the host's keyboard handling rather than about
    /// the document.
    /// </summary>
    private void SubscribeAcceleratorKeyPressed()
    {
        var controller = _controller;
        if (controller is null)
        {
            return;
        }

        _acceleratorHandler = new AcceleratorKeyPressedHandler(OnAcceleratorKeyPressed);
        var callback = CreateCallback(_acceleratorHandler, WebViewIid.AcceleratorKeyPressedHandler);
        if (callback == 0)
        {
            Log.Error("webview: could not expose the accelerator callback to the browser");
            return;
        }

        try
        {
            var hr = controller.Target.AddAcceleratorKeyPressed(callback, out _acceleratorKeyPressedToken);
            if (hr < 0)
            {
                Log.Error($"webview: add_AcceleratorKeyPressed returned 0x{hr:X8}");
            }
        }
        finally
        {
            Marshal.Release(callback);
        }
    }

    private void OnAcceleratorKeyPressed(nint sender, nint argsPointer)
    {
        var handler = AcceleratorPressed;
        if (handler is null)
        {
            return;
        }

        using var args = ComHandle<ICoreWebView2AcceleratorKeyPressedEventArgs>.Borrow(argsPointer);
        if (args is null)
        {
            return;
        }

        // Key down only. Every key raises this twice, and acting on both runs the command twice.
        if (args.Target.GetKeyEventKind(out var kind) < 0 || kind != KeyEventKind.KeyDown)
        {
            return;
        }

        if (args.Target.GetVirtualKey(out var key) < 0)
        {
            return;
        }

        if (handler(key))
        {
            args.Target.PutHandled(1);
        }
    }

    private void SubscribeNavigationCompleted()
    {
        var view = _webView;
        if (view is null)
        {
            return;
        }

        _navigationHandler = new NavigationCompletedHandler(OnNavigationCompleted);
        var callback = CreateCallback(_navigationHandler, WebViewIid.NavigationCompletedHandler);
        if (callback == 0)
        {
            return;
        }

        try
        {
            view.Target.AddNavigationCompleted(callback, out _navigationCompletedToken);
        }
        finally
        {
            Marshal.Release(callback);
        }
    }

    /// <summary>
    /// Subscribes to messages from the page. The event sits past the point the first revision of
    /// the view interface is declared to, so this needs the later one and does nothing without it.
    /// </summary>
    private void SubscribeWebMessageReceived()
    {
        var view = _view;
        if (view is null)
        {
            return;
        }

        _messageHandler = new WebMessageReceivedHandler(OnWebMessageReceived);
        var callback = CreateCallback(_messageHandler, WebViewIid.WebMessageReceivedHandler);
        if (callback == 0)
        {
            Log.Error("webview: could not expose the message callback to the browser");
            return;
        }

        try
        {
            var hr = view.Target.AddWebMessageReceived(callback, out _webMessageReceivedToken);
            if (hr < 0)
            {
                Log.Error($"webview: add_WebMessageReceived returned 0x{hr:X8}");
                return;
            }

            Log.Info("webview: listening for messages from the page");
        }
        finally
        {
            // The browser holds its own reference for as long as the subscription lasts.
            Marshal.Release(callback);
        }
    }

    /// <summary>
    /// Sends the surface to the editing bundle.
    ///
    /// There is one document. An earlier version chose between this and a diagnostic page by
    /// looking at what happened to be on disk, and once both existed it put the wrong one in the
    /// wrong window; the failure presented as a rendering fault rather than as a wrong document.
    /// A surface with nothing to show is a broken install, and it says so in the log rather than
    /// rendering something that looks deliberate.
    /// </summary>
    private void Navigate()
    {
        if (_view is null)
        {
            Log.Error("webview: this runtime is older than the folder mapping, so nothing can be served");
            return;
        }

        var directory = ShimModule.Directory;
        if (directory is null)
        {
            Log.Error("webview: the shim could not locate its own directory, so the bundle cannot be found");
            return;
        }

        var entry = WebViewPaths.EditorEntryDocument(directory);
        if (!File.Exists(entry))
        {
            Log.Error($"webview: no editor bundle at {entry}, so there is nothing to show");
            return;
        }

        // Loopback first: the folder mapping brokers every byte through the browser's host
        // pipe at about two megabytes a second, which billed two seconds of every start-up to
        // fetching a bundle that sits on local disk. A socket to 127.0.0.1 serves the same
        // bytes in tens of milliseconds. The mapping remains the fallback below, because a
        // slow editor beats no editor.
        _pageServer ??= LoopbackPageServer.Start(WebViewPaths.EditorContentRoot(directory));
        if (_pageServer is not null && NavigateToUrl($"{_pageServer.BaseUrl}/index.html"))
        {
            return;
        }

        _pageServer?.Dispose();
        _pageServer = null;

        // The mapping has to be in place before the navigation, not after: the address only
        // resolves because of it.
        if (!SetContentRoot(WebViewPaths.EditorHostName, WebViewPaths.EditorContentRoot(directory)))
        {
            return;
        }

        if (!NavigateToUrl(WebViewPaths.EditorEntryUrl(WebViewPaths.EditorHostName)))
        {
            // The mapping outlives a failed navigation and would shadow the host name for anything
            // that came later, so it is taken back rather than left behind.
            _view.Target.ClearVirtualHostNameToFolderMapping(WebViewPaths.EditorHostName);
        }
    }

    private void OnNavigationCompleted(nint sender, nint argsPointer)
    {
        using var args = ComHandle<ICoreWebView2NavigationCompletedEventArgs>.Borrow(argsPointer);
        if (args is null)
        {
            Log.Info("webview: navigated");
            return;
        }

        args.Target.GetIsSuccess(out var success);
        if (success != 0)
        {
            Log.Info("webview: navigated");
            return;
        }

        args.Target.GetWebErrorStatus(out var status);
        Log.Error($"webview: navigation failed with web error status {status}");
    }

    private void OnWebMessageReceived(nint sender, nint argsPointer)
    {
        using var args = ComHandle<ICoreWebView2WebMessageReceivedEventArgs>.Borrow(argsPointer);
        if (args is null)
        {
            Log.Error("webview: a page message arrived without arguments");
            return;
        }

        var message = ReadMessage(args.Target);
        if (message is null)
        {
            Log.Error("webview: a page message could not be read as text");
            return;
        }

        var handler = MessageReceived;
        if (handler is null)
        {
            // Said once, not once per message: this runs on the host user interface thread and the
            // log is a file append. One line proves the channel works and nothing is listening yet;
            // the rest would only be the same line at whatever rate the page posts.
            if (!_reportedUnhandledMessage)
            {
                _reportedUnhandledMessage = true;
                Log.Info($"webview: the page is posting messages and nothing is listening, {message.Length} characters");
            }

            return;
        }

        handler(message);
    }

    /// <summary>
    /// Reads a message as text whichever way the page posted it. A posted string comes back from
    /// the first call; a posted object fails it and is only available as its JSON form.
    /// </summary>
    private static string? ReadMessage(ICoreWebView2WebMessageReceivedEventArgs args)
    {
        if (args.TryGetWebMessageAsString(out var text) >= 0 && text != 0)
        {
            return TakeString(text);
        }

        return args.GetWebMessageAsJson(out var json) >= 0 && json != 0 ? TakeString(json) : null;
    }

    private string? ReadBrowserVersion()
    {
        var environment = _environment;
        return environment is not null
            && environment.Target.GetBrowserVersionString(out var text) >= 0
            && text != 0
            ? TakeString(text)
            : null;
    }

    /// <summary>
    /// Copies a string the browser returned and frees the original.
    ///
    /// Every string handed back through an out parameter here is allocated with the task allocator
    /// and owned by the caller. Nothing else frees it, and one leaked per message would be a leak
    /// for the life of the host process.
    /// </summary>
    private static string? TakeString(nint value)
    {
        try
        {
            return Marshal.PtrToStringUni(value);
        }
        finally
        {
            Marshal.FreeCoTaskMem(value);
        }
    }

    private static bool EnsureLoaderLoaded()
    {
        if (_loaderLoaded)
        {
            return true;
        }

        var directory = ShimModule.Directory;
        if (directory is null)
        {
            Log.Error("webview: the shim could not locate its own directory, so the loader cannot be found");
            return false;
        }

        var path = WebViewPaths.LoaderLibrary(directory);
        if (!File.Exists(path))
        {
            Log.Error($"webview: no loader library at {path}");
            return false;
        }

        // Loaded by full path and never freed. The search path the runtime would otherwise use
        // starts at the host executable's directory, which belongs to Office, not to us.
        if (!NativeLibrary.TryLoad(path, out _))
        {
            Log.Error($"webview: the loader library at {path} could not be loaded");
            return false;
        }

        _loaderLoaded = true;
        Log.Info($"webview: loader library loaded from {path}");
        return true;
    }

    /// <summary>
    /// Builds a callback object the browser can hold, and hands back the one reference we own.
    /// </summary>
    private static nint CreateCallback(object handler, in Guid interfaceId)
    {
        var unknown = ComRuntime.Wrappers.GetOrCreateComInterfaceForObject(handler, CreateComInterfaceFlags.None);
        try
        {
            return Marshal.QueryInterface(unknown, in interfaceId, out var callback) < 0 ? 0 : callback;
        }
        finally
        {
            Marshal.Release(unknown);
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        MessageReceived = null;
        AcceleratorPressed = null;

        _pageServer?.Dispose();
        _pageServer = null;

        if (_controller is not null && _acceleratorKeyPressedToken.Value != 0)
        {
            _controller.Target.RemoveAcceleratorKeyPressed(_acceleratorKeyPressedToken);
            _acceleratorKeyPressedToken = default;
        }

        // Subscriptions come off first, while the objects they were made against are still alive.
        var view = _webView;
        if (view is not null && _navigationCompletedToken.Value != 0)
        {
            view.Target.RemoveNavigationCompleted(_navigationCompletedToken);
            _navigationCompletedToken = default;
        }

        if (_view is not null && _webMessageReceivedToken.Value != 0)
        {
            _view.Target.RemoveWebMessageReceived(_webMessageReceivedToken);
            _webMessageReceivedToken = default;
        }

        // Closing the controller is what actually tears down the browser processes. Releasing the
        // references alone would leave them running until the host exits.
        _controller?.Target.Close();

        // Its own reference, taken by the query that produced it, and released here exactly once.
        _view?.Dispose();
        _view = null;

        _webView?.Dispose();
        _webView = null;

        _controller?.Dispose();
        _controller = null;

        _environment?.Dispose();
        _environment = null;

        _environmentHandler = null;
        _controllerHandler = null;
        _navigationHandler = null;
        _messageHandler = null;
        _acceleratorHandler = null;
    }
}

/// <summary>Which half of a keystroke the browser is reporting.</summary>
internal static class KeyEventKind
{
    public const int KeyDown = 0;
    public const int KeyUp = 1;
    public const int SystemKeyDown = 2;
    public const int SystemKeyUp = 3;
}

/// <summary>Receives the environment once the loader has created it.</summary>
[GeneratedComClass]
internal sealed partial class EnvironmentCompletedHandler : ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
{
    private readonly Action<int, nint> _completed;

    public EnvironmentCompletedHandler(Action<int, nint> completed) => _completed = completed;

    public int Invoke(int errorCode, nint createdEnvironment)
    {
        try
        {
            _completed(errorCode, createdEnvironment);
        }
        catch (Exception ex)
        {
            Log.Error("webview: the environment callback failed", ex);
        }

        return HResult.Ok;
    }
}

/// <summary>Receives the controller once the environment has created it.</summary>
[GeneratedComClass]
internal sealed partial class ControllerCompletedHandler : ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
{
    private readonly Action<int, nint> _completed;

    public ControllerCompletedHandler(Action<int, nint> completed) => _completed = completed;

    public int Invoke(int errorCode, nint createdController)
    {
        try
        {
            _completed(errorCode, createdController);
        }
        catch (Exception ex)
        {
            Log.Error("webview: the controller callback failed", ex);
        }

        return HResult.Ok;
    }
}

/// <summary>Reports the outcome of each navigation.</summary>
[GeneratedComClass]
internal sealed partial class NavigationCompletedHandler : ICoreWebView2NavigationCompletedEventHandler
{
    private readonly Action<nint, nint> _completed;

    public NavigationCompletedHandler(Action<nint, nint> completed) => _completed = completed;

    public int Invoke(nint sender, nint args)
    {
        try
        {
            _completed(sender, args);
        }
        catch (Exception ex)
        {
            Log.Error("webview: the navigation callback failed", ex);
        }

        return HResult.Ok;
    }
}

/// <summary>Delivers each accelerator key before the page sees it.</summary>
[GeneratedComClass]
internal sealed partial class AcceleratorKeyPressedHandler : ICoreWebView2AcceleratorKeyPressedEventHandler
{
    private readonly Action<nint, nint> _pressed;

    public AcceleratorKeyPressedHandler(Action<nint, nint> pressed) => _pressed = pressed;

    public int Invoke(nint sender, nint args)
    {
        try
        {
            _pressed(sender, args);
        }
        catch (Exception ex)
        {
            Log.Error("webview: the accelerator callback failed", ex);
        }

        return HResult.Ok;
    }
}

/// <summary>
/// Delivers each message the page posts. The browser calls this on the host user interface thread
/// and treats a failure as the page's problem, so the exception never gets past here.
/// </summary>
[GeneratedComClass]
internal sealed partial class WebMessageReceivedHandler : ICoreWebView2WebMessageReceivedEventHandler
{
    private readonly Action<nint, nint> _received;

    public WebMessageReceivedHandler(Action<nint, nint> received) => _received = received;

    public int Invoke(nint sender, nint args)
    {
        try
        {
            _received(sender, args);
        }
        catch (Exception ex)
        {
            Log.Error("webview: the message callback failed", ex);
        }

        return HResult.Ok;
    }
}
