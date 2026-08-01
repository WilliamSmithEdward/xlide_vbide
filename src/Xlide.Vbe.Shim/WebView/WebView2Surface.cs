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
/// </summary>
internal sealed class WebView2Surface : IDisposable
{
    private static bool _loaderLoaded;

    private readonly nint _parentWindow;

    private ComHandle<ICoreWebView2Environment>? _environment;
    private ComHandle<ICoreWebView2Controller>? _controller;
    private ComHandle<ICoreWebView2>? _webView;

    private EnvironmentCompletedHandler? _environmentHandler;
    private ControllerCompletedHandler? _controllerHandler;
    private NavigationCompletedHandler? _navigationHandler;
    private EventRegistrationToken _navigationCompletedToken;

    private PixelRect _bounds;
    private bool _disposed;

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

    /// <summary>Sizes the browser to the window's client area.</summary>
    public void SetBounds(PixelRect bounds)
    {
        _bounds = bounds;

        var controller = _controller;
        if (controller is null)
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
    }

    /// <summary>Gives keyboard focus to the browser.</summary>
    public void Focus() => _controller?.Target.MoveFocus(MoveFocusReason.Programmatic);

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

        Log.Info("webview: controller created");

        SetBounds(_bounds);
        _controller.Target.PutIsVisible(1);

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

        SubscribeNavigationCompleted();
        Navigate();
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

    private void Navigate()
    {
        var view = _webView;
        if (view is null)
        {
            return;
        }

        var document = LoadShellDocument();

        // The document is pushed rather than fetched. A file URL would put the surface in an origin
        // that blocks module scripts and storage, which the real editor surface needs; the surface
        // will move to a virtual host name mapped to the UI directory when it does.
        var hr = view.Target.NavigateToString(document);
        if (hr < 0)
        {
            Log.Error($"webview: NavigateToString returned 0x{hr:X8}");
            return;
        }

        Log.Info("webview: navigation started");
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

    private string LoadShellDocument()
    {
        var directory = ShimModule.Directory;
        var path = directory is null ? null : WebViewPaths.ShellDocument(directory);

        if (path is null || !File.Exists(path))
        {
            Log.Warn($"webview: no shell document at {path ?? "an unknown location"}");
            return ShellDocument.Missing(path ?? WebViewPaths.ShellDocumentRelativePath);
        }

        try
        {
            return ShellDocument.Compose(File.ReadAllText(path), ReadBrowserVersion());
        }
        catch (IOException ex)
        {
            Log.Error($"webview: could not read the shell document at {path}", ex);
            return ShellDocument.Missing(path);
        }
        catch (UnauthorizedAccessException ex)
        {
            Log.Error($"webview: could not read the shell document at {path}", ex);
            return ShellDocument.Missing(path);
        }
    }

    private string? ReadBrowserVersion()
    {
        var environment = _environment;
        if (environment is null || environment.Target.GetBrowserVersionString(out var text) < 0 || text == 0)
        {
            return null;
        }

        try
        {
            return Marshal.PtrToStringUni(text);
        }
        finally
        {
            // The string is allocated by the browser with the task allocator and owned by us.
            Marshal.FreeCoTaskMem(text);
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

        var view = _webView;
        if (view is not null && _navigationCompletedToken.Value != 0)
        {
            view.Target.RemoveNavigationCompleted(_navigationCompletedToken);
            _navigationCompletedToken = default;
        }

        // Closing the controller is what actually tears down the browser processes. Releasing the
        // references alone would leave them running until the host exits.
        _controller?.Target.Close();

        _webView?.Dispose();
        _webView = null;

        _controller?.Dispose();
        _controller = null;

        _environment?.Dispose();
        _environment = null;

        _environmentHandler = null;
        _controllerHandler = null;
        _navigationHandler = null;
    }
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
