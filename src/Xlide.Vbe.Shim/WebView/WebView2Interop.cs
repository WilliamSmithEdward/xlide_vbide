using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.WebView;

// These declarations are transcribed from WebView2.h in the Microsoft.Web.WebView2 package, which
// is also where the loader binary beside the shim comes from. The same rule as the OLE interfaces
// applies and matters more here, because these vtables are long: a member declared out of order
// does not fail, it calls the neighbouring function.
//
// Two interfaces are declared only as far as the last member this shim calls. That is safe in one
// direction only. Calling through a truncated declaration is fine, because a vtable is read by
// index and the entries beyond the last declared one are never touched. Implementing one would not
// be, and nothing here implements them. Where a member is added later, it is appended, never
// inserted.

/// <summary>Token returned when subscribing to a browser event, and required to unsubscribe.</summary>
[StructLayout(LayoutKind.Sequential)]
internal struct EventRegistrationToken
{
    public long Value;
}

/// <summary>Why focus is being moved into the browser.</summary>
internal static class MoveFocusReason
{
    public const int Programmatic = 0;
}

/// <summary>
/// The browser environment: a browser version, a user data folder, and the factory for controllers
/// bound to it. Created asynchronously by the loader.
/// </summary>
/// <remarks>Declared in full; the interface has five members.</remarks>
[GeneratedComInterface]
[Guid("b96d755e-0319-4e92-a296-23436f46a1fc")]
internal partial interface ICoreWebView2Environment
{
    [PreserveSig]
    int CreateCoreWebView2Controller(nint parentWindow, nint handler);

    [PreserveSig]
    int CreateWebResourceResponse(nint content, int statusCode, nint reasonPhrase, nint headers, out nint response);

    [PreserveSig]
    int GetBrowserVersionString(out nint versionInfo);

    [PreserveSig]
    int AddNewBrowserVersionAvailable(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveNewBrowserVersionAvailable(EventRegistrationToken token);
}

/// <summary>
/// The browser attached to one parent window. It owns positioning, visibility, and focus; the
/// content itself is reached through the view it exposes.
/// </summary>
/// <remarks>Declared in full; the interface has twenty-three members.</remarks>
[GeneratedComInterface]
[Guid("4d00c0d1-9434-4eb6-8078-8697a560334f")]
internal partial interface ICoreWebView2Controller
{
    [PreserveSig]
    int GetIsVisible(out int isVisible);

    [PreserveSig]
    int PutIsVisible(int isVisible);

    [PreserveSig]
    int GetBounds(out Rect bounds);

    [PreserveSig]
    int PutBounds(Rect bounds);

    [PreserveSig]
    int GetZoomFactor(out double zoomFactor);

    [PreserveSig]
    int PutZoomFactor(double zoomFactor);

    [PreserveSig]
    int AddZoomFactorChanged(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveZoomFactorChanged(EventRegistrationToken token);

    [PreserveSig]
    int SetBoundsAndZoomFactor(Rect bounds, double zoomFactor);

    [PreserveSig]
    int MoveFocus(int reason);

    [PreserveSig]
    int AddMoveFocusRequested(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveMoveFocusRequested(EventRegistrationToken token);

    [PreserveSig]
    int AddGotFocus(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveGotFocus(EventRegistrationToken token);

    [PreserveSig]
    int AddLostFocus(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveLostFocus(EventRegistrationToken token);

    [PreserveSig]
    int AddAcceleratorKeyPressed(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveAcceleratorKeyPressed(EventRegistrationToken token);

    [PreserveSig]
    int GetParentWindow(out nint parentWindow);

    [PreserveSig]
    int PutParentWindow(nint parentWindow);

    [PreserveSig]
    int NotifyParentWindowPositionChanged();

    [PreserveSig]
    int Close();

    [PreserveSig]
    int GetCoreWebView2(out nint coreWebView2);
}

/// <summary>
/// The content surface: navigation, scripting, and the message channel that will eventually carry
/// everything between the editor surface and the shim.
/// </summary>
/// <remarks>
/// Declared as far as the fourteenth member, which is where the navigation completion event sits.
/// The interface continues past that point and is only ever called through, never implemented.
/// </remarks>
[GeneratedComInterface]
[Guid("76eceacb-0462-4d94-ac83-423a6793775e")]
internal partial interface ICoreWebView2
{
    [PreserveSig]
    int GetSettings(out nint settings);

    [PreserveSig]
    int GetSource(out nint uri);

    [PreserveSig]
    int Navigate([MarshalAs(UnmanagedType.LPWStr)] string uri);

    [PreserveSig]
    int NavigateToString([MarshalAs(UnmanagedType.LPWStr)] string htmlContent);

    [PreserveSig]
    int AddNavigationStarting(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveNavigationStarting(EventRegistrationToken token);

    [PreserveSig]
    int AddContentLoading(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveContentLoading(EventRegistrationToken token);

    [PreserveSig]
    int AddSourceChanged(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveSourceChanged(EventRegistrationToken token);

    [PreserveSig]
    int AddHistoryChanged(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveHistoryChanged(EventRegistrationToken token);

    [PreserveSig]
    int AddNavigationCompleted(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveNavigationCompleted(EventRegistrationToken token);
}

/// <summary>Outcome of a navigation. Read from the completion event.</summary>
/// <remarks>Declared in full; the interface has three members.</remarks>
[GeneratedComInterface]
[Guid("30d68b7d-20d9-4752-a9ca-ec8448fbb5c1")]
internal partial interface ICoreWebView2NavigationCompletedEventArgs
{
    [PreserveSig]
    int GetIsSuccess(out int isSuccess);

    [PreserveSig]
    int GetWebErrorStatus(out int webErrorStatus);

    [PreserveSig]
    int GetNavigationId(out ulong navigationId);
}

/// <summary>Callback the loader invokes once the environment exists, or once it has failed to.</summary>
[GeneratedComInterface]
[Guid("4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d")]
internal partial interface ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
{
    [PreserveSig]
    int Invoke(int errorCode, nint createdEnvironment);
}

/// <summary>Callback the environment invokes once the controller exists, or once it has failed to.</summary>
[GeneratedComInterface]
[Guid("6c4819f3-c9b7-4260-8127-c9f5bde7f68c")]
internal partial interface ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
{
    [PreserveSig]
    int Invoke(int errorCode, nint createdController);
}

/// <summary>Callback raised when a navigation finishes, successfully or not.</summary>
[GeneratedComInterface]
[Guid("d33a35bf-1c49-4f98-93ab-006e0533fe1c")]
internal partial interface ICoreWebView2NavigationCompletedEventHandler
{
    [PreserveSig]
    int Invoke(nint sender, nint args);
}

/// <summary>Identifiers of the interfaces this shim implements for the browser to call back on.</summary>
internal static class WebViewIid
{
    public static readonly Guid EnvironmentCompletedHandler = new("4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d");
    public static readonly Guid ControllerCompletedHandler = new("6c4819f3-c9b7-4260-8127-c9f5bde7f68c");
    public static readonly Guid NavigationCompletedHandler = new("d33a35bf-1c49-4f98-93ab-006e0533fe1c");
    public static readonly Guid NavigationCompletedEventArgs = new("30d68b7d-20d9-4752-a9ca-ec8448fbb5c1");
    public static readonly Guid Environment = new("b96d755e-0319-4e92-a296-23436f46a1fc");
    public static readonly Guid Controller = new("4d00c0d1-9434-4eb6-8078-8697a560334f");
    public static readonly Guid WebView = new("76eceacb-0462-4d94-ac83-423a6793775e");
}

/// <summary>
/// The one entry point the browser runtime exposes as a plain export.
///
/// It lives in WebView2Loader.dll, which is deployed beside the shim rather than relied upon to be
/// anywhere on the search path. Excel's own directory is the first place the loader would look and
/// it is not ours to write to, so the library is loaded explicitly by full path before the first
/// call. Once loaded, the module is found by name and this declaration binds to it.
/// </summary>
internal static partial class WebView2Loader
{
    [LibraryImport("WebView2Loader.dll", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int CreateCoreWebView2EnvironmentWithOptions(
        string? browserExecutableFolder,
        string? userDataFolder,
        nint environmentOptions,
        nint environmentCreatedHandler);
}
