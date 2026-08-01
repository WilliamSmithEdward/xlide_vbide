using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.WebView;

// These declarations are transcribed from WebView2.h in the Microsoft.Web.WebView2 package, which
// is also where the loader binary beside the shim comes from. The same rule as the OLE interfaces
// applies and matters more here, because these vtables are long: a member declared out of order
// does not fail, it calls the neighbouring function.
//
// Some interfaces are declared only as far as the last member this shim calls. That is safe in one
// direction only. Calling through a truncated declaration is fine, because a vtable is read by
// index and the entries beyond the last declared one are never touched. Implementing one would not
// be, and nothing here implements them. Where a member is added later, it is appended, never
// inserted.
//
// Interfaces the header derives from another are declared flat here, repeating every inherited
// member in order. C# interface inheritance is not vtable inheritance: the source generator lays
// out one vtable per interface from that interface's own members, so a derived declaration would
// describe a vtable that starts where the base one ends.

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
/// How much of a folder mapped to a virtual host name the browser will serve.
///
/// COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND in the header, passed as a plain integer because the
/// value crosses the boundary as one and an enumeration type here would only add a cast.
/// </summary>
internal static class HostResourceAccessKind
{
    /// <summary>Nothing under the folder is served.</summary>
    public const int Deny = 0;

    /// <summary>Everything is served, including to requests from other origins.</summary>
    public const int Allow = 1;

    /// <summary>
    /// Served to the document loaded from the mapped host, refused to cross-origin requests.
    ///
    /// This is what a local application bundle wants. The editor surface loads its own files, and
    /// nothing else has a reason to read them; Allow would let any page the browser ever loads
    /// fetch the contents of a folder inside the install directory.
    /// </summary>
    public const int DenyCors = 2;
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

/// <summary>
/// The same content surface two revisions later, which is where the two capabilities this shim
/// needs actually live: mapping a folder to a virtual host name, and the message channel.
///
/// Declared flat, listing every inherited member in vtable order, rather than as a C# interface
/// deriving from <see cref="ICoreWebView2"/>. That is deliberate and it is the same choice the OLE
/// interfaces in this project make. C# interface inheritance says nothing about layout: the source
/// generator builds one vtable per interface from the members declared on that interface alone, so
/// a derived declaration would produce a five-entry vtable that gets called as if it were seventy.
/// The header's inheritance chain is a layout fact, and the only way to state a layout fact here is
/// to write the layout out.
///
/// Seventy slots, in this order: the fifty-eight of ICoreWebView2, then the seven ICoreWebView2_2
/// appends, then the five ICoreWebView2_3 appends. Slots this shim never calls carry a placeholder
/// signature and a name that says so; the signature of a slot that is never called is irrelevant,
/// its position is not. Slot numbers are marked every ten members so the list can be checked
/// against the header without counting from the top.
/// </summary>
[GeneratedComInterface]
[Guid("a0d6df20-3b92-416d-aa0c-437a9c727857")]
internal partial interface ICoreWebView2_3
{
    // ---- ICoreWebView2, slots 1 to 58 --------------------------------------------------------

    // 1
    [PreserveSig]
    int UnusedGetSettings();

    [PreserveSig]
    int UnusedGetSource();

    [PreserveSig]
    int Navigate([MarshalAs(UnmanagedType.LPWStr)] string uri);

    [PreserveSig]
    int NavigateToString([MarshalAs(UnmanagedType.LPWStr)] string htmlContent);

    [PreserveSig]
    int UnusedAddNavigationStarting();

    [PreserveSig]
    int UnusedRemoveNavigationStarting();

    [PreserveSig]
    int UnusedAddContentLoading();

    [PreserveSig]
    int UnusedRemoveContentLoading();

    [PreserveSig]
    int UnusedAddSourceChanged();

    [PreserveSig]
    int UnusedRemoveSourceChanged();

    // 11
    [PreserveSig]
    int UnusedAddHistoryChanged();

    [PreserveSig]
    int UnusedRemoveHistoryChanged();

    [PreserveSig]
    int AddNavigationCompleted(nint eventHandler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveNavigationCompleted(EventRegistrationToken token);

    [PreserveSig]
    int UnusedAddFrameNavigationStarting();

    [PreserveSig]
    int UnusedRemoveFrameNavigationStarting();

    [PreserveSig]
    int UnusedAddFrameNavigationCompleted();

    [PreserveSig]
    int UnusedRemoveFrameNavigationCompleted();

    [PreserveSig]
    int UnusedAddScriptDialogOpening();

    [PreserveSig]
    int UnusedRemoveScriptDialogOpening();

    // 21
    [PreserveSig]
    int UnusedAddPermissionRequested();

    [PreserveSig]
    int UnusedRemovePermissionRequested();

    [PreserveSig]
    int UnusedAddProcessFailed();

    [PreserveSig]
    int UnusedRemoveProcessFailed();

    [PreserveSig]
    int UnusedAddScriptToExecuteOnDocumentCreated();

    [PreserveSig]
    int UnusedRemoveScriptToExecuteOnDocumentCreated();

    [PreserveSig]
    int ExecuteScript([MarshalAs(UnmanagedType.LPWStr)] string javaScript, nint handler);

    [PreserveSig]
    int UnusedCapturePreview();

    [PreserveSig]
    int UnusedReload();

    [PreserveSig]
    int PostWebMessageAsJson([MarshalAs(UnmanagedType.LPWStr)] string webMessageAsJson);

    // 31
    [PreserveSig]
    int PostWebMessageAsString([MarshalAs(UnmanagedType.LPWStr)] string webMessageAsString);

    [PreserveSig]
    int AddWebMessageReceived(nint handler, out EventRegistrationToken token);

    [PreserveSig]
    int RemoveWebMessageReceived(EventRegistrationToken token);

    [PreserveSig]
    int UnusedCallDevToolsProtocolMethod();

    [PreserveSig]
    int UnusedGetBrowserProcessId();

    [PreserveSig]
    int UnusedGetCanGoBack();

    [PreserveSig]
    int UnusedGetCanGoForward();

    [PreserveSig]
    int UnusedGoBack();

    [PreserveSig]
    int UnusedGoForward();

    [PreserveSig]
    int UnusedGetDevToolsProtocolEventReceiver();

    // 41
    [PreserveSig]
    int UnusedStop();

    [PreserveSig]
    int UnusedAddNewWindowRequested();

    [PreserveSig]
    int UnusedRemoveNewWindowRequested();

    [PreserveSig]
    int UnusedAddDocumentTitleChanged();

    [PreserveSig]
    int UnusedRemoveDocumentTitleChanged();

    [PreserveSig]
    int UnusedGetDocumentTitle();

    [PreserveSig]
    int UnusedAddHostObjectToScript();

    [PreserveSig]
    int UnusedRemoveHostObjectFromScript();

    [PreserveSig]
    int UnusedOpenDevToolsWindow();

    [PreserveSig]
    int UnusedAddContainsFullScreenElementChanged();

    // 51
    [PreserveSig]
    int UnusedRemoveContainsFullScreenElementChanged();

    [PreserveSig]
    int UnusedGetContainsFullScreenElement();

    [PreserveSig]
    int UnusedAddWebResourceRequested();

    [PreserveSig]
    int UnusedRemoveWebResourceRequested();

    [PreserveSig]
    int UnusedAddWebResourceRequestedFilter();

    [PreserveSig]
    int UnusedRemoveWebResourceRequestedFilter();

    [PreserveSig]
    int UnusedAddWindowCloseRequested();

    [PreserveSig]
    int UnusedRemoveWindowCloseRequested();

    // ---- ICoreWebView2_2 appends, slots 59 to 65 ---------------------------------------------

    // 59
    [PreserveSig]
    int UnusedAddWebResourceResponseReceived();

    [PreserveSig]
    int UnusedRemoveWebResourceResponseReceived();

    // 61
    [PreserveSig]
    int UnusedNavigateWithWebResourceRequest();

    [PreserveSig]
    int UnusedAddDomContentLoaded();

    [PreserveSig]
    int UnusedRemoveDomContentLoaded();

    [PreserveSig]
    int UnusedGetCookieManager();

    [PreserveSig]
    int UnusedGetEnvironment();

    // ---- ICoreWebView2_3 appends, slots 66 to 70 ---------------------------------------------

    // 66
    [PreserveSig]
    int UnusedTrySuspend();

    [PreserveSig]
    int UnusedResume();

    [PreserveSig]
    int UnusedGetIsSuspended();

    [PreserveSig]
    int SetVirtualHostNameToFolderMapping(
        [MarshalAs(UnmanagedType.LPWStr)] string hostName,
        [MarshalAs(UnmanagedType.LPWStr)] string folderPath,
        int accessKind);

    // 70
    [PreserveSig]
    int ClearVirtualHostNameToFolderMapping([MarshalAs(UnmanagedType.LPWStr)] string hostName);
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

/// <summary>One message the page posted through window.chrome.webview.postMessage.</summary>
/// <remarks>
/// Declared in full; the interface has three members. Every string it hands back is allocated with
/// the task allocator and owned by the caller, so each one is freed here rather than leaked for the
/// life of the host process.
/// </remarks>
[GeneratedComInterface]
[Guid("0f99a40c-e962-4207-9e92-e3d542eff849")]
internal partial interface ICoreWebView2WebMessageReceivedEventArgs
{
    [PreserveSig]
    int GetSource(out nint uri);

    [PreserveSig]
    int GetWebMessageAsJson(out nint json);

    /// <summary>
    /// The message as text, which succeeds only when the page posted a string. A page that posted
    /// an object gets E_INVALIDARG here and has to be read as JSON instead.
    /// </summary>
    [PreserveSig]
    int TryGetWebMessageAsString(out nint text);
}

/// <summary>Callback raised for each message the page posts.</summary>
[GeneratedComInterface]
[Guid("57213f19-00e6-49fa-8e07-898ea01ecbd2")]
internal partial interface ICoreWebView2WebMessageReceivedEventHandler
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
    public static readonly Guid WebMessageReceivedHandler = new("57213f19-00e6-49fa-8e07-898ea01ecbd2");
    public static readonly Guid WebMessageReceivedEventArgs = new("0f99a40c-e962-4207-9e92-e3d542eff849");
    public static readonly Guid Environment = new("b96d755e-0319-4e92-a296-23436f46a1fc");
    public static readonly Guid Controller = new("4d00c0d1-9434-4eb6-8078-8697a560334f");
    public static readonly Guid WebView = new("76eceacb-0462-4d94-ac83-423a6793775e");

    /// <summary>
    /// The content surface two revisions on. A runtime older than the revision that introduced it
    /// answers the query with E_NOINTERFACE, which is the only way this shim has of telling whether
    /// the virtual host mapping and the message channel exist.
    /// </summary>
    public static readonly Guid WebView3 = new("a0d6df20-3b92-416d-aa0c-437a9c727857");
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
