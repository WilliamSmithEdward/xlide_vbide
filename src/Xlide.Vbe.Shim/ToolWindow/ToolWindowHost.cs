using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Core.Registration;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.UI;
using Xlide.Vbe.Shim.WebView;

namespace Xlide.Vbe.Shim.ToolWindow;

/// <summary>
/// The object the editor sites inside a docked tool window, and the thing that owns the browser
/// surface the developer actually sees.
///
/// A control container negotiates entirely through interfaces. It creates the class, asks it for
/// every interface it expects a control to have, and abandons the whole arrangement if any one of
/// them is missing. Source-generated interop exposes exactly what a class declares, so the list on
/// this class is the contract: an interface left out is not a missing feature, it is a tool window
/// that never appears and no diagnostic anywhere. That is the same failure the add-in class hit
/// with IDispatch, which is why IDispatch is on this class too even though it answers nothing.
///
/// Everything here runs on the host user interface thread.
/// </summary>
[GeneratedComClass]
internal sealed unsafe partial class ToolWindowHost :
    IOleObject,
    IOleInPlaceObject,
    IOleInPlaceActiveObject,
    IOleWindow,
    IOleControl,
    IPersistStreamInit,
    IViewObject2,
    IProvideClassInfo,
    IDispatch,
    IDisposable
{
    /// <summary>HIMETRIC units per pixel at the reference density of 96 dots per inch.</summary>
    private const int HiMetricPerInch = 2540;
    private const int PixelsPerInch = 96;

    private static readonly Guid ClassId = new(ProductIdentity.ToolWindowHostClsid);

    private ComHandle<IOleClientSite>? _clientSite;
    private ComHandle<IOleInPlaceSite>? _inPlaceSite;
    private HostWindow? _window;
    private WebView2Surface? _surface;

    private SizeL _extent = new() { Width = ToHiMetric(640), Height = ToHiMetric(400) };
    private bool _inPlaceActive;
    private bool _uiActive;

    public ToolWindowHost() => Log.Info("tool window host: created");

    // ---- IOleObject ------------------------------------------------------------------------

    public int SetClientSite(nint clientSite)
    {
        Log.Info($"tool window host: SetClientSite 0x{clientSite:X}");

        _inPlaceSite?.Dispose();
        _inPlaceSite = null;
        _clientSite?.Dispose();
        _clientSite = ComHandle<IOleClientSite>.Borrow(clientSite);

        return HResult.Ok;
    }

    public int GetClientSite(out nint clientSite)
    {
        var site = _clientSite;
        if (site is null)
        {
            clientSite = 0;
            return HResult.Fail;
        }

        // The caller owns what it receives, so the reference count goes up on the way out.
        Marshal.AddRef(site.Pointer);
        clientSite = site.Pointer;
        return HResult.Ok;
    }

    public int SetHostNames(nint containerApplication, nint containerObject) => HResult.Ok;

    public int Close(int saveOption)
    {
        Log.Info("tool window host: Close");
        InPlaceDeactivate();
        SetClientSite(0);
        return HResult.Ok;
    }

    public int SetMoniker(int whichMoniker, nint moniker) => HResult.Ok;

    public int GetMoniker(int assign, int whichMoniker, out nint moniker)
    {
        moniker = 0;
        return HResult.NotImplemented;
    }

    public int InitFromData(nint dataObject, int creation, int reserved) => HResult.NotImplemented;

    public int GetClipboardData(int reserved, out nint dataObject)
    {
        dataObject = 0;
        return HResult.NotImplemented;
    }

    public int DoVerb(int verb, nint message, nint activeSite, int index, nint parentWindow, nint positionRect)
    {
        Log.Info($"tool window host: DoVerb {verb}");

        switch (verb)
        {
            case OleVerb.Show:
            case OleVerb.InPlaceActivate:
                return InPlaceActivate(parentWindow, positionRect, uiActivate: false);

            case OleVerb.Primary:
            case OleVerb.UIActivate:
                return InPlaceActivate(parentWindow, positionRect, uiActivate: true);

            case OleVerb.Hide:
                UIDeactivate();
                _window?.SetBounds(default);
                return HResult.Ok;

            case OleVerb.Open:
                // There is no separate window for this control, and opening one is not a fallback
                // worth having.
                return HResult.NotImplemented;

            default:
                return HResult.OleInvalidVerb;
        }
    }

    public int EnumVerbs(out nint enumerator)
    {
        enumerator = 0;

        // A container that gets this asks the registry for the verb list instead, which is the
        // behaviour we want: this control has no verbs to offer a user.
        return HResult.OleUseRegistry;
    }

    public int Update() => HResult.Ok;

    public int IsUpToDate() => HResult.Ok;

    public int GetUserClassID(out Guid classId)
    {
        classId = ClassId;
        return HResult.Ok;
    }

    public int GetUserType(int formOfType, out nint userType)
    {
        userType = 0;
        return HResult.OleUseRegistry;
    }

    public int SetExtent(int drawAspect, nint size)
    {
        if ((drawAspect & DvAspect.Content) == 0 || size == 0)
        {
            return HResult.InvalidArg;
        }

        _extent = *(SizeL*)size;
        return HResult.Ok;
    }

    public int GetExtent(int drawAspect, nint size)
    {
        if ((drawAspect & DvAspect.Content) == 0 || size == 0)
        {
            return HResult.InvalidArg;
        }

        *(SizeL*)size = _extent;
        return HResult.Ok;
    }

    public int Advise(nint adviseSink, out int connection)
    {
        // Advisory notifications describe changes to a document's data. This control has none, so
        // there is nothing a sink could ever be told. Accepting the sink and never calling it would
        // be worse than refusing.
        connection = 0;
        return HResult.OleAdviseNotSupported;
    }

    public int Unadvise(int connection) => HResult.OleNoConnection;

    public int EnumAdvise(out nint enumerator)
    {
        enumerator = 0;
        return HResult.OleAdviseNotSupported;
    }

    public int GetMiscStatus(int aspect, out int status)
    {
        // The same value the class registration advertises. A container may read either, and a
        // disagreement between them is a control that is sited but never activated.
        status = ControlMiscStatus.ToolWindowHost;
        return HResult.Ok;
    }

    public int SetColorScheme(nint logicalPalette) => HResult.NotImplemented;

    // ---- IOleWindow, IOleInPlaceObject, IOleInPlaceActiveObject -----------------------------

    public int GetWindow(out nint window)
    {
        window = _window?.Handle ?? 0;
        return window == 0 ? HResult.Fail : HResult.Ok;
    }

    public int ContextSensitiveHelp(int enterMode) => HResult.Ok;

    public int InPlaceDeactivate()
    {
        if (!_inPlaceActive)
        {
            return HResult.Ok;
        }

        Log.Info("tool window host: InPlaceDeactivate");

        UIDeactivate();
        DestroySurface();

        _inPlaceActive = false;
        _inPlaceSite?.Target.OnInPlaceDeactivate();
        _inPlaceSite?.Dispose();
        _inPlaceSite = null;

        return HResult.Ok;
    }

    public int UIDeactivate()
    {
        if (!_uiActive)
        {
            return HResult.Ok;
        }

        _uiActive = false;
        _inPlaceSite?.Target.OnUIDeactivate(0);
        return HResult.Ok;
    }

    public int SetObjectRects(nint positionRect, nint clipRect)
    {
        if (positionRect == 0)
        {
            return HResult.InvalidArg;
        }

        var position = ToPixelRect((Rect*)positionRect);
        if (clipRect != 0)
        {
            // The container is entitled to clip us to less than the position it gave. Sizing to the
            // position alone produces a control that draws over its neighbours when docked panes
            // overlap during a drag.
            position = position.Intersect(ToPixelRect((Rect*)clipRect));
        }

        _window?.SetBounds(position);
        return HResult.Ok;
    }

    public int ReactivateAndUndo() => HResult.NotImplemented;

    public int TranslateAccelerator(nint message) => HResult.False;

    public int OnFrameWindowActivate(int activate) => HResult.Ok;

    public int OnDocWindowActivate(int activate) => HResult.Ok;

    public int ResizeBorder(nint border, nint uiWindow, int frameWindow) => HResult.Ok;

    public int EnableModeless(int enable) => HResult.Ok;

    // ---- IOleControl -----------------------------------------------------------------------

    public int GetControlInfo(nint controlInfo) => HResult.NotImplemented;

    public int OnMnemonic(nint message) => HResult.NotImplemented;

    public int OnAmbientPropertyChange(int dispId) => HResult.Ok;

    public int FreezeEvents(int freeze) => HResult.Ok;

    // ---- IPersistStreamInit ----------------------------------------------------------------

    public int GetClassID(out Guid classId)
    {
        classId = ClassId;
        return HResult.Ok;
    }

    public int IsDirty() => HResult.False;

    public int Load(nint stream) => HResult.Ok;

    public int Save(nint stream, int clearDirty) => HResult.Ok;

    public int GetSizeMax(out ulong size)
    {
        size = 0;
        return HResult.Ok;
    }

    public int InitNew() => HResult.Ok;

    // ---- IViewObject2 ----------------------------------------------------------------------

    public int Draw(
        int drawAspect,
        int index,
        nint aspect,
        nint targetDevice,
        nint targetDeviceContext,
        nint drawContext,
        nint bounds,
        nint windowBounds,
        nint continueFunction,
        nint continueParameter)
    {
        // The control has its own window and the browser paints into it. There is no metafile
        // representation to produce, and reporting failure here makes containers treat the control
        // as broken rather than as windowed.
        return HResult.Ok;
    }

    public int GetColorSet(int drawAspect, int index, nint aspect, nint targetDevice, nint targetDeviceContext, out nint colorSet)
    {
        colorSet = 0;
        return HResult.NotImplemented;
    }

    public int Freeze(int drawAspect, int index, nint aspect, out int freeze)
    {
        freeze = 0;
        return HResult.NotImplemented;
    }

    public int Unfreeze(int freeze) => HResult.NotImplemented;

    public int SetAdvise(int aspects, int advf, nint adviseSink) => HResult.Ok;

    public int GetAdvise(nint aspects, nint advf, nint adviseSink)
    {
        if (aspects != 0)
        {
            *(int*)aspects = DvAspect.Content;
        }

        if (advf != 0)
        {
            *(int*)advf = 0;
        }

        if (adviseSink != 0)
        {
            *(nint*)adviseSink = 0;
        }

        return HResult.Ok;
    }

    public int GetExtent(int drawAspect, int index, nint targetDevice, nint size)
    {
        if ((drawAspect & DvAspect.Content) == 0 || size == 0)
        {
            return HResult.InvalidArg;
        }

        *(SizeL*)size = _extent;
        return HResult.Ok;
    }

    // ---- IProvideClassInfo -----------------------------------------------------------------

    public int GetClassInfo(out nint typeInfo)
    {
        typeInfo = 0;

        // The shim ships no type library. Containers use class information to find a default event
        // interface; this control raises no events, and a refusal is the honest answer.
        return HResult.NotImplemented;
    }

    // ---- IDispatch -------------------------------------------------------------------------

    public int GetTypeInfoCount(out uint count)
    {
        count = 0;
        return HResult.Ok;
    }

    public int GetTypeInfo(uint typeInfoIndex, uint lcid, out nint typeInfo)
    {
        typeInfo = 0;
        return HResult.NotImplemented;
    }

    public int GetIDsOfNames(in Guid riid, nint names, uint nameCount, uint lcid, nint dispIds)
    {
        if (dispIds == 0 || nameCount == 0)
        {
            return HResult.InvalidArg;
        }

        var results = (int*)dispIds;
        for (var i = 0u; i < nameCount; i++)
        {
            results[i] = DispId.Unknown;
        }

        return HResult.DispUnknownName;
    }

    public int Invoke(
        int dispIdMember,
        in Guid riid,
        uint lcid,
        ushort flags,
        nint dispParams,
        nint result,
        nint exceptionInfo,
        nint argumentError) => HResult.DispMemberNotFound;

    // ---- activation --------------------------------------------------------------------------

    private int InPlaceActivate(nint fallbackParent, nint positionRect, bool uiActivate)
    {
        var clientSite = _clientSite;
        if (clientSite is null)
        {
            Log.Error("tool window host: activation requested before the container supplied a site");
            return HResult.Fail;
        }

        _inPlaceSite ??= clientSite.As<IOleInPlaceSite>(OleIid.OleInPlaceSite);
        if (_inPlaceSite is null)
        {
            Log.Error("tool window host: the container site does not support in-place activation");
            return HResult.Fail;
        }

        var site = _inPlaceSite.Target;

        if (!_inPlaceActive)
        {
            if (site.CanInPlaceActivate() != HResult.Ok)
            {
                Log.Error("tool window host: the container declined in-place activation");
                return HResult.Fail;
            }

            site.OnInPlaceActivate();
            _inPlaceActive = true;
        }

        var position = default(Rect);
        var clip = default(Rect);
        var frameInfo = default(OleInPlaceFrameInfo);
        frameInfo.Size = (uint)sizeof(OleInPlaceFrameInfo);

        nint frame = 0;
        nint document = 0;
        var haveContext = site.GetWindowContext(
            out frame,
            out document,
            (nint)(&position),
            (nint)(&clip),
            (nint)(&frameInfo)) >= 0;

        try
        {
            var bounds = positionRect != 0
                ? ToPixelRect((Rect*)positionRect)
                : haveContext ? ToPixelRect(&position) : PixelRect.FromSize(0, 0, ToPixels(_extent.Width), ToPixels(_extent.Height));

            if (haveContext && clip.Right > clip.Left && clip.Bottom > clip.Top)
            {
                bounds = bounds.Intersect(ToPixelRect(&clip));
            }

            nint parent = 0;
            if (site.GetWindow(out var siteWindow) >= 0)
            {
                parent = siteWindow;
            }

            if (parent == 0)
            {
                parent = fallbackParent;
            }

            if (!EnsureWindow(parent, bounds))
            {
                return HResult.Fail;
            }

            if (uiActivate && !_uiActive)
            {
                _uiActive = true;
                site.OnUIActivate();

                if (haveContext)
                {
                    AnnounceActiveObject(frame);
                    AnnounceActiveObject(document);
                }
            }

            clientSite.Target.ShowObject();
            return HResult.Ok;
        }
        finally
        {
            // Only a successful call promises anything about these two. A failed one may leave the
            // locals holding whatever they held before, so releasing them then would be releasing
            // something that was never ours.
            if (haveContext)
            {
                if (frame != 0)
                {
                    Marshal.Release(frame);
                }

                if (document != 0)
                {
                    Marshal.Release(document);
                }
            }
        }
    }

    /// <summary>
    /// Tells the frame or document window which object currently owns the user interface, so that
    /// border space negotiation and accelerator routing reach us rather than the previous holder.
    /// </summary>
    private void AnnounceActiveObject(nint uiWindow)
    {
        if (uiWindow == 0)
        {
            return;
        }

        using var window = ComHandle<IOleInPlaceUIWindow>.Borrow(uiWindow);
        if (window is null)
        {
            return;
        }

        var self = QueryOwnInterface(OleIid.OleInPlaceActiveObject);
        if (self == 0)
        {
            return;
        }

        try
        {
            window.Target.SetActiveObject(self, 0);
        }
        finally
        {
            // The window took its own reference if it kept us.
            Marshal.Release(self);
        }
    }

    private bool EnsureWindow(nint parent, PixelRect bounds)
    {
        if (_window is not null)
        {
            _window.SetBounds(bounds);
            return true;
        }

        var window = HostWindow.Create(parent, bounds);
        if (window is null)
        {
            return false;
        }

        _window = window;
        window.Resized = OnWindowResized;
        window.FocusReceived = OnWindowFocused;
        window.Destroying = OnWindowDestroying;

        _surface = WebView2Surface.Start(window.Handle, window.ClientBounds());

        // Register with the bus so analysis can reach whatever panel is showing. The two have
        // unrelated lifetimes, so neither holds a reference to the other.
        if (_surface is not null)
        {
            PanelBus.Attach(_surface);
        }
        if (_surface is null)
        {
            Log.Error("tool window host: the browser surface could not be started");
        }

        return true;
    }

    private void OnWindowResized(PixelRect clientBounds) => _surface?.SetBounds(clientBounds);

    private void OnWindowFocused() => _surface?.Focus();

    private void OnWindowDestroying()
    {
        DetachSurface();
    }

    private void DestroySurface()
    {
        // Disposing the window drives WM_DESTROY, which runs the callback above. Doing the browser
        // first means that path finds nothing left to do rather than racing it.
        DetachSurface();

        _window?.Dispose();
        _window = null;
    }

    /// <summary>
    /// Takes the surface off the bus and shuts it down.
    ///
    /// Unregistering comes first. The bus holds the surface to post to and would otherwise be left
    /// pointing at a disposed one, which is a use-after-free reachable from a background analysis
    /// pass rather than from anything happening here.
    /// </summary>
    private void DetachSurface()
    {
        var surface = _surface;
        _surface = null;

        if (surface is null)
        {
            return;
        }

        PanelBus.Detach(surface);
        surface.Dispose();
    }

    /// <summary>
    /// Produces a pointer to one of our own interfaces. It goes through the shared wrapper table,
    /// so the container sees the identity it already has rather than a second object.
    /// </summary>
    private nint QueryOwnInterface(in Guid interfaceId)
    {
        var unknown = ComRuntime.Wrappers.GetOrCreateComInterfaceForObject(this, CreateComInterfaceFlags.None);
        try
        {
            return Marshal.QueryInterface(unknown, in interfaceId, out var result) < 0 ? 0 : result;
        }
        finally
        {
            Marshal.Release(unknown);
        }
    }

    private static PixelRect ToPixelRect(Rect* rect) => new(rect->Left, rect->Top, rect->Right, rect->Bottom);

    private static int ToHiMetric(int pixels) => (int)((long)pixels * HiMetricPerInch / PixelsPerInch);

    private static int ToPixels(int hiMetric) => (int)((long)hiMetric * PixelsPerInch / HiMetricPerInch);

    /// <summary>
    /// Safety net for the case where the container releases the control without closing it. COM
    /// owns this object's lifetime, so the normal path is Close followed by SetClientSite(null).
    /// </summary>
    public void Dispose()
    {
        DestroySurface();

        _inPlaceSite?.Dispose();
        _inPlaceSite = null;

        _clientSite?.Dispose();
        _clientSite = null;
    }
}
