using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.ToolWindow;

// Every interface below is declared flat: where COM defines one interface as deriving from another,
// the base members are repeated here in their original slots rather than expressed as C# interface
// inheritance. This matches how the extensibility interface is declared elsewhere in the shim and
// it has one purpose, which is that the vtable layout is visible in the source. Source-generated
// interop assigns slots in declaration order, the container calls through those slots without any
// negotiation, and a member inserted in the wrong place does not fail - it calls a different
// function with the wrong arguments.
//
// Pointer-shaped parameters stay as nint. Marshalling them would mean declaring types for
// monikers, data objects, advise sinks, and enumerators that this control never uses, and every one
// of those declarations would be another vtable to get wrong.

/// <summary>Interface identifiers, kept in one place because they are passed to QueryInterface by hand.</summary>
internal static class OleIid
{
    public static readonly Guid OleWindow = new("00000114-0000-0000-C000-000000000046");
    public static readonly Guid OleInPlaceObject = new("00000113-0000-0000-C000-000000000046");
    public static readonly Guid OleInPlaceActiveObject = new("00000117-0000-0000-C000-000000000046");
    public static readonly Guid OleInPlaceSite = new("00000119-0000-0000-C000-000000000046");
    public static readonly Guid OleInPlaceUIWindow = new("00000115-0000-0000-C000-000000000046");
}

/// <summary>Verbs a container asks an embedded object to perform.</summary>
internal static class OleVerb
{
    public const int Primary = 0;
    public const int Show = -1;
    public const int Open = -2;
    public const int Hide = -3;
    public const int UIActivate = -4;
    public const int InPlaceActivate = -5;
}

/// <summary>Drawing aspects. Only the content aspect is meaningful for a control.</summary>
internal static class DvAspect
{
    public const int Content = 1;
}

/// <summary>
/// What the container tells an activating object about its frame. The size field is validated by
/// the container, so the caller must set it before the call rather than after.
/// </summary>
[StructLayout(LayoutKind.Sequential)]
internal struct OleInPlaceFrameInfo
{
    public uint Size;
    public int IsMdiApplication;
    public nint FrameWindow;
    public nint AcceleratorTable;
    public uint AcceleratorEntryCount;
}

/// <summary>
/// Common base of the in-place interfaces. A container asks for it directly when it only needs the
/// window handle, so it is declared and implemented separately as well as being repeated inside the
/// interfaces that derive from it.
/// </summary>
[GeneratedComInterface]
[Guid("00000114-0000-0000-C000-000000000046")]
internal partial interface IOleWindow
{
    [PreserveSig]
    int GetWindow(out nint window);

    [PreserveSig]
    int ContextSensitiveHelp(int enterMode);
}

/// <summary>
/// The interface that makes an object embeddable. The container drives the entire lifetime through
/// it: siting, sizing, activation, and shutdown.
/// </summary>
[GeneratedComInterface]
[Guid("00000112-0000-0000-C000-000000000046")]
internal partial interface IOleObject
{
    [PreserveSig]
    int SetClientSite(nint clientSite);

    [PreserveSig]
    int GetClientSite(out nint clientSite);

    [PreserveSig]
    int SetHostNames(nint containerApplication, nint containerObject);

    [PreserveSig]
    int Close(int saveOption);

    [PreserveSig]
    int SetMoniker(int whichMoniker, nint moniker);

    [PreserveSig]
    int GetMoniker(int assign, int whichMoniker, out nint moniker);

    [PreserveSig]
    int InitFromData(nint dataObject, int creation, int reserved);

    [PreserveSig]
    int GetClipboardData(int reserved, out nint dataObject);

    [PreserveSig]
    int DoVerb(int verb, nint message, nint activeSite, int index, nint parentWindow, nint positionRect);

    [PreserveSig]
    int EnumVerbs(out nint enumerator);

    [PreserveSig]
    int Update();

    [PreserveSig]
    int IsUpToDate();

    [PreserveSig]
    int GetUserClassID(out Guid classId);

    [PreserveSig]
    int GetUserType(int formOfType, out nint userType);

    [PreserveSig]
    int SetExtent(int drawAspect, nint size);

    [PreserveSig]
    int GetExtent(int drawAspect, nint size);

    [PreserveSig]
    int Advise(nint adviseSink, out int connection);

    [PreserveSig]
    int Unadvise(int connection);

    [PreserveSig]
    int EnumAdvise(out nint enumerator);

    [PreserveSig]
    int GetMiscStatus(int aspect, out int status);

    [PreserveSig]
    int SetColorScheme(nint logicalPalette);
}

/// <summary>
/// How the container moves, clips, and deactivates an object that is active in place. The first two
/// members are the window interface repeated in its original slots.
/// </summary>
[GeneratedComInterface]
[Guid("00000113-0000-0000-C000-000000000046")]
internal partial interface IOleInPlaceObject
{
    [PreserveSig]
    int GetWindow(out nint window);

    [PreserveSig]
    int ContextSensitiveHelp(int enterMode);

    [PreserveSig]
    int InPlaceDeactivate();

    [PreserveSig]
    int UIDeactivate();

    [PreserveSig]
    int SetObjectRects(nint positionRect, nint clipRect);

    [PreserveSig]
    int ReactivateAndUndo();
}

/// <summary>
/// How the frame talks to whichever embedded object currently owns the user interface: accelerator
/// routing, activation changes, and modal transitions.
/// </summary>
[GeneratedComInterface]
[Guid("00000117-0000-0000-C000-000000000046")]
internal partial interface IOleInPlaceActiveObject
{
    [PreserveSig]
    int GetWindow(out nint window);

    [PreserveSig]
    int ContextSensitiveHelp(int enterMode);

    [PreserveSig]
    int TranslateAccelerator(nint message);

    [PreserveSig]
    int OnFrameWindowActivate(int activate);

    [PreserveSig]
    int OnDocWindowActivate(int activate);

    [PreserveSig]
    int ResizeBorder(nint border, nint uiWindow, int frameWindow);

    [PreserveSig]
    int EnableModeless(int enable);
}

/// <summary>Control-specific notifications: mnemonics, ambient properties, and event freezing.</summary>
[GeneratedComInterface]
[Guid("B196B288-BAB4-101A-B69C-00AA00341D07")]
internal partial interface IOleControl
{
    [PreserveSig]
    int GetControlInfo(nint controlInfo);

    [PreserveSig]
    int OnMnemonic(nint message);

    [PreserveSig]
    int OnAmbientPropertyChange(int dispId);

    [PreserveSig]
    int FreezeEvents(int freeze);
}

/// <summary>
/// Persistence. A control container calls InitNew or Load before activating the control and refuses
/// to proceed if neither succeeds, so this is required even for a control with no state.
/// </summary>
[GeneratedComInterface]
[Guid("7FD52380-4E07-101B-AE2D-08002B2EC713")]
internal partial interface IPersistStreamInit
{
    [PreserveSig]
    int GetClassID(out Guid classId);

    [PreserveSig]
    int IsDirty();

    [PreserveSig]
    int Load(nint stream);

    [PreserveSig]
    int Save(nint stream, int clearDirty);

    [PreserveSig]
    int GetSizeMax(out ulong size);

    [PreserveSig]
    int InitNew();
}

/// <summary>
/// How a container renders an object that is not active. A windowed control never renders through
/// this path, but the container asks for the interface during siting and treats its absence as a
/// sign that the object is not a control.
/// </summary>
[GeneratedComInterface]
[Guid("00000127-0000-0000-C000-000000000046")]
internal partial interface IViewObject2
{
    [PreserveSig]
    int Draw(
        int drawAspect,
        int index,
        nint aspect,
        nint targetDevice,
        nint targetDeviceContext,
        nint drawContext,
        nint bounds,
        nint windowBounds,
        nint continueFunction,
        nint continueParameter);

    [PreserveSig]
    int GetColorSet(int drawAspect, int index, nint aspect, nint targetDevice, nint targetDeviceContext, out nint colorSet);

    [PreserveSig]
    int Freeze(int drawAspect, int index, nint aspect, out int freeze);

    [PreserveSig]
    int Unfreeze(int freeze);

    [PreserveSig]
    int SetAdvise(int aspects, int advf, nint adviseSink);

    [PreserveSig]
    int GetAdvise(nint aspects, nint advf, nint adviseSink);

    [PreserveSig]
    int GetExtent(int drawAspect, int index, nint targetDevice, nint size);
}

/// <summary>
/// Hands out the type information for the coclass. Containers use it to discover the default event
/// interface, and accept a refusal from a control that raises no events.
/// </summary>
[GeneratedComInterface]
[Guid("B196B283-BAB4-101A-B69C-00AA00341D07")]
internal partial interface IProvideClassInfo
{
    [PreserveSig]
    int GetClassInfo(out nint typeInfo);
}

/// <summary>
/// The container's side of the embedding relationship. Called out to, never implemented here.
/// </summary>
[GeneratedComInterface]
[Guid("00000118-0000-0000-C000-000000000046")]
internal partial interface IOleClientSite
{
    [PreserveSig]
    int SaveObject();

    [PreserveSig]
    int GetMoniker(int assign, int whichMoniker, out nint moniker);

    [PreserveSig]
    int GetContainer(out nint container);

    [PreserveSig]
    int ShowObject();

    [PreserveSig]
    int OnShowWindow(int show);

    [PreserveSig]
    int RequestNewObjectLayout();
}

/// <summary>
/// The container's side of in-place activation. This is where the parent window handle and the
/// position and clipping rectangles come from, so it is the interface the whole hosting arrangement
/// depends on.
/// </summary>
[GeneratedComInterface]
[Guid("00000119-0000-0000-C000-000000000046")]
internal partial interface IOleInPlaceSite
{
    [PreserveSig]
    int GetWindow(out nint window);

    [PreserveSig]
    int ContextSensitiveHelp(int enterMode);

    [PreserveSig]
    int CanInPlaceActivate();

    [PreserveSig]
    int OnInPlaceActivate();

    [PreserveSig]
    int OnUIActivate();

    [PreserveSig]
    int GetWindowContext(out nint frame, out nint document, nint positionRect, nint clipRect, nint frameInfo);

    [PreserveSig]
    int Scroll(Interop.Size scrollExtent);

    [PreserveSig]
    int OnUIDeactivate(int undoable);

    [PreserveSig]
    int OnInPlaceDeactivate();

    [PreserveSig]
    int DiscardUndoState();

    [PreserveSig]
    int DeactivateAndUndo();

    [PreserveSig]
    int OnPosRectChange(nint positionRect);
}

/// <summary>
/// The border-negotiation half of the frame and document windows. Only SetActiveObject is used, to
/// tell the frame which object owns the user interface. The frame interface derives from this one,
/// so the same pointer answers to both identifiers and the member sits at the same slot in each.
/// </summary>
[GeneratedComInterface]
[Guid("00000115-0000-0000-C000-000000000046")]
internal partial interface IOleInPlaceUIWindow
{
    [PreserveSig]
    int GetWindow(out nint window);

    [PreserveSig]
    int ContextSensitiveHelp(int enterMode);

    [PreserveSig]
    int GetBorder(nint border);

    [PreserveSig]
    int RequestBorderSpace(nint borderWidths);

    [PreserveSig]
    int SetBorderSpace(nint borderWidths);

    [PreserveSig]
    int SetActiveObject(nint activeObject, nint objectName);
}
