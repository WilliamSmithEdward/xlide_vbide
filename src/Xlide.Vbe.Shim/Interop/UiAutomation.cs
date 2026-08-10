using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xlide.Vbe.Shim.Interop;

/*
 * Just enough of the accessibility API to read a control that draws its own text.
 *
 * The editor's Immediate window is custom drawn: it has no child edit control, and asking the
 * window for its text answers with its caption. The only thing that can read it is the interface a
 * screen reader would use, and that does read it, including while the window is hidden.
 *
 * Every interface here is declared flat, with a placeholder for each method that is not used, so
 * the slots line up with the real vtable. A missing placeholder does not fail to compile and does
 * not throw: it calls whatever member happens to occupy that slot, with the wrong arguments. The
 * counts in the comments are what keeps that honest.
 */

/// <summary>IUIAutomation. Only element lookup and a true condition are used.</summary>
[GeneratedComInterface]
[Guid("30cbe57d-d9d0-452a-ab13-7ac5ac4825ee")]
internal partial interface IUIAutomation
{
    // 1..3
    [PreserveSig] int CompareElements(nint first, nint second, out int areSame);
    [PreserveSig] int CompareRuntimeIds(nint first, nint second, out int areSame);
    [PreserveSig] int GetRootElement(out nint root);

    // 4
    [PreserveSig] int ElementFromHandle(nint window, out nint element);

    // 5..18
    [PreserveSig] int ElementFromPoint(long point, out nint element);
    [PreserveSig] int GetFocusedElement(out nint element);
    [PreserveSig] int GetRootElementBuildCache(nint request, out nint element);
    [PreserveSig] int ElementFromHandleBuildCache(nint window, nint request, out nint element);
    [PreserveSig] int ElementFromPointBuildCache(long point, nint request, out nint element);
    [PreserveSig] int GetFocusedElementBuildCache(nint request, out nint element);
    [PreserveSig] int CreateTreeWalker(nint condition, out nint walker);
    [PreserveSig] int GetControlViewWalker(out nint walker);
    [PreserveSig] int GetContentViewWalker(out nint walker);
    [PreserveSig] int GetRawViewWalker(out nint walker);
    [PreserveSig] int GetRawViewCondition(out nint condition);
    [PreserveSig] int GetControlViewCondition(out nint condition);
    [PreserveSig] int GetContentViewCondition(out nint condition);
    [PreserveSig] int CreateCacheRequest(out nint request);

    // 19
    [PreserveSig] int CreateTrueCondition(out nint condition);
}

/// <summary>IUIAutomationElement. Finding a descendant, and asking it for a pattern.</summary>
[GeneratedComInterface]
[Guid("d22108aa-8ac5-49a5-837b-37bbb3d7591e")]
internal partial interface IUIAutomationElement
{
    // 1..2
    [PreserveSig] int SetFocus();
    [PreserveSig] int GetRuntimeId(out nint runtimeId);

    // 3
    [PreserveSig] int FindFirst(int scope, nint condition, out nint found);

    // 4..13
    [PreserveSig] int FindAll(int scope, nint condition, out nint found);
    [PreserveSig] int FindFirstBuildCache(int scope, nint condition, nint request, out nint found);
    [PreserveSig] int FindAllBuildCache(int scope, nint condition, nint request, out nint found);
    [PreserveSig] int BuildUpdatedCache(nint request, out nint updated);
    [PreserveSig] int GetCurrentPropertyValue(int propertyId, out UiVariant value);
    [PreserveSig] int GetCurrentPropertyValueEx(int propertyId, int ignoreDefault, out ComVariantBlock value);
    [PreserveSig] int GetCachedPropertyValue(int propertyId, out ComVariantBlock value);
    [PreserveSig] int GetCachedPropertyValueEx(int propertyId, int ignoreDefault, out ComVariantBlock value);
    [PreserveSig] int GetCurrentPatternAs(int patternId, in Guid interfaceId, out nint pattern);
    [PreserveSig] int GetCachedPatternAs(int patternId, in Guid interfaceId, out nint pattern);

    // 14
    [PreserveSig] int GetCurrentPattern(int patternId, out nint pattern);
}

/// <summary>IUIAutomationTextPattern. Only the whole-document range is needed.</summary>
[GeneratedComInterface]
[Guid("32eba289-3583-42c9-9c59-3b6d9a1e9b6a")]
internal partial interface IUIAutomationTextPattern
{
    // 1..4
    [PreserveSig] int RangeFromPoint(long point, out nint range);
    [PreserveSig] int RangeFromChild(nint child, out nint range);
    [PreserveSig] int GetSelection(out nint ranges);
    [PreserveSig] int GetVisibleRanges(out nint ranges);

    // 5
    [PreserveSig] int GetDocumentRange(out nint range);
}

/// <summary>IUIAutomationTextRange. Only the text is needed.</summary>
[GeneratedComInterface]
[Guid("a543cc6a-f4ae-494b-8239-c814481187a8")]
internal partial interface IUIAutomationTextRange
{
    // 1..9
    [PreserveSig] int Clone(out nint range);
    [PreserveSig] int Compare(nint other, out int areSame);
    [PreserveSig] int CompareEndpoints(int endpoint, nint other, int otherEndpoint, out int result);
    [PreserveSig] int ExpandToEnclosingUnit(int unit);
    [PreserveSig] int FindAttribute(int attribute, ComVariantBlock value, int backward, out nint found);
    [PreserveSig] int FindText(nint text, int backward, int ignoreCase, out nint found);
    [PreserveSig] int GetAttributeValue(int attribute, out ComVariantBlock value);
    [PreserveSig] int GetBoundingRectangles(out nint rectangles);
    [PreserveSig] int GetEnclosingElement(out nint element);

    // 10
    [PreserveSig] int GetText(int maxLength, out nint text);
}

/// <summary>IUIAutomationElementArray. What FindAll answers with.</summary>
[GeneratedComInterface]
[Guid("14314595-b4bc-4055-95f2-58f2e42c9855")]
internal partial interface IUIAutomationElementArray
{
    // 1..2
    [PreserveSig] int GetLength(out int length);
    [PreserveSig] int GetElement(int index, out nint element);
}

/// <summary>
/// A variant sized for the vtable, never read.
///
/// Several methods above return one. They are placeholders, so this only has to be the right
/// size - and the right size is TWENTY-FOUR bytes on x64: eight of type tag and padding, then a
/// sixteen-byte data union (its widest member is a record's two pointers). Sixteen was the x86
/// size, and an out-parameter eight bytes short is not harmless even on a placeholder: the
/// callee initialises the whole variant, so the last eight bytes land on whatever the caller
/// keeps beside the buffer. That is the corruption that took the Locals reader down for a whole
/// day (2026-08-05): Release stack layouts had left the overhang on dead space, the Debug
/// builds the dev loop switched to that morning put a live slot there, and every property read
/// died in a frameless NullReferenceException while the same reads from outside the process -
/// made with a full-size variant - worked all along.
/// </summary>
[StructLayout(LayoutKind.Sequential, Size = 24)]
internal struct ComVariantBlock;

/// <summary>
/// A variant that is actually read: the twenty-four bytes above with their real shape. The
/// first word is the type tag; the data starts at offset eight, which is where a BSTR pointer
/// or a 32-bit integer sits. Only the two types the property reads produce are handled.
/// </summary>
[StructLayout(LayoutKind.Sequential, Size = 24)]
internal struct UiVariant
{
    public ushort Type;
    public ushort Reserved1;
    public ushort Reserved2;
    public ushort Reserved3;
    public nint Value;

    public const ushort TypeBstr = 8;
    public const ushort TypeInt32 = 3;

    /// <summary>The integer inside, or zero when the variant holds something else.</summary>
    public readonly int AsInt32() => Type == TypeInt32 ? (int)Value : 0;

    /// <summary>The string inside, freed after reading, or null for anything else.</summary>
    public string? TakeString()
    {
        if (Type != TypeBstr || Value == 0)
        {
            return null;
        }

        var text = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(Value);
        System.Runtime.InteropServices.Marshal.FreeBSTR(Value);
        Value = 0;
        return text;
    }
}

/// <summary>Identifiers used with the interfaces above.</summary>
internal static class UiAutomationIds
{
    /// <summary>CLSID_CUIAutomation.</summary>
    public static readonly Guid AutomationClass = new("ff48dba4-60ef-4201-aa87-54103eef594e");

    public static readonly Guid Automation = new("30cbe57d-d9d0-452a-ab13-7ac5ac4825ee");
    public static readonly Guid Element = new("d22108aa-8ac5-49a5-837b-37bbb3d7591e");
    public static readonly Guid ElementArray = new("14314595-b4bc-4055-95f2-58f2e42c9855");
    public static readonly Guid TextPattern = new("32eba289-3583-42c9-9c59-3b6d9a1e9b6a");
    public static readonly Guid TextRange = new("a543cc6a-f4ae-494b-8239-c814481187a8");

    /// <summary>UIA_TextPatternId.</summary>
    public const int TextPatternId = 10014;

    /// <summary>TreeScope_Descendants.</summary>
    public const int Descendants = 4;

    /// <summary>The whole range, however long it is.</summary>
    public const int WholeRange = -1;

    /// <summary>UIA_NamePropertyId.</summary>
    public const int NameProperty = 30005;

    /// <summary>UIA_ControlTypePropertyId.</summary>
    public const int ControlTypeProperty = 30003;

    /// <summary>UIA_ListItemControlTypeId: the rows of the Locals window.</summary>
    public const int ListItemControl = 50007;

    /// <summary>UIA_EditControlTypeId: the context strip naming the broken procedure.</summary>
    public const int EditControl = 50004;

    /// <summary>
    /// UIA_PaneControlTypeId: what the Locals window's context box actually reads as - a bare
    /// child window with no richer role (measured 2026-08-05).
    /// </summary>
    public const int PaneControl = 50033;
}
