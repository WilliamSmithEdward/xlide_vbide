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
    [PreserveSig] int GetCachedPropertyValue(int propertyId, out UiVariant value);
    [PreserveSig] int GetCachedPropertyValueEx(int propertyId, int ignoreDefault, out ComVariantBlock value);
    [PreserveSig] int GetCurrentPatternAs(int patternId, in Guid interfaceId, out nint pattern);
    [PreserveSig] int GetCachedPatternAs(int patternId, in Guid interfaceId, out nint pattern);

    // 14
    [PreserveSig] int GetCurrentPattern(int patternId, out nint pattern);

    // 15..17
    [PreserveSig] int GetCachedPattern(int patternId, out nint pattern);
    [PreserveSig] int GetCachedParent(out nint parent);
    [PreserveSig] int GetCachedChildren(out nint children);
}

/// <summary>
/// IUIAutomationCacheRequest. Fetches a whole subtree's properties in one call, so a reading
/// of a long Locals tree crosses to the host thread once instead of once per property.
/// </summary>
[GeneratedComInterface]
[Guid("b32a92b5-bc25-4078-9c08-d7ee95c48e03")]
internal partial interface IUIAutomationCacheRequest
{
    // 1..9
    [PreserveSig] int AddProperty(int propertyId);
    [PreserveSig] int AddPattern(int patternId);
    [PreserveSig] int Clone(out nint cloned);
    [PreserveSig] int GetTreeScope(out int scope);
    [PreserveSig] int SetTreeScope(int scope);
    [PreserveSig] int GetTreeFilter(out nint filter);
    [PreserveSig] int SetTreeFilter(nint filter);
    [PreserveSig] int GetAutomationElementMode(out int mode);
    [PreserveSig] int SetAutomationElementMode(int mode);
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

    // 11..14. Selection places the Immediate caret without global keyboard state.
    [PreserveSig] int Move(int unit, int count, out int moved);
    [PreserveSig] int MoveEndpointByUnit(int endpoint, int unit, int count, out int moved);
    [PreserveSig] int MoveEndpointByRange(int endpoint, nint range, int targetEndpoint);
    [PreserveSig] int Select();
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
/// IUIAutomationTreeWalker. The parent link, which the Locals provider gets right, and the next
/// row, each with the asked properties fetched in the same call.
/// </summary>
[GeneratedComInterface]
[Guid("4042c624-389c-4afc-a630-9df854a541fc")]
internal partial interface IUIAutomationTreeWalker
{
    // 1..10
    [PreserveSig] int GetParentElement(nint element, out nint parent);
    [PreserveSig] int GetFirstChildElement(nint element, out nint first);
    [PreserveSig] int GetLastChildElement(nint element, out nint last);
    [PreserveSig] int GetNextSiblingElement(nint element, out nint next);
    [PreserveSig] int GetPreviousSiblingElement(nint element, out nint previous);
    [PreserveSig] int NormalizeElement(nint element, out nint normalized);
    [PreserveSig] int GetParentElementBuildCache(nint element, nint request, out nint parent);
    [PreserveSig] int GetFirstChildElementBuildCache(nint element, nint request, out nint first);
    [PreserveSig] int GetLastChildElementBuildCache(nint element, nint request, out nint last);
    [PreserveSig] int GetNextSiblingElementBuildCache(nint element, nint request, out nint next);
}

/// <summary>IUIAutomationSelectionPattern. Which row a list has selected.</summary>
[GeneratedComInterface]
[Guid("5ed5202e-b2ac-47a6-b638-4b0bf140d78e")]
internal partial interface IUIAutomationSelectionPattern
{
    // 1
    [PreserveSig] int GetCurrentSelection(out nint elements);
}

/// <summary>IUIAutomationSelectionItemPattern. Selecting a row, which also scrolls it into view.</summary>
[GeneratedComInterface]
[Guid("a8efa66a-0fda-421a-9194-38021f3578ea")]
internal partial interface IUIAutomationSelectionItemPattern
{
    // 1
    [PreserveSig] int Select();
}

/// <summary>IUIAutomationScrollPattern. Moving a list by a line, and putting it back.</summary>
[GeneratedComInterface]
[Guid("88f4d42a-e881-459d-a77c-73bbbb7e02dc")]
internal partial interface IUIAutomationScrollPattern
{
    // 1..4
    [PreserveSig] int Scroll(int horizontalAmount, int verticalAmount);
    [PreserveSig] int SetScrollPercent(double horizontalPercent, double verticalPercent);
    [PreserveSig] int GetCurrentHorizontalScrollPercent(out double percent);
    [PreserveSig] int GetCurrentVerticalScrollPercent(out double percent);
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
    public const ushort TypeInt32Array = 0x2003;
    public const ushort TypeDoubleArray = 0x2005;

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

    /// <summary>UIA_BoundingRectanglePropertyId is a SAFEARRAY of four doubles.</summary>
    public (double X, double Y, double Width, double Height)? TakeRectangle()
    {
        if (Type != TypeDoubleArray || Value == 0)
        {
            return null;
        }

        try
        {
            var values = new double[4];
            for (var index = 0; index < values.Length; index++)
            {
                if (SafeArrayGetElement(Value, in index, out values[index]) < 0)
                {
                    return null;
                }
            }

            return (values[0], values[1], values[2], values[3]);
        }
        finally
        {
            _ = SafeArrayDestroy(Value);
            Value = 0;
        }
    }

    /// <summary>
    /// UIA_RuntimeIdPropertyId is a SAFEARRAY of integers; answered joined with dots, the array
    /// freed, or null for anything else.
    /// </summary>
    public string? TakeRuntimeId()
    {
        if (Type != TypeInt32Array || Value == 0)
        {
            return null;
        }

        try
        {
            if (SafeArrayGetLBound(Value, 1, out var lower) < 0 || SafeArrayGetUBound(Value, 1, out var upper) < 0)
            {
                return null;
            }

            var parts = new string[Math.Max(0, upper - lower + 1)];
            for (var index = lower; index <= upper; index++)
            {
                if (SafeArrayGetElement(Value, in index, out int part) < 0)
                {
                    return null;
                }

                parts[index - lower] = part.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            return string.Join('.', parts);
        }
        finally
        {
            _ = SafeArrayDestroy(Value);
            Value = 0;
        }
    }

    [DllImport("oleaut32.dll")]
    private static extern int SafeArrayGetElement(nint array, in int index, out double value);

    [DllImport("oleaut32.dll")]
    private static extern int SafeArrayGetElement(nint array, in int index, out int value);

    [DllImport("oleaut32.dll")]
    private static extern int SafeArrayGetLBound(nint array, uint dimension, out int bound);

    [DllImport("oleaut32.dll")]
    private static extern int SafeArrayGetUBound(nint array, uint dimension, out int bound);

    [DllImport("oleaut32.dll")]
    private static extern int SafeArrayDestroy(nint array);
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

    /// <summary>TreeScope_Subtree: the element, its children and all their descendants.</summary>
    public const int Subtree = 7;

    /// <summary>TreeScope_Element: the element alone.</summary>
    public const int ElementScope = 1;

    /// <summary>AutomationElementMode_None: cached values only, no live element behind them.</summary>
    public const int ElementModeNone = 0;
    public const int ElementModeFull = 1;

    /// <summary>The selection, selection item and scroll patterns, and their interfaces.</summary>
    public const int SelectionPatternId = 10001;
    public const int ScrollPatternId = 10004;
    public const int SelectionItemPatternId = 10010;
    public static readonly Guid SelectionPattern = new("5ed5202e-b2ac-47a6-b638-4b0bf140d78e");
    public static readonly Guid SelectionItemPattern = new("a8efa66a-0fda-421a-9194-38021f3578ea");
    public static readonly Guid ScrollPattern = new("88f4d42a-e881-459d-a77c-73bbbb7e02dc");

    /// <summary>ScrollAmount_NoAmount and ScrollAmount_SmallIncrement; UIA_ScrollPatternNoScroll.</summary>
    public const int ScrollNoAmount = 2;
    public const int ScrollSmallIncrement = 4;
    public const double NoScroll = -1;

    /// <summary>The whole range, however long it is.</summary>
    public const int WholeRange = -1;

    /// <summary>UIA_NamePropertyId.</summary>
    public const int NameProperty = 30005;
    public const int RuntimeIdProperty = 30000;
    public const int BoundingRectangleProperty = 30001;
    public const int NativeWindowHandleProperty = 30020;

    /// <summary>UIA_ExpandCollapsePatternId, its state property, and the state's Expanded value.</summary>
    public const int ExpandCollapsePatternId = 10005;
    public const int ExpandCollapseStateProperty = 30070;
    public const int Expanded = 1;

    /// <summary>UIA_ControlTypePropertyId.</summary>
    public const int ControlTypeProperty = 30003;

    /// <summary>UIA_ListItemControlTypeId: the rows of the Locals window.</summary>
    public const int ListItemControl = 50007;
    public const int ListControl = 50008;

    /// <summary>UIA_EditControlTypeId: the context strip naming the broken procedure.</summary>
    public const int EditControl = 50004;

    /// <summary>
    /// UIA_PaneControlTypeId: what the Locals window's context box actually reads as - a bare
    /// child window with no richer role (measured 2026-08-05).
    /// </summary>
    public const int PaneControl = 50033;
}
