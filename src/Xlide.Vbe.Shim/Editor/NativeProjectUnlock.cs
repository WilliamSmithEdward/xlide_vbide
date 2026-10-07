using System.Runtime.InteropServices;
using Xlide.Vbe.Shim.Interop;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>Asks the VBE's own Project Explorer to expand one locked project.</summary>
internal static class NativeProjectUnlock
{
    private const uint TvmExpand = 0x1102;
    private const uint TvmGetNextItem = 0x110A;
    private const uint TvmGetItemW = 0x113E;
    private const int TvgnRoot = 0;
    private const int TvgnNext = 1;
    private const int TveExpand = 2;
    private const uint TvifText = 1;

    [StructLayout(LayoutKind.Sequential)]
    private unsafe struct TreeItem
    {
        public uint Mask;
        public nint Handle;
        public uint State;
        public uint StateMask;
        public char* Text;
        public int TextCapacity;
        public int Image;
        public int SelectedImage;
        public int ChildCount;
        public nint Parameter;
    }

    internal static unsafe bool Expand(nint frame, string explorerCaption, string workbookName)
    {
        // A floating Project Explorer is an owned top-level palette. A docked one is a frame
        // child. In either layout, only the native tree of this editor may receive the message.
        var palette = explorerCaption.Length == 0
            ? 0
            : CodePaneTracker.FindTopLevelByCaption(explorerCaption);
        var trees = palette != 0
            ? CodePaneTracker.FindChildrenByClass(palette, "SysTreeView32")
            : CodePaneTracker.FindChildrenByClass(frame, "SysTreeView32");

        var exactMatches = new List<(nint Tree, nint Item)>();
        var stemMatches = new List<(nint Tree, nint Item)>();
        var stem = Path.GetFileNameWithoutExtension(workbookName);
        var buffer = stackalloc char[512];
        foreach (var tree in trees)
        {
            for (var item = Win32.SendMessage(tree, TvmGetNextItem, TvgnRoot, 0);
                 item != 0;
                 item = Win32.SendMessage(tree, TvmGetNextItem, TvgnNext, item))
            {
                buffer[0] = '\0';
                var data = new TreeItem
                {
                    Mask = TvifText,
                    Handle = item,
                    Text = buffer,
                    TextCapacity = 512
                };
                if (Win32.SendMessage(tree, TvmGetItemW, 0, (nint)(&data)) == 0)
                {
                    continue;
                }

                var caption = new string(buffer);
                if (caption.EndsWith($"({workbookName})", StringComparison.OrdinalIgnoreCase))
                {
                    exactMatches.Add((tree, item));
                }
                else if (caption.EndsWith($"({stem})", StringComparison.OrdinalIgnoreCase))
                {
                    // Word drops .docm from its native tree label. Only use this fallback
                    // when it identifies exactly one project in the visible native tree.
                    stemMatches.Add((tree, item));
                }
            }
        }

        // Never send a password prompt to an arbitrary project when two open files share
        // the same native label. The caller reports the ambiguity instead.
        var selected = exactMatches.Count == 1
            ? exactMatches[0]
            : exactMatches.Count == 0 && stemMatches.Count == 1
                ? stemMatches[0]
                : default;
        // Expanding a protected item opens a modal VBE password dialog. Queue the message
        // so the WebView command returns before VBE enters that modal loop.
        return selected.Item != 0 && Win32.PostMessage(selected.Tree, TvmExpand, TveExpand, selected.Item);
    }
}
