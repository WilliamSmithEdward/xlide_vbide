namespace Xlide.Vbe.Core;

/// <summary>
/// The identifiers that outlive everything else. These appear in the registry, in the MSI, and in
/// the add-in list the VBE shows the user. Changing one after release orphans installed copies.
/// </summary>
public static class ProductIdentity
{
    /// <summary>ProgID the VBE uses as the add-in key name, and to resolve the CLSID.</summary>
    public const string AddInProgId = "Xlide.VbeAddIn";

    /// <summary>CLSID of the add-in coclass.</summary>
    public const string AddInClsid = "588903F2-4CDE-4607-828A-6870A1F3FDC1";

    /// <summary>ProgID of the ActiveX control the VBE sites inside a docked tool window.</summary>
    public const string ToolWindowHostProgId = "Xlide.VbeToolWindowHost";

    /// <summary>CLSID of the tool window host coclass.</summary>
    public const string ToolWindowHostClsid = "3DCD653C-57B6-4F7E-8704-7A14C5800152";

    /// <summary>Name shown in the VBE add-in manager.</summary>
    public const string FriendlyName = "xlide";

    /// <summary>Description shown in the VBE add-in manager.</summary>
    public const string Description = "Modern VBA development inside the Visual Basic Editor.";

    /// <summary>Folder name used under LOCALAPPDATA for logs and state.</summary>
    public const string DataFolderName = "xlide_vbide";

    /// <summary>File name of the native shim library.</summary>
    public const string ShimFileName = "Xlide.Vbe.Shim.dll";

    /// <summary>File name of the language engine sidecar.</summary>
    public const string EngineFileName = "xlide-engine.exe";
}
