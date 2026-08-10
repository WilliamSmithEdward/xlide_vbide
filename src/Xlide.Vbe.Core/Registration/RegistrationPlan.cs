using System.Globalization;

namespace Xlide.Vbe.Core.Registration;

/// <summary>Registry hive a registration entry targets.</summary>
public enum RegistryScope
{
    /// <summary>Per-user. Requires no administrator rights and is the default for this product.</summary>
    CurrentUser,

    /// <summary>Machine-wide. Required when the host runs elevated, because an elevated process
    /// does not see HKCU class registrations.</summary>
    LocalMachine,
}

/// <summary>Office bitness, which selects both the add-in key and the COM registration layout.</summary>
public enum HostBitness
{
    X86,
    X64,
}

/// <summary>One registry value that must exist for the add-in to load.</summary>
/// <param name="Path">Key path relative to the hive root.</param>
/// <param name="Name">Value name, or null for the key's default value.</param>
/// <param name="Value">Value data as a string, or the decimal digits of a DWORD.</param>
/// <param name="IsDword">True when <paramref name="Value"/> should be written as REG_DWORD.</param>
public sealed record RegistryEntry(string Path, string? Name, string Value, bool IsDword = false);

/// <summary>
/// Produces the exact set of registry values that register the add-in with the VBE.
///
/// This type is the single source of truth. The installer authoring and the development
/// registration script are both checked against it by a test, so the three cannot drift apart.
///
/// The VBE enumerates add-ins under Software\Microsoft\VBA\VBE\6.0\Addins for 32-bit hosts and
/// Addins64 for 64-bit hosts. The subkey name is the ProgID. The ProgID resolves to a CLSID, and
/// the CLSID resolves to the server library. Because the shim is a native library, InprocServer32
/// points at it directly rather than at a runtime host.
/// </summary>
public static class RegistrationPlan
{
    /// <summary>LoadBehavior value meaning "loaded, and load at host startup".</summary>
    public const int LoadAtStartup = 3;

    /// <summary>Apartment model the shim requires. The VBE is single-threaded apartment.</summary>
    public const string ThreadingModel = "Apartment";

    /// <summary>
    /// Builds every registry value required to register the add-in.
    /// </summary>
    /// <param name="shimPath">Absolute path to the native shim library.</param>
    /// <param name="bitness">Bitness of the Office installation being registered for.</param>
    /// <param name="scope">Hive to target.</param>
    public static IReadOnlyList<RegistryEntry> Build(string shimPath, HostBitness bitness, RegistryScope scope)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(shimPath);

        var entries = new List<RegistryEntry>();
        var classes = ClassesRoot(scope);

        AddCoClass(entries, classes, ProductIdentity.AddInClsid, ProductIdentity.AddInProgId,
            ProductIdentity.FriendlyName, shimPath, bitness);

        var addInsKey = AddInsKeyPath(bitness);
        entries.Add(new RegistryEntry($@"{addInsKey}\{ProductIdentity.AddInProgId}", "Description", ProductIdentity.Description));
        entries.Add(new RegistryEntry($@"{addInsKey}\{ProductIdentity.AddInProgId}", "FriendlyName", ProductIdentity.FriendlyName));
        entries.Add(new RegistryEntry($@"{addInsKey}\{ProductIdentity.AddInProgId}", "LoadBehavior", LoadAtStartup.ToString(), IsDword: true));
        entries.Add(new RegistryEntry($@"{addInsKey}\{ProductIdentity.AddInProgId}", "CommandLineSafe", "0", IsDword: true));

        return entries;
    }

    /// <summary>
    /// Root of the Click-to-Run registry overlay, relative to HKLM. C2R Office resolves the
    /// Software\Microsoft\VBA namespace through this overlay - the editor's machine-level VBA
    /// values exist only here - and on some machines the per-user Addins64 read resolves into it
    /// and comes back empty, hiding a correct HKCU registration. Registration planted here is in
    /// the view those reads consult. Office updates can rebuild the overlay, so anything written
    /// under it must be re-asserted rather than assumed durable.
    /// </summary>
    public const string OverlayRoot = @"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE";

    /// <summary>Key whose presence identifies a Click-to-Run Office whose VBA namespace lives in
    /// the overlay. Writable only by administrators, like everything under the overlay.</summary>
    public const string OverlayVbaKey = OverlayRoot + @"\SOFTWARE\Microsoft\VBA";

    /// <summary>
    /// Rebases a plan entry's hive-relative path into the overlay. The overlay's MACHINE branch
    /// mirrors the ordinary hive shape, so the same entries apply verbatim beneath it.
    /// </summary>
    public static string OverlayPath(string entryPath) => $@"{OverlayRoot}\{entryPath}";

    /// <summary>
    /// Key path, relative to the hive root, under which the VBE looks for add-ins.
    /// </summary>
    public static string AddInsKeyPath(HostBitness bitness) => bitness switch
    {
        HostBitness.X64 => @"Software\Microsoft\VBA\VBE\6.0\Addins64",
        HostBitness.X86 => @"Software\Microsoft\VBA\VBE\6.0\Addins",
        _ => throw new ArgumentOutOfRangeException(nameof(bitness)),
    };

    /// <summary>
    /// Root under which COM classes are registered for the given scope. HKCU\Software\Classes is
    /// merged into HKCR for non-elevated processes, which is what makes admin-free install work.
    /// </summary>
    public static string ClassesRoot(RegistryScope scope) => scope switch
    {
        RegistryScope.CurrentUser => @"Software\Classes",
        RegistryScope.LocalMachine => @"Software\Classes",
        _ => throw new ArgumentOutOfRangeException(nameof(scope)),
    };

    private static void AddCoClass(
        List<RegistryEntry> entries,
        string classes,
        string clsid,
        string progId,
        string friendlyName,
        string serverPath,
        HostBitness bitness)
    {
        var clsidKey = $@"{classes}\CLSID\{{{clsid}}}";

        entries.Add(new RegistryEntry(clsidKey, null, friendlyName));
        entries.Add(new RegistryEntry($@"{clsidKey}\InprocServer32", null, serverPath));
        entries.Add(new RegistryEntry($@"{clsidKey}\InprocServer32", "ThreadingModel", ThreadingModel));
        entries.Add(new RegistryEntry($@"{clsidKey}\ProgID", null, progId));
        entries.Add(new RegistryEntry($@"{classes}\{progId}", null, friendlyName));
        entries.Add(new RegistryEntry($@"{classes}\{progId}\CLSID", null, $"{{{clsid}}}"));

        // A 32-bit server registered on 64-bit Windows lives under the WOW node so that 32-bit
        // Office resolves it. The add-in key itself is not redirected, only the class registration.
        if (bitness == HostBitness.X86)
        {
            var wowKey = $@"Software\Classes\WOW6432Node\CLSID\{{{clsid}}}";
            entries.Add(new RegistryEntry($@"{wowKey}\InprocServer32", null, serverPath));
            entries.Add(new RegistryEntry($@"{wowKey}\InprocServer32", "ThreadingModel", ThreadingModel));
        }
    }
}
