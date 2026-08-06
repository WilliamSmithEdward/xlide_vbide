using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using Microsoft.Win32;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Registration;

namespace Xlide.Setup;

/// <summary>
/// Installs and removes the product for the current user.
///
/// The primary registration lives under the user's own profile and registry hive and needs no
/// administrator rights — the editor resolves class registration through the user hive, and a
/// development tool that demands elevation excludes exactly the people most likely to want it.
///
/// Click-to-Run Office resolves some registry namespaces through its own overlay. When the
/// installer happens to run elevated on such a machine it plants the same registration inside the
/// overlay as well; it never asks for elevation itself, because the per-user registration is the
/// documented, complete mechanism. The --overlay-only verb exists for running that one step
/// deliberately.
///
/// The installer reuses the same registration description the product and its tests use, so there
/// is no second copy of the registry layout that could disagree with the first.
/// </summary>
internal static class Program
{
    private const string PayloadPrefix = "payload/";
    private const string CompressedSuffix = ".br";
    private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\xlide";

    private static int Main(string[] args)
    {
        var silent = args.Contains("--silent", StringComparer.OrdinalIgnoreCase);
        var uninstall = args.Contains("--uninstall", StringComparer.OrdinalIgnoreCase);
        var relaunched = args.Contains("--relaunched", StringComparer.OrdinalIgnoreCase);

        // The elevated pass of this same executable: touch only the Click-to-Run overlay and exit.
        if (args.Contains("--overlay-only", StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                return uninstall ? RemoveOverlayRegistration() : WriteOverlayRegistration(ShimArgument(args));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"The overlay step failed: {ex.Message}");
                return 1;
            }
        }

        try
        {
            var keepData = args.Contains("--keep-data", StringComparer.OrdinalIgnoreCase);
            var result = uninstall ? Uninstall(silent, relaunched, keepData) : Install(silent);

            if (!silent && !relaunched)
            {
                WaitForAcknowledgement();
            }

            return result;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine($"The operation failed: {ex.Message}");

            if (!silent)
            {
                WaitForAcknowledgement();
            }

            return 1;
        }
    }

    private static int Install(bool silent)
    {
        var payload = ReadPayload();
        if (payload.Count == 0)
        {
            Console.Error.WriteLine("This installer was built without a payload and cannot install anything.");
            return 2;
        }

        var target = DefaultInstallFolder();

        Report(silent, $"Installing {ProductIdentity.FriendlyName} to {target}");

        if (IsHostRunning(out var hostPid))
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine($"Excel is running (process {hostPid}).");
            Console.Error.WriteLine("Close it and run this installer again. The editor holds the add-in open while it runs.");
            return 3;
        }

        Directory.CreateDirectory(target);

        foreach (var (name, content) in payload)
        {
            var destination = Path.Combine(target, name);
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.WriteAllBytes(destination, content);
            Report(silent, $"  {name} ({content.Length / 1024.0:N0} KB)");
        }

        // Keep a copy of the installer so the product can be removed without the original download.
        var uninstaller = Path.Combine(target, "xlide-setup.exe");
        var self = Environment.ProcessPath;
        if (self is not null && !string.Equals(self, uninstaller, StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(self, uninstaller, overwrite: true);
        }

        var shim = Path.Combine(target, ProductIdentity.ShimFileName);
        if (!File.Exists(shim))
        {
            throw new InvalidOperationException($"The payload did not contain {ProductIdentity.ShimFileName}.");
        }

        ApplyRegistration(shim);
        WriteUninstallEntry(target, uninstaller, payload.Sum(p => p.Content.Length));

        Report(silent, string.Empty);
        Report(silent, "Registered for the current user. No administrator rights were needed.");

        SupplementOverlay(shim, silent);

        if (!HasWebViewRuntime())
        {
            Report(silent, string.Empty);
            Report(silent, "Note: the WebView2 runtime was not detected. Panels will not render until it is present.");
            Report(silent, "It ships with Windows 11 and with Microsoft Edge, so this is unusual.");
        }

        Report(silent, string.Empty);
        Report(silent, "Start Excel and open the Visual Basic Editor to use it.");
        return 0;
    }

    private static int Uninstall(bool silent, bool relaunched, bool keepData)
    {
        var target = DefaultInstallFolder();
        var self = Environment.ProcessPath;

        // A running executable cannot delete itself. When the uninstaller is the copy living in the
        // installation folder, it re-runs from a temporary location so the folder can be removed
        // whole rather than left behind with one file in it.
        if (!relaunched && self is not null && self.StartsWith(target, StringComparison.OrdinalIgnoreCase))
        {
            var staged = Path.Combine(Path.GetTempPath(), $"xlide-setup-{Environment.ProcessId}.exe");
            File.Copy(self, staged, overwrite: true);

            var arguments = "--uninstall --relaunched"
                + (silent ? " --silent" : string.Empty)
                + (keepData ? " --keep-data" : string.Empty);
            Process.Start(new ProcessStartInfo(staged, arguments) { UseShellExecute = false });
            return 0;
        }

        Report(silent, $"Removing {ProductIdentity.FriendlyName}");

        if (IsHostRunning(out var hostPid))
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine($"Excel is running (process {hostPid}). Close it and try again.");
            return 3;
        }

        RemoveRegistration();
        RemoveOverlaySupplement(silent);
        DeleteKeyIfPresent(UninstallKey);

        if (Directory.Exists(target))
        {
            TryDeleteDirectory(target);
        }

        // Logs and the browser surface's cache are ours, not the user's work. Their VBA lives in
        // their workbooks and is never touched by any of this, so removing the data folder leaves
        // the machine as it was before installation.
        var data = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            ProductIdentity.DataFolderName);

        if (Directory.Exists(data) && !keepData)
        {
            TryDeleteDirectory(data);
        }

        Report(silent, "Removed.");
        return 0;
    }

    private static string DefaultInstallFolder() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Programs",
        "xlide");

    private static List<(string Name, byte[] Content)> ReadPayload()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var files = new List<(string, byte[])>();

        foreach (var resource in assembly.GetManifestResourceNames())
        {
            if (!resource.StartsWith(PayloadPrefix, StringComparison.Ordinal))
            {
                continue;
            }

            using var stream = assembly.GetManifestResourceStream(resource);
            if (stream is null)
            {
                continue;
            }

            var name = resource[PayloadPrefix.Length..];

            // tools\Xlide.Payload.Pack compresses the payload before it is embedded, because most of
            // it is a language runtime and a runtime binary is about seventy per cent air. An
            // uncompressed entry is still honoured, so a payload staged by hand installs too.
            using var buffer = new MemoryStream();
            if (name.EndsWith(CompressedSuffix, StringComparison.OrdinalIgnoreCase))
            {
                name = name[..^CompressedSuffix.Length];
                using var brotli = new BrotliStream(stream, CompressionMode.Decompress);
                brotli.CopyTo(buffer);
            }
            else
            {
                stream.CopyTo(buffer);
            }

            files.Add((name.Replace('/', Path.DirectorySeparatorChar), buffer.ToArray()));
        }

        return files;
    }

    private static void ApplyRegistration(string shimPath)
    {
        foreach (var entry in RegistrationPlan.Build(shimPath, HostBitness.X64, RegistryScope.CurrentUser))
        {
            WriteEntry(Registry.CurrentUser, entry.Path, entry);
        }
    }

    private static void WriteEntry(RegistryKey hive, string path, RegistryEntry entry)
    {
        using var key = hive.CreateSubKey(path, writable: true)
            ?? throw new InvalidOperationException($"Could not create {hive.Name}\\{path}.");

        if (entry.IsDword)
        {
            key.SetValue(entry.Name, int.Parse(entry.Value, System.Globalization.CultureInfo.InvariantCulture), RegistryValueKind.DWord);
        }
        else
        {
            key.SetValue(entry.Name, entry.Value, RegistryValueKind.String);
        }
    }

    /// <summary>
    /// Plants the registration inside the Click-to-Run overlay too, when that is possible without
    /// asking anyone for anything: only on C2R machines, and only when already elevated.
    /// </summary>
    private static void SupplementOverlay(string shimPath, bool silent)
    {
        using var overlay = Registry.LocalMachine.OpenSubKey(RegistrationPlan.OverlayVbaKey);
        if (overlay is null)
        {
            return;
        }

        // Never prompt for elevation: the per-user registration is the documented, complete
        // mechanism. The overlay copy is written only when the installer already runs elevated,
        // or explicitly via --overlay-only, for machines whose virtualized view hides HKLM.
        if (IsElevated())
        {
            WriteOverlayRegistration(shimPath);
            ReportOverlayWritten(silent);
        }
    }

    private static void ReportOverlayWritten(bool silent)
    {
        Report(silent, "Registered inside Office's Click-to-Run overlay as well.");
        Report(silent, "Office updates can rebuild that overlay; if xlide disappears after one, run this installer again.");
    }

    /// <summary>Removes the overlay supplement, elevating only when there is one to remove.</summary>
    private static void RemoveOverlaySupplement(bool silent)
    {
        var addInKey = $@"{RegistrationPlan.OverlayPath(RegistrationPlan.AddInsKeyPath(HostBitness.X64))}\{ProductIdentity.AddInProgId}";

        using (var probe = Registry.LocalMachine.OpenSubKey(addInKey))
        {
            if (probe is null)
            {
                return;
            }
        }

        if (IsElevated())
        {
            RemoveOverlayRegistration();
            return;
        }

        Report(silent, $"An overlay registration under HKLM\\{addInKey} remains. It is inert without the");
        Report(silent, "files; an elevated run with --uninstall --overlay-only removes it.");
    }

    private static int WriteOverlayRegistration(string shimPath)
    {
        foreach (var entry in RegistrationPlan.Build(shimPath, HostBitness.X64, RegistryScope.CurrentUser))
        {
            WriteEntry(Registry.LocalMachine, RegistrationPlan.OverlayPath(entry.Path), entry);
        }

        return 0;
    }

    private static int RemoveOverlayRegistration()
    {
        var classes = RegistrationPlan.OverlayPath(RegistrationPlan.ClassesRoot(RegistryScope.CurrentUser));

        foreach (var path in new[]
                 {
                     $@"{RegistrationPlan.OverlayPath(RegistrationPlan.AddInsKeyPath(HostBitness.X64))}\{ProductIdentity.AddInProgId}",
                     $@"{classes}\CLSID\{{{ProductIdentity.AddInClsid}}}",
                     $@"{classes}\{ProductIdentity.AddInProgId}",
                 })
        {
            using var probe = Registry.LocalMachine.OpenSubKey(path);
            if (probe is null)
            {
                continue;
            }

            probe.Dispose();
            Registry.LocalMachine.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
        }

        return 0;
    }

    private static string ShimArgument(string[] args)
    {
        for (var i = 0; i < args.Length - 1; i++)
        {
            if (args[i].Equals("--shim", StringComparison.OrdinalIgnoreCase))
            {
                var path = Path.GetFullPath(args[i + 1]);
                return File.Exists(path)
                    ? path
                    : throw new FileNotFoundException($"No shim at {path}.", path);
            }
        }

        throw new ArgumentException("--overlay-only needs --shim <path> when installing.");
    }

    private static bool IsElevated()
    {
        using var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
        return new System.Security.Principal.WindowsPrincipal(identity)
            .IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
    }

    private static void RemoveRegistration()
    {
        var classes = RegistrationPlan.ClassesRoot(RegistryScope.CurrentUser);

        DeleteKeyIfPresent($@"{RegistrationPlan.AddInsKeyPath(HostBitness.X64)}\{ProductIdentity.AddInProgId}");
        DeleteKeyIfPresent($@"{RegistrationPlan.AddInsKeyPath(HostBitness.X86)}\{ProductIdentity.AddInProgId}");
        DeleteKeyIfPresent($@"{classes}\CLSID\{{{ProductIdentity.AddInClsid}}}");
        DeleteKeyIfPresent($@"{classes}\{ProductIdentity.AddInProgId}");
        DeleteKeyIfPresent($@"{classes}\WOW6432Node\CLSID\{{{ProductIdentity.AddInClsid}}}");
    }

    private static void WriteUninstallEntry(string target, string uninstaller, int payloadBytes)
    {
        using var key = Registry.CurrentUser.CreateSubKey(UninstallKey, writable: true)
            ?? throw new InvalidOperationException("Could not write the uninstall entry.");

        var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.1.0";

        key.SetValue("DisplayName", ProductIdentity.FriendlyName);
        key.SetValue("DisplayVersion", version);
        key.SetValue("Publisher", "William Smith Edward");
        key.SetValue("InstallLocation", target);
        key.SetValue("UninstallString", $"\"{uninstaller}\" --uninstall");
        key.SetValue("QuietUninstallString", $"\"{uninstaller}\" --uninstall --silent");
        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        key.SetValue("EstimatedSize", Math.Max(1, payloadBytes / 1024), RegistryValueKind.DWord);
    }

    private static void DeleteKeyIfPresent(string path)
    {
        using var probe = Registry.CurrentUser.OpenSubKey(path);
        if (probe is null)
        {
            return;
        }

        probe.Dispose();
        Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
    }

    private static void TryDeleteDirectory(string path)
    {
        foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
        {
            try
            {
                File.Delete(file);
            }
            catch (IOException)
            {
                // The staged uninstaller may still be mapped. Leave it; the folder removal below
                // reports what remains rather than pretending the removal was complete.
            }
        }

        try
        {
            Directory.Delete(path, recursive: true);
        }
        catch (IOException)
        {
            Console.WriteLine($"Some files under {path} were in use and remain. They can be deleted manually.");
        }
    }

    /// <summary>
    /// Reports whether a host holds the add-in open. Overwriting a loaded library silently produces
    /// an installation that appears to succeed and does not take effect until the next restart.
    /// </summary>
    private static bool IsHostRunning(out int processId)
    {
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            processId = process.Id;
            process.Dispose();
            return true;
        }

        processId = 0;
        return false;
    }

    private static bool HasWebViewRuntime()
    {
        const string clients = @"Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";

        foreach (var root in new[]
                 {
                     Registry.LocalMachine.OpenSubKey($@"SOFTWARE\WOW6432Node\{clients}"),
                     Registry.LocalMachine.OpenSubKey($@"SOFTWARE\{clients}"),
                     Registry.CurrentUser.OpenSubKey($@"SOFTWARE\{clients}"),
                 })
        {
            using (root)
            {
                if (root?.GetValue("pv") is string version && version.Length > 0 && version != "0.0.0.0")
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static void Report(bool silent, string message)
    {
        if (!silent)
        {
            Console.WriteLine(message);
        }
    }

    private static void WaitForAcknowledgement()
    {
        if (Console.IsInputRedirected)
        {
            return;
        }

        Console.WriteLine();
        Console.Write("Press any key to close. ");

        try
        {
            Console.ReadKey(intercept: true);
        }
        catch (InvalidOperationException)
        {
            // No console is attached. Nothing to wait for.
        }

        Console.WriteLine();
    }
}
