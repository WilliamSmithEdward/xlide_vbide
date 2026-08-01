using Microsoft.Win32;
using Xlide.Vbe.Core;
using Xlide.Vbe.Core.Registration;

namespace Xlide.Vbe.Register;

/// <summary>
/// Applies, removes, or prints the add-in registration.
///
/// The installer writes the same values from its own authoring. Both come from
/// <see cref="RegistrationPlan"/>, and a test asserts the two agree, so this tool and a real
/// installation cannot drift apart.
/// </summary>
internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            var options = Options.Parse(args);
            if (options is null)
            {
                PrintUsage();
                return 2;
            }

            var entries = RegistrationPlan.Build(options.ShimPath, options.Bitness, RegistryScope.CurrentUser);

            switch (options.Action)
            {
                case Action.Print:
                    foreach (var entry in entries)
                    {
                        var name = entry.Name ?? "(default)";
                        var kind = entry.IsDword ? "DWORD" : "SZ";
                        Console.WriteLine($"HKCU\\{entry.Path}\\{name} = {entry.Value} [{kind}]");
                    }

                    break;

                case Action.Apply:
                    Apply(entries);
                    Console.WriteLine($"Registered {ProductIdentity.AddInProgId} for {options.Bitness} Office.");
                    Console.WriteLine($"Server: {options.ShimPath}");
                    break;

                case Action.Remove:
                    Remove(options.Bitness);
                    Console.WriteLine($"Removed {ProductIdentity.AddInProgId}.");
                    break;
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"error: {ex.Message}");
            return 1;
        }
    }

    private static void Apply(IReadOnlyList<RegistryEntry> entries)
    {
        foreach (var entry in entries)
        {
            using var key = Registry.CurrentUser.CreateSubKey(entry.Path, writable: true);
            if (key is null)
            {
                throw new InvalidOperationException($"Could not create HKCU\\{entry.Path}.");
            }

            if (entry.IsDword)
            {
                key.SetValue(entry.Name, int.Parse(entry.Value, System.Globalization.CultureInfo.InvariantCulture), RegistryValueKind.DWord);
            }
            else
            {
                key.SetValue(entry.Name, entry.Value, RegistryValueKind.String);
            }
        }
    }

    private static void Remove(HostBitness bitness)
    {
        // Inspect before deleting: only remove keys that name this product.
        DeleteSubTreeIfPresent($@"{RegistrationPlan.AddInsKeyPath(bitness)}\{ProductIdentity.AddInProgId}");

        var classes = RegistrationPlan.ClassesRoot(RegistryScope.CurrentUser);
        DeleteSubTreeIfPresent($@"{classes}\CLSID\{{{ProductIdentity.AddInClsid}}}");
        DeleteSubTreeIfPresent($@"{classes}\CLSID\{{{ProductIdentity.ToolWindowHostClsid}}}");
        DeleteSubTreeIfPresent($@"{classes}\{ProductIdentity.AddInProgId}");
        DeleteSubTreeIfPresent($@"{classes}\{ProductIdentity.ToolWindowHostProgId}");
        DeleteSubTreeIfPresent($@"{classes}\WOW6432Node\CLSID\{{{ProductIdentity.AddInClsid}}}");
        DeleteSubTreeIfPresent($@"{classes}\WOW6432Node\CLSID\{{{ProductIdentity.ToolWindowHostClsid}}}");
    }

    private static void DeleteSubTreeIfPresent(string path)
    {
        using var probe = Registry.CurrentUser.OpenSubKey(path);
        if (probe is null)
        {
            return;
        }

        probe.Dispose();
        Registry.CurrentUser.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
        Console.WriteLine($"removed HKCU\\{path}");
    }

    private static void PrintUsage()
    {
        Console.WriteLine("""
            xlide-register --apply|--remove|--print [options]

              --shim <path>       Path to Xlide.Vbe.Shim.dll. Required for --apply and --print.
              --bitness x64|x86   Office bitness to register for. Defaults to x64.

            Registration is per user and needs no administrator rights.
            """);
    }

    private enum Action
    {
        Apply,
        Remove,
        Print,
    }

    private sealed record Options(Action Action, string ShimPath, HostBitness Bitness)
    {
        public static Options? Parse(string[] args)
        {
            Action? action = null;
            string? shim = null;
            var bitness = HostBitness.X64;

            for (var i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--apply":
                        action = Action.Apply;
                        break;
                    case "--remove":
                        action = Action.Remove;
                        break;
                    case "--print":
                        action = Action.Print;
                        break;
                    case "--shim" when i + 1 < args.Length:
                        shim = Path.GetFullPath(args[++i]);
                        break;
                    case "--bitness" when i + 1 < args.Length:
                        bitness = args[++i].Equals("x86", StringComparison.OrdinalIgnoreCase)
                            ? HostBitness.X86
                            : HostBitness.X64;
                        break;
                    default:
                        return null;
                }
            }

            if (action is null)
            {
                return null;
            }

            if (action != Action.Remove)
            {
                if (shim is null)
                {
                    Console.Error.WriteLine("error: --shim is required.");
                    return null;
                }

                if (!File.Exists(shim))
                {
                    Console.Error.WriteLine($"error: no such file: {shim}");
                    return null;
                }
            }

            return new Options(action.Value, shim ?? string.Empty, bitness);
        }
    }
}
