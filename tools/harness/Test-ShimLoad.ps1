<#
.SYNOPSIS
    Loads the shim into a process of a chosen width, with no Office involved, and checks what it
    does there.

.DESCRIPTION
    32-bit and 64-bit Office cannot be installed side by side, so a machine that develops xlide
    has no host for the 32-bit shim. What every 64-bit Windows does have is a 32-bit Windows
    PowerShell, and a process is all a library needs to be loaded into.

    For each architecture asked for, this starts the Windows PowerShell of that width and, inside
    it, without registering anything:

      1. loads the shim and asks its DllGetClassObject for the add-in's class factory, the way
         COM does for a host;
      2. creates the add-in object and calls it by name through IDispatch, the way a host that
         late-binds does, including a call that must be refused;
      3. in a Debug shim, calls the XlideSelfCheck export, which runs the product's own window,
         automation, type library, font, picture, UI Automation, job, engine and browser code in
         that process. See src\Xlide.Vbe.Shim\Diagnostics\SelfCheck.cs.

    With both architectures, the 64-bit run is the known-good one: it shows the checks pass where
    the product is known to work, and every line the self-check marks "same" must then read alike
    in the two reports.

    A Release shim has no self-check, and this fails if it has one.

    -InstalledLayout loads the shims from a copy laid out the way the installer lays them out:
    the 64-bit shim with the page, the engine and the icon around it, and the 32-bit shim and its
    loader in an x86 folder below, looking one level up for the rest.

    -Registered loads nothing by path. It asks COM for the add-in by its ProgID, the way the
    editor does, and checks which file COM loaded into a process of each width. It is for a
    machine xlide has just been installed on: a CI runner. It reads the registration and
    changes nothing.

    Nothing here proves the editor works in a 32-bit Office. It proves the library loads in a
    32-bit process, that its own code runs there, and that an installation hands each width its
    own library.

.EXAMPLE
    tools\harness\Test-ShimLoad.ps1
    Both architectures of the Debug shim, compared. Publish them first:
    dotnet publish src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj -c Debug -r win-x86   (and win-x64)

.EXAMPLE
    tools\harness\Test-ShimLoad.ps1 -InstalledLayout
    The same, from a copy laid out as an installation.

.EXAMPLE
    tools\harness\Test-ShimLoad.ps1 -Architecture x86 -Configuration Release
    The load and the late-bound calls only, against the library that ships.

.EXAMPLE
    tools\harness\Test-ShimLoad.ps1 -Skip engine
    Leave a stage out, for a tree with no packaged engine. The report names what was left out.

.EXAMPLE
    tools\harness\Test-ShimLoad.ps1 -Registered -Configuration Release
    After an install: each width of process is handed its own library by COM.
#>
[CmdletBinding()]
param(
    [ValidateSet('x86', 'x64', 'both')]
    [string] $Architecture = 'both',

    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Debug',

    # Self-check stages to run, comma separated. Every stage when empty.
    [string] $Stage = '',

    # Self-check stages to leave out, comma separated.
    [string] $Skip = '',

    # A shim somewhere other than this tree's publish folder. Needs one architecture.
    [string] $Shim = '',

    # Load from a copy of the Debug publishes laid out as an installation.
    [switch] $InstalledLayout,

    # Create the add-in through its registration instead of loading a file.
    [switch] $Registered,

    # Where -Registered expects the installation to be.
    [string] $InstallFolder = (Join-Path $env:LOCALAPPDATA 'Programs\xlide'),

    # How long one architecture's run may take before it is stopped.
    [int] $TimeoutSeconds = 300,

    # Internal. This process is the host of the width asked for; the script started it.
    [switch] $Inner,

    # Internal. Where the inner process writes its findings.
    [string] $Report = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$allStages = @('layout', 'window', 'dispatch', 'typelib', 'fonts', 'picture', 'uia', 'job', 'engine', 'webview')

function Split-Names([string] $text) {
    @($text -split ',' | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
}

# ---------------------------------------------------------------------------------------------
# Inside the host of the chosen width.
# ---------------------------------------------------------------------------------------------
if ($Inner) {
    $width = [IntPtr]::Size * 8
    $wanted = if ($Architecture -eq 'x86') { 32 } else { 64 }
    if ($width -ne $wanted) {
        throw "This is a $width-bit process and a $wanted-bit one was asked for."
    }

    # C# 5, because Windows PowerShell compiles this with the compiler in the .NET Framework.
    Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

public static class ShimLoadProbe
{
    [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern IntPtr LoadLibraryW(string path);

    [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Ansi, BestFitMapping = false)]
    static extern IntPtr GetProcAddress(IntPtr module, string name);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int GetClassObject(ref Guid classId, ref Guid interfaceId, out IntPtr factory);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int CanUnloadNow();

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int SelfCheck(
        [MarshalAs(UnmanagedType.LPWStr)] string reportPath,
        [MarshalAs(UnmanagedType.LPWStr)] string stages);

    [ComImport, Guid("00000001-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IClassFactory
    {
        [PreserveSig] int CreateInstance(IntPtr outer, ref Guid interfaceId, out IntPtr instance);
        [PreserveSig] int LockServer(bool hold);
    }

    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatchRaw
    {
        [PreserveSig] int GetTypeInfoCount(out uint count);
        [PreserveSig] int GetTypeInfo(uint index, uint locale, out IntPtr info);
        [PreserveSig] int GetIDsOfNames(ref Guid reserved,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] names,
            uint count, uint locale, [Out, MarshalAs(UnmanagedType.LPArray)] int[] ids);
    }

    const string ProgId = "Xlide.VbeAddIn";
    const string LibraryName = "Xlide.Vbe.Shim.dll";

    static string report;
    static int failed;

    static void Write(string line) { File.AppendAllText(report, line + "\r\n"); }
    static void Note(string text) { Write("note " + text); }

    static bool Check(string what, bool passed, string detail)
    {
        if (!passed) { failed++; }
        Write((passed ? "ok   load: " : "FAIL load: ") + what + (string.IsNullOrEmpty(detail) ? "" : " -- " + detail));
        return passed;
    }

    static string Hex(int value) { return "0x" + value.ToString("X8"); }

    static void Begin(string reportPath)
    {
        report = reportPath;
        failed = 0;
        Note("host " + Process.GetCurrentProcess().MainModule.FileName + ", pointer size " + IntPtr.Size
            + ", apartment " + System.Threading.Thread.CurrentThread.GetApartmentState());
    }

    // The last line of every report, and how the driver knows the run reached its end.
    static int End(int more)
    {
        int total = failed + Math.Max(more, 0);
        Write("end " + total + " failed");
        return total;
    }

    // What a late-bound call through the runtime's own IDispatch client came back with.
    //
    // Read off the exception, never through Marshal.GetHRForException: that call also leaves the
    // exception on the thread as its COM error, and the next automation object to fail without
    // describing itself is then reported in THIS failure's words. It put "Unknown name" into
    // the self-check's file system call the first time this ran.
    static int CallByName(object target, string member, object[] arguments)
    {
        try
        {
            target.GetType().InvokeMember(member, BindingFlags.InvokeMethod, null, target, arguments);
            return 0;
        }
        catch (TargetInvocationException ex)
        {
            return (ex.InnerException ?? ex).HResult;
        }
        catch (Exception ex)
        {
            return ex.HResult;
        }
    }

    // Called the way a host that late-binds calls it: names resolved, arguments packed as
    // variants, Invoke.
    static void LateBound(object addIn)
    {
        var dispatch = (IDispatchRaw)addIn;

        uint count;
        int result = dispatch.GetTypeInfoCount(out count);
        Check("it says it has no type information", result == 0 && count == 0, Hex(result) + ", " + count);

        foreach (string name in new[] { "OnConnection", "OnDisconnection", "OnAddInsUpdate", "OnStartupComplete", "OnBeginShutdown" })
        {
            Guid empty = Guid.Empty;
            int[] ids = new int[1];
            result = dispatch.GetIDsOfNames(ref empty, new[] { name }, 1, 0, ids);
            Check("it knows " + name + " by name", result == 0 && ids[0] > 0, Hex(result) + ", id " + ids[0]);
        }

        // The first is accepted and does nothing. The second is OnConnection handed three
        // numbers where an editor and an add-in object belong, which it must refuse as a bad
        // argument without starting anything.
        result = CallByName(addIn, "OnAddInsUpdate", new object[] { 0 });
        Check("a late-bound call arrives and is accepted", result == 0, Hex(result));

        result = CallByName(addIn, "OnConnection", new object[] { 1, 2, 3 });
        Check("a late-bound OnConnection without an editor is refused as a bad argument",
            result == unchecked((int)0x80070057), Hex(result));

        result = CallByName(addIn, "NoSuchMember", new object[0]);
        Check("a member it does not have is an unknown name", result == unchecked((int)0x80020006), Hex(result));
    }

    public static int Run(string shim, string reportPath, bool debug, string stages)
    {
        Begin(reportPath);
        Note("shim " + shim);

        IntPtr module = LoadLibraryW(shim);
        if (!Check("the library loads into this process", module != IntPtr.Zero,
            module != IntPtr.Zero ? "" : "error " + Marshal.GetLastWin32Error() + " (193 is a library of the other width)"))
        {
            return End(0);
        }

        IntPtr canUnload = GetProcAddress(module, "DllCanUnloadNow");
        IntPtr entry = GetProcAddress(module, "DllGetClassObject");
        if (!Check("it exports the two functions COM looks for, by their plain names",
            entry != IntPtr.Zero && canUnload != IntPtr.Zero, ""))
        {
            return End(0);
        }

        var unload = (CanUnloadNow)Marshal.GetDelegateForFunctionPointer(canUnload, typeof(CanUnloadNow));
        int answer = unload();
        Check("DllCanUnloadNow says never", answer == 1, Hex(answer));

        var get = (GetClassObject)Marshal.GetDelegateForFunctionPointer(entry, typeof(GetClassObject));
        Guid classId = new Guid("588903F2-4CDE-4607-828A-6870A1F3FDC1");
        Guid factoryId = new Guid("00000001-0000-0000-C000-000000000046");
        IntPtr factoryPointer;
        int result = get(ref classId, ref factoryId, out factoryPointer);
        if (!Check("DllGetClassObject hands over the add-in's class factory",
            result == 0 && factoryPointer != IntPtr.Zero, Hex(result)))
        {
            return End(0);
        }

        Guid other = new Guid("11111111-2222-3333-4444-555555555555");
        IntPtr none;
        result = get(ref other, ref factoryId, out none);
        Check("and refuses a class it does not have", result == unchecked((int)0x80040111) && none == IntPtr.Zero, Hex(result));

        var factory = (IClassFactory)Marshal.GetObjectForIUnknown(factoryPointer);
        Guid dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
        IntPtr dispatchPointer;
        result = factory.CreateInstance(IntPtr.Zero, ref dispatchId, out dispatchPointer);
        if (!Check("the factory creates the add-in object", result == 0 && dispatchPointer != IntPtr.Zero, Hex(result)))
        {
            return End(0);
        }

        LateBound(Marshal.GetObjectForIUnknown(dispatchPointer));

        Marshal.Release(dispatchPointer);
        Marshal.Release(factoryPointer);

        IntPtr check = GetProcAddress(module, "XlideSelfCheck");
        if (!debug)
        {
            Check("a Release library carries no self-check", check == IntPtr.Zero, "");
            Note("Release: the load and the late-bound calls are all that was checked");
            return End(0);
        }

        if (!Check("a Debug library exports its self-check", check != IntPtr.Zero, ""))
        {
            return End(0);
        }

        var run = (SelfCheck)Marshal.GetDelegateForFunctionPointer(check, typeof(SelfCheck));
        int inside = run(reportPath, stages);
        if (inside < 0)
        {
            Check("the self-check ran", false, "it answered " + inside + "; see the shim's log");
        }

        return End(inside);
    }

    // Nothing is loaded by path here. COM is asked for the add-in by name, as the editor asks,
    // and whichever file it loads into this process is the one this width is registered with.
    public static int RunRegistered(string expected, string reportPath)
    {
        Begin(reportPath);
        Note("expecting " + expected);

        Type type = Type.GetTypeFromProgID(ProgId, false);
        if (!Check("the ProgID resolves to a class", type != null, ProgId))
        {
            return End(0);
        }

        object addIn;
        try
        {
            addIn = Activator.CreateInstance(type);
        }
        catch (Exception ex)
        {
            Exception inner = ex.InnerException ?? ex;
            Check("COM creates the add-in for a process of this width", false,
                Hex(inner.HResult) + " " + inner.Message.Trim() + " (0x80040154 is a class this width cannot see)");
            return End(0);
        }

        Check("COM creates the add-in for a process of this width", addIn != null, "");

        string loaded = null;
        foreach (ProcessModule module in Process.GetCurrentProcess().Modules)
        {
            if (string.Equals(module.ModuleName, LibraryName, StringComparison.OrdinalIgnoreCase))
            {
                loaded = module.FileName;
            }
        }

        Check("and the library it loaded is the one installed for this width",
            string.Equals(loaded, expected, StringComparison.OrdinalIgnoreCase), loaded ?? "(none loaded)");

        LateBound(addIn);
        return End(0);
    }
}
'@

    $failed = if ($Registered) {
        [ShimLoadProbe]::RunRegistered($Shim, $Report)
    } else {
        [ShimLoadProbe]::Run($Shim, $Report, ($Configuration -eq 'Debug'), $Stage)
    }

    exit ([Math]::Min($failed, 100))
}

# ---------------------------------------------------------------------------------------------
# The driver.
# ---------------------------------------------------------------------------------------------
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$architectures = if ($Architecture -eq 'both') { @('x86', 'x64') } else { @($Architecture) }

if ($Shim -and $architectures.Count -ne 1) {
    throw '-Shim names one library, so it needs -Architecture x86 or x64.'
}

if ($Shim -and ($InstalledLayout -or $Registered)) {
    throw '-Shim is for loading one file by path; -InstalledLayout and -Registered find their own.'
}

if ($InstalledLayout -and $Registered) {
    throw '-InstalledLayout stages a copy and -Registered reads a real installation. Ask for one.'
}

if ($InstalledLayout -and $Configuration -ne 'Debug') {
    throw '-InstalledLayout is about what the shim finds around it, which only the Debug self-check looks at.'
}

# Each call wrapped: a function that returns no names returns nothing at all, and nothing piped
# onward arrives as one empty item that is in no list.
$named = @(Split-Names $Stage) + @(Split-Names $Skip)
$unknown = @($named | Where-Object { $_ -and $_ -notin $allStages })
if ($unknown.Count -gt 0) {
    throw "No such stage: $($unknown -join ', '). The stages are $($allStages -join ', ')."
}

$skipped = @(Split-Names $Skip)
$asked = @(Split-Names $Stage)
if ($asked.Count -eq 0) { $asked = $allStages }
$stages = @($asked | Where-Object { $_ -notin $skipped })
$selfCheck = $Configuration -eq 'Debug' -and -not $Registered

# The Windows PowerShell of each width. From a 32-bit process System32 is redirected to the
# 32-bit folder, and Sysnative is the name that reaches the real one.
$system64 = if ([Environment]::Is64BitProcess) { 'System32' } else { 'Sysnative' }
$hosts = @{
    x86 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    x64 = Join-Path $env:WINDIR "$system64\WindowsPowerShell\v1.0\powershell.exe"
}

$reportRoot = Join-Path $repoRoot 'artifacts\shim-load'
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null

function Get-PublishFolder([string] $arch) {
    Join-Path $repoRoot "artifacts\publish\Xlide.Vbe.Shim\$($Configuration.ToLowerInvariant())_win-$arch"
}

$failures = 0
$staged = $null

if ($InstalledLayout) {
    # The layout installer\build.ps1 stages and InstallLayout describes: everything from the
    # 64-bit publish, the engine beside it, and the 32-bit shim and its loader in x86 below.
    $staged = Join-Path $reportRoot 'installed'
    if (Test-Path $staged) { Remove-Item $staged -Recurse -Force }
    New-Item -ItemType Directory -Path (Join-Path $staged 'x86') -Force | Out-Null

    foreach ($arch in 'x64', 'x86') {
        if (-not (Test-Path (Join-Path (Get-PublishFolder $arch) 'Xlide.Vbe.Shim.dll'))) {
            throw "No $Configuration shim for win-$arch to lay out. Publish both first."
        }
    }

    Copy-Item (Join-Path (Get-PublishFolder 'x64') '*') $staged -Recurse -Force -Exclude '*.pdb'
    foreach ($name in 'Xlide.Vbe.Shim.dll', 'WebView2Loader.dll') {
        Copy-Item (Join-Path (Get-PublishFolder 'x86') $name) (Join-Path $staged "x86\$name") -Force
    }

    $engine = Join-Path $repoRoot 'engine\dist\xlide-engine.exe'
    if ('engine' -in $stages) {
        if (-not (Test-Path $engine)) { throw "No engine at $engine to lay out. Package it, or pass -Skip engine." }
        Copy-Item $engine (Join-Path $staged 'xlide-engine.exe') -Force
    }

    Write-Host "Laid out as an installation at $staged"
}

$same = @{}

foreach ($arch in $architectures) {
    $library = if ($Shim) {
        (Resolve-Path $Shim).Path
    } elseif ($Registered) {
        if ($arch -eq 'x86') { Join-Path $InstallFolder 'x86\Xlide.Vbe.Shim.dll' } else { Join-Path $InstallFolder 'Xlide.Vbe.Shim.dll' }
    } elseif ($staged) {
        if ($arch -eq 'x86') { Join-Path $staged 'x86\Xlide.Vbe.Shim.dll' } else { Join-Path $staged 'Xlide.Vbe.Shim.dll' }
    } else {
        Join-Path (Get-PublishFolder $arch) 'Xlide.Vbe.Shim.dll'
    }

    $mode = if ($Registered) { 'as registered' } elseif ($staged) { "$Configuration, laid out as installed" } else { $Configuration }
    Write-Host ''
    Write-Host "==> $arch, $mode" -ForegroundColor Cyan

    if (-not (Test-Path $library)) {
        $hint = if ($Registered) { 'xlide is not installed there.' }
            else { "Publish it: dotnet publish src\Xlide.Vbe.Shim\Xlide.Vbe.Shim.csproj -c $Configuration -r win-$arch" }
        Write-Host "FAIL $arch no shim at $library. $hint" -ForegroundColor Red
        $failures++
        continue
    }

    $tag = if ($Registered) { 'registered' } elseif ($staged) { 'installed' } else { $Configuration.ToLowerInvariant() }
    $report = Join-Path $reportRoot "$tag-$arch.txt"
    if (Test-Path $report) { Remove-Item $report -Force }

    if ($selfCheck -and $skipped.Count -gt 0) {
        Write-Host "note left out on request: $($skipped -join ', ')" -ForegroundColor Yellow
    }

    $arguments = @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-STA',
        '-File', "`"$PSCommandPath`"",
        '-Inner', '-Architecture', $arch, '-Configuration', $Configuration,
        '-Shim', "`"$library`"", '-Report', "`"$report`"", '-Stage', "`"$($stages -join ',')`""
    )
    if ($Registered) { $arguments += '-Registered' }

    $console = Join-Path $reportRoot "$tag-$arch.console.txt"
    $process = Start-Process -FilePath $hosts[$arch] -ArgumentList $arguments -NoNewWindow -PassThru `
        -RedirectStandardOutput $console -RedirectStandardError "$console.err"

    # Read now: a process started this way does not keep its exit code readable once it is gone
    # unless something already holds its handle.
    $null = $process.Handle

    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        try { $process.Kill() } catch { Write-Verbose "The host had already gone: $_" }
        Write-Host "FAIL $arch the host did not finish in $TimeoutSeconds seconds and was stopped" -ForegroundColor Red
        $failures++
    }

    $lines = @(if (Test-Path $report) { Get-Content -LiteralPath $report -Encoding UTF8 })
    foreach ($line in $lines) {
        $colour = if ($line.StartsWith('FAIL')) { 'Red' } elseif ($line.StartsWith('note')) { 'DarkGray' } else { 'Gray' }
        Write-Host $line -ForegroundColor $colour
    }

    foreach ($extra in @($console, "$console.err")) {
        $said = @(if (Test-Path $extra) { Get-Content -LiteralPath $extra -Encoding UTF8 | Where-Object { $_.Trim() } })
        foreach ($line in $said) { Write-Host "host: $line" -ForegroundColor Yellow }
    }

    $failed = @($lines | Where-Object { $_.StartsWith('FAIL') }).Count
    $passed = @($lines | Where-Object { $_.StartsWith('ok  ') }).Count
    $exit = if ($process.HasExited) { $process.ExitCode } else { -1 }

    # A report is whole only when its last line is the probe's own. A process that died part-way
    # leaves no failing line of its own, and counting failures alone would pass it.
    $ended = $lines.Count -gt 0 -and $lines[-1].StartsWith('end ')
    if ($failed -eq 0 -and (-not $ended -or $exit -ne 0)) {
        $last = if ($lines.Count -gt 0) { $lines[-1] } else { '(nothing was written)' }
        Write-Host "FAIL $arch the run did not reach its end (exit $exit). Its last line: $last" -ForegroundColor Red
        $failed = 1
    }

    # Laid out as installed, the point is where the 32-bit shim looked for what both share.
    if ($staged -and $selfCheck) {
        $shared = @($lines | Where-Object { $_.StartsWith('note shared folder ') })
        $found = if ($shared.Count -eq 1) { $shared[0].Substring('note shared folder '.Length) } else { '(not reported)' }
        if ($found -ne $staged) {
            Write-Host "FAIL $arch the shim took $found as the shared folder; the installation is $staged" -ForegroundColor Red
            $failed++
        }
    }

    $failures += $failed
    $same[$arch] = @($lines | Where-Object { $_.StartsWith('same ') })

    $colour = if ($failed -eq 0) { 'Green' } else { 'Red' }
    Write-Host "RESULT ${arch}: $passed passed, $failed failed" -ForegroundColor $colour
}

if ($architectures.Count -eq 2 -and $selfCheck -and $same.ContainsKey('x86') -and $same.ContainsKey('x64')) {
    Write-Host ''
    Write-Host '==> the two reports, compared' -ForegroundColor Cyan

    $differences = @(Compare-Object -ReferenceObject $same['x64'] -DifferenceObject $same['x86'] -CaseSensitive)
    foreach ($difference in $differences) {
        $where = if ($difference.SideIndicator -eq '<=') { 'x64 only' } else { 'x86 only' }
        Write-Host "FAIL ${where}: $($difference.InputObject)" -ForegroundColor Red
    }

    # The stages that say something a process of either width must say alike. With one of them
    # run and nothing to compare, the comparison has checked nothing and must not read as passed.
    $comparable = @($stages | Where-Object { $_ -in @('dispatch', 'typelib', 'fonts', 'picture', 'engine') })
    if ($same['x64'].Count -eq 0 -and $comparable.Count -gt 0) {
        Write-Host 'FAIL nothing to compare: the 64-bit report has no "same" lines' -ForegroundColor Red
        $failures++
    }

    $failures += $differences.Count
    $colour = if ($differences.Count -eq 0) { 'Green' } else { 'Red' }
    Write-Host "RESULT same: $($same['x64'].Count) line(s) in the 64-bit report, $($differences.Count) difference(s)" -ForegroundColor $colour
}

Write-Host ''
if ($failures -eq 0) {
    Write-Host 'RESULT: passed' -ForegroundColor Green
    exit 0
}

Write-Host "RESULT: $failures failed" -ForegroundColor Red
exit 1
