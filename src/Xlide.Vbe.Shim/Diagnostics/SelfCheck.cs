#if DEBUG
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using Xlide.Vbe.Core.Hosting;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;
using Xlide.Vbe.Shim.Interop;
using Xlide.Vbe.Shim.WebView;

namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// Runs the parts of this library that depend on the width of a pointer, in whatever process
/// loaded it, with no editor present.
///
/// It exists because the 32-bit build has no host on a developer's machine: 32-bit and 64-bit
/// Office cannot be installed side by side, so a machine set up to develop this has only the
/// 64-bit one. What a 32-bit process CAN do without Office is load the library and run its own
/// code: make a window and read its slots back, call an automation object and take its error,
/// walk a type library, read a property through UI Automation, start the engine, and bring up
/// the browser. Those are the places a struct of the wrong size or an entry point that only one
/// user32 exports would show, and each is run here through the product's own code, not a copy.
///
/// It is not a test of the editor. Nothing here talks to a Visual Basic Editor, so everything
/// that needs one is still unproven in a 32-bit host until someone runs it in one.
///
/// The same checks run in a 64-bit process, where the product is known to work. That run is
/// what says the checks themselves are sound, and the lines marked "same" must then read alike
/// in both: a type library's members and a picture's pixels do not depend on who is reading.
/// tools\harness\Test-ShimLoad.ps1 drives both and compares them.
///
/// Debug builds only. The export does not exist in a shipped library.
/// </summary>
internal static unsafe partial class SelfCheck
{
    private const string ClassName = "XlideSelfCheckWindow";
    private const string Title = "xlide self-check";

    private static readonly Guid DispatchId = new("00020400-0000-0000-C000-000000000046");

    private static readonly (string Name, Action Body)[] Stages =
    [
        ("layout", Layout),
        ("window", Window),
        ("dispatch", DispatchCalls),
        ("typelib", TypeLibraries),
        ("fonts", Fonts),
        ("picture", Pictures),
        ("uia", UiAutomationReads),
        ("job", JobObject),
        ("engine", EngineRoundTrip),
        ("webview", BrowserSurface),
    ];

    private static string _report = string.Empty;
    private static string _stage = string.Empty;
    private static int _failed;

    private static nint _window;
    private static int _sizeMessages;

    /// <summary>
    /// Runs the named stages, or every stage when none is named, appending one line per finding
    /// to <paramref name="reportPath"/> as it goes, so a process that dies part-way leaves the
    /// line it died after. Answers how many checks failed.
    /// </summary>
    public static int Run(string reportPath, string stages)
    {
        _report = reportPath;
        _failed = 0;

        Note($"process {Environment.ProcessPath ?? "unknown"}, pointer size {nint.Size}");
        Note($"shim folder {ShimModule.Directory ?? "unknown"}");
        Note($"shared folder {ShimModule.SharedDirectory ?? "unknown"}");

        var wanted = stages.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (wanted.Length == 0)
        {
            wanted = [.. Stages.Select(stage => stage.Name)];
        }

        try
        {
            foreach (var name in wanted)
            {
                _stage = name;
                var body = Stages.FirstOrDefault(stage => string.Equals(stage.Name, name, StringComparison.OrdinalIgnoreCase)).Body;
                if (body is null)
                {
                    Fail("is a stage", $"there is none by that name; the stages are {string.Join(", ", Stages.Select(stage => stage.Name))}");
                    continue;
                }

                try
                {
                    body();
                }
                catch (Exception ex)
                {
                    Fail("ran to its end", $"{ex.GetType().Name}: {ex.Message}");
                }
            }
        }
        finally
        {
            _stage = "window";
            DestroyWindow();
        }

        Write($"done {_failed} failed");
        return _failed;
    }

    // ---- reporting ----------------------------------------------------------------------------

    private static void Write(string line)
    {
        Log.Info($"self-check: {line}");
        File.AppendAllText(_report, line + "\r\n");
    }

    private static void Note(string text) => Write($"note {text}");

    /// <summary>A fact that must read the same in a process of either width.</summary>
    private static void Same(string name, string value) => Write($"same {_stage}: {name} = {value}");

    private static void Fail(string check, string detail)
    {
        _failed++;
        Write($"FAIL {_stage}: {check} -- {detail}");
    }

    private static bool Check(string check, bool passed, string detail = "")
    {
        if (passed)
        {
            Write(detail.Length == 0 ? $"ok   {_stage}: {check}" : $"ok   {_stage}: {check} -- {detail}");
        }
        else
        {
            Fail(check, detail.Length == 0 ? "it did not hold" : detail);
        }

        return passed;
    }

    /// <summary>A short, stable digest of some text, so a long answer compares as one word.</summary>
    private static string Digest(IEnumerable<string> lines)
    {
        var hash = 14695981039346656037UL;
        foreach (var line in lines)
        {
            foreach (var character in line)
            {
                hash = (hash ^ character) * 1099511628211UL;
            }

            hash = (hash ^ '\n') * 1099511628211UL;
        }

        return hash.ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }

    // ---- stages -------------------------------------------------------------------------------

    /// <summary>The variant the accessibility reads go through is the size of a real variant.</summary>
    private static void Layout()
    {
        var real = Unsafe.SizeOf<ComVariant>();
        Check("the read variant is as wide as a variant", sizeof(UiVariant) == real, $"{sizeof(UiVariant)} against {real}");
        Check("the placeholder variant is as wide as a variant", sizeof(ComVariantBlock) == real, $"{sizeof(ComVariantBlock)} against {real}");

        var probe = default(UiVariant);
        var offset = (int)((byte*)&probe.Value - (byte*)&probe);
        Check("a variant's data starts at offset eight", offset == 8, $"offset {offset}");
    }

    /// <summary>
    /// A window of our own class: its slots hold a pointer, its class answers for its procedure,
    /// and the product's frame subclass is called back when it is resized.
    /// </summary>
    private static void Window()
    {
        var window = EnsureWindow();

        nint first = 0x1234ABCD;
        _ = Win32.SetWindowLongPtr(window, Win32.GwlpUserData, first);
        var read = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        Check("a window slot gives back what was put in it", read == first, $"0x{read:X}");

        var second = nint.Size == 8 ? unchecked((nint)0x00007FFE1234ABCDL) : unchecked((nint)0x7FFE1234);
        var replaced = Win32.SetWindowLongPtr(window, Win32.GwlpUserData, second);
        read = Win32.GetWindowLongPtr(window, Win32.GwlpUserData);
        Check("writing a slot answers the value it replaced", replaced == first, $"0x{replaced:X}");
        Check("a slot holds a value as wide as a pointer", read == second, $"0x{read:X}");

        const int classWindowProcedure = -24;
        var procedure = Win32.GetClassLongPtr(window, classWindowProcedure);
        var ours = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&OnMessage;
        Check("the window's class answers with its procedure", procedure == ours, $"0x{procedure:X} against 0x{ours:X}");
        _ = Win32.GetClassLongPtr(window, Win32.GclpHIconSm);
        Check("the window's class answers for its icon", true);

        var resized = 0;
        using (var subclass = FrameSubclass.Install(window, () => resized++))
        {
            if (Check("the frame subclass installs", subclass is not null))
            {
                var before = _sizeMessages;
                _ = Win32.SetWindowPos(window, 0, 0, 0, 800, 600, Win32.SwpNoMove | Win32.SwpNoZOrder | Win32.SwpNoActivate);
                Check("a resize calls the subclass back", resized > 0, $"{resized} call(s)");
                Check("and still reaches the window's own procedure", _sizeMessages > before, $"{_sizeMessages - before} message(s)");
            }
        }
    }

    /// <summary>
    /// The automation layer every call into the editor goes through, pointed at an object that is
    /// on every Windows: arguments of both kinds, an object coming back, a collection, type
    /// information, and a call that fails with a description.
    /// </summary>
    private static void DispatchCalls()
    {
        var hr = Win32.CLSIDFromProgID("Scripting.FileSystemObject", out var classId);
        if (!Check("the file system object has a class", hr >= 0, $"0x{hr:X8}"))
        {
            return;
        }

        hr = Win32.CoCreateInstance(in classId, 0, Win32.ClassContextInProcessServer, in DispatchId, out var pointer);
        if (!Check("it can be created", hr >= 0 && pointer != 0, $"0x{hr:X8}"))
        {
            return;
        }

        using var files = DispatchObject.Attach(pointer);
        if (!Check("it answers as an automation object", files is not null) || files is null)
        {
            return;
        }

        var baseName = files.CallToString("GetBaseName", @"C:\folder\name.txt");
        Check("a call with one string answers a string", baseName == "name", baseName);

        var built = files.CallToString("BuildPath", @"C:\a", "b");
        Check("a call with two strings answers a string", built == @"C:\a\b", built);

        Check("a member that is not there is reported as unknown", files.GetDispId("NoSuchMember") == DispId.Unknown);

        using (var windows = files.CallObject("GetSpecialFolder", 0))
        {
            if (Check("a call with a number answers an object", windows is not null) && windows is not null)
            {
                var name = windows.GetString("Name") ?? string.Empty;
                Check("the object's string property reads", name.Length > 0, name);
                Same("special folder", name);
                Same("its type", windows.TypeName() ?? "(none)");
                Same("its attributes", windows.GetInt32("Attributes").ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        using (var drives = files.GetObject("Drives"))
        {
            if (Check("a property answers a collection", drives is not null) && drives is not null)
            {
                var count = drives.GetInt32("Count");
                Check("the collection counts its items", count > 0, $"{count} drive(s)");
                Same("drives", count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        try
        {
            using var missing = files.CallObject("GetFolder", @"Q:\xlide-self-check\no-such-folder");
            Fail("a call that fails throws", "it answered instead");
        }
        catch (Exception ex)
        {
            Check("a call that fails with a bare code throws", ex.Message.Length > 0, ex.Message);
            Same("the bare failure", $"{ex.GetType().Name}: {ex.Message}");
        }

        DescribedFailure();
    }

    /// <summary>
    /// A failure that comes with a description, which the file system object never gives: its
    /// errors are bare codes. The XML parser describes its own, so its refusal of a malformed
    /// query is what reads the three strings out of an exception record and frees them.
    /// </summary>
    private static void DescribedFailure()
    {
        var hr = Win32.CLSIDFromProgID("Msxml2.DOMDocument.6.0", out var classId);
        if (!Check("the XML document has a class", hr >= 0, $"0x{hr:X8}"))
        {
            return;
        }

        hr = Win32.CoCreateInstance(in classId, 0, Win32.ClassContextInProcessServer, in DispatchId, out var pointer);
        using var document = hr >= 0 ? DispatchObject.Attach(pointer) : null;
        if (!Check("it can be created", document is not null, $"0x{hr:X8}") || document is null)
        {
            return;
        }

        try
        {
            using var node = document.CallObject("selectSingleNode", "///");
            Fail("a call that fails with a description throws", "it answered instead");
        }
        catch (Exception ex)
        {
            var words = ex.Message.ReplaceLineEndings(" ");
            Check("a call that fails with a description throws, in the object's own words",
                ex is InvalidOperationException && words.Length > 0, $"{ex.GetType().Name}: {words}");
            Same("the described failure", words);
        }
    }

    /// <summary>Every type and member of a library Windows always has, as the Object Browser reads them.</summary>
    private static void TypeLibraries()
    {
        var path = Path.Combine(Environment.SystemDirectory, "stdole2.tlb");
        using var catalog = new TypeLibraryCatalog();
        catalog.AddLibrary("stdole", "OLE Automation", path);

        var types = catalog.TypesOf("stdole");
        if (!Check("the library loads and lists its types", types is { Count: > 0 }, path) || types is null)
        {
            return;
        }

        var rows = new List<string>();
        var members = 0;
        foreach (var type in types)
        {
            rows.Add($"{type.Kind} {type.Name}");
            foreach (var member in catalog.MembersOf("stdole", type.Name) ?? [])
            {
                members++;
                rows.Add($"  {member.Kind} {member.Name} {member.Signature} {member.Description}");
            }
        }

        Check("its types list their members", members > 0, $"{types.Count} type(s), {members} member(s)");
        Same("types", types.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Same("members", members.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Same("every row", Digest(rows));
    }

    /// <summary>The font walk, which is a callback from GDI into this library.</summary>
    private static void Fonts()
    {
        var fonts = InstalledFonts.All;
        Check("the font walk calls back with families", fonts.Count > 0, $"{fonts.Count} family(ies)");
        Same("families", fonts.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Same("every name", Digest(fonts));
    }

    /// <summary>The product's icon, loaded as a picture and turned back into pixels.</summary>
    private static void Pictures()
    {
        var icon = Path.Combine(ShimModule.SharedDirectory ?? string.Empty, "xlide.ico");
        if (!Check("the icon is where the editor's frame would look", File.Exists(icon), icon))
        {
            return;
        }

        using var picture = PictureBytes.FromPath(icon);
        var size = PictureBytes.SizeOf(picture);
        Check("the picture has a size", size is { Width: > 0, Height: > 0 }, size is null ? "none" : $"{size.Value.Width} x {size.Value.Height}");

        var uri = PictureBytes.DataUriOf(picture);
        Check("the picture turns into pixels", uri is { Length: > 100 }, uri is null ? "none" : $"{uri.Length} character(s)");
        Same("what it is", PictureBytes.Describe(picture));
        Same("its size", size is null ? "none" : $"{size.Value.Width} x {size.Value.Height}");
        Same("its pixels", uri is null ? "none" : Digest([uri]));
    }

    /// <summary>
    /// The four kinds of value the Locals and Immediate readers take through UI Automation, each
    /// read into a variant with a word on either side of it that the callee must not touch.
    /// </summary>
    private static void UiAutomationReads()
    {
        var window = EnsureWindow();

        var hr = Win32.CoCreateInstance(
            in UiAutomationIds.AutomationClass, 0, Win32.ClassContextInProcessServer, in UiAutomationIds.Automation, out var automationPointer);
        if (!Check("the accessibility interface can be created", hr >= 0 && automationPointer != 0, $"0x{hr:X8}"))
        {
            return;
        }

        using var automation = ComHandle<IUIAutomation>.Own(automationPointer);
        if (automation is null || automation.Target.ElementFromHandle(window, out var elementPointer) < 0 || elementPointer == 0)
        {
            Fail("the window has an accessible element", "it has none");
            return;
        }

        using var element = ComHandle<IUIAutomationElement>.Own(elementPointer);
        if (!Check("the window has an accessible element", element is not null) || element is null)
        {
            return;
        }

        var name = ReadGuarded(element.Target, UiAutomationIds.NameProperty, "its name");
        var text = name.TakeString();
        Check("its name is the window's title", text == Title, text ?? "(none)");

        const int windowControl = 50032;
        var kind = ReadGuarded(element.Target, UiAutomationIds.ControlTypeProperty, "its control type");
        Check("its control type is a window", kind.AsInt32() == windowControl, kind.AsInt32().ToString(System.Globalization.CultureInfo.InvariantCulture));

        var bounds = ReadGuarded(element.Target, UiAutomationIds.BoundingRectangleProperty, "its bounds");
        var rectangle = bounds.TakeRectangle();
        Check("its bounds are four numbers", rectangle is not null, rectangle is null ? "none" : $"{rectangle.Value.Width} x {rectangle.Value.Height}");

        var runtime = ReadGuarded(element.Target, UiAutomationIds.RuntimeIdProperty, "its runtime id");
        var id = runtime.TakeRuntimeId();
        Check("its runtime id is a list of numbers", id is { Length: > 0 }, id ?? "none");
    }

    private const ulong Canary = 0xC0FFEE11DEC0DE22UL;

    [StructLayout(LayoutKind.Sequential)]
    private struct GuardedVariant
    {
        public ulong Before;
        public UiVariant Value;
        public ulong After;
    }

    private static UiVariant ReadGuarded(IUIAutomationElement element, int property, string what)
    {
        var guarded = new GuardedVariant { Before = Canary, After = Canary };
        var hr = element.GetCurrentPropertyValue(property, out guarded.Value);
        Check($"{what} reads", hr >= 0, $"0x{hr:X8}");
        Check($"{what} is written inside its variant", guarded.Before == Canary && guarded.After == Canary,
            $"before 0x{guarded.Before:X16}, after 0x{guarded.After:X16}");
        return guarded.Value;
    }

    /// <summary>The job the engine is put in, so it cannot outlive the host.</summary>
    private static void JobObject()
    {
        using var job = KillOnCloseJob.Create();
        Check("a kill-on-close job can be made", job is not null);
    }

    /// <summary>The engine, started and asked one question, as a session starts it.</summary>
    private static void EngineRoundTrip()
    {
        var path = AnalysisService.EnginePath();
        if (!Check("the engine is where a session would look", File.Exists(path), path))
        {
            return;
        }

        using var patience = new CancellationTokenSource(TimeSpan.FromSeconds(90));
        var client = EngineClient.StartAsync(path, patience.Token).GetAwaiter().GetResult();
        if (!Check("the engine starts and connects", client is { IsRunning: true }) || client is null)
        {
            return;
        }

        try
        {
            var rules = client.RulesAsync(patience.Token).GetAwaiter().GetResult();
            var count = rules?.Rules.Length ?? 0;
            Check("the engine answers a request", count > 0, $"{count} analysis rule(s)");
            Same("analysis rules", count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        finally
        {
            client.DisposeAsync().AsTask().GetAwaiter().GetResult();
        }
    }

    /// <summary>
    /// The browser, brought up inside the window through the product's own surface: the loader of
    /// this process's width, the environment and controller callbacks, script in and an answer
    /// out, and the editor page's first message back.
    /// </summary>
    private static void BrowserSurface()
    {
        var shared = ShimModule.SharedDirectory;
        var entry = shared is null ? null : WebViewPaths.EditorEntryDocument(shared);
        if (!Check("the editor page is where the surface would look", entry is not null && File.Exists(entry), entry ?? "unknown"))
        {
            return;
        }

        var window = EnsureWindow();
        string? firstMessage = null;

        var surface = WebView2Surface.Start(window, new PixelRect(0, 0, 800, 600));
        if (!Check("the browser surface starts", surface is not null) || surface is null)
        {
            return;
        }

        try
        {
            surface.DebugName = "self-check";
            surface.MessageReceived = message => firstMessage ??= message;

            var asked = false;
            var failure = 0;
            string? answer = null;

            Pump(
                () =>
                {
                    // Refused until the browser exists, which takes two callbacks from now.
                    asked = asked || surface.ExecuteScript("6 * 7", (code, json) =>
                    {
                        failure = code;
                        answer = json;
                    });

                    return answer is not null || failure != 0;
                },
                TimeSpan.FromSeconds(60));

            Check("the browser comes up, runs script and answers", answer == "42",
                asked ? $"answer {answer ?? "(none)"}, code 0x{failure:X8}" : "the browser never came up");

            Pump(() => firstMessage is not null, TimeSpan.FromSeconds(60));
            Check("the editor page posts its first message", firstMessage is not null,
                firstMessage is null ? "none within a minute" : firstMessage[..Math.Min(firstMessage.Length, 60)]);
        }
        finally
        {
            surface.Dispose();

            // The browser's own teardown is posted, so it is given a moment to run.
            Pump(static () => false, TimeSpan.FromMilliseconds(500));
        }
    }

    // ---- the window the stages share ----------------------------------------------------------

    private static nint EnsureWindow()
    {
        if (_window != 0)
        {
            return _window;
        }

        fixed (char* className = ClassName)
        {
            var windowClass = new WndClassExW
            {
                Size = (uint)sizeof(WndClassExW),
                WindowProc = (nint)(delegate* unmanaged<nint, uint, nint, nint, nint>)&OnMessage,
                Instance = ShimModule.Handle,
                ClassName = className,
            };

            if (Win32.RegisterClassEx(&windowClass) == 0)
            {
                var error = Marshal.GetLastPInvokeError();
                if (error != Win32.ErrorClassAlreadyExists)
                {
                    throw new InvalidOperationException($"the window class would not register, error {error}");
                }
            }
        }

        // Never shown. Everything here works on a window nobody can see, and a self-check that
        // put one on the developer's screen would be a nuisance every time it ran.
        _window = Win32.CreateWindowEx(
            0, ClassName, Title, Win32.WsOverlappedWindow, 0, 0, 640, 480, 0, 0, ShimModule.Handle, 0);
        if (_window == 0)
        {
            throw new InvalidOperationException($"the window would not be made, error {Marshal.GetLastPInvokeError()}");
        }

        return _window;
    }

    private static void DestroyWindow()
    {
        if (_window == 0)
        {
            return;
        }

        var window = _window;
        _window = 0;
        if (!Win32.DestroyWindow(window))
        {
            Fail("the window goes when asked", $"error {Marshal.GetLastPInvokeError()}");
        }
    }

    private const uint WmSize = 0x0005;

    [UnmanagedCallersOnly]
    private static nint OnMessage(nint window, uint message, nint wParam, nint lParam)
    {
        if (message == WmSize)
        {
            _sizeMessages++;
        }

        return Win32.DefWindowProc(window, message, wParam, lParam);
    }

    // ---- a message loop, which a host normally lends us ---------------------------------------

    private const uint PeekRemove = 0x0001;
    private const uint WakeOnAnyInput = 0x04FF;
    private const uint WaitForInputAvailable = 0x0004;

    /// <summary>
    /// Dispatches this thread's messages until <paramref name="done"/> says so or the time is up.
    /// The browser's callbacks arrive as messages on the thread that asked, and in a host that
    /// thread is already in a loop; here nothing is, so the loop is ours.
    /// </summary>
    private static void Pump(Func<bool> done, TimeSpan patience)
    {
        var until = Environment.TickCount64 + (long)patience.TotalMilliseconds;
        Message message;

        while (!done() && Environment.TickCount64 < until)
        {
            while (PeekMessage(&message, 0, 0, 0, PeekRemove))
            {
                _ = TranslateMessage(&message);
                _ = DispatchMessage(&message);
            }

            _ = MsgWaitForMultipleObjectsEx(0, null, 50, WakeOnAnyInput, WaitForInputAvailable);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Message
    {
        public nint Window;
        public uint Id;
        public nuint WParam;
        public nint LParam;
        public uint Time;
        public int X;
        public int Y;
        public uint Private;
    }

    [LibraryImport("user32.dll", EntryPoint = "PeekMessageW")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool PeekMessage(Message* message, nint window, uint first, uint last, uint remove);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool TranslateMessage(Message* message);

    [LibraryImport("user32.dll", EntryPoint = "DispatchMessageW")]
    private static partial nint DispatchMessage(Message* message);

    [LibraryImport("user32.dll")]
    private static partial uint MsgWaitForMultipleObjectsEx(uint count, nint* handles, uint milliseconds, uint wakeMask, uint flags);
}
#endif
