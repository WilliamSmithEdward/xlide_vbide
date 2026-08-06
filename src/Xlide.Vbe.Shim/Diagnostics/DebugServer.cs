#if DEBUG
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
using Xlide.Vbe.Core;

namespace Xlide.Vbe.Shim.Diagnostics;

/// <summary>
/// The dev build's local door: a token-gated HTTP endpoint on 127.0.0.1 that answers
/// questions about the running session and acts on it by name, so the harness can ask and
/// act semantically instead of posting mouse messages at measured pixels (developer,
/// 2026-08-05). The route shapes are deliberately product-grade: the developer's stated
/// vision is xlide_vscode and xlide_vbide talking over this api one day, so routes stay
/// noun-shaped and stable, and the discovery file carries a schema number.
///
/// This file compiles ONLY in Debug. A release build carries no server, no port, and no
/// route; the product's contract is untouched. The port is random, the token is random,
/// and both are announced in the shim log and in a per-process discovery file that the
/// session deletes when it stops.
///
/// The server owns sockets and nothing else. Every answer comes from the delegate its
/// owner provides; marshaling to the host thread is the owner's business, and a slow
/// answer is the owner saying so, not this class guessing.
/// </summary>
internal sealed class DebugServer : IDisposable
{
    /// <summary>One parsed request: the route, its query arguments, and its body when it sent one.</summary>
    public sealed record DebugRequest(string Route, IReadOnlyDictionary<string, string> Query, string Body);

    /// <summary>One answer: its content type and its bytes. JSON is the norm; capture is not.</summary>
    public sealed record DebugReply(string ContentType, byte[] Bytes)
    {
        public static DebugReply Json(string body) =>
            new("application/json; charset=utf-8", Encoding.UTF8.GetBytes(body));
    }

    /// <summary>The discovery file's schema number; bumped on breaking shape changes.</summary>
    private const int ApiVersion = 1;

    private readonly TcpListener _listener;
    private readonly Func<DebugRequest, DebugReply> _answer;
    private readonly string _token;
    private readonly string _discoveryPath;
    private volatile bool _stopped;

    private DebugServer(TcpListener listener, string token, string discoveryPath, Func<DebugRequest, DebugReply> answer)
    {
        _listener = listener;
        _token = token;
        _discoveryPath = discoveryPath;
        _answer = answer;
    }

    public static DebugServer? Start(Func<DebugRequest, DebugReply> answer)
    {
        try
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();

            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var token = Convert.ToHexString(RandomNumberGenerator.GetBytes(8)).ToLowerInvariant();

            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ProductIdentity.DataFolderName);
            Directory.CreateDirectory(directory);
            SweepDeadDiscoveryFiles(directory);
            var discoveryPath = Path.Combine(directory, $"debug-api-{Environment.ProcessId}.json");

            var server = new DebugServer(listener, token, discoveryPath, answer);

            // Everything a client needs to pick THIS instance from several running Excels:
            // the file is per pid, and the state route answers with the shown workbook. A
            // client discovers by globbing debug-api-*.json and probing each state.
            File.WriteAllText(discoveryPath,
                $"{{\"api\":{ApiVersion},\"port\":{port},\"token\":\"{token}\"," +
                $"\"devtoolsPort\":{WebView.WebView2Surface.DevToolsPort},\"pid\":{Environment.ProcessId}," +
                $"\"startedAt\":\"{DateTime.UtcNow:O}\"}}");

            _ = Task.Run(server.Loop);
            Log.Info($"debug api: listening on 127.0.0.1:{port}/{token} (dev build only)");
            return server;
        }
        catch (Exception ex)
        {
            Log.Info($"debug api: could not start ({ex.GetType().Name}: {ex.Message})");
            return null;
        }
    }

    /// <summary>
    /// Deletes discovery files whose process is gone. Clean stops delete their own file, but
    /// a killed Excel leaves one naming a dead port, and a client that globs the directory
    /// would probe corpses. Each new session sweeps for all of them.
    /// </summary>
    private static void SweepDeadDiscoveryFiles(string directory)
    {
        foreach (var stale in Directory.EnumerateFiles(directory, "debug-api-*.json"))
        {
            try
            {
                var name = Path.GetFileNameWithoutExtension(stale);
                if (int.TryParse(name["debug-api-".Length..], out var pid) && pid != Environment.ProcessId)
                {
                    using var process = System.Diagnostics.Process.GetProcessById(pid);
                }
            }
            catch (ArgumentException)
            {
                // No such process: the file names a corpse.
                try
                {
                    File.Delete(stale);
                }
                catch
                {
                    // Another session may be sweeping the same file; losing the race is fine.
                }
            }
            catch
            {
                // A file that cannot be judged is left alone.
            }
        }
    }

    private async Task Loop()
    {
        while (!_stopped)
        {
            TcpClient client;
            try
            {
                client = await _listener.AcceptTcpClientAsync().ConfigureAwait(false);
            }
            catch
            {
                // The listener stopping is the one way out of Accept, and it is not news.
                return;
            }

            _ = Task.Run(() => Serve(client));
        }
    }

    private void Serve(TcpClient client)
    {
        try
        {
            using var owned = client;
            owned.NoDelay = true;
            using var stream = owned.GetStream();
            stream.ReadTimeout = 5000;
            stream.WriteTimeout = 15000;

            if (ReadRequest(stream) is not { } request)
            {
                return;
            }

            var (method, target, requestBody) = request;
            if (method is not ("GET" or "POST"))
            {
                WriteReply(stream, "405 Method Not Allowed", DebugReply.Json("""{"error":"GET or POST"}"""));
                return;
            }

            var path = target;
            var query = new Dictionary<string, string>(StringComparer.Ordinal);
            var mark = path.IndexOf('?', StringComparison.Ordinal);
            if (mark >= 0)
            {
                foreach (var pair in path[(mark + 1)..].Split('&'))
                {
                    var eq = pair.IndexOf('=', StringComparison.Ordinal);
                    if (eq > 0)
                    {
                        query[Uri.UnescapeDataString(pair[..eq])] = Uri.UnescapeDataString(pair[(eq + 1)..]);
                    }
                }

                path = path[..mark];
            }

            var prefix = $"/{_token}/";
            if (!path.StartsWith(prefix, StringComparison.Ordinal))
            {
                WriteReply(stream, "404 Not Found", DebugReply.Json("""{"error":"unknown"}"""));
                return;
            }

            DebugReply body;
            try
            {
                body = _answer(new DebugRequest(path[prefix.Length..], query, requestBody));
            }
            catch (Exception ex)
            {
                WriteReply(stream, "500 Internal Server Error", DebugReply.Json($"{{\"error\":\"{ex.GetType().Name}\"}}"));
                return;
            }

            WriteReply(stream, "200 OK", body);
        }
        catch
        {
            // A connection the client abandoned is routine.
        }
    }

    private static (string Method, string Target, string Body)? ReadRequest(NetworkStream stream)
    {
        // Headers first; a body follows only when Content-Length says so, and it is capped
        // hard - the biggest legitimate body is a module's text.
        var buffer = new byte[1 << 20];
        var held = 0;
        var headerEnd = -1;

        while (held < buffer.Length)
        {
            var read = stream.Read(buffer, held, buffer.Length - held);
            if (read <= 0)
            {
                return null;
            }

            held += read;
            headerEnd = buffer.AsSpan(0, held).IndexOf("\r\n\r\n"u8);
            if (headerEnd >= 0)
            {
                break;
            }
        }

        if (headerEnd < 0)
        {
            return null;
        }

        var head = Encoding.ASCII.GetString(buffer, 0, headerEnd);
        var lineEnd = head.IndexOf("\r\n", StringComparison.Ordinal);
        var requestLine = lineEnd < 0 ? head : head[..lineEnd];
        var parts = requestLine.Split(' ');
        if (parts.Length != 3)
        {
            return null;
        }

        var contentLength = 0;
        foreach (var header in head.Split("\r\n").Skip(1))
        {
            if (header.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(header["Content-Length:".Length..].Trim(), out var declared))
            {
                contentLength = Math.Clamp(declared, 0, buffer.Length - (headerEnd + 4));
            }
        }

        var bodyStart = headerEnd + 4;
        while (held - bodyStart < contentLength)
        {
            var read = stream.Read(buffer, held, buffer.Length - held);
            if (read <= 0)
            {
                break;
            }

            held += read;
        }

        var body = contentLength > 0
            ? Encoding.UTF8.GetString(buffer, bodyStart, Math.Min(contentLength, held - bodyStart))
            : string.Empty;

        return (parts[0], parts[1], body);
    }

    private static void WriteReply(NetworkStream stream, string status, DebugReply reply)
    {
        var header = $"HTTP/1.1 {status}\r\n"
            + $"Content-Type: {reply.ContentType}\r\n"
            + $"Content-Length: {reply.Bytes.Length}\r\n"
            + "Cache-Control: no-store\r\n"
            + "Connection: close\r\n"
            + "\r\n";
        stream.Write(Encoding.ASCII.GetBytes(header));
        stream.Write(reply.Bytes);
    }

    public void Dispose()
    {
        _stopped = true;
        try
        {
            _listener.Stop();
        }
        catch
        {
            // Stopping twice, or after a socket fault, changes nothing.
        }

        try
        {
            File.Delete(_discoveryPath);
        }
        catch
        {
            // A leftover discovery file names a dead port; the next session overwrites it.
        }
    }
}

/// <summary>
/// Renders a native window into a bottom-up 32bpp BMP for the capture route. Plain GDI
/// through its own imports - the shim is ahead-of-time compiled and carries no drawing
/// library. PW_RENDERFULLCONTENT is what makes a composited browser child appear.
/// </summary>
internal static partial class DebugCapture
{
    [LibraryImport("user32.dll")]
    private static partial nint GetDC(nint window);

    [LibraryImport("user32.dll")]
    private static partial int ReleaseDC(nint window, nint dc);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetWindowRect(nint window, out Rect rect);

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool PrintWindow(nint window, nint dc, uint flags);

    [LibraryImport("gdi32.dll")]
    private static partial nint CreateCompatibleDC(nint dc);

    [LibraryImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool DeleteDC(nint dc);

    [LibraryImport("gdi32.dll")]
    private static partial nint SelectObject(nint dc, nint gdiObject);

    [LibraryImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool DeleteObject(nint gdiObject);

    [LibraryImport("gdi32.dll")]
    private static unsafe partial nint CreateDIBSection(nint dc, BitmapInfo* info, uint usage, out nint bits, nint section, uint offset);

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BitmapInfo
    {
        public uint Size;
        public int Width;
        public int Height;
        public ushort Planes;
        public ushort BitCount;
        public uint Compression;
        public uint SizeImage;
        public int XPelsPerMeter;
        public int YPelsPerMeter;
        public uint ClrUsed;
        public uint ClrImportant;
    }

    /// <summary>The window as a BMP file's bytes, or null when it would not render.</summary>
    public static unsafe byte[]? CaptureBmp(nint window)
    {
        if (window == 0 || !GetWindowRect(window, out var rect))
        {
            return null;
        }

        var width = rect.Right - rect.Left;
        var height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0)
        {
            return null;
        }

        var screen = GetDC(0);
        var memory = CreateCompatibleDC(screen);
        nint bitmap = 0;
        nint previous = 0;

        try
        {
            var info = new BitmapInfo
            {
                Size = (uint)sizeof(BitmapInfo),
                Width = width,
                Height = height,
                Planes = 1,
                BitCount = 32,
                Compression = 0,
            };

            bitmap = CreateDIBSection(memory, &info, 0, out var bits, 0, 0);
            if (bitmap == 0)
            {
                return null;
            }

            previous = SelectObject(memory, bitmap);
            if (!PrintWindow(window, memory, 0x2))
            {
                return null;
            }

            var pixelBytes = width * height * 4;
            var file = new byte[14 + 40 + pixelBytes];

            // BITMAPFILEHEADER: BM, file size, reserved, offset to pixels.
            file[0] = (byte)'B';
            file[1] = (byte)'M';
            BitConverter.GetBytes(file.Length).CopyTo(file, 2);
            BitConverter.GetBytes(14 + 40).CopyTo(file, 10);

            // BITMAPINFOHEADER, the same shape the section was made with.
            BitConverter.GetBytes(40).CopyTo(file, 14);
            BitConverter.GetBytes(width).CopyTo(file, 18);
            BitConverter.GetBytes(height).CopyTo(file, 22);
            BitConverter.GetBytes((ushort)1).CopyTo(file, 26);
            BitConverter.GetBytes((ushort)32).CopyTo(file, 28);
            BitConverter.GetBytes(pixelBytes).CopyTo(file, 34);

            new ReadOnlySpan<byte>((void*)bits, pixelBytes).CopyTo(file.AsSpan(54));
            return file;
        }
        finally
        {
            if (previous != 0)
            {
                SelectObject(memory, previous);
            }

            if (bitmap != 0)
            {
                DeleteObject(bitmap);
            }

            DeleteDC(memory);
            _ = ReleaseDC(0, screen);
        }
    }
}

/// <summary>What GET state answers.</summary>
public sealed record DebugStateReply(
    [property: JsonPropertyName("configuration")] string Configuration,
    [property: JsonPropertyName("shownModule")] string? ShownModule,
    [property: JsonPropertyName("shownProject")] string? ShownProject,
    [property: JsonPropertyName("debugMode")] string? DebugMode,
    [property: JsonPropertyName("hasUnwrittenEdits")] bool HasUnwrittenEdits,
    [property: JsonPropertyName("engineUp")] bool EngineUp,
    [property: JsonPropertyName("frame")] string Frame,
    [property: JsonPropertyName("frameRect")] string FrameRect,
    [property: JsonPropertyName("documentArea")] string DocumentArea,
    [property: JsonPropertyName("documentAreaRect")] string DocumentAreaRect,
    [property: JsonPropertyName("paletteOpen")] bool PaletteOpen,
    [property: JsonPropertyName("paletteVisible")] bool PaletteVisible,
    [property: JsonPropertyName("surfaceReady")] bool SurfaceReady,
    [property: JsonPropertyName("devtoolsPort")] int DevToolsPort);

/// <summary>One native editor window, as GET windows lists them.</summary>
public sealed record DebugWindowRow(
    [property: JsonPropertyName("type")] int Type,
    [property: JsonPropertyName("caption")] string Caption,
    [property: JsonPropertyName("visible")] bool Visible);

public sealed record DebugWindowsReply(
    [property: JsonPropertyName("windows")] DebugWindowRow[] Windows);

/// <summary>What POST command, placement, breakpoint, and immediate answer.</summary>
public sealed record DebugCommandReply(
    [property: JsonPropertyName("ran")] bool Ran,
    [property: JsonPropertyName("command")] int Command);

/// <summary>A slice of the shim log, from a byte offset the caller advances.</summary>
public sealed record DebugLogReply(
    [property: JsonPropertyName("lines")] string[] Lines,
    [property: JsonPropertyName("next")] long Next);

/// <summary>One message over a surface's wire, as GET messages lists them.</summary>
public sealed record DebugMessageRow(
    [property: JsonPropertyName("seq")] long Seq,
    [property: JsonPropertyName("at")] string At,
    [property: JsonPropertyName("surface")] string Surface,
    [property: JsonPropertyName("direction")] string Direction,
    [property: JsonPropertyName("text")] string Text);

public sealed record DebugMessagesReply(
    [property: JsonPropertyName("messages")] DebugMessageRow[] Messages);

public sealed record DebugErrorReply(
    [property: JsonPropertyName("error")] string Error);

/// <summary>
/// A timeout that knows why. The editor's commonest reason for one is a modal dialog owning
/// the host thread, and window enumeration can see that without the thread the dialog is
/// holding - so a blocked answer names the dialog, its buttons, how long the host thread has
/// been silent, and whether the door answered a dialog it had raised itself.
/// </summary>
public sealed record DebugBlockedReply(
    [property: JsonPropertyName("error")] string Error,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs,
    [property: JsonPropertyName("blockedBy")] string? BlockedBy,
    [property: JsonPropertyName("buttons")] string[] Buttons,
    [property: JsonPropertyName("dismissed")] string? Dismissed,
    [property: JsonPropertyName("retried")] bool Retried);

/// <summary>One native dialog standing in the process.</summary>
public sealed record DebugDialogRow(
    [property: JsonPropertyName("window")] string Window,
    [property: JsonPropertyName("caption")] string Caption,
    [property: JsonPropertyName("buttons")] string[] Buttons,
    [property: JsonPropertyName("enabled")] bool Enabled);

public sealed record DebugDialogsReply(
    [property: JsonPropertyName("dialogs")] DebugDialogRow[] Dialogs,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs);

/// <summary>
/// The start-of-session questions, answered together: is this session running the code that
/// was just built, and is everything that should be attached attached. Findings is empty when
/// nothing is wrong, which is the only summary worth trusting.
/// </summary>
public sealed record DebugDoctorReply(
    [property: JsonPropertyName("healthy")] bool Healthy,
    [property: JsonPropertyName("shimPath")] string ShimPath,
    [property: JsonPropertyName("shimBuiltUtc")] string ShimBuiltUtc,
    [property: JsonPropertyName("bundleBuiltUtc")] string BundleBuiltUtc,
    [property: JsonPropertyName("pageBuildStamp")] string PageBuildStamp,
    [property: JsonPropertyName("engineUp")] bool EngineUp,
    [property: JsonPropertyName("ghostReadersUp")] bool GhostReadersUp,
    [property: JsonPropertyName("surfaceReady")] bool SurfaceReady,
    [property: JsonPropertyName("findings")] string[] Findings);

/// <summary>Recent raw durations, for percentiles a running maximum cannot give.</summary>
public sealed record DebugPerfReply(
    [property: JsonPropertyName("placementMs")] long[] PlacementMs,
    [property: JsonPropertyName("marshalMs")] long[] MarshalMs,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs);

/// <summary>What a script run in the page answered with; result is JSON as the browser encodes it.</summary>
public sealed record DebugEvalReply(
    [property: JsonPropertyName("answered")] bool Answered,
    [property: JsonPropertyName("errorCode")] int ErrorCode,
    [property: JsonPropertyName("result")] string Result);

/// <summary>The Locals ghost's feed as data, for the debug-side suite.</summary>
public sealed record DebugLocalsReply(
    [property: JsonPropertyName("context")] string? Context,
    [property: JsonPropertyName("rows")] Editor.SurfaceLocalRow[] Rows);

/// <summary>The Watch ghost's feed as data.</summary>
public sealed record DebugWatchesReply(
    [property: JsonPropertyName("stopped")] bool Stopped,
    [property: JsonPropertyName("rows")] Editor.SurfaceWatchRow[] Rows);

/// <summary>One analyzer finding, as GET problems lists them.</summary>
public sealed record DebugFindingRow(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message);

public sealed record DebugProblemsReply(
    [property: JsonPropertyName("findings")] DebugFindingRow[] Findings);

/// <summary>A module's text, read through the session's own reader.</summary>
public sealed record DebugModuleReply(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("text")] string Text);

/// <summary>The counters behind the regression classes this product has lived through.</summary>
public sealed record DebugStatsReply(
    [property: JsonPropertyName("uptimeSeconds")] long UptimeSeconds,
    [property: JsonPropertyName("managedMemoryBytes")] long ManagedMemoryBytes,
    [property: JsonPropertyName("workingSetBytes")] long WorkingSetBytes,
    [property: JsonPropertyName("handleCount")] int HandleCount,
    [property: JsonPropertyName("gcCounts")] int[] GcCounts,
    [property: JsonPropertyName("placementFullPasses")] long PlacementFullPasses,
    [property: JsonPropertyName("placementFastPasses")] long PlacementFastPasses,
    [property: JsonPropertyName("placementLastMs")] long PlacementLastMs,
    [property: JsonPropertyName("placementMaxMs")] long PlacementMaxMs,
    [property: JsonPropertyName("marshalCount")] long MarshalCount,
    [property: JsonPropertyName("marshalLastMs")] long MarshalLastMs,
    [property: JsonPropertyName("marshalMaxMs")] long MarshalMaxMs,
    [property: JsonPropertyName("logLines")] long LogLines,
    [property: JsonPropertyName("pollIntervalMs")] long PollIntervalMs,
    [property: JsonPropertyName("messagesToPage")] long MessagesToPage,
    [property: JsonPropertyName("messagesToHost")] long MessagesToHost,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs,
    [property: JsonPropertyName("dialogsStanding")] int DialogsStanding);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(DebugStateReply))]
[JsonSerializable(typeof(DebugWindowRow))]
[JsonSerializable(typeof(DebugWindowsReply))]
[JsonSerializable(typeof(DebugCommandReply))]
[JsonSerializable(typeof(DebugLogReply))]
[JsonSerializable(typeof(DebugMessageRow))]
[JsonSerializable(typeof(DebugMessagesReply))]
[JsonSerializable(typeof(DebugLocalsReply))]
[JsonSerializable(typeof(DebugWatchesReply))]
[JsonSerializable(typeof(DebugFindingRow))]
[JsonSerializable(typeof(DebugProblemsReply))]
[JsonSerializable(typeof(DebugModuleReply))]
[JsonSerializable(typeof(DebugStatsReply))]
[JsonSerializable(typeof(DebugErrorReply))]
[JsonSerializable(typeof(DebugBlockedReply))]
[JsonSerializable(typeof(DebugDialogRow))]
[JsonSerializable(typeof(DebugDialogsReply))]
[JsonSerializable(typeof(DebugEvalReply))]
[JsonSerializable(typeof(DebugDoctorReply))]
[JsonSerializable(typeof(DebugPerfReply))]
internal sealed partial class DebugJsonContext : JsonSerializerContext;
#endif
