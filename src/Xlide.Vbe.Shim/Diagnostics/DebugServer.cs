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

            var route = path[prefix.Length..];
            RecordRequest(method, route, query);

            DebugReply body;
            var servedFrom = Environment.TickCount64;
            try
            {
                body = _answer(new DebugRequest(route, query, requestBody));
            }
            catch (Exception ex)
            {
                WriteReply(stream, "500 Internal Server Error", DebugReply.Json($"{{\"error\":\"{ex.GetType().Name}\"}}"));
                return;
            }

            // Timed around the answer, not around the write: a slow client reading its bytes
            // is not a slow route, and conflating the two sends a perf hunt at the network.
            RecordDuration(route, Environment.TickCount64 - servedFrom);

            WriteReply(stream, "200 OK", body);
        }
        catch
        {
            // A connection the client abandoned is routine.
        }
    }

    /*
     * Every request this door has served, so a session that went somewhere interesting can be
     * replayed instead of remembered. What actually happens after a live investigation is
     * that the useful sequence is reconstructed from a scrollback, badly; this keeps it.
     * The history route hands it back as a runnable script. Bounded, and no bodies: a module
     * write's payload is large and belongs to the workbook, not to a transcript.
     */
    private const int HistoryKept = 300;
    private static readonly Queue<string> History = new(HistoryKept + 1);
    private static readonly Lock HistoryGate = new();

    private static void RecordRequest(string method, string route, Dictionary<string, string> query)
    {
        if (route is "history" or "log" or "journal" or "state" or "dialogs")
        {
            // The routes a probe polls would otherwise be the whole transcript.
            return;
        }

        var arguments = query.Count == 0
            ? string.Empty
            : "?" + string.Join("&", query.Select(pair => $"{pair.Key}={pair.Value}"));

        lock (HistoryGate)
        {
            History.Enqueue($"{method} {route}{arguments}");
            while (History.Count > HistoryKept)
            {
                History.Dequeue();
            }
        }
    }

    /// <summary>What this door has been asked to do, oldest first.</summary>
    public static string[] Requests()
    {
        lock (HistoryGate)
        {
            return [.. History];
        }
    }

    /*
     * How long each route takes, in aggregate.
     *
     * The door is the instrument every other measurement is taken through, so a route that has
     * quietly become slow makes everything measured with it look slow, and there was no way to
     * notice. Ranked by TOTAL rather than worst case: the route to look at is the one a session
     * actually spends its seconds in.
     */
    private static readonly Dictionary<string, (long Count, long TotalMs, long MaxMs)> Durations =
        new(StringComparer.Ordinal);

    private static void RecordDuration(string route, long elapsedMs)
    {
        lock (HistoryGate)
        {
            var (count, total, max) = Durations.TryGetValue(route, out var seen) ? seen : (0, 0, 0);
            Durations[route] = (count + 1, total + elapsedMs, Math.Max(max, elapsedMs));
        }
    }

    /// <summary>Each route's count, total and worst, slowest total first.</summary>
    public static DebugRouteCost[] RouteCosts()
    {
        lock (HistoryGate)
        {
            return [.. Durations
                .Select(entry => new DebugRouteCost(
                    entry.Key, entry.Value.Count, entry.Value.TotalMs, entry.Value.MaxMs))
                .OrderByDescending(row => row.TotalMs)];
        }
    }

    /// <summary>
    /// The largest request body this door will read: 32 MB.
    ///
    /// Sized from what a module can legally be rather than from what is convenient. VBA holds up
    /// to 65,534 lines in one module, which at the shape of ordinary code is around 13 MB, and
    /// this is a BYTE count so a Cyrillic or CJK module of the same length costs two to three
    /// times that. Anything beyond is refused loudly rather than truncated quietly, which is the
    /// property this reader exists to keep.
    /// </summary>
    private const int LargestBody = 32 * 1024 * 1024;

    private static (string Method, string Target, string Body)? ReadRequest(NetworkStream stream)
    {
        // Headers first, into a buffer sized for headers. The BODY is read into its own buffer,
        // sized from Content-Length, because the two have nothing to do with each other and
        // treating them as one thing set the body's ceiling to whatever was left over after the
        // headers.
        //
        // That ceiling was 1 MB, and the biggest legitimate body is a module's text: VBA holds up
        // to 65,534 lines in one module, which at the shape of ordinary code is some 13 MB. So a
        // module a developer can legally write was more than ten times what could be sent through
        // here, and the refusal it earned looked like `fetch failed` at the caller (2026-08-08,
        // found while building a fixture at the line ceiling).
        var buffer = new byte[64 * 1024];
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
        var tooLarge = false;
        foreach (var header in head.Split("\r\n").Skip(1))
        {
            if (header.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(header["Content-Length:".Length..].Trim(), out var declared))
            {
                // REFUSED rather than clamped. A body larger than the buffer used to be silently
                // shortened to whatever fitted, and the caller most likely to send one is
                // `module` POST — so an oversized module would be written to the workbook
                // TRUNCATED, losing the developer's code behind a reply that said it worked. "A
                // write that fails must fail in its reply, not only in the log" applies to the
                // door as much as it did to the writer (lessons-2026-08-07).
                //
                // Non-ASCII brings the ceiling closer than it looks: this is a BYTE count, and a
                // Cyrillic or CJK module costs two to three bytes a character.
                if (declared > LargestBody)
                {
                    tooLarge = true;
                }

                contentLength = Math.Clamp(declared, 0, LargestBody);
            }
        }

        if (tooLarge)
        {
            // Refused by closing rather than by answering: the request was never fully read, and
            // a 200 for a body that was thrown away is the failure this exists to prevent.
            Log.Info($"debug api: a body over {LargestBody} bytes was refused rather than truncated");
            return null;
        }

        // The body gets its own buffer, sized to what was declared, with whatever of it already
        // arrived alongside the headers copied in first.
        var bodyStart = headerEnd + 4;
        var alreadyHere = Math.Min(contentLength, held - bodyStart);
        var bodyBytes = new byte[contentLength];
        if (alreadyHere > 0)
        {
            Array.Copy(buffer, bodyStart, bodyBytes, 0, alreadyHere);
        }

        var bodyHeld = Math.Max(0, alreadyHere);
        while (bodyHeld < contentLength)
        {
            var read = stream.Read(bodyBytes, bodyHeld, contentLength - bodyHeld);
            if (read <= 0)
            {
                break;
            }

            bodyHeld += read;
        }

        var body = bodyHeld > 0 ? Encoding.UTF8.GetString(bodyBytes, 0, bodyHeld) : string.Empty;

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

    /// <summary>
    /// Cuts a rectangle out of a captured bitmap, in the SCREEN coordinates a page reports.
    ///
    /// A whole editor frame is a big picture in which a 54-pixel drop zone is invisible, and
    /// a surface built by measuring numbers rather than looking at it is a surface built with
    /// one eye shut. The page can say where a widget is (`inspect`); this cuts that out of
    /// the frame so it can actually be seen.
    ///
    /// The rows of a bottom-up DIB run last to first, which is why the copy walks backwards.
    /// </summary>
    public static byte[]? CropBmp(byte[] source, int frameLeft, int frameTop, int x, int y, int width, int height)
    {
        ArgumentNullException.ThrowIfNull(source);

        if (source.Length < 54 || width <= 0 || height <= 0)
        {
            return null;
        }

        var fullWidth = BitConverter.ToInt32(source, 18);
        var fullHeight = BitConverter.ToInt32(source, 22);

        // Screen coordinates into the frame's own.
        var left = x - frameLeft;
        var top = y - frameTop;

        // Clamped rather than refused: a widget half off the frame is still worth seeing.
        left = Math.Clamp(left, 0, Math.Max(0, fullWidth - 1));
        top = Math.Clamp(top, 0, Math.Max(0, fullHeight - 1));
        width = Math.Clamp(width, 1, fullWidth - left);
        height = Math.Clamp(height, 1, fullHeight - top);

        var pixelBytes = width * height * 4;
        var file = new byte[54 + pixelBytes];

        file[0] = (byte)'B';
        file[1] = (byte)'M';
        BitConverter.GetBytes(file.Length).CopyTo(file, 2);
        BitConverter.GetBytes(54).CopyTo(file, 10);
        BitConverter.GetBytes(40).CopyTo(file, 14);
        BitConverter.GetBytes(width).CopyTo(file, 18);
        BitConverter.GetBytes(height).CopyTo(file, 22);
        BitConverter.GetBytes((ushort)1).CopyTo(file, 26);
        BitConverter.GetBytes((ushort)32).CopyTo(file, 28);
        BitConverter.GetBytes(pixelBytes).CopyTo(file, 34);

        for (var row = 0; row < height; row++)
        {
            // Row 0 of the crop is its BOTTOM, which sits at (top + height - 1) from the
            // frame's top, and that row lives at (fullHeight - 1 - thatRow) in the file.
            var sourceRow = fullHeight - 1 - (top + height - 1 - row);
            var from = 54 + ((sourceRow * fullWidth) + left) * 4;
            var to = 54 + row * width * 4;

            if (from < 54 || from + width * 4 > source.Length)
            {
                continue;
            }

            Array.Copy(source, from, file, to, width * 4);
        }

        return file;
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

/// <summary>
/// One native dialog standing in the process.
///
/// Text, not just Caption: every VBA compile error wears the caption "Microsoft Visual Basic for
/// Applications", so the caption alone cannot tell one from another and a harness reading only
/// captions learns that something is wrong but never what.
/// </summary>
public sealed record DebugDialogRow(
    [property: JsonPropertyName("window")] string Window,
    [property: JsonPropertyName("caption")] string Caption,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("buttons")] string[] Buttons,
    [property: JsonPropertyName("enabled")] bool Enabled);

public sealed record DebugDialogsReply(
    [property: JsonPropertyName("dialogs")] DebugDialogRow[] Dialogs,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs);

/// <summary>
/// The dialog guard: whether it is on, and what it has taken off the screen.
///
/// Cleared is the important half. A guard that silently swallows a compile error turns a hang
/// into a mystery, which is a worse trade than the hang.
/// </summary>
public sealed record DebugGuardReply(
    [property: JsonPropertyName("on")] bool On,
    [property: JsonPropertyName("cleared")] string[] Cleared,
    [property: JsonPropertyName("standing")] int Standing);

/// <summary>
/// What a compile said, as data rather than as a modal nobody can read from a script. Errors
/// is empty when the project compiles, which is the only summary worth trusting.
/// </summary>
public sealed record DebugCompileReply(
    [property: JsonPropertyName("compiled")] bool Compiled,
    [property: JsonPropertyName("errors")] string[] Errors,
    [property: JsonPropertyName("project")] string Project);

/// <summary>
/// What became of a component the door was asked to add, rename or remove.
///
/// Name is read BACK from the component rather than echoed: the editor normalises names it
/// dislikes, and a fixture built on the name that was asked for rather than the one that exists
/// is a fixture that addresses something else.
/// </summary>
public sealed record DebugComponentReply(
    [property: JsonPropertyName("ok")] bool Ok,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("action")] string Action);

/// <summary>One component of a VBA project, as the object model has it.</summary>
public sealed record DebugComponentRow(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("type")] int Type,
    [property: JsonPropertyName("lines")] int Lines,
    [property: JsonPropertyName("hasPane")] bool HasPane);

/// <summary>
/// What a workbook's VBA project actually contains, read from the object model.
///
/// The one question a harness could not ask without "Trust access to the VBA project object
/// model": not what the surface is SHOWING, but what is THERE. Building a fixture and then
/// checking it is the same question twice, and both were reaching outside for it.
/// </summary>
public sealed record DebugProjectReply(
    [property: JsonPropertyName("project")] string Project,
    [property: JsonPropertyName("projectId")] string? ProjectId,
    [property: JsonPropertyName("mode")] int Mode,
    [property: JsonPropertyName("components")] DebugComponentRow[] Components);

/// <summary>
/// What an Immediate window line came to.
///
/// The route used to answer that an evaluation had been ASKED FOR: it posted the line to the host
/// thread and returned without waiting, so what the expression evaluated to, and whether it
/// failed, went only to the page. That is why the Immediate window had a route and no suite.
///
/// With no `text` the route reads instead, and `text` carries the whole window as it stands.
/// </summary>
public sealed record DebugImmediateReply(
    /// <summary>False when the evaluation did not finish inside the wait. The line still went in.</summary>
    [property: JsonPropertyName("ran")] bool Ran,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("failed")] bool Failed);

/// <summary>
/// One open workbook, named the way the tree and the tab strip name it.
///
/// The plural of `project`, which answers about one. With two workbooks open there was no way to
/// discover the other's name from the host at all: a probe either knew it in advance or read the
/// page's tree, which is the surface's view rather than the object model's. Two workbooks holding
/// a module of the same name is a designed case here, and three separate defects have lived in
/// it, so a suite that cannot name the workbook it means cannot test any of them.
/// </summary>
public sealed record DebugProjectRow(
    [property: JsonPropertyName("project")] string Project,
    [property: JsonPropertyName("projectId")] string ProjectId,
    /// <summary>Components, with this product's own scratch module left out of the count.</summary>
    [property: JsonPropertyName("components")] int Components,
    /// <summary>Whether the surface is showing a module of this workbook.</summary>
    [property: JsonPropertyName("shown")] bool Shown);

public sealed record DebugProjectsReply(
    [property: JsonPropertyName("projects")] DebugProjectRow[] Projects);

/// <summary>
/// The ENGINE's copy of a module against the surface's.
///
/// Every finding is computed against the engine's copy, and it is maintained incrementally by
/// didChange rather than re-sent whole. When a squiggle is drawn in the wrong place, this is the
/// question that settles which side drifted, and it could not be asked before 2026-08-08.
/// </summary>
public sealed record DebugEngineSourceReply(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("engineHolds")] bool EngineHolds,
    [property: JsonPropertyName("engineLines")] int EngineLines,
    [property: JsonPropertyName("surfaceLines")] int SurfaceLines,
    [property: JsonPropertyName("engineContent")] string? EngineContent,
    [property: JsonPropertyName("surfaceContent")] string? SurfaceContent,
    [property: JsonPropertyName("engineText")] string? EngineText,
    [property: JsonPropertyName("surfaceText")] string? SurfaceText);

/// <summary>One pane the host's own editor holds open.</summary>
public sealed record DebugNativePaneRow(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("project")] string? Project,
    /// <summary>The workbook's text for this module, reduced to a comparable form.</summary>
    [property: JsonPropertyName("hostContent")] string? HostContent,
    /// <summary>
    /// The surface's copy of the same module, reduced the same way. Null when the surface holds
    /// no text for it, which is a different answer from holding the wrong text.
    /// </summary>
    [property: JsonPropertyName("surfaceContent")] string? SurfaceContent);

/// <summary>
/// The HOST's editor, underneath the surface that covers it.
///
/// Run, Step, Compile and ToggleBreakpoint all act on the native active code pane and the caret
/// inside it, not on the page. When the two disagree, a Run executes somewhere the developer is
/// not looking and a breakpoint lands on the wrong line, with nothing on screen to say so. The
/// surface's own idea is carried alongside so the comparison is one reply rather than two.
/// </summary>
public sealed record DebugNativeReply(
    [property: JsonPropertyName("activeModule")] string? ActiveModule,
    [property: JsonPropertyName("activeProject")] string? ActiveProject,
    [property: JsonPropertyName("caretLine")] int CaretLine,
    [property: JsonPropertyName("caretColumn")] int CaretColumn,
    [property: JsonPropertyName("panes")] DebugNativePaneRow[] Panes,
    /// <summary>What the SURFACE believes it is showing, for the same-reply comparison.</summary>
    [property: JsonPropertyName("surfaceModule")] string? SurfaceModule,
    [property: JsonPropertyName("surfaceProject")] string? SurfaceProject,
    /// <summary>Lines in the native pane's code module. Names agreeing is not parity.</summary>
    [property: JsonPropertyName("nativeLines")] int NativeLines,
    /// <summary>Lines the surface holds for what it is showing.</summary>
    [property: JsonPropertyName("surfaceLines")] int SurfaceLines,
    /// <summary>
    /// The native pane's text reduced to a comparable form: length and hash, with line endings
    /// normalised and trailing blanks dropped. Null when there is no pane.
    /// </summary>
    [property: JsonPropertyName("nativeContent")] string? NativeContent,
    /// <summary>The same reduction of what the surface holds, so the two compare exactly.</summary>
    [property: JsonPropertyName("surfaceContent")] string? SurfaceContent,
    /// <summary>The native text itself, only when asked with `text=1`.</summary>
    [property: JsonPropertyName("nativeText")] string? NativeText,
    /// <summary>The surface's text itself, only when asked with `text=1`.</summary>
    [property: JsonPropertyName("surfaceText")] string? SurfaceText);

/// <summary>The breakpoints one module carries, in line order.</summary>
public sealed record DebugBreakpointRow(
    [property: JsonPropertyName("module")] string Module,
    /// <summary>The workbook it belongs to. Two open workbooks can each hold the same module.</summary>
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("lines")] int[] Lines);

/// <summary>
/// Every breakpoint the session is holding, and the mode it is in.
///
/// There was a way to SET one from the moment this door landed and no way to ask what is set,
/// which makes every debugger assertion a matter of remembering what the test did.
/// </summary>
public sealed record DebugBreakpointsReply(
    [property: JsonPropertyName("breakpoints")] DebugBreakpointRow[] Breakpoints,
    [property: JsonPropertyName("mode")] string? Mode);

/// <summary>The developer's settings as they stand, after any change this request asked for.</summary>
public sealed record DebugSettingsReply(
    [property: JsonPropertyName("blockLayout")] string BlockLayout,
    [property: JsonPropertyName("continueCommentOnNewline")] bool ContinueCommentOnNewline,
    [property: JsonPropertyName("mirrorCommentSpacing")] bool MirrorCommentSpacing,
    [property: JsonPropertyName("treeFollowsEditor")] bool TreeFollowsEditor,
    [property: JsonPropertyName("formatIndentSize")] int FormatIndentSize,
    [property: JsonPropertyName("formatCanonicalKeywords")] bool FormatCanonicalKeywords);

/// <summary>Where a marker landed in the log, so a caller can read back from exactly there.</summary>
public sealed record DebugMarkReply(
    [property: JsonPropertyName("marked")] string Marked,
    [property: JsonPropertyName("at")] long At);

/// <summary>One procedure in a module, as the analyzer sees it.</summary>
public sealed record DebugProcedureRow(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("line")] int Line);

/// <summary>A module's shape without reading its text and parsing it again in the caller.</summary>
public sealed record DebugOutlineReply(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("procedures")] DebugProcedureRow[] Procedures);

/// <summary>One document the HOST is holding text for, and whether the page was given it.</summary>
public sealed record DebugDocumentRow(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("lines")] int Lines,
    [property: JsonPropertyName("unwritten")] bool Unwritten,
    [property: JsonPropertyName("active")] bool Active);

/// <summary>
/// The documents the surface holds. Which modules have TEXT and which merely have tabs are
/// different lists, and two defects came from nothing ever showing the difference.
/// </summary>
public sealed record DebugDocumentsReply(
    [property: JsonPropertyName("documents")] DebugDocumentRow[] Documents,
    [property: JsonPropertyName("shownModule")] string? ShownModule);

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

/// <summary>
/// Everything a bug report needs, captured at one moment. Evidence gathered request by
/// request describes several different moments, and the interesting one is usually the one
/// that has already passed.
/// </summary>
public sealed record DebugJournalReply(
    [property: JsonPropertyName("capturedAt")] string CapturedAt,
    [property: JsonPropertyName("pid")] int Pid,
    [property: JsonPropertyName("state")] string State,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs,
    [property: JsonPropertyName("dialogs")] DebugDialogRow[] Dialogs,
    [property: JsonPropertyName("placementMs")] long[] PlacementMs,
    [property: JsonPropertyName("marshalMs")] long[] MarshalMs,
    [property: JsonPropertyName("messages")] DebugMessageRow[] Messages,
    [property: JsonPropertyName("log")] string[] Log);

/// <summary>Whether a stated claim held, and what was there instead when it did not.</summary>
public sealed record DebugAssertReply(
    [property: JsonPropertyName("held")] bool Held,
    [property: JsonPropertyName("claim")] string Claim,
    [property: JsonPropertyName("expected")] string Expected,
    [property: JsonPropertyName("saw")] string Saw);

/// <summary>The requests this door has served, as a script that can be replayed.</summary>
/// <summary>One route's share of the door's own time.</summary>
public sealed record DebugRouteCost(
    [property: JsonPropertyName("route")] string Route,
    [property: JsonPropertyName("count")] long Count,
    [property: JsonPropertyName("totalMs")] long TotalMs,
    [property: JsonPropertyName("maxMs")] long MaxMs);

public sealed record DebugHistoryReply(
    [property: JsonPropertyName("requests")] string[] Requests,
    [property: JsonPropertyName("script")] string Script,
    /// <summary>Each route's count, total and worst time, slowest total first.</summary>
    [property: JsonPropertyName("routeCosts")] DebugRouteCost[] RouteCosts);

/// <summary>Recent raw durations, for percentiles a running maximum cannot give.</summary>
public sealed record DebugPerfReply(
    [property: JsonPropertyName("placementMs")] long[] PlacementMs,
    [property: JsonPropertyName("marshalMs")] long[] MarshalMs,
    [property: JsonPropertyName("heartbeatAgeMs")] long HeartbeatAgeMs,
    /// <summary>Per analyzer method, most time spent first. The largest cost in the product.</summary>
    [property: JsonPropertyName("engine")] EngineMethodCost[] Engine,
    /// <summary>Individual analyzer calls over the slow threshold, newest first: the tail.</summary>
    [property: JsonPropertyName("engineSlowest")] EngineSlowCall[] EngineSlowest,
    /// <summary>How long the engine figures cover, since they can be reset per experiment.</summary>
    [property: JsonPropertyName("engineWindowMs")] long EngineWindowMs);

/// <summary>
/// What a script run in the page answered with.
///
/// Two fields for one answer, because the encoding bit twice. `result` is the browser's own JSON
/// of the value, so a script returning a STRING comes back quoted — and a script returning a
/// JSON string comes back quoted twice. One parse then leaves a string that reads as an object
/// right up until every property of it is undefined, which is a probe reporting false for
/// something that worked (2026-08-07). `value` is that unwrapped as far as it goes, so a caller
/// that wants the answer can read the answer.
/// </summary>
public sealed record DebugEvalReply(
    [property: JsonPropertyName("answered")] bool Answered,
    [property: JsonPropertyName("errorCode")] int ErrorCode,
    [property: JsonPropertyName("result")] string Result,
    [property: JsonPropertyName("value")] System.Text.Json.Nodes.JsonNode? Value);

/// <summary>
/// A condition waited for in the page: whether it came true, and how long it took. Elapsed
/// is the interesting half when it did — a condition met in 4ms and one met in 4 seconds are
/// different facts about the same PASS.
/// </summary>
public sealed record DebugAwaitReply(
    [property: JsonPropertyName("met")] bool Met,
    [property: JsonPropertyName("elapsedMs")] int ElapsedMs,
    [property: JsonPropertyName("detail")] string Detail);

/// <summary>
/// One element the page holds: where it is, what it is, and — when asked — which CSS rules
/// claim a property and what the winner computed to. The rule list is the point: a page
/// sharing a stylesheet with a large bundle loses arguments it never knew it was having.
/// </summary>
public sealed record DebugElementRow(
    [property: JsonPropertyName("tag")] string Tag,
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("classes")] string Classes,
    [property: JsonPropertyName("hidden")] bool Hidden,
    [property: JsonPropertyName("x")] int X,
    [property: JsonPropertyName("y")] int Y,
    [property: JsonPropertyName("w")] int W,
    [property: JsonPropertyName("h")] int H,
    [property: JsonPropertyName("styles")] Dictionary<string, string> Styles,
    [property: JsonPropertyName("rules")] string[] Rules);

/// <summary>Elements matching a selector, and how many there were before any cap.</summary>
public sealed record DebugInspectReply(
    [property: JsonPropertyName("selector")] string Selector,
    [property: JsonPropertyName("matched")] int Matched,
    [property: JsonPropertyName("elements")] DebugElementRow[] Elements);

/// <summary>One timed run of a named scenario, in milliseconds.</summary>
public sealed record DebugBenchReply(
    [property: JsonPropertyName("what")] string What,
    [property: JsonPropertyName("runs")] int Runs,
    [property: JsonPropertyName("minMs")] double MinMs,
    [property: JsonPropertyName("medianMs")] double MedianMs,
    [property: JsonPropertyName("p95Ms")] double P95Ms,
    [property: JsonPropertyName("maxMs")] double MaxMs,
    [property: JsonPropertyName("samplesMs")] double[] SamplesMs,
    [property: JsonPropertyName("detail")] string Detail);

/// <summary>A page reload, and what came back: how long to ready, and which bundle it is.</summary>
public sealed record DebugReloadReply(
    [property: JsonPropertyName("ready")] bool Ready,
    [property: JsonPropertyName("elapsedMs")] int ElapsedMs,
    [property: JsonPropertyName("pageBuildStamp")] string PageBuildStamp,
    [property: JsonPropertyName("bundleBuiltUtc")] string BundleBuiltUtc,
    [property: JsonPropertyName("stale")] bool Stale);

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

/// <summary>
/// What a forced finalizer drain saw. `survived: true` in the reply is the whole answer: if a
/// leaked wrapper had been waiting, this call would not have returned at all.
/// </summary>
public sealed record DebugDrainReply(
    [property: JsonPropertyName("wrappersLiveBefore")] long WrappersLiveBefore,
    [property: JsonPropertyName("wrappersLiveAfter")] long WrappersLiveAfter,
    [property: JsonPropertyName("survived")] bool Survived);

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
    [property: JsonPropertyName("placementFastTotalMs")] long PlacementFastTotalMs,
    [property: JsonPropertyName("placementFastMaxMs")] long PlacementFastMaxMs,
    [property: JsonPropertyName("windowEvents")] long WindowEvents,
    [property: JsonPropertyName("refreshPasses")] long RefreshPasses,
    [property: JsonPropertyName("refreshTotalMs")] long RefreshTotalMs,
    [property: JsonPropertyName("refreshMaxMs")] long RefreshMaxMs,
    [property: JsonPropertyName("overlayMs")] long OverlayMs,
    [property: JsonPropertyName("browserMs")] long BrowserMs,
    [property: JsonPropertyName("browserCalls")] long BrowserCalls,
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
    [property: JsonPropertyName("dialogsStanding")] int DialogsStanding,
    /// <summary>
    /// Wrappers built over the editor's objects, given back, and still held.
    ///
    /// COUNTED BECAUSE THE ALTERNATIVE IS A CRASH REPORT. A wrapper holds its own reference on an
    /// apartment-threaded object; one that is never disposed is given back by the FINALIZER
    /// THREAD instead, where releasing it is an access violation the runtime cannot throw and so
    /// ends the host process. The damage also surfaces late and elsewhere: on 2026-08-07 one leak
    /// was reported once against this library, once against VBE7.DLL, and twice as heap
    /// corruption inside ntdll. `live` should return to its resting level after an operation; one
    /// that only climbs is that crash, seen early enough to fix.
    /// </summary>
    [property: JsonPropertyName("comWrappersTaken")] long ComWrappersTaken,
    [property: JsonPropertyName("comWrappersGivenBack")] long ComWrappersGivenBack,
    [property: JsonPropertyName("comWrappersLive")] long ComWrappersLive);

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
[JsonSerializable(typeof(DebugDrainReply))]
[JsonSerializable(typeof(DebugModuleReply))]
[JsonSerializable(typeof(DebugStatsReply))]
[JsonSerializable(typeof(DebugErrorReply))]
[JsonSerializable(typeof(DebugBlockedReply))]
[JsonSerializable(typeof(DebugDialogRow))]
[JsonSerializable(typeof(DebugDialogsReply))]
[JsonSerializable(typeof(DebugGuardReply))]
[JsonSerializable(typeof(DebugComponentReply))]
[JsonSerializable(typeof(DebugImmediateReply))]
[JsonSerializable(typeof(DebugProjectsReply))]
[JsonSerializable(typeof(DebugProjectRow))]
[JsonSerializable(typeof(DebugProjectReply))]
[JsonSerializable(typeof(DebugSettingsReply))]
[JsonSerializable(typeof(DebugEngineSourceReply))]
[JsonSerializable(typeof(DebugNativeReply))]
[JsonSerializable(typeof(DebugNativePaneRow))]
[JsonSerializable(typeof(DebugNativePaneRow[]))]
[JsonSerializable(typeof(DebugBreakpointsReply))]
[JsonSerializable(typeof(DebugMarkReply))]
[JsonSerializable(typeof(DebugOutlineReply))]
[JsonSerializable(typeof(DebugCompileReply))]
[JsonSerializable(typeof(DebugDocumentsReply))]
[JsonSerializable(typeof(DebugEvalReply))]
[JsonSerializable(typeof(DebugAwaitReply))]
[JsonSerializable(typeof(DebugReloadReply))]
[JsonSerializable(typeof(DebugElementRow))]
[JsonSerializable(typeof(DebugInspectReply))]
[JsonSerializable(typeof(DebugBenchReply))]
[JsonSerializable(typeof(DebugDoctorReply))]
[JsonSerializable(typeof(DebugPerfReply))]
[JsonSerializable(typeof(EngineMethodCost))]
[JsonSerializable(typeof(EngineMethodCost[]))]
[JsonSerializable(typeof(EngineSlowCall))]
[JsonSerializable(typeof(EngineSlowCall[]))]
[JsonSerializable(typeof(DebugJournalReply))]
[JsonSerializable(typeof(DebugAssertReply))]
[JsonSerializable(typeof(DebugHistoryReply))]
[JsonSerializable(typeof(DebugRouteCost))]
[JsonSerializable(typeof(DebugRouteCost[]))]
internal sealed partial class DebugJsonContext : JsonSerializerContext;
#endif
