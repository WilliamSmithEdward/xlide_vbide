using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.WebView;

/// <summary>
/// Serves the editor bundle to the browser over a loopback socket.
///
/// The browser's own folder mapping brokers every byte through its host-process pipe, and that
/// pipe moves the multi-megabyte bundle at roughly two megabytes a second: the editor spent two
/// seconds of every start-up fetching a file that sits on local disk. Loopback TCP moves the
/// same bytes in tens of milliseconds, so this small server is the difference between a
/// two-second boot and a sub-half-second one. The folder mapping stays behind it as the
/// fallback, because a slow editor beats no editor.
///
/// The listener binds 127.0.0.1 on an ephemeral port, answers GET and HEAD only, serves only
/// what lies under the one directory it was given, and requires a random path token minted per
/// session — a stranger process scanning local ports learns nothing and fetches nothing, and
/// the served content is the product's own page bundle in any case.
/// </summary>
internal sealed class LoopbackPageServer : IDisposable
{
    private readonly TcpListener _listener;
    private readonly string _root;
    private readonly string _token;
    private volatile bool _stopping;

    private LoopbackPageServer(TcpListener listener, string root, string token, int port)
    {
        _listener = listener;
        _root = root;
        _token = token;
        Port = port;
    }

    public int Port { get; }

    /// <summary>The origin plus the session token; append the document path to navigate.</summary>
    public string BaseUrl => $"http://127.0.0.1:{Port}/{_token}";

    /// <summary>A running server over the directory, or null when the socket could not be had.</summary>
    public static LoopbackPageServer? Start(string contentRoot)
    {
        try
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();

            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var token = Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();
            var server = new LoopbackPageServer(listener, Path.GetFullPath(contentRoot), token, port);

            new Thread(server.AcceptLoop)
            {
                IsBackground = true,
                Name = "xlide-page-server",
            }.Start();

            Log.Info($"page server: serving the editor bundle on 127.0.0.1:{port}");
            return server;
        }
        catch (Exception ex)
        {
            Log.Warn($"page server: could not listen on loopback, {ex.Message}");
            return null;
        }
    }

    private void AcceptLoop()
    {
        while (!_stopping)
        {
            TcpClient client;
            try
            {
                client = _listener.AcceptTcpClient();
            }
            catch (Exception)
            {
                // The listener was stopped, or the socket faulted; either way the loop is over.
                return;
            }

            ThreadPool.QueueUserWorkItem(static state =>
            {
                var (server, accepted) = ((LoopbackPageServer, TcpClient))state!;
                server.Serve(accepted);
            }, (this, client));
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

            var (method, target) = request;
            if (method != "GET" && method != "HEAD")
            {
                WriteStatus(stream, "405 Method Not Allowed");
                return;
            }

            var file = Resolve(target);
            if (file is null)
            {
                WriteStatus(stream, "404 Not Found");
                return;
            }

            var body = File.ReadAllBytes(file);
            var header = "HTTP/1.1 200 OK\r\n"
                + $"Content-Type: {ContentTypeFor(file)}\r\n"
                + $"Content-Length: {body.Length}\r\n"
                // Always the bundle as published right now: a cached stale editor is the
                // debugging problem nothing else in this product has, and the fetch is cheap.
                + "Cache-Control: no-store\r\n"
                + "Connection: close\r\n"
                + "\r\n";

            stream.Write(Encoding.ASCII.GetBytes(header));
            if (method == "GET")
            {
                stream.Write(body);
            }
        }
        catch (Exception)
        {
            // A connection the browser abandoned mid-answer is routine, not news.
        }
    }

    /// <summary>The request line's method and target, with the headers drained; null if malformed.</summary>
    private static (string Method, string Target)? ReadRequest(NetworkStream stream)
    {
        // Requests here are a line and a few headers; 16KB is generous and bounds a bad actor.
        var buffer = new byte[16384];
        var held = 0;

        while (held < buffer.Length)
        {
            var read = stream.Read(buffer, held, buffer.Length - held);
            if (read <= 0)
            {
                return null;
            }

            held += read;
            if (held >= 4 && buffer.AsSpan(0, held).IndexOf("\r\n\r\n"u8) >= 0)
            {
                break;
            }
        }

        var text = Encoding.ASCII.GetString(buffer, 0, held);
        var lineEnd = text.IndexOf("\r\n", StringComparison.Ordinal);
        if (lineEnd < 0)
        {
            return null;
        }

        var parts = text[..lineEnd].Split(' ');
        return parts.Length == 3 ? (parts[0], parts[1]) : null;
    }

    /// <summary>The file a target names, or null: wrong token, a dodged root, or nothing there.</summary>
    private string? Resolve(string target)
    {
        var path = target;
        var query = path.IndexOf('?', StringComparison.Ordinal);
        if (query >= 0)
        {
            path = path[..query];
        }

        var prefix = $"/{_token}/";
        if (!path.StartsWith(prefix, StringComparison.Ordinal))
        {
            return null;
        }

        var relative = Uri.UnescapeDataString(path[prefix.Length..]);
        if (relative.Length == 0
            || relative.Contains("..", StringComparison.Ordinal)
            || relative.Contains('\\', StringComparison.Ordinal))
        {
            return null;
        }

        var full = Path.GetFullPath(Path.Combine(_root, relative));
        if (!full.StartsWith(_root, StringComparison.OrdinalIgnoreCase) || !File.Exists(full))
        {
            return null;
        }

        return full;
    }

    private static string ContentTypeFor(string file) => Path.GetExtension(file).ToLowerInvariant() switch
    {
        ".html" => "text/html; charset=utf-8",
        ".js" => "text/javascript; charset=utf-8",
        ".css" => "text/css; charset=utf-8",
        ".json" => "application/json",
        ".ttf" => "font/ttf",
        ".svg" => "image/svg+xml",
        ".png" => "image/png",
        _ => "application/octet-stream",
    };

    private static void WriteStatus(NetworkStream stream, string status)
    {
        var header = $"HTTP/1.1 {status}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n";
        stream.Write(Encoding.ASCII.GetBytes(header));
    }

    public void Dispose()
    {
        _stopping = true;
        try
        {
            _listener.Stop();
        }
        catch (Exception)
        {
            // Closing a listener that already faulted is still closed.
        }
    }
}
