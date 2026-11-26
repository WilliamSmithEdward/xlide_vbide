using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using Xlide.Vbe.Core.Updates;
using Xlide.Vbe.Shim.Diagnostics;

namespace Xlide.Vbe.Shim.Updates;

/// <summary>Something the updater could not do, in words the dialog can show as they are.</summary>
internal sealed class UpdateException(string message) : Exception(message);

/// <summary>
/// The network and file half of self-update: reads the release feed, downloads the installer
/// while hashing it, and starts it. No COM and no UI, so every call here runs on a pool thread
/// and the session marshals the outcome back to the host thread itself.
///
/// What stands in for a code signature is here. The installer is not signed, so the chain of
/// trust is TLS to the feed's own address, a release whose every address the source vouched for
/// before anything was fetched, and the SHA-256 GitHub computed on upload, checked against every
/// byte as it arrives. A download whose length or hash disagrees is deleted before anything can
/// run it.
/// </summary>
internal sealed class UpdateService : IDisposable
{
    /// <summary>The feed is a few kilobytes; anything near this is not a release.</summary>
    private const int FeedLimit = 2 * 1024 * 1024;

    private const int BufferSize = 1 << 16;

    private readonly Lazy<HttpClient> _http = new(CreateClient);

    private static HttpClient CreateClient()
    {
        var handler = new SocketsHttpHandler
        {
            // GitHub's download address answers with a redirect to its object store. HttpClient
            // follows it, and refuses on its own to follow one that steps down from https.
            AllowAutoRedirect = true,
            MaxAutomaticRedirections = 5,
            ConnectTimeout = TimeSpan.FromSeconds(15),
            PooledConnectionLifetime = TimeSpan.FromMinutes(5),
        };

        // Deadlines are per call, below: a feed read is seconds and a download is minutes.
        var client = new HttpClient(handler) { Timeout = Timeout.InfiniteTimeSpan };
        client.DefaultRequestHeaders.UserAgent.ParseAdd(string.Create(CultureInfo.InvariantCulture,
            $"xlide_vbide/{ProductVersion.Current}"));
        return client;
    }

    /// <summary>Reads the feed's answer as text. Throws <see cref="UpdateException"/> with the reason.</summary>
    public async Task<string> ReadFeedAsync(UpdateSource source, CancellationToken cancellation)
    {
        ArgumentNullException.ThrowIfNull(source);

        using var request = new HttpRequestMessage(HttpMethod.Get, source.FeedUrl);
        request.Headers.Accept.ParseAdd("application/vnd.github+json");
        request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");

        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        deadline.CancelAfter(TimeSpan.FromSeconds(30));

        HttpResponseMessage response;
        try
        {
            response = await _http.Value.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (!cancellation.IsCancellationRequested)
        {
            throw new UpdateException("the release feed did not answer within 30 seconds");
        }
        catch (HttpRequestException ex)
        {
            throw new UpdateException($"the release feed could not be reached: {ex.Message.Trim()}");
        }

        using (response)
        {
            if (response.StatusCode is HttpStatusCode.Forbidden or HttpStatusCode.TooManyRequests)
            {
                // Sixty unauthenticated requests an hour per address. A whole office behind one
                // address can spend them; the honest answer is to try later, not to retry now.
                throw new UpdateException("GitHub is limiting requests from this address for a while; try again later");
            }

            if (!response.IsSuccessStatusCode)
            {
                throw new UpdateException(string.Create(CultureInfo.InvariantCulture,
                    $"the release feed answered {(int)response.StatusCode} {response.ReasonPhrase}"));
            }

            if (response.Content.Headers.ContentLength is > FeedLimit)
            {
                throw new UpdateException("the release feed answered more than a release");
            }

            await using var body = await response.Content.ReadAsStreamAsync(deadline.Token).ConfigureAwait(false);
            using var text = new MemoryStream();
            var buffer = new byte[BufferSize];
            int read;
            while ((read = await body.ReadAsync(buffer, deadline.Token).ConfigureAwait(false)) > 0)
            {
                if (text.Length + read > FeedLimit)
                {
                    throw new UpdateException("the release feed answered more than a release");
                }

                text.Write(buffer, 0, read);
            }

            return System.Text.Encoding.UTF8.GetString(text.GetBuffer(), 0, (int)text.Length);
        }
    }

    /// <summary>
    /// Downloads a release's installer into <paramref name="folder"/>, checking its length and
    /// SHA-256 against what the release published, and answers the verified file's path. A file
    /// already there that passes the same check is answered without a download. Progress is
    /// reported in bytes received, on the calling pool thread.
    /// </summary>
    public async Task<string> DownloadAsync(
        Release release,
        string folder,
        Action<long> progress,
        CancellationToken cancellation)
    {
        ArgumentNullException.ThrowIfNull(release);
        ArgumentException.ThrowIfNullOrWhiteSpace(folder);
        ArgumentNullException.ThrowIfNull(progress);

        Directory.CreateDirectory(folder);
        var final = Path.Combine(folder, ReleaseFeed.DownloadFileName(release.Version));
        var partial = final + ".partial";

        if (File.Exists(final) && await MatchesAsync(final, release.Installer, cancellation).ConfigureAwait(false))
        {
            Log.Info($"update: {Path.GetFileName(final)} is already here and verifies");
            progress(release.Installer.Size);
            RemoveOtherInstallers(folder, final);
            return final;
        }

        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
        deadline.CancelAfter(TimeSpan.FromMinutes(20));

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, release.Installer.DownloadUrl);
            using var response = await _http.Value.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token)
                .ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new UpdateException(string.Create(CultureInfo.InvariantCulture,
                    $"the download answered {(int)response.StatusCode} {response.ReasonPhrase}"));
            }

            if (response.Content.Headers.ContentLength is { } declared && declared != release.Installer.Size)
            {
                throw new UpdateException(string.Create(CultureInfo.InvariantCulture,
                    $"the download is {declared:N0} bytes where the release says {release.Installer.Size:N0}"));
            }

            await using var body = await response.Content.ReadAsStreamAsync(deadline.Token).ConfigureAwait(false);
            string hex;
            await using (var file = new FileStream(partial, FileMode.Create, FileAccess.Write, FileShare.None, BufferSize, FileOptions.Asynchronous))
            using (var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
            {
                var buffer = new byte[BufferSize];
                long received = 0;
                int read;
                while ((read = await body.ReadAsync(buffer, deadline.Token).ConfigureAwait(false)) > 0)
                {
                    received += read;
                    if (received > release.Installer.Size)
                    {
                        throw new UpdateException("the download is longer than the release says");
                    }

                    hash.AppendData(buffer, 0, read);
                    await file.WriteAsync(buffer.AsMemory(0, read), deadline.Token).ConfigureAwait(false);
                    progress(received);
                }

                if (received != release.Installer.Size)
                {
                    throw new UpdateException(string.Create(CultureInfo.InvariantCulture,
                        $"the download ended at {received:N0} of {release.Installer.Size:N0} bytes"));
                }

                hex = Convert.ToHexStringLower(hash.GetHashAndReset());
            }

            if (!string.Equals(hex, release.Installer.Sha256, StringComparison.Ordinal))
            {
                throw new UpdateException(
                    "the download's SHA-256 does not match the digest the release published, so it was deleted");
            }

            File.Move(partial, final, overwrite: true);
        }
        catch (OperationCanceledException) when (!cancellation.IsCancellationRequested)
        {
            TryDelete(partial);
            throw new UpdateException("the download did not finish within 20 minutes");
        }
        catch (HttpRequestException ex)
        {
            TryDelete(partial);
            throw new UpdateException($"the download failed: {ex.Message.Trim()}");
        }
        catch (IOException ex)
        {
            TryDelete(partial);
            throw new UpdateException($"the download could not be written: {ex.Message.Trim()}");
        }
        catch
        {
            TryDelete(partial);
            throw;
        }

        RemoveOtherInstallers(folder, final);
        return final;
    }

    /// <summary>
    /// Starts a downloaded installer in its own window. The installer finds the Office host
    /// open, says so, and installs the moment it closes; nothing here waits for that.
    /// </summary>
    public static Process Launch(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        // A console program started from a windowed process gets a console of its own, which
        // is the window the developer reads the installer in. Not through the shell: the
        // file is ours and verified, and a direct start leaves no association to consult.
        var start = new ProcessStartInfo(path)
        {
            UseShellExecute = false,
            WorkingDirectory = Path.GetDirectoryName(path) ?? string.Empty,
        };

        return Process.Start(start) ?? throw new UpdateException("Windows did not start the installer");
    }

    private static async Task<bool> MatchesAsync(string path, ReleaseAsset asset, CancellationToken cancellation)
    {
        try
        {
            if (new FileInfo(path).Length != asset.Size)
            {
                return false;
            }

            await using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize, FileOptions.Asynchronous);
            var hash = await SHA256.HashDataAsync(file, cancellation).ConfigureAwait(false);
            return string.Equals(Convert.ToHexStringLower(hash), asset.Sha256, StringComparison.Ordinal);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>One installer in the folder at a time: the one just verified.</summary>
    private static void RemoveOtherInstallers(string folder, string keep)
    {
        try
        {
            foreach (var file in Directory.EnumerateFiles(folder, "xlide-setup-*.exe*"))
            {
                if (!string.Equals(file, keep, StringComparison.OrdinalIgnoreCase))
                {
                    TryDelete(file);
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // An old download left behind changes nothing about the one just verified.
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Log.Warn($"update: {path} could not be removed ({ex.GetType().Name})");
        }
    }

    public void Dispose()
    {
        if (_http.IsValueCreated)
        {
            _http.Value.Dispose();
        }
    }
}
