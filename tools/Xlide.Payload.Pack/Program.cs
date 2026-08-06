// The installer carries everything the product needs inside itself, and most of that is a language
// runtime: xlide-engine.exe is node.exe with the analyzer's JavaScript embedded, ninety megabytes
// against an add-in of six. Embedded resources are stored raw, so the installer was as large as its
// contents. A runtime binary compresses by about seventy per cent, which is the difference between
// a hundred-megabyte download and a thirty-megabyte one.
//
// This runs over the staged payload immediately before the installer is built, replacing each file
// with a Brotli-compressed copy named alongside it. Xlide.Setup reverses it on install, and it
// still accepts uncompressed entries, so a payload that never passed through here works too.
//
//   xlide-payload-pack <payload-directory>
//
// The port in decision 3a is what makes the runtime go away entirely. Until then the runtime ships,
// and there is no reason to ship it uncompressed.

using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;

if (args.Length is not (1 or 2))
{
    Console.Error.WriteLine("usage: xlide-payload-pack <payload-directory> [cache-directory]");
    return 2;
}

var root = Path.GetFullPath(args[0]);
if (!Directory.Exists(root))
{
    Console.Error.WriteLine($"No payload directory at {root}.");
    return 2;
}

const string Suffix = ".br";
var files = Directory.GetFiles(root, "*", SearchOption.AllDirectories)
    .Where(f => !f.EndsWith(Suffix, StringComparison.OrdinalIgnoreCase))
    .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
    .ToArray();

if (files.Length == 0)
{
    Console.WriteLine("    (nothing to compress)");
    return 0;
}

// SmallestSize is brotli's slowest setting, and on a ninety-megabyte runtime it costs two minutes.
// It is still the right setting -- the payload is decompressed on every install and compressed only
// when it changes -- but "when it changes" has to be true, so results are cached by content. The
// engine is the expensive file and it changes rarely; every build after the first reuses it.
var cache = args.Length == 2 ? Path.GetFullPath(args[1]) : null;
if (cache is not null) Directory.CreateDirectory(cache);

long raw = 0;
long packed = 0;
var reused = 0;
var clock = Stopwatch.StartNew();
var report = new System.Collections.Concurrent.ConcurrentBag<(string Line, long Before, long After, bool Reused)>();

Parallel.ForEach(files, file =>
{
    var target = file + Suffix;
    var before = new FileInfo(file).Length;

    string? cached = null;
    if (cache is not null)
    {
        using var source = File.OpenRead(file);
        cached = Path.Combine(cache, Convert.ToHexString(SHA256.HashData(source)) + Suffix);
    }

    var fromCache = cached is not null && File.Exists(cached);
    if (fromCache)
    {
        File.Copy(cached!, target, overwrite: true);
    }
    else
    {
        using (var input = File.OpenRead(file))
        using (var output = File.Create(target))
        using (var brotli = new BrotliStream(output, CompressionLevel.SmallestSize))
        {
            input.CopyTo(brotli);
        }

        if (cached is not null) File.Copy(target, cached, overwrite: true);
    }

    var after = new FileInfo(target).Length;
    File.Delete(file);

    report.Add((
        $"    {Path.GetRelativePath(root, file),-40} {before / 1048576.0,7:N1} MB -> {after / 1048576.0,6:N1} MB{(fromCache ? "  (cached)" : string.Empty)}",
        before, after, fromCache));
});

clock.Stop();

foreach (var entry in report.OrderBy(e => e.Line, StringComparer.OrdinalIgnoreCase))
{
    Console.WriteLine(entry.Line);
    raw += entry.Before;
    packed += entry.After;
    if (entry.Reused) reused++;
}
Console.WriteLine(
    $"    payload {raw / 1048576.0:N1} MB -> {packed / 1048576.0:N1} MB " +
    $"({(1 - (double)packed / raw) * 100:N0}% smaller, {clock.Elapsed.TotalSeconds:N0}s, " +
    $"{reused}/{files.Length} reused)");

return 0;
