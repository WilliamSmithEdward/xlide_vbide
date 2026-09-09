using System.Globalization;
using System.Text;

namespace Xlide.Vbe.Core.Scm;

/// <summary>Reads `git cat-file --batch` output.</summary>
public static class GitCatFile
{
    /// <summary>
    /// Parses the answers for the requested object names, in order. Each answer is a header line
    /// and then the object: sha, type and size, then exactly size bytes and a newline; or the
    /// name and `missing` with nothing after it. The header names the OBJECT, not the name that
    /// was asked for, so answers pair with requests by position and nothing else. Content is
    /// decoded as UTF-8 with a byte-order mark tolerated; missing (and ambiguous) names are absent
    /// from the result. Bytes, not text, because a size counts bytes and a decoded string cannot
    /// be walked by it.
    /// </summary>
    public static IReadOnlyDictionary<string, string> ParseBatch(
        byte[] output, IReadOnlyList<string> requested)
    {
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(requested);

        var answers = new Dictionary<string, string>(StringComparer.Ordinal);
        var at = 0;
        var index = 0;

        while (at < output.Length && index < requested.Count)
        {
            var lineEnd = Array.IndexOf(output, (byte)'\n', at);
            if (lineEnd < 0)
            {
                break;
            }

            var header = Encoding.UTF8.GetString(output, at, lineEnd - at).TrimEnd('\r');
            at = lineEnd + 1;
            var name = requested[index++];

            var parts = header.Split(' ');
            if (parts.Length < 3 || !long.TryParse(
                parts[^1], NumberStyles.None, CultureInfo.InvariantCulture, out var declared))
            {
                continue;
            }

            // A truncated stream - git killed at a deadline - still yields what arrived.
            var size = (int)Math.Min(declared, output.Length - at);
            var start = at;
            var length = size;
            var marked = length >= 3
                && output[start] == 0xEF && output[start + 1] == 0xBB && output[start + 2] == 0xBF;
            if (marked)
            {
                start += 3;
                length -= 3;
            }

            answers[name] = Encoding.UTF8.GetString(output, start, length);
            at += size + 1;
        }

        return answers;
    }
}
