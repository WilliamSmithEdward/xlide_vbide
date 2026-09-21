namespace Xlide.Vbe.Core.Editor;

/// <summary>Separates native command echo from output without trimming printed whitespace.</summary>
public static class ImmediateTranscript
{
    /// <summary>New output, including when the native window discards its oldest lines.</summary>
    public static string? Added(string previous, string current)
    {
        ArgumentNullException.ThrowIfNull(previous);
        ArgumentNullException.ThrowIfNull(current);
        if (current.StartsWith(previous, StringComparison.Ordinal))
        {
            return current[previous.Length..];
        }

        // The buffer rolls by whole lines. The first matching suffix is the longest overlap;
        // searching only line boundaries avoids mistaking part of a value for retained history.
        var at = previous.IndexOf('\n');
        while (at >= 0 && at + 1 < previous.Length)
        {
            var retained = previous.AsSpan(at + 1);
            if (current.AsSpan().StartsWith(retained, StringComparison.Ordinal))
            {
                return current[retained.Length..];
            }
            at = previous.IndexOf('\n', at + 1);
        }

        return null; // Cleared or rewritten: adopt the baseline without replaying old output.
    }

    public static string? Output(string? transcript, string marker)
    {
        ArgumentException.ThrowIfNullOrEmpty(marker);
        if (transcript is null)
        {
            return null;
        }

        var at = transcript.IndexOf(marker + "\r\n", StringComparison.Ordinal);
        if (at < 0)
        {
            return null;
        }

        var output = transcript[(at + marker.Length + 2)..];
        // The accessibility provider appends the blank caret line and a NUL, including when
        // the printed value is empty. Remove that furniture, then one output line terminator.
        if (!output.EndsWith("\r\n\0", StringComparison.Ordinal))
        {
            return null;
        }

        output = output[..^3];
        return output.EndsWith("\r\n", StringComparison.Ordinal) ? output[..^2] : output;
    }
}
