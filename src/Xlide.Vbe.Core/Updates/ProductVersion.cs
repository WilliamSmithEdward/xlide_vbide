using System.Globalization;

namespace Xlide.Vbe.Core.Updates;

/// <summary>
/// The version this build carries, and the one way a release tag is read as one.
///
/// Every assembly in the product takes its version from the single <c>&lt;Version&gt;</c> in
/// Directory.Build.props, so the version of this library is the version of the product. Reading
/// it from the assembly rather than from a constant means a release bump is one edit, and a
/// unit test holds this to the props file so the two cannot drift.
/// </summary>
public static class ProductVersion
{
    private static readonly Version Built =
        typeof(ProductVersion).Assembly.GetName().Version ?? new Version(0, 0, 0);

    /// <summary>The three-part version of this build, as a release tag names it: "0.20.21".</summary>
    public static string Current { get; } = Format(Built);

    /// <summary>Formats a version as the three parts a release tag carries.</summary>
    public static string Format(Version version)
    {
        ArgumentNullException.ThrowIfNull(version);
        return string.Create(CultureInfo.InvariantCulture,
            $"{version.Major}.{version.Minor}.{Math.Max(0, version.Build)}");
    }

    /// <summary>
    /// Reads "1.2.3" or "v1.2.3" and nothing else. A tag with two parts, four parts, or a
    /// suffix is not one of this product's releases, and refusing it is what keeps a mistyped
    /// tag from being offered as an update.
    /// </summary>
    public static bool TryParse(string? text, out Version version)
    {
        version = new Version(0, 0, 0);
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var span = text.AsSpan().Trim();
        if (span.Length > 1 && (span[0] == 'v' || span[0] == 'V'))
        {
            span = span[1..];
        }

        var parts = new int[3];
        var index = 0;
        foreach (var piece in span.Split('.'))
        {
            if (index == 3)
            {
                return false;
            }

            var digits = span[piece];
            if (digits.Length is 0 or > 9
                || !int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out parts[index]))
            {
                return false;
            }

            index++;
        }

        if (index != 3)
        {
            return false;
        }

        version = new Version(parts[0], parts[1], parts[2]);
        return true;
    }

    /// <summary>
    /// Orders two versions: negative when <paramref name="left"/> is older. Text that is not a
    /// version sorts as the oldest thing there is, so an unreadable tag is never "newer".
    /// </summary>
    public static int Compare(string? left, string? right)
    {
        var leftRead = TryParse(left, out var leftVersion);
        var rightRead = TryParse(right, out var rightVersion);

        if (!leftRead || !rightRead)
        {
            return leftRead.CompareTo(rightRead);
        }

        return leftVersion.CompareTo(rightVersion);
    }
}
