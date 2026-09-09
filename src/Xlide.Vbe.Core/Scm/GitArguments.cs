namespace Xlide.Vbe.Core.Scm;

/// <summary>
/// The arguments every git invocation shares, so the runner and the parsers agree.
/// </summary>
public static class GitArguments
{
    /// <summary>
    /// What every invocation starts with, before the command. No pager, because there is no
    /// terminal to page to and `less` would wait for one; quotepath off, so a module name with a
    /// non-ASCII letter comes back as itself rather than as octal escapes; colour off, so no
    /// escape sequence ever reaches a parser, whatever the developer's own git config says.
    /// </summary>
    public static IReadOnlyList<string> Prefix { get; } =
        ["--no-pager", "-c", "core.quotepath=false", "-c", "color.ui=never"];

    /// <summary>
    /// The log format <see cref="GitLog.Parse"/> reads. A record starts with a record separator
    /// (0x1E), the fields are NUL-separated, and the body ends at a unit separator (0x1F), because
    /// a body can hold anything a developer types - blank lines, tabs, a line that looks like a
    /// name-status row - and only characters no commit message can contain make safe delimiters.
    /// </summary>
    public const string LogFormat = "%x1e%H%x00%h%x00%an%x00%ae%x00%aI%x00%s%x00%b%x00%P%x1f";
}
