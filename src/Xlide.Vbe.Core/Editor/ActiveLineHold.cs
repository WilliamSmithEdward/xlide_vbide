namespace Xlide.Vbe.Core.Editor;

/// <summary>
/// The VBE's contract with a line being typed: a line is validated when the caret leaves it,
/// never while the hand is still in it. This is the publish-side half of that contract - while
/// a module's line is held, findings that touch it are kept out of what the surface shows, and
/// the moment the caret settles anywhere else the hold ends and the caller republishes from
/// findings it already has. No re-analysis is involved in either direction.
///
/// A hold begins only when typing actually edits a line, so clicking into a line that carries
/// findings hides nothing: an old verdict about untouched text stays on screen. What is held
/// back is a NEW verdict about text that is mid-keystroke, which is the analyzer flagging
/// `MsgBox ` for its argument count while the arguments are still on their way.
/// </summary>
public sealed class ActiveLineHold
{
    /// <summary>The module whose line is held, or null while nothing is.</summary>
    public string? Module { get; private set; }

    /// <summary>The held 1-based line. Meaningless while Module is null.</summary>
    public int Line { get; private set; }

    public bool IsHolding => Module is not null;

    /// <summary>
    /// Typing edited this line: verdicts about it now wait for the caret to leave. True when
    /// this changed the hold - the caller republishes so verdicts already on screen about the
    /// line's previous text retire immediately.
    /// </summary>
    public bool Begin(string module, int line)
    {
        if (IsSame(module, line))
        {
            return false;
        }

        Module = module;
        Line = line;
        return true;
    }

    /// <summary>
    /// The caret settled here. Releases the hold when "here" is anywhere but the held line;
    /// true when a hold was released, which is the caller's cue to republish.
    /// </summary>
    public bool Release(string? module, int line)
    {
        if (Module is null || (module is not null && IsSame(module, line)))
        {
            return false;
        }

        Module = null;
        return true;
    }

    /// <summary>Unconditional release - the module switched, or the surface reset.</summary>
    public bool Release()
    {
        if (Module is null)
        {
            return false;
        }

        Module = null;
        return true;
    }

    /// <summary>True when a finding spanning these lines is held out of the current publish.</summary>
    public bool Hides(string module, int startLine, int endLine) =>
        Module is not null
        && startLine <= Line
        && Line <= endLine
        && string.Equals(Module, module, StringComparison.OrdinalIgnoreCase);

    private bool IsSame(string module, int line) =>
        line == Line && string.Equals(Module, module, StringComparison.OrdinalIgnoreCase);
}
