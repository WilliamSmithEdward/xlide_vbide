using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// The two lines the attribute features have to recognise in a module's code: a procedure's
/// header and a module-level variable's declaration. One spelling of each, shared, so that what
/// binds an annotation (AttributeAnnotations), what places an attribute line (AttributeRewriter)
/// and what reports an attribute nothing annotates (AttributeDrift) cannot drift from each other.
/// </summary>
internal static class DeclarationSyntax
{
    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

    /// <summary>A Sub, Function or Property leg, with its name in <c>name</c>.</summary>
    public static readonly Regex ProcedureHeader = new(
        @"^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(?<name>\p{L}[\p{L}\p{N}_]*)",
        Options);

    /// <summary>
    /// A module-level variable, with its name in <c>name</c>: Dim, Private, Public or Global,
    /// WithEvents or not, an array or not. Const, Type, Enum, Declare and Event open with the same
    /// keywords and are not variables - none of them carries a VB_VarDescription - so the word
    /// after the visibility is checked before it is taken for a name; without that check an
    /// annotation above <c>Private Const Limit = 9</c> bound to a variable called Const.
    /// </summary>
    public static readonly Regex VariableDeclaration = new(
        @"^\s*(?:dim|private|public|global)\s+(?:withevents\s+)?(?!(?:const|type|enum|declare|event|withevents|sub|function|property|static)\b)(?<name>\p{L}[\p{L}\p{N}_]*)\b",
        Options);

    /// <summary>
    /// The index of the last line of a statement that starts at <paramref name="at"/>: the line
    /// itself, or the last of the lines a trailing underscore continues it onto. An attribute
    /// line belongs after the whole statement, not in the middle of it.
    /// </summary>
    public static int LastLineOf(System.Collections.Generic.IReadOnlyList<string> lines, int at)
    {
        var last = at;
        while (last < lines.Count - 1 && lines[last].TrimEnd().EndsWith('_'))
        {
            last++;
        }
        return last;
    }
}
