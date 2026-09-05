using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Vba;

/// <summary>The hidden attributes of one procedure, as an exported module carries them.</summary>
/// <param name="Description"><c>VB_Description</c>, or null.</param>
/// <param name="UserMemId"><c>VB_UserMemId</c>: 0 for the default member, -4 for the enumerator, or null.</param>
/// <param name="Hotkey">The letter of <c>VB_ProcData.VB_Invoke_Func</c>, or null.</param>
public sealed record MemberAttributes(string? Description, int? UserMemId, string? Hotkey)
{
    public static readonly MemberAttributes None = new(null, null, null);
    public bool IsEmpty => Description is null && UserMemId is null && Hotkey is null;
}

/// <summary>
/// Everything an exported module says about itself outside its code: the attribute lines the
/// code pane never shows, module-level and per member.
/// </summary>
/// <param name="Name"><c>VB_Name</c>.</param>
/// <param name="Description"><c>VB_Description</c> of the module, or null.</param>
/// <param name="PredeclaredId"><c>VB_PredeclaredId</c>, or null when the header does not say (a standard module).</param>
/// <param name="Exposed"><c>VB_Exposed</c>, or null when the header does not say.</param>
/// <param name="Members">Per procedure, by name (a property's legs share the name and the attributes attach to whichever leg carries them).</param>
/// <param name="VariableDescriptions">Per module-level variable, its <c>VB_VarDescription</c>.</param>
public sealed record AttributeSet(
    string? Name,
    string? Description,
    bool? PredeclaredId,
    bool? Exposed,
    IReadOnlyDictionary<string, MemberAttributes> Members,
    IReadOnlyDictionary<string, string> VariableDescriptions)
{
    public MemberAttributes Member(string name) =>
        Members.TryGetValue(name, out var found) ? found : MemberAttributes.None;
}

/// <summary>
/// Reads and writes the attribute lines of an exported module.
///
/// An exported <c>.bas</c> or <c>.cls</c> is the code as the editor shows it plus the lines it
/// hides: a header of <c>Attribute VB_X = value</c> lines (a class also opens with a VERSION and
/// BEGIN...END block), and inside the code an <c>Attribute Proc.VB_X = value</c> line directly
/// under each procedure header that carries one, and an <c>Attribute var.VB_VarDescription</c>
/// line under a described variable. This is the only shape the editor accepts attributes in, so
/// it is the shape this product reads them from (out of the saved package) and writes them in
/// (through Export, a rewrite, and Import).
/// </summary>
public static class ModuleAttributes
{
    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

    private static readonly Regex AttributeLine = new(
        @"^\s*Attribute\s+(?:(?<owner>\p{L}[\p{L}\p{N}_]*)\.)?(?<name>VB_[A-Za-z_]+(?:\.VB_[A-Za-z_]+)?)\s*=\s*(?<value>.*?)\s*$",
        Options);

    /// <summary>The attributes an exported module's text carries.</summary>
    public static AttributeSet Read(string exported)
    {
        ArgumentNullException.ThrowIfNull(exported);

        string? name = null;
        string? description = null;
        bool? predeclared = null;
        bool? exposed = null;
        var members = new Dictionary<string, MemberAttributes>(StringComparer.OrdinalIgnoreCase);
        var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var raw in exported.Split('\n'))
        {
            var match = AttributeLine.Match(raw.TrimEnd('\r'));
            if (!match.Success)
            {
                continue;
            }

            var attribute = match.Groups["name"].Value;
            var value = match.Groups["value"].Value;
            if (!match.Groups["owner"].Success)
            {
                switch (attribute.ToUpperInvariant())
                {
                    case "VB_NAME": name = StringValue(value); break;
                    case "VB_DESCRIPTION": description = StringValue(value); break;
                    case "VB_PREDECLAREDID": predeclared = BooleanValue(value); break;
                    case "VB_EXPOSED": exposed = BooleanValue(value); break;
                }
                continue;
            }

            var owner = match.Groups["owner"].Value;
            switch (attribute.ToUpperInvariant())
            {
                case "VB_DESCRIPTION":
                    members[owner] = Member(owner) with { Description = StringValue(value) };
                    break;
                case "VB_USERMEMID":
                    members[owner] = Member(owner) with { UserMemId = int.TryParse(value, out var id) ? id : null };
                    break;
                case "VB_PROCDATA.VB_INVOKE_FUNC":
                    members[owner] = Member(owner) with { Hotkey = HotkeyOf(StringValue(value)) };
                    break;
                case "VB_VARDESCRIPTION":
                    if (StringValue(value) is { } text)
                    {
                        variables[owner] = text;
                    }
                    break;
            }
        }

        return new AttributeSet(name, description, predeclared, exposed, members, variables);

        MemberAttributes Member(string owner) =>
            members.TryGetValue(owner, out var found) ? found : MemberAttributes.None;
    }

    /// <summary>A VBA string literal's text: the quotes off and doubled quotes undoubled. Null when it is not one.</summary>
    public static string? StringValue(string literal)
    {
        var trimmed = literal.Trim();
        if (trimmed.Length < 2 || trimmed[0] != '"' || trimmed[^1] != '"')
        {
            return null;
        }
        return trimmed[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
    }

    /// <summary>The literal for a text, quotes doubled.</summary>
    public static string Literal(string text) => $"\"{text.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";

    private static bool? BooleanValue(string value) =>
        value.Trim().ToUpperInvariant() switch
        {
            "TRUE" or "-1" => true,
            "FALSE" or "0" => false,
            _ => null,
        };

    /// <summary>
    /// The letter of an Excel macro hotkey out of the value the editor stores for it. The stored
    /// form is the letter, a literal backslash-n, and 14 - <c>"D\n14"</c> - so the letter is
    /// everything before the backslash.
    /// </summary>
    public static string? HotkeyOf(string? invokeFunc)
    {
        if (string.IsNullOrEmpty(invokeFunc))
        {
            return null;
        }
        var cut = invokeFunc.IndexOf('\\');
        var letter = cut < 0 ? invokeFunc : invokeFunc[..cut];
        return letter.Length == 1 && char.IsLetter(letter[0]) ? letter : null;
    }

    /// <summary>The value the editor stores for a hotkey letter.</summary>
    public static string InvokeFuncFor(string letter) => $"{letter}\\n14";
}
