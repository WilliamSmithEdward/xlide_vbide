using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Vba;

/// <summary>How an annotation and the module's attributes disagree.</summary>
public enum DriftKind
{
    /// <summary>The annotation is there and the attribute does not match it yet.</summary>
    AnnotationNotApplied,
    /// <summary>The attribute is set and nothing in the code says so.</summary>
    AttributeNotAnnotated,
    /// <summary>The annotation cannot be written on this kind of module.</summary>
    AnnotationNotApplicable,
    /// <summary>The annotation is malformed or misplaced; see the message.</summary>
    AnnotationProblem,
}

/// <summary>One disagreement, placed on the line the developer would look at.</summary>
/// <param name="Kind">What sort of disagreement.</param>
/// <param name="Line">The 1-based code line: the annotation's, or the header of the member whose attribute has none.</param>
/// <param name="Message">The finding's words.</param>
/// <param name="Annotation">The annotation kind involved, when one is.</param>
/// <param name="Target">The procedure or variable, or null for the module.</param>
/// <param name="Occurrence">Which same-named header the target is, for a property's legs.</param>
public sealed record DriftItem(DriftKind Kind, int Line, string Message, AnnotationKind? Annotation, string? Target, int Occurrence)
{
    /// <summary>The finding code the surface files it under.</summary>
    public string Code => Kind switch
    {
        DriftKind.AnnotationNotApplied => "annotation-not-applied",
        DriftKind.AttributeNotAnnotated => "attribute-not-annotated",
        DriftKind.AnnotationNotApplicable => "annotation-not-applicable",
        _ => "annotation-problem",
    };

    /// <summary>The severity the surface draws it at.</summary>
    public string Severity => Kind == DriftKind.AttributeNotAnnotated ? "info" : "warning";
}

/// <summary>
/// Where a module's annotations and its attributes disagree, in either direction.
///
/// An annotation without its attribute is the common case: the developer typed it and has not
/// applied it. An attribute without its annotation is the other: set by an import, another tool,
/// or an annotation since deleted, and invisible in the code pane. Both are reported, because a
/// developer reading the code should be able to trust that it says what the module is.
/// </summary>
public static class AttributeDrift
{
    private static readonly Regex ProcedureHeader = new(
        @"^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(?<name>\p{L}[\p{L}\p{N}_]*)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex VariableDeclaration = new(
        @"^\s*(?:dim|private|public|global)\s+(?:withevents\s+)?(?<name>\p{L}[\p{L}\p{N}_]*)\b(?!\s*\()",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>
    /// The disagreements for one module. <paramref name="actual"/> is null when the saved
    /// workbook cannot vouch for the module's attributes; an annotation is then reported as not
    /// applied, in words that say so, and the reverse direction is not asked at all.
    /// <paramref name="moduleKind"/> is the analyzer's word: standard, class, document, userform.
    /// </summary>
    public static IReadOnlyList<DriftItem> Between(string source, ModuleAnnotations annotations, AttributeSet? actual, string moduleKind, string moduleName)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(annotations);

        var items = new List<DriftItem>();
        var isClass = string.Equals(moduleKind, "class", StringComparison.OrdinalIgnoreCase);
        var writable = isClass || string.Equals(moduleKind, "standard", StringComparison.OrdinalIgnoreCase);

        foreach (var problem in annotations.Problems)
        {
            items.Add(new DriftItem(DriftKind.AnnotationProblem, problem.Line, problem.Message, null, null, 0));
        }

        foreach (var annotation in annotations.Annotations)
        {
            var spelled = AttributeAnnotations.Spelled(annotation.Kind);
            var occurrence = annotation.Target is null || annotation.Kind == AnnotationKind.VariableDescription
                ? 0
                : AttributeRewriter.OccurrenceOf(source, annotation.Target, annotation.TargetLine ?? 0);

            if (!writable)
            {
                items.Add(new DriftItem(DriftKind.AnnotationNotApplicable, annotation.Line,
                    $"{spelled} cannot be applied to {moduleName}: the editor only takes attributes through an import, and a {moduleKind} module cannot be imported.",
                    annotation.Kind, annotation.Target, occurrence));
                continue;
            }

            if (!isClass && annotation.Kind is AnnotationKind.PredeclaredId or AnnotationKind.Exposed)
            {
                items.Add(new DriftItem(DriftKind.AnnotationNotApplicable, annotation.Line,
                    $"{spelled} means something only on a class module; {moduleName} is a standard module.",
                    annotation.Kind, annotation.Target, occurrence));
                continue;
            }

            if (actual is null)
            {
                items.Add(new DriftItem(DriftKind.AnnotationNotApplied, annotation.Line,
                    $"{spelled} is not known to be applied: the saved workbook does not carry {moduleName}'s attributes yet. It is written when you save, or now from the quick fix.",
                    annotation.Kind, annotation.Target, occurrence));
                continue;
            }

            var (wanted, has) = Compare(annotation, actual);
            if (!string.Equals(wanted, has, StringComparison.Ordinal))
            {
                var attribute = AttributeName(annotation.Kind);
                var where = annotation.Target is null ? moduleName : $"{moduleName}.{annotation.Target}";
                items.Add(new DriftItem(DriftKind.AnnotationNotApplied, annotation.Line,
                    $"{spelled} is annotated, but {where}'s {attribute} is {has ?? "not set"}. It is written when you save, or now from the quick fix.",
                    annotation.Kind, annotation.Target, occurrence));
            }
        }

        if (actual is not null && writable)
        {
            ReverseDirection(source, annotations, actual, isClass, moduleName, items);
        }

        return items;
    }

    /// <summary>The attribute an annotation stands for, as the export spells it.</summary>
    public static string AttributeName(AnnotationKind kind) => kind switch
    {
        AnnotationKind.ModuleDescription or AnnotationKind.Description => "VB_Description",
        AnnotationKind.PredeclaredId => "VB_PredeclaredId",
        AnnotationKind.Exposed => "VB_Exposed",
        AnnotationKind.DefaultMember or AnnotationKind.Enumerator => "VB_UserMemId",
        AnnotationKind.ExcelHotkey => "VB_ProcData.VB_Invoke_Func",
        _ => "VB_VarDescription",
    };

    /// <summary>What the annotation wants the attribute to read, and what it reads.</summary>
    private static (string Wanted, string? Has) Compare(Annotation annotation, AttributeSet actual)
    {
        switch (annotation.Kind)
        {
            case AnnotationKind.ModuleDescription:
                return (Quoted(annotation.Argument), actual.Description is null ? null : Quoted(actual.Description));
            case AnnotationKind.PredeclaredId:
                return ("True", actual.PredeclaredId?.ToString());
            case AnnotationKind.Exposed:
                return ("True", actual.Exposed?.ToString());
            case AnnotationKind.VariableDescription:
                return (Quoted(annotation.Argument),
                    actual.VariableDescriptions.TryGetValue(annotation.Target ?? string.Empty, out var text) ? Quoted(text) : null);
        }

        var member = actual.Member(annotation.Target ?? string.Empty);
        return annotation.Kind switch
        {
            AnnotationKind.Description => (Quoted(annotation.Argument), member.Description is null ? null : Quoted(member.Description)),
            AnnotationKind.DefaultMember => ("0", member.UserMemId?.ToString()),
            AnnotationKind.Enumerator => ("-4", member.UserMemId?.ToString()),
            _ => (Quoted(annotation.Argument), member.Hotkey is null ? null : Quoted(member.Hotkey)),
        };
    }

    private static string Quoted(string? text) => $"\"{text}\"";

    private static void ReverseDirection(string source, ModuleAnnotations annotations, AttributeSet actual, bool isClass, string moduleName, List<DriftItem> items)
    {
        bool Annotated(AnnotationKind kind, string? target = null) =>
            annotations.Annotations.Any(one => one.Kind == kind
                && string.Equals(one.Target, target, StringComparison.OrdinalIgnoreCase));

        if (actual.Description is not null && !Annotated(AnnotationKind.ModuleDescription))
        {
            items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, 1,
                $"{moduleName}'s VB_Description is \"{actual.Description}\", and no '@ModuleDescription says so.",
                AnnotationKind.ModuleDescription, null, 0));
        }
        if (isClass && actual.PredeclaredId == true && !Annotated(AnnotationKind.PredeclaredId))
        {
            items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, 1,
                $"{moduleName} has VB_PredeclaredId = True, a default instance, and no '@PredeclaredId says so.",
                AnnotationKind.PredeclaredId, null, 0));
        }
        if (isClass && actual.Exposed == true && !Annotated(AnnotationKind.Exposed))
        {
            items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, 1,
                $"{moduleName} has VB_Exposed = True, and no '@Exposed says so.",
                AnnotationKind.Exposed, null, 0));
        }

        // Members and variables are placed on their own lines in the code.
        var headers = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        var declarations = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var lines = source.Split('\n');
        var inDeclarations = true;
        for (var at = 0; at < lines.Length; at++)
        {
            var line = lines[at].TrimEnd('\r');
            var header = ProcedureHeader.Match(line);
            if (header.Success)
            {
                inDeclarations = false;
                if (!headers.TryGetValue(header.Groups["name"].Value, out var list))
                {
                    headers[header.Groups["name"].Value] = list = [];
                }
                list.Add(at + 1);
                continue;
            }
            if (inDeclarations)
            {
                var variable = VariableDeclaration.Match(line);
                if (variable.Success)
                {
                    declarations.TryAdd(variable.Groups["name"].Value, at + 1);
                }
            }
        }

        foreach (var (name, member) in actual.Members)
        {
            if (!headers.TryGetValue(name, out var at))
            {
                continue;
            }
            // Attributes are read per name and the annotation may sit above any leg, so a member
            // counts as annotated when any leg carries the annotation.
            var line = at[0];
            if (member.Description is not null && !Annotated(AnnotationKind.Description, name))
            {
                items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, line,
                    $"{name}'s VB_Description is \"{member.Description}\", and no '@Description says so.",
                    AnnotationKind.Description, name, 0));
            }
            if (member.UserMemId == 0 && !Annotated(AnnotationKind.DefaultMember, name))
            {
                items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, line,
                    $"{name} is the default member (VB_UserMemId = 0), and no '@DefaultMember says so.",
                    AnnotationKind.DefaultMember, name, 0));
            }
            if (member.UserMemId == -4 && !Annotated(AnnotationKind.Enumerator, name))
            {
                items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, line,
                    $"{name} is the enumerator (VB_UserMemId = -4), and no '@Enumerator says so.",
                    AnnotationKind.Enumerator, name, 0));
            }
            if (member.Hotkey is not null && !Annotated(AnnotationKind.ExcelHotkey, name))
            {
                items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, line,
                    $"{name} has the hotkey Ctrl+{(char.IsUpper(member.Hotkey[0]) ? "Shift+" : string.Empty)}{member.Hotkey.ToUpperInvariant()}, and no '@ExcelHotkey says so.",
                    AnnotationKind.ExcelHotkey, name, 0));
            }
        }

        foreach (var (name, text) in actual.VariableDescriptions)
        {
            if (declarations.TryGetValue(name, out var line) && !Annotated(AnnotationKind.VariableDescription, name))
            {
                items.Add(new DriftItem(DriftKind.AttributeNotAnnotated, line,
                    $"{name}'s VB_VarDescription is \"{text}\", and no '@VariableDescription says so.",
                    AnnotationKind.VariableDescription, name, 0));
            }
        }
    }
}
