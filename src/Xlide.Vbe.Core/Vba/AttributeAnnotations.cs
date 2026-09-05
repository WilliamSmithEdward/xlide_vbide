using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Vba;

/// <summary>Which hidden attribute an annotation stands for.</summary>
public enum AnnotationKind
{
    /// <summary><c>'@ModuleDescription("text")</c>: the module's <c>VB_Description</c>.</summary>
    ModuleDescription,
    /// <summary><c>'@PredeclaredId</c>: <c>VB_PredeclaredId = True</c>, a class with a default instance.</summary>
    PredeclaredId,
    /// <summary><c>'@Exposed</c>: <c>VB_Exposed = True</c>, a class other projects can see.</summary>
    Exposed,
    /// <summary><c>'@Description("text")</c> above a procedure: its <c>VB_Description</c>.</summary>
    Description,
    /// <summary><c>'@DefaultMember</c> above a procedure: <c>VB_UserMemId = 0</c>.</summary>
    DefaultMember,
    /// <summary><c>'@Enumerator</c> above a procedure: <c>VB_UserMemId = -4</c>, what For Each walks.</summary>
    Enumerator,
    /// <summary><c>'@ExcelHotkey("D")</c> above a procedure: <c>VB_ProcData.VB_Invoke_Func</c>, a Ctrl key.</summary>
    ExcelHotkey,
    /// <summary><c>'@VariableDescription("text")</c> above a module-level variable: its <c>VB_VarDescription</c>.</summary>
    VariableDescription,
}

/// <summary>
/// One annotation as read from the code: where it is, what it says, and the procedure or
/// variable it is bound to - null for a module annotation.
/// </summary>
/// <param name="Kind">Which attribute it stands for.</param>
/// <param name="Line">The 1-based line the comment is on.</param>
/// <param name="Argument">The text between the brackets, unquoted, or null for a bare annotation.</param>
/// <param name="Target">The procedure or variable the annotation is above, or null for the module.</param>
/// <param name="TargetLine">The 1-based line of that procedure's header or variable's declaration.</param>
public sealed record Annotation(AnnotationKind Kind, int Line, string? Argument, string? Target, int? TargetLine)
{
    /// <summary>The spelling this product writes for it.</summary>
    public string Canonical => Kind switch
    {
        AnnotationKind.ModuleDescription => $"'@ModuleDescription(\"{Argument}\")",
        AnnotationKind.Description => $"'@Description(\"{Argument}\")",
        AnnotationKind.ExcelHotkey => $"'@ExcelHotkey(\"{Argument}\")",
        AnnotationKind.VariableDescription => $"'@VariableDescription(\"{Argument}\")",
        _ => $"'@{Kind}",
    };
}

/// <summary>An annotation that cannot mean what it says where it sits, with the reason.</summary>
public sealed record AnnotationProblem(int Line, string Message);

/// <summary>Everything the annotations of one module say.</summary>
public sealed record ModuleAnnotations(IReadOnlyList<Annotation> Annotations, IReadOnlyList<AnnotationProblem> Problems)
{
    public static readonly ModuleAnnotations None = new([], []);
}

/// <summary>
/// Reads the annotations that stand for hidden attributes out of a module's code.
///
/// The convention is Rubberduck's. A VBA module carries attributes the code pane never shows -
/// <c>VB_PredeclaredId</c>, <c>VB_Description</c>, <c>VB_UserMemId</c> and the rest - and the
/// editor offers no way to set them; they decide whether a class has a default instance, what
/// <c>For Each</c> walks, and what IntelliSense says about a member. An annotation is a comment
/// that names the attribute the developer wants, in the code, where it can be read, diffed and
/// edited; this product then writes the attribute to match (see AttributeRewriter).
///
/// PLACEMENT decides what an annotation binds to. A module annotation lives in the declarations
/// section, above the first procedure, anywhere in it. A member annotation lives in the run of
/// comment and blank lines directly above a procedure's header, and binds to that procedure; the
/// same run above a module-level variable binds a <c>'@VariableDescription</c> to the variable. A
/// member annotation above anything else, or a module annotation below the first procedure, is a
/// problem rather than a guess, reported with its line.
///
/// Read leniently: the name in any case, the argument in brackets with or without quotes, or
/// after a space. The canonical spelling is the documented one, and it is the one written.
/// </summary>
public static class AttributeAnnotations
{
    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

    // '@Name("arg")   '@Name(arg)   '@Name "arg"   '@Name arg   '@Name
    private static readonly Regex AnnotationLine = new(
        @"^\s*'\s*@(?<name>[A-Za-z]+)(?=\s|\(|""|$)\s*(?:\(\s*(?:""(?<quoted>(?:[^""]|"""")*)""|(?<unquoted>[^)]*?))\s*\)|""(?<bare>(?:[^""]|"""")*)""|(?<word>\S+))?\s*$",
        Options);

    private static readonly Regex ProcedureHeader = new(
        @"^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(?<name>\p{L}[\p{L}\p{N}_]*)",
        Options);

    // A module-level variable: Dim/Private/Public/Global, optionally WithEvents, then the name.
    // Const, Type, Enum, Declare, Event and Implements are declarations too, but none of them
    // carries a VB_VarDescription.
    private static readonly Regex VariableDeclaration = new(
        @"^\s*(?:dim|private|public|global)\s+(?:withevents\s+)?(?<name>\p{L}[\p{L}\p{N}_]*)\b(?!\s*\()",
        Options);

    private static readonly Regex BlankOrComment = new(@"^\s*(?:'|Rem\b|$)", Options);

    /// <summary>The annotations of a module's code, with the problems found on the way.</summary>
    public static ModuleAnnotations Read(string? source)
    {
        if (string.IsNullOrEmpty(source))
        {
            return ModuleAnnotations.None;
        }

        var found = new List<Annotation>();
        var problems = new List<AnnotationProblem>();
        var pending = new List<(AnnotationKind Kind, int Line, string? Argument)>();
        var inDeclarations = true;
        var defaultMemberAt = 0;

        var lines = source.Split('\n');
        for (var at = 0; at < lines.Length; at++)
        {
            var line = lines[at].TrimEnd('\r');
            var number = at + 1;

            var match = AnnotationLine.Match(line);
            if (match.Success)
            {
                if (!TryKind(match.Groups["name"].Value, out var kind))
                {
                    // Not one of ours: '@Folder, '@Ignore, '@TestMethod and any prose that starts
                    // with an at-sign are somebody else's, and are left alone.
                    continue;
                }

                var argument = ArgumentOf(match);
                if (!IsModuleAnnotation(kind))
                {
                    pending.Add((kind, number, argument));
                    continue;
                }

                if (!inDeclarations)
                {
                    problems.Add(new AnnotationProblem(number, $"{Spelled(kind)} is a module annotation and belongs in the declarations section, above the first procedure."));
                    continue;
                }

                if (NeedsArgument(kind) && string.IsNullOrEmpty(argument))
                {
                    problems.Add(new AnnotationProblem(number, $"{Spelled(kind)} needs the text to write, in brackets: {Spelled(kind)}(\"...\")."));
                    continue;
                }

                if (found.Exists(one => one.Kind == kind))
                {
                    problems.Add(new AnnotationProblem(number, $"{Spelled(kind)} appears more than once; the first one counts."));
                    continue;
                }

                found.Add(new Annotation(kind, number, argument, null, null));
                continue;
            }

            if (BlankOrComment.IsMatch(line))
            {
                continue;
            }

            // A line of code. Whatever member annotations were pending bind to it or fail here.
            var header = ProcedureHeader.Match(line);
            if (header.Success)
            {
                inDeclarations = false;
                BindToProcedure(header.Groups["name"].Value, number);
                continue;
            }

            var variable = inDeclarations ? VariableDeclaration.Match(line) : Match.Empty;
            if (variable.Success)
            {
                BindToVariable(variable.Groups["name"].Value, number);
                continue;
            }

            foreach (var (kind, at2, _) in pending)
            {
                problems.Add(new AnnotationProblem(at2, $"{Spelled(kind)} is above a line that is not a {(kind == AnnotationKind.VariableDescription ? "module-level variable" : "procedure")}, so there is nothing to bind it to."));
            }
            pending.Clear();
        }

        foreach (var (kind, at, _) in pending)
        {
            problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} is above nothing, so there is nothing to bind it to."));
        }

        return new ModuleAnnotations(found, problems);

        void BindToProcedure(string name, int headerLine)
        {
            foreach (var (kind, at, argument) in pending)
            {
                if (kind == AnnotationKind.VariableDescription)
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} describes a module-level variable, and '{name}' is a procedure. Use '@Description for a procedure."));
                    continue;
                }

                if (NeedsArgument(kind) && string.IsNullOrEmpty(argument))
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} needs the text to write, in brackets: {Spelled(kind)}(\"...\")."));
                    continue;
                }

                if (kind == AnnotationKind.ExcelHotkey && !IsHotkey(argument))
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} takes one letter: a lower-case letter is Ctrl+letter, an upper-case one Ctrl+Shift+letter."));
                    continue;
                }

                if (kind == AnnotationKind.DefaultMember)
                {
                    if (defaultMemberAt != 0)
                    {
                        problems.Add(new AnnotationProblem(at, $"'@DefaultMember appears again; a class has one default member, and line {defaultMemberAt} already names it."));
                        continue;
                    }
                    defaultMemberAt = at;
                }

                if (found.Exists(one => one.Kind == kind && one.Target == name && one.TargetLine == headerLine))
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} appears more than once above '{name}'; the first one counts."));
                    continue;
                }

                found.Add(new Annotation(kind, at, argument, name, headerLine));
            }
            pending.Clear();
        }

        void BindToVariable(string name, int declarationLine)
        {
            foreach (var (kind, at, argument) in pending)
            {
                if (kind != AnnotationKind.VariableDescription)
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} describes a procedure, and '{name}' is a variable. Use '@VariableDescription for a variable."));
                    continue;
                }

                if (string.IsNullOrEmpty(argument))
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} needs the text to write, in brackets: {Spelled(kind)}(\"...\")."));
                    continue;
                }

                if (found.Exists(one => one.Kind == kind && one.Target == name))
                {
                    problems.Add(new AnnotationProblem(at, $"{Spelled(kind)} appears more than once above '{name}'; the first one counts."));
                    continue;
                }

                found.Add(new Annotation(kind, at, argument, name, declarationLine));
            }
            pending.Clear();
        }
    }

    /// <summary>The documented spelling of a kind's name, apostrophe and at-sign included.</summary>
    public static string Spelled(AnnotationKind kind) => $"'@{kind}";

    /// <summary>Whether the kind binds to the module rather than to a procedure or variable.</summary>
    public static bool IsModuleAnnotation(AnnotationKind kind) =>
        kind is AnnotationKind.ModuleDescription or AnnotationKind.PredeclaredId or AnnotationKind.Exposed;

    /// <summary>Whether the kind carries text.</summary>
    public static bool NeedsArgument(AnnotationKind kind) =>
        kind is AnnotationKind.ModuleDescription or AnnotationKind.Description
            or AnnotationKind.ExcelHotkey or AnnotationKind.VariableDescription;

    /// <summary>A hotkey is one letter; its case says whether Shift is held.</summary>
    public static bool IsHotkey(string? argument) =>
        argument is { Length: 1 } && char.IsLetter(argument[0]);

    private static bool TryKind(string name, out AnnotationKind kind) =>
        Enum.TryParse(name, ignoreCase: true, out kind) && Enum.IsDefined(kind);

    private static string? ArgumentOf(Match match)
    {
        if (match.Groups["quoted"].Success)
        {
            return match.Groups["quoted"].Value.Replace("\"\"", "\"", StringComparison.Ordinal);
        }
        if (match.Groups["bare"].Success)
        {
            return match.Groups["bare"].Value.Replace("\"\"", "\"", StringComparison.Ordinal);
        }
        if (match.Groups["unquoted"].Success && match.Groups["unquoted"].Value.Length > 0)
        {
            return match.Groups["unquoted"].Value;
        }
        if (match.Groups["word"].Success)
        {
            return match.Groups["word"].Value;
        }
        return null;
    }
}
