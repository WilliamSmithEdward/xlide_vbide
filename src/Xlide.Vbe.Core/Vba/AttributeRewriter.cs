using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Xlide.Vbe.Core.Sync;

namespace Xlide.Vbe.Core.Vba;

/// <summary>One attribute the rewrite set, changed or took away, in words a notice can carry.</summary>
public sealed record AttributeChange(string Target, string Attribute, string? From, string? To)
{
    public override string ToString() => To is null
        ? $"{Target}: {Attribute} removed"
        : From is null ? $"{Target}: {Attribute} = {To}" : $"{Target}: {Attribute} {From} -> {To}";
}

/// <summary>The rewritten export and what changed in it. An empty change list means the text is as it was.</summary>
public sealed record RewriteResult(string Text, IReadOnlyList<AttributeChange> Changes, IReadOnlyList<string> Skipped);

/// <summary>
/// Writes attributes into an exported module's text, so that an Import puts them on the module.
///
/// The text is edited in place and otherwise left alone: every line that is not the attribute
/// being set stays where it was, byte for byte, because the file is about to be imported over the
/// developer's module and anything this touched by accident would be a change to their code. The
/// managed attributes are the ones the annotations stand for; an attribute this product does not
/// know is carried through untouched.
/// </summary>
public static class AttributeRewriter
{
    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

    private static readonly Regex ProcedureHeader = new(
        @"^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property\s+(?:get|let|set))\s+(?<name>\p{L}[\p{L}\p{N}_]*)",
        Options);

    private static readonly Regex VariableDeclaration = new(
        @"^\s*(?:dim|private|public|global)\s+(?:withevents\s+)?(?<name>\p{L}[\p{L}\p{N}_]*)\b(?!\s*\()",
        Options);

    private static readonly Regex OwnedAttribute = new(
        @"^\s*Attribute\s+(?<owner>\p{L}[\p{L}\p{N}_]*)\.(?<name>VB_[A-Za-z_]+(?:\.VB_[A-Za-z_]+)?)\s*=",
        Options);

    private static readonly Regex ModuleAttribute = new(
        @"^\s*Attribute\s+(?<name>VB_[A-Za-z_]+)\s*=\s*(?<value>.*?)\s*$",
        Options);

    /// <summary>
    /// Writes every attribute the annotations name. Attributes the annotations say nothing about
    /// are left exactly as the export had them; taking one away is <see cref="Remove"/>, a
    /// separate and deliberate act.
    /// </summary>
    public static RewriteResult Apply(string exported, ModuleAnnotations annotations)
    {
        ArgumentNullException.ThrowIfNull(exported);
        ArgumentNullException.ThrowIfNull(annotations);

        var text = new ExportText(exported);
        var changes = new List<AttributeChange>();
        var skipped = new List<string>();

        foreach (var annotation in annotations.Annotations)
        {
            switch (annotation.Kind)
            {
                case AnnotationKind.ModuleDescription:
                    text.SetModule("VB_Description", ModuleAttributes.Literal(annotation.Argument ?? string.Empty), changes, skipped, canInsert: true);
                    break;
                case AnnotationKind.PredeclaredId:
                    text.SetModule("VB_PredeclaredId", "True", changes, skipped, canInsert: false);
                    break;
                case AnnotationKind.Exposed:
                    text.SetModule("VB_Exposed", "True", changes, skipped, canInsert: false);
                    break;
                case AnnotationKind.Description:
                    text.SetMember(annotation, "VB_Description", ModuleAttributes.Literal(annotation.Argument ?? string.Empty), changes, skipped);
                    break;
                case AnnotationKind.DefaultMember:
                    text.SetMember(annotation, "VB_UserMemId", "0", changes, skipped);
                    break;
                case AnnotationKind.Enumerator:
                    text.SetMember(annotation, "VB_UserMemId", "-4", changes, skipped);
                    break;
                case AnnotationKind.ExcelHotkey:
                    text.SetMember(annotation, "VB_ProcData.VB_Invoke_Func", ModuleAttributes.Literal(ModuleAttributes.InvokeFuncFor(annotation.Argument ?? string.Empty)), changes, skipped);
                    break;
                case AnnotationKind.VariableDescription:
                    text.SetVariable(annotation, ModuleAttributes.Literal(annotation.Argument ?? string.Empty), changes, skipped);
                    break;
            }
        }

        return new RewriteResult(text.ToString(), changes, skipped);
    }

    /// <summary>
    /// Takes one managed attribute away: a boolean module attribute goes back to False, any other
    /// line is removed. <paramref name="target"/> names the procedure or variable, or is null for
    /// the module; <paramref name="occurrence"/> picks the leg when a property's legs share a name.
    /// </summary>
    public static RewriteResult Remove(string exported, AnnotationKind kind, string? target, int occurrence = 0)
    {
        ArgumentNullException.ThrowIfNull(exported);

        var text = new ExportText(exported);
        var changes = new List<AttributeChange>();
        var skipped = new List<string>();

        switch (kind)
        {
            case AnnotationKind.ModuleDescription:
                text.RemoveModule("VB_Description", changes);
                break;
            case AnnotationKind.PredeclaredId:
                text.SetModule("VB_PredeclaredId", "False", changes, skipped, canInsert: false);
                break;
            case AnnotationKind.Exposed:
                text.SetModule("VB_Exposed", "False", changes, skipped, canInsert: false);
                break;
            case AnnotationKind.Description when target is not null:
                text.RemoveMember(target, occurrence, "VB_Description", changes, skipped);
                break;
            case AnnotationKind.DefaultMember or AnnotationKind.Enumerator when target is not null:
                text.RemoveMember(target, occurrence, "VB_UserMemId", changes, skipped);
                break;
            case AnnotationKind.ExcelHotkey when target is not null:
                text.RemoveMember(target, occurrence, "VB_ProcData.VB_Invoke_Func", changes, skipped);
                break;
            case AnnotationKind.VariableDescription when target is not null:
                text.RemoveVariable(target, changes, skipped);
                break;
            default:
                skipped.Add($"{AttributeAnnotations.Spelled(kind)} needs a procedure or variable to take it from.");
                break;
        }

        return new RewriteResult(text.ToString(), changes, skipped);
    }

    /// <summary>
    /// Which occurrence of a procedure's name an annotation's target header is: a property's Get
    /// and Let share a name, and the annotation above the second binds to the second.
    /// </summary>
    public static int OccurrenceOf(string source, string target, int targetLine)
    {
        var seen = 0;
        var lines = source.Split('\n');
        for (var at = 0; at < Math.Min(lines.Length, targetLine - 1); at++)
        {
            var header = ProcedureHeader.Match(lines[at].TrimEnd('\r'));
            if (header.Success && string.Equals(header.Groups["name"].Value, target, StringComparison.OrdinalIgnoreCase))
            {
                seen++;
            }
        }
        return seen;
    }

    /// <summary>The export as lines, with the edits that keep everything else in place.</summary>
    private sealed class ExportText
    {
        private readonly List<string> lines;
        private readonly string newline;

        public ExportText(string exported)
        {
            newline = exported.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
            lines = [.. exported.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n')];
        }

        public override string ToString() => string.Join(newline, lines);

        /// <summary>The index one past the last header attribute or preamble line.</summary>
        private int HeaderEnd()
        {
            var end = 0;
            for (var at = 0; at < lines.Count; at++)
            {
                if (ModuleSync.IsAttributeLine(lines[at]) || ModuleSync.IsHeaderPreamble(lines[at]))
                {
                    end = at + 1;
                    continue;
                }
                if (end > 0)
                {
                    break;
                }
            }
            return end;
        }

        private int ModuleAttributeIndex(string attribute)
        {
            var end = HeaderEnd();
            for (var at = 0; at < end; at++)
            {
                var match = ModuleAttribute.Match(lines[at]);
                if (match.Success && match.Groups["name"].Value.Equals(attribute, StringComparison.OrdinalIgnoreCase))
                {
                    return at;
                }
            }
            return -1;
        }

        public void SetModule(string attribute, string value, List<AttributeChange> changes, List<string> skipped, bool canInsert)
        {
            var at = ModuleAttributeIndex(attribute);
            if (at >= 0)
            {
                var was = ModuleAttribute.Match(lines[at]).Groups["value"].Value;
                if (!string.Equals(was, value, StringComparison.Ordinal))
                {
                    lines[at] = $"Attribute {attribute} = {value}";
                    changes.Add(new AttributeChange("module", attribute, was, value));
                }
                return;
            }

            if (!canInsert)
            {
                // A standard module's header has no VB_PredeclaredId or VB_Exposed line, and the
                // editor gives those meaning only on a class.
                skipped.Add($"{attribute} is not an attribute this kind of module carries.");
                return;
            }

            var end = HeaderEnd();
            if (end == 0)
            {
                skipped.Add($"the export has no header to put {attribute} in.");
                return;
            }
            lines.Insert(end, $"Attribute {attribute} = {value}");
            changes.Add(new AttributeChange("module", attribute, null, value));
        }

        public void RemoveModule(string attribute, List<AttributeChange> changes)
        {
            var at = ModuleAttributeIndex(attribute);
            if (at < 0)
            {
                return;
            }
            var was = ModuleAttribute.Match(lines[at]).Groups["value"].Value;
            lines.RemoveAt(at);
            changes.Add(new AttributeChange("module", attribute, was, null));
        }

        /// <summary>The line index of the last line of the nth header named <paramref name="name"/>, or -1.</summary>
        private int HeaderIndex(string name, int occurrence)
        {
            var seen = 0;
            for (var at = HeaderEnd(); at < lines.Count; at++)
            {
                var match = ProcedureHeader.Match(lines[at]);
                if (!match.Success || !match.Groups["name"].Value.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (seen++ != occurrence)
                {
                    continue;
                }
                // The attributes follow the header's LAST line, which is the last one ending in a
                // continuation character.
                var last = at;
                while (last < lines.Count - 1 && lines[last].TrimEnd().EndsWith('_'))
                {
                    last++;
                }
                return last;
            }
            return -1;
        }

        private int VariableIndex(string name)
        {
            var headerEnd = HeaderEnd();
            for (var at = headerEnd; at < lines.Count; at++)
            {
                if (ProcedureHeader.IsMatch(lines[at]))
                {
                    break;
                }
                var match = VariableDeclaration.Match(lines[at]);
                if (match.Success && match.Groups["name"].Value.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return at;
                }
            }
            return -1;
        }

        /// <summary>The index of the owned attribute line directly under <paramref name="after"/>, or -1.</summary>
        private int OwnedIndex(int after, string owner, string attribute)
        {
            for (var at = after + 1; at < lines.Count; at++)
            {
                var match = OwnedAttribute.Match(lines[at]);
                if (!match.Success)
                {
                    return -1;
                }
                if (match.Groups["owner"].Value.Equals(owner, StringComparison.OrdinalIgnoreCase)
                    && match.Groups["name"].Value.Equals(attribute, StringComparison.OrdinalIgnoreCase))
                {
                    return at;
                }
            }
            return -1;
        }

        private void SetOwned(int after, string owner, string attribute, string value, List<AttributeChange> changes)
        {
            var at = OwnedIndex(after, owner, attribute);
            var line = $"Attribute {owner}.{attribute} = {value}";
            if (at >= 0)
            {
                var was = lines[at][(lines[at].IndexOf('=') + 1)..].Trim();
                if (!string.Equals(was, value, StringComparison.Ordinal))
                {
                    lines[at] = line;
                    changes.Add(new AttributeChange(owner, attribute, was, value));
                }
                return;
            }
            lines.Insert(after + 1, line);
            changes.Add(new AttributeChange(owner, attribute, null, value));
        }

        private void RemoveOwned(int after, string owner, string attribute, List<AttributeChange> changes)
        {
            var at = OwnedIndex(after, owner, attribute);
            if (at < 0)
            {
                return;
            }
            var was = lines[at][(lines[at].IndexOf('=') + 1)..].Trim();
            lines.RemoveAt(at);
            changes.Add(new AttributeChange(owner, attribute, was, null));
        }

        public void SetMember(Annotation annotation, string attribute, string value, List<AttributeChange> changes, List<string> skipped)
        {
            var owner = annotation.Target!;
            var header = HeaderIndex(owner, OccurrenceIn(annotation));
            if (header < 0)
            {
                skipped.Add($"no procedure named '{owner}' was found in the export for {AttributeAnnotations.Spelled(annotation.Kind)}.");
                return;
            }
            SetOwned(header, owner, attribute, value, changes);
        }

        public void RemoveMember(string owner, int occurrence, string attribute, List<AttributeChange> changes, List<string> skipped)
        {
            var header = HeaderIndex(owner, occurrence);
            if (header < 0)
            {
                skipped.Add($"no procedure named '{owner}' was found in the export.");
                return;
            }
            RemoveOwned(header, owner, attribute, changes);
        }

        public void SetVariable(Annotation annotation, string value, List<AttributeChange> changes, List<string> skipped)
        {
            var owner = annotation.Target!;
            var at = VariableIndex(owner);
            if (at < 0)
            {
                skipped.Add($"no module-level variable named '{owner}' was found in the export for '@VariableDescription.");
                return;
            }
            SetOwned(at, owner, "VB_VarDescription", value, changes);
        }

        public void RemoveVariable(string owner, List<AttributeChange> changes, List<string> skipped)
        {
            var at = VariableIndex(owner);
            if (at < 0)
            {
                skipped.Add($"no module-level variable named '{owner}' was found in the export.");
                return;
            }
            RemoveOwned(at, owner, "VB_VarDescription", changes);
        }

        /// <summary>
        /// The annotation's target as the nth header of its name. The annotation's own line
        /// numbers are code coordinates, and this text holds the header plus attribute lines, so
        /// the count of same-named headers ABOVE the target is what carries across.
        /// </summary>
        private int OccurrenceIn(Annotation annotation)
        {
            // Reconstruct the code coordinates from this text: every non-attribute, non-preamble
            // line is a code line, in order.
            var codeLine = 0;
            var seen = 0;
            for (var at = 0; at < lines.Count; at++)
            {
                if (ModuleSync.IsAttributeLine(lines[at]) || (at < HeaderEnd() && ModuleSync.IsHeaderPreamble(lines[at])))
                {
                    continue;
                }
                codeLine++;
                if (codeLine >= (annotation.TargetLine ?? int.MaxValue))
                {
                    break;
                }
                var match = ProcedureHeader.Match(lines[at]);
                if (match.Success && match.Groups["name"].Value.Equals(annotation.Target, StringComparison.OrdinalIgnoreCase))
                {
                    seen++;
                }
            }
            return seen;
        }
    }
}
