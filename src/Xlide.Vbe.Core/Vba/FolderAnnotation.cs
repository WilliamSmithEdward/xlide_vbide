using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Xlide.Vbe.Core.Vba;

/// <summary>
/// The folder a module says it belongs to, read off a comment at the top of its code.
///
/// The convention is Rubberduck's: a comment of the form <c>'@Folder("Parent.Child")</c> in the
/// declarations section, with a dot between a parent and its child. A project that already
/// carries these keeps its organisation here without editing a line, and one written here reads
/// the same way there. The folders are not a thing the workbook stores: they are a way of
/// DRAWING the annotations that are already in the modules, which is also how the original
/// describes its own ("folders don't really exist: they are merely a UI rendering of annotation
/// comments in existing modules").
///
/// What is accepted is wider than what is written. The documented form is the one above; the
/// same annotation is also read without its parentheses, without its quotes, and in any letter
/// case, because the comment is typed by hand and a folder that vanishes over a missing bracket
/// would read as a tree that lost a module. <see cref="Canonical"/> is the one spelling this
/// product ever writes.
///
/// Only the declarations section is read. An annotation below the first procedure is a comment
/// like any other, which is where the convention puts it too; and the first one found is the one
/// that counts, so a module carrying two does not flicker between them.
/// </summary>
public static class FolderAnnotation
{
    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

    // 'Folder("A.B")   'Folder "A.B"   'Folder(A.B)   'Folder A.B   - and @folder in any case.
    // The name has to END there: whitespace, a bracket, a quote or the line's end may follow it,
    // so "@Folder-ish" and "@Folders" are prose. The parenthesised, quoted form is tried first
    // so a path holding a space survives; the bare form ends at whitespace because the rest of
    // such a line is prose.
    private static readonly Regex Annotation = new(
        @"^\s*'\s*@folder(?=\s|\(|""|$)\s*(?:\(\s*""(?<quoted>[^""]*)""\s*\)|""(?<bare>[^""]*)""|\(\s*(?<unquoted>[^)""]*?)\s*\)|(?<word>\S+))?",
        Options);

    // The first line of a procedure ends the declarations section. Visibility and Static are
    // optional; the keyword that decides is the one right after them.
    private static readonly Regex ProcedureHeader = new(
        @"^\s*(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property)\b",
        Options);

    /// <summary>
    /// The folder path the module's declarations section names, normalized: every segment
    /// trimmed, empty segments dropped, joined by dots. Null when there is no annotation, or the
    /// annotation names nothing.
    /// </summary>
    public static string? Of(string? source)
    {
        if (string.IsNullOrEmpty(source))
        {
            return null;
        }

        foreach (var line in Lines(source))
        {
            if (ProcedureHeader.IsMatch(line))
            {
                return null;
            }

            var match = Annotation.Match(line);
            if (!match.Success)
            {
                continue;
            }

            var raw = match.Groups["quoted"].Success ? match.Groups["quoted"].Value
                : match.Groups["bare"].Success ? match.Groups["bare"].Value
                : match.Groups["unquoted"].Success ? match.Groups["unquoted"].Value
                : match.Groups["word"].Success ? match.Groups["word"].Value
                : string.Empty;

            return Normalize(raw);
        }

        return null;
    }

    /// <summary>A path with its segments trimmed and the empty ones dropped, or null when nothing is left.</summary>
    public static string? Normalize(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        var kept = new List<string>();
        foreach (var segment in path.Split('.'))
        {
            var trimmed = segment.Trim();
            if (trimmed.Length > 0)
            {
                kept.Add(trimmed);
            }
        }

        return kept.Count == 0 ? null : string.Join('.', kept);
    }

    /// <summary>The segments of a normalized path, root first.</summary>
    public static string[] Segments(string path) => path.Split('.');

    /// <summary>The one spelling this product writes: the documented, parenthesised, quoted form.</summary>
    public static string Canonical(string path) => $"'@Folder(\"{path}\")";

    private static IEnumerable<string> Lines(string source)
    {
        var start = 0;
        for (var at = 0; at < source.Length; at++)
        {
            var c = source[at];
            if (c != '\r' && c != '\n')
            {
                continue;
            }

            yield return source[start..at];
            if (c == '\r' && at + 1 < source.Length && source[at + 1] == '\n')
            {
                at++;
            }

            start = at + 1;
        }

        if (start < source.Length)
        {
            yield return source[start..];
        }
    }
}
