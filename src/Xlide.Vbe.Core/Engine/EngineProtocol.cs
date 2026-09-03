using System.Text.Json;
using System.Text.Json.Serialization;

namespace Xlide.Vbe.Core.Engine;

/// <summary>One module as the add-in reads it out of the editor.</summary>
public sealed record EngineModule(
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("source")] string Source,
    [property: JsonPropertyName("type")] string Type,
    /// <summary>
    /// Members the module's text never declares but the host knows are there: a UserForm's
    /// controls, name and type each, read off the designer. The analyzer resolves them and
    /// completion offers their types' members (xlide_vscode#17); absent for every other kind.
    /// </summary>
    [property: JsonPropertyName("implicitMembers")] EngineImplicitMember[]? ImplicitMembers = null,
    /// <summary>
    /// True when the module carries `Attribute VB_PredeclaredId = True`, which gives a class a
    /// default instance and so makes its own name usable as a value: `Ticket.ChangeTest` compiles
    /// against a predeclared class and is `Variable not defined` against a plain one
    /// (xlide_vscode#47).
    ///
    /// THREE STATES, and null is not false. The attribute never appears in a code pane, so it is
    /// read from the document already saved on disk (<see cref="Vba.SavedModules"/>) and a module
    /// the file cannot answer for stays null. A null sent as false would put a red squiggle under
    /// every use of every predeclared singleton in the project.
    /// </summary>
    [property: JsonPropertyName("predeclaredId")] bool? PredeclaredId = null);

/// <summary>One implicit member: the control's name, and the type completion resolves it as.</summary>
public sealed record EngineImplicitMember(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("type")] string Type);

/// <summary>A finding, with a span in character offsets into the module source.</summary>
public sealed record EngineDiagnostic(
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("span")] EngineSpan Span,
    [property: JsonPropertyName("at")] EnginePosition? At = null);

public sealed record EngineSpan(
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End);

/// <summary>
/// The span as one-based lines and columns, measured by the engine against the text it analysed.
///
/// Null only from an engine older than this field. An offset is meaningless without the text it
/// was counted in, and a request that sends no source leaves the engine to choose between its live
/// copy and its seeded one, a choice the caller cannot see. Converting here from a text the caller
/// merely happened to hold put findings columns away from the words they were about.
/// </summary>
public sealed record EnginePosition(
    [property: JsonPropertyName("startLine")] int StartLine,
    [property: JsonPropertyName("startColumn")] int StartColumn,
    [property: JsonPropertyName("endLine")] int EndLine,
    [property: JsonPropertyName("endColumn")] int EndColumn);

public sealed record EngineDiagnostics(
    [property: JsonPropertyName("diagnostics")] EngineDiagnostic[] Diagnostics,
    [property: JsonPropertyName("mode")] string? Mode);

/// <summary>
/// One analyzer rule as `analysis/rules` lists it: its stable code, its words, and the severity
/// moves the analyzer permits. An empty `allowed` is a rule that cannot be changed - most error
/// rules mirror a VBE compile failure and take no override at all.
/// </summary>
public sealed record EngineAnalysisRule(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("title")] string Title,
    [property: JsonPropertyName("category")] string Category,
    [property: JsonPropertyName("defaultSeverity")] string DefaultSeverity,
    [property: JsonPropertyName("allowed")] string[] Allowed,
    [property: JsonPropertyName("suppressionScopes")] string[] SuppressionScopes);

public sealed record EngineAnalysisRules(
    [property: JsonPropertyName("rules")] EngineAnalysisRule[] Rules);

/// <summary>One completion, in the analyzer's own vocabulary of kinds.</summary>
public sealed record EngineCompletionItem(
    [property: JsonPropertyName("label")] string Label,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("detail")] string? Detail,
    [property: JsonPropertyName("documentation")] string? Documentation,
    [property: JsonPropertyName("insertText")] string? InsertText,
    [property: JsonPropertyName("filterText")] string? FilterText,
    [property: JsonPropertyName("sortText")] string? SortText);

public sealed record EngineCompletions(
    [property: JsonPropertyName("items")] EngineCompletionItem[] Items);

/// <summary>A resolved hover: the declaration line, plain-text facts, and documentation.</summary>
public sealed record EngineHoverPayload(
    [property: JsonPropertyName("signature")] string Signature,
    [property: JsonPropertyName("details")] string[] Details,
    [property: JsonPropertyName("documentation")] string? Documentation,
    [property: JsonPropertyName("span")] EngineSpan Span);

public sealed record EngineHover(
    [property: JsonPropertyName("hover")] EngineHoverPayload? Hover);

/// <summary>One parameter slot, its label exactly as it appears in the signature line.</summary>
public sealed record EngineSignatureParameter(
    [property: JsonPropertyName("label")] string Label,
    [property: JsonPropertyName("documentation")] string? Documentation);

/// <summary>A resolved call tip: the signature line and which parameter is active.</summary>
public sealed record EngineSignatureInfo(
    [property: JsonPropertyName("label")] string Label,
    [property: JsonPropertyName("parameters")] EngineSignatureParameter[] Parameters,
    [property: JsonPropertyName("activeParameter")] int ActiveParameter,
    [property: JsonPropertyName("documentation")] string? Documentation,
    [property: JsonPropertyName("details")] string[]? Details);

public sealed record EngineSignatureHelp(
    [property: JsonPropertyName("signature")] EngineSignatureInfo? Signature);

/// <summary>A text replacement, offsets into the request's source; an insertion has Start == End.</summary>
public sealed record EngineTextEdit(
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End,
    [property: JsonPropertyName("text")] string Text);

/// <summary>What Enter should leave behind: edits, and where the caret belongs once they apply.</summary>
public sealed record EngineSmartEnter(
    [property: JsonPropertyName("edits")] EngineTextEdit[] Edits,
    [property: JsonPropertyName("caret")] int? Caret);

/// <summary>A plain set of edits: canonical casing and loop-iterator sync both answer with one.</summary>
public sealed record EngineTextEdits(
    [property: JsonPropertyName("edits")] EngineTextEdit[] Edits);

/// <summary>
/// One quick fix: what to call it, the finding it answers, and the edits that apply it.
/// </summary>
public sealed record EngineCodeAction(
    [property: JsonPropertyName("title")] string Title,
    [property: JsonPropertyName("isPreferred")] bool? IsPreferred,
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("span")] EngineSpan Span,
    [property: JsonPropertyName("edits")] EngineTextEdit[] Edits);

/// <summary>The quick fixes offered over a span.</summary>
public sealed record EngineCodeActions(
    [property: JsonPropertyName("actions")] EngineCodeAction[] Actions);

/// <summary>One module a rename rewrites: its name, what it says afterwards, and how many went.</summary>
public sealed record EngineRenamedModule(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("source")] string Source,
    [property: JsonPropertyName("replaced")] int Replaced);

/// <summary>
/// The result of a rename: every module it rewrites, whole. Refused says why nothing changed -
/// a rename that cannot reach every use must do none of them.
/// </summary>
public sealed record EngineRename(
    [property: JsonPropertyName("modules")] EngineRenamedModule[] Modules,
    [property: JsonPropertyName("oldName")] string? OldName,
    [property: JsonPropertyName("refused")] string? Refused,
    /// <summary>
    /// Set when what is being renamed is a MODULE rather than a symbol in one. Its name lives on
    /// the component, not in any module's text, so the add-in owns that half.
    /// </summary>
    [property: JsonPropertyName("module")] string? Module = null);

/// <summary>
/// The result of an Extract Method: the one module it rewrites, whole. Refused says why nothing
/// changed, in the words the surface shows - a refactoring that cannot be made safely must make
/// no part of itself.
/// </summary>
public sealed record EngineExtractMethod(
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("source")] string? Source,
    /// <summary>The new procedure's name, as it was actually created.</summary>
    [property: JsonPropertyName("procedure")] string? Procedure,
    /// <summary>Its header, which is the part a developer checks first.</summary>
    [property: JsonPropertyName("signature")] string? Signature,
    /// <summary>The procedure the statements came out of.</summary>
    [property: JsonPropertyName("from")] string? From,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The result of an Implement Interface: the one class it writes stubs into, whole. Refused says
/// why nothing was written.
/// </summary>
public sealed record EngineImplementInterface(
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("source")] string? Source,
    /// <summary>The interfaces answered, so the surface can say whose members these are.</summary>
    [property: JsonPropertyName("interfaces")] string[]? Interfaces,
    /// <summary>The members written, by the name each stub carries.</summary>
    [property: JsonPropertyName("added")] string[]? Added,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The result of an Encapsulate Field: the one module it rewrites, whole. Refused says why nothing
/// changed.
/// </summary>
public sealed record EngineEncapsulateField(
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("source")] string? Source,
    /// <summary>The name that is now a property.</summary>
    [property: JsonPropertyName("field")] string? Field,
    /// <summary>The private variable it now stands in front of.</summary>
    [property: JsonPropertyName("backingField")] string? BackingField,
    /// <summary>The two members written.</summary>
    [property: JsonPropertyName("accessors")] string[]? Accessors,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The result of an Extract Variable: the one module it rewrites, whole, plus what the analyzer
/// said the expression was. Refused says why nothing changed.
/// </summary>
public sealed record EngineExtractVariable(
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("source")] string? Source,
    [property: JsonPropertyName("variable")] string? Variable,
    /// <summary>The declared type, from the analyzer rather than from a guess.</summary>
    [property: JsonPropertyName("type")] string? Type,
    /// <summary>Whether its assignment needed Set, which is the other half of that answer.</summary>
    [property: JsonPropertyName("isObject")] bool? IsObject,
    [property: JsonPropertyName("expression")] string? Expression,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The result of an Inline Variable: the one module it rewrites, whole, with what the name stood
/// for and how many uses took its place. Refused says why nothing changed.
/// </summary>
public sealed record EngineInlineVariable(
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("source")] string? Source,
    [property: JsonPropertyName("variable")] string? Variable,
    [property: JsonPropertyName("value")] string? Value,
    [property: JsonPropertyName("replaced")] int? Replaced,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The result of a Move to Module: every module it rewrites, whole. Refused says why nothing moved.
/// </summary>
public sealed record EngineMoveToModule(
    [property: JsonPropertyName("modules")] EngineRenamedModule[]? Modules,
    [property: JsonPropertyName("moved")] string? Moved,
    [property: JsonPropertyName("from")] string? From,
    [property: JsonPropertyName("to")] string? To,
    /// <summary>How many qualified call sites were repointed at the new module.</summary>
    [property: JsonPropertyName("requalified")] int? Requalified,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>One place in a workbook: a module, and a 1-based line and column into its live text.</summary>
public sealed record EngineLocation(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column,
    [property: JsonPropertyName("length")] int Length,
    /// <summary>The line it sits on, so a module with no tab open can still be listed.</summary>
    [property: JsonPropertyName("preview")] string? Preview = null,
    /// <summary>
    /// What the occurrence does - read, write, or readwrite - classified by the analyzer from
    /// the statement it sits in (xlide_vscode#55). Null where nothing classifies: definitions,
    /// and a type name's occurrences.
    /// </summary>
    [property: JsonPropertyName("kind")] string? Kind = null);

/// <summary>Where a symbol is declared, or everywhere it is used. Never crosses a workbook.</summary>
public sealed record EngineLocations(
    [property: JsonPropertyName("locations")] EngineLocation[] Locations);

/// <summary>
/// One coloured span. The type is the analyzer's vocabulary - class, enum, struct, type,
/// variable - and the only modifier used is defaultLibrary, which marks a host global.
/// </summary>
public sealed record EngineSemanticToken(
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("modifiers")] string[]? Modifiers);

/// <summary>A module's colouring, in position order.</summary>
public sealed record EngineSemanticTokens(
    [property: JsonPropertyName("tokens")] EngineSemanticToken[] Tokens);

/// <summary>One procedure: the kind as a tree spells it ("Sub", "Property Get") and its 1-based line.</summary>
public sealed record EngineOutlineProcedure(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("line")] int Line);

/// <summary>A module's procedures, in declaration order.</summary>
public sealed record EngineOutline(
    [property: JsonPropertyName("procedures")] EngineOutlineProcedure[] Procedures);

/// <summary>One search hit: where it is, and the line it sits on for the results list.</summary>
public sealed record EngineSearchMatch(
    [property: JsonPropertyName("projectId")] string ProjectId,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column,
    [property: JsonPropertyName("length")] int Length,
    [property: JsonPropertyName("preview")] string Preview);

/// <summary>What a workspace search found; truncated says the limit spoke, not the text.</summary>
public sealed record EngineSearchResult(
    [property: JsonPropertyName("matches")] EngineSearchMatch[] Matches,
    [property: JsonPropertyName("truncated")] bool Truncated);

/// <summary>
/// What opening a project taught the engine: how many modules, and the project's own words -
/// names that denote types and names that denote procedures, for the surface's tokenizer.
/// </summary>
public sealed record EngineProjectOpened(
    [property: JsonPropertyName("modules")] int Modules,
    [property: JsonPropertyName("types")] string[] Types,
    [property: JsonPropertyName("procedures")] string[] Procedures);

/// <summary>
/// Serialisation for the engine protocol.
///
/// Every type is registered on a source-generated context because the add-in is compiled ahead of
/// time, where reflection-based serialisation cannot be used at all. Adding a type to the protocol
/// means adding it here, and forgetting to shows up as a run-time failure rather than a build one.
/// </summary>
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(EngineModule))]
[JsonSerializable(typeof(EngineModule[]))]
[JsonSerializable(typeof(EngineImplicitMember))]
[JsonSerializable(typeof(EngineImplicitMember[]))]
[JsonSerializable(typeof(EngineDiagnostic))]
[JsonSerializable(typeof(EngineDiagnostics))]
[JsonSerializable(typeof(EngineAnalysisRule))]
[JsonSerializable(typeof(EngineAnalysisRules))]
// Severity overrides ride the request dictionary boxed, like the booleans above.
[JsonSerializable(typeof(IReadOnlyDictionary<string, string>))]
[JsonSerializable(typeof(Dictionary<string, string>))]
[JsonSerializable(typeof(EngineSpan))]
[JsonSerializable(typeof(EnginePosition))]
[JsonSerializable(typeof(EngineCompletionItem))]
[JsonSerializable(typeof(EngineCompletions))]
[JsonSerializable(typeof(EngineHoverPayload))]
[JsonSerializable(typeof(EngineHover))]
[JsonSerializable(typeof(EngineSignatureParameter))]
[JsonSerializable(typeof(EngineSignatureInfo))]
[JsonSerializable(typeof(EngineSignatureHelp))]
[JsonSerializable(typeof(EngineTextEdit))]
[JsonSerializable(typeof(EngineTextEdit[]))]
[JsonSerializable(typeof(EngineSmartEnter))]
[JsonSerializable(typeof(EngineTextEdits))]
[JsonSerializable(typeof(EngineCodeAction))]
[JsonSerializable(typeof(EngineCodeActions))]
[JsonSerializable(typeof(EngineSemanticToken))]
[JsonSerializable(typeof(EngineSemanticTokens))]
[JsonSerializable(typeof(EngineRenamedModule))]
[JsonSerializable(typeof(EngineRename))]
[JsonSerializable(typeof(EngineExtractMethod))]
[JsonSerializable(typeof(EngineImplementInterface))]
[JsonSerializable(typeof(EngineEncapsulateField))]
[JsonSerializable(typeof(EngineExtractVariable))]
[JsonSerializable(typeof(EngineInlineVariable))]
[JsonSerializable(typeof(EngineMoveToModule))]
[JsonSerializable(typeof(EngineLocation))]
[JsonSerializable(typeof(EngineLocations))]
[JsonSerializable(typeof(EngineOutlineProcedure))]
[JsonSerializable(typeof(EngineOutline))]
[JsonSerializable(typeof(EngineProjectOpened))]
[JsonSerializable(typeof(JsonElement))]
// Booleans ride the request dictionaries boxed, and a boxed value serialises only if its own
// type is registered; leaving it out fails at run time, in the middle of a keystroke.
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(Dictionary<string, object>))]
// The live modules an import/export plan is worked out from travel as a list of these. Ahead-of-time
// serialisation needs the SHAPE named, and a nested one is not covered by naming the outer: without
// this line the call throws NoMetadataForType at run time and the planner quietly falls back, which
// is a green test against the wrong implementation (2026-08-09).
[JsonSerializable(typeof(List<Dictionary<string, object>>))]
[JsonSerializable(typeof(EngineSearchResult))]
[JsonSerializable(typeof(EngineSearchMatch))]
public sealed partial class EngineJsonContext : JsonSerializerContext;

/// <summary>
/// Converts between the editor's positions and the engine's.
///
/// The editor counts lines from one and gives no column at all for a whole line; the engine works
/// in character offsets into the module text. Converting in one place keeps the arithmetic from
/// being repeated, and repeated slightly differently, at every call site.
/// </summary>
public static class TextPositions
{
    /// <summary>
    /// Offsets at which each line begins. Built once per module version and reused, because
    /// converting a diagnostic without it costs a scan of the whole module per diagnostic.
    /// </summary>
    public static int[] LineStarts(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        var starts = new List<int>(Math.Max(16, text.Length / 40)) { 0 };

        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == '\n')
            {
                starts.Add(i + 1);
            }
        }

        return [.. starts];
    }

    /// <summary>
    /// Converts a character offset to a one-based line and column, which is what both the editor
    /// and the editing surface use.
    /// </summary>
    public static (int Line, int Column) ToLineColumn(int[] lineStarts, int offset)
    {
        ArgumentNullException.ThrowIfNull(lineStarts);

        if (lineStarts.Length == 0)
        {
            return (1, 1);
        }

        var low = 0;
        var high = lineStarts.Length - 1;

        while (low < high)
        {
            var middle = (low + high + 1) / 2;
            if (lineStarts[middle] <= offset)
            {
                low = middle;
            }
            else
            {
                high = middle - 1;
            }
        }

        return (low + 1, offset - lineStarts[low] + 1);
    }

    /// <summary>Converts a one-based line and column back to a character offset.</summary>
    public static int ToOffset(int[] lineStarts, int line, int column)
    {
        ArgumentNullException.ThrowIfNull(lineStarts);

        if (lineStarts.Length == 0)
        {
            return 0;
        }

        var index = Math.Clamp(line - 1, 0, lineStarts.Length - 1);
        return lineStarts[index] + Math.Max(0, column - 1);
    }
}
