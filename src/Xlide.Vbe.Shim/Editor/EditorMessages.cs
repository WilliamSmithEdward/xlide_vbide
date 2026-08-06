using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>Tells the surface which module to show and what its text is.</summary>
public sealed record LoadDocumentMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("text")] string Text);

/// <summary>
/// Tells the surface there is nothing to show: every pane is closed. The surface stays on screen
/// with its empty workspace rather than yielding the frame back to the native editor.
/// </summary>
public sealed record ClearDocumentMessage(
    [property: JsonPropertyName("type")] string Type);

/// <summary>One squiggle. Positions are one-based lines and columns, as the surface expects.</summary>
public sealed record EditorMarker(
    [property: JsonPropertyName("startLine")] int StartLine,
    [property: JsonPropertyName("startColumn")] int StartColumn,
    [property: JsonPropertyName("endLine")] int EndLine,
    [property: JsonPropertyName("endColumn")] int EndColumn,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("code")] string? Code);

/// <summary>Replaces every squiggle shown on the module.</summary>
public sealed record SetDiagnosticsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("markers")] EditorMarker[] Markers);

/// <summary>Chooses the surface's theme.</summary>
public sealed record SetThemeMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("theme")] string Theme);

/// <summary>Scrolls a line into view without moving the caret.</summary>
public sealed record RevealLineMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("line")] int Line);

/// <summary>
/// Puts the caret somewhere and reveals it. Distinct from revealLine, which only scrolls:
/// the caret decides what the editor's own commands act on, and the host syncs it into the
/// native pane before every one of them, so nothing that wants to run a particular procedure
/// can get there by scrolling alone.
/// </summary>
public sealed record SetCaretMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column);

/// <summary>
/// The modules the editor has open, and which one is showing. Projects runs parallel to
/// Modules — the workbook each tab belongs to, by the name the tree uses — so the strip can
/// say WHICH Module1 when two workbooks hold one.
/// </summary>
public sealed record SetModulesMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("modules")] string[] Modules,
    [property: JsonPropertyName("projects")] string?[] Projects,
    [property: JsonPropertyName("active")] string? Active,
    [property: JsonPropertyName("activeProject")] string? ActiveProject,
    [property: JsonPropertyName("dirty")] bool[]? Dirty = null);

/// <summary>The developer's settings, for the page's dialog and its typing behaviour.</summary>
public sealed record SetSettingsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("blockLayout")] string BlockLayout,
    [property: JsonPropertyName("continueCommentOnNewline")] bool ContinueCommentOnNewline,
    [property: JsonPropertyName("mirrorCommentSpacing")] bool MirrorCommentSpacing,
    [property: JsonPropertyName("formatIndentSize")] int FormatIndentSize,
    [property: JsonPropertyName("formatUseTabs")] bool FormatUseTabs,
    [property: JsonPropertyName("formatCanonicalKeywords")] bool FormatCanonicalKeywords);

/// <summary>One search hit, as the results list draws it. Workbook is the display name.</summary>
public sealed record SurfaceSearchMatch(
    [property: JsonPropertyName("workbook")] string? Workbook,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column,
    [property: JsonPropertyName("length")] int Length,
    [property: JsonPropertyName("preview")] string Preview);

/// <summary>A search's answer, echoing the id the page asked with.</summary>
public sealed record SearchResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("matches")] SurfaceSearchMatch[] Matches,
    [property: JsonPropertyName("truncated")] bool Truncated,
    [property: JsonPropertyName("replaced")] int Replaced = 0);

/// <summary>A breakpoint the host refused, so the page can mark the line briefly.</summary>
public sealed record BreakpointRefusedMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("line")] int Line);

/// <summary>One row of the Watch panel: a watch expression and where it stands.</summary>
public sealed record SurfaceWatchRow(
    [property: JsonPropertyName("expression")] string Expression,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("context")] string Context);

/// <summary>
/// One library the Object Browser lists: a referenced type library, or an open project —
/// the kind says which, because only a project's members can be navigated to.
/// </summary>
public sealed record ObLibraryRow(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("description")] string Description,
    [property: JsonPropertyName("kind")] string Kind);

/// <summary>One browsable type of a library.</summary>
public sealed record ObTypeRow(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind);

/// <summary>
/// One member of a type: its signature spelled the way VBA would. The line is where the
/// member lives in its module — meaningful only for project members, zero elsewhere.
/// </summary>
public sealed record ObMemberRow(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("signature")] string Signature,
    [property: JsonPropertyName("description")] string Description,
    [property: JsonPropertyName("line")] int Line);

/// <summary>The referenced libraries, answering an Object Browser request.</summary>
public sealed record ObLibrariesResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("libraries")] ObLibraryRow[] Libraries);

/// <summary>A library's types, answering an Object Browser request.</summary>
public sealed record ObTypesResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("types")] ObTypeRow[] Types);

/// <summary>A type's members, answering an Object Browser request.</summary>
public sealed record ObMembersResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("members")] ObMemberRow[] Members);

/// <summary>Which debug mode the editor is in: design, run, or break.</summary>
public sealed record SetDebugStateMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("mode")] string Mode);

/// <summary>The Watch panel's rows, replaced whole; stopped false is the idle state.</summary>
public sealed record SetWatchesMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("stopped")] bool Stopped,
    [property: JsonPropertyName("rows")] SurfaceWatchRow[] Rows);

/// <summary>
/// A tab close the host is holding until the developer answers for the module's unsaved
/// changes. The page asks again with their choice.
/// </summary>
public sealed record ConfirmCloseMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("project")] string? Project);

/// <summary>One row of the Locals panel: a variable in scope at the break.</summary>
public sealed record SurfaceLocalRow(
    [property: JsonPropertyName("expression")] string Expression,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("kind")] string Kind);

/// <summary>
/// What the debugger has in scope. Context names the broken procedure. Stopped false is the
/// idle state; stopped true with no rows is a break whose variables cannot be read yet, or a
/// scope with nothing in it — the panel must not claim "not stopped" during a break.
/// </summary>
public sealed record SetLocalsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("stopped")] bool Stopped,
    [property: JsonPropertyName("context")] string? Context,
    [property: JsonPropertyName("rows")] SurfaceLocalRow[] Rows);

/// <summary>One finding as the surface's panel wants it.</summary>
public sealed record SurfaceFinding(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("severity")] string Severity,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column,
    [property: JsonPropertyName("project")] string? Project = null);

/// <summary>Everything the panel lists, for every module.</summary>
public sealed record SetFindingsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("findings")] SurfaceFinding[] Findings);

/// <summary>One component in a project, with the kind the editor reports for it.</summary>
public sealed record SurfaceComponent(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] int Kind);

/// <summary>One project and everything in it.</summary>
public sealed record SurfaceProject(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("components")] SurfaceComponent[] Components);

/// <summary>The whole project tree, for the explorer.</summary>
public sealed record SetProjectsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("projects")] SurfaceProject[] Projects);

/// <summary>
/// The module's text as the editor now holds it, to be adopted without disturbing the developer.
///
/// Distinct from loading a document: loading replaces the model and resets the undo stack and the
/// caret, which is right when the developer opens a different module and wrong when the text they
/// are in the middle of typing has merely been normalised underneath them.
/// </summary>
public sealed record SyncDocumentMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("text")] string Text);

/// <summary>Asks the surface to run one of the editor's own commands.</summary>
public sealed record EditorCommandMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] string Id);

/// <summary>Something the developer should be told, shown briefly and not dwelt on.</summary>
public sealed record NoticeMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("text")] string Text);

/// <summary>One line of output for the Immediate panel.</summary>
public sealed record ImmediateResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("failed")] bool Failed);

/// <summary>The line execution is stopped on, or null when nothing is stopped.</summary>
public sealed record SetCurrentLineMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("line")] int? Line);

/// <summary>Every line in the shown module that carries a breakpoint.</summary>
public sealed record SetBreakpointsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("lines")] int[] Lines);

/// <summary>
/// One entry in a menu. The index is the item's real position in the editor's own control
/// collection, which is how the page addresses it back; hidden items are skipped but positions are
/// not renumbered around them.
/// </summary>
public sealed record SurfaceMenuItem(
    [property: JsonPropertyName("index")] int Index,
    [property: JsonPropertyName("caption")] string Caption,
    [property: JsonPropertyName("enabled")] bool Enabled,
    [property: JsonPropertyName("separator")] bool Separator,
    [property: JsonPropertyName("popup")] bool Popup,
    [property: JsonPropertyName("checked")] bool Checked,
    [property: JsonPropertyName("shortcut")] string? Shortcut);

/// <summary>The items of one menu, named by the path the page asked about.</summary>
public sealed record SetMenuMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("path")] int[] Path,
    [property: JsonPropertyName("items")] SurfaceMenuItem[] Items);

/// <summary>
/// Which parts of the surface's own chrome are drawn. The menu bar is withdrawn while the surface
/// retreats to the document area, because the native bar is visible then and two menu bars answer
/// the same question twice.
/// </summary>
public sealed record SetChromeMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("menuBar")] bool MenuBar);

/// <summary>
/// One property of the selected component, rendered for display. Writable says whether an edit
/// will be attempted, not promised: the editor can still refuse one, and the refusal is reported.
/// Boolean marks a value that offers True and False rather than free text.
/// </summary>
public sealed record SurfacePropertyEntry(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("writable")] bool Writable,
    [property: JsonPropertyName("boolean")] bool Boolean);

/// <summary>
/// The properties of the selected component, with the class name shown in the panel's object
/// header the way the editor's own window names what is selected.
/// </summary>
public sealed record SetPropertiesMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("component")] string Component,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("properties")] SurfacePropertyEntry[] Properties);

/// <summary>One completion offered to the editor. The kind is the analyzer's vocabulary.</summary>
public sealed record SurfaceCompletionItem(
    [property: JsonPropertyName("label")] string Label,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("detail")] string? Detail,
    [property: JsonPropertyName("documentation")] string? Documentation,
    [property: JsonPropertyName("insertText")] string? InsertText,
    [property: JsonPropertyName("filterText")] string? FilterText,
    [property: JsonPropertyName("sortText")] string? SortText);

/// <summary>The answer to one completion request, matched to it by its identifier.</summary>
public sealed record CompletionResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("items")] SurfaceCompletionItem[] Items);

/// <summary>A resolved hover: declaration line, plain-text facts, spans into the live source.</summary>
public sealed record SurfaceHoverPayload(
    [property: JsonPropertyName("signature")] string Signature,
    [property: JsonPropertyName("details")] string[] Details,
    [property: JsonPropertyName("documentation")] string? Documentation,
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End);

/// <summary>The answer to one hover request; a null hover means nothing under the cursor.</summary>
public sealed record HoverResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("hover")] SurfaceHoverPayload? Hover);

/// <summary>One parameter slot, its label exactly as it appears in the signature line.</summary>
public sealed record SurfaceSignatureParameter(
    [property: JsonPropertyName("label")] string Label,
    [property: JsonPropertyName("documentation")] string? Documentation);

/// <summary>A resolved call tip: the signature line and which parameter is active.</summary>
public sealed record SurfaceSignatureInfo(
    [property: JsonPropertyName("label")] string Label,
    [property: JsonPropertyName("parameters")] SurfaceSignatureParameter[] Parameters,
    [property: JsonPropertyName("activeParameter")] int ActiveParameter,
    [property: JsonPropertyName("documentation")] string? Documentation,
    [property: JsonPropertyName("details")] string[]? Details);

/// <summary>The answer to one call-tip request; null means the caret is not inside a call.</summary>
public sealed record SignatureHelpResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("signature")] SurfaceSignatureInfo? Signature);

/// <summary>A text replacement, offsets into the live source; an insertion has Start == End.</summary>
public sealed record SurfaceTextEdit(
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End,
    [property: JsonPropertyName("text")] string Text);

/// <summary>The answer to one Smart Enter request: edits, and the caret once they apply.</summary>
public sealed record SmartEnterResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("edits")] SurfaceTextEdit[] Edits,
    [property: JsonPropertyName("caret")] int? Caret);

/// <summary>The answer to one canonical-case request; no edits means the span was canonical.</summary>
public sealed record CanonicalCaseResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("edits")] SurfaceTextEdit[] Edits);

/// <summary>The answer to one loop-sync request; at most one edit, the paired rename.</summary>
public sealed record LoopSyncResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("edits")] SurfaceTextEdit[] Edits);

/// <summary>One procedure in a module's outline, the kind spelled the way the tree shows it.</summary>
public sealed record SurfaceOutlineProcedure(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("line")] int Line);

/// <summary>
/// The answer to one outline request: the module's procedures in declaration order. Failed
/// marks an answer that is a shrug rather than a statement — the engine timed out or threw —
/// so the page keeps what it already shows instead of blanking an unfolded list.
/// </summary>
public sealed record OutlineResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("procedures")] SurfaceOutlineProcedure[] Procedures,
    [property: JsonPropertyName("failed")] bool Failed = false);

/// <summary>
/// The project's own words for the tokenizer: names that denote types and names that denote
/// procedures, so a name reads as what it is wherever it appears.
/// </summary>
public sealed record SetLanguageFactsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("types")] string[] Types,
    [property: JsonPropertyName("procedures")] string[] Procedures);

/// <summary>
/// Serialisation for surface messages. Source generated, because ahead-of-time compilation has no
/// reflection to fall back on and a message type that is not registered here fails at run time
/// rather than at build time.
/// </summary>
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(LoadDocumentMessage))]
[JsonSerializable(typeof(ClearDocumentMessage))]
[JsonSerializable(typeof(SetDiagnosticsMessage))]
[JsonSerializable(typeof(SetThemeMessage))]
[JsonSerializable(typeof(RevealLineMessage))]
[JsonSerializable(typeof(SetCaretMessage))]
[JsonSerializable(typeof(SetModulesMessage))]
[JsonSerializable(typeof(SetSettingsMessage))]
[JsonSerializable(typeof(SetLocalsMessage))]
[JsonSerializable(typeof(SurfaceLocalRow))]
[JsonSerializable(typeof(BreakpointRefusedMessage))]
[JsonSerializable(typeof(ConfirmCloseMessage))]
[JsonSerializable(typeof(SetWatchesMessage))]
[JsonSerializable(typeof(SurfaceWatchRow))]
[JsonSerializable(typeof(SetDebugStateMessage))]
[JsonSerializable(typeof(ObLibraryRow))]
[JsonSerializable(typeof(ObTypeRow))]
[JsonSerializable(typeof(ObMemberRow))]
[JsonSerializable(typeof(ObLibrariesResultMessage))]
[JsonSerializable(typeof(ObTypesResultMessage))]
[JsonSerializable(typeof(ObMembersResultMessage))]
[JsonSerializable(typeof(SearchResultMessage))]
[JsonSerializable(typeof(SurfaceSearchMatch))]
[JsonSerializable(typeof(SetFindingsMessage))]
[JsonSerializable(typeof(SetProjectsMessage))]
[JsonSerializable(typeof(SyncDocumentMessage))]
[JsonSerializable(typeof(EditorCommandMessage))]
[JsonSerializable(typeof(NoticeMessage))]
[JsonSerializable(typeof(ImmediateResultMessage))]
[JsonSerializable(typeof(SetCurrentLineMessage))]
[JsonSerializable(typeof(SetBreakpointsMessage))]
[JsonSerializable(typeof(SetMenuMessage))]
[JsonSerializable(typeof(SetChromeMessage))]
[JsonSerializable(typeof(SurfaceMenuItem))]
[JsonSerializable(typeof(SetPropertiesMessage))]
[JsonSerializable(typeof(SurfacePropertyEntry))]
[JsonSerializable(typeof(CompletionResultMessage))]
[JsonSerializable(typeof(SurfaceCompletionItem))]
[JsonSerializable(typeof(HoverResultMessage))]
[JsonSerializable(typeof(SurfaceHoverPayload))]
[JsonSerializable(typeof(SignatureHelpResultMessage))]
[JsonSerializable(typeof(SurfaceSignatureInfo))]
[JsonSerializable(typeof(SurfaceSignatureParameter))]
[JsonSerializable(typeof(SurfaceTextEdit))]
[JsonSerializable(typeof(SmartEnterResultMessage))]
[JsonSerializable(typeof(CanonicalCaseResultMessage))]
[JsonSerializable(typeof(LoopSyncResultMessage))]
[JsonSerializable(typeof(SurfaceOutlineProcedure))]
[JsonSerializable(typeof(OutlineResultMessage))]
[JsonSerializable(typeof(SetLanguageFactsMessage))]
[JsonSerializable(typeof(SurfaceProject))]
[JsonSerializable(typeof(SurfaceComponent))]
[JsonSerializable(typeof(SurfaceFinding))]
[JsonSerializable(typeof(EditorMarker))]
public sealed partial class EditorMessageContext : JsonSerializerContext;
