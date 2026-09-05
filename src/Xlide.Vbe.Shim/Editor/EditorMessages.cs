using System.Text.Json.Serialization;

namespace Xlide.Vbe.Shim.Editor;

/// <summary>
/// Tells the surface a module is open and what its text is. The page keeps a live model per
/// open module (decision 12); the project is the workbook display name that, with the module
/// name, is the document's identity. Idempotent on the page: a model that already exists
/// adopts the text instead of being replaced.
/// </summary>
public sealed record OpenDocumentMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("text")] string Text);

/// <summary>
/// A form's design, answering the page's requestFormMarkup: the markup TEXT and the walked
/// SPEC it was printed from, in one message so the designer tab's two halves - the document
/// and the visual - are the same walk and cannot skew. `markup` null means the form could not
/// be projected and `reason` says why - the tab shows the reason rather than an empty
/// document pretending to be a form.
/// </summary>
public sealed record FormMarkupMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("markup")] string? Markup,
    [property: JsonPropertyName("reason")] string? Reason,
    [property: JsonPropertyName("form")] FormMarkupBox? Form = null,
    [property: JsonPropertyName("controls")] FormMarkupControl[]? Controls = null);

/// <summary>
/// How an apply of the markup tab's document ended: what landed, and why it stopped if it
/// did. A fresh formMarkup follows it either way, because what LANDED is on the form even
/// when the apply stopped partway - the tab shows the truth, the refusal explains the gap.
/// </summary>
public sealed record FormMarkupAppliedMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("ok")] bool Ok,
    [property: JsonPropertyName("added")] string[] Added,
    [property: JsonPropertyName("removed")] string[] Removed,
    [property: JsonPropertyName("set")] int Set,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The lint of a markup document, answering the page's lintFormMarkup: every finding the
/// tolerant parse collects, where the strict apply stops at the first. Pure text both ways -
/// no designer is touched, so no window can stir.
/// </summary>
public sealed record FormMarkupLintMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("findings")] FormMarkupLintFinding[] Findings,
    // The DRAFT the text describes, when it parses strictly: the same ride the squiggles
    // take, so the canvas can follow the document as it is typed without a second parser
    // or a second round trip - and without the form being touched. Dialect fields only;
    // the page carries display extras over from the last applied projection by name.
    [property: JsonPropertyName("draftForm")] FormMarkupBox? DraftForm = null,
    [property: JsonPropertyName("draft")] FormMarkupControl[]? Draft = null);

/// <summary>
/// The host's Ctrl+S with a designer tab active: the page applies the tab's document to
/// the form and calls back for the raw save ("saveOnly"). The document lives in the page,
/// so the host cannot flush it itself the way it flushes code edits before a save.
///
/// `Run` is the same handshake asked for by F5, and the callback it comes back through is
/// "runOnly" - the launch with no save, because the native editor never saves on Run and the
/// guarantee F5 needs, that the window which opens is the document, is the apply itself. The
/// intent travels WITH the request rather than being remembered here, because a refused apply
/// calls nothing back at all: a flag waiting on the host for a callback that never comes
/// would fire on somebody else's Ctrl+S minutes later.
/// </summary>
public sealed record DesignerApplySaveMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("run")] bool Run = false);

/// <summary>
/// The markup language's whole vocabulary, answering the page's requestFormMarkupVocabulary:
/// every kind a document can spell and every property each kind holds, with what it takes, what
/// it holds untouched, and what its type library says about it.
///
/// MEASURED, never written down: each kind's is read from a bare instance of its coclass and its
/// own ITypeInfo (ControlDefaults), so the completions offer what MSForms on THIS machine has
/// rather than what a table of ours remembers. Sent once per session - a registered coclass does
/// not change while Excel is up - and the page's providers answer from it without a round trip.
/// </summary>
public sealed record FormMarkupVocabularyMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("kinds")] FormMarkupKind[] Kinds);

/// <summary>One kind a document can name: its ProgID, whether it holds children, and its
/// properties. `Form` is here too, described from the live form rather than from a coclass,
/// because a document always has exactly one and its lines are the first a developer types.</summary>
public sealed record FormMarkupKind(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("progId")] string? ProgId,
    [property: JsonPropertyName("container")] bool Container,
    [property: JsonPropertyName("properties")] FormMarkupProperty[] Properties,
    /// <summary>What the kind IS, in a line - the sentence a hover leads with. The wording is
    /// this product's, because MSForms ships no help strings to read it out of.</summary>
    [property: JsonPropertyName("doc")] string? Doc = null);

/// <summary>One property line a document may carry: the path, the type it is declared as, the
/// value an untouched control holds, the library's own sentence about it, and its enum's members
/// where it has them. A colour says so, because the dialect spells one as &amp;Hbbggrr&amp;.</summary>
public sealed record FormMarkupProperty(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("type")] string? Type,
    [property: JsonPropertyName("default")] string? Default,
    [property: JsonPropertyName("doc")] string? Doc,
    [property: JsonPropertyName("members")] FormMarkupEnumMember[]? Members,
    [property: JsonPropertyName("colour")] bool Colour,
    /// <summary>Values that are not an enum's members but are still worth offering, because the
    /// machine can be asked what they are: a font's faces, today, which is the whole list. The
    /// property still takes anything - MSForms stores a face as a plain string.</summary>
    [property: JsonPropertyName("values")] string[]? Values = null);

/// <summary>One member of a property's enum: the name the developer writes, and the number
/// behind it - both, because the dialect takes either.</summary>
public sealed record FormMarkupEnumMember(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("value")] int Value);

/// <summary>One squiggle: 1-based line, the reason, and "error" or "warning".</summary>
public sealed record FormMarkupLintFinding(
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("severity")] string Severity);

/// <summary>The form's own box, as the markup layer projects it: points, like every bound.
/// Colours arrive as CSS, converted host-side through the system palette, so the canvas
/// paints what this machine's real form surface would; insides are the designer's own
/// client area, the parity numbers the canvas derives its chrome from.</summary>
public sealed record FormMarkupBox(
    [property: JsonPropertyName("caption")] string? Caption,
    [property: JsonPropertyName("width")] double? Width,
    [property: JsonPropertyName("height")] double? Height,
    [property: JsonPropertyName("backColor")] string? BackColor = null,
    [property: JsonPropertyName("foreColor")] string? ForeColor = null,
    [property: JsonPropertyName("insideWidth")] double? InsideWidth = null,
    [property: JsonPropertyName("insideHeight")] double? InsideHeight = null,
    [property: JsonPropertyName("picture")] FormMarkupPicture? Picture = null);

/// <summary>
/// A picture and how it sits, for the canvas to paint. Display truth of the purest kind: the
/// dialect cannot speak a picture - it is binary in the form's .frx, and MSForms does not
/// remember the file it came from - so it rides the projection or it is not drawn at all.
///
/// The two placements are the two families MSForms has, and a control has one or the other. A
/// SURFACE picture (the form, an Image, a Frame, a Page) is placed by size mode, alignment and
/// tiling; a CAPTION picture (a button, a Label, a check box) by its position around the caption.
/// </summary>
public sealed record FormMarkupPicture(
    [property: JsonPropertyName("src")] string Source,
    [property: JsonPropertyName("sizeMode")] int? SizeMode = null,
    [property: JsonPropertyName("alignment")] int? Alignment = null,
    [property: JsonPropertyName("tiling")] bool? Tiling = null,
    [property: JsonPropertyName("position")] int? Position = null);

/// <summary>
/// One control of the projection, flat with a parent NAME - the walk's own shape. Bounds are
/// points relative to the parent's client area, which is MSForms' own coordinate model; the
/// canvas composes them by nesting rather than by arithmetic. The display fields beyond the
/// markup's vocabulary - font, colours, a container's real client area - ride here for the
/// canvas's parity with the real form surface; the DOCUMENT deliberately does not speak them.
/// </summary>
public sealed record FormMarkupControl(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("caption")] string? Caption,
    [property: JsonPropertyName("left")] double? Left,
    [property: JsonPropertyName("top")] double? Top,
    [property: JsonPropertyName("width")] double? Width,
    [property: JsonPropertyName("height")] double? Height,
    [property: JsonPropertyName("parent")] string? Parent,
    [property: JsonPropertyName("fontName")] string? FontName = null,
    [property: JsonPropertyName("fontSize")] double? FontSize = null,
    [property: JsonPropertyName("fontBold")] bool? FontBold = null,
    [property: JsonPropertyName("fontItalic")] bool? FontItalic = null,
    [property: JsonPropertyName("backColor")] string? BackColor = null,
    [property: JsonPropertyName("foreColor")] string? ForeColor = null,
    [property: JsonPropertyName("insideWidth")] double? InsideWidth = null,
    [property: JsonPropertyName("insideHeight")] double? InsideHeight = null,
    [property: JsonPropertyName("tabs")] string[]? Tabs = null,
    /// <summary>Where the control sits in its container's TAB ORDER. Display truth, like the
    /// fonts and the client areas beside it: the dialect does not print a tab index, and the
    /// tab-order dialog reads it from the projection rather than walking the form again.</summary>
    [property: JsonPropertyName("tabIndex")] int? TabIndex = null,
    [property: JsonPropertyName("picture")] FormMarkupPicture? Picture = null);

/// <summary>
/// Tells the surface there is nothing to show: every pane is closed. The surface drops every
/// model and stays on screen with its empty workspace rather than yielding the frame back to
/// the native editor.
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

/// <summary>Replaces every squiggle shown on one open module's model.</summary>
public sealed record SetDiagnosticsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("moduleName")] string ModuleName,
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("markers")] EditorMarker[] Markers);

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
/// Modules - the workbook each tab belongs to, by the name the tree uses - so the strip can
/// say WHICH Module1 when two workbooks hold one. Faces runs parallel too: null or "code" is
/// a code pane mirrored from the host's own pane list, "design" is a form's designer tab,
/// which is THIS product's state rather than a mirror - the native designer window stays
/// down on purpose (the Toolbox trap), so these tabs exist only here and in the page.
/// </summary>
public sealed record SetModulesMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("modules")] string[] Modules,
    [property: JsonPropertyName("projects")] string?[] Projects,
    [property: JsonPropertyName("active")] string? Active,
    [property: JsonPropertyName("activeProject")] string? ActiveProject,
    [property: JsonPropertyName("dirty")] bool[]? Dirty = null,
    [property: JsonPropertyName("faces")] string?[]? Faces = null,
    [property: JsonPropertyName("activeFace")] string? ActiveFace = null);

/// <summary>The developer's settings, for the page's dialog and its typing behaviour.</summary>
public sealed record SetSettingsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("blockLayout")] string BlockLayout,
    [property: JsonPropertyName("continueCommentOnNewline")] bool ContinueCommentOnNewline,
    [property: JsonPropertyName("mirrorCommentSpacing")] bool MirrorCommentSpacing,
    [property: JsonPropertyName("insertOptionExplicit")] bool InsertOptionExplicit,
    [property: JsonPropertyName("treeFollowsEditor")] bool TreeFollowsEditor,
    [property: JsonPropertyName("formatIndentSize")] int FormatIndentSize,
    [property: JsonPropertyName("syncEngine")] string SyncEngine,
    [property: JsonPropertyName("designerSnap")] string DesignerSnap,
    [property: JsonPropertyName("designerGridSize")] int DesignerGridSize,
    /// <summary>Which explorer layout is showing: "tree" or "folders".</summary>
    [property: JsonPropertyName("explorerView")] string ExplorerView = "tree");

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

/// <summary>One row of the Watch panel: a watch expression and where it stands.</summary>
public sealed record SurfaceWatchRow(
    [property: JsonPropertyName("expression")] string Expression,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("context")] string Context);

/// <summary>
/// One library the Object Browser lists: a referenced type library, or an open project -
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
/// member lives in its module - meaningful only for project members, zero elsewhere.
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
/// scope with nothing in it - the panel must not claim "not stopped" during a break.
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

/// <summary>
/// One component in a project, with the kind the editor reports for it, and the folder its
/// <c>'@Folder("Parent.Child")</c> annotation names - null for a module that carries none, which
/// the folder view draws at the workbook's root.
/// </summary>
public sealed record SurfaceComponent(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] int Kind,
    [property: JsonPropertyName("folder")] string? Folder = null);

/// <summary>One project and everything in it.</summary>
public sealed record SurfaceProject(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("components")] SurfaceComponent[] Components);

/// <summary>
/// The whole project tree, for the explorer - and WHICH APPLICATION the tree belongs to, so the
/// page can shape what it offers: Access VBA has no UserForms, and a "New UserForm" item there
/// would promise what the host cannot do (the owner, 2026-08-19).
/// </summary>
public sealed record SetProjectsMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("projects")] SurfaceProject[] Projects,
    [property: JsonPropertyName("host")] string Host);

/// <summary>One test as the Tests pane draws it: identity, place, directive facts, and outcome.</summary>
public sealed record TestRowMessage(
    [property: JsonPropertyName("id")] string Id,
    /// <summary>
    /// The file the test lives in, as it is shown - a workbook in Excel, a document in Word.
    /// A module name is not an identity across files, so a row without this cannot be placed.
    /// </summary>
    [property: JsonPropertyName("file")] string File,
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("procedure")] string Procedure,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("status")] string Status,
    [property: JsonPropertyName("message")] string? Message,
    [property: JsonPropertyName("durationMs")] double DurationMs,
    [property: JsonPropertyName("output")] string[] Output,
    [property: JsonPropertyName("tags")] string[] Tags,
    [property: JsonPropertyName("owner")] string? Owner,
    [property: JsonPropertyName("requirement")] string? Requirement,
    [property: JsonPropertyName("timeoutMs")] int? TimeoutMs,
    [property: JsonPropertyName("expectedError")] string? ExpectedError);

/// <summary>
/// One open file's standing in the runner: what it is called, whether it carries a current
/// XlideAssert, and how many tests were discovered in it. The support module is installed per
/// file, because it is a module IN the file - so a session with two files open has two answers,
/// and a pane showing one number for both would be lying to whichever it did not mean.
/// </summary>
public sealed record TestFileMessage(
    [property: JsonPropertyName("file")] string File,
    [property: JsonPropertyName("support")] string Support,
    [property: JsonPropertyName("tests")] int Tests);

/// <summary>
/// The Tests pane's whole picture in one message: every open file's support state, whether a run
/// is in flight and where it stands, and every known test with its latest outcome. Sent whole on
/// every change - discovery, run start, each landing result - because a panel diffing partial
/// updates is a panel that drifts.
/// </summary>
public sealed record SetTestsMessage(
    [property: JsonPropertyName("type")] string Type,
    /// <summary>
    /// The whole session's support standing, worst first: missing if any file holding tests has
    /// no XlideAssert, outdated if any is behind, installed when every file that needs one has a
    /// current one. `files` carries the per-file truth the summary is drawn from.
    /// </summary>
    [property: JsonPropertyName("support")] string Support,
    [property: JsonPropertyName("running")] bool Running,
    [property: JsonPropertyName("currentTest")] string? CurrentTest,
    /// <summary>When the last run finished, round-trip ISO 8601 with offset; null before any run.</summary>
    [property: JsonPropertyName("ranAt")] string? RanAt,
    [property: JsonPropertyName("files")] TestFileMessage[] Files,
    [property: JsonPropertyName("rows")] TestRowMessage[] Rows);

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
    [property: JsonPropertyName("project")] string? Project,
    [property: JsonPropertyName("text")] string Text);

/// <summary>Asks the surface to run one of the editor's own commands.</summary>
public sealed record EditorCommandMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] string Id);

/// <summary>
/// Something the developer should be told, shown briefly and not dwelt on.
///
/// `sticky` is for the other kind: a condition that lasts an unknown length of time and ends when
/// something happens rather than when a timer expires. A five second notice cannot describe a wait
/// - it either vanishes while the wait continues or lingers after it ends. An empty text with
/// sticky set clears the line.
/// </summary>
public sealed record NoticeMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("sticky")] bool Sticky = false);

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
    [property: JsonPropertyName("shortcut")] string? Shortcut,
    /// <summary>
    /// A codicon name to draw INSTEAD of the caption, for the entries this product composes rather
    /// than mirrors. Null for everything read from the editor, which has captions and no icons.
    /// The caption is still carried when this is set: it becomes the accessible name, because an
    /// icon on its own is a button a screen reader can only call "button".
    /// </summary>
    [property: JsonPropertyName("icon")] string? Icon = null);

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
/// Where the add-in is loaded from, for the surface's About dialog. Only the host can answer it:
/// the page has no filesystem, and the question behind it is always the same one, which is whether
/// the build being looked at is the installed one or a development publish.
/// </summary>
public sealed record SetInstallPathMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("path")] string? Path);

/// <summary>
/// The system colours, for the colour picker's System half: what each one is called, the value
/// that ASKS for it, and what this machine answers today. Sent once when the surface loads, like
/// the install path - the names do not change and the values only change with the theme, which
/// the picker can afford to learn on the next reload.
///
/// Only the host can answer the third field, which is the whole reason this message exists: a
/// page has no way to ask Windows what a button face is, and a picker whose System list is grey
/// squares is a list of words.
/// </summary>
public sealed record SetSystemColoursMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("colours")] SystemColourEntry[] Colours);

/// <summary>One system colour: the name the panel shows, the value a property takes, and the
/// CSS this machine resolves it to right now.</summary>
public sealed record SystemColourEntry(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("css")] string Css);

/// <summary>
/// One property of the selected component, rendered for display. Writable says whether an edit
/// will be attempted, not promised: the editor can still refuse one, and the refusal is reported.
/// Boolean marks a value that offers True and False rather than free text.
/// </summary>
/// <summary>
/// One row of the Properties panel. `Options`, when the type library named them, are the values
/// this property can take - an enum's members, spelled the way the developer writes them - and
/// the row becomes a list rather than a text box. `Boolean` is the same idea from before the
/// typelib was read, kept because True/False needs no library to know.
/// </summary>
public sealed record SurfacePropertyEntry(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("writable")] bool Writable,
    [property: JsonPropertyName("boolean")] bool Boolean,
    [property: JsonPropertyName("options")] string[]? Options = null,
    /// <summary>A colour property's value as CSS, system colours resolved - what a swatch
    /// paints, and what the picker opens on. Null for everything that is not a colour.</summary>
    [property: JsonPropertyName("swatch")] string? Swatch = null,
    /// <summary>True for a PICTURE row, which is edited by choosing a file rather than by
    /// typing: the row offers Browse and Clear instead of a text box.</summary>
    [property: JsonPropertyName("picture")] bool Picture = false,
    /// <summary>The picture itself as a data URI, for the row's thumbnail. Null when the
    /// property holds nothing, or holds something this side cannot turn into pixels.</summary>
    [property: JsonPropertyName("preview")] string? Preview = null);

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

/// <summary>
/// One quick fix: what to call it, the finding it answers, and the edits that apply it. The code
/// and span are the finding's, so the surface can attach the fix to the squiggle it belongs to.
/// </summary>
public sealed record SurfaceCodeAction(
    [property: JsonPropertyName("title")] string Title,
    [property: JsonPropertyName("isPreferred")] bool IsPreferred,
    [property: JsonPropertyName("code")] string? Code,
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End,
    [property: JsonPropertyName("edits")] SurfaceTextEdit[] Edits);

/// <summary>
/// One analyzer rule for the rules modal: its stable code, its words, its default severity, and
/// the override values the analyzer permits. An empty `allowed` renders as fixed - most error
/// rules mirror a VBE compile failure and take no override at all.
/// </summary>
public sealed record SurfaceAnalysisRule(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("title")] string Title,
    [property: JsonPropertyName("category")] string Category,
    [property: JsonPropertyName("defaultSeverity")] string DefaultSeverity,
    [property: JsonPropertyName("allowed")] string[] Allowed);

/// <summary>
/// The rules modal's one payload: the catalog and the machine's standing overrides together,
/// because a list of rules with the ticks read from somewhere else is two requests that can
/// disagree. `failed` distinguishes "the engine is not up yet" from a catalog of none.
/// </summary>
public sealed record AnalysisRulesResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("rules")] SurfaceAnalysisRule[] Rules,
    [property: JsonPropertyName("overrides")] IReadOnlyDictionary<string, string> Overrides,
    [property: JsonPropertyName("failed")] bool Failed);

/// <summary>The answer to one quick-fix request; empty means nothing on that span can be fixed.</summary>
public sealed record CodeActionResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("actions")] SurfaceCodeAction[] Actions);

/// <summary>
/// The answer to one rename: which modules changed and how many uses went in each, or the reason
/// nothing changed. The new text is not sent back - the host has already written it, and the open
/// tabs are refreshed by the ordinary document sync.
/// </summary>
public sealed record RenameResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("oldName")] string? OldName,
    [property: JsonPropertyName("newName")] string? NewName,
    [property: JsonPropertyName("modules")] string[] Modules,
    [property: JsonPropertyName("replaced")] int Replaced,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Introduce Parameter: the name that is now a parameter, the value the callers
/// pass, and how many took it - or the reason nothing changed.
/// </summary>
public sealed record IntroduceParameterResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("parameter")] string? Parameter,
    [property: JsonPropertyName("declaredType")] string? DeclaredType,
    [property: JsonPropertyName("value")] string? Value,
    [property: JsonPropertyName("procedure")] string? Procedure,
    [property: JsonPropertyName("modules")] string[] Modules,
    [property: JsonPropertyName("callSites")] int CallSites,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Move to Module: what moved, where from and to, which modules were rewritten
/// and how many qualified call sites followed it - or the reason nothing moved.
/// </summary>
public sealed record MoveToModuleResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("moved")] string? Moved,
    [property: JsonPropertyName("from")] string? From,
    [property: JsonPropertyName("to")] string? To,
    [property: JsonPropertyName("modules")] string[] Modules,
    [property: JsonPropertyName("requalified")] int Requalified,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Inline Variable: the name that is gone, what stands in its place, and how
/// many uses took it - or the reason nothing changed.
/// </summary>
public sealed record InlineVariableResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("variable")] string? Variable,
    [property: JsonPropertyName("value")] string? Value,
    [property: JsonPropertyName("replaced")] int Replaced,
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Extract Variable: the name it made, the type the analyzer gave it, and
/// whether it needed Set - or the reason nothing was made.
/// </summary>
public sealed record ExtractVariableResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("variable")] string? Variable,
    [property: JsonPropertyName("declaredType")] string? DeclaredType,
    [property: JsonPropertyName("isObject")] bool IsObject,
    [property: JsonPropertyName("expression")] string? Expression,
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Encapsulate Field: what became a property, what the variable behind it is
/// called, and the two members written - or the reason nothing changed.
/// </summary>
public sealed record EncapsulateResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("field")] string? Field,
    [property: JsonPropertyName("backingField")] string? BackingField,
    [property: JsonPropertyName("accessors")] string[] Accessors,
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Implement Interface: whose members were written and which ones, or the reason
/// none were. The new text is not sent back, for the reason a rename's is not.
/// </summary>
public sealed record ImplementResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("interfaces")] string[] Interfaces,
    [property: JsonPropertyName("added")] string[] Added,
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// The answer to one Extract Method: what was made, out of what, and the header it was given - or
/// the reason nothing was made. Like a rename, the new text is not sent back: the host has already
/// written it and the open tab is refreshed by the ordinary document sync.
/// </summary>
public sealed record ExtractResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("procedure")] string? Procedure,
    [property: JsonPropertyName("from")] string? From,
    [property: JsonPropertyName("signature")] string? Signature,
    [property: JsonPropertyName("module")] string? Module,
    [property: JsonPropertyName("refused")] string? Refused);

/// <summary>
/// One place in the workbook: the module, its workbook's display name, and a 1-based line and
/// column into the module's live text.
/// </summary>
public sealed record SurfaceLocation(
    [property: JsonPropertyName("module")] string Module,
    [property: JsonPropertyName("workbook")] string? Workbook,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column,
    [property: JsonPropertyName("length")] int Length,
    /// <summary>The line it sits on, which is what makes an unopened module renderable.</summary>
    [property: JsonPropertyName("preview")] string? Preview = null,
    /// <summary>read | write | readwrite, from the analyzer; null where nothing classifies.</summary>
    [property: JsonPropertyName("kind")] string? Kind = null);

/// <summary>
/// The answer to one navigation request: where a symbol is declared, or everywhere it is used.
/// Always within the one workbook, because two workbooks can each hold a Module1.
/// </summary>
public sealed record NavigationResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("locations")] SurfaceLocation[] Locations);

/// <summary>
/// One coloured span, offsets into the live source. The type is the analyzer's vocabulary and
/// the only modifier used is defaultLibrary, which marks a host global.
/// </summary>
public sealed record SurfaceSemanticToken(
    [property: JsonPropertyName("start")] int Start,
    [property: JsonPropertyName("end")] int End,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("modifiers")] string[]? Modifiers);

/// <summary>
/// The answer to one colouring request, in position order. Failed says the module could not be
/// coloured rather than that it has no colour: the surface keeps what it is already showing,
/// because dropping to a plain grammar mid-session reads as the analysis having gone wrong.
/// </summary>
public sealed record SemanticTokensResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("tokens")] SurfaceSemanticToken[] Tokens,
    [property: JsonPropertyName("failed")] bool Failed);

/// <summary>One procedure in a module's outline, the kind spelled the way the tree shows it.</summary>
public sealed record SurfaceOutlineProcedure(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("line")] int Line);

/// <summary>
/// The answer to one outline request: the module's procedures in declaration order. Failed
/// marks an answer that is a shrug rather than a statement - the engine timed out or threw -
/// so the page keeps what it already shows instead of blanking an unfolded list.
/// </summary>
public sealed record OutlineResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("procedures")] SurfaceOutlineProcedure[] Procedures,
    [property: JsonPropertyName("failed")] bool Failed = false);

/// <summary>
/// What MSForms' own AutoSize makes a control, answered for "Size to Fit".
///
/// The size cannot be worked out on the page - a check box's glyph, a button's chrome and above
/// all a picture drawn at natural size are not in the caption's ink - so the gesture asks the
/// host, which measures and puts the control straight back. Null width and height are a control
/// with no AutoSize to ask about, which is an answer rather than a failure.
/// </summary>
public sealed record DesignerAutoSizeResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("width")] double? Width,
    [property: JsonPropertyName("height")] double? Height);

/// <summary>
/// The answer to an import/export request, carried as the JSON the service produced rather than as
/// a shape of its own.
///
/// Passing it through verbatim is deliberate: the xlide api answers the very same string from the
/// very same call, so the dialog cannot be looking at a different plan from the one a harness
/// reads. A shape declared twice is a shape that drifts once.
/// </summary>
public sealed record SyncResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("json")] string Json);

/// <summary>
/// The change log's answer to something the pane asked, carried as the route's own JSON so the
/// pane and the xlide api are reading one reply rather than two shapes that can disagree.
/// </summary>
public sealed record ChangesResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("json")] string Json);

/// <summary>
/// The api door's answer to the agent card: whether it is open, and the address it hands out.
/// Carried as JSON for the same reason the change log's is - one shape, read in one place.
/// </summary>
public sealed record ApiResultMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("json")] string Json);

/// <summary>
/// One number, meaning "the change log has moved on since you last read it". Carries nothing
/// else on purpose: the pane's counts are whole-text comparisons and are never computed on the
/// write path, so this is a tap on the shoulder rather than an update.
/// </summary>
public sealed record ChangesStampMessage(
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("stamp")] int Stamp);

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
[JsonSerializable(typeof(OpenDocumentMessage))]
[JsonSerializable(typeof(FormMarkupMessage))]
[JsonSerializable(typeof(FormMarkupBox))]
[JsonSerializable(typeof(FormMarkupControl))]
[JsonSerializable(typeof(FormMarkupAppliedMessage))]
[JsonSerializable(typeof(FormMarkupLintMessage))]
[JsonSerializable(typeof(DesignerApplySaveMessage))]
[JsonSerializable(typeof(FormMarkupLintFinding))]
[JsonSerializable(typeof(FormMarkupVocabularyMessage))]
[JsonSerializable(typeof(FormMarkupKind))]
[JsonSerializable(typeof(FormMarkupProperty))]
[JsonSerializable(typeof(FormMarkupEnumMember))]
[JsonSerializable(typeof(ClearDocumentMessage))]
[JsonSerializable(typeof(SetDiagnosticsMessage))]
[JsonSerializable(typeof(RevealLineMessage))]
[JsonSerializable(typeof(SetCaretMessage))]
[JsonSerializable(typeof(SetModulesMessage))]
[JsonSerializable(typeof(SetSettingsMessage))]
[JsonSerializable(typeof(SetLocalsMessage))]
[JsonSerializable(typeof(SurfaceLocalRow))]
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
[JsonSerializable(typeof(SetTestsMessage))]
[JsonSerializable(typeof(SyncDocumentMessage))]
[JsonSerializable(typeof(EditorCommandMessage))]
[JsonSerializable(typeof(NoticeMessage))]
[JsonSerializable(typeof(ImmediateResultMessage))]
[JsonSerializable(typeof(SetCurrentLineMessage))]
[JsonSerializable(typeof(SetBreakpointsMessage))]
[JsonSerializable(typeof(SetMenuMessage))]
[JsonSerializable(typeof(SetChromeMessage))]
[JsonSerializable(typeof(SetInstallPathMessage))]
[JsonSerializable(typeof(SetSystemColoursMessage))]
[JsonSerializable(typeof(SystemColourEntry))]
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
[JsonSerializable(typeof(ChangesResultMessage))]
[JsonSerializable(typeof(ApiResultMessage))]
[JsonSerializable(typeof(ChangesStampMessage))]
[JsonSerializable(typeof(SurfaceCodeAction))]
[JsonSerializable(typeof(CodeActionResultMessage))]
[JsonSerializable(typeof(SurfaceAnalysisRule))]
[JsonSerializable(typeof(AnalysisRulesResultMessage))]
[JsonSerializable(typeof(SurfaceSemanticToken))]
[JsonSerializable(typeof(SemanticTokensResultMessage))]
[JsonSerializable(typeof(RenameResultMessage))]
[JsonSerializable(typeof(ExtractResultMessage))]
[JsonSerializable(typeof(ImplementResultMessage))]
[JsonSerializable(typeof(EncapsulateResultMessage))]
[JsonSerializable(typeof(ExtractVariableResultMessage))]
[JsonSerializable(typeof(InlineVariableResultMessage))]
[JsonSerializable(typeof(MoveToModuleResultMessage))]
[JsonSerializable(typeof(IntroduceParameterResultMessage))]
[JsonSerializable(typeof(SurfaceLocation))]
[JsonSerializable(typeof(NavigationResultMessage))]
[JsonSerializable(typeof(SurfaceOutlineProcedure))]
[JsonSerializable(typeof(OutlineResultMessage))]
[JsonSerializable(typeof(DesignerAutoSizeResultMessage))]
[JsonSerializable(typeof(SyncResultMessage))]
[JsonSerializable(typeof(SetLanguageFactsMessage))]
[JsonSerializable(typeof(SurfaceProject))]
[JsonSerializable(typeof(SurfaceComponent))]
[JsonSerializable(typeof(SurfaceFinding))]
[JsonSerializable(typeof(EditorMarker))]
public sealed partial class EditorMessageContext : JsonSerializerContext;
