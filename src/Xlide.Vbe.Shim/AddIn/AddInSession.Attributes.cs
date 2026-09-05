using System.Text;
using Xlide.Vbe.Core.Engine;
using Xlide.Vbe.Core.Sync;
using Xlide.Vbe.Core.Vba;
using Xlide.Vbe.Shim.Com;
using Xlide.Vbe.Shim.Diagnostics;
using Xlide.Vbe.Shim.Editor;
using Xlide.Vbe.Shim.Engine;

namespace Xlide.Vbe.Shim.AddIn;

/// <summary>
/// Annotations that control the attributes the code pane cannot show.
///
/// A VBA module carries attributes the editor hides and offers no way to set: VB_PredeclaredId
/// (a class with a default instance), VB_Description (what IntelliSense says about a module or a
/// member), VB_UserMemId (the default member, and the enumerator For Each walks), the Excel macro
/// hotkey, and a variable's description. They decide real behaviour and they are invisible. The
/// Rubberduck convention puts each one in the code as a comment - '@PredeclaredId,
/// '@Description("...") and the rest - where it can be read, diffed and edited, and this file makes
/// the module match what the comments say.
///
/// THREE SURFACES, ONE TRUTH. The drift between what the annotations say and what the saved module
/// carries is computed on every analysis pass from the snapshot the pass already read, and
/// published as findings in the Problems pane: an annotation not yet applied, an attribute nothing
/// annotates, an annotation that cannot mean anything where it sits. Each finding offers its fix
/// as a quick fix - apply the module's annotations, add the missing annotation, or take the
/// attribute away - and a hover on an annotation says what it writes and what the module has now.
///
/// THE ONE WAY TO WRITE AN ATTRIBUTE is an import. The code pane rejects an Attribute line, the
/// object model has no property for one, and the only dialog that sets one is modal. So applying
/// exports the module to a temporary file, rewrites exactly the attribute lines the annotations
/// name, removes the component and imports the file back under the same name. That is a designed
/// exception to this product's rule against reading files it writes: the file is read by the
/// editor, not by us, and it is the editor's own import path - the same one the sync feature
/// creates modules through. It costs the module its undo history and its native breakpoints; the
/// breakpoints are put back from this session's own record, and the tab is reopened where it was.
/// A document module cannot be imported, so its annotations are reported as inapplicable rather
/// than silently ignored; a form is not offered yet.
///
/// Until the workbook is saved the saved package does not carry what was applied, so the applied
/// set is ASSERTED to <see cref="SavedModules"/>, which answers from it - for the drift, for the
/// analyzer's predeclared-class seed, and for the api - until the file is saved and read again.
/// </summary>
internal sealed partial class AddInSession
{
    /// <summary>The attribute findings of every project, replaced per project on each pass.</summary>
    private readonly Dictionary<string, List<Finding>> _attributeFindings = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The drift items behind those findings, by project and module, for the fixes and the api.</summary>
    private readonly Dictionary<string, IReadOnlyList<DriftItem>> _attributeDrift = new(StringComparer.OrdinalIgnoreCase);

    private const int StandardModuleType = 1;
    private const int ClassModuleType = 2;

    private static string DriftKey(string projectId, string module) => $"{projectId}\0{module}";

    /// <summary>Every attribute finding, for the Problems pane and the problems route.</summary>
    private IEnumerable<Finding> AttributeFindings()
    {
        lock (_attributeFindings)
        {
            return [.. _attributeFindings.Values.SelectMany(list => list)];
        }
    }

    /// <summary>
    /// The drift of a whole project, from the snapshot an analysis pass just read. Runs on the
    /// pass's worker thread; the saved package is read through SavedModules' own guarded cache.
    /// Republishes the findings through a hop to the host thread only when they changed.
    /// </summary>
    private void RememberAttributeDrift(string projectId, IReadOnlyList<EngineModule> modules)
    {
        var findings = new List<Finding>();
        var saved = SavedModules.For(projectId);
        foreach (var module in modules)
        {
            var drift = DriftFor(saved, module.ModuleName, module.Type, module.Source);
            lock (_attributeDrift)
            {
                _attributeDrift[DriftKey(projectId, module.ModuleName)] = drift;
            }
            findings.AddRange(drift.Select(item => FindingOf(item, module.ModuleName, module.Source, projectId)));
        }

        bool changed;
        lock (_attributeFindings)
        {
            var before = _attributeFindings.TryGetValue(projectId, out var held) ? held : [];
            changed = before.Count != findings.Count
                || before.Zip(findings).Any(pair => pair.First != pair.Second);
            _attributeFindings[projectId] = findings;
        }

        if (changed)
        {
            Log.Verbose($"attributes: {System.IO.Path.GetFileName(projectId)} has {findings.Count} annotation finding(s)");
            _editorSurface?.RunOnHostThread(() =>
            {
                PublishFindingsToSurface();
                PublishMarkersToSurface();
            });
        }
    }

    private static IReadOnlyList<DriftItem> DriftFor(SavedModules? saved, string module, string kind, string source)
    {
        var annotations = AttributeAnnotations.Read(source);
        if (annotations.Annotations.Count == 0 && annotations.Problems.Count == 0 && saved?.Knows(module) != true)
        {
            return [];
        }

        var actual = saved?.Knows(module) == true ? saved.AttributesOf(module) : null;
        return AttributeDrift.Between(source, annotations, actual, kind, module);
    }

    private static Finding FindingOf(DriftItem item, string module, string source, string projectId)
    {
        var length = LineLength(source, item.Line);
        return new Finding(module, item.Code, item.Message, item.Severity, item.Line, 1, item.Line, Math.Max(2, length + 1), projectId);
    }

    private static int LineLength(string source, int line)
    {
        var lines = source.Split('\n');
        return line >= 1 && line <= lines.Length ? lines[line - 1].TrimEnd('\r').Length : 0;
    }

    /// <summary>The 1-based line a character offset of the source falls on.</summary>
    private static int LineOfOffset(string source, int offset)
    {
        var line = 1;
        for (var at = 0; at < Math.Min(offset, source.Length); at++)
        {
            if (source[at] == '\n')
            {
                line++;
            }
        }
        return line;
    }

    /// <summary>The character offset at which a 1-based line starts.</summary>
    private static int OffsetOfLine(string source, int line)
    {
        var offset = 0;
        var seen = 1;
        while (seen < line && offset < source.Length)
        {
            var next = source.IndexOf('\n', offset);
            if (next < 0)
            {
                return source.Length;
            }
            offset = next + 1;
            seen++;
        }
        return offset;
    }

    /// <summary>
    /// The quick fixes the attribute findings of the shown module offer over a span, alongside the
    /// engine's own. An annotation not applied offers the apply; an attribute nothing annotates
    /// offers the annotation as a text edit and the removal as a host action.
    /// </summary>
    private SurfaceCodeAction[] AttributeCodeActions(string module, string? projectId, string source, int start, int end)
    {
        if (projectId is null)
        {
            return [];
        }

        IReadOnlyList<DriftItem> drift;
        lock (_attributeDrift)
        {
            if (!_attributeDrift.TryGetValue(DriftKey(projectId, module), out var held))
            {
                return [];
            }
            drift = held;
        }

        var firstLine = LineOfOffset(source, start);
        var lastLine = LineOfOffset(source, end);
        var display = DisplayFromProjectId(projectId);
        var actions = new List<SurfaceCodeAction>();
        var offeredApply = false;

        foreach (var item in drift.Where(one => one.Line >= firstLine && one.Line <= lastLine))
        {
            var lineStart = OffsetOfLine(source, item.Line);
            var lineEnd = lineStart + LineLength(source, item.Line);
            switch (item.Kind)
            {
                case DriftKind.AnnotationNotApplied when !offeredApply:
                    offeredApply = true;
                    actions.Add(new SurfaceCodeAction(
                        $"Apply annotations to {module}'s attributes now",
                        IsPreferred: true, item.Code, lineStart, lineEnd, [],
                        "applyAttributes", [module, display]));
                    break;

                case DriftKind.AttributeNotAnnotated when item.Annotation is { } kind:
                {
                    var argument = ArgumentFor(item, projectId, module);
                    var annotation = new Annotation(kind, item.Line, argument, item.Target, item.Line);
                    actions.Add(new SurfaceCodeAction(
                        $"Add {annotation.Canonical} above",
                        IsPreferred: true, item.Code, lineStart, lineEnd,
                        [new SurfaceTextEdit(lineStart, lineStart, annotation.Canonical + "\r\n")]));
                    actions.Add(new SurfaceCodeAction(
                        $"Remove {AttributeDrift.AttributeName(kind)} from {item.Target ?? module}",
                        IsPreferred: false, item.Code, lineStart, lineEnd, [],
                        "removeAttribute", [module, display, kind.ToString(), item.Target, item.Occurrence.ToString()]));
                    break;
                }
            }
        }

        return [.. actions];
    }

    /// <summary>The text the missing annotation should carry: the attribute's current value.</summary>
    private static string? ArgumentFor(DriftItem item, string projectId, string module)
    {
        if (item.Annotation is not { } kind || !AttributeAnnotations.NeedsArgument(kind))
        {
            return null;
        }
        var actual = SavedModules.For(projectId)?.AttributesOf(module);
        return kind switch
        {
            AnnotationKind.ModuleDescription => actual?.Description,
            AnnotationKind.Description => actual?.Member(item.Target ?? string.Empty).Description,
            AnnotationKind.ExcelHotkey => actual?.Member(item.Target ?? string.Empty).Hotkey,
            _ => actual?.VariableDescriptions.TryGetValue(item.Target ?? string.Empty, out var text) == true ? text : null,
        } ?? string.Empty;
    }

    /// <summary>
    /// What an annotation under the caret writes, and what the module carries now, or null when
    /// the offset is not on one of our annotations.
    /// </summary>
    private static SurfaceHoverPayload? AttributeHover(string module, string? projectId, string source, int offset)
    {
        var line = LineOfOffset(source, offset);
        var annotations = AttributeAnnotations.Read(source);
        var annotation = annotations.Annotations.FirstOrDefault(one => one.Line == line);
        if (annotation is null)
        {
            return null;
        }

        var attribute = AttributeDrift.AttributeName(annotation.Kind);
        var owner = annotation.Target is null ? string.Empty : $"{annotation.Target}.";
        var writes = annotation.Kind switch
        {
            AnnotationKind.PredeclaredId or AnnotationKind.Exposed => $"Attribute {attribute} = True",
            AnnotationKind.DefaultMember => $"Attribute {owner}{attribute} = 0",
            AnnotationKind.Enumerator => $"Attribute {owner}{attribute} = -4",
            AnnotationKind.ExcelHotkey => $"Attribute {owner}{attribute} = {ModuleAttributes.Literal(ModuleAttributes.InvokeFuncFor(annotation.Argument ?? string.Empty))}",
            _ => $"Attribute {owner}{attribute} = {ModuleAttributes.Literal(annotation.Argument ?? string.Empty)}",
        };

        var saved = projectId is null ? null : SavedModules.For(projectId);
        var actual = saved?.Knows(module) == true ? saved.AttributesOf(module) : null;
        string now;
        if (actual is null)
        {
            now = "The saved workbook does not carry this module's attributes yet; save it, or apply the annotations.";
        }
        else
        {
            var member = actual.Member(annotation.Target ?? string.Empty);
            var value = annotation.Kind switch
            {
                AnnotationKind.ModuleDescription => actual.Description is null ? null : ModuleAttributes.Literal(actual.Description),
                AnnotationKind.PredeclaredId => actual.PredeclaredId?.ToString(),
                AnnotationKind.Exposed => actual.Exposed?.ToString(),
                AnnotationKind.Description => member.Description is null ? null : ModuleAttributes.Literal(member.Description),
                AnnotationKind.DefaultMember or AnnotationKind.Enumerator => member.UserMemId?.ToString(),
                AnnotationKind.ExcelHotkey => member.Hotkey is null ? null : ModuleAttributes.Literal(ModuleAttributes.InvokeFuncFor(member.Hotkey)),
                _ => actual.VariableDescriptions.TryGetValue(annotation.Target ?? string.Empty, out var text) ? ModuleAttributes.Literal(text) : null,
            };
            now = value is null
                ? $"The module has no {attribute}{(annotation.Target is null ? string.Empty : $" on {annotation.Target}")} yet."
                : $"The module has {attribute} = {value}.";
        }

        var documentation = annotation.Kind switch
        {
            AnnotationKind.PredeclaredId => "A class with a default instance: its own name is an object, so `Registry.Lookup` works without New.",
            AnnotationKind.Exposed => "A class other projects can see and create.",
            AnnotationKind.DefaultMember => "The member VBA calls when the object is used without one: `bag(1)` means `bag.Item(1)`.",
            AnnotationKind.Enumerator => "The member For Each walks; conventionally `NewEnum` returning IUnknown.",
            AnnotationKind.ExcelHotkey => "An Excel macro shortcut: a lower-case letter is Ctrl+letter, an upper-case one Ctrl+Shift+letter.",
            AnnotationKind.VariableDescription => "The description the Object Browser shows for a module-level variable.",
            _ => "The description the Object Browser and IntelliSense show.",
        };

        var lineStart = OffsetOfLine(source, line);
        return new SurfaceHoverPayload(
            annotation.Canonical,
            [$"Writes {writes}", now],
            documentation,
            lineStart,
            lineStart + LineLength(source, line));
    }

    /// <summary>The page chose a fix or a menu item the host performs.</summary>
    private void OnHostActionRequested(string command, string?[] arguments)
    {
        string? Arg(int at) => at < arguments.Length ? arguments[at] : null;

        switch (command)
        {
            case "applyAttributes":
            {
                var refused = ApplyAttributes(Arg(0) ?? string.Empty, Arg(1), out var changes, out var skipped);
                _editorSurface?.Notify(refused ?? AppliedNotice(Arg(0) ?? string.Empty, changes, skipped));
                break;
            }
            case "removeAttribute":
            {
                if (!Enum.TryParse<AnnotationKind>(Arg(2), ignoreCase: true, out var kind))
                {
                    _editorSurface?.Notify($"'{Arg(2)}' is not an attribute annotation this product knows.");
                    break;
                }
                var occurrence = int.TryParse(Arg(4), out var parsed) ? parsed : 0;
                var refused = RemoveAttribute(Arg(0) ?? string.Empty, Arg(1), kind, Arg(3), occurrence, out var changes);
                _editorSurface?.Notify(refused ?? (changes.Count == 0
                    ? $"{Arg(0)} did not carry that attribute."
                    : $"Removed from {Arg(0)}: {string.Join("; ", changes)}. Save the workbook to keep it."));
                break;
            }
            default:
                Log.Warn($"hostAction: '{command}' is not an action this host performs");
                _editorSurface?.Notify($"'{command}' is not something this editor can do.");
                break;
        }
    }

    private static string AppliedNotice(string module, IReadOnlyList<AttributeChange> changes, IReadOnlyList<string> skipped)
    {
        var said = changes.Count == 0
            ? $"{module}'s attributes already match its annotations."
            : $"Applied to {module}: {string.Join("; ", changes)}. Save the workbook to keep it.";
        return skipped.Count == 0 ? said : $"{said} Skipped: {string.Join(" ", skipped)}";
    }

    /// <summary>
    /// Writes a module's annotations into its attributes. Answers the refusal, or null when the
    /// module was rewritten (or already matched). Host thread.
    /// </summary>
    private string? ApplyAttributes(string module, string? projectDisplay, out IReadOnlyList<AttributeChange> changes, out IReadOnlyList<string> skipped)
    {
        changes = [];
        skipped = [];
        return Reimport(module, projectDisplay, "apply",
            (exported, source) =>
            {
                var annotations = AttributeAnnotations.Read(source);
                if (annotations.Annotations.Count == 0)
                {
                    return (null, $"{module} carries no attribute annotations. Add one - '@PredeclaredId, '@Description(\"...\") and the rest - above the module's first procedure or above a member, then apply.");
                }
                return (AttributeRewriter.Apply(exported, annotations), null);
            },
            out changes, out skipped);
    }

    /// <summary>Takes one managed attribute off a module, a member or a variable. Host thread.</summary>
    private string? RemoveAttribute(string module, string? projectDisplay, AnnotationKind kind, string? target, int occurrence, out IReadOnlyList<AttributeChange> changes)
    {
        var refused = Reimport(module, projectDisplay, "remove",
            (exported, _) => (AttributeRewriter.Remove(exported, kind, target, occurrence), null),
            out changes, out _);
        return refused;
    }

    /// <summary>
    /// The export, rewrite and import that every attribute write goes through. <paramref name="rewrite"/>
    /// answers the rewritten export or a refusal. Answers the refusal, or null.
    /// </summary>
    private string? Reimport(
        string module,
        string? projectDisplay,
        string verb,
        Func<string, string, (RewriteResult? Result, string? Refused)> rewrite,
        out IReadOnlyList<AttributeChange> changes,
        out IReadOnlyList<string> skipped)
    {
        changes = [];
        skipped = [];

        if (string.IsNullOrWhiteSpace(module))
        {
            return "No module was named.";
        }

        if (ProjectModeNow() != DesignMode)
        {
            return $"The project is stopped in the debugger, so {module} was not touched: re-importing a module would reset it and lose the run. Press Reset and try again.";
        }

        var projectId = ProjectIdFromDisplay(projectDisplay);
        using var component = FindComponent(module, projectId, out var foundIn);
        if (component is null)
        {
            return $"There is no module named {module}" + (projectDisplay is null ? "." : $" in {projectDisplay}.");
        }

        var type = component.GetInt32("Type");
        if (type is not (StandardModuleType or ClassModuleType))
        {
            return type == DocumentComponent
                ? $"{module} is a document module. The editor only takes attributes through an import, and a document module cannot be imported, so its annotations cannot be applied."
                : $"{module} is a UserForm. Applying attributes to a form is not offered yet.";
        }

        var owner = foundIn ?? projectId;
        var display = DisplayFromProjectId(owner);
        using var project = FindProjectByDisplayName(display) ?? _editor.GetObject("ActiveVBProject");
        using var components = project?.GetObject("VBComponents");
        if (project is null || components is null)
        {
            return $"{module}'s project would not open its component list.";
        }

        // The page's typing reaches the module before the module is exported, or the export would
        // carry text the developer no longer has.
        _editorSurface?.FlushEdits();
        var source = ProjectReader.ReadSource(component) ?? string.Empty;

        var temporary = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"xlide-attributes-{Guid.NewGuid():N}{(type == ClassModuleType ? ".cls" : ".bas")}");
        try
        {
            component.Invoke("Export", temporary);
            var exported = File.ReadAllText(temporary, Encoding.Latin1);

            var (result, refused) = rewrite(exported, source);
            if (refused is not null || result is null)
            {
                return refused ?? "nothing to write";
            }

            changes = result.Changes;
            skipped = result.Skipped;
            var applied = ModuleAttributes.Read(result.Text);
            var savedPath = ProjectReader.SavedPathOf(project);

            if (result.Changes.Count == 0)
            {
                // Already as annotated. The saved file may not know that yet - a module never
                // saved since its import - so the answer stands in until it does, and the
                // findings settle now rather than at the next save.
                SavedModules.Assert(savedPath, module, applied);
                RefreshAttributeDriftFor(owner, module, type, source);
                Log.Info($"attributes: {verb} on {module} changed nothing");
                return null;
            }

            File.WriteAllText(temporary, result.Text, Encoding.Latin1);

            // What the module had on the surface, to put back once the new one is in.
            var wasOpen = _editorSurface?.TextOf(module, display) is not null;
            var wasShown = string.Equals(_editorSurface?.Module, module, StringComparison.OrdinalIgnoreCase)
                && string.Equals(_shownProject, owner, StringComparison.OrdinalIgnoreCase);
            var shownBefore = _editorSurface?.Module;
            var shownProjectBefore = DisplayFromProjectId(_shownProject);
            var shownLineBefore = _editorSurface?.CaretLine ?? 1;
            var shownColumnBefore = _editorSurface?.CaretColumn ?? 1;
            var caretLine = wasShown ? shownLineBefore : 1;
            var caretColumn = wasShown ? shownColumnBefore : 1;
            var breakpoints = BreakpointsFor(module, display).ToArray();
            foreach (var key in new[] { WrittenKey(module, display), WrittenKey(module, null) })
            {
                _breakpoints.Remove(key);
                _writtenModules.Remove(key);
            }

            _editorSurface?.DiscardEdits(module, display);
            components.InvokeWithObject("Remove", component);
            components.Invoke("Import", temporary);

            using var imported = FindComponent(module, owner, out _);
            if (imported is null)
            {
                Log.Warn($"attributes: the editor imported {module} but no module of that name came back");
                ComponentsChanged();
                return $"{module} was exported and removed, but the import did not bring it back under its name. Its text is in {temporary}.";
            }

            var roundTrip = ProjectReader.ReadSource(imported) ?? string.Empty;
            if (!ModuleSync.SameText(ModuleSync.CodeWithoutHeader(roundTrip), ModuleSync.CodeWithoutHeader(source)))
            {
                Log.Warn($"attributes: {module} came back from the import with different code");
            }

            SavedModules.Assert(savedPath, module, applied);
            Log.Info($"attributes: {verb} on {module}: {string.Join("; ", result.Changes)}");

            // THE MODULE NEVER LEAVES THE SURFACE. Its pane went with the old component; it is
            // opened again HERE, before the tree and the tab strip are republished, so both lists
            // still hold the module when they are drawn and nothing flickers out and back in. The
            // native breakpoints went too, and this session's record puts them back through the
            // same toggle a click makes; then the caret returns to where the developer had it,
            // and the module that was showing before shows again if this was not it.
            if (wasOpen || breakpoints.Length > 0)
            {
                GoTo(module, breakpoints.Length > 0 ? breakpoints[0] : caretLine, 1, display);
                foreach (var line in breakpoints)
                {
                    ToggleBreakpoint(line);
                }

                if (wasShown)
                {
                    GoTo(module, caretLine, caretColumn, display);
                }
                else if (shownBefore is not null)
                {
                    // The developer was in another module; this one was open behind it, or only
                    // carried breakpoints. Back to theirs, caret where it was.
                    GoTo(shownBefore, shownLineBefore, shownColumnBefore, shownProjectBefore);
                }
            }

            ComponentsChanged();

            RefreshAttributeDriftFor(owner, module, type, roundTrip);
            _analysis?.Reanalyse();
            return null;
        }
        catch (Exception ex)
        {
            Log.Warn($"attributes: {verb} on {module} failed ({ex.GetType().Name}: {ex.Message})");
            ComponentsChanged();
            return $"{module} could not be rewritten: {ex.Message}";
        }
        finally
        {
            try
            {
                if (File.Exists(temporary))
                {
                    File.Delete(temporary);
                }
            }
            catch (IOException)
            {
                // A temp file that would not go is not worth failing the developer over.
            }
        }
    }

    /// <summary>
    /// Writes the annotations of every module of the shown workbook whose attributes do not yet
    /// match, before a save. Only modules with drift of that kind are touched; a refusal is said
    /// and the save goes on regardless, because a save the developer asked for is not something
    /// an attribute may hold up.
    /// </summary>
    private void ApplyAnnotationsBeforeSave()
    {
        if (_shownProject is not { } projectId)
        {
            return;
        }

        List<string> pending;
        lock (_attributeDrift)
        {
            var prefix = DriftKey(projectId, string.Empty);
            pending = [.. _attributeDrift
                .Where(pair => pair.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    && pair.Value.Any(item => item.Kind == DriftKind.AnnotationNotApplied))
                .Select(pair => pair.Key[prefix.Length..])];
        }

        if (pending.Count == 0)
        {
            return;
        }

        ApplyAnnotationsTo(projectId, pending, "before saving");
    }

    /// <summary>The same, for modules a sync import just created or rewrote from files.</summary>
    private void ApplyAnnotationsAfterImport(string? projectId, IReadOnlyList<string> modules)
    {
        if (projectId is null || modules.Count == 0 || !_settings.ApplyAttributesOnSave)
        {
            return;
        }

        // The drift cache may not have seen the imported text yet; each module is read now.
        var saved = SavedModules.For(projectId);
        var pending = new List<string>();
        foreach (var module in modules)
        {
            using var component = FindComponent(module, projectId, out _);
            if (component is null)
            {
                continue;
            }
            var type = component.GetInt32("Type");
            if (type is not (StandardModuleType or ClassModuleType))
            {
                continue;
            }
            var source = ProjectReader.ReadSource(component) ?? string.Empty;
            if (DriftFor(saved, module, ProjectReader.TypeName(type), source).Any(item => item.Kind == DriftKind.AnnotationNotApplied))
            {
                pending.Add(module);
            }
        }

        ApplyAnnotationsTo(projectId, pending, "after the import");
    }

    private void ApplyAnnotationsTo(string projectId, List<string> modules, string when)
    {
        if (modules.Count == 0)
        {
            return;
        }

        var display = DisplayFromProjectId(projectId);
        var written = new List<string>();
        var refusals = new List<string>();
        foreach (var module in modules)
        {
            var refused = ApplyAttributes(module, display, out var changes, out _);
            if (refused is not null)
            {
                refusals.Add($"{module}: {refused}");
            }
            else if (changes.Count > 0)
            {
                written.Add($"{module} ({changes.Count})");
            }
        }

        if (written.Count > 0)
        {
            Log.Info($"attributes: {when}, wrote annotations into {string.Join(", ", written)}");
            _editorSurface?.Notify($"Annotations written into {string.Join(", ", written)} {when}.");
        }
        if (refusals.Count > 0)
        {
            Log.Warn($"attributes: {when}, refused: {string.Join(" | ", refusals)}");
            _editorSurface?.Notify($"Annotations not written {when}: {string.Join(" ", refusals)}");
        }
    }

    /// <summary>Recomputes one module's drift now, so the findings do not wait for the next pass.</summary>
    private void RefreshAttributeDriftFor(string? projectId, string module, int type, string source)
    {
        if (projectId is null)
        {
            return;
        }

        var kind = ProjectReader.TypeName(type);
        var drift = DriftFor(SavedModules.For(projectId), module, kind, source);
        lock (_attributeDrift)
        {
            _attributeDrift[DriftKey(projectId, module)] = drift;
        }
        lock (_attributeFindings)
        {
            var held = _attributeFindings.TryGetValue(projectId, out var list) ? list : [];
            _attributeFindings[projectId] = [
                .. held.Where(finding => !string.Equals(finding.Module, module, StringComparison.OrdinalIgnoreCase)),
                .. drift.Select(item => FindingOf(item, module, source, projectId)),
            ];
        }
        PublishFindingsToSurface();
        PublishMarkersToSurface(module);
    }

    /// <summary>
    /// The module's annotations, its saved attributes and the drift between them, for the api's
    /// attributes route. Reads the module through the object model, so it needs no pass to have run.
    /// </summary>
    private DebugAttributesReply? DescribeAttributes(string module, string? projectDisplay, out string? refused)
    {
        refused = null;
        var projectId = ProjectIdFromDisplay(projectDisplay);
        using var component = FindComponent(module, projectId, out var foundIn);
        if (component is null)
        {
            refused = $"There is no module named {module}" + (projectDisplay is null ? "." : $" in {projectDisplay}.");
            return null;
        }

        var owner = foundIn ?? projectId ?? string.Empty;
        var type = component.GetInt32("Type");
        var kind = ProjectReader.TypeName(type);
        var source = ProjectReader.ReadSource(component) ?? string.Empty;
        var annotations = AttributeAnnotations.Read(source);
        var saved = SavedModules.For(owner);
        var actual = saved?.Knows(module) == true ? saved.AttributesOf(module) : null;
        var drift = AttributeDrift.Between(source, annotations, actual, kind, module);

        return new DebugAttributesReply(
            module,
            DisplayFromProjectId(owner),
            kind,
            type is StandardModuleType or ClassModuleType,
            actual is not null,
            SavedModules.AssertedFor(owner, module) is not null,
            [.. annotations.Annotations.Select(one => new DebugAnnotationRow(one.Kind.ToString(), one.Line, one.Argument, one.Target, one.TargetLine, one.Canonical))],
            [.. annotations.Problems.Select(one => new DebugAnnotationProblemRow(one.Line, one.Message))],
            actual is null ? null : new DebugAttributeSetRow(
                actual.Description,
                actual.PredeclaredId,
                actual.Exposed,
                [.. actual.Members.Select(pair => new DebugMemberAttributesRow(pair.Key, pair.Value.Description, pair.Value.UserMemId, pair.Value.Hotkey))],
                [.. actual.VariableDescriptions.Select(pair => new DebugVariableDescriptionRow(pair.Key, pair.Value))]),
            [.. drift.Select(one => new DebugDriftRow(one.Code, one.Severity, one.Line, one.Message, one.Annotation?.ToString(), one.Target, one.Occurrence))]);
    }
}
