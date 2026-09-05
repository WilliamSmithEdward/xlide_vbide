# Attributes through annotations

A VBA module carries attributes the code pane never shows. `VB_PredeclaredId` gives a class a
default instance, so its own name is an object. `VB_UserMemId = 0` names the default member, so
`bag(1)` means `bag.Item(1)`; `-4` names the enumerator `For Each` walks. `VB_Description` is
what the Object Browser and IntelliSense say about a module or a member. `VB_ProcData.VB_Invoke_Func`
is an Excel macro hotkey. `VB_VarDescription` describes a module-level variable. The editor
offers no way to set any of them, and Rubberduck's answer was a comment in the code that names
the attribute. This product reads the same comments and makes the module match.

## The annotations

| Comment | Attribute | Binds to |
| --- | --- | --- |
| `'@ModuleDescription("text")` | `VB_Description = "text"` | the module |
| `'@PredeclaredId` | `VB_PredeclaredId = True` | a class |
| `'@Exposed` | `VB_Exposed = True` | a class |
| `'@Description("text")` | `Member.VB_Description = "text"` | the procedure below |
| `'@DefaultMember` | `Member.VB_UserMemId = 0` | the procedure below, one per class |
| `'@Enumerator` | `Member.VB_UserMemId = -4` | the procedure below |
| `'@ExcelHotkey("D")` | `Member.VB_ProcData.VB_Invoke_Func = "D\n14"` | the procedure below; lower case is Ctrl, upper case Ctrl+Shift |
| `'@VariableDescription("text")` | `Variable.VB_VarDescription = "text"` | the module-level variable below |

A module annotation lives anywhere in the declarations section. A member annotation lives in the
run of comment and blank lines directly above a procedure's header; the same run above a
module-level variable takes a variable description. The reader is lenient about the spelling (any
case; with or without brackets and quotes) and writes the documented form. An annotation that
cannot bind, or a module annotation below the first procedure, is reported on its line rather
than guessed at (`AttributeAnnotations.Read`, `src/Xlide.Vbe.Core/Vba`).

## What the file carries

Attributes are read out of the saved package: each module is a compressed stream in
`vbaProject.bin`, and the stream holds the source with its attribute lines, module-level ones in
a header and member ones directly under each procedure's header. `SavedModules.AttributesOf`
decompresses the whole stream, once per save, for the module asked about
(`ModuleAttributes.Read`). A module the saved file does not carry - added since the last save -
has unknown attributes, and the drift says so rather than pretending.

## The drift

`AttributeDrift.Between` compares the two and answers items in four kinds, each filed in the
Problems pane on the line the developer would look at:

- `annotation-not-applied` (warning): the annotation is there and the attribute does not match.
- `attribute-not-annotated` (info): the attribute is set and nothing in the code says so - an
  import, another tool, or an annotation since deleted.
- `annotation-not-applicable` (warning): a document module (cannot be imported, so cannot take
  attributes), a form (not offered yet), or `'@PredeclaredId` on a standard module.
- `annotation-problem` (warning): malformed or misplaced, with the reason.

The drift is computed on every analysis pass from the snapshot the pass already read
(`RememberAttributeDrift`), so it costs no read of its own, and again the moment an apply lands.

## The write

There is one way to put an attribute on a module the editor accepts: an import. The code pane
rejects an `Attribute` line, the object model has no property for one, and the dialog that sets
one is modal. So `ApplyAttributes` exports the module to a temporary file, rewrites only the
attribute lines the annotations name (`AttributeRewriter.Apply`, byte for byte otherwise),
removes the component and imports the file back under the same name. Decision 17 records why
this is the designed exception to the rule against files.

What it costs, and what is put back: the module's undo history goes; its native breakpoints go
and are re-set from the session's own record; its tab reopens where the caret was; its code is
read back and compared with what went in. Until the workbook is saved the saved package does not
carry what was written, so the applied set is asserted to `SavedModules`, which answers from it
for the drift, for the api, and for the analyzer's predeclared-class seed - so `Registry.Lookup`
stops being reported the moment `'@PredeclaredId` is applied, not at the next save.

Applying writes what the annotations say and leaves every other attribute alone. Taking an
attribute away is a separate act: the `attribute-not-annotated` finding offers it, beside the
text edit that adds the annotation instead.

## When it happens

A save writes the annotations first. Every module of the workbook being saved whose annotations
are not yet its attributes is re-imported with them, and then the save goes; a module with nothing
to write is not touched. An import through the sync feature does the same for the modules it
created or rewrote: the file said what the module should be. Both are under one setting,
`attributes.applyOnSave`, on by default; off leaves the drift in the Problems pane with its quick
fix for whoever wants to choose the moment. There is no menu item: the owner's call
(2026-09-05) was that the save and the import are the moments, and a right-click for it would be
a chore. The re-import keeps the surface still: the module's pane is reopened before the tree
and the tab strip are republished, so the same tabs stay in the same order, the module stays
active and unfolded, and the caret returns to where it was.

## Surfaces

- The Problems pane, with quick fixes: apply the module's annotations now (a host action, since
  no text edit can re-import a module), add the missing annotation (a text edit), remove the
  attribute (a host action).
- A hover on an annotation: what it writes, and what the module has now.
- The api: `attributes` GET for the annotations, the saved attributes and the drift; POST
  `action=apply` and `action=remove`; `settings` for `applyAttributesOnSave`.
  `tools/harness/attributes.mjs` drives the whole loop against `AttributesFixture.xlsm`
  (`tools/New-AttributesFixture.ps1`), saving the workbook on purpose to prove the package
  reader and the save path, and putting the file back afterwards.

## Not done

Forms are not offered: their export is a `.frm` and `.frx` pair and the round trip has not been
proven. Document modules cannot take attributes at all, by the editor's own rule. `@MemberAttribute`
and `@ModuleAttribute`, Rubberduck's escape hatches for arbitrary attributes, are not read.
