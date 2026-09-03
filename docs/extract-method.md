# Extract Method

The first refactoring beyond rename, and the head of a set. **Shipped 2026-09-03.**

What is below is the design as it was written, kept because it is still what the feature does.
What shipped differs from it in three places, each recorded at the end under
[What shipped](#what-shipped): the name is asked for in a dialog rather than the editor's inline
rename box, only a plainly-value-typed result becomes a `Function`, and the refusal list grew.

`rename` and `renameModule` are the only refactorings this product has - measured against the
live action vocabulary on 2026-08-30, and there is no groundwork for another in this repo or in
the analyzer checkout. Inspections have 122 rules and, since 0.10.0, a settings surface.
The test runner shipped 2026-08-20. Refactoring is what is left.

## What it looks like

Select statements, press Ctrl+. , type a name. The lightbulb that already carries analyzer fixes
gains **Extract method...**, which opens the same inline name box `rename` uses. The new
procedure lands below the one it came from, `Private`, and the selection is replaced by the call.
One round in the Changes pane, one undo.

```vba
Public Sub PostInvoices()
    Dim lastRow As Long
    Dim total As Currency
    Dim i As Long
    lastRow = Sheet1.UsedRange.Rows.Count

    total = 0                                        ' <-- selection starts
    For i = 2 To lastRow
        If Sheet1.Cells(i, 3).Value > 0 Then
            total = total + Sheet1.Cells(i, 3).Value
        End If
    Next i                                           ' <-- selection ends

    Debug.Print "Posted " & total
End Sub
```

becomes

```vba
Public Sub PostInvoices()
    Dim lastRow As Long
    Dim total As Currency
    lastRow = Sheet1.UsedRange.Rows.Count

    total = SumPositiveColumn(lastRow)

    Debug.Print "Posted " & total
End Sub

Private Function SumPositiveColumn(ByVal lastRow As Long) As Currency
    Dim i As Long
    Dim total As Currency

    total = 0
    For i = 2 To lastRow
        If Sheet1.Cells(i, 3).Value > 0 Then
            total = total + Sheet1.Cells(i, 3).Value
        End If
    Next i

    SumPositiveColumn = total
End Function
```

Four decisions came out of the analysis, and each is the feature:

- `lastRow` is **read before it is written** inside the selection, so it is a parameter, `ByVal`
  because the callee never assigns it.
- `total` is **written inside and read after**, so it is the result. One such variable means a
  `Function`, and the caller's assignment is synthesised.
- `i` is **written inside and never read after**, so its `Dim` moves into the new procedure.
- `Sheet1` is **not a local**, so it stays a free reference. Getting this wrong is how a
  refactoring tool produces a six-argument procedure nobody wants.

Two variables written and used after: one becomes the return, the rest become explicit `ByRef`
parameters, because VBA returns one value.

## What the spike measured, 2026-08-30

Driven through the same providers the editor calls, against a real module in DebugFixture.

**The declared type and the scope are already answerable.** `hover` on a local returns
``total As Currency`` and, under it, `Local in PostInvoices`. That is the `As` clause and the
fact that the name is a local of that specific procedure rather than a module-level or a global -
two of the three things a parameter needs, available today with no engine work.

**Read versus write is not.** `textDocument/references` answers module, workbook, line, column,
length and the line's text, and nothing about what the reference DOES. For
`total = total + Sheet1.Cells(i, 3).Value` it returns two entries on the same line, columns 13
and 21, and nothing marks the first as the assignment. Deriving it here would mean re-parsing the
line outside the binder that already knows - the wrong place. **This is the one piece of genuine
engine work the feature needs**, and it belongs upstream in the analyzer with the rest of the
binder, per the standing rule that the analyzer is shared with xlide_vscode.

**`With` scope is modelled, and well.** Hover on `.Address` inside `With Sheet1.Range("A1")`
answers `Range.Address As String`, "Excel host property (read-only)", with the library's own
documentation; the dot offers 201 `Excel.Range` members. So a selection inside a `With` can be
analysed rather than blindly refused - the transformation still has to carry the `With` or decline,
but the information is there.

> Both negative findings in the first pass were wrong, and both were mis-aimed probes rather than
> missing capability: a completion asked at column 16 of `        caption = .Address`, which is a
> space, and a reference asked for `{word: "i"}`, which resolved inside `Option`. Re-aim before
> concluding. The second of those turned out to be a real defect in the word lookup and is fixed.

## Where VBA makes this harder than the same feature elsewhere

- **`ByRef` is the default.** An unmarked parameter silently lets the callee write through to the
  caller's variable. Every parameter this emits is explicitly `ByVal` or `ByRef`, chosen from
  whether the selection assigns it, and never left blank.
- **Arrays cannot be passed `ByVal`.** The rule above yields to the language.
- **`Option Explicit` off** makes undeclared names implicit Variants at procedure scope.
  Extraction changes their scope and with it their lifetime, which is a behaviour change hiding
  inside a refactoring. Refuse, and say so.
- **A `With` block** carries its receiver in leading dots. Either the `With` goes into the new
  procedure or the extraction declines.

## What it refuses

The refusals matter more than the transformation, and each is a check a suite can drive:

- `Exit Sub`, `Exit Function`, or an `Exit For` whose loop is outside the selection
- a `GoTo` crossing the boundary in either direction, and `Resume` or error-handler labels
- a selection that starts inside a block and ends outside it - `If` without its `End If`
- `Static` locals, whose whole meaning is the lifetime extraction would change

## The api

`act("extractMethod", { name, startLine, endLine })`, answering `{did, detail}` with the
refusal's own wording in `detail` - the shape `rename` already uses, so it mirrors the UI and a
suite can drive every refusal above.

## The order of work

1. **Upstream, filed as xlide_vscode#55 - LANDED 2026-08-30** (`9277e10`), wired through the
   engine, the shim and the act the same day and pinned in rename-features.mjs. Measured on the
   spike's own module: `total = total + i` answers write@9 and read@17 as separate rows, the Dim
   and the assignments write, `Debug.Print total` reads. A `kind` on each reference - `read`,
   `write`, `readwrite`. That turned out to be the whole ask. With kinds plus positions, a consumer
   computes the signature itself: a local whose first in-range reference is a read is a
   parameter; one written in range with a read after the end is the return or a `ByRef`; one
   written and never read after moves its `Dim`. No liveness engine, because extraction refuses
   control flow crossing the boundary anyway. The binder already establishes write-ness -
   `set-required` and `assignment-type-mismatch` are diagnostics about writes - and discards it
   on the way out.
2. **Here**: the transformation, the refusals, the lightbulb entry, the name box, the api route,
   and the suite.

Step 1 was the whole risk and it is done. Step 2 is ordinary work on top of the rename
machinery, which already applies multi-site edits through the host and has an undo.

## What shipped

Both steps, 2026-09-03. `engine/src/extractMethod.ts` works out the transformation and every
refusal; `textDocument/extractMethod` carries it; `AddInSession.OnExtractMethodRequested` writes
the one module, syncs the open tab and fills the undo slot; the page offers it from the lightbulb
(`refactor.extract.function`) and the right-click menu, and `ui/editor/src/extractdialog.ts` asks
for the name. 34 checks in `engine/test/extract-method.mjs`, 16 in
`tools\harness\extract-method.mjs`, and a round in `com-leak.mjs`.

Three departures from the design above, each deliberate:

- **A dialog, not the inline rename box.** The editor's rename widget is bound to a symbol at a
  position and is not public API; reusing it for an operation on a selection is not something the
  editor offers. The dialog is one labelled field, and it earns its keep on the refusals: it
  STAYS UP carrying the reason, so a developer who selected half an `If` reads what happened and
  reselects, rather than watching it go past as a toast.

- **Only a plainly-value-typed result becomes the `Function`'s return.** VBA assigns an object
  with `Set` and a value without it, and the declared type is the only evidence in reach. So a
  result declared `Long`, `String`, `Currency` and the rest becomes the return; anything else -
  an object, a user-defined type, an array, and `Variant`, which can hold either - becomes a
  `ByRef` parameter and the procedure stays a `Sub`. Always correct, at the cost of a less
  pretty signature in the cases where nobody can tell in advance.

- **The refusal list grew** by what the first suite found: a selection taking the header line or
  the closer, a procedure with no closer, a blank selection, a name that is not a VBA name, a
  name the module already uses, and `GoSub`. The `With` case turned out to have a third shape
  the design did not name - a selection wholly INSIDE the block, which is neither opening nor
  closing within it - and it is refused only when the selection actually uses a leading dot,
  because statements inside a `With` that never use one move perfectly well.

One correction to the analysis above, found by the suite. `total = total + step` reports the
write at column 5 and the read at column 13, so reading the reference kinds in COLUMN order says
the callee writes `total` before it reads it - which makes it the return value instead of a
`ByRef` parameter, and silently drops the caller's running total. A statement evaluates its
right-hand side before it assigns, so the kinds are read by LINE, and a read anywhere on the
first line that writes counts as reading first. That errs the safe way too: passing a value the
callee did not need costs an argument, and not passing one it did costs the developer their data.
