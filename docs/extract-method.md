# Extract Method

The first refactoring beyond rename, and the head of a set.

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

1. **Upstream**: a range query on the analyzer that answers, for a statement range, which locals
   are read before written, which are written and live after, and each one's declared type. The
   binder has the facts; nothing exposes them.
2. **Here**: the transformation, the refusals, the lightbulb entry, the name box, the api route,
   and the suite.

Step 1 is the whole risk. Step 2 is ordinary work on top of the rename machinery, which already
applies multi-site edits through the host and has an undo.
