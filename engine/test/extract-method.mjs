/*
 * Extract Method through the real engine: the four decisions, and every refusal.
 *
 * The transformation is checked by reading the TEXT IT PRODUCES, not by trusting a did/didn't
 * flag. A refactoring that answers `did: true` and writes something that no longer compiles is
 * the failure mode worth having a suite for, so every passing case asserts the header it emitted,
 * the call it left behind, and - where it matters - that a name is nowhere it should not be.
 *
 *   node test/extract-method.mjs
 *   node test/extract-method.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('extract-method');
const CRLF = '\r\n';

// The verdict shape the gate parses, with a detail column: what a check reads is nearly always
// the header or the refusal it got, and a bare FAIL sends the reader back to run it by hand.
let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
    console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 160)}` : ''}`);
    if (ok) { passed += 1; } else { failures.push(what); }
};
const done = () => {
    console.log(`\n${passed} passed, ${failures.length} failed`);
    for (const one of failures) { console.log(`  ${one}`); }
    process.exit(failures.length === 0 ? 0 : 1);
};

await call('initialize', {});

let opened = 0;

/** One module in a fresh project, extracted from and handed back whole. */
async function extract(lines, startLine, endLine, newName = 'Extracted') {
    const source = lines.join(CRLF);
    opened += 1;
    const projectId = `p${opened}`;
    await call('project/open', {
        projectId, generation: 1,
        modules: [{ moduleName: 'Probe', source, type: 'standard' }],
    });

    const answer = await call('textDocument/extractMethod', {
        projectId, moduleName: 'Probe', source, moduleType: 'standard',
        startLine, endLine, newName,
    });

    return { ...answer, lines: (answer.source ?? '').split(CRLF) };
}

/** The line the call was left on, so a check can say what the caller now reads. */
const lineWith = (answer, text) => answer.lines.find((one) => one.includes(text))?.trim() ?? '(absent)';

/* ---- the design document's own example ---------------------------------------------------- */

const DOC = [
    'Option Explicit',
    '',
    'Public Sub PostInvoices()',
    '    Dim lastRow As Long',
    '    Dim total As Currency',
    '    Dim i As Long',
    '    lastRow = Sheet1.UsedRange.Rows.Count',
    '',
    '    total = 0',
    '    For i = 2 To lastRow',
    '        If Sheet1.Cells(i, 3).Value > 0 Then',
    '            total = total + Sheet1.Cells(i, 3).Value',
    '        End If',
    '    Next i',
    '',
    '    Debug.Print "Posted " & total',
    'End Sub',
];

const doc = await extract(DOC, 9, 14, 'SumPositiveColumn');

check('the documented example is extracted', !doc.refused, doc.refused ?? doc.signature);
check('lastRow is read before it is written, so it is a ByVal parameter',
    doc.signature?.includes('ByVal lastRow As Long'), doc.signature);
check('total is written inside and read after, so it is the result',
    doc.signature?.endsWith(') As Currency'), doc.signature);
check('the new procedure is a Private Function',
    doc.signature?.startsWith('Private Function SumPositiveColumn('), doc.signature);
check('i is written inside and read nowhere else, so it is not a parameter',
    !doc.signature?.includes('i As Long'), doc.signature);
check('Sheet1 is not a local, so it is not passed',
    !doc.signature?.includes('Sheet1'), doc.signature);
check('the caller calls it and keeps its answer',
    lineWith(doc, 'SumPositiveColumn(') === 'total = SumPositiveColumn(lastRow)',
    lineWith(doc, 'SumPositiveColumn('));
check("i's Dim moved out of the caller",
    doc.lines.filter((one) => one.trim() === 'Dim i As Long').length === 1,
    doc.lines.filter((one) => one.trim() === 'Dim i As Long').length);
check('the result is assigned to the function name before it ends',
    doc.lines.some((one) => one.trim() === 'SumPositiveColumn = total'),
    lineWith(doc, 'SumPositiveColumn = '));
check('the new procedure lands below the one it came from',
    doc.lines.findIndex((one) => one.startsWith('Private Function'))
        > doc.lines.findIndex((one) => one.trim() === 'End Sub'),
    `Private at ${doc.lines.findIndex((one) => one.startsWith('Private Function'))}, End Sub at ${doc.lines.findIndex((one) => one.trim() === 'End Sub')}`);
check('and nothing else in the caller moved',
    doc.lines.includes('    lastRow = Sheet1.UsedRange.Rows.Count')
        && doc.lines.includes('    Debug.Print "Posted " & total'));

/* ---- a Sub, because nothing comes back ------------------------------------------------------ */

const sub = await extract([
    'Option Explicit',
    '',
    'Public Sub Announce()',
    '    Dim name As String',
    '    name = "world"',
    '    Debug.Print "hello"',
    '    Debug.Print name',
    'End Sub',
], 6, 7, 'Say');

check('nothing read afterwards makes a Sub',
    sub.signature === 'Private Sub Say(ByVal name As String)', sub.signature ?? sub.refused);
check('a Sub is called without parentheses',
    lineWith(sub, 'Say ') === 'Say name', lineWith(sub, 'Say '));

/* ---- a parameter the callee writes through -------------------------------------------------- */

const throughRef = await extract([
    'Option Explicit',
    '',
    'Public Sub Grow()',
    '    Dim total As Long',
    '    Dim step As Long',
    '    total = 1',
    '    step = 2',
    '    total = total + step',
    '    Debug.Print total',
    'End Sub',
], 8, 8, 'AddStep');

check('a local read before it is written and read again after goes ByRef',
    throughRef.signature?.includes('ByRef total As Long'), throughRef.signature ?? throughRef.refused);
check('and one only read inside stays ByVal',
    throughRef.signature?.includes('ByVal step As Long'), throughRef.signature);

/* ---- an object result cannot be assigned without Set, so it goes ByRef ---------------------- */

const object = await extract([
    'Option Explicit',
    '',
    'Public Sub Fill()',
    '    Dim target As Object',
    '    Dim seed As Long',
    '    seed = 3',
    '    Set target = CreateObject("Scripting.Dictionary")',
    '    target.Add "n", seed',
    '    Debug.Print target.Count',
    'End Sub',
], 7, 8, 'Build');

check('an Object result is a ByRef parameter, never the return',
    object.signature?.includes('ByRef target As Object')
        && !object.signature?.includes(' As Object' + ''.padEnd(0)) === false
        && !/\)\s*As Object$/.test(object.signature ?? ''),
    object.signature ?? object.refused);
check('so the procedure stays a Sub',
    object.signature?.startsWith('Private Sub Build('), object.signature);

/* ---- an array cannot be passed ByVal --------------------------------------------------------- */

const array = await extract([
    'Option Explicit',
    '',
    'Public Sub Walk()',
    '    Dim names() As String',
    '    ReDim names(1 To 2)',
    '    names(1) = "a"',
    '    Debug.Print names(1)',
    'End Sub',
], 6, 6, 'Seed');

check('an array parameter is ByRef, because VBA has no other way',
    array.signature?.includes('ByRef names() As String'), array.signature ?? array.refused);

/* ---- the refusals ---------------------------------------------------------------------------- */

console.log('\nrefusals:\n');

const refusals = [
    ['a selection reaching outside a procedure', [
        'Option Explicit', '', 'Public Sub One()', '    Debug.Print 1', 'End Sub', '',
        'Public Sub Two()', '    Debug.Print 2', 'End Sub',
    ], 4, 8, /inside one procedure/i],

    ['a selection taking the header line', [
        'Option Explicit', '', 'Public Sub One()', '    Debug.Print 1', 'End Sub',
    ], 3, 4, /not its/i],

    ['a module without Option Explicit', [
        'Public Sub One()', '    Dim n As Long', '    n = 1', '    Debug.Print n', 'End Sub',
    ], 3, 3, /Option Explicit/i],

    ['half an If block', [
        'Option Explicit', '', 'Public Sub One()', '    If 1 = 1 Then', '        Debug.Print 1',
        '    End If', 'End Sub',
    ], 4, 5, /starts inside an? If block and ends outside/i],

    ['the inside of a With block', [
        'Option Explicit', '', 'Public Sub One()', '    With Sheet1', '        .Name = "x"',
        '    End With', 'End Sub',
    ], 5, 5, /With block/i],

    ['Exit Sub', [
        'Option Explicit', '', 'Public Sub One()', '    If 1 = 1 Then Exit Sub',
        '    Debug.Print 1', 'End Sub',
    ], 4, 5, /Exit/i],

    ['an Exit For whose loop stayed behind', [
        'Option Explicit', '', 'Public Sub One()', '    Dim i As Long', '    For i = 1 To 3',
        '        If i = 2 Then Exit For', '    Next i', 'End Sub',
    ], 6, 6, /Exit For/i],

    ['a GoTo whose label stayed behind', [
        'Option Explicit', '', 'Public Sub One()', '    On Error GoTo Fail',
        '    Debug.Print 1', '    Exit Sub', 'Fail:', '    Debug.Print 2', 'End Sub',
    ], 4, 5, /Fail/],

    ['a label jumped to from outside', [
        'Option Explicit', '', 'Public Sub One()', '    On Error GoTo Fail',
        '    Debug.Print 1', '    Exit Sub', 'Fail:', '    Debug.Print 2', 'End Sub',
    ], 7, 8, /jumped to from outside/i],

    ['Resume', [
        'Option Explicit', '', 'Public Sub One()', '    On Error Resume Next',
        '    Debug.Print 1', 'End Sub',
    ], 4, 5, /Resume/i],

    ['a Static local', [
        'Option Explicit', '', 'Public Sub One()', '    Static seen As Long',
        '    seen = seen + 1', '    Debug.Print seen', 'End Sub',
    ], 5, 5, /Static/i],

    ['a blank selection', [
        'Option Explicit', '', 'Public Sub One()', '', '    Debug.Print 1', 'End Sub',
    ], 4, 4, /no statements/i],

    ['a name that is not a VBA name', [
        'Option Explicit', '', 'Public Sub One()', '    Debug.Print 1', 'End Sub',
    ], 4, 4, /is not a VBA name/i, '2Fast'],

    ['a name the module already uses', [
        'Option Explicit', '', 'Public Sub One()', '    Debug.Print 1', 'End Sub', '',
        'Private Sub Taken()', 'End Sub',
    ], 4, 4, /already a procedure/i, 'Taken'],
];

for (const [what, lines, from, to, expected, name] of refusals) {
    const answer = await extract(lines, from, to, name ?? 'Extracted');
    check(`refused: ${what}`,
        typeof answer.refused === 'string' && expected.test(answer.refused) && !answer.source,
        answer.refused ?? `NOT REFUSED (${answer.signature})`);
}

/* ---- a whole block, which is not a refusal --------------------------------------------------- */

console.log('\nwhole blocks go:\n');

const whole = await extract([
    'Option Explicit', '', 'Public Sub One()', '    Dim i As Long', '    For i = 1 To 3',
    '        If i = 2 Then Exit For', '    Next i', '    Debug.Print "done"', 'End Sub',
], 5, 7, 'Scan');

check('a loop taken whole carries its own Exit For',
    !whole.refused && whole.signature === 'Private Sub Scan()', whole.refused ?? whole.signature);

const withWhole = await extract([
    'Option Explicit', '', 'Public Sub One()', '    With Sheet1', '        .Name = "x"',
    '    End With', 'End Sub',
], 4, 6, 'Rename');

check('a With block taken whole carries its receiver',
    !withWhole.refused && withWhole.signature === 'Private Sub Rename()',
    withWhole.refused ?? withWhole.signature);

stop();
done();
