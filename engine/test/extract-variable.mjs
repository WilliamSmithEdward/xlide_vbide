/*
 * Extract Variable through the real engine: the type and the Set both come from the analyzer.
 *
 * The type is what makes this refactoring hard and the `Set` is what makes it dangerous: VBA has
 * no assignment form that works for both an object and a value, so a wrong answer emits a line
 * that does not compile. Both come from `resolveExpressionType` (xlide_vscode#61) rather than from
 * a rule of thumb here, and both halves are read as TEXT below.
 *
 *   node test/extract-variable.mjs
 *   node test/extract-variable.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('extract-variable');
const CRLF = '\r\n';

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
    console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 170)}` : ''}`);
    if (ok) { passed += 1; } else { failures.push(what); }
};
const done = () => {
    console.log(`\n${passed} passed, ${failures.length} failed`);
    for (const one of failures) { console.log(`  ${one}`); }
    process.exit(failures.length === 0 ? 0 : 1);
};

await call('initialize', {});

let opened = 0;

/** Extracts the first occurrence of `expression` in the module and hands the result back whole. */
async function extract(lines, expression, newName = 'extracted', { at = 0, modules } = {}) {
    const source = lines.join(CRLF);
    opened += 1;
    const projectId = `p${opened}`;

    await call('project/open', {
        projectId, generation: 1, host: 'excel',
        modules: modules ?? [{ moduleName: 'Probe', source, type: 'standard' }],
    });

    let start = -1;
    for (let n = 0; n <= at; n += 1) {
        start = source.indexOf(expression, start + 1);
    }

    const answer = await call('textDocument/extractVariable', {
        projectId, moduleName: 'Probe', source, moduleType: 'standard',
        startOffset: start, endOffset: start + expression.length, newName,
    });

    return { ...answer, lines: (answer.source ?? '').split(CRLF) };
}

const has = (answer, text) => answer.lines.some((one) => one.trim() === text);

/* ---- a value expression ------------------------------------------------------------------------ */

const value = await extract([
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Dim total As Long',
    '    total = 2',
    '    Debug.Print total * 3',
    'End Sub',
], 'total * 3', 'scaled');

check('a value expression is extracted', !value.refused, value.refused ?? `${value.variable} As ${value.type}`);
// WHAT THE ANALYZER SAYS, not what this suite would have guessed. `total * 3` on a Long is a
// Double to the shared inference, which widens every numeric literal - right for checking
// compatibility, wider than a person would write on a declaration. It is safe (a Double holds any
// Long product), it is one answer rather than two, and a refactoring that overruled the analyzer
// on types would be a second type system living in a text rewriter.
check('the analyzer supplies the type',
    value.type === 'Double' && has(value, 'Dim scaled As Double'),
    value.lines.find((one) => one.includes('Dim scaled')));
check('it is assigned without Set, because it is not an object',
    has(value, 'scaled = total * 3') && value.isObject === false,
    value.lines.find((one) => one.includes('scaled =')));
check('and the selection is replaced by the name',
    has(value, 'Debug.Print scaled'), value.lines.find((one) => one.includes('Debug.Print')));
check('the declaration goes directly above the statement it came from',
    value.lines.indexOf('    Dim scaled As Double') + 2 === value.lines.indexOf('    Debug.Print scaled'),
    value.lines.map((one) => one.trim()).join(' | '));
check('and it keeps the indentation of the statement it came from',
    value.lines.includes('    Dim scaled As Double') && value.lines.includes('    scaled = total * 3'),
    value.lines.filter((one) => one.includes('scaled')).map((one) => JSON.stringify(one)).join(' '));

/* ---- an object expression, where VBA wants Set -------------------------------------------------- */

const object = await extract([
    'Option Explicit',
    '',
    'Public Sub Fill()',
    '    Debug.Print New Collection.Count',
    'End Sub',
], 'New Collection', 'items');

check('an object expression is assigned with Set',
    object.refused ? false : (object.isObject === true && has(object, 'Set items = New Collection')),
    object.refused ?? object.lines.find((one) => one.includes('items =')));
check('and declared as its own class rather than as Object',
    object.type === 'Collection' && has(object, 'Dim items As Collection'),
    object.refused ?? object.type);

/* ---- a string, and a call ----------------------------------------------------------------------- */

const text = await extract([
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Debug.Print "hello" & " world"',
    'End Sub',
], '"hello" & " world"', 'greeting');

check('a string expression is a String',
    text.type === 'String' && has(text, 'Dim greeting As String'),
    text.refused ?? text.type);

const called = await extract([
    'Option Explicit',
    '',
    'Private Function Rate() As Currency',
    '    Rate = 1',
    'End Function',
    '',
    'Public Sub Post()',
    '    Debug.Print Rate() * 2',
    'End Sub',
], 'Rate()', 'perUnit', { at: 1 });

check("a call's declared return type is used",
    called.type === 'Currency' && has(called, 'Dim perUnit As Currency'),
    called.refused ?? called.type);

/* ---- inside a block, the declaration stays inside it -------------------------------------------- */

const inBlock = await extract([
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Dim n As Long',
    '    n = 2',
    '    If n > 0 Then',
    '        Debug.Print n + 1',
    '    End If',
    'End Sub',
], 'n + 1', 'bumped');

check('a statement inside a block keeps its declaration inside the block',
    inBlock.lines.indexOf('        Dim bumped As Double') > inBlock.lines.findIndex((one) => one.includes('If n > 0'))
    && inBlock.lines.indexOf('        Dim bumped As Double') < inBlock.lines.findIndex((one) => one.includes('End If')),
    inBlock.refused ?? inBlock.lines.map((one) => one.trim()).join(' | '));

/* ---- the refusals ------------------------------------------------------------------------------- */

console.log('\nrefusals:\n');

const BASE = [
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Dim total As Long',
    '    total = 2',
    '    Debug.Print total * 3',
    'End Sub',
];

const partial = await extract(BASE, 'total *', 'half');
check('refused: half an expression',
    /not a whole expression/i.test(String(partial.refused)) && !partial.source, partial.refused);

const target = await extract(BASE, 'total', 'copy', { at: 1 });
check('refused: the name being assigned to',
    /is the name being assigned to/i.test(String(target.refused)) && !target.source, target.refused);

const outside = await extract(BASE, 'Option Explicit', 'nope');
check('refused: something outside a procedure',
    /inside a procedure/i.test(String(outside.refused)) && !outside.source, outside.refused);

const declaration = await extract(BASE, 'Long', 'kind');
check('refused: part of a declaration',
    typeof declaration.refused === 'string' && !declaration.source, declaration.refused);

const taken = await extract(BASE, 'total * 3', 'total');
check('refused: a name already declared here',
    /already declared here/i.test(String(taken.refused)) && !taken.source, taken.refused);

const keyword = await extract(BASE, 'total * 3', 'Next');
check('refused: a VBA keyword',
    /is a VBA keyword/i.test(String(keyword.refused)) && !keyword.source, keyword.refused);

const illegal = await extract(BASE, 'total * 3', '2fast');
check('refused: a name VBA would not accept',
    /is not a VBA name/i.test(String(illegal.refused)) && !illegal.source, illegal.refused);

/* ---- and the result analyses --------------------------------------------------------------------- */

console.log('\nthe result is code the analyzer accepts:\n');

await call('project/open', {
    projectId: 'clean', generation: 1, host: 'excel',
    modules: [{ moduleName: 'Probe', source: value.source, type: 'standard' }],
});
const findings = await call('textDocument/diagnostics', {
    documentKey: 'clean/Probe', projectId: 'clean', generation: 1,
    source: value.source, moduleName: 'Probe', moduleType: 'standard',
});
check('the module the variable went into reports nothing',
    (findings.diagnostics ?? []).length === 0,
    (findings.diagnostics ?? []).map((one) => `${one.code}@${one.at.startLine}`).join(', ') || '(none)');

stop();
done();
