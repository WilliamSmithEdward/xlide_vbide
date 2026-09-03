/*
 * Inline Variable through the real engine, and the two VBA traps it declines rather than springs.
 *
 * BRACKETS. `Foo x` passes by reference and `Foo (x)` passes by value, so an inliner that
 * parenthesises for precedence changes how a call binds. EVALUATION COUNT. A variable read three
 * times becomes three evaluations of its initialiser. Both are settled by inlining only an ATOMIC
 * initialiser - a literal, or a name with its member chain - and the refusals below are where that
 * rule is visible.
 *
 *   node test/inline-variable.mjs
 *   node test/inline-variable.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('inline-variable');
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

/** Inlines the name at its declaration and hands the module back whole. */
async function inline(lines, name) {
    const source = lines.join(CRLF);
    opened += 1;
    const projectId = `p${opened}`;

    await call('project/open', {
        projectId, generation: 1, host: 'excel',
        modules: [{ moduleName: 'Probe', source, type: 'standard' }],
    });

    // The caret on the declaration, which is where a developer puts it.
    const at = source.indexOf(` ${name} `, source.indexOf('Dim') >= 0 ? source.indexOf('Dim') : 0);
    const answer = await call('textDocument/inlineVariable', {
        projectId, moduleName: 'Probe', source, moduleType: 'standard',
        offset: at + 1,
    });

    return { ...answer, lines: (answer.source ?? '').split(CRLF) };
}

const has = (answer, text) => answer.lines.some((one) => one.trim() === text);

/* ---- a literal, read more than once ------------------------------------------------------------ */

const literal = await inline([
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Dim limit As Long',
    '    limit = 10',
    '    Debug.Print limit',
    '    Debug.Print limit * 2',
    'End Sub',
], 'limit');

check('a literal is inlined', !literal.refused, literal.refused ?? `${literal.variable} -> ${literal.value}`);
check('every use is replaced',
    has(literal, 'Debug.Print 10') && has(literal, 'Debug.Print 10 * 2') && literal.replaced === 2,
    `${literal.replaced} use(s)`);
check('the declaration is gone',
    !literal.lines.some((one) => one.includes('Dim limit')), literal.lines.map((one) => one.trim()).join(' | '));
check('and so is the assignment',
    !literal.lines.some((one) => one.includes('limit =')), literal.lines.map((one) => one.trim()).join(' | '));

/* ---- an object, which is the case people actually reach for ------------------------------------- */

const object = await inline([
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Dim ws As Object',
    '    Set ws = Sheet1',
    '    Debug.Print ws.Name',
    '    Debug.Print ws.Index',
    'End Sub',
], 'ws');

check('a Set assignment of a plain name is inlined',
    !object.refused && object.value === 'Sheet1', object.refused ?? object.value);
check('and its member chains come out qualified by the name',
    has(object, 'Debug.Print Sheet1.Name') && has(object, 'Debug.Print Sheet1.Index'),
    object.lines.filter((one) => one.includes('Debug.Print')).map((one) => one.trim()).join(' | '));

/* ---- a member chain is atomic too --------------------------------------------------------------- */

const chain = await inline([
    'Option Explicit',
    '',
    'Public Sub Post()',
    '    Dim who As String',
    '    who = Application.UserName',
    '    Debug.Print who',
    'End Sub',
], 'who');

check('a member chain with no call is atomic',
    !chain.refused && has(chain, 'Debug.Print Application.UserName'),
    chain.refused ?? chain.value);

/* ---- the refusals -------------------------------------------------------------------------------- */

console.log('\nrefusals:\n');

const arithmetic = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Dim total As Long', '    total = 2 + 3', '    Debug.Print total', 'End Sub',
], 'total');
check('refused: an expression that would need brackets',
    /brackets/i.test(String(arithmetic.refused)) && !arithmetic.source, arithmetic.refused);

const called = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Dim cell As Object', '    Set cell = Sheet1.Cells(1, 1)',
    '    Debug.Print cell.Value', '    Debug.Print cell.Row', 'End Sub',
], 'cell');
check('refused: a call, which would be made once per use',
    /not a literal or a plain name/i.test(String(called.refused)) && !called.source, called.refused);

const twice = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Dim n As Long', '    n = 1', '    n = 2', '    Debug.Print n', 'End Sub',
], 'n');
check('refused: assigned more than once',
    /assigned 2 times/i.test(String(twice.refused)) && !twice.source, twice.refused);

const readFirst = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Dim n As Long', '    Debug.Print n', '    n = 1', '    Debug.Print n', 'End Sub',
], 'n');
check('refused: read before it is assigned',
    /read before it is assigned/i.test(String(readFirst.refused)) && !readFirst.source, readFirst.refused);

const unread = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Dim n As Long', '    n = 1', 'End Sub',
], 'n');
check('refused: assigned and never read',
    /never read/i.test(String(unread.refused)) && !unread.source, unread.refused);

const parameter = await inline([
    'Option Explicit', '', 'Public Sub Post(ByVal seed As Long)',
    '    Debug.Print seed', 'End Sub',
], 'seed');
check('refused: a parameter, whose value is the caller\'s',
    /is a parameter/i.test(String(parameter.refused)) && !parameter.source, parameter.refused);

const constant = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Const limit As Long = 10', '    Debug.Print limit', 'End Sub',
], 'limit');
check('refused: a Const, which is already a name for a value',
    /is a Const/i.test(String(constant.refused)) && !constant.source, constant.refused);

const shared = await inline([
    'Option Explicit', '', 'Public Sub Post()',
    '    Dim a As Long, b As Long', '    a = 1', '    Debug.Print a + b', 'End Sub',
], 'a');
check('refused: a declaration shared with another name',
    /shares its declaration/i.test(String(shared.refused)) && !shared.source, shared.refused);

/* ---- and the result analyses ---------------------------------------------------------------------- */

console.log('\nthe result is code the analyzer accepts:\n');

await call('project/open', {
    projectId: 'clean', generation: 1, host: 'excel',
    modules: [{ moduleName: 'Probe', source: literal.source, type: 'standard' }],
});
const findings = await call('textDocument/diagnostics', {
    documentKey: 'clean/Probe', projectId: 'clean', generation: 1,
    source: literal.source, moduleName: 'Probe', moduleType: 'standard',
});
check('the module the variable left reports nothing',
    (findings.diagnostics ?? []).length === 0,
    (findings.diagnostics ?? []).map((one) => `${one.code}@${one.at.startLine}`).join(', ') || '(none)');

stop();
done();
