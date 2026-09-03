/*
 * Introduce Parameter through the real engine: the signature grows, and every caller is given the
 * value the local used to be assigned.
 *
 * THE VALUE HAS TO TRAVEL. A local's initialiser is written in the procedure's own scope, so an
 * expression naming another local means nothing at a call site in another module. That check is
 * what most of the refusals below are about, and it is the difference between a refactoring and a
 * project that stops compiling somewhere the developer was not looking.
 *
 *   node test/introduce-parameter.mjs
 *   node test/introduce-parameter.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('introduce-parameter');
const CRLF = '\r\n';

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
    console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 190)}` : ''}`);
    if (ok) { passed += 1; } else { failures.push(what); }
};
const done = () => {
    console.log(`\n${passed} passed, ${failures.length} failed`);
    for (const one of failures) { console.log(`  ${one}`); }
    process.exit(failures.length === 0 ? 0 : 1);
};

await call('initialize', {});

let opened = 0;

/** A project of named modules, with one local of the first turned into a parameter. */
async function introduce(modules, localName, { from = 'Source' } = {}) {
    opened += 1;
    const projectId = `p${opened}`;
    const seeded = Object.entries(modules).map(([moduleName, lines]) => ({
        moduleName, source: lines.join(CRLF), type: moduleName === 'Book' ? 'document' : 'standard',
    }));

    await call('project/open', { projectId, generation: 1, host: 'excel', modules: seeded });

    const source = seeded.find((one) => one.moduleName === from).source;
    const at = source.indexOf(` ${localName} `);
    const answer = await call('workspace/introduceParameter', {
        projectId, moduleName: from, source, moduleType: 'standard', offset: at + 1,
    });

    const byModule = new Map((answer.modules ?? []).map((one) => [one.module, one.source.split(CRLF)]));
    return { ...answer, byModule };
}

const has = (lines, text) => (lines ?? []).some((one) => one.trim() === text);

/* ---- a literal, with callers in two modules ----------------------------------------------------- */

const made = await introduce({
    Source: [
        'Option Explicit',
        '',
        'Public Sub Post(ByVal label As String)',
        '    Dim rate As Double',
        '    rate = 1.2',
        '    Debug.Print label, rate',
        'End Sub',
        '',
        'Public Sub Nearby()',
        '    Post "here"',
        'End Sub',
    ],
    Caller: [
        'Option Explicit',
        '',
        'Public Sub Far()',
        '    Post "there"',
        'End Sub',
    ],
}, 'rate');

check('the local becomes a parameter', !made.refused,
    made.refused ?? `${made.parameter} As ${made.type} = ${made.value}`);
check('the signature grew, at the end where the argument goes',
    has(made.byModule.get('Source'), 'Public Sub Post(ByVal label As String, ByVal rate As Double)'),
    (made.byModule.get('Source') ?? []).find((one) => one.includes('Sub Post')));
check('the declaration and the assignment are gone',
    !(made.byModule.get('Source') ?? []).some((one) => one.includes('Dim rate'))
    && !(made.byModule.get('Source') ?? []).some((one) => one.trim().startsWith('rate =')),
    (made.byModule.get('Source') ?? []).map((one) => one.trim()).filter(Boolean).join(' | '));
check('the body still reads it, now as the parameter',
    has(made.byModule.get('Source'), 'Debug.Print label, rate'));
check('a caller in the same module passes the value',
    has(made.byModule.get('Source'), 'Post "here", 1.2'),
    (made.byModule.get('Source') ?? []).find((one) => one.includes('"here"')));
check('and so does one in another module',
    has(made.byModule.get('Caller'), 'Post "there", 1.2'),
    (made.byModule.get('Caller') ?? []).find((one) => one.includes('"there"')));
check('both call sites are counted', made.callSites === 2, String(made.callSites));

/* ---- a parenthesised call, and one with no arguments at all ------------------------------------- */

const shapes = await introduce({
    Source: [
        'Option Explicit',
        '',
        'Public Function Total() As Long',
        '    Dim seed As Long',
        '    seed = 3',
        '    Total = seed',
        'End Function',
        '',
        'Public Sub Use()',
        '    Debug.Print Total()',
        'End Sub',
    ],
}, 'seed');

check('a call with parentheses takes the argument inside them',
    has(shapes.byModule.get('Source'), 'Debug.Print Total(3)'),
    (shapes.byModule.get('Source') ?? []).find((one) => one.includes('Debug.Print')));
check('and the empty parameter list becomes one parameter',
    has(shapes.byModule.get('Source'), 'Public Function Total(ByVal seed As Long) As Long'),
    (shapes.byModule.get('Source') ?? []).find((one) => one.includes('Function Total')));

/* ---- the refusals --------------------------------------------------------------------------------- */

console.log('\nrefusals:\n');

const stranded = await introduce({
    Source: [
        'Option Explicit', '',
        'Public Sub Post()',
        '    Dim basePrice As Double', '    basePrice = 10',
        '    Dim rate As Double', '    rate = basePrice * 1.2',
        '    Debug.Print rate',
        'End Sub',
    ],
}, 'rate');
check('refused: the value names a local, which no caller can see',
    /'baseprice'/i.test(String(stranded.refused)) && stranded.modules.length === 0, stranded.refused);

const privateState = await introduce({
    Source: [
        'Option Explicit', '', 'Private held As Double', '',
        'Public Sub Post()', '    Dim rate As Double', '    rate = held',
        '    Debug.Print rate', 'End Sub',
    ],
}, 'rate');
check('refused: the value names something Private to the module',
    /'held'/i.test(String(privateState.refused)) && privateState.modules.length === 0, privateState.refused);

const twice = await introduce({
    Source: [
        'Option Explicit', '', 'Public Sub Post()',
        '    Dim rate As Double', '    rate = 1', '    rate = 2', '    Debug.Print rate', 'End Sub',
    ],
}, 'rate');
check('refused: assigned more than once',
    /assigned 2 times/i.test(String(twice.refused)), twice.refused);

const never = await introduce({
    Source: [
        'Option Explicit', '', 'Public Sub Post()',
        '    Dim rate As Double', '    Debug.Print rate', 'End Sub',
    ],
}, 'rate');
check('refused: never assigned, so there is no value to pass',
    /never assigned/i.test(String(never.refused)), never.refused);

const handler = await introduce({
    Source: ['Option Explicit', '', 'Public Sub Book_Open()',
        '    Dim rate As Double', '    rate = 1', '    Debug.Print rate', 'End Sub'],
    Book: ['Option Explicit'],
}, 'rate');
check('refused: an event handler, whose signature is not this project\'s to change',
    /bound by its name/i.test(String(handler.refused)), handler.refused);

const already = await introduce({
    Source: ['Option Explicit', '', 'Public Sub Post(ByVal rate As Double)',
        '    Debug.Print rate', 'End Sub'],
}, 'rate');
check('refused: already a parameter',
    /already a parameter/i.test(String(already.refused)), already.refused);

const constant = await introduce({
    Source: ['Option Explicit', '', 'Public Sub Post()',
        '    Const rate As Double = 1.2', '    Debug.Print rate', 'End Sub'],
}, 'rate');
check('refused: a Const, whose value is fixed at compile time',
    /is a Const/i.test(String(constant.refused)), constant.refused);

/* ---- and what it produced analyses ------------------------------------------------------------------ */

console.log('\nthe result is code the analyzer accepts:\n');

await call('project/open', {
    projectId: 'clean', generation: 1, host: 'excel',
    modules: made.modules.map((one) => ({ moduleName: one.module, source: one.source, type: 'standard' })),
});

let findings = [];
for (const one of made.modules) {
    const answer = await call('textDocument/diagnostics', {
        documentKey: `clean/${one.module}`, projectId: 'clean', generation: 1,
        source: one.source, moduleName: one.module, moduleType: 'standard',
    });
    findings = findings.concat((answer.diagnostics ?? []).map((d) => `${one.module}:${d.code}`));
}

check('every module it touched reports nothing', findings.length === 0, findings.join(', ') || '(none)');

stop();
done();
