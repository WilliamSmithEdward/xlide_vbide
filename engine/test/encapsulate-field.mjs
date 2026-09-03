/*
 * Encapsulate Field through the real engine: a public variable behind a property pair.
 *
 * Read as TEXT, because the whole value of this refactoring is that nothing which used the field
 * has to change - the property keeps its name. A rewrite that renamed the field, or emitted a Let
 * where VBA wants a Set, would answer `did: true` and leave a class that does not compile.
 *
 *   node test/encapsulate-field.mjs
 *   node test/encapsulate-field.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('encapsulate-field');
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

async function encapsulate(lines, fieldName, { type = 'class' } = {}) {
    const source = lines.join(CRLF);
    opened += 1;
    const projectId = `p${opened}`;

    await call('project/open', {
        projectId, generation: 1,
        modules: [{ moduleName: 'Store', source, type }],
    });

    const answer = await call('textDocument/encapsulateField', {
        projectId, moduleName: 'Store', source, moduleType: type, fieldName,
    });

    return { ...answer, lines: (answer.source ?? '').split(CRLF) };
}

const has = (answer, text) => answer.lines.some((one) => one.trim() === text);

/* ---- a value field ---------------------------------------------------------------------------- */

const value = await encapsulate([
    'Option Explicit',
    '',
    'Public Name As String',
    'Private m_Count As Long',
    '',
    'Public Sub Announce()',
    '    Debug.Print Name',
    'End Sub',
], 'Name');

check('a public field is encapsulated', !value.refused, value.refused ?? value.accessors?.join(' / '));
check('the variable becomes private, under a new name',
    has(value, 'Private m_Name As String') && !has(value, 'Public Name As String'),
    value.lines.find((one) => one.includes('m_Name')));
check('the declaration stays where it was, above the procedures',
    value.lines.indexOf('Private m_Name As String')
    < value.lines.findIndex((one) => one.startsWith('Public Sub')),
    `${value.lines.indexOf('Private m_Name As String')} against ${value.lines.findIndex((one) => one.startsWith('Public Sub'))}`);
check('a Property Get returns it',
    has(value, 'Public Property Get Name() As String') && has(value, 'Name = m_Name'));
check('a Property Let writes it, because a String is a value',
    has(value, 'Public Property Let Name(ByVal RHS As String)') && has(value, 'm_Name = RHS'));
check('nothing that used the field had to change',
    has(value, 'Debug.Print Name'), value.lines.find((one) => one.includes('Debug.Print')));
check('the answer names both accessors',
    (value.accessors ?? []).join(' / ') === 'Property Get Name / Property Let Name',
    (value.accessors ?? []).join(' / '));
check('and the backing variable it made',
    value.backingField === 'm_Name' && value.field === 'Name',
    `${value.field} -> ${value.backingField}`);

/* ---- an object field, where VBA wants Set ------------------------------------------------------ */

const object = await encapsulate([
    'Option Explicit', '',
    'Public Log As Object',
], 'Log');

check('an Object field gets a Property Set, not a Let',
    has(object, 'Public Property Set Log(ByVal RHS As Object)') && !has(object, 'Public Property Let Log(ByVal RHS As Object)'),
    (object.accessors ?? []).join(' / ') || object.refused);
check('and both sides of it assign with Set',
    has(object, 'Set Log = m_Log') && has(object, 'Set m_Log = RHS'),
    object.lines.filter((one) => one.trim().startsWith('Set ')).map((one) => one.trim()).join(' | '));

/* ---- a type suffix and a bare declaration ------------------------------------------------------ */

const suffix = await encapsulate(['Option Explicit', '', 'Public Total&'], 'Total');
check('a type suffix is read as its type',
    has(suffix, 'Private m_Total As Long') && has(suffix, 'Public Property Get Total() As Long'),
    suffix.refused ?? suffix.lines.find((one) => one.includes('m_Total')));

const bare = await encapsulate(['Option Explicit', '', 'Public Anything'], 'Anything');
check('a declaration with no type at all is a Variant',
    has(bare, 'Private m_Anything As Variant') && has(bare, 'Public Property Let Anything(ByVal RHS As Variant)'),
    bare.refused ?? bare.lines.find((one) => one.includes('m_Anything')));

const friend = await encapsulate(['Option Explicit', '', 'Friend Shared As Long'], 'Shared');
check('a Friend field keeps its visibility',
    has(friend, 'Friend Property Get Shared() As Long') && has(friend, 'Friend Property Let Shared(ByVal RHS As Long)'),
    friend.refused ?? (friend.accessors ?? []).join(' / '));

/* ---- the refusals ------------------------------------------------------------------------------ */

console.log('\nrefusals:\n');

const refusals = [
    ['a name the module does not declare',
        ['Option Explicit', '', 'Public Name As String'], 'Nothing', /declares no module-level variable/i],
    ['a field that is already Private',
        ['Option Explicit', '', 'Private Name As String'], 'Name', /already Private/i],
    ['a Const',
        ['Option Explicit', '', 'Public Const Limit As Long = 3'], 'Limit', /is a Const/i],
    ['a WithEvents field',
        ['Option Explicit', '', 'Public WithEvents Book As Workbook'], 'Book', /WithEvents/i],
    ['an array',
        ['Option Explicit', '', 'Public Values() As String'], 'Values', /is an array/i],
    ['a declaration shared with another name',
        ['Option Explicit', '', 'Public First As Long, Second As Long'], 'First', /shares its declaration with 'Second'/i],
    ['a name a procedure already has',
        ['Option Explicit', '', 'Public Name As String', '', 'Public Sub Name()', 'End Sub'], 'Name', /already declares/i],
    ['a backing name that is taken',
        ['Option Explicit', '', 'Public Name As String', 'Private m_Name As Long'], 'Name', /already declares 'm_Name'/i],
];

for (const [what, lines, field, expected] of refusals) {
    const answer = await encapsulate(lines, field);
    check(`refused: ${what}`,
        typeof answer.refused === 'string' && expected.test(answer.refused) && !answer.source,
        answer.refused ?? `NOT REFUSED (${(answer.accessors ?? []).join(' / ')})`);
}

/* ---- and the result analyses ------------------------------------------------------------------- */

console.log('\nthe result is code the analyzer accepts:\n');

await call('project/open', {
    projectId: 'clean', generation: 1,
    modules: [{ moduleName: 'Store', source: value.source, type: 'class' }],
});
const findings = await call('textDocument/diagnostics', {
    documentKey: 'clean/Store', projectId: 'clean', generation: 1,
    source: value.source, moduleName: 'Store', moduleType: 'class',
});
check('the encapsulated class reports nothing',
    (findings.diagnostics ?? []).length === 0,
    (findings.diagnostics ?? []).map((one) => `${one.code}@${one.at.startLine}`).join(', ') || '(none)');

stop();
done();
