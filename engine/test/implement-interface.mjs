/*
 * Implement Interface through the real engine: what a class must carry, written out for it.
 *
 * The stubs are read as TEXT. A signature that does not match the interface's does not compile,
 * and VBA's message names the member rather than the difference, so "did it emit the right
 * parameters" is the whole question and a did/didn't flag cannot answer it.
 *
 *   node test/implement-interface.mjs
 *   node test/implement-interface.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('implement-interface');
const CRLF = '\r\n';

let passed = 0;
const failures = [];
const check = (what, ok, detail) => {
    console.log(`  ${ok ? 'ok  ' : 'FAIL'} ${what}${detail !== undefined ? `  -- ${String(detail).slice(0, 180)}` : ''}`);
    if (ok) { passed += 1; } else { failures.push(what); }
};
const done = () => {
    console.log(`\n${passed} passed, ${failures.length} failed`);
    for (const one of failures) { console.log(`  ${one}`); }
    process.exit(failures.length === 0 ? 0 : 1);
};

await call('initialize', {});

let opened = 0;

/** A two-module project - the interface and the class - with the stubs written into the class. */
async function implement(iface, impl, { interfaceName, ifaceName = 'IStore', implName = 'Store' } = {}) {
    const interfaceSource = iface.join(CRLF);
    const source = impl.join(CRLF);
    opened += 1;
    const projectId = `p${opened}`;

    await call('project/open', {
        projectId, generation: 1,
        modules: [
            { moduleName: ifaceName, source: interfaceSource, type: 'class' },
            { moduleName: implName, source, type: 'class' },
        ],
    });

    const answer = await call('textDocument/implementInterface', {
        projectId, moduleName: implName, source, moduleType: 'class',
        ...(interfaceName ? { interfaceName } : {}),
    });

    return { ...answer, lines: (answer.source ?? '').split(CRLF) };
}

const has = (answer, text) => answer.lines.some((one) => one.trim() === text);
const headers = (answer) => answer.lines.filter((one) => /^Private (Sub|Function|Property)/.test(one));

/* ---- every member kind an interface can declare ---------------------------------------------- */

const ALL = [
    'Option Explicit', '',
    'Public Count As Long',
    'Public Log As Object', '',
    'Public Sub Save()', 'End Sub', '',
    'Public Function Total(ByVal rate As Double, Optional ByVal since As Date) As Currency',
    'End Function', '',
    'Public Property Get Name() As String', 'End Property', '',
    'Public Property Let Name(ByVal value As String)', 'End Property', '',
    'Private Sub Helper()', 'End Sub',
];

const all = await implement(ALL, [
    'Option Explicit', '', 'Implements IStore', '',
]);

check('the class is given its members', !all.refused, all.refused ?? headers(all).length + ' stub(s)');
check('a Sub becomes a Private Sub',
    has(all, 'Private Sub IStore_Save()'), headers(all).join(' | '));
check("a Function keeps the interface's parameters, word for word",
    has(all, 'Private Function IStore_Total(ByVal rate As Double, Optional ByVal since As Date) As Currency'),
    all.lines.find((one) => one.includes('IStore_Total')));
check('a Property Get keeps its return type',
    has(all, 'Private Property Get IStore_Name() As String'));
check('a Property Let keeps its value parameter',
    has(all, 'Private Property Let IStore_Name(ByVal value As String)'));
check('a public field becomes a Get and a Let',
    has(all, 'Private Property Get IStore_Count() As Long')
    && has(all, 'Private Property Let IStore_Count(ByVal RHS As Long)'),
    headers(all).filter((one) => one.includes('Count')).join(' | '));
check('a public OBJECT field becomes a Get and a Set, because VBA assigns one with Set',
    has(all, 'Private Property Get IStore_Log() As Object')
    && has(all, 'Private Property Set IStore_Log(ByVal RHS As Object)'),
    headers(all).filter((one) => one.includes('Log')).join(' | '));
check('a Private member of the interface is not required',
    !all.lines.some((one) => one.includes('IStore_Helper')), headers(all).join(' | '));
check('every stub raises rather than silently doing nothing',
    all.lines.filter((one) => one.trim().startsWith('Err.Raise 5')).length === headers(all).length,
    `${all.lines.filter((one) => one.trim().startsWith('Err.Raise 5')).length} raise(s) for ${headers(all).length} stub(s)`);
check('and each closes with the right keyword',
    has(all, 'End Sub') && has(all, 'End Function') && has(all, 'End Property'));
check('the answer names what it wrote',
    (all.added ?? []).length === headers(all).length && (all.interfaces ?? []).includes('IStore'),
    `${(all.added ?? []).join(', ')} for ${(all.interfaces ?? []).join(', ')}`);

/* ---- what is already there is left alone ------------------------------------------------------ */

const partly = await implement(ALL, [
    'Option Explicit', '', 'Implements IStore', '',
    'Private Sub IStore_Save()',
    '    Debug.Print "saved"',
    'End Sub',
]);

check('a member already written is not written again',
    partly.lines.filter((one) => one.includes('Private Sub IStore_Save(')).length === 1,
    `${partly.lines.filter((one) => one.includes('Private Sub IStore_Save(')).length} copies`);
check('and its body is untouched',
    has(partly, 'Debug.Print "saved"'));
check('the rest are still written',
    (partly.added ?? []).length === (all.added ?? []).length - 1,
    (partly.added ?? []).join(', '));

/* ---- two interfaces --------------------------------------------------------------------------- */

const both = await call('project/open', {
    projectId: 'two', generation: 1,
    modules: [
        { moduleName: 'IStore', source: ['Option Explicit', '', 'Public Sub Save()', 'End Sub'].join(CRLF), type: 'class' },
        { moduleName: 'ILog', source: ['Option Explicit', '', 'Public Sub Write(ByVal text As String)', 'End Sub'].join(CRLF), type: 'class' },
        { moduleName: 'Store', source: ['Option Explicit', '', 'Implements IStore', 'Implements ILog'].join(CRLF), type: 'class' },
    ],
});
void both;

const two = await call('textDocument/implementInterface', {
    projectId: 'two', moduleName: 'Store', moduleType: 'class',
    source: ['Option Explicit', '', 'Implements IStore', 'Implements ILog'].join(CRLF),
});
check('with no name given, every interface is answered',
    (two.interfaces ?? []).join(',') === 'IStore,ILog'
    && (two.added ?? []).join(',') === 'IStore_Save,ILog_Write',
    `${(two.interfaces ?? []).join(',')} -> ${(two.added ?? []).join(',')}`);

const one = await call('textDocument/implementInterface', {
    projectId: 'two', moduleName: 'Store', moduleType: 'class', interfaceName: 'ILog',
    source: ['Option Explicit', '', 'Implements IStore', 'Implements ILog'].join(CRLF),
});
check('and a name narrows it to that one',
    (one.added ?? []).join(',') === 'ILog_Write', (one.added ?? []).join(','));

/* ---- the refusals ------------------------------------------------------------------------------ */

console.log('\nrefusals:\n');

const noImplements = await implement(ALL, ['Option Explicit', '', 'Public Sub Go()', 'End Sub']);
check('refused: a class that implements nothing',
    /does not declare Implements/i.test(String(noImplements.refused)) && !noImplements.source,
    noImplements.refused);

const wrongName = await implement(ALL, ['Option Explicit', '', 'Implements IStore'], { interfaceName: 'IOther' });
check('refused: an interface the class does not implement',
    /does not implement 'IOther'/i.test(String(wrongName.refused)) && /IStore/.test(String(wrongName.refused)),
    wrongName.refused);

const missing = await call('textDocument/implementInterface', {
    projectId: 'gone', moduleName: 'Store', moduleType: 'class',
    source: ['Option Explicit', '', 'Implements INowhere'].join(CRLF),
});
check('refused: an interface no module in the project declares',
    /no module called 'INowhere'/i.test(String(missing.refused)), missing.refused);

const already = await implement(
    ['Option Explicit', '', 'Public Sub Save()', 'End Sub'],
    ['Option Explicit', '', 'Implements IStore', '', 'Private Sub IStore_Save()', 'End Sub']);
check('refused: a class that already implements every member',
    /already implements every member/i.test(String(already.refused)) && !already.source,
    already.refused);

const empty = await implement(
    ['Option Explicit', '', 'Private Sub Hidden()', 'End Sub'],
    ['Option Explicit', '', 'Implements IStore']);
check('refused: an interface with no public members',
    /declares no public members/i.test(String(empty.refused)), empty.refused);

/* ---- and what it produced still analyses ------------------------------------------------------- */

console.log('\nthe result is code the analyzer accepts:\n');

await call('project/open', {
    projectId: 'clean', generation: 1,
    modules: [
        { moduleName: 'IStore', source: ALL.join(CRLF), type: 'class' },
        { moduleName: 'Store', source: all.source, type: 'class' },
    ],
});
const findings = await call('textDocument/diagnostics', {
    documentKey: 'clean/Store', projectId: 'clean', generation: 1,
    source: all.source, moduleName: 'Store', moduleType: 'class',
});
check('the class the stubs were written into reports nothing',
    (findings.diagnostics ?? []).length === 0,
    (findings.diagnostics ?? []).map((one) => `${one.code}@${one.at.startLine}`).join(', ') || '(none)');

stop();
done();
