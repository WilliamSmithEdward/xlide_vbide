/*
 * Move to Module through the real engine: what goes, what follows it, and what stops it.
 *
 * The move is the easy half. What this suite is really about is the refusals - a procedure that
 * touches something Private to the module it is leaving cannot go, because nothing outside that
 * module can reach it, and a tool that moved it anyway would leave a project that does not compile
 * and a developer with no idea why.
 *
 *   node test/move-to-module.mjs
 *   node test/move-to-module.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('move-to-module');
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

/** A project of named modules, with one procedure moved out of the first. */
async function move(modules, procedureName, targetModule, { from = 'Source' } = {}) {
    opened += 1;
    const projectId = `p${opened}`;
    const seeded = Object.entries(modules).map(([moduleName, lines]) => ({
        moduleName, source: lines.join(CRLF), type: moduleName === 'Book' ? 'document' : 'standard',
    }));

    await call('project/open', { projectId, generation: 1, host: 'excel', modules: seeded });

    const source = seeded.find((one) => one.moduleName === from).source;
    const answer = await call('workspace/moveToModule', {
        projectId, moduleName: from, source, moduleType: 'standard',
        offset: source.indexOf(procedureName), targetModule,
    });

    const byModule = new Map((answer.modules ?? []).map((one) => [one.module, one.source.split(CRLF)]));
    return { ...answer, byModule };
}

const has = (lines, text) => (lines ?? []).some((one) => one.trim() === text);

/* ---- a clean move ------------------------------------------------------------------------------- */

const clean = await move({
    Source: [
        'Option Explicit',
        '',
        'Public Sub Stay()',
        '    Debug.Print "here"',
        'End Sub',
        '',
        'Public Sub Go(ByVal n As Long)',
        '    Debug.Print n * 2',
        'End Sub',
    ],
    Target: ['Option Explicit', '', 'Public Sub Already()', 'End Sub'],
}, 'Go', 'Target');

check('the procedure moves', !clean.refused, clean.refused ?? `${clean.moved}: ${clean.from} -> ${clean.to}`);
check('it is gone from the module it left',
    !(clean.byModule.get('Source') ?? []).some((one) => one.includes('Sub Go(')),
    (clean.byModule.get('Source') ?? []).map((one) => one.trim()).filter(Boolean).join(' | '));
check('and whole in the one it arrived at',
    has(clean.byModule.get('Target'), 'Public Sub Go(ByVal n As Long)')
    && has(clean.byModule.get('Target'), 'Debug.Print n * 2')
    && has(clean.byModule.get('Target'), 'End Sub'),
    (clean.byModule.get('Target') ?? []).map((one) => one.trim()).filter(Boolean).join(' | '));
check('what stayed behind is untouched',
    has(clean.byModule.get('Source'), 'Public Sub Stay()'));
check('and the module it arrived at kept what it had',
    has(clean.byModule.get('Target'), 'Public Sub Already()'));
check('two modules are rewritten and no more',
    clean.modules.length === 2, clean.modules.map((one) => one.module).join(', '));

/* ---- a qualified call site follows the move ------------------------------------------------------ */

const qualified = await move({
    Source: ['Option Explicit', '', 'Public Sub Go()', '    Debug.Print 1', 'End Sub'],
    Target: ['Option Explicit'],
    Caller: [
        'Option Explicit',
        '',
        'Public Sub Use()',
        '    Source.Go',
        '    Go',
        'End Sub',
    ],
}, 'Go', 'Target');

check('a qualified call is repointed at the new module',
    has(qualified.byModule.get('Caller'), 'Target.Go'),
    (qualified.byModule.get('Caller') ?? []).map((one) => one.trim()).filter(Boolean).join(' | '));
check('and an unqualified one is left alone, because VBA still finds it',
    has(qualified.byModule.get('Caller'), 'Go'),
    (qualified.byModule.get('Caller') ?? []).map((one) => one.trim()).filter(Boolean).join(' | '));
check('the answer counts what it repointed', qualified.requalified === 1, String(qualified.requalified));

/* ---- the refusals --------------------------------------------------------------------------------- */

console.log('\nrefusals:\n');

const privateVariable = await move({
    Source: [
        'Option Explicit', '', 'Private held As Long', '',
        'Public Sub Go()', '    Debug.Print held', 'End Sub',
    ],
    Target: ['Option Explicit'],
}, 'Go', 'Target');
check('refused: it uses a Private variable of the module it would leave',
    /'held'/.test(String(privateVariable.refused)) && /Private to 'Source'/.test(String(privateVariable.refused))
    && privateVariable.modules.length === 0,
    privateVariable.refused);

const privateProcedure = await move({
    Source: [
        'Option Explicit', '',
        'Private Sub Helper()', 'End Sub', '',
        'Public Sub Go()', '    Helper', 'End Sub',
    ],
    Target: ['Option Explicit'],
}, 'Go', 'Target');
check('refused: it calls a Private procedure of that module',
    /'Helper'/.test(String(privateProcedure.refused)) && privateProcedure.modules.length === 0,
    privateProcedure.refused);

const publicVariable = await move({
    Source: [
        'Option Explicit', '', 'Public shared As Long', '',
        'Public Sub Go()', '    Debug.Print shared', 'End Sub',
    ],
    Target: ['Option Explicit'],
}, 'Go', 'Target');
check('but a PUBLIC one is fine, because the project can still see it',
    !publicVariable.refused && has(publicVariable.byModule.get('Target'), 'Debug.Print shared'),
    publicVariable.refused ?? 'moved');

const missing = await move({
    Source: ['Option Explicit', '', 'Public Sub Go()', 'End Sub'],
    Target: ['Option Explicit'],
}, 'Go', 'Nowhere');
check('refused: a module the project does not have',
    /no module called 'Nowhere'/.test(String(missing.refused)), missing.refused);

const taken = await move({
    Source: ['Option Explicit', '', 'Public Sub Go()', 'End Sub'],
    Target: ['Option Explicit', '', 'Public Sub Go()', 'End Sub'],
}, 'Go', 'Target');
check('refused: the target already declares that name',
    /already declares a procedure called 'Go'/.test(String(taken.refused)), taken.refused);

const itself = await move({
    Source: ['Option Explicit', '', 'Public Sub Go()', 'End Sub'],
    Target: ['Option Explicit'],
}, 'Go', 'Source');
check('refused: moving it where it already is',
    /already in 'Source'/.test(String(itself.refused)), itself.refused);

const privateModule = await move({
    Source: ['Option Explicit', 'Option Private Module', '', 'Public Sub Go()', 'End Sub'],
    Target: ['Option Explicit'],
}, 'Go', 'Target');
check('refused: Option Private Module, which changes what Public means',
    /Option Private Module/.test(String(privateModule.refused)), privateModule.refused);

const handler = await move({
    Source: ['Option Explicit', '', 'Public Sub Book_Open()', 'End Sub'],
    Target: ['Option Explicit'],
    Book: ['Option Explicit'],
}, 'Book_Open', 'Target');
check('refused: an event handler, which is bound by its name',
    /event handler/.test(String(handler.refused)), handler.refused);

/* ---- and what it produced analyses ------------------------------------------------------------------ */

console.log('\nthe result is code the analyzer accepts:\n');

await call('project/open', {
    projectId: 'clean', generation: 1, host: 'excel',
    modules: clean.modules.map((one) => ({ moduleName: one.module, source: one.source, type: 'standard' })),
});

let findings = [];
for (const one of clean.modules) {
    const answer = await call('textDocument/diagnostics', {
        documentKey: `clean/${one.module}`, projectId: 'clean', generation: 1,
        source: one.source, moduleName: one.module, moduleType: 'standard',
    });
    findings = findings.concat((answer.diagnostics ?? []).map((d) => `${one.module}:${d.code}`));
}

check('both modules report nothing afterwards', findings.length === 0, findings.join(', ') || '(none)');

stop();
done();
