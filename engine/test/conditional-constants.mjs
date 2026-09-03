/*
 * A project's own conditional compilation arguments, and what they decide.
 *
 * The analyzer knows the compiler's constants - VBA7, Win64, Mac - and drops the arms they settle.
 * A project's OWN constants live in the VBE's Project Properties box, and without them every
 * `#If MY_FLAG Then` is undecidable: both arms are analyzed, both contribute declarations, and a
 * finding can be reported from an arm the compiler never compiles.
 *
 * BOTH HALVES OR NEITHER. The symbols and the rules must be built under the same constants, or a
 * branch dropped from the symbols would still be analyzed - so this asks the question from both
 * sides: does a name declared only in a dead arm resolve, and does an error only in a dead arm
 * report.
 *
 *   node test/conditional-constants.mjs
 *   node test/conditional-constants.mjs --exe
 */

import { startEngine } from './harness.mjs';

const { call, stop } = await startEngine('conditional-constants');
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

/** One module analysed with, or without, a project's own constants. */
async function findings(lines, conditionalConstants) {
    const source = lines.join(CRLF);
    opened += 1;
    const projectId = `p${opened}`;

    await call('project/open', {
        projectId, generation: 1, host: 'excel',
        modules: [{ moduleName: 'Probe', source, type: 'standard' }],
        ...(conditionalConstants ? { conditionalConstants } : {}),
    });

    const answer = await call('textDocument/diagnostics', {
        documentKey: `${projectId}/Probe`, projectId, generation: 1,
        source, moduleName: 'Probe', moduleType: 'standard',
    });

    return (answer.diagnostics ?? []).map((one) => `${one.code}@${one.at.startLine}`);
}

/*
 * An error that exists ONLY in the arm a flag turns off. Under `MY_FLAG = 1` the compiler never
 * sees it, so neither should the analyzer; with no constants supplied both arms are live and it
 * must still report, or this suite would be measuring nothing.
 */
const ARMS = [
    'Option Explicit',
    '',
    'Public Sub Go()',
    '#If MY_FLAG Then',
    '    Debug.Print 1',
    '#Else',
    '    undeclaredInTheDeadArm = 2',
    '#End If',
    'End Sub',
];

const undecided = await findings(ARMS, undefined);
check('with no constants, both arms are analysed and the dead one reports',
    undecided.some((one) => one.startsWith('undeclared-variable')), undecided.join(', ') || '(none)');

const decided = await findings(ARMS, 'MY_FLAG = 1');
check('with the project\'s constant, the arm it does not compile is dropped',
    decided.length === 0, decided.join(', ') || '(none)');

const off = await findings(ARMS, 'MY_FLAG = 0');
check('and turning the flag off keeps the other arm live',
    off.some((one) => one.startsWith('undeclared-variable')), off.join(', ') || '(none)');

/* ---- the symbols are built under them too, not only the rules --------------------------------- */

// A name declared only in the arm that IS compiled must resolve; the same code under the other
// setting must not. That is the half a rules-only integration would get wrong.
const DECLARED = [
    'Option Explicit',
    '',
    '#If MY_FLAG Then',
    'Private held As Long',
    '#End If',
    '',
    'Public Sub Go()',
    '    held = 1',
    'End Sub',
];

const visible = await findings(DECLARED, 'MY_FLAG = 1');
check('a declaration in the live arm resolves', visible.length === 0, visible.join(', ') || '(none)');

const hidden = await findings(DECLARED, 'MY_FLAG = 0');
check('and the same declaration in a dead arm does not, so its use reports',
    hidden.some((one) => one.startsWith('undeclared-variable')), hidden.join(', ') || '(none)');

/* ---- the format the VBE writes ------------------------------------------------------------------ */

const pairs = await findings(ARMS, 'OTHER = 0 : MY_FLAG = 1 : THIRD = -1');
check('several arguments, colon separated, as the VBE property holds them',
    pairs.length === 0, pairs.join(', ') || '(none)');

const boolean = await findings([
    'Option Explicit', '', 'Public Sub Go()',
    '#If MY_FLAG Then', '    Debug.Print 1', '#Else', '    stillUndeclared = 2', '#End If',
    'End Sub',
], 'MY_FLAG = -1');
check('True written as -1, which is how the VBE writes it',
    boolean.length === 0, boolean.join(', ') || '(none)');

stop();
done();
