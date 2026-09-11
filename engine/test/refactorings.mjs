/*
 * The lightbulb's verdicts through the real engine.
 *
 * textDocument/refactorings runs each refactoring's own planner at a caret or a selection and
 * answers null where it would go through and the planner's refusal where it would not, so the
 * lightbulb offers only what applies. Pinned here: the case the owner reported (2026-09-10) - the
 * bulb lit on `End Function` offering to inline a keyword - one case each way for every
 * refactoring the lightbulb can raise, Move to Module asked without a destination, and the cost
 * of a caret's worth of questions.
 *
 *   node test/refactorings.mjs
 *   node test/refactorings.mjs --exe
 */

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('refactorings');
const { check, done } = reporter();
const CRLF = '\r\n';

const LEDGER = [
    'Option Explicit',                                                                  // 1
    '',                                                                                 // 2
    'Private held As Long',                                                             // 3
    '',                                                                                 // 4
    'Public Function Balance(ByVal opening As Long, ByVal movement As Long) As Long',   // 5
    '    Balance = opening + movement',                                                 // 6
    '    Debug.Print "Ledger.Balance"',                                                 // 7
    'End Function',                                                                     // 8
    '',                                                                                 // 9
    'Public Sub Settle(ByVal amount As Long)',                                          // 10
    '    Dim limit As Long',                                                            // 11
    '    limit = 10',                                                                   // 12
    '    Debug.Print limit, amount',                                                    // 13
    '    Dim twice As Long',                                                            // 14
    '    twice = 1',                                                                    // 15
    '    twice = 2',                                                                    // 16
    '    Debug.Print twice * 3',                                                        // 17
    'End Sub',                                                                          // 18
    '',                                                                                 // 19
    'Public Sub Uses()',                                                                // 20
    '    Debug.Print held',                                                             // 21
    'End Sub',                                                                          // 22
].join(CRLF);

const OTHER = ['Option Explicit', '', 'Public Sub Elsewhere()', 'End Sub'].join(CRLF);
const ISTORE = ['Option Explicit', '', 'Public Sub Save()', 'End Sub'].join(CRLF);
const STORE = ['Option Explicit', '', 'Implements IStore', '', 'Public Label As String', 'Public Items(3) As Long'].join(CRLF);
const DONE = ['Option Explicit', '', 'Implements IStore', '', 'Private Sub IStore_Save()', 'End Sub'].join(CRLF);
const ALONE = ['Option Explicit', '', 'Public Sub Only()', 'End Sub'].join(CRLF);

await call('initialize', {});
await call('project/open', {
    projectId: 'ledger', generation: 1, host: 'excel',
    modules: [
        { moduleName: 'Ledger', source: LEDGER, type: 'standard' },
        { moduleName: 'Other', source: OTHER, type: 'standard' },
        { moduleName: 'IStore', source: ISTORE, type: 'class' },
        { moduleName: 'Store', source: STORE, type: 'class' },
        { moduleName: 'Done', source: DONE, type: 'class' },
    ],
});
await call('project/open', {
    projectId: 'alone', generation: 1, host: 'excel',
    modules: [{ moduleName: 'Alone', source: ALONE, type: 'standard' }],
});

/** The UTF-16 offset of `text` on a 1-based line of `source`. */
function offsetOf(source, line, text) {
    const lines = source.split(CRLF);
    const column = lines[line - 1].indexOf(text);
    assert.ok(column >= 0, `'${text}' is not on line ${line}`);
    return lines.slice(0, line - 1).reduce((sum, one) => sum + one.length + CRLF.length, 0) + column;
}

async function verdicts(projectId, moduleName, candidates) {
    const answer = await call('textDocument/refactorings', { projectId, moduleName, candidates });
    return answer.verdicts;
}

function offered(verdict) {
    assert.equal(verdict.refused, null, `withheld: ${verdict.refused}`);
}

function withheld(verdict, pattern) {
    assert.notEqual(verdict.refused, null, 'it was offered');
    assert.match(String(verdict.refused), pattern);
}

/* ---- the owner's case: End Function ------------------------------------------------------------ */

const endFunction = offsetOf(LEDGER, 8, 'Function');
const onKeyword = await verdicts('ledger', 'Ledger', [
    { kind: 'introduceParameter', offset: endFunction },
    { kind: 'inlineVariable', offset: endFunction },
]);

check('a keyword under the caret is neither inlined nor made a parameter - the End Function the owner reported', () => {
    withheld(onKeyword[0], /not a local/);
    withheld(onKeyword[1], /not a local/);
});

check('the verdicts come back in the order they were asked, kind for kind', () => {
    assert.deepEqual(onKeyword.map((one) => one.kind), ['introduceParameter', 'inlineVariable']);
});

/* ---- a name under the caret: offered where it applies, withheld where it does not --------------- */

const limit = offsetOf(LEDGER, 11, 'limit');
const twice = offsetOf(LEDGER, 14, 'twice');
const amount = offsetOf(LEDGER, 10, 'amount');
const names = await verdicts('ledger', 'Ledger', [
    { kind: 'inlineVariable', offset: limit },
    { kind: 'introduceParameter', offset: limit },
    { kind: 'inlineVariable', offset: twice },
    { kind: 'introduceParameter', offset: twice },
    { kind: 'inlineVariable', offset: amount },
    { kind: 'introduceParameter', offset: amount },
]);

check('a local assigned once from a literal is offered both ways', () => {
    offered(names[0]);
    offered(names[1]);
});

check('a local assigned twice is withheld both ways, saying so', () => {
    withheld(names[2], /assigned 2 times/);
    withheld(names[3], /assigned 2 times/);
});

check('a parameter is withheld both ways, saying what it already is', () => {
    withheld(names[4], /is a parameter/);
    withheld(names[5], /already a parameter/);
});

/* ---- Move to Module, asked without a destination ------------------------------------------------ */

const moves = await verdicts('ledger', 'Ledger', [
    { kind: 'moveToModule', offset: offsetOf(LEDGER, 10, 'Settle') },
    { kind: 'moveToModule', offset: offsetOf(LEDGER, 20, 'Uses') },
]);

check('a procedure another standard module would take is offered the move', () => {
    offered(moves[0]);
});

check('one that reaches for a Private of its module is withheld, naming it, with no destination in the words', () => {
    withheld(moves[1], /'held'.*Private to 'Ledger'.*any other module/);
});

const [inAClass] = await verdicts('ledger', 'Done', [{ kind: 'moveToModule', offset: offsetOf(DONE, 5, 'IStore_Save') }]);
check('a class member is withheld: only a standard module\'s procedures move', () => {
    withheld(inAClass, /not a standard module/);
});

const [alone] = await verdicts('alone', 'Alone', [{ kind: 'moveToModule', offset: offsetOf(ALONE, 3, 'Only') }]);
check('a project with nowhere else to go is withheld, saying so', () => {
    withheld(alone, /no other standard module/);
});

/* ---- the selections ----------------------------------------------------------------------------- */

const expression = offsetOf(LEDGER, 17, 'twice * 3');
const selections = await verdicts('ledger', 'Ledger', [
    { kind: 'extractVariable', startOffset: expression, endOffset: expression + 'twice * 3'.length },
    { kind: 'extractVariable', startOffset: expression, endOffset: expression + 'twice *'.length },
    { kind: 'extractMethod', startLine: 11, endLine: 13 },
    { kind: 'extractMethod', startLine: 10, endLine: 11 },
]);

check('a whole expression is offered Extract Variable, before any name has been typed', () => {
    offered(selections[0]);
});

check('half of one is withheld', () => {
    withheld(selections[1], /whole expression/);
});

check('whole statements are offered Extract Method, before any name has been typed', () => {
    offered(selections[2]);
});

check('a selection taking the procedure\'s own header line is withheld', () => {
    withheld(selections[3], /Sub line/);
});

/* ---- the declaration lines ---------------------------------------------------------------------- */

const declarations = await verdicts('ledger', 'Store', [
    { kind: 'encapsulateField', fieldName: 'Label' },
    { kind: 'encapsulateField', fieldName: 'Items' },
    { kind: 'implementInterface', interfaceName: 'IStore' },
]);

check('a public field is offered encapsulation', () => {
    offered(declarations[0]);
});

check('an array is withheld: VBA cannot pass one to a Property Let', () => {
    withheld(declarations[1], /array/);
});

check('an Implements line with members still owed is offered them', () => {
    offered(declarations[2]);
});

const [finished] = await verdicts('ledger', 'Done', [{ kind: 'implementInterface', interfaceName: 'IStore' }]);
check('and one that owes nothing is withheld, which is the bulb not lighting on a finished class', () => {
    withheld(finished, /already implements every member/);
});

/* ---- what is not a question --------------------------------------------------------------------- */

const odd = await verdicts('ledger', 'Ledger', [
    { kind: 'renameEverything', offset: 1 },
    { kind: 'inlineVariable' },
]);

check('an unknown kind is refused rather than guessed at', () => {
    withheld(odd[0], /not a refactoring this engine knows/);
});

check('a candidate missing its offset is refused rather than run at offset zero', () => {
    withheld(odd[1], /missing what it needs/);
});

const [unheld] = await verdicts('ledger', 'NoSuchModule', [{ kind: 'inlineVariable', offset: 0 }]);
check('a module the engine does not hold refuses every candidate', () => {
    withheld(unheld, /not one the engine holds/);
});

/* ---- cheap enough for every caret move ----------------------------------------------------------- */

// A caret's worth: the two a name raises and the move a header line raises, which is the most a
// bare caret asks at once. Over the pipe, so the figure is what the shim waits for.
const samples = [];
for (let run = 0; run < 60; run += 1) {
    const began = performance.now();
    await verdicts('ledger', 'Ledger', [
        { kind: 'introduceParameter', offset: limit },
        { kind: 'inlineVariable', offset: limit },
        { kind: 'moveToModule', offset: offsetOf(LEDGER, 10, 'Settle') },
    ]);
    samples.push(performance.now() - began);
}

samples.sort((a, b) => a - b);
const median = samples[Math.floor(samples.length / 2)];
const p95 = samples[Math.floor(samples.length * 0.95)];
console.log(`     a caret's worth of verdicts: median ${median.toFixed(2)}ms, p95 ${p95.toFixed(2)}ms`);

check('a caret\'s worth of verdicts stays well inside what a caret move can spend', () => {
    assert.ok(p95 < 50, `p95 ${p95.toFixed(2)}ms is past the 50ms budget`);
});

stop();
process.exit(done());
