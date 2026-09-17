// A caller is re-analysed when the callee it calls changes shape, with its own text untouched.
//
// THE DEFECT THIS EXISTS FOR. Diagnostics are answered from the last analysis when a module's
// text and the project facts it depends on are both unchanged, which is what turned a pass over
// an 82,000-line project from 476ms into 34ms. The saving is only correct if "the project facts
// it depends on" is exact, and for a while it was not.
//
// The facts are fingerprinted from `projectAnalysisOptionsForModule`, and that object carries
// `projectProcedures` as a **Map**. `JSON.stringify(new Map([...]))` is `"{}"` - not an error,
// not a warning, just an empty object where every procedure signature in the project should be.
// So the fingerprint could not see a signature change at all. Give a procedure an extra
// parameter and every call to it elsewhere stayed exactly as it was: no squiggle, nothing red,
// nothing slow, and an hour spent on a call the editor said was fine.
//
// It reproduced through the editor only about two runs in five, because a caller whose memo had
// happened to be evicted was re-analysed anyway and looked correct. Three sampling errors in the
// end-to-end suite were found and fixed while chasing that, and the real defect outlived all of
// them. Here there is no Excel, no fixture and no timing: seed, ask, reseed, ask.
//
//   node test/freshness.mjs

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('freshness');
const { check, done } = reporter();

const PROJECT = 'freshness.xlsm';

// The caller's text, written once and never changed again. Everything below happens because the
// OTHER module moved.
// `r` is READ so the caller is clean on its own: every check here is about a finding that
// appears or clears because Add moved, and the analyzer's 8.3.0 dead-code rules report a result
// that is only ever assigned (2026-09-17).
const CALLER = [
    'Option Explicit',
    '',
    'Public Sub CallsAcross()',
    '    Dim r As Long',
    '    r = Add(1)',
    '    Debug.Print r',
    'End Sub',
    '',
].join('\r\n');

const oneParameter = [
    'Option Explicit',
    '',
    'Public Function Add(ByVal seed As Long) As Long',
    '    Add = seed',
    'End Function',
    '',
].join('\r\n');

const twoParameters = [
    'Option Explicit',
    '',
    'Public Function Add(ByVal seed As Long, ByVal extra As Long) As Long',
    '    Add = seed + extra',
    'End Function',
    '',
].join('\r\n');

// A body change that alters no declaration, to prove the saving is still being made. If this
// re-analysed too, the check above would pass for the wrong reason: a fingerprint that changes
// on everything is as useless as one that changes on nothing, it just fails safely.
const oneParameterOtherBody = [
    'Option Explicit',
    '',
    'Public Function Add(ByVal seed As Long) As Long',
    '    Dim doubled As Long',
    '    doubled = seed * 2',
    '    Add = doubled - seed',
    'End Function',
    '',
].join('\r\n');

const seed = (calleeSource, generation) => call('project/open', {
    projectId: PROJECT,
    generation,
    modules: [
        { moduleName: 'Callee', source: calleeSource, type: 'standard' },
        { moduleName: 'Caller', source: CALLER, type: 'standard' },
    ],
});

const askCaller = (generation) => call('textDocument/diagnostics', {
    projectId: PROJECT,
    generation,
    documentKey: `${PROJECT}\0caller`,
    moduleName: 'Caller',
    moduleType: 'standard',
    source: CALLER,
});

await call('initialize', {});

await seed(oneParameter, 1);
const clean = await askCaller(1);
check('the caller is clean while Add takes one argument', () => {
    assert.equal(clean.diagnostics.length, 0, `got ${JSON.stringify(clean.diagnostics.map((d) => d.code))}`);
});

// Asked twice at the same generation: this is the memo hit, and it must not change the answer.
const again = await askCaller(1);
check('asking again with nothing changed gives the same answer', () => {
    assert.deepEqual(again.diagnostics, clean.diagnostics);
});

await seed(twoParameters, 2);
const broken = await askCaller(2);
check('the caller reports the call once Add takes two, with its own text untouched', () => {
    assert.equal(
        broken.diagnostics.length,
        1,
        broken.diagnostics.length === 0
            ? 'no finding: the memo served the answer from before the signature changed'
            : JSON.stringify(broken.diagnostics.map((d) => d.code)));
    assert.equal(broken.diagnostics[0].code, 'argument-count');
});

await seed(oneParameter, 3);
const healed = await askCaller(3);
check('and it clears when Add goes back to one', () => {
    assert.equal(healed.diagnostics.length, 0, `got ${JSON.stringify(healed.diagnostics.map((d) => d.code))}`);
});

await seed(oneParameterOtherBody, 4);
const stillClean = await askCaller(4);
check('a change to Add\'s BODY alone leaves the caller clean', () => {
    assert.equal(stillClean.diagnostics.length, 0, `got ${JSON.stringify(stillClean.diagnostics.map((d) => d.code))}`);
});

/*
 * ONE ANALYSIS PER DOCUMENT VERSION.
 *
 * The live path asks with the caret, so a half-typed expression is not underlined yet. The pass
 * asks afterwards without one. Same module, same text, analysed twice: on the 64,802-line fixture
 * that was 3,868ms of a 5,252ms edit.
 *
 * A request without a caret may now reuse an answer computed WITH one. The reverse must not
 * happen, and that asymmetry is what these two checks pin: serving an unsuppressed answer to a
 * live request would put the error back under the cursor mid-expression, which is the whole
 * reason the suppression exists.
 */
const HALF_TYPED = [
    'Option Explicit',
    '',
    'Public Sub Typing()',
    '    Dim value As Long',
    '    value = Add(',
    'End Sub',
    '',
].join('\r\n');

const caretAtOpenParen = HALF_TYPED.indexOf('Add(') + 'Add('.length;

const askTyping = (caretOffset) => call('textDocument/diagnostics', {
    projectId: PROJECT,
    generation: 5,
    documentKey: `${PROJECT} typing`,
    moduleName: 'Typing',
    moduleType: 'standard',
    source: HALF_TYPED,
    ...(caretOffset === null ? {} : { activeIncompleteExpressionOffset: caretOffset }),
});

await call('project/open', {
    projectId: PROJECT,
    generation: 5,
    modules: [
        { moduleName: 'Callee', source: oneParameter, type: 'standard' },
        { moduleName: 'Caller', source: CALLER, type: 'standard' },
        { moduleName: 'Typing', source: HALF_TYPED, type: 'standard' },
    ],
});

const whileTyping = await askTyping(caretAtOpenParen);
const afterTyping = await askTyping(null);

check('a request with no caret reuses the answer computed with one', () => {
    assert.deepEqual(
        afterTyping.diagnostics,
        whileTyping.diagnostics,
        'the pass computed its own answer instead of taking the live one');
});

// And the other direction: the caret request must not be served the answer that has no
// suppression in it. Asked in that order deliberately, so the memo holds the no-caret answer.
const noCaretFirst = await askTyping(null);
const caretAfter = await askTyping(caretAtOpenParen);

check('a request with a caret is not served an answer computed without one', () => {
    assert.ok(
        caretAfter.diagnostics.length <= noCaretFirst.diagnostics.length,
        `caret request got ${caretAfter.diagnostics.length} findings against `
        + `${noCaretFirst.diagnostics.length} without a caret; suppression was lost`);
});

/*
 * A RENAME IS THE ONE EDIT THAT CHANGES A NAME AND NO TEXT.
 *
 * The symbol assembly is memoised per project and the memo was keyed on the modules' SOURCES. A
 * reseed after a rename carries the same sources under a new name, so the memo hit and handed back
 * an assembly that still knew the module by the name it no longer had. The next rename of that
 * module was then refused as not being a module of this workbook, by an assembly one rename out of
 * date, while the object model and the tree both showed the new name.
 *
 * It only ever showed when the rename had nothing to replace. Replacing a mention rewrites another
 * module's text, which misses the memo and rebuilds it, so every rename that did something worked
 * and only the ones that did nothing else broke: a module nothing references, which in practice
 * means a new module or a document (found 2026-08-10, from a document, which is not special).
 *
 * The memo is cleared on closeProject and on nothing else, so a reseed alone never cleared it.
 */
const LONELY = [
    'Option Explicit',
    '',
    'Public Sub Only()',
    'End Sub',
    '',
].join('\r\n');

const seedLonely = (moduleName) => call('project/open', {
    projectId: 'renamememo.xlsm',
    generation: 9,
    modules: [{ moduleName, source: LONELY, type: 'standard' }],
});

await seedLonely('Lonely');

const firstRename = await call('workspace/renameModule', {
    projectId: 'renamememo.xlsm',
    moduleName: 'Lonely',
    newName: 'LonelyA',
});

check('a module nothing references can be renamed', () => {
    assert.equal(firstRename.refused, undefined, `refused: ${firstRename.refused}`);
});

// The shim reseeds after a rename. Same sources, new name: the exact input the memo could not see.
await seedLonely('LonelyA');

const secondRename = await call('workspace/renameModule', {
    projectId: 'renamememo.xlsm',
    moduleName: 'LonelyA',
    newName: 'LonelyB',
});

check('and renamed AGAIN, with no edit in between', () => {
    assert.equal(
        secondRename.refused,
        undefined,
        `refused: ${secondRename.refused}. The assembly is a rename behind the seed.`);
});

stop();
process.exitCode = done();
