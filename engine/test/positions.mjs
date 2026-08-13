// A finding's line and column, against the text the engine actually analysed.
//
// THE DEFECT THIS EXISTS FOR. A finding crosses the pipe as a character offset, and an offset
// means nothing without the text it was counted in. A diagnostics request may carry its own
// source, or leave it out; when it leaves it out the engine picks between the live copy of the
// module, kept up to date by didChange, and the seeded copy the project was opened with. That
// choice is invisible to the caller. The caller converted the offsets with whatever text it
// happened to be holding, which was the seeded one, and while the two agreed nothing looked
// wrong.
//
// Formatting a module is what made them disagree. The page held the formatted text and the
// editor still held the original, so the engine measured a finding in one and the caller
// converted it in the other. A statement at indent 0 gained a tab, every character after it
// moved, and a finding that had been at 6:11 came back at 6:6 and stayed there: a red underline
// six columns to the left of the word it was about, on the wrong word entirely.
//
// So this test drives the two copies apart on purpose and asks for the positions. The assertion
// is not a hardcoded line and column but the invariant itself: whatever text the engine chose,
// the line and column it reports must describe the offset it reports, in that same text.
//
//   node test/positions.mjs

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('positions');
const { check, done } = reporter();

/** Where each line begins, counting the three line endings VBA modules turn up with. */
function lineStarts(source) {
    const starts = [0];
    for (let at = 0; at < source.length; at++) {
        if (source[at] === '\n') {
            starts.push(at + 1);
        } else if (source[at] === '\r') {
            starts.push(source[at + 1] === '\n' ? at + 2 : at + 1);
            if (source[at + 1] === '\n') { at++; }
        }
    }
    return starts;
}

/** An offset as a one-based line and column, worked out the long way on purpose. */
function positionOf(source, offset) {
    const starts = lineStarts(source);
    let line = 0;
    while (line + 1 < starts.length && starts[line + 1] <= offset) { line++; }
    return { line: line + 1, column: offset - starts[line] + 1 };
}

// The developer's own recipe: the offending statement at INDENT 0, so that formatting has to ADD
// an indent rather than swap four spaces for a tab of the same width. Adding one moves every
// character on the line, which is what exposes a conversion done against the wrong text.
const BEFORE = [
    'Option Explicit',
    '',
    'Public Sub Probe()',
    '    Dim n As Long',
    '    n = 1',
    'Workbooks.Close n',
    '    Debug.Print n',
    'End Sub',
    '',
].join('\r\n');

const AFTER = [
    'Option Explicit',
    '',
    'Public Sub Probe()',
    '\tDim n As Long',
    '\tn = 1',
    '\tWorkbooks.Close n',
    '\tDebug.Print n',
    'End Sub',
    '',
].join('\r\n');

const projectId = 'Positions';
const request = {
    documentKey: `${projectId}/Probe`,
    projectId,
    generation: 1,
    moduleName: 'Probe',
    moduleType: 'standard',
};

await call('initialize', {});
await call('project/open', {
    projectId,
    generation: 1,
    modules: [{ moduleName: 'Probe', source: BEFORE, type: 'standard' }],
});

console.log('\n== the seeded copy, which nothing has moved ==\n');

const seeded = await call('textDocument/diagnostics', request);
const seededFindings = seeded.diagnostics ?? [];

check('the module produces findings to measure', () => assert.ok(seededFindings.length > 0));

check('every position describes its own offset in the seeded text', () => {
    for (const finding of seededFindings) {
        assert.ok(finding.at, `no position on ${JSON.stringify(finding.message)}`);
        assert.deepEqual(
            { line: finding.at.startLine, column: finding.at.startColumn },
            positionOf(BEFORE, finding.span.start),
            `start of ${JSON.stringify(finding.message)}`,
        );
        assert.deepEqual(
            { line: finding.at.endLine, column: finding.at.endColumn },
            positionOf(BEFORE, finding.span.end),
            `end of ${JSON.stringify(finding.message)}`,
        );
    }
});

// THE DIVERGENCE. The live copy takes the formatted text and the seeded copy keeps the original,
// exactly as they stand between a format in the page and the write-back to the editor.
console.log('\n== the live copy moved ahead, as a format leaves it ==\n');

await call('textDocument/didChange', { projectId, moduleName: 'Probe', source: AFTER });

const held = await call('debug/liveSource', { projectId, moduleName: 'Probe', includeText: true });
check('the engine is holding the formatted text', () => {
    assert.equal(held.held, true, 'the engine reports holding nothing');
    assert.equal(held.source, AFTER);
});

// No source on the request, so the engine chooses. It chooses the live copy, and the positions
// it reports have to describe that copy rather than the one the caller was opened with.
const live = await call('textDocument/diagnostics', request);
const liveFindings = live.diagnostics ?? [];

check('every position describes its own offset in the text the engine chose', () => {
    for (const finding of liveFindings) {
        assert.ok(finding.at, `no position on ${JSON.stringify(finding.message)}`);
        assert.deepEqual(
            { line: finding.at.startLine, column: finding.at.startColumn },
            positionOf(AFTER, finding.span.start),
            `start of ${JSON.stringify(finding.message)}`,
        );
    }
});

// And the shape of the original defect, named outright: the finding about Workbooks.Close must
// point AT Close in the formatted text. Converting the same offset against the pre-format text
// puts it at 6:6, on the "b" of Workbooks, which is what the developer saw underlined.
const closeFinding = liveFindings.find((finding) => (finding.message ?? '').includes('Close'));

check('the finding about Close is reported', () => assert.ok(closeFinding));

if (closeFinding) {
    const wanted = positionOf(AFTER, AFTER.indexOf('Close'));
    const wrong = positionOf(BEFORE, AFTER.indexOf('Close'));

    check('it points at Close in the formatted text', () =>
        assert.deepEqual(
            { line: closeFinding.at.startLine, column: closeFinding.at.startColumn },
            wanted,
        ));

    check('it does not point where the pre-format text would put it', () =>
        assert.notDeepEqual(
            { line: closeFinding.at.startLine, column: closeFinding.at.startColumn },
            wrong,
            `converted against the wrong text, ${wrong.line}:${wrong.column}`,
        ));
}

const code = done();
stop();
process.exit(code);
