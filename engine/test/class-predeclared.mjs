// A class module's name used as a value: does the caller go red?
//
// THE CASE ISSUE #4 WAS ABOUT, from the other side. `Ticket.ChangeTest` compiles against a class
// marked `Attribute VB_PredeclaredId = True`, which has a default instance, and is `Variable not
// defined` against a plain class, which is a TYPE and not a value. The VBE refuses to compile the
// second one, and the analyzer said nothing about either until xlide_vscode#47.
//
// That fix reads the attribute out of the module's text, which is right for a file on disk and
// unreachable from a code pane: the `Attribute VB_` lines live only in the EXPORTED .cls. So the
// analyzer grew a host-supplied field (xlide_vscode#50), and this pins what the three states of
// it mean - because the wrong one is not a missed finding, it is a red squiggle under working
// code, twelve times over in a project that uses stdVBA.
//
// The host half reads the flag out of the document Excel has already SAVED - the package's
// vbaProject.bin, where each module's stream carries its attribute header - because a code pane
// never shows it and writing a temporary export to read it back is not something this product
// does. That half is pinned in SavedModulesTests; this is what the engine does with the answer.
//
//   node test/class-predeclared.mjs

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('class-predeclared');
const { check, done } = reporter();

const CRLF = '\r\n';

// A perfectly ordinary class. Nothing in this text says whether it is predeclared - that is the
// whole difficulty, and why the fact has to arrive beside the source rather than inside it.
const TICKET = [
    'Option Explicit',
    '',
    'Public Sub ChangeTest()',
    'End Sub',
].join(CRLF);

// And a caller naming the class bare, as singleton-style VBA does all day.
const CALLER = [
    'Option Explicit',
    '',
    'Public Sub Go()',
    '    Ticket.ChangeTest',
    'End Sub',
].join(CRLF);

const project = (projectId, predeclaredId) => call('project/open', {
    projectId,
    generation: 1,
    modules: [
        {
            moduleName: 'Ticket',
            source: TICKET,
            type: 'class',
            ...(predeclaredId === undefined ? {} : { predeclaredId }),
        },
        { moduleName: 'Caller', source: CALLER, type: 'standard' },
    ],
});

const diagnose = (projectId) => call('textDocument/diagnostics', {
    documentKey: `${projectId}/Caller`,
    projectId,
    generation: 1,
    source: CALLER,
    moduleName: 'Caller',
    moduleType: 'standard',
});

/** Findings that name `Ticket` as an undeclared name, which is the one rule in question. */
const aboutTicket = (answer) => (answer.diagnostics ?? [])
    .filter((one) => /not defined|undeclared/i.test(one.message))
    .filter((one) => /ticket/i.test(one.message));

await call('initialize', {});

// ---- the attribute header was NEVER READ ------------------------------------------------------
//
// Which is every class module a VBE host seeds without doing the export. Unknown is not "no": the
// name might well be a legitimate default instance, and calling it undefined would be a false
// report on correct code. Silence is the honest answer.

await project('Unread', undefined);
const unread = await diagnose('Unread');

check('a class whose attribute header was never read makes no claim either way', () => {
    assert.deepEqual(aboutTicket(unread).map((one) => one.message), [],
        'an unread header cannot convict a name. Got '
        + `${JSON.stringify(unread.diagnostics)}`);
});

// ---- the host read it, and it is FALSE --------------------------------------------------------
//
// A plain class. `Ticket` is a type, not a value, and the VBE will not compile the caller. This
// is the finding the whole exercise exists to produce, and the only state that produces it.

await project('Plain', false);
const plain = await diagnose('Plain');

check('a class VOUCHED FOR as not predeclared does report its name used as a value', () => {
    assert.ok(aboutTicket(plain).length > 0,
        `expected a finding against Ticket; got ${JSON.stringify(plain.diagnostics)}`);
});

// And it has to land on the word, in the module that used it. A finding whose span points
// somewhere else is a finding the developer reads as wrong about something it never mentioned.
check('and it underlines Ticket in the caller, not something else', () => {
    const [finding] = aboutTicket(plain);
    const span = CALLER.slice(finding.span.start, finding.span.end);
    assert.equal(span, 'Ticket', `the span covers ${JSON.stringify(span)}`);
    assert.equal(finding.at?.startLine, 4, 'the line the caller names it on');
});

// ---- the host read it, and it is TRUE ---------------------------------------------------------
//
// The singleton style: the class has a default instance, its bare name IS a value, and the caller
// is correct code. Reporting here is the failure that costs a developer their trust in the pane.

await project('Singleton', true);
const singleton = await diagnose('Singleton');

check('a class vouched for as predeclared is silent about its name used as a value', () => {
    assert.deepEqual(aboutTicket(singleton).map((one) => one.message), [],
        'a predeclared class name is a value. Got '
        + `${JSON.stringify(singleton.diagnostics)}`);
});

// ---- and the answer is not remembered across a reseed -----------------------------------------
//
// The flag rides the seed, so a project reseeded with a different answer must be analysed again
// rather than served from the memo. Same project id, so nothing but the flag distinguishes the
// two passes - which is exactly the confusion a memo keyed on source alone would fall into.

await project('Flip', true);
const beforeFlip = await diagnose('Flip');
await call('project/open', {
    projectId: 'Flip',
    generation: 2,
    modules: [
        { moduleName: 'Ticket', source: TICKET, type: 'class', predeclaredId: false },
        { moduleName: 'Caller', source: CALLER, type: 'standard' },
    ],
});
const afterFlip = await call('textDocument/diagnostics', {
    documentKey: 'Flip/Caller',
    projectId: 'Flip',
    generation: 2,
    source: CALLER,
    moduleName: 'Caller',
    moduleType: 'standard',
});

check('reseeding the same project with the other answer changes the finding', () => {
    assert.deepEqual(aboutTicket(beforeFlip).map((one) => one.message), [],
        'the first pass should be silent');
    assert.ok(aboutTicket(afterFlip).length > 0,
        'the second pass should report; a memo that cannot see the flag would stay silent. Got '
        + `${JSON.stringify(afterFlip.diagnostics)}`);
});

// ---- and null on the wire is absent, not false ------------------------------------------------
//
// JSON can spell "nobody knows" two ways, and the analyzer reads only one: a null passes its
// `!== undefined` test and is stored as a supplied answer. This holds the WIRE to the same meaning
// either way. Measured: it passes with the normalisation removed too, because nothing downstream
// currently distinguishes them - so this pins a contract rather than guarding a known break, and
// it is the check that would notice if that stopped being true.

await call('project/open', {
    projectId: 'Null',
    generation: 1,
    modules: [
        { moduleName: 'Ticket', source: TICKET, type: 'class', predeclaredId: null },
        { moduleName: 'Caller', source: CALLER, type: 'standard' },
    ],
});
const spelledNull = await diagnose('Null');

check('predeclaredId sent as null is treated as unknown, not as false', () => {
    assert.deepEqual(aboutTicket(spelledNull).map((one) => one.message), [],
        'null is an unknown spelt differently, and unknown is silent. Got '
        + `${JSON.stringify(spelledNull.diagnostics)}`);
});

await stop();
done();
