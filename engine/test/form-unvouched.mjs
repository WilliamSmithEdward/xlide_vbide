// A form whose designer could not be read: does its OWN code-behind go red?
//
// THE CASE ISSUE #5 IS ABOUT. Showing a UserForm makes the VBE stop handing out its designer,
// so the host seeds that form with its control list ABSENT - the honest answer, and one the
// protocol can carry (`implicitMembers` is nullable, and the host's reader is explicit that
// "Null and empty are DIFFERENT seeds - unreadable vs vouched-for-empty").
//
// forms.mjs already pins the neighbouring case: another module poking at `Bare.Nope` makes no
// absence claim when the list is unvouched. That is the MEMBER rule standing down. This asks the
// other one - the form's own code-behind naming its own controls bare, which is the
// undeclared-variable rule under Option Explicit, and which nothing covered.
//
// In the VBE a developer saw seven of them at once on code the compiler had just accepted.
//
// NOT IN THE GATE YET, DELIBERATELY. Its third check fails on current behaviour, and a suite that
// is expected to fail is a suite people learn to skip. It goes into engine/package.json's test
// script the moment xlide_vscode#48 lands, and its failing until then is the point: it is the
// reproduction that issue was filed with.
//
// The host half needs nothing: the reader already seeds null rather than empty, and measured on
// a live session the findings clear the instant the form is unloaded and the designer answers
// again. What is left is only this.
//
//   node test/form-unvouched.mjs

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('form-unvouched');
const { check, done } = reporter();

// A form's code-behind referring to its own controls, the way every form's does.
const BEHIND = [
    'Option Explicit',
    '',
    'Private Sub Recalculate()',
    '    lblPreview.Caption = txtAmount.Text',
    'End Sub',
].join('\r\n');

const form = (implicitMembers) => ({
    moduleName: 'frmLoan',
    source: BEHIND,
    type: 'userform',
    ...(implicitMembers === undefined ? {} : { implicitMembers }),
});

const diagnose = (projectId, generation) => call('textDocument/diagnostics', {
    documentKey: `${projectId}/frmLoan`,
    projectId,
    generation,
    source: BEHIND,
    moduleName: 'frmLoan',
    moduleType: 'userform',
});

const undeclared = (answer) => (answer.diagnostics ?? [])
    .filter((one) => /not defined|undeclared/i.test(one.message))
    .map((one) => one.message);

await call('initialize', {});

// ---- the control list is KNOWN ---------------------------------------------------------------

await call('project/open', {
    projectId: 'Vouched',
    generation: 1,
    modules: [form([
        { name: 'txtAmount', type: 'MSForms.TextBox' },
        { name: 'lblPreview', type: 'MSForms.Label' },
    ])],
});

const vouched = await diagnose('Vouched', 1);
check('a form whose controls are seeded reports nothing against its own code-behind', () => {
    assert.deepEqual(undeclared(vouched), [],
        `expected silence; got ${JSON.stringify(vouched.diagnostics)}`);
});

// ---- the control list is VOUCHED-FOR EMPTY ---------------------------------------------------
//
// The host looked, and the form genuinely has no controls. Every reference IS undeclared, and
// saying so is correct - this is the case that must keep working.

await call('project/open', { projectId: 'Empty', generation: 1, modules: [form([])] });
const empty = await diagnose('Empty', 1);
check('a form vouched for as EMPTY does report its code-behind as undeclared', () => {
    assert.ok(undeclared(empty).length > 0,
        `expected findings against a genuinely empty form; got ${JSON.stringify(empty.diagnostics)}`);
});

// ---- and the list is UNKNOWN ------------------------------------------------------------------
//
// The host could not read the designer, so it seeded no list at all. An unknown is not an
// emptiness: the honest answer is that these names cannot be judged, and the rule should stand
// down for this module rather than call every control a missing variable.

await call('project/open', { projectId: 'Unknown', generation: 1, modules: [form(undefined)] });
const unknown = await diagnose('Unknown', 1);

check('a form whose designer could not be READ makes no undeclared claim about its controls', () => {
    assert.deepEqual(undeclared(unknown), [],
        'a form with no control list is a form whose names cannot be judged. Got '
        + `${JSON.stringify(unknown.diagnostics)}`);
});

await stop();
done();
