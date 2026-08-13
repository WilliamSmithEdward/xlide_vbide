// The form language over the pipe: controls the DESIGNER declares, not the text.
//
// Every other module kind declares its names in its own source. A form does not: the host reads
// the designer and seeds each control as an implicit member (xlide_vscode#17), and the analyzer
// is supposed to treat them as declared everywhere - completion, hover, diagnostics, and now
// semantic colouring (#20), whose acceptance table is exact enough to pin verbatim: a RESOLVED
// METHOD call on a control paints `function`, and everything else - a property read, an
// unresolved member, the same text in a module with no controls - paints nothing.
//
// This file exists because the seeding is a chain with three links (host seed, engine
// pass-through, analyzer resolution) and the middle link is this repo's. The analyzer's own
// tests hold the third link still; the live designer suite holds the first against a real
// workbook; this holds the middle one headlessly, so a break in it fails a commit rather than
// a live run.
//
//   node test/forms.mjs
//   node test/forms.mjs --exe

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('forms');
const { check, done } = reporter();

/*
 * The acceptance table's expressions, one per line, in a form whose controls are seeded below.
 * `Me.Hide` engages through meType rather than through a control, so it exercises the OTHER
 * field of the pass-through; `Me.RegionPick.AddItem` is the same resolved call reached through
 * the form, which upstream pins as painting identically.
 */
const FORM_CODE = [
    'Option Explicit',
    '',
    'Private Sub Describe()',
    '    RegionPick.AddItem "North"',
    '    NameBox.SetFocus',
    '    Me.Hide',
    '    Taxable.Value = True',
    '    RegionPick.NotAMember',
    '    Me.RegionPick.AddItem "South"',
    'End Sub',
    '',
].join('\r\n');

const CONTROLS = [
    { name: 'RegionPick', type: 'MSForms.ComboBox' },
    { name: 'NameBox', type: 'MSForms.TextBox' },
    { name: 'Taxable', type: 'MSForms.CheckBox' },
];

const hello = await call('initialize', {});
check('the engine answers the handshake', () => assert.equal(hello.engine, 'xlide'));

const opened = await call('project/open', {
    projectId: 'Forms',
    generation: 1,
    modules: [{ moduleName: 'EntryForm', source: FORM_CODE, type: 'userform', implicitMembers: CONTROLS }],
});
check('a form module opens with its designer seed', () => assert.equal(opened.modules, 1));

const coloured = await call('textDocument/semanticTokens', {
    projectId: 'Forms',
    moduleName: 'EntryForm',
    source: FORM_CODE,
    moduleType: 'userform',
});

console.log(`  -> ${coloured.tokens.length} semantic token(s):`);
for (const token of coloured.tokens) {
    console.log(`     ${token.type} ${JSON.stringify(FORM_CODE.slice(token.start, token.end))}`);
}

/** Every token whose span is exactly the given word. */
const over = (word) => coloured.tokens.filter(
    (token) => FORM_CODE.slice(token.start, token.end) === word);

check('a resolved method call paints as a function: RegionPick.AddItem, both mentions', () => {
    const hits = over('AddItem');
    assert.equal(hits.length, 2, 'the plain call and the Me.-qualified one');
    assert.ok(hits.every((token) => token.type === 'function'),
        `types were ${hits.map((token) => token.type).join(', ')}`);
});

check('a method from the Control BASE class resolves too: NameBox.SetFocus', () => {
    // SetFocus is not in TextBox's own dump - it lives on MSForms.Control, which the library
    // declares once and the surface lookup merges into every placed control (#20's side find).
    const hits = over('SetFocus');
    assert.equal(hits.length, 1);
    assert.equal(hits[0].type, 'function');
});

check("Me engages as the form itself: Me.Hide paints", () => {
    const hits = over('Hide');
    assert.equal(hits.length, 1);
    assert.equal(hits[0].type, 'function');
});

check('a property stays as it is: Taxable.Value takes no token', () =>
    assert.deepEqual(over('Value'), []));

check('an unresolved member is never painted: RegionPick.NotAMember takes no token', () =>
    assert.deepEqual(over('NotAMember'), []));

check('tokens arrive in position order, which the surface encodes as deltas', () => {
    for (let i = 1; i < coloured.tokens.length; i++) {
        assert.ok(coloured.tokens[i].start >= coloured.tokens[i - 1].start,
            `token ${i} starts before token ${i - 1}`);
    }
});

/*
 * The same text in a STANDARD module: no controls seeded, no MSForms Me. Nothing may paint -
 * this is the table's last row generalised, and it is also the meType gate: a module whose Me
 * is not MSForms must not engage the collector.
 */
await call('project/open', {
    projectId: 'FormsPlain',
    generation: 1,
    modules: [{ moduleName: 'Plain', source: FORM_CODE, type: 'standard' }],
});

const plain = await call('textDocument/semanticTokens', {
    projectId: 'FormsPlain',
    moduleName: 'Plain',
    source: FORM_CODE,
    moduleType: 'standard',
});

check('the same text in a module with no controls paints no calls', () => {
    const painted = plain.tokens
        .filter((token) => token.type === 'function')
        .map((token) => FORM_CODE.slice(token.start, token.end));
    assert.deepEqual(painted, []);
});

/*
 * The base-class merge reaches completion through the same seed: NameBox. offers what
 * MSForms.Control declares alongside what TextBox does, and none of the `_`-prefixed dispatch
 * internals the VBE's own list hides.
 */
const caret = FORM_CODE.indexOf('NameBox.') + 'NameBox.'.length;
const completions = await call('textDocument/completion', {
    projectId: 'Forms',
    moduleName: 'EntryForm',
    source: FORM_CODE,
    offset: caret,
    moduleType: 'userform',
});

check('NameBox. offers the base class alongside its own members', () => {
    const labels = (completions.items ?? []).map((item) => item.label);
    for (const wanted of ['SetFocus', 'Move', 'Visible', 'Text']) {
        assert.ok(labels.includes(wanted), `expected ${wanted} among ${labels.length} items`);
    }
    assert.ok(!labels.includes('AddItem'), 'a TextBox must not offer a ComboBox member');
});

check('the dispatch internals stay hidden, the way the VBE hides them', () => {
    const underscored = (completions.items ?? [])
        .map((item) => item.label)
        .filter((label) => label.startsWith('_'));
    assert.deepEqual(underscored, []);
});

stop();
process.exit(done());
