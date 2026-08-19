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
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
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

/*
 * AN EMPTY CONTROL LIST IS AN ANSWER. Upstream proves a member ABSENT on a form only when the
 * host vouched for the control list - an EMPTY array included (xlide_vscode#26) - and the shim
 * folded "walked the designer, found nothing" into null until 2026-08-19, which silenced every
 * unknown-member finding on a control-less form. Both sides of the contract are pinned: a
 * supplied empty list proves absence; an absent list proves nothing.
 */
const POKE = 'Public Sub Poke()\r\n    Bare.Nope\r\nEnd Sub\r\n';
await call('project/open', {
    projectId: 'EmptyKnown', generation: 1,
    modules: [
        { moduleName: 'Bare', source: 'Option Explicit\r\n', type: 'userform', implicitMembers: [] },
        { moduleName: 'Uses', source: POKE, type: 'standard' },
    ],
});

const emptyKnown = await call('textDocument/diagnostics', {
    documentKey: 'EmptyKnown/Uses', projectId: 'EmptyKnown', generation: 1,
    source: POKE, moduleName: 'Uses', moduleType: 'standard',
});

check('a vouched-for EMPTY form proves absence: Bare.Nope is a finding', () => {
    assert.ok(emptyKnown.diagnostics.some((one) => one.message.includes('Nope')),
        `expected an unknown-member finding; got ${JSON.stringify(emptyKnown.diagnostics)}`);
});

await call('project/open', {
    projectId: 'EmptyUnknown', generation: 1,
    modules: [
        { moduleName: 'Bare', source: 'Option Explicit\r\n', type: 'userform' },
        { moduleName: 'Uses', source: POKE, type: 'standard' },
    ],
});

const emptyUnknown = await call('textDocument/diagnostics', {
    documentKey: 'EmptyUnknown/Uses', projectId: 'EmptyUnknown', generation: 1,
    source: POKE, moduleName: 'Uses', moduleType: 'standard',
});

check('an unvouched form proves nothing: no control list, no absence claim', () => {
    assert.ok(!emptyUnknown.diagnostics.some((one) => one.message.includes('Nope')),
        `an absent list must stay silent; got ${JSON.stringify(emptyUnknown.diagnostics)}`);
});

/*
 * AND THE FINGERPRINT SEES A DESIGNER CHANGE. The diagnostics memo replays an answer while
 * the cross-module facts hold still, and a control list can change without a single source
 * changing - a designer apply is exactly that. The fingerprint's index ignored implicit
 * members until 2026-08-19, so a removed control kept resolving as a ghost across a reseed:
 * same sources, same fingerprint, replayed findings. This is that regression, as a sequence.
 */
const GHOST = 'Public Sub Poke()\r\n    Wardrobe.Zed.Visible = True\r\nEnd Sub\r\n';
const wardrobe = (members) => [
    { moduleName: 'Wardrobe', source: 'Option Explicit\r\n', type: 'userform', implicitMembers: members },
    { moduleName: 'Pokes', source: GHOST, type: 'standard' },
];
const ghostDiag = (generation) => call('textDocument/diagnostics', {
    documentKey: 'Ghost/Pokes', projectId: 'Ghost', generation,
    source: GHOST, moduleName: 'Pokes', moduleType: 'standard',
});

await call('project/open', { projectId: 'Ghost', generation: 1, modules: wardrobe([{ name: 'Zed', type: 'MSForms.TextBox' }]) });
const withControl = await ghostDiag(1);
await call('project/open', { projectId: 'Ghost', generation: 2, modules: wardrobe([]) });
const withoutControl = await ghostDiag(2);

check('a control that exists resolves; removed across a reseed, it is a finding, not a ghost', () => {
    assert.ok(!withControl.diagnostics.some((one) => one.message.includes('Zed')),
        `Zed exists at generation 1; got ${JSON.stringify(withControl.diagnostics)}`);
    assert.ok(withoutControl.diagnostics.some((one) => one.message.includes('Zed')),
        `the same text after the removal must flag; got ${JSON.stringify(withoutControl.diagnostics)}`);
});

/*
 * THE COLLECTORS' LAST TWO BLIND SPOTS, filed 2026-08-19 and fixed the same day (xlide_vscode
 * #30 and #31, landed with 4.0.2 and consumed in engine/src/semantic.ts - the global collector
 * takes the controls as shadows, the method collector takes meType). These were watcher rows
 * that announced the fixes' arrival; they pin the fixed behaviour now.
 */
const SHADOWED = 'Public Sub T()\r\n    ActiveSheet.Clear\r\nEnd Sub\r\n';
await call('project/open', {
    projectId: 'ShadowTint', generation: 1,
    modules: [{ moduleName: 'F', source: SHADOWED, type: 'userform',
        implicitMembers: [{ name: 'ActiveSheet', type: 'MSForms.ListBox' }] }],
});
const shadowTint = await call('textDocument/semanticTokens', {
    projectId: 'ShadowTint', moduleName: 'F', source: SHADOWED, moduleType: 'userform',
});

check('a control named like a host global wins the binding: no tint, and Clear is the ListBox\'s', () => {
    const receiver = shadowTint.tokens.filter(
        (token) => SHADOWED.slice(token.start, token.end) === 'ActiveSheet');
    assert.deepEqual(receiver, [],
        `the control must not wear the host global's tint; it wore ${JSON.stringify(receiver)}`);
    const clear = shadowTint.tokens.filter(
        (token) => SHADOWED.slice(token.start, token.end) === 'Clear');
    assert.equal(clear.length, 1, `Clear took ${clear.length} token(s)`);
    assert.equal(clear[0].type, 'function');
});

const ME_DOC = 'Public Sub T()\r\n    Me.Calculate\r\n    Sheet1.Calculate\r\nEnd Sub\r\n';
await call('project/open', {
    projectId: 'MeDoc', generation: 1,
    modules: [{ moduleName: 'Sheet1', source: ME_DOC, type: 'document', documentType: 'worksheet' }],
});
const meDoc = await call('textDocument/semanticTokens', {
    projectId: 'MeDoc', moduleName: 'Sheet1', source: ME_DOC, moduleType: 'document',
});

check('Me.Calculate and Sheet1.Calculate paint as one thing: both mentions, both calls', () => {
    const calculates = meDoc.tokens.filter(
        (token) => ME_DOC.slice(token.start, token.end) === 'Calculate');
    assert.equal(calculates.length, 2,
        `both the Me-qualified and the code-name mention paint; got ${JSON.stringify(calculates)}`);
    assert.ok(calculates.every((token) => token.type === 'function'),
        `types were ${calculates.map((token) => token.type).join(', ')}`);
});

/*
 * MEMBER ACCESS FOLLOWS A CONTROL MEMBER INTO WHAT IT RETURNS (xlide_vscode#32, filed on the
 * owner's screenshot - "caption doesn't have a rollover" - and fixed the same day): TabStrip's
 * SelectedItem hovers `As Tab`, and the surface INSIDE the returned Tab now answers too, the
 * way a host chain always did. Upstream's own boundary stands: a primitive or Object-typed
 * return still ends the chain, the way the VBE's own list does.
 */
const CHAIN_DOT = 'Private Sub T()\r\n    Views.SelectedItem.\r\nEnd Sub\r\n';
await call('project/open', {
    projectId: 'Chain', generation: 1,
    modules: [{
        moduleName: 'TabForm', source: CHAIN_DOT, type: 'userform',
        implicitMembers: [{ name: 'Views', type: 'MSForms.TabStrip' }],
    }],
});
const chainItems = await call('textDocument/completion', {
    projectId: 'Chain', moduleName: 'TabForm', source: CHAIN_DOT,
    offset: CHAIN_DOT.indexOf('SelectedItem.') + 'SelectedItem.'.length, moduleType: 'userform',
});
const DIRECT_DOT = 'Private Sub T()\r\n    Views.\r\nEnd Sub\r\n';
const directItems = await call('textDocument/completion', {
    projectId: 'Chain', moduleName: 'TabForm', source: DIRECT_DOT,
    offset: DIRECT_DOT.indexOf('Views.') + 'Views.'.length, moduleType: 'userform',
});

check('a control member resolves directly, and the object it returns resolves too', () => {
    assert.ok((directItems.items ?? []).some((item) => item.label === 'SelectedItem'),
        `Views. must offer SelectedItem; got ${directItems.items?.length ?? 0} item(s)`);
    assert.ok((chainItems.items ?? []).some((item) => item.label === 'Caption'),
        `Views.SelectedItem. must offer the Tab's Caption; got ${chainItems.items?.length ?? 0} item(s)`);
});

const CHAIN_HOVER = 'Private Sub T()\r\n    Dim t As String\r\n    t = Views.SelectedItem.Caption\r\nEnd Sub\r\n';
const chainHover = await call('textDocument/hover', {
    projectId: 'Chain', moduleName: 'TabForm', source: CHAIN_HOVER,
    offset: CHAIN_HOVER.lastIndexOf('Caption') + 3, moduleType: 'userform',
});

check('and the rollover that started it answers: SelectedItem.Caption hovers as the Tab\'s', () => {
    assert.ok(chainHover.hover, 'expected a hover');
    assert.ok(/Caption/.test(chainHover.hover.signature),
        `signature was '${chainHover.hover.signature}'`);
});

/*
 * THE NEXT TWO RECEIVER GAPS, filed and watched (xlide_vscode#33 and #34, the announce
 * idiom): a DECLARED LOCAL with a host type paints nothing while its hover says method, and
 * the host GLOBAL interface's own methods (Word's InchesToPoints) resolve nothing anywhere.
 * Both rows pin what is true today and shout the day the analyzer moves - #33's fix may want
 * the locals threaded through semantic.ts, which is exactly what these exist to catch.
 */
const LOCAL_RECV = 'Public Sub T()\r\n    Dim rng As Range\r\n    Set rng = ActiveDocument.Range(0, 0)\r\n    rng.InsertParagraphAfter\r\n    TopM = InchesToPoints(1)\r\nEnd Sub\r\n';
await call('project/open', {
    projectId: 'WordRecv', generation: 1, host: 'word',
    modules: [{ moduleName: 'M', source: LOCAL_RECV, type: 'standard' }],
});
const localPaint = await call('textDocument/semanticTokens', {
    projectId: 'WordRecv', moduleName: 'M', source: LOCAL_RECV, moduleType: 'standard',
});
const localHover = await call('textDocument/hover', {
    projectId: 'WordRecv', moduleName: 'M', source: LOCAL_RECV,
    offset: LOCAL_RECV.indexOf('InsertParagraphAfter') + 4, moduleType: 'standard',
});

check('a declared local resolves its host member to hover; the paint is xlide_vscode#33, watched', () => {
    assert.equal(localHover.hover?.details?.[0], 'Word host method',
        `the hover read '${localHover.hover?.details?.[0] ?? '(none)'}'`);
    const painted = localPaint.tokens.filter(
        (token) => LOCAL_RECV.slice(token.start, token.end) === 'InsertParagraphAfter');
    if (painted.length > 0) {
        console.log('     UPSTREAM FIXED: local receivers paint - check semantic.ts threads the locals, pin it, close xlide_vscode#33.');
        assert.ok(painted.every((token) => token.type === 'function'),
            `painted as ${painted.map((token) => token.type).join(', ')}`);
    }
});

const globalFnHover = await call('textDocument/hover', {
    projectId: 'WordRecv', moduleName: 'M', source: LOCAL_RECV,
    offset: LOCAL_RECV.indexOf('InchesToPoints') + 4, moduleType: 'standard',
});

check('the host Global interface is xlide_vscode#34, watched: InchesToPoints answers when modelled', () => {
    if (globalFnHover.hover) {
        console.log('     UPSTREAM FIXED: the Global interface resolves - pin the hover here and close xlide_vscode#34.');
        assert.ok(/InchesToPoints/.test(globalFnHover.hover.signature),
            `signature was '${globalFnHover.hover.signature}'`);
    }
});

// NOTE: the word open above re-hosts the engine; the sync fork block below opens access and
// expects to be LAST. Anything excel-hosted goes before the WordRecv open.

/*
 * THE SYNC FORK FOLLOWS THE HOST (decision 15's narrowing, 2026-08-19): a .frm/.frx pair is a
 * CREATE where the applier's VBComponents.Import can land it, and Access's VBE carries no
 * MSForms at all - no UserForms exist there, so the planner's refusal is the truth and the
 * promotion must stand down. The host is the process's identity, set by project/open, so the
 * access half runs LAST in this file: everything above assumes Excel, and the first version
 * of this block sat before the watchers above and quietly re-hosted them (caught the same
 * hour it was written: the #30 row announced a fix that had not happened).
 */
const syncFolder = mkdtempSync(join(tmpdir(), 'xlide-forms-sync-'));
writeFileSync(join(syncFolder, 'GhostForm.frm'), [
    'VERSION 5.00',
    'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GhostForm ',
    '   Caption         =   "Ghost"',
    'End',
    'Attribute VB_Name = "GhostForm"',
    'Option Explicit',
    '',
].join('\r\n'));
writeFileSync(join(syncFolder, 'GhostForm.frx'), Buffer.from([0, 1, 2, 3]));

const planFor = () => call('sync/plan', {
    direction: 'import',
    workbookPath: 'X:\\nowhere\\Probe.xlsm',
    folder: syncFolder,
    mode: 'updateOnly',
    modules: [{ name: 'Module1', type: 'standard', source: 'Option Explicit\r\n' }],
});

const rowOf = (plan) => (plan.items ?? []).find(
    (item) => /GhostForm\.frm$/i.test(item.relativeName ?? ''));

try {
    const inExcel = await planFor();

    check('in excel, a .frm with its .frx beside it plans as a create', () => {
        const row = rowOf(inExcel);
        assert.ok(row, `no GhostForm row in ${JSON.stringify(inExcel.items?.map((i) => i.relativeName))}`);
        assert.equal(row.status, 'will-create', `the row read ${row.status}`);
    });

    // Now the engine learns it is running in Access, the way project/open tells it.
    await call('project/open', {
        projectId: 'AccessProbe', generation: 1, host: 'access',
        modules: [{ moduleName: 'M', source: 'Option Explicit\r\n', type: 'standard' }],
    });

    const inAccess = await planFor();

    check('in access, the same pair stays refused: no host without MSForms offers a create', () => {
        const row = rowOf(inAccess);
        assert.ok(row, `no GhostForm row in ${JSON.stringify(inAccess.items?.map((i) => i.relativeName))}`);
        assert.equal(row.status, 'skipping-import', `the row read ${row.status}`);
    });
} finally {
    rmSync(syncFolder, { recursive: true, force: true });
}

stop();
process.exit(done());
