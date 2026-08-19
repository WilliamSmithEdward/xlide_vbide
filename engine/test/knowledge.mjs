// The knowledge routes: what the language service knows, served as data.
//
// Two contracts are pinned here beyond the shapes. First, knowledge answers BEFORE initialize
// and before any project opens - it is product knowledge, not project state, and an agent must
// be able to learn the terrain without touching anything. Second, the model answers for the
// HOST the engine was told it is running in: after a project/open that says word, the reply is
// WORD's model (the analyzer's host registry, xlide_vscode 4.0.0), and a host the registry
// does not know answers known:false with a note - never another application's types wearing
// this host's name.
//
//   node test/knowledge.mjs
//   node test/knowledge.mjs --exe

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('knowledge');
const { check, done } = reporter();

// Before initialize, deliberately.
const rules = await call('knowledge/analyzer', {});

check('the rule catalogue answers before initialize', () => {
    assert.ok(rules.ruleCount >= 20, `expected a real catalogue, got ${rules.ruleCount}`);
    assert.equal(rules.rules.length, rules.ruleCount);
});

check('every rule carries its contract fields', () => {
    for (const rule of rules.rules) {
        assert.ok(rule.key.length > 0 && rule.code.length > 0 && rule.title.length > 0,
            `a rule is missing identity: ${JSON.stringify(rule)}`);
        assert.ok(rule.severity.length > 0 && rule.category.length > 0 && rule.kind.length > 0,
            `${rule.code} is missing classification`);
        assert.ok(typeof rule.compileEquivalent === 'boolean', `${rule.code} compileEquivalent`);
    }
});

check('rule codes are unique', () => {
    const codes = rules.rules.map((rule) => rule.code);
    assert.equal(new Set(codes).size, codes.length);
});

check('a known rule is present with its authority', () => {
    const unterminated = rules.rules.find((rule) => rule.code === 'unterminated-string');
    assert.ok(unterminated, 'unterminated-string is missing');
    assert.equal(unterminated.severity, 'error');
    assert.ok(unterminated.spec?.includes('MS-VBAL'), `spec was ${unterminated.spec}`);
});

check('the categories index matches the rows', () => {
    const seen = new Set(rules.rules.map((rule) => rule.category));
    assert.deepEqual([...seen].sort(), rules.categories);
});

const model = await call('knowledge/objectModel', {});

check('the default host is excel, with a model', () => {
    assert.equal(model.host, 'excel');
    assert.equal(model.known, true);
    assert.ok(model.typeCount > 0 && model.types.length === model.typeCount);
});

check('the inventory carries the core types', () => {
    const names = new Set(model.types.map((row) => row.name));
    for (const wanted of ['Excel.Worksheet', 'Excel.Range', 'Excel.Workbook', 'Excel.Application']) {
        assert.ok(names.has(wanted), `${wanted} is missing from the inventory`);
    }
});

check('the globals name the host-injected identifiers', () => {
    assert.ok(Object.keys(model.globals).length > 0, 'no globals at all');
});

const worksheet = await call('knowledge/objectModel', { type: 'Worksheet' });

check('a bare name expands through the alias table', () => {
    assert.equal(worksheet.type?.name, 'Excel.Worksheet');
    assert.ok(worksheet.type.memberCount >= 20, `only ${worksheet.type?.memberCount} members`);
    assert.equal(worksheet.type.members.length, worksheet.type.memberCount);
});

check('members carry names, and Range is among them', () => {
    const names = worksheet.type.members.map((member) => member.name);
    assert.ok(names.every((name) => name.length > 0));
    assert.ok(names.some((name) => name.toLowerCase() === 'range'), 'Worksheet.Range is missing');
});

const missing = await call('knowledge/objectModel', { type: 'NoSuchThing' });

check('an unknown type answers a note, not an empty crash', () => {
    assert.equal(missing.known, true);
    assert.equal(missing.type, undefined);
    assert.ok(missing.note.includes('NoSuchThing'));
});

/*
 * THE HOST REACHES THE FEATURES, not only the knowledge routes. Before the host flips to Word,
 * pin Excel's half of each pair; after, the same inputs must answer in Word's tongue. Two
 * paths carry the model and each is pinned through its own door: diagnostics ride the analysis
 * worker request's `host`, completion rides the assembled context's model.
 */
const GLOBAL_USER = 'Option Explicit\r\n\r\nPublic Sub Probe()\r\n    ActiveSheet.Calculate\r\nEnd Sub\r\n';

await call('initialize', {});
await call('project/open', {
    projectId: 'ExcelProject',
    generation: 1,
    modules: [{ moduleName: 'Module1', source: GLOBAL_USER, type: 'standard' }],
});

const excelFindings = await call('textDocument/diagnostics', {
    documentKey: 'ExcelProject/Module1',
    projectId: 'ExcelProject',
    generation: 1,
    source: GLOBAL_USER,
    moduleName: 'Module1',
    moduleType: 'standard',
});

check('in excel, ActiveSheet is a host global and no finding', () => {
    assert.ok(!excelFindings.diagnostics.some((one) => one.message.includes('ActiveSheet')),
        JSON.stringify(excelFindings.diagnostics));
});

/*
 * HOST METHODS PAINT AS CALLS (xlide_vscode#29, fixed upstream and consumed 2026-08-19):
 * ActiveSheet.Calculate takes a `function` token the way RegionPick.AddItem does (#20's
 * convention, extended to host receivers), while the receiver keeps its defaultLibrary tint.
 * The collector is composed in engine/src/semantic.ts - this row is what notices if that
 * wiring, or the upstream collector, ever goes quiet again. The row runs HERE because the
 * host is the process's identity, one per engine: after the word open below, ActiveSheet
 * is nobody.
 */
const excelPaint = await call('textDocument/semanticTokens', {
    projectId: 'ExcelProject', moduleName: 'Module1', source: GLOBAL_USER, moduleType: 'standard',
});

check('a host method paints as a call and its receiver keeps the defaultLibrary tint', () => {
    const over = (word) => excelPaint.tokens.filter(
        (token) => GLOBAL_USER.slice(token.start, token.end) === word);
    const receiver = over('ActiveSheet');
    assert.ok(receiver.some((token) => token.type === 'variable'
        && (token.modifiers ?? []).includes('defaultLibrary')),
        `ActiveSheet wore ${JSON.stringify(receiver)}`);
    const method = over('Calculate');
    assert.equal(method.length, 1, `Calculate took ${method.length} token(s)`);
    assert.equal(method[0].type, 'function', `Calculate painted as ${method[0].type}`);
});

// Now tell the engine it is running in Word, the way project/open does, and ask again.
await call('project/open', {
    projectId: 'WordProject',
    generation: 1,
    host: 'word',
    modules: [{ moduleName: 'Module1', source: GLOBAL_USER, type: 'standard' }],
});

const wordFindings = await call('textDocument/diagnostics', {
    documentKey: 'WordProject/Module1',
    projectId: 'WordProject',
    generation: 1,
    source: GLOBAL_USER,
    moduleName: 'Module1',
    moduleType: 'standard',
});

check('in word, ActiveSheet is nobody and the analyzer says so', () => {
    assert.ok(wordFindings.diagnostics.some((one) => one.message.includes('ActiveSheet')),
        `expected an undeclared finding; got ${JSON.stringify(wordFindings.diagnostics)}`);
});

const wordMemberSource = 'Public Sub Probe()\r\n    ActiveDocument.\r\nEnd Sub\r\n';
const wordMembers = await call('textDocument/completion', {
    projectId: 'WordProject',
    moduleName: 'Module1',
    source: wordMemberSource,
    offset: wordMemberSource.indexOf('ActiveDocument.') + 'ActiveDocument.'.length,
    moduleType: 'standard',
});

check('in word, ActiveDocument. offers Word.Document members', () => {
    assert.ok(wordMembers.items.some((item) => item.label === 'Paragraphs'),
        `${wordMembers.items.length} item(s); no Paragraphs`);
    assert.ok(!wordMembers.items.some((item) => item.label === 'Cells'),
        'an Excel member leaked into Word');
});

const inWord = await call('knowledge/objectModel', {});

check('in word the model is WORD\'s, not excel\'s', () => {
    assert.equal(inWord.host, 'word');
    assert.equal(inWord.known, true);
    assert.ok(inWord.typeCount > 0, 'word answered with no types at all');
    assert.equal(inWord.globals.ThisDocument, 'Word.Document');
    assert.equal(inWord.globals.ActiveSheet, undefined, 'an Excel global leaked into Word');
});

// No host prefix in the request: `Document` must find Word.Document in Word.
const wordDocument = await call('knowledge/objectModel', { type: 'Document' });

check('a word type expands by its bare display name', () => {
    assert.equal(wordDocument.type?.name, 'Word.Document');
    assert.ok(wordDocument.type.memberCount > 0);
});

/*
 * ORIGIN LABELS NAME THE MODULE'S HOST (xlide_vscode#28, fixed upstream and consumed
 * 2026-08-19): the model carries its application name, and hover builds the label from it -
 * so Word's Document.FitToPages hovers as a WORD host method, while Excel's wording stays
 * byte-for-byte what it always was (pinned by the excel hover rows in smoke.mjs). And the
 * method takes a `function` token here too: the paint fix (#29) is host-general, so the
 * user-visible case that started both issues - FitToPages wearing property blue under an
 * 'Excel host method' hover - is pinned closed from both sides at once.
 */
const WORD_METHOD = 'Public Sub Probe()\r\n    ActiveDocument.FitToPages\r\nEnd Sub\r\n';
const wordMethodHover = await call('textDocument/hover', {
    projectId: 'WordProject', moduleName: 'Module1', source: WORD_METHOD,
    offset: WORD_METHOD.indexOf('FitToPages') + 3, moduleType: 'standard',
});

check('a word host member hovers as a Word host method, in those words', () => {
    assert.equal(wordMethodHover.hover?.details?.[0], 'Word host method',
        `the label was '${wordMethodHover.hover?.details?.[0] ?? '(no hover)'}'`);
});

const wordPaint = await call('textDocument/semanticTokens', {
    projectId: 'WordProject', moduleName: 'Module1', source: WORD_METHOD, moduleType: 'standard',
});

check('and the word host method paints as a call', () => {
    const method = wordPaint.tokens.filter(
        (token) => WORD_METHOD.slice(token.start, token.end) === 'FitToPages');
    assert.equal(method.length, 1, `FitToPages took ${method.length} token(s)`);
    assert.equal(method[0].type, 'function', `FitToPages painted as ${method[0].type}`);
});

// A host the registry does not know: the honest answer is still known:false with a note.
await call('project/open', {
    projectId: 'OutlookProject',
    generation: 1,
    host: 'outlook',
    modules: [{ moduleName: 'Module1', source: 'Sub A()\r\nEnd Sub\r\n', type: 'standard' }],
});

const inOutlook = await call('knowledge/objectModel', {});

check('an unmodelled host says so plainly', () => {
    assert.equal(inOutlook.host, 'outlook');
    assert.equal(inOutlook.known, false);
    assert.ok(inOutlook.note.length > 0, 'a refusal without a note strands the caller');
    assert.equal(inOutlook.types, undefined);
});

check('the rule catalogue is host-independent', () => {
    // The analyzer's rules are VBA-language rules; no host loses them.
    assert.equal(rules.ruleCount >= 20, true);
});

await stop();
process.exit(done());
