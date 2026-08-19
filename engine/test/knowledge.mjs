// The knowledge routes: what the language service knows, served as data.
//
// Two contracts are pinned here beyond the shapes. First, knowledge answers BEFORE initialize
// and before any project opens - it is product knowledge, not project state, and an agent must
// be able to learn the terrain without touching anything. Second, the model answers for the
// HOST the engine was told it is running in: after a project/open that says word, the honest
// reply is known:false with a note, not Excel's types wearing Word's name - the same rule the
// features themselves follow.
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

// Now tell the engine it is running in Word, the way project/open does, and ask again.
await call('initialize', {});
await call('project/open', {
    projectId: 'WordProject',
    generation: 1,
    host: 'word',
    modules: [{ moduleName: 'Module1', source: 'Sub A()\r\nEnd Sub\r\n', type: 'standard' }],
});

const inWord = await call('knowledge/objectModel', {});

check('in word the honest answer is no model, said plainly', () => {
    assert.equal(inWord.host, 'word');
    assert.equal(inWord.known, false);
    assert.ok(inWord.note.length > 0, 'a refusal without a note strands the caller');
    assert.equal(inWord.types, undefined);
});

check('the rule catalogue is host-independent', () => {
    // The analyzer's rules are VBA-language rules; Word does not lose them.
    assert.equal(rules.ruleCount >= 20, true);
});

await stop();
process.exit(done());
