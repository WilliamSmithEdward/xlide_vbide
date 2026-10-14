// The engine's mirror of the text being typed (#105): edits are offsets into the surface's text,
// so they apply only onto a mirror the host sent, never onto the workbook's text as the last pass
// read it; a mirror whose length disagrees with the surface's after an edit is dropped; and the
// diagnostics answer says whether a mirror stands behind it, so the host can send the text again.
//
//   node test/live-mirror.mjs
//   node test/live-mirror.mjs --exe

import assert from 'node:assert/strict';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('live-mirror');
const { check, done } = reporter();

const SEEDED = 'Option Explicit\r\n\r\nPublic Sub A()\r\nEnd Sub\r\n';
const SURFACE = 'Option Explicit\r\n\r\nPublic Sub A()\r\n    Debug.Print 1 + 1\r\nEnd Sub\r\n';

await call('initialize', {});
await call('project/open', {
    projectId: 'Mirror',
    generation: 1,
    modules: [{ moduleName: 'Typed', source: SEEDED, type: 'standard' }],
});

const held = async () => call('debug/liveSource', { projectId: 'Mirror', moduleName: 'Typed', includeText: true });
const change = (params) => call('textDocument/didChange', { projectId: 'Mirror', moduleName: 'Typed', ...params });

// An edit made against the surface's text, which has a line the seed does not.
const at = SURFACE.indexOf('1 + 1') + '1 + 1'.length;
const typed = { start: at, end: at, text: ' + 1' };

await change({ edits: [typed], length: SURFACE.length + 4 });
const unapplied = await held();
check('an edit with no mirror to apply to is dropped, not applied to the seeded text', () =>
    assert.equal(unapplied.held, false, unapplied.source ?? ''));

const answer = await call('textDocument/diagnostics', { projectId: 'Mirror', generation: 1, moduleName: 'Typed', moduleType: 'standard' });
check('and the diagnostics answer says no mirror stands behind it', () => assert.equal(answer.liveHeld, false));

await change({ source: SURFACE });
await change({ edits: [typed], length: SURFACE.length + 4 });
const applied = await held();
check('onto a mirror, an edit whose length agrees applies', () =>
    assert.equal(applied.source, SURFACE.replace('1 + 1', '1 + 1 + 1')));

const after = await call('textDocument/diagnostics', { projectId: 'Mirror', generation: 1, moduleName: 'Typed', moduleType: 'standard' });
check('and the diagnostics answer says one does', () => assert.equal(after.liveHeld, true));

await change({ edits: [{ start: 0, end: 0, text: "'" }], length: 1 });
const dropped = await held();
check('an edit whose resulting length disagrees with the surface drops the mirror', () =>
    assert.equal(dropped.held, false));

await change({ source: SURFACE });
await change({ edits: [{ start: 0, end: 0, text: "'" }] });
const unchecked = await held();
check('an edit sent without a length still applies, as an older host sends it', () =>
    assert.equal(unchecked.source, `'${SURFACE}`));

stop();
process.exit(done());
