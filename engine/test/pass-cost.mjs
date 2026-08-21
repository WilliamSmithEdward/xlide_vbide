// Analysing a module a second time must not cost more than analysing it the first time.
//
// THE DEFECT THIS EXISTS FOR. The analyzer memoises per source string and finds its entries with
// `===`. Inside one analysis that is free: the caller keeps handing back the very string the entry
// was stored under, and V8 settles `===` on the pointer. Across two analyses it is not. A request
// arrives over the pipe, `JSON.parse` builds a NEW string with the same 1.5 MB of content, and
// every lookup then compares the two character by character. `statementTokensCached` is asked
// "hundreds of thousands of times" per pass by its own account, so the pass turns quadratic in
// module size.
//
// Measured on the 64,802-line perf fixture on 2026-08-21: the FIRST analysis 0.9s - fast only
// because the cache was empty and stored the caller's own instance - and every analysis after it
// 15.8 SECONDS. Not the incremental path's doing: a forced FULL pass cost 9.1s where the identical
// cold full pass cost 0.5s. That is what the owner had been reporting as the editor freezing for
// ten to fifteen seconds after a keystroke in a large module, and as analysis passes that gave up
// before they published, which reads from the outside as the analyzer having stopped working.
//
// The engine's side of the bargain is to hand a memo a stable key: text equal to text already held
// is answered with the instance already held (dispatcher, stableSources). A few of them, not one,
// because a developer alternates - break something, look, undo, break it again - and holding only
// the latest sends the version coming BACK as a fresh instance every time.
//
// So this measures a RATIO rather than a duration: how much a repeat analysis costs against the
// first one, which normalises away whatever machine it runs on. Broken, that ratio was 8 to 18.
// Whole, it is under 1.
//
//   node test/pass-cost.mjs

import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { reporter, startEngine } from './harness.mjs';

const { call, stop } = await startEngine('pass-cost');
const { check, done } = reporter();

const PROJECT = 'passcost.xlsm';

// Big enough for a quadratic to be unmistakable, small enough to stay a test: about 32,400
// lines, half of VBA's per-module ceiling. Sized by measurement rather than by guess - at 2,400
// procedures a repeat cost 4.7x on the defective engine against a threshold of 4, which is too
// thin to call a proof. Here it costs 8x, and a whole engine costs 0.4x.
const PROCEDURES = 3600;

function moduleText(count, marker) {
    const lines = ['Option Explicit', '', `' ${marker}`, ''];
    for (let n = 0; n < count; n++) {
        lines.push(`Public Function V${n}(ByVal seed As Long) As Long`);
        lines.push('    Dim total As Long');
        lines.push(`    total = seed * ${n + 1}`);
        if (n > 0) {
            lines.push('    If total > 100 Then');
            lines.push(`        total = total - V${n - 1}(seed)`);
            lines.push('    End If');
        }
        lines.push(`    V${n} = total`);
        lines.push('End Function');
        lines.push('');
    }
    return lines.join('\r\n');
}

const FIRST = moduleText(PROCEDURES, 'first');
const SECOND = moduleText(PROCEDURES, 'second');

let generation = 1;

const seed = (source) => call('project/open', {
    projectId: PROJECT,
    generation,
    host: 'excel',
    modules: [{ moduleName: 'Big', source, type: 'standard' }],
});

// The caret moves on every call so the engine's OWN answer memo cannot serve it: what is being
// measured is the cost of actually analysing, not the cost of remembering.
async function analyse(source, caret) {
    const started = performance.now();
    const answer = await call('textDocument/diagnostics', {
        projectId: PROJECT,
        generation,
        documentKey: `${PROJECT}\0big`,
        moduleName: 'Big',
        moduleType: 'standard',
        source,
        activeIncompleteExpressionOffset: caret,
    });
    return { took: performance.now() - started, diagnostics: answer.diagnostics };
}

const median = (values) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)];

await call('initialize', {});
await seed(FIRST);

// The first analysis of a module IS the honest full cost of analysing it, and it is the yardstick.
const cold = await analyse(FIRST, 10);

const repeats = [];
for (let i = 0; i < 4; i++) {
    repeats.push((await analyse(FIRST, 20 + i)).took);
}

check('analysing the same text again does not cost more than analysing it the first time', () => {
    const ratio = median(repeats) / cold.took;
    assert.ok(
        ratio < 4,
        `repeat analysis cost ${ratio.toFixed(1)}x the first one `
        + `(first ${cold.took.toFixed(0)}ms, repeats ${repeats.map((t) => t.toFixed(0)).join(', ')}ms). `
        + 'The analyzer is re-comparing the whole module on every memo lookup, which is quadratic '
        + 'in module size: see stableSources in the dispatcher, and xlide_vscode#45.');
});

// ALTERNATING, which is what a developer actually does: break it, look at it, put it back. Holding
// one instance per module is not enough for this - the version coming back is a fresh string while
// the memo still holds the older equal one, which is the same quadratic reached the other way.
const swapped = [];
for (let i = 0; i < 4; i++) {
    generation += 1;
    const text = i % 2 === 0 ? SECOND : FIRST;
    await seed(text);
    swapped.push((await analyse(text, 30 + i)).took);
}

check('alternating between two versions of the module costs no more than the first analysis', () => {
    const ratio = median(swapped) / cold.took;
    assert.ok(
        ratio < 4,
        `alternating analysis cost ${ratio.toFixed(1)}x the first one `
        + `(first ${cold.took.toFixed(0)}ms, alternating ${swapped.map((t) => t.toFixed(0)).join(', ')}ms). `
        + 'Only the newest version of each module is being held by instance; see STABLE_SOURCES_HELD.');
});

// Fast for the right reason. A memo that answered without analysing would pass every check above.
generation += 1;
const broken = FIRST.replace('End Function\r\n\r\nPublic Function V1(', 'End Functiona\r\n\r\nPublic Function V1(');
assert.notEqual(broken, FIRST, 'the probe did not find the closer it meant to break');
await seed(broken);
const complained = await analyse(broken, 40);

check('and it is still analysing: a broken procedure closer is reported', () => {
    assert.ok(
        complained.diagnostics.length > 0,
        'no findings on a module whose first End Function was broken');
});

generation += 1;
await seed(FIRST);
const healed = await analyse(FIRST, 50);

check('and the findings clear when it is put back', () => {
    assert.equal(
        healed.diagnostics.length,
        0,
        `got ${JSON.stringify(healed.diagnostics.map((d) => d.code))}`);
});

console.log(
    `\nfirst ${cold.took.toFixed(0)}ms, repeats ${median(repeats).toFixed(0)}ms, `
    + `alternating ${median(swapped).toFixed(0)}ms (medians)`);

stop();
done();
