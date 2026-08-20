// IDENTIFIER CASING IN THE GENERATED VBA MODULES, held to the host object models.
//
// VBA cases identifiers PROJECT-WIDE to the latest declaration it sees. A module that declares
// `ByVal value` re-spells every `.Value` in the developer's project the moment it is installed,
// and XlideAssert is installed permanently - so a lowercase parameter here is a rename of the
// user's code that nobody asked for (xlide_vscode#38). The four names that issue listed were
// not the whole set: `Condition` and `MacroName` were re-casing projects too, found by
// measuring rather than by reading (xlide_vbide#3).
//
// So this measures, and it measures against the ANALYZER's own tables - the four host object
// models and the VBA runtime - through the same oracle the sibling product's test uses. Any
// name these modules declare that matches a canonical name case-insensitively must match it
// exactly. The next helper that declares `ByVal count` fails here instead of shipping.
//
// The module text is read out of the SHIM's own sources, because that is where this product's
// copies live: a check against a transcription would pass while the shipped module was wrong.
//
//   node test/module-casing.mjs

import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { reporter } from './harness.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const shim = resolve(here, '..', '..', 'src', 'Xlide.Vbe.Shim', 'Editor');
const { check, done } = reporter();

// The oracle is TypeScript reaching into the analyzer checkout, so it is bundled the way
// engine/src is. A missing checkout fails loudly here rather than passing vacuously.
const out = await mkdtemp(join(tmpdir(), 'xlide-casing-'));
const bundle = join(out, 'oracle.mjs');
await build({
    entryPoints: [join(here, 'module-casing.oracle.ts')],
    bundle: true,
    format: 'esm',
    platform: 'node',
    outfile: bundle,
    logLevel: 'error',
});

const oracle = await import(`file:///${bundle.replace(/\\/g, '/')}`);
const canonical = oracle.canonicalSpellings();

check('the analyzer hands over a real name set to measure against', () => {
    assert.ok(canonical.size > 10_000, `expected the models' names, got ${canonical.size}`);
});

/**
 * Every VBA line a C# file spells, from its string literals: the modules are built by
 * concatenation and by Append, and both forms are one quoted literal per line. Interpolated
 * lines carry `{...}` holes, which are left as they are - a hole is not an identifier.
 */
async function vbaTextOf(file) {
    const source = await readFile(join(shim, file), 'utf8');
    const lines = [];
    for (const match of source.matchAll(/"((?:[^"\\]|\\.)*)\\r\\n"/g)) {
        // Braces in order: the DOUBLED ones are literal text and are parked first, so the
        // interpolation holes that remain can be replaced without eating the JSON that the
        // runner's own lines are made of - `{""outcome"":...}` is not a hole.
        lines.push(match[1]
            .replace(/\\"/g, '"')
            .replace(/\\\\/g, '\\')
            .replace(/\{\{/g, '')
            .replace(/\}\}/g, '')
            .replace(/\{[^{}]*\}/g, 'Placeholder')
            .replace(//g, '{')
            .replace(//g, '}'));
    }

    return lines.join('\r\n');
}

for (const file of ['TestRunService.AssertModule.cs', 'TestRunService.cs']) {
    const text = await vbaTextOf(file);
    check(`${file} yields its VBA text`, () => {
        assert.ok(text.split('\r\n').length > 20, `only ${text.split('\r\n').length} line(s) came out`);
    });

    const declared = oracle.declaredNames(text);
    check(`${file} declares identifiers this check can see`, () => {
        assert.ok(declared.length > 3, `only ${declared.length} declaration(s) were found`);
    });

    check(`${file} spells every host name the way the models spell it`, () => {
        const wrong = [...new Set(declared)]
            .filter((name) => {
                const spellings = canonical.get(name.toLowerCase());
                return spellings !== undefined && !spellings.includes(name);
            })
            .map((name) => `${name} should be ${canonical.get(name.toLowerCase()).join(' or ')}`);
        assert.deepEqual(wrong, [], `these declarations re-case the developer's project:\n  ${wrong.join('\n  ')}`);
    });
}

// The wire protocol is spelled in string literals, and a re-casing that did not know the
// difference would rename the JSON keys the page parses (xlide_vbide#3, item 3).
const runner = await vbaTextOf('TestRunService.cs');
check('the result JSON keys stay lower case, whatever the identifiers do', () => {
    for (const key of ['outcome', 'number', 'source', 'message', 'output']) {
        assert.ok(runner.includes(`""${key}""`), `the ${key} key is not spelled in lower case any more`);
    }
});

process.exitCode = done();
