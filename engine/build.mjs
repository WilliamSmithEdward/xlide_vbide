// Bundles the engine, and optionally packages it as one executable.
//
// The analyzer is imported from the editor extension's source rather than copied, so there is one
// implementation shared by both products. That coupling is real and load bearing, so it is checked
// before anything else runs and reported plainly when it is missing.

import { execFileSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import esbuild from 'esbuild';
import { analyzerPin, pinnedAnalyzerPlugin, readFromPin } from './pinned-analyzer.mjs';

const here = dirname(fileURLToPath(import.meta.url));
// THE SIBLING CHECKOUT BY DEFAULT, A PIN WHEN ONE IS ASKED FOR. That checkout is a working tree
// somebody is usually working in, and a release cannot be built from a tree that is mid-edit:
// the currency guard compares file times, so any edit in progress makes the packaged engine look
// stale, and repackaging to satisfy it would ship whatever was half-written. `tools\Pin-Analyzer.ps1`
// clones that repository at one commit and prints the root; XLIDE_ANALYZER_ROOT points here at it.
// Day to day nobody sets it and this builds against the checkout next door, which is the whole
// point of the coupling.
const pin = analyzerPin();
const analyzerRoot = pin ?? resolve(here, '..', '..', 'xlide_vscode', 'src');
const outDir = join(here, 'dist');
const bundlePath = join(outDir, 'engine.cjs');
const exePath = join(outDir, 'xlide-engine.exe');
const wantsPackage = process.argv.includes('--package');

if (!existsSync(join(analyzerRoot, 'analysisWorkerLogic.ts'))) {
    console.error(`The analyzer source was not found at ${analyzerRoot}.`);
    console.error('The engine builds against the editor extension checkout so both share one analyzer.');
    console.error('Clone it next to this repository, or adjust the path in engine/build.mjs.');
    process.exit(1);
}

mkdirSync(outDir, { recursive: true });

// THE PLUGIN IS WHAT MAKES THE PIN REAL. Setting XLIDE_ANALYZER_ROOT used to do two things only:
// pass the existence check above, and point the currency guards at a folder whose timestamps
// never move. The imports in engine/src are literal relative paths, so esbuild went on reading the
// checkout next door - and the guard, now comparing against a static directory, had nothing to
// say about it. A pin that silences the alarm while the build reads the tree is worse than no
// pin: v0.17.0 was cut believing it shipped pinned bits and did not (2026-09-21). The redirect
// lives in pinned-analyzer.mjs, shared with the typecheck and the tests that bundle analyzer code.
const built = await esbuild.build({
    entryPoints: [join(here, 'src', 'main.ts')],
    ...(pin ? { plugins: [pinnedAnalyzerPlugin(pin)] } : {}),
    bundle: true,
    // A single executable embeds a script, and the embedded script is evaluated as CommonJS.
    format: 'cjs',
    platform: 'node',
    target: 'node22',
    // XLIDE_NO_MINIFY=1 builds one that can be PROFILED. A minified bundle names its functions
    // `XP` and `Tt`, so `node --cpu-prof` answers a perf question with two-letter labels and the
    // question stays open; unminified it answers `statementTokensCached`, which is how the
    // quadratic memo lookup behind the editor's freezes was found (2026-08-21). Never set for a
    // shipped build: the bundle is 10.7 MB instead of 8.8 MB.
    minify: process.env.XLIDE_NO_MINIFY ? false : true,
    sourcemap: false,
    outfile: bundlePath,
    logLevel: 'info',
    define: {
        // The engine names its own build (initialize answers it). Deliberately per-build: an
        // identical rebuild used to produce an identical executable, and a machine whose
        // application-control policy caches verdicts per file hash could then never re-judge a
        // blocked engine however many times it was repackaged (2026-08-19).
        __ENGINE_BUILT__: JSON.stringify(new Date().toISOString()),
    },
    // Kept so the build can prove WHICH analyzer it read, below.
    metafile: true,
});

// THE BUILD PROVES ITS OWN PROVENANCE. A redirect that silently does nothing is the failure this
// exists for: for one release XLIDE_ANALYZER_ROOT moved the guards and not the build, and the
// only thing that could have caught it - which files esbuild actually read - was sitting in the
// metafile nobody asked for. Now the answer is asserted at the point of truth, every build, at
// no cost. A pinned build that read one file from the working tree is not a pinned build.
if (pin) {
    const { sibling, pinned } = readFromPin(Object.keys(built.metafile.inputs), pin);
    if (sibling.length > 0) {
        console.error(
            `The analyzer was pinned at ${pin}, but ${sibling.length} file(s) were read `
            + `from the checkout next door instead, starting with ${sibling[0]}.`);
        process.exit(1);
    }

    console.log(`analyzer: ${pinned} file(s), all from the pin at ${pin}`);
}

const bundleSize = statSync(bundlePath).size;
console.log(`bundle ${bundlePath} (${(bundleSize / 1024 / 1024).toFixed(2)} MB)`);

if (!wantsPackage) {
    process.exit(0);
}

// Package as a single executable: the runtime is copied, the script is compiled to a blob, and the
// blob is injected into a reserved section of the copy. The result needs nothing installed.
const seaConfigPath = join(outDir, 'sea-config.json');
const blobPath = join(outDir, 'engine.blob');

writeFileSync(
    seaConfigPath,
    JSON.stringify(
        {
            main: bundlePath,
            output: blobPath,
            disableExperimentalSEAWarning: true,
            useSnapshot: false,

            // MEASURED, AND IT DOES NOTHING HERE (2026-08-11). The obvious reading of a 3.2s engine
            // start is that V8 is compiling the 2.26 MB bundle on every launch, and `useCodeCache`
            // is the switch that would cache that compilation into the blob. Six warm runs before:
            // 190-194ms to the listening line. Six after: 187-192ms. It is noise, and it costs 2 MB
            // of executable. The cost is not compilation.
            //
            // What the 3.2s actually is: reading a 90 MB image off disk while Excel, the VBE and a
            // WebView2 browser process are all starting at the same time. Standalone the same
            // launch is 1.27s cold and 190ms warm. Do not re-try this switch without a measurement
            // showing something changed.
            useCodeCache: false,
        },
        null,
        2,
    ),
);

execFileSync(process.execPath, ['--experimental-sea-config', seaConfigPath], { stdio: 'inherit' });
copyFileSync(process.execPath, exePath);

const postject = join(here, 'node_modules', 'postject', 'dist', 'cli.js');
execFileSync(
    process.execPath,
    [
        postject,
        exePath,
        'NODE_SEA_BLOB',
        blobPath,
        '--sentinel-fuse',
        'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2',
    ],
    { stdio: 'inherit' },
);

const exeSize = statSync(exePath).size;
console.log(`executable ${exePath} (${(exeSize / 1024 / 1024).toFixed(2)} MB)`);
