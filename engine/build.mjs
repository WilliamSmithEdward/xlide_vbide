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

const here = dirname(fileURLToPath(import.meta.url));
const analyzerRoot = resolve(here, '..', '..', 'xlide_vscode', 'src');
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

await esbuild.build({
    entryPoints: [join(here, 'src', 'main.ts')],
    bundle: true,
    // A single executable embeds a script, and the embedded script is evaluated as CommonJS.
    format: 'cjs',
    platform: 'node',
    target: 'node22',
    minify: true,
    sourcemap: false,
    outfile: bundlePath,
    logLevel: 'info',
});

// The reference lexer, exposed as a command so its output can be compared with the ported one over
// a corpus. Built alongside because it shares this toolchain; it is not part of the product and
// leaves when the port is finished.
await esbuild.build({
    entryPoints: [join(here, 'src', 'lexdump.ts')],
    bundle: true,
    format: 'cjs',
    platform: 'node',
    target: 'node22',
    minify: false,
    sourcemap: false,
    outfile: join(outDir, 'lexdump.cjs'),
    logLevel: 'warning',
});

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
