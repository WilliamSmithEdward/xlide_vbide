// Typechecks the engine against the analyzer the build reads.
//
// Day to day that is the checkout next door, and this is `tsc --noEmit`, unchanged, which is what
// CI runs. With XLIDE_ANALYZER_ROOT set the build reads a PIN, and so must this. The imports in
// engine/src are literal relative paths into `xlide_vscode/src`, so tsc on its own went on checking
// the checkout next door while the build shipped the pin: the 0.20.0 gate failed on a half-written
// line in that checkout's projectService.ts, a file the release was never going to contain
// (lessons.md 97). A check of what does not ship says nothing about what does, in either direction.

import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';
import { analyzerPin, pinnedTarget, readFromPin } from './pinned-analyzer.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const pin = analyzerPin();

if (!pin) {
    const tsc = createRequire(import.meta.url).resolve('typescript/bin/tsc');
    try {
        execFileSync(process.execPath, [tsc, '--noEmit'], { cwd: here, stdio: 'inherit' });
    } catch (error) {
        process.exit(typeof error.status === 'number' ? error.status : 1);
    }
    process.exit(0);
}

const parsed = ts.getParsedCommandLineOfConfigFile(join(here, 'tsconfig.json'), {}, {
    ...ts.sys,
    onUnRecoverableConfigFileDiagnostic: (diagnostic) => {
        console.error(ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n'));
        process.exit(1);
    },
});

/*
 * A PACKAGE the pin imports is typed as the checkout next door types it. The pin is a clone
 * without node_modules, so `import * as vscode from 'vscode'` found nothing and every callback the
 * analyzer hands the workspace came out implicitly `any` - two errors in a file the checkout itself
 * compiles cleanly. CI installs the analyzer's packages beside it and typechecks with them; asking
 * from the matching path in the checkout is the same answer. Only package imports take this route:
 * no source is read from the checkout, which the assertion below holds to.
 */
const siblingRoot = resolve(here, '..', '..', 'xlide_vscode', 'src');
const packagesFrom = (containingFile) => {
    const inside = resolve(containingFile).toLowerCase().startsWith(pin.toLowerCase());
    return inside && existsSync(siblingRoot)
        ? join(siblingRoot, resolve(containingFile).slice(pin.length))
        : containingFile;
};

const missing = [];
const host = ts.createCompilerHost(parsed.options);
host.resolveModuleNameLiterals = (literals, containingFile, redirected, options, containingSourceFile) =>
    literals.map((literal) => {
        const relative = /^\.{1,2}\//.test(literal.text);
        if (relative) {
            // The build's own redirect, through the same function.
            const redirect = pinnedTarget(resolve(dirname(containingFile), literal.text), pin);
            if (redirect.moved) {
                if (!redirect.target) {
                    missing.push(redirect.tail);
                    return { resolvedModule: undefined };
                }

                return {
                    resolvedModule: {
                        resolvedFileName: redirect.target,
                        extension: ts.Extension.Ts,
                        isExternalLibraryImport: false,
                    },
                };
            }
        }

        return ts.resolveModuleName(literal.text, relative ? containingFile : packagesFrom(containingFile),
            options, host, undefined, redirected, ts.getModeForUsageLocation(containingSourceFile, literal, options));
    });

const program = ts.createProgram({ rootNames: parsed.fileNames, options: parsed.options, host });
const diagnostics = ts.getPreEmitDiagnostics(program);

if (missing.length > 0) {
    console.error(`the pinned analyzer at ${pin} has no ${missing[0]}`
        + (missing.length > 1 ? ` (and ${missing.length - 1} more)` : ''));
    process.exit(1);
}

// THE CHECK PROVES WHICH ANALYZER IT READ, as the build does from its metafile. A pinned check
// that read one file from the working tree is not a pinned check.
const { sibling, pinned } = readFromPin(program.getSourceFiles().map((file) => file.fileName), pin);
if (sibling.length > 0) {
    console.error(`The analyzer was pinned at ${pin}, but ${sibling.length} file(s) were read `
        + `from the checkout next door instead, starting with ${sibling[0]}.`);
    process.exit(1);
}

if (diagnostics.length > 0) {
    console.error(ts.formatDiagnostics(diagnostics, {
        getCanonicalFileName: (name) => name,
        getCurrentDirectory: () => here,
        getNewLine: () => '\n',
    }));
    process.exit(1);
}

console.log(`typecheck: clean, ${pinned} analyzer file(s), all from the pin at ${pin}`);
