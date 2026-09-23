// The one place the analyzer pin is honoured.
//
// The analyzer is imported from the editor extension's source by literal relative paths into
// `xlide_vscode/src`. XLIDE_ANALYZER_ROOT names a pinned copy of that source (tools\Pin-Analyzer.ps1),
// and every step that reads the analyzer - the build, the typecheck, a test that bundles analyzer
// code - resolves those imports through here. A pin is only real in the readers it reaches: the
// build was pinned and the typecheck was not, and the 0.20.0 gate failed on a half-written line in
// the checkout that the release could never have contained (lessons.md 97).

import { existsSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';

const marker = /\/xlide_vscode\/src\//;

/** The pinned analyzer's root, or null when nothing is pinned and the checkout next door is read. */
export const analyzerPin = () =>
    (process.env.XLIDE_ANALYZER_ROOT ? resolve(process.env.XLIDE_ANALYZER_ROOT) : null);

/**
 * Where a resolved relative import goes under the pin. `moved: false` when the path is not inside
 * an `xlide_vscode/src`, and nothing moves. `target: null` when the pin has no such file: a broken
 * pin is an error, never a silent fallback to the checkout.
 */
export function pinnedTarget(full, pin) {
    const path = full.split('\\').join('/');
    const found = marker.exec(path);
    if (!found) {
        return { moved: false };
    }

    const tail = path.slice(found.index + found[0].length);
    for (const candidate of [join(pin, tail), join(pin, `${tail}.ts`), join(pin, tail, 'index.ts')]) {
        if (existsSync(candidate) && statSync(candidate).isFile()) {
            return { moved: true, target: candidate, tail };
        }
    }

    return { moved: true, target: null, tail };
}

/** An esbuild plugin sending every relative import that lands in `xlide_vscode/src` to the pin. */
export function pinnedAnalyzerPlugin(pin) {
    return {
        name: 'pinned-analyzer',
        setup(build) {
            build.onResolve({ filter: /^\.{1,2}\// }, (args) => {
                const redirect = pinnedTarget(resolve(args.resolveDir, args.path), pin);
                if (!redirect.moved) {
                    return null;
                }

                return redirect.target
                    ? { path: redirect.target }
                    : { errors: [{ text: `the pinned analyzer at ${pin} has no ${redirect.tail}` }] };
            });
        },
    };
}

/**
 * What a pinned read actually read, from the paths it loaded: the files that came from the
 * checkout next door (there must be none) and how many came from the pin.
 */
export function readFromPin(paths, pin) {
    return {
        sibling: paths.filter((one) => /xlide_vscode[\\/]src[\\/]/.test(one)),
        pinned: paths.filter((one) => resolve(one).toLowerCase().startsWith(pin.toLowerCase())).length,
    };
}
