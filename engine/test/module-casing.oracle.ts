// The oracle for the generated modules' identifier casing: every name the four host object
// models and the VBA runtime canonicalize, and the analyzer's own reading of what a module
// declares. Both come from the ANALYZER, so this product measures against the same tables the
// sibling does (xlide_vscode's tests/vbaTestModuleIdentifierCasing.test.ts) rather than against
// a copied list that can rot.
//
// TypeScript, and bundled at test time, because the analyzer is TypeScript source in a sibling
// checkout - the same way engine/src reaches it.

import {
    buildModuleSymbols,
    getHostConstants,
    getHostGlobals,
    getHostMembers,
    VBA_RUNTIME_CONSTANTS,
    VBA_RUNTIME_FUNCTIONS,
    VBA_RUNTIME_OBJECTS,
    type VbaSymbol,
} from '../../../xlide_vscode/src/analyzer';
import { getExcelObjectModel, type HostObjectModel } from '../../../xlide_vscode/src/analyzer/host/excelObjectModel';
import { getWordObjectModel } from '../../../xlide_vscode/src/analyzer/host/wordObjectModel';
import { getPowerPointObjectModel } from '../../../xlide_vscode/src/analyzer/host/powerpointObjectModel';
import { getAccessObjectModel } from '../../../xlide_vscode/src/analyzer/host/accessObjectModel';

/** Lowercased name to every canonical spelling the models and the runtime carry. */
export function canonicalSpellings(): Map<string, string[]> {
    const byLower = new Map<string, Set<string>>();
    const add = (name: string): void => {
        if (!/^[A-Za-z_]\w*$/.test(name)) {
            return;
        }

        const key = name.toLowerCase();
        const spellings = byLower.get(key) ?? new Set<string>();
        spellings.add(name);
        byLower.set(key, spellings);
    };

    const models: HostObjectModel[] = [
        getExcelObjectModel(),
        getWordObjectModel(),
        getPowerPointObjectModel(),
        getAccessObjectModel(),
    ];

    for (const model of models) {
        for (const qualified of Object.keys(model.types)) {
            for (const member of getHostMembers(qualified, model)) {
                add(member.name);
            }
        }

        for (const global of getHostGlobals(model)) {
            add(global.name);
        }

        for (const constant of getHostConstants(model)) {
            add(constant.name);
        }
    }

    for (const object of VBA_RUNTIME_OBJECTS) {
        add(object.name);
        for (const member of object.members ?? []) {
            add(member.name);
        }
    }

    for (const fn of VBA_RUNTIME_FUNCTIONS) {
        add(fn.name);
    }

    for (const constant of VBA_RUNTIME_CONSTANTS) {
        add(constant.name);
    }

    return new Map([...byLower].map(([key, spellings]) => [key, [...spellings]]));
}

/** Every name a module's own text declares: procedures, parameters, locals, constants. */
export function declaredNames(source: string): string[] {
    const names: string[] = [];
    const walk = (symbols: readonly VbaSymbol[] | undefined): void => {
        for (const symbol of symbols ?? []) {
            names.push(symbol.name);
            walk(symbol.children);
        }
    };

    walk(buildModuleSymbols('Probe', 'standard', source).root.children);
    return names;
}
